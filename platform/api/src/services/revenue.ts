/**
 * Revenue dashboard aggregation service.
 *
 * Computes revenue from completed marking items using the pricing engine,
 * then aggregates by date, contractor, pricing group, and work order.
 *
 * Replaces Code.js _buildRevenuePayload_ (lines 12594–12842).
 */
import { eq, and, between, asc } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { markingItems, workOrders, contractors, signinEntries, payRates, invoices } from '../db/schema.js';
import { priceMarkingItem, REVENUE_BUCKETS, type RateRow, type PricingGroup } from './pricing.js';
import { billingRemap, type RemapRule } from './billingRemap.js';

export interface RevenueFilters {
  startDate: string;
  endDate: string;
}

export interface DailyRevenue {
  date: string;
  revenue: number;
  byGroup: Record<string, number>;
}

export interface ContractorRevenue {
  contractorId: string;
  contractorName: string;
  revenue: number;
}

export interface NeedsPricingItem {
  itemId: string;
  woNumber: string;
  category: string;
  quantity: string | null;
  unit: string | null;
  reason: string;
  contractor?: string;
  contractNum?: string;
  regionCode?: string;
}

export interface DailyLabor {
  date: string;
  cost: number;
}

export interface RevenueData {
  range: { start: string; end: string };
  totals: {
    revenue: number;
    byBucket: Record<string, number>;
    byGroup: Record<string, number>;
    items: number;
    needsPricing: number;
    invoicedRevenue: number;
    wipRevenue: number;
    pctInvoiced: number;
  };
  daily: DailyRevenue[];
  byContractor: ContractorRevenue[];
  needsPricingItems: NeedsPricingItem[];
  laborDaily: DailyLabor[];
  laborTotals: { total: number };
}

/**
 * Compute revenue data for a date range.
 * Prices every completed marking item using the pricing engine,
 * then aggregates into the shape the Revenue dashboard expects.
 */
export async function getRevenueData(
  db: Db,
  orgId: string,
  filters: RevenueFilters,
  rates: RateRow[],
  categoryGroupMap: Record<string, string>,
  multipliers: {
    lineWidth: Record<string, number>;
    line12: Record<string, number>;
    extrudedUnit: Record<string, number>;
    preformedUnit: Record<string, number>;
  },
  remapRules: RemapRule[] = [],
): Promise<RevenueData> {
  const { startDate, endDate } = filters;

  // Load completed marking items with WO metadata
  const rows = await db.select({
    itemId: markingItems.id,
    category: markingItems.category,
    quantity: markingItems.quantity,
    unit: markingItems.unit,
    dateCompleted: markingItems.dateCompleted,
    woId: markingItems.woId,
    woNumber: workOrders.woNumber,
    contractorId: workOrders.contractorId,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    location: workOrders.location,
  })
    .from(markingItems)
    .innerJoin(workOrders, eq(markingItems.woId, workOrders.id))
    .where(and(
      eq(markingItems.orgId, orgId),
      eq(markingItems.status, 'completed'),
      between(markingItems.dateCompleted, startDate, endDate),
    ))
    .orderBy(asc(markingItems.dateCompleted));

  // Load contractor names
  const contractorRows = await db.select({ id: contractors.id, name: contractors.name })
    .from(contractors).where(eq(contractors.orgId, orgId));
  const contractorNameMap = new Map(contractorRows.map(c => [c.id, c.name]));

  // Load invoiced WO set — any WO with an invoice record is "invoiced"
  const invoiceRows = await db.select({ woId: invoices.woId })
    .from(invoices)
    .where(eq(invoices.orgId, orgId));
  const invoicedWoIds = new Set(invoiceRows.map(i => i.woId).filter(Boolean));

  // Price each item and aggregate
  const dailyMap = new Map<string, { revenue: number; byGroup: Record<string, number> }>();
  const contractorTotals = new Map<string, number>();
  const groupTotals: Record<string, number> = {};
  const needsPricing: NeedsPricingItem[] = [];
  const woRevenue = new Map<string, { woNumber: string; contractorName: string; location: string; revenue: number; items: number }>();
  let totalRevenue = 0;
  let totalItems = 0;
  let invoicedRevenue = 0;
  let wipRevenue = 0;

  for (const row of rows) {
    const contractorName = contractorNameMap.get(row.contractorId) || '';

    // Strip /EXT suffix BEFORE remap — Code.js _buildRevenuePayload_ (line 12618)
    // does split('/')[0] before calling _billingRemap_.
    const cnStripped = String(row.contractNum || '').split('/')[0].trim();

    // Apply billing remap so sub-prime work (e.g., Metro Express BK → M)
    // matches the correct pricing row.
    const billed = billingRemap(
      remapRules,
      cnStripped,
      row.regionCode || '',
      contractorName,
    );

    const result = priceMarkingItem(
      {
        category: row.category,
        quantity: row.quantity,
        unit: row.unit,
        dateCompleted: row.dateCompleted,
      },
      {
        contractor: contractorName,
        contractNum: billed.contractNum,
        regionCode: billed.regionCode,
      },
      rates,
      categoryGroupMap,
      multipliers,
    );

    totalItems++;

    if (result.reason) {
      needsPricing.push({
        itemId: row.itemId,
        woNumber: row.woNumber,
        category: row.category,
        quantity: row.quantity,
        unit: row.unit,
        reason: result.reason,
        contractor: contractorName,           // M-18
        contractNum: billed.contractNum,      // M-18
        regionCode: billed.regionCode,        // M-18
      });
      continue;
    }

    totalRevenue += result.revenue;

    // Invoiced vs WIP — check if this WO has an invoice
    if (invoicedWoIds.has(row.woId)) {
      invoicedRevenue += result.revenue;
    } else {
      wipRevenue += result.revenue;
    }

    // Daily aggregation
    const date = String(row.dateCompleted || '');
    const daily = dailyMap.get(date) || { revenue: 0, byGroup: {} };
    daily.revenue += result.revenue;
    daily.byGroup[result.group] = (daily.byGroup[result.group] || 0) + result.revenue;
    dailyMap.set(date, daily);

    // Per-contractor
    const prevC = contractorTotals.get(row.contractorId) || 0;
    contractorTotals.set(row.contractorId, prevC + result.revenue);

    // Per-group
    groupTotals[result.group] = (groupTotals[result.group] || 0) + result.revenue;

    // Per-WO (DASH-2: top_wos computation)
    const woEntry = woRevenue.get(row.woId) || {
      woNumber: row.woNumber,
      contractorName: contractorName,
      location: row.location || '',
      revenue: 0,
      items: 0,
    };
    woEntry.revenue += result.revenue;
    woEntry.items++;
    woRevenue.set(row.woId, woEntry);
  }

  // Compute bucket totals (Thermo, MMA, Preform)
  const byBucket: Record<string, number> = {};
  for (const bucket of REVENUE_BUCKETS) {
    byBucket[bucket.key] = bucket.groups.reduce((sum, g) => sum + (groupTotals[g] || 0), 0);
  }

  // ─── Labor Cost Computation ───────────────────────────────
  // Walk sign-in entries for the date range and compute fully-loaded labor cost
  const signinRows = await db.select({
    workDate: signinEntries.workDate,
    classification: signinEntries.classification,
    hoursWorked: signinEntries.hoursWorked,
    otHours: signinEntries.otHours,
  })
    .from(signinEntries)
    .where(and(
      eq(signinEntries.orgId, orgId),
      between(signinEntries.workDate, startDate, endDate),
    ));

  const payRateRows = await db.select()
    .from(payRates)
    .where(eq(payRates.orgId, orgId));

  const laborDailyMap = new Map<string, number>();
  let laborTotal = 0;

  for (const sr of signinRows) {
    const hours = Number(sr.hoursWorked) || 0;
    const ot = Number(sr.otHours) || 0;
    const st = Math.max(0, hours - ot);
    const cls = sr.classification;
    const date = String(sr.workDate || '');

    // Resolve rate for this classification
    const applicableRates = payRateRows
      .filter(r => r.classificationCode === cls && r.effectiveDate <= endDate)
      .sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate));

    if (applicableRates.length === 0) continue;
    const rate = applicableRates[0];
    const rateSt = Number(rate.rateSt) + Number(rate.suppSt);
    const rateOt = Number(rate.rateOt) + Number(rate.suppOt);
    const cost = st * rateSt + ot * rateOt;

    laborTotal += cost;
    laborDailyMap.set(date, (laborDailyMap.get(date) || 0) + cost);
  }

  return {
    range: { start: startDate, end: endDate },
    totals: {
      revenue: totalRevenue,
      byBucket,
      byGroup: groupTotals,
      items: totalItems,
      needsPricing: needsPricing.length,
      invoicedRevenue,
      wipRevenue,
      pctInvoiced: totalRevenue > 0 ? invoicedRevenue / totalRevenue : 0,
    },
    daily: Array.from(dailyMap.entries()).map(([date, d]) => ({ date, ...d })),
    byContractor: Array.from(contractorTotals.entries()).map(([cid, revenue]) => ({
      contractorId: cid,
      contractorName: contractorNameMap.get(cid) || 'Unknown',
      revenue,
    })),
    needsPricingItems: needsPricing,
    byWo: Array.from(woRevenue.values()),
    laborDaily: Array.from(laborDailyMap.entries()).map(([date, cost]) => ({ date, cost: Math.round(cost * 100) / 100 })),
    laborTotals: { total: Math.round(laborTotal * 100) / 100 },
  };
}
