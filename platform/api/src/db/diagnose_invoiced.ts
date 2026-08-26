/**
 * Diagnose invoiced vs WIP revenue split.
 * Compares our computed split against the old app's expected values.
 *
 * Old app (season to date): Invoiced=$813,012  WIP=$60,147  Total=$873,159
 *
 * Usage: npx tsx src/db/diagnose_invoiced.ts
 */
import 'dotenv/config';
import { db } from './client.js';
import { eq, and, between } from 'drizzle-orm';
import { markingItems, workOrders, contractors, invoices } from './schema.js';
import {
  priceMarkingItem, LINE_WIDTH_MULTIPLIER, LINE12_MULTIPLIER,
  EXTRUDED_UNIT_COUNT, PREFORMED_UNIT_COUNT, type RateRow,
} from '../services/pricing.js';
import { billingRemap, type RemapRule } from '../services/billingRemap.js';
import {
  listContractPricing, listMultipliers, listCategories, listBillingRemaps,
} from './queries/settings.js';

const ORG = '3520b16b-892c-490c-a60f-30defa7e15f2';

async function run() {
  // Load all config
  const pricingRows = await listContractPricing(db, ORG);
  const cRows = await db.select({ id: contractors.id, name: contractors.name })
    .from(contractors).where(eq(contractors.orgId, ORG));
  const cMap = new Map(cRows.map(c => [c.id, c.name]));

  const rates: RateRow[] = pricingRows.map(r => ({
    contractor: cMap.get(r.contractorId) || '',
    contractNum: r.contractNum,
    regionCode: r.regionCode || '',
    effectiveDate: r.effectiveDate ? new Date(r.effectiveDate) : null,
    rates: {
      line4: r.rateLine4 ? Number(r.rateLine4) : null,
      line12: r.rateLine12 ? Number(r.rateLine12) : null,
      preformed: r.ratePreformed ? Number(r.ratePreformed) : null,
      extruded: r.rateExtruded ? Number(r.rateExtruded) : null,
      colorSurface: r.rateColorSurface ? Number(r.rateColorSurface) : null,
    },
  }));

  const catRows = await listCategories(db, ORG);
  const catMap: Record<string, string> = {};
  for (const c of catRows) if (c.pricingGroup) catMap[c.name] = c.pricingGroup;

  const mulRows = await listMultipliers(db, ORG);
  const muls = {
    lineWidth: { ...LINE_WIDTH_MULTIPLIER },
    line12: { ...LINE12_MULTIPLIER },
    extrudedUnit: { ...EXTRUDED_UNIT_COUNT },
    preformedUnit: { ...PREFORMED_UNIT_COUNT },
  };
  for (const m of mulRows) {
    const v = Number(m.value);
    if (m.multiplierType === 'line_width') muls.lineWidth[m.categoryName] = v;
    else if (m.multiplierType === 'line12') muls.line12[m.categoryName] = v;
    else if (m.multiplierType === 'extruded_unit') muls.extrudedUnit[m.categoryName] = v;
    else if (m.multiplierType === 'preformed_unit') muls.preformedUnit[m.categoryName] = v;
  }

  const remapRows = await listBillingRemaps(db, ORG);
  const remapRules: RemapRule[] = remapRows.map(r => ({
    sourceContract: r.sourceContract,
    sourceRegion: r.sourceRegion,
    sourceContractor: r.sourceContractor,
    targetContract: r.targetContract,
    targetRegion: r.targetRegion,
    effectiveDate: r.effectiveDate,
  }));

  // Load invoiced WO set
  const invRows = await db.select({ woId: invoices.woId, amount: invoices.amount, invoiceNumber: invoices.invoiceNumber })
    .from(invoices).where(eq(invoices.orgId, ORG));
  const invoicedWoIds = new Set(invRows.map(i => i.woId).filter(Boolean));

  console.log(`\n=== CONFIG ===`);
  console.log(`Rate rows: ${rates.length}`);
  console.log(`Categories with pricing: ${Object.keys(catMap).length}`);
  console.log(`Remap rules: ${remapRules.length}`);
  console.log(`Invoiced WOs: ${invoicedWoIds.size}`);

  // Load all completed items
  const items = await db.select({
    id: markingItems.id,
    category: markingItems.category,
    quantity: markingItems.quantity,
    unit: markingItems.unit,
    dateCompleted: markingItems.dateCompleted,
    woId: markingItems.woId,
    woNumber: workOrders.woNumber,
    contractorId: workOrders.contractorId,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
  })
    .from(markingItems)
    .innerJoin(workOrders, eq(markingItems.woId, workOrders.id))
    .where(and(
      eq(markingItems.orgId, ORG),
      eq(markingItems.status, 'completed'),
      between(markingItems.dateCompleted, '2026-01-01', '2026-08-31'),
    ));

  console.log(`Completed items: ${items.length}`);

  // Price every item and track per-WO revenue
  let totalRevenue = 0;
  let invoicedRevenue = 0;
  let wipRevenue = 0;
  let unpricedCount = 0;
  const woRevenue = new Map<string, { woNumber: string; contractor: string; revenue: number; invoiced: boolean; items: number }>();

  for (const item of items) {
    const contractorName = cMap.get(item.contractorId) || '';
    const cn = String(item.contractNum || '').split('/')[0].trim();
    const billed = billingRemap(remapRules, cn, item.regionCode || '', contractorName);

    const result = priceMarkingItem(
      { category: item.category, quantity: item.quantity, unit: item.unit, dateCompleted: item.dateCompleted },
      { contractor: contractorName, contractNum: billed.contractNum, regionCode: billed.regionCode },
      rates, catMap, muls,
    );

    if (result.reason) { unpricedCount++; continue; }

    totalRevenue += result.revenue;
    const isInvoiced = invoicedWoIds.has(item.woId);
    if (isInvoiced) invoicedRevenue += result.revenue;
    else wipRevenue += result.revenue;

    // Track per-WO
    const woEntry = woRevenue.get(item.woId) || { woNumber: item.woNumber, contractor: contractorName, revenue: 0, invoiced: isInvoiced, items: 0 };
    woEntry.revenue += result.revenue;
    woEntry.items++;
    woRevenue.set(item.woId, woEntry);
  }

  console.log(`\n=== RESULTS (our system) ===`);
  console.log(`Total Revenue:    $${totalRevenue.toFixed(2)}`);
  console.log(`Invoiced Revenue: $${invoicedRevenue.toFixed(2)}`);
  console.log(`WIP Revenue:      $${wipRevenue.toFixed(2)}`);
  console.log(`Pct Invoiced:     ${(totalRevenue > 0 ? invoicedRevenue / totalRevenue * 100 : 0).toFixed(1)}%`);
  console.log(`Unpriced:         ${unpricedCount} items`);

  console.log(`\n=== OLD APP (expected) ===`);
  console.log(`Total Revenue:    $873,159`);
  console.log(`Invoiced Revenue: $813,012`);
  console.log(`WIP Revenue:      $60,147`);
  console.log(`Pct Invoiced:     93%`);

  console.log(`\n=== DIFFERENCES ===`);
  console.log(`Total diff:       $${(totalRevenue - 873159).toFixed(2)}`);
  console.log(`Invoiced diff:    $${(invoicedRevenue - 813012).toFixed(2)}`);
  console.log(`WIP diff:         $${(wipRevenue - 60147).toFixed(2)}`);

  // Show top WIP WOs
  const wipWos = [...woRevenue.values()].filter(w => !w.invoiced).sort((a, b) => b.revenue - a.revenue);
  console.log(`\n=== WIP WOs (${wipWos.length} total, $${wipRevenue.toFixed(2)}) ===`);
  for (const w of wipWos) {
    console.log(`  ${w.woNumber} (${w.contractor}): $${w.revenue.toFixed(2)} (${w.items} items)`);
  }

  // Show top invoiced WOs for comparison
  const invWos = [...woRevenue.values()].filter(w => w.invoiced).sort((a, b) => b.revenue - a.revenue);
  console.log(`\n=== Top 10 Invoiced WOs (of ${invWos.length}) ===`);
  for (const w of invWos.slice(0, 10)) {
    console.log(`  ${w.woNumber} (${w.contractor}): $${w.revenue.toFixed(2)} (${w.items} items)`);
  }

  process.exit(0);
}

run();
