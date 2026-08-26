import { Router } from 'express';
import { eq } from 'drizzle-orm';
import { db } from '../db/client.js';
import { getDashboardData } from '../db/queries/workOrders.js';
import { getPendingCounts } from '../db/queries/documents.js';
import { listContractPricing, listMultipliers, listCategories, listBillingRemaps } from '../db/queries/settings.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { getRevenueData } from '../services/revenue.js';
import { getProductionData } from '../services/production.js';
import { contractors } from '../db/schema.js';
import {
  LINE_WIDTH_MULTIPLIER, LINE12_MULTIPLIER,
  EXTRUDED_UNIT_COUNT, PREFORMED_UNIT_COUNT, type RateRow,
} from '../services/pricing.js';
import type { RemapRule } from '../services/billingRemap.js';
import { statusToDisplay } from '../utils/statusFormat.js';

const router = Router();

/** Build rate rows and multiplier maps from tenant's DB config. */
async function loadPricingConfig(orgId: string) {
  const [pricingRows, multiplierRows, categoryRows, contractorRows] = await Promise.all([
    listContractPricing(db, orgId),
    listMultipliers(db, orgId),
    listCategories(db, orgId),
    db.select({ id: contractors.id, name: contractors.name }).from(contractors).where(eq(contractors.orgId, orgId)),
  ]);

  const contractorNameMap = new Map(contractorRows.map(c => [c.id, c.name]));

  // Build rate rows
  const rates: RateRow[] = pricingRows.map(r => ({
    contractor: contractorNameMap.get(r.contractorId) || '',
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

  // Build category → pricing group map
  const categoryGroupMap: Record<string, string> = {};
  for (const cat of categoryRows) {
    if (cat.pricingGroup) categoryGroupMap[cat.name] = cat.pricingGroup;
  }

  // Build multiplier maps from DB, falling back to NYC DOT defaults
  const lineWidth = { ...LINE_WIDTH_MULTIPLIER };
  const line12 = { ...LINE12_MULTIPLIER };
  const extrudedUnit = { ...EXTRUDED_UNIT_COUNT };
  const preformedUnit = { ...PREFORMED_UNIT_COUNT };

  for (const m of multiplierRows) {
    const val = Number(m.value);
    switch (m.multiplierType) {
      case 'line_width': lineWidth[m.categoryName] = val; break;
      case 'line12': line12[m.categoryName] = val; break;
      case 'extruded_unit': extrudedUnit[m.categoryName] = val; break;
      case 'preformed_unit': preformedUnit[m.categoryName] = val; break;
    }
  }

  // Load billing remap rules
  const remapRows = await listBillingRemaps(db, orgId);
  const remapRules: RemapRule[] = remapRows.map(r => ({
    sourceContract: r.sourceContract,
    sourceRegion: r.sourceRegion,
    sourceContractor: r.sourceContractor,
    targetContract: r.targetContract,
    targetRegion: r.targetRegion,
    effectiveDate: r.effectiveDate,
  }));

  return { rates, categoryGroupMap, multipliers: { lineWidth, line12, extrudedUnit, preformedUnit }, remapRules };
}

/** GET /api/dashboard — WO list + stats. */
router.get('/dashboard', requireRole('owner', 'admin'), async (req, res) => {
  const data = await getDashboardData(db, getOrgId(req));
  // Normalize status to Title Case for frontend
  data.wos = data.wos.map((wo: any) => ({ ...wo, status: statusToDisplay(wo.status) }));
  if (data.attention) {
    data.attention = data.attention.map((wo: any) => ({ ...wo, status: statusToDisplay(wo.status) }));
  }
  res.json(data);
});

/** GET /api/revenue — Revenue data for a date range. */
router.get('/revenue', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  // L-11: Use org timezone for default date range, not UTC
  const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
  const start = (req.query.start as string) || today.slice(0, 8) + '01';
  const end = (req.query.end as string) || today;

  const { rates, categoryGroupMap, multipliers, remapRules } = await loadPricingConfig(orgId);
  const data = await getRevenueData(db, orgId, { startDate: start, endDate: end }, rates, categoryGroupMap, multipliers, remapRules);

  // ── byGroup: zero-fill all 5 pricing groups (M-19) ──────────
  const ALL_GROUPS = ['line4', 'line12', 'preformed', 'extruded', 'colorSurface'];
  const groupMap = new Map(Object.entries(data.totals.byGroup || {}));
  const byGroup = ALL_GROUPS.map(group => ({
    group,
    revenue: Number(groupMap.get(group) || 0),
    items: 0, // not computed per-group currently
  }));

  const buckets = [
    { key: 'thermo', label: 'Thermo', revenue: (data.totals.byBucket?.thermo || 0) },
    { key: 'mma', label: 'MMA', revenue: (data.totals.byBucket?.mma || 0) },
    { key: 'preform', label: 'Preform', revenue: (data.totals.byBucket?.preform || 0) },
  ];

  // ── top_wos: top 25 WOs by revenue (DASH-2) ────────────────
  const topWos = (data.byWo || [])
    .sort((a: any, b: any) => (b.revenue || 0) - (a.revenue || 0))
    .slice(0, 25)
    .map((w: any) => ({
      wo_id: w.woNumber,         // old-app field name
      woNumber: w.woNumber,      // new-app field name
      contractor: w.contractorName,
      contractorName: w.contractorName,
      location: w.location || '',
      revenue: w.revenue || 0,
      items: w.items || 0,
    }));

  res.json({
    range: data.range,
    totals: {
      revenue: data.totals.revenue,
      items: data.totals.items,
      needsPricing: data.totals.needsPricing,
      invoicedRevenue: data.totals.invoicedRevenue,
      wipRevenue: data.totals.wipRevenue,
      pctInvoiced: data.totals.pctInvoiced,
      pct_invoiced: data.totals.pctInvoiced,  // snake_case alias (M-22)
    },
    buckets,
    daily: data.daily,
    byContractor: data.byContractor.map((c: any) => ({
      contractor: c.contractorName,      // old-app field name (M-21)
      contractorName: c.contractorName,  // new-app field name
      revenue: c.revenue,
      items: c.items || 0,
    })),
    byGroup,
    needsPricingItems: data.needsPricingItems,
    needs_pricing: data.needsPricingItems,  // React also reads this alias
    top_wos: topWos,
    laborDaily: data.laborDaily,
    laborTotals: data.laborTotals,
  });
});

/** GET /api/production — Production data for a date range. */
router.get('/production', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  const todayProd = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
  const start = (req.query.start as string) || todayProd.slice(0, 8) + '01';
  const end = (req.query.end as string) || todayProd;

  const data = await getProductionData(db, orgId, { startDate: start, endDate: end });
  res.json(data);
});

/** GET /api/pending-counts — Nav badge counts (all roles). */
router.get('/pending-counts', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const counts = await getPendingCounts(db, getOrgId(req));
  res.json(counts);
});

/** GET /api/pending-counts/doc-status — Doc status pending count. */
router.get('/pending-counts/doc-status', requireRole('owner', 'admin'), async (req, res) => {
  const counts = await getPendingCounts(db, getOrgId(req));
  res.json({ doc_status_pending: counts.doc_status_pending });
});

export default router;
