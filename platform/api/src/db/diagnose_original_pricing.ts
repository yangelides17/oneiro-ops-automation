/**
 * Diagnose revenue using ONLY the original Excel pricing (no extra rows we added).
 * This should match the old app's numbers exactly.
 */
import 'dotenv/config';
import { db } from './client.js';
import { eq, and, between } from 'drizzle-orm';
import { markingItems, workOrders, contractors, invoices } from './schema.js';
import {
  priceMarkingItem, resolveRateRow, LINE_WIDTH_MULTIPLIER, LINE12_MULTIPLIER,
  EXTRUDED_UNIT_COUNT, PREFORMED_UNIT_COUNT, type RateRow,
} from '../services/pricing.js';
import { billingRemap, type RemapRule } from '../services/billingRemap.js';
import { listCategories, listMultipliers, listBillingRemaps } from './queries/settings.js';

const ORG = '3520b16b-892c-490c-a60f-30defa7e15f2';

async function run() {
  const cRows = await db.select({ id: contractors.id, name: contractors.name })
    .from(contractors).where(eq(contractors.orgId, ORG));
  const cMap = new Map(cRows.map(c => [c.id, c.name]));

  // ORIGINAL pricing from Excel only — exclude the rows we added
  const originalRates: RateRow[] = [
    { contractor: 'Denville', contractNum: '84125MBTP701', regionCode: 'BX', effectiveDate: null,
      rates: { line4: 0.65, line12: 1.4, preformed: 120, extruded: 80, colorSurface: 3.5 } },
    { contractor: 'Metro Express', contractNum: '84125MBTP701', regionCode: 'M', effectiveDate: null,
      rates: { line4: 0.7, line12: 1.45, preformed: 120, extruded: 90, colorSurface: 3.4 } },
    { contractor: 'Metro Express', contractNum: '84123MBTP564', regionCode: 'M', effectiveDate: null,
      rates: { line4: null, line12: null, preformed: null, extruded: null, colorSurface: 3.4 } },
    { contractor: 'Metro Express', contractNum: '84123MBTP564', regionCode: 'BK', effectiveDate: null,
      rates: { line4: null, line12: null, preformed: null, extruded: null, colorSurface: 3.4 } },
    { contractor: 'Metro Express', contractNum: '84123MBTP564', regionCode: 'SI', effectiveDate: null,
      rates: { line4: null, line12: null, preformed: null, extruded: null, colorSurface: 3.4 } },
    { contractor: 'Denville', contractNum: '84125MBTP701', regionCode: 'QU', effectiveDate: null,
      rates: { line4: 0.65, line12: 1.4, preformed: 120, extruded: 80, colorSurface: 3.5 } },
    { contractor: 'Delan', contractNum: '84125MBTP701', regionCode: 'SI', effectiveDate: null,
      rates: { line4: 0.72, line12: 1.55, preformed: 125, extruded: 90, colorSurface: 3.65 } },
    { contractor: 'Metro Express', contractNum: '84123MBTP599', regionCode: 'BK', effectiveDate: null,
      rates: { line4: null, line12: null, preformed: null, extruded: null, colorSurface: null } },
  ];

  const catRows = await listCategories(db, ORG);
  const catMap: Record<string, string> = {};
  for (const c of catRows) if (c.pricingGroup) catMap[c.name] = c.pricingGroup;

  const mulRows = await listMultipliers(db, ORG);
  const muls = {
    lineWidth: { ...LINE_WIDTH_MULTIPLIER }, line12: { ...LINE12_MULTIPLIER },
    extrudedUnit: { ...EXTRUDED_UNIT_COUNT }, preformedUnit: { ...PREFORMED_UNIT_COUNT },
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
    sourceContract: r.sourceContract, sourceRegion: r.sourceRegion, sourceContractor: r.sourceContractor,
    targetContract: r.targetContract, targetRegion: r.targetRegion, effectiveDate: r.effectiveDate,
  }));

  const invRows = await db.select({ woId: invoices.woId }).from(invoices).where(eq(invoices.orgId, ORG));
  const invoicedWoIds = new Set(invRows.map(i => i.woId).filter(Boolean));

  const items = await db.select({
    category: markingItems.category, quantity: markingItems.quantity, unit: markingItems.unit,
    dateCompleted: markingItems.dateCompleted, woId: markingItems.woId, woNumber: workOrders.woNumber,
    contractorId: workOrders.contractorId, contractNum: workOrders.contractNum, regionCode: workOrders.regionCode,
  }).from(markingItems).innerJoin(workOrders, eq(markingItems.woId, workOrders.id))
    .where(and(eq(markingItems.orgId, ORG), eq(markingItems.status, 'completed'), between(markingItems.dateCompleted, '2026-01-01', '2026-08-31')));

  let total = 0, inv = 0, wip = 0, unpriced = 0;
  const reasonCounts: Record<string, number> = {};

  for (const item of items) {
    const contractorName = cMap.get(item.contractorId) || '';
    const cn = String(item.contractNum || '').split('/')[0].trim();
    const billed = billingRemap(remapRules, cn, item.regionCode || '', contractorName);
    const result = priceMarkingItem(
      { category: item.category, quantity: item.quantity, unit: item.unit, dateCompleted: item.dateCompleted },
      { contractor: contractorName, contractNum: billed.contractNum, regionCode: billed.regionCode },
      originalRates, catMap, muls,
    );

    if (result.reason) {
      unpriced++;
      reasonCounts[result.reason] = (reasonCounts[result.reason] || 0) + 1;
      continue;
    }
    total += result.revenue;
    if (invoicedWoIds.has(item.woId)) inv += result.revenue;
    else wip += result.revenue;
  }

  console.log(`\n=== WITH ORIGINAL EXCEL PRICING (no extra rows) ===`);
  console.log(`Total:    $${total.toFixed(2)}`);
  console.log(`Invoiced: $${inv.toFixed(2)}`);
  console.log(`WIP:      $${wip.toFixed(2)}`);
  console.log(`Pct:      ${(total > 0 ? inv/total*100 : 0).toFixed(1)}%`);
  console.log(`Unpriced: ${unpriced} items`);
  console.log(`Reasons:  ${JSON.stringify(reasonCounts)}`);
  console.log(`\n=== OLD APP (expected) ===`);
  console.log(`Total:    $863,418`);
  console.log(`Invoiced: $843,457`);
  console.log(`WIP:      $29,702 (? - user to confirm)`);
  console.log(`Pct:      97%`);

  process.exit(0);
}
run();
