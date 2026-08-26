/**
 * Diagnostic: compare pricing between old and new system.
 * Usage: npx tsx src/db/diagnose_pricing.ts
 */
import 'dotenv/config';
import { db } from './client.js';
import { eq, and, between, asc } from 'drizzle-orm';
import { markingItems, workOrders, contractors } from './schema.js';
import { priceMarkingItem, type RateRow, LINE_WIDTH_MULTIPLIER, LINE12_MULTIPLIER, EXTRUDED_UNIT_COUNT, PREFORMED_UNIT_COUNT } from '../services/pricing.js';
import { billingRemap, type RemapRule } from '../services/billingRemap.js';
import { listContractPricing, listMultipliers, listCategories, listBillingRemaps } from './queries/settings.js';

const ORG = '3520b16b-892c-490c-a60f-30defa7e15f2';

async function diagnose() {
  // Load pricing config
  const pricingRows = await listContractPricing(db, ORG);
  const cRows = await db.select({ id: contractors.id, name: contractors.name }).from(contractors).where(eq(contractors.orgId, ORG));
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
  const catGroupMap: Record<string, string> = {};
  for (const c of catRows) { if (c.pricingGroup) catGroupMap[c.name] = c.pricingGroup; }

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

  console.log(`Rates: ${rates.length} rows`);
  console.log(`Categories with pricing group: ${Object.keys(catGroupMap).length}`);
  console.log(`Remap rules: ${remapRules.length}`);

  // Load all completed items in range
  const items = await db.select({
    category: markingItems.category,
    quantity: markingItems.quantity,
    unit: markingItems.unit,
    dateCompleted: markingItems.dateCompleted,
    contractNum: workOrders.contractNum,
    regionCode: workOrders.regionCode,
    contractorId: workOrders.contractorId,
    woNumber: workOrders.woNumber,
  })
  .from(markingItems)
  .innerJoin(workOrders, eq(markingItems.woId, workOrders.id))
  .where(and(
    eq(markingItems.orgId, ORG),
    eq(markingItems.status, 'completed'),
    between(markingItems.dateCompleted, '2026-01-01', '2026-08-31'),
  ));

  console.log(`\nCompleted items in range: ${items.length}`);

  // Price each item
  let totalRevenue = 0;
  let priced = 0;
  let unpriced = 0;
  const reasonCounts: Record<string, number> = {};
  const reasonExamples: Record<string, string[]> = {};
  const revenueByGroup: Record<string, number> = {};

  for (const item of items) {
    const contractorName = cMap.get(item.contractorId) || '';

    // Strip /EXT before remap (matching old Code.js _buildRevenuePayload_ behavior)
    const cn = String(item.contractNum || '').split('/')[0].trim();
    const billed = billingRemap(remapRules, cn, item.regionCode || '', contractorName);

    const result = priceMarkingItem(
      { category: item.category, quantity: item.quantity, unit: item.unit, dateCompleted: item.dateCompleted },
      { contractor: contractorName, contractNum: billed.contractNum, regionCode: billed.regionCode },
      rates, catGroupMap, muls,
    );

    if (result.reason) {
      unpriced++;
      reasonCounts[result.reason] = (reasonCounts[result.reason] || 0) + 1;
      if (!reasonExamples[result.reason]) reasonExamples[result.reason] = [];
      if (reasonExamples[result.reason].length < 3) {
        reasonExamples[result.reason].push(
          `${item.woNumber} ${item.category} qty=${item.quantity} contractor=${contractorName} cn=${cn} region=${item.regionCode} billed=${billed.contractNum}/${billed.regionCode}`
        );
      }
    } else {
      priced++;
      totalRevenue += result.revenue;
      revenueByGroup[result.group] = (revenueByGroup[result.group] || 0) + result.revenue;
    }
  }

  console.log(`\nPriced: ${priced} items → $${totalRevenue.toFixed(2)}`);
  console.log(`Unpriced: ${unpriced} items`);
  console.log(`\nRevenue by group:`);
  for (const [g, v] of Object.entries(revenueByGroup)) {
    console.log(`  ${g}: $${v.toFixed(2)}`);
  }
  console.log(`\nUnpriced reasons:`);
  for (const [reason, count] of Object.entries(reasonCounts)) {
    console.log(`  ${reason}: ${count} items`);
    for (const ex of (reasonExamples[reason] || [])) {
      console.log(`    e.g. ${ex}`);
    }
  }

  process.exit(0);
}

diagnose();
