/**
 * Seeds a new organization with NYC DOT defaults.
 * Run: npm run db:seed
 *
 * Reads from platform/seed/*.json and inserts into the org's tables.
 */
import 'dotenv/config';
import { db } from './client.js';
import {
  regions, markingCategories, pricingMultipliers,
  payClassifications, payRates, overtimeRules,
} from './schema.js';
import { readFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const seedDir = resolve(__dirname, '../../../seed');

function loadJson(filename: string) {
  return JSON.parse(readFileSync(resolve(seedDir, filename), 'utf-8'));
}

export async function seedOrg(orgId: string) {
  console.log(`Seeding org ${orgId} with NYC DOT defaults...`);

  // Regions
  const regionData = loadJson('nyc_regions.json');
  for (const r of regionData) {
    await db.insert(regions).values({ orgId, ...r }).onConflictDoNothing();
  }
  console.log(`  ${regionData.length} regions`);

  // Marking categories
  const categoryData = loadJson('nyc_dot_categories.json');
  for (const c of categoryData) {
    await db.insert(markingCategories).values({ orgId, ...c }).onConflictDoNothing();
  }
  console.log(`  ${categoryData.length} marking categories`);

  // Pricing multipliers
  const multiplierData = loadJson('nyc_dot_multipliers.json');
  for (const m of multiplierData) {
    await db.insert(pricingMultipliers).values({ orgId, ...m }).onConflictDoNothing();
  }
  console.log(`  ${multiplierData.length} pricing multipliers`);

  // Pay classifications
  const classData = loadJson('nyc_dot_classifications.json');
  for (const c of classData.classifications) {
    await db.insert(payClassifications).values({ orgId, ...c }).onConflictDoNothing();
  }
  console.log(`  ${classData.classifications.length} pay classifications`);

  // Pay rates
  for (const r of classData.rates) {
    await db.insert(payRates).values({ orgId, ...r }).onConflictDoNothing();
  }
  console.log(`  ${classData.rates.length} pay rates`);

  // Overtime rules (should already exist from signup, but ensure)
  await db.insert(overtimeRules).values({ orgId }).onConflictDoNothing();
  console.log('  overtime rules (NYC defaults)');

  console.log('Seed complete.');
}

// CLI entry point: pass org ID as argument
const orgId = process.argv[2];
if (!orgId) {
  console.error('Usage: npm run db:seed -- <org-id>');
  process.exit(1);
}
seedOrg(orgId).then(() => process.exit(0)).catch((e) => {
  console.error(e);
  process.exit(1);
});
