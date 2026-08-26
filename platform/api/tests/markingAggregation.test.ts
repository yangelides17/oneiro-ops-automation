/**
 * Tests for marking item aggregation into Production Log shape.
 * Validates against Code.js aggregateMarkingItemsForPL_.
 */
import {
  aggregateForProductionLog,
  PL_CATEGORY_MAP,
  PL_FOLD_MULTIPLIER,
  type MarkingItemRow,
} from '../src/services/markingAggregation.js';

let passed = 0;
let failed = 0;

function assert(condition: boolean, label: string) {
  if (condition) { passed++; } else { failed++; console.error(`  FAIL: ${label}`); }
}

function item(overrides: Partial<MarkingItemRow> & { category: string }): MarkingItemRow {
  return {
    quantity: 0,
    unit: 'LF',
    colorMaterial: null,
    dateCompleted: '2026-08-24',
    status: 'completed',
    crewChief: '',
    ...overrides,
  };
}

// ─── Basic line aggregation ──────────────────────────────────
console.log('Basic line aggregation:');

let result = aggregateForProductionLog([
  item({ category: '4" Line', quantity: 100 }),
  item({ category: '4" Line', quantity: 50 }),
], '2026-08-24', '');

assert(result.markings['4" Lines'] === 150, '4" Line: 100 + 50 = 150');

// ─── Double White Line folds at 2× ──────────────────────────
console.log('Double White Line fold:');

result = aggregateForProductionLog([
  item({ category: 'Double White Line', quantity: 100 }),
], '2026-08-24', '');

assert(result.markings['4" Lines'] === 200, 'Double White Line 100 → 4" Lines 200 (2× fold)');

// Double Yellow Line keeps its own row (no fold)
result = aggregateForProductionLog([
  item({ category: 'Double Yellow Line', quantity: 100 }),
], '2026-08-24', '');

assert(result.markings['Double Yellow Line (Center Line)'] === 100, 'Double Yellow: own row, no fold');

// ─── HVX Crosswalk + Stop Line shared row ────────────────────
console.log('HVX + Stop Line:');

result = aggregateForProductionLog([
  item({ category: 'HVX Crosswalk', quantity: 75 }),
  item({ category: 'Stop Line', quantity: 30 }),
], '2026-08-24', '');

assert(result.markings['CrossWalks/Stop Lines'] === '75/30', 'HVX 75 + Stop 30 → "75/30"');

// Only crosswalk, no stop line
result = aggregateForProductionLog([
  item({ category: 'HVX Crosswalk', quantity: 50 }),
], '2026-08-24', '');

assert(result.markings['CrossWalks/Stop Lines'] === '50/0', 'HVX only → "50/0"');

// ─── MMA SF categories ──────────────────────────────────────
console.log('MMA SF:');

result = aggregateForProductionLog([
  item({ category: 'Bike Lane', quantity: 500, unit: 'SF', colorMaterial: 'Green' }),
  item({ category: 'Bus Lane', quantity: 300, unit: 'SF', colorMaterial: 'Red' }),
], '2026-08-24', '');

assert(result.sqft === 800, 'MMA SF: 500 + 300 = 800');
assert(result.paint === 'Green, Red', 'MMA paint: Green, Red (sorted)');
assert(Object.keys(result.markings).length === 0, 'MMA SF does NOT go into grid markings');

// ─── Pedestrian Men ──────────────────────────────────────────
console.log('Pedestrian Men:');

result = aggregateForProductionLog([
  item({ category: 'Pedestrian Men', quantity: 3, unit: 'EA' }),
], '2026-08-24', '');

assert(result.markings['PED X-ING Message'] === '3 PED MEN', 'Ped Men → "3 PED MEN"');

// ─── EA messages ─────────────────────────────────────────────
console.log('EA messages:');

result = aggregateForProductionLog([
  item({ category: 'Stop Msg', quantity: 4, unit: 'EA' }),
  item({ category: 'L/R Arrow', quantity: 6, unit: 'EA' }),
], '2026-08-24', '');

assert(result.markings['Stop Message'] === 4, 'Stop Msg → Stop Message = 4');
assert(result.markings['Left & or Right Arrows'] === 6, 'L/R Arrow → Left & or Right Arrows = 6');

// ─── Combination Arrow rollup ────────────────────────────────
console.log('Combination Arrow rollup:');

result = aggregateForProductionLog([
  item({ category: 'Combination Arrow', quantity: 2, unit: 'EA' }),
  item({ category: 'Combination Arrow (L/R)', quantity: 3, unit: 'EA' }),
], '2026-08-24', '');

assert(result.markings['Combination Arrow'] === 5, 'Combo + Combo L/R = 5 (shared row)');

// ─── Old + New Bike Symbol rollup ────────────────────────────
console.log('Bike Symbol rollup:');

result = aggregateForProductionLog([
  item({ category: 'Old Bike Symbol (w/ rider)', quantity: 2, unit: 'EA' }),
  item({ category: 'New Bike Symbol (just bike)', quantity: 1, unit: 'EA' }),
], '2026-08-24', '');

assert(result.markings['Bicycle Lane Symbol'] === 3, 'Old + New Bike Symbol = 3 (shared row)');

// ─── Date filtering ──────────────────────────────────────────
console.log('Date filtering:');

result = aggregateForProductionLog([
  item({ category: '4" Line', quantity: 100, dateCompleted: '2026-08-24' }),
  item({ category: '4" Line', quantity: 50, dateCompleted: '2026-08-23' }), // wrong date
], '2026-08-24', '');

assert(result.markings['4" Lines'] === 100, 'Only items completed on target date');

// ─── Crew chief filtering ───────────────────────────────────
console.log('Crew chief filtering:');

result = aggregateForProductionLog([
  item({ category: '4" Line', quantity: 100, crewChief: 'John' }),
  item({ category: '4" Line', quantity: 50, crewChief: 'Mike' }),
], '2026-08-24', 'John');

assert(result.markings['4" Lines'] === 100, 'Only items from crew chief John');

// Blank crew chief matches only untagged items
result = aggregateForProductionLog([
  item({ category: '4" Line', quantity: 100, crewChief: '' }),
  item({ category: '4" Line', quantity: 50, crewChief: 'John' }),
], '2026-08-24', '');

assert(result.markings['4" Lines'] === 100, 'Blank chief matches only untagged');

// ─── Status filtering ────────────────────────────────────────
console.log('Status filtering:');

result = aggregateForProductionLog([
  item({ category: '4" Line', quantity: 100, status: 'completed' }),
  item({ category: '4" Line', quantity: 50, status: 'pending' }),
], '2026-08-24', '');

assert(result.markings['4" Lines'] === 100, 'Only completed items counted');

// ─── Unmapped categories drop silently ───────────────────────
console.log('Unmapped categories:');

result = aggregateForProductionLog([
  item({ category: 'Custom Msg', quantity: 5, unit: 'EA' }),
  item({ category: 'Others', quantity: 10, unit: 'EA' }),
], '2026-08-24', '');

assert(Object.keys(result.markings).length === 0, 'Unmapped categories produce no output');

// ─── Empty input ─────────────────────────────────────────────
console.log('Empty input:');

result = aggregateForProductionLog([], '2026-08-24', '');
assert(Object.keys(result.markings).length === 0, 'Empty items → empty markings');
assert(result.sqft === '', 'Empty items → empty sqft');
assert(result.paint === '', 'Empty items → empty paint');

// ─── Verify PL_CATEGORY_MAP completeness ─────────────────────
console.log('PL_CATEGORY_MAP:');

assert(PL_CATEGORY_MAP['Lane Lines'] === 'Lane Lines 4" (Skips)', 'Lane Lines mapped');
assert(PL_CATEGORY_MAP['Shark Teeth 24x36'] === 'Sharks Teeth 24" 36"', 'Shark Teeth mapped');
assert(PL_CATEGORY_MAP['Speed Hump Markings'] === 'Speed Hump Marking', 'Speed Hump mapped');
assert(PL_FOLD_MULTIPLIER['Double White Line'] === 2, 'Double White Line fold = 2');

// ─── Summary ─────────────────────────────────────────────────
console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failed > 0) process.exit(1);
