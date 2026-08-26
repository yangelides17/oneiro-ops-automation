/**
 * Tests for the pricing engine.
 * Validates against the exact behavior of Code.js priceMarkingItem_.
 */
import {
  rate4, money2, resolveRateRow, priceMarkingItem, resolvePricingGroup,
  LINE_WIDTH_MULTIPLIER, LINE12_MULTIPLIER, EXTRUDED_UNIT_COUNT, PREFORMED_UNIT_COUNT,
  type RateRow, type WoMeta,
} from '../src/services/pricing.js';

let passed = 0;
let failed = 0;

function assert(condition: boolean, label: string) {
  if (condition) { passed++; } else { failed++; console.error(`  FAIL: ${label}`); }
}

function assertClose(actual: number, expected: number, label: string, tolerance = 0.005) {
  assert(Math.abs(actual - expected) < tolerance, `${label} (expected ${expected}, got ${actual})`);
}

// Build standard multiplier maps from the exported constants
const multipliers = {
  lineWidth: LINE_WIDTH_MULTIPLIER,
  line12: LINE12_MULTIPLIER,
  extrudedUnit: EXTRUDED_UNIT_COUNT,
  preformedUnit: PREFORMED_UNIT_COUNT,
};

// Standard category → group mapping (mirrors NYC DOT defaults)
const categoryGroupMap: Record<string, string> = {
  '4" Line': 'line4', '6" Line': 'line4', '8" Line': 'line4',
  '12" Line': 'line4', '16" Line': 'line4', '24" Line': 'line4',
  'Lane Lines': 'line4', 'Double Yellow Line': 'line4', 'Double White Line': 'line4',
  'HVX Crosswalk': 'line12', 'Stop Line': 'line12',
  'Old Bike Symbol (w/ rider)': 'preformed', 'New Bike Symbol (just bike)': 'preformed',
  'Pedestrian Men': 'preformed', 'Bike Lane Symbol': 'preformed',
  'Stop Msg': 'extruded', 'Only Msg': 'extruded', 'Bus Msg': 'extruded',
  'Bump Msg': 'extruded', '20 MPH Msg': 'extruded',
  'Railroad (RR)': 'extruded', 'Railroad (X)': 'extruded',
  'L/R Arrow': 'extruded', 'Straight Arrow': 'extruded',
  'Combination Arrow': 'extruded', 'Combination Arrow (L/R)': 'extruded',
  'Speed Hump Markings': 'extruded', 'Shark Teeth 12x18': 'extruded',
  'Shark Teeth 24x36': 'extruded', 'Bike Lane Arrow': 'extruded',
  'Bike Lane': 'color_surface', 'Bus Lane': 'color_surface',
  'Pedestrian Space': 'color_surface', 'Bike Lane Green Bar': 'color_surface',
  'Custom Msg': 'unpriced', 'Others': 'unpriced', 'Gores': 'unpriced',
};

// Sample rate rows
const sampleRates: RateRow[] = [
  {
    contractor: 'Metro Express',
    contractNum: '84125MBTP701',
    regionCode: 'BK',
    effectiveDate: new Date('2025-07-01'),
    rates: { line4: 3.40, line12: 3.40, preformed: 95.00, extruded: 95.00, colorSurface: 3.40 },
  },
  {
    contractor: 'Metro Express',
    contractNum: '84125MBTP701',
    regionCode: 'BK',
    effectiveDate: null, // blank-date fallback
    rates: { line4: 3.00, line12: 3.00, preformed: 90.00, extruded: 90.00, colorSurface: 3.00 },
  },
  {
    contractor: 'Metro Express',
    contractNum: '84125MBTP701',
    regionCode: 'MN',
    effectiveDate: null,
    rates: { line4: 3.50, line12: 3.50, preformed: 100.00, extruded: 100.00, colorSurface: 3.50 },
  },
];

const woMeta: WoMeta = { contractor: 'Metro Express', contractNum: '84125MBTP701', regionCode: 'BK' };

// ─── rate4 ───────────────────────────────────────────────────
console.log('rate4:');

assertClose(rate4(3.40 * 1.0), 3.4, '3.40 × 1.0 = 3.4');
assertClose(rate4(3.40 * 2.0), 6.8, '3.40 × 2.0 = 6.8');
assertClose(rate4(3.40 * 1.5), 5.1, '3.40 × 1.5 = 5.1');
assertClose(rate4(95.00 * 1.35), 128.25, '95.00 × 1.35 = 128.25');
assertClose(rate4(95.00 * 0.31), 29.45, '95.00 × 0.31 = 29.45');
assertClose(rate4(0.975), 0.975, '0.975 stays 0.975');

// ─── money2 ──────────────────────────────────────────────────
console.log('money2:');

assertClose(money2(100, 3.40), 340.00, '100 × 3.40 = 340.00');
assertClose(money2(250, 6.80), 1700.00, '250 × 6.80 = 1700.00');
assertClose(money2(1, 128.25), 128.25, '1 × 128.25 = 128.25');
assertClose(money2(3, 128.25), 384.75, '3 × 128.25 = 384.75');
assertClose(money2(150, 3.40), 510.00, '150 × 3.40 = 510.00');
// Edge: very small quantities
assertClose(money2(0.05, 95.00), 4.75, '0.05 × 95.00 = 4.75');
// Edge: verify no floating-point drift
assertClose(money2(33, 3.03), 99.99, '33 × 3.03 = 99.99');

// ─── resolveRateRow ──────────────────────────────────────────
console.log('resolveRateRow:');

// Should pick the dated row (2025-07-01) when item date is after it
let row = resolveRateRow(sampleRates, 'Metro Express', '84125MBTP701', 'BK', '2025-08-15');
assert(row !== null, 'finds rate row for BK after 2025-07-01');
assertClose(row!.rates.line4!, 3.40, 'BK dated rate = 3.40');

// Should pick the dated row even without /EXT suffix
row = resolveRateRow(sampleRates, 'Metro Express', '84125MBTP701/EXT', 'BK', '2025-08-15');
assert(row !== null, 'handles /EXT suffix stripping');

// Should fall back to blank-date row when item date is before dated row
row = resolveRateRow(sampleRates, 'Metro Express', '84125MBTP701', 'BK', '2025-01-01');
assert(row !== null, 'finds blank-date fallback for early date');
assertClose(row!.rates.line4!, 3.00, 'BK blank-date rate = 3.00');

// Should pick MN rate when region is MN
row = resolveRateRow(sampleRates, 'Metro Express', '84125MBTP701', 'MN', '2025-08-15');
assert(row !== null, 'finds MN rate');
assertClose(row!.rates.line4!, 3.50, 'MN rate = 3.50');

// Should return null for unknown contractor
row = resolveRateRow(sampleRates, 'Unknown', '84125MBTP701', 'BK', '2025-08-15');
assert(row === null, 'null for unknown contractor');

// Should return null for unknown region
row = resolveRateRow(sampleRates, 'Metro Express', '84125MBTP701', 'QU', '2025-08-15');
assert(row === null, 'null for unknown region');

// Empty rates array
row = resolveRateRow([], 'Metro Express', '84125MBTP701', 'BK', '2025-08-15');
assert(row === null, 'null for empty rates');

// ─── priceMarkingItem: line4 group ───────────────────────────
console.log('priceMarkingItem — line4:');

// 4" Line: qty=100, rate=3.40, mult=1.0 → revenue=340.00
let result = priceMarkingItem(
  { category: '4" Line', quantity: 100, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === null, '4" Line prices cleanly');
assert(result.group === 'line4', '4" Line → line4 group');
assertClose(result.revenue, 340.00, '100 × 3.40 × 1.0 = 340.00');

// 8" Line: qty=200, rate=3.40, mult=2.0 → revenue=1360.00
result = priceMarkingItem(
  { category: '8" Line', quantity: 200, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 1360.00, '200 × 3.40 × 2.0 = 1360.00');

// 24" Line: qty=50, rate=3.40, mult=6.0 → revenue=1020.00
result = priceMarkingItem(
  { category: '24" Line', quantity: 50, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 1020.00, '50 × 3.40 × 6.0 = 1020.00');

// Lane Lines: qty=10 (EA), rate=3.40, mult=10 → revenue=340.00
result = priceMarkingItem(
  { category: 'Lane Lines', quantity: 10, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 340.00, '10 skips × 3.40 × 10 = 340.00');

// Double Yellow Line: qty=100, rate=3.40, mult=2.0 → revenue=680.00
result = priceMarkingItem(
  { category: 'Double Yellow Line', quantity: 100, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 680.00, '100 × 3.40 × 2.0 = 680.00');

// ─── priceMarkingItem: line12 group ──────────────────────────
console.log('priceMarkingItem — line12:');

// HVX Crosswalk: qty=75, rate=3.40, mult=1.0 → 255.00
result = priceMarkingItem(
  { category: 'HVX Crosswalk', quantity: 75, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 255.00, '75 × 3.40 × 1.0 = 255.00');

// Stop Line: qty=30, rate=3.40, mult=2.0 → 204.00
result = priceMarkingItem(
  { category: 'Stop Line', quantity: 30, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 204.00, '30 × 3.40 × 2.0 = 204.00');

// ─── priceMarkingItem: extruded group ────────────────────────
console.log('priceMarkingItem — extruded:');

// Stop Msg: qty=4, rate=95.00, unitCount=1.35 → rate4=128.25, revenue=513.00
result = priceMarkingItem(
  { category: 'Stop Msg', quantity: 4, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 513.00, '4 × 95.00 × 1.35 = 513.00');
assertClose(result.rate!, 128.25, 'Stop Msg rate = 128.25');

// Railroad (X): qty=2, rate=95.00, unitCount=0.31 → rate4=29.45, revenue=58.90
result = priceMarkingItem(
  { category: 'Railroad (X)', quantity: 2, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 58.90, '2 × 95.00 × 0.31 = 58.90');

// ─── priceMarkingItem: preformed group ───────────────────────
console.log('priceMarkingItem — preformed:');

// Old Bike Symbol: qty=3, rate=95.00, unitCount=0.91 → rate4=86.45, revenue=259.35
result = priceMarkingItem(
  { category: 'Old Bike Symbol (w/ rider)', quantity: 3, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 259.35, '3 × 95.00 × 0.91 = 259.35');

// ─── priceMarkingItem: color_surface group ───────────────────
console.log('priceMarkingItem — color_surface:');

// Bike Lane: qty=500 SF, rate=3.40 → 1700.00
result = priceMarkingItem(
  { category: 'Bike Lane', quantity: 500, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assertClose(result.revenue, 1700.00, '500 × 3.40 = 1700.00');

// ─── priceMarkingItem: error cases ───────────────────────────
console.log('priceMarkingItem — error cases:');

// Unpriced category
result = priceMarkingItem(
  { category: 'Custom Msg', quantity: 2, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'unpriced_category', 'Custom Msg → unpriced_category');
assertClose(result.revenue, 0, 'unpriced → $0');

// Bad quantity (zero)
result = priceMarkingItem(
  { category: '4" Line', quantity: 0, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'bad_qty', 'qty=0 → bad_qty');

// Bad quantity (NaN)
result = priceMarkingItem(
  { category: '4" Line', quantity: 'abc', dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'bad_qty', 'qty=abc → bad_qty');

// No date completed
result = priceMarkingItem(
  { category: '4" Line', quantity: 100, dateCompleted: null },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'no_date', 'no date → no_date');

// No matching rate row
result = priceMarkingItem(
  { category: '4" Line', quantity: 100, dateCompleted: '2025-08-15' },
  { contractor: 'Unknown', contractNum: '99999', regionCode: 'BK' },
  sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'no_rate', 'unknown contractor → no_rate');

// Empty category
result = priceMarkingItem(
  { category: '', quantity: 100, dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'unpriced_category', 'empty category → unpriced_category');

// Bike Lane Green Bar as EA → unit_migration
result = priceMarkingItem(
  { category: 'Bike Lane Green Bar', quantity: 5, unit: 'EA', dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === 'unit_migration', 'Bike Lane Green Bar EA → unit_migration');

// Bike Lane Green Bar as SF → prices normally
result = priceMarkingItem(
  { category: 'Bike Lane Green Bar', quantity: 500, unit: 'SF', dateCompleted: '2025-08-15' },
  woMeta, sampleRates, categoryGroupMap, multipliers,
);
assert(result.reason === null, 'Bike Lane Green Bar SF → prices cleanly');
assertClose(result.revenue, 1700.00, '500 × 3.40 = 1700.00');

// ─── Summary ─────────────────────────────────────────────────
console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failed > 0) process.exit(1);
