/**
 * Tests for billing remap service.
 * Validates against Code.js _billingRemap_, _billingRemapAsOf_, _billingRemapForMonth_, _hasBillingRemap_.
 */
import {
  billingRemap, billingRemapAsOf, billingRemapForMonth, hasBillingRemap,
  type RemapRule,
} from '../src/services/billingRemap.js';

let passed = 0;
let failed = 0;

function assert(condition: boolean, label: string) {
  if (condition) { passed++; } else { failed++; console.error(`  FAIL: ${label}`); }
}

// Replicate the exact rules from Code.js _BILLING_REMAP_
const rules: RemapRule[] = [
  {
    sourceContract: '84125MBTP701', sourceRegion: 'BK', sourceContractor: 'Metro Express',
    targetContract: '84125MBTP701', targetRegion: 'M',
    effectiveDate: '2026-07-05',
  },
  {
    sourceContract: '84125MBTP701', sourceRegion: 'BK', sourceContractor: 'Denville',
    targetContract: '84125MBTP701', targetRegion: 'QU',
    effectiveDate: '2026-07-05',
  },
];

// ─── billingRemap (always-on) ────────────────────────────────
console.log('billingRemap:');

// Metro Express BK → M
let r = billingRemap(rules, '84125MBTP701', 'BK', 'Metro Express');
assert(r.contractNum === '84125MBTP701', 'Metro BK: contract unchanged');
assert(r.regionCode === 'M', 'Metro BK → M');

// Denville BK → QU
r = billingRemap(rules, '84125MBTP701', 'BK', 'Denville');
assert(r.contractNum === '84125MBTP701', 'Denville BK: contract unchanged');
assert(r.regionCode === 'QU', 'Denville BK → QU');

// Unknown contractor → identity (no remap)
r = billingRemap(rules, '84125MBTP701', 'BK', 'Unknown Co');
assert(r.contractNum === '84125MBTP701', 'Unknown: contract identity');
assert(r.regionCode === 'BK', 'Unknown: region identity (BK stays BK)');

// Different contract → identity
r = billingRemap(rules, '99999', 'MN', 'Metro Express');
assert(r.regionCode === 'MN', 'Different contract: region identity');

// Empty rules → identity
r = billingRemap([], '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'BK', 'Empty rules: region identity');

// ─── billingRemapAsOf (date-gated) ──────────────────────────
console.log('billingRemapAsOf:');

// Before cutover (2026-07-04) → raw identity
r = billingRemapAsOf(rules, '2026-07-04', '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'BK', 'Before cutover: BK stays BK');

// On cutover (2026-07-05) → remapped
r = billingRemapAsOf(rules, '2026-07-05', '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'M', 'On cutover: BK → M');

// After cutover (2026-08-15) → remapped
r = billingRemapAsOf(rules, '2026-08-15', '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'M', 'After cutover: BK → M');

// Null date → raw identity
r = billingRemapAsOf(rules, null, '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'BK', 'Null date: BK stays BK');

// /EXT suffix stripped
r = billingRemapAsOf(rules, '2026-08-15', '84125MBTP701/EXT', 'BK', 'Metro Express');
assert(r.contractNum === '84125MBTP701', '/EXT stripped');
assert(r.regionCode === 'M', '/EXT stripped: BK → M');

// Denville post-cutover → QU
r = billingRemapAsOf(rules, '2026-08-15', '84125MBTP701', 'BK', 'Denville');
assert(r.regionCode === 'QU', 'Denville post-cutover: BK → QU');

// ─── billingRemapForMonth (month-granularity) ────────────────
console.log('billingRemapForMonth:');

// Month before cutover month (2026-06) → raw
r = billingRemapForMonth(rules, '2026-06', '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'BK', 'June 2026: BK stays BK');

// Cutover month (2026-07) → entire month remapped
r = billingRemapForMonth(rules, '2026-07', '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'M', 'July 2026 (cutover month): BK → M');

// Month after cutover (2026-08) → remapped
r = billingRemapForMonth(rules, '2026-08', '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'M', 'Aug 2026: BK → M');

// Null month → raw
r = billingRemapForMonth(rules, null, '84125MBTP701', 'BK', 'Metro Express');
assert(r.regionCode === 'BK', 'Null month: BK stays BK');

// ─── hasBillingRemap ─────────────────────────────────────────
console.log('hasBillingRemap:');

assert(hasBillingRemap(rules, '84125MBTP701', 'BK') === true, '701 BK has remap');
assert(hasBillingRemap(rules, '84125MBTP701', 'MN') === false, '701 MN has no remap');
assert(hasBillingRemap(rules, '99999', 'BK') === false, '99999 BK has no remap');
assert(hasBillingRemap([], '84125MBTP701', 'BK') === false, 'Empty rules: no remap');

// ─── Summary ─────────────────────────────────────────────────
console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failed > 0) process.exit(1);
