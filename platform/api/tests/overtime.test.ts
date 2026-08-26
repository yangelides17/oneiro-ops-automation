/**
 * Tests for the overtime calculation service.
 * Validates against the exact behavior of Code.js _allocateDayOvertime_.
 */
import {
  allocateDayOvertime,
  computeRowHours,
  splitStOt,
  isWeekendDate,
  parseTimeOfDay,
  to12h,
  to24h,
  normalizeEmployeeName,
  type OvertimeRules,
} from '../src/services/overtime.js';

const NYC_RULES: OvertimeRules = {
  dailyThresholdHours: 8,
  weeklyThresholdHours: null,
  weekendAllOt: true,
  crossGroupLookback: true,
};

let passed = 0;
let failed = 0;

function assert(condition: boolean, label: string) {
  if (condition) {
    passed++;
  } else {
    failed++;
    console.error(`  FAIL: ${label}`);
  }
}

function assertClose(actual: number, expected: number, label: string, tolerance = 0.01) {
  assert(Math.abs(actual - expected) < tolerance, `${label} (expected ${expected}, got ${actual})`);
}

// ─── parseTimeOfDay ──────────────────────────────────────────
console.log('parseTimeOfDay:');

assert(parseTimeOfDay('7:00 AM')?.hours === 7, '7:00 AM → 7h');
assert(parseTimeOfDay('7:00 AM')?.minutes === 0, '7:00 AM → 0m');
assert(parseTimeOfDay('12:00 PM')?.hours === 12, '12:00 PM → 12h');
assert(parseTimeOfDay('12:00 AM')?.hours === 0, '12:00 AM → 0h');
assert(parseTimeOfDay('11:30 PM')?.hours === 23, '11:30 PM → 23h');
assert(parseTimeOfDay('11:30 PM')?.minutes === 30, '11:30 PM → 30m');
assert(parseTimeOfDay('07:00')?.hours === 7, '07:00 (24h) → 7h');
assert(parseTimeOfDay('23:45')?.hours === 23, '23:45 (24h) → 23h');
assert(parseTimeOfDay('') === null, 'empty → null');
assert(parseTimeOfDay('garbage') === null, 'garbage → null');

// ─── computeRowHours ─────────────────────────────────────────
console.log('computeRowHours:');

assertClose(computeRowHours('7:00 AM', '3:30 PM'), 8.5, '7AM–3:30PM = 8.5h');
assertClose(computeRowHours('07:00', '15:30'), 8.5, '07:00–15:30 = 8.5h');
assertClose(computeRowHours('9:00 PM', '5:00 AM'), 8, '9PM–5AM cross-midnight = 8h');
assertClose(computeRowHours('11:00 PM', '2:00 AM'), 3, '11PM–2AM cross-midnight = 3h');
assertClose(computeRowHours('6:00 AM', '6:00 AM'), 24, '6AM–6AM = 24h (full day)');
assert(computeRowHours('', '3:00 PM') === 0, 'missing timeIn = 0');
assert(computeRowHours('7:00 AM', '') === 0, 'missing timeOut = 0');

// ─── isWeekendDate ───────────────────────────────────────────
console.log('isWeekendDate:');

// 2026-08-24 is a Monday
assert(isWeekendDate('2026-08-24') === false, '2026-08-24 (Mon) → false');
// 2026-08-22 is a Saturday
assert(isWeekendDate('2026-08-22') === true, '2026-08-22 (Sat) → true');
// 2026-08-23 is a Sunday
assert(isWeekendDate('2026-08-23') === true, '2026-08-23 (Sun) → true');
assert(isWeekendDate('') === false, 'empty → false');

// ─── allocateDayOvertime: basic weekday ──────────────────────
console.log('allocateDayOvertime — weekday basics:');

// Single employee, 8h → 0 OT
let result = allocateDayOvertime(
  [{ key: 'john', hours: 8 }],
  false, {}, NYC_RULES,
);
assertClose(result[0], 0, '8h weekday → 0 OT');

// Single employee, 10h → 2 OT
result = allocateDayOvertime(
  [{ key: 'john', hours: 10 }],
  false, {}, NYC_RULES,
);
assertClose(result[0], 2, '10h weekday → 2 OT');

// Single employee, 6h → 0 OT
result = allocateDayOvertime(
  [{ key: 'john', hours: 6 }],
  false, {}, NYC_RULES,
);
assertClose(result[0], 0, '6h weekday → 0 OT');

// ─── allocateDayOvertime: weekend ────────────────────────────
console.log('allocateDayOvertime — weekend:');

// Weekend: 8h → ALL OT
result = allocateDayOvertime(
  [{ key: 'john', hours: 8 }],
  true, {}, NYC_RULES,
);
assertClose(result[0], 8, '8h weekend → 8 OT');

// Weekend: 3h → ALL OT
result = allocateDayOvertime(
  [{ key: 'john', hours: 3 }],
  true, {}, NYC_RULES,
);
assertClose(result[0], 3, '3h weekend → 3 OT');

// ─── allocateDayOvertime: multiple entries same employee ─────
console.log('allocateDayOvertime — multiple entries same employee:');

// Two entries: 6h + 4h = 10h total, second entry gets 2h OT
result = allocateDayOvertime(
  [
    { key: 'john', hours: 6 },
    { key: 'john', hours: 4 },
  ],
  false, {}, NYC_RULES,
);
assertClose(result[0], 0, 'first 6h → 0 OT');
assertClose(result[1], 2, 'next 4h → 2 OT (total 10, threshold 8)');

// Three entries: 4h + 4h + 2h
result = allocateDayOvertime(
  [
    { key: 'john', hours: 4 },
    { key: 'john', hours: 4 },
    { key: 'john', hours: 2 },
  ],
  false, {}, NYC_RULES,
);
assertClose(result[0], 0, '4h → 0 OT');
assertClose(result[1], 0, '4h → 0 OT (total 8, at threshold)');
assertClose(result[2], 2, '2h → 2 OT (total 10, over threshold)');

// ─── allocateDayOvertime: cross-group lookback ───────────────
console.log('allocateDayOvertime — cross-group lookback:');

// Prior: John already worked 5h on an earlier sign-in group (LP sheet)
// Current: John works 4h on a different group (SAT sheet)
// Total = 9h, threshold = 8h → 1h OT on the current entry
result = allocateDayOvertime(
  [{ key: 'john', hours: 4 }],
  false, { john: 5 }, NYC_RULES,
);
assertClose(result[0], 1, 'prior 5h + current 4h → 1 OT');

// Prior: 8h already worked → ALL current hours are OT
result = allocateDayOvertime(
  [{ key: 'john', hours: 3 }],
  false, { john: 8 }, NYC_RULES,
);
assertClose(result[0], 3, 'prior 8h + current 3h → 3 OT');

// Prior: 0h → same as no prior
result = allocateDayOvertime(
  [{ key: 'john', hours: 6 }],
  false, { john: 0 }, NYC_RULES,
);
assertClose(result[0], 0, 'prior 0h + current 6h → 0 OT');

// ─── allocateDayOvertime: multiple employees ─────────────────
console.log('allocateDayOvertime — multiple employees:');

// Two different employees, each under 8h
result = allocateDayOvertime(
  [
    { key: 'john', hours: 7 },
    { key: 'mike', hours: 9 },
  ],
  false, {}, NYC_RULES,
);
assertClose(result[0], 0, 'john 7h → 0 OT');
assertClose(result[1], 1, 'mike 9h → 1 OT');

// Two employees with prior hours from cross-group
result = allocateDayOvertime(
  [
    { key: 'john', hours: 4 },
    { key: 'mike', hours: 4 },
  ],
  false, { john: 6, mike: 2 }, NYC_RULES,
);
assertClose(result[0], 2, 'john: prior 6 + 4 → 2 OT');
assertClose(result[1], 0, 'mike: prior 2 + 4 → 0 OT');

// ─── allocateDayOvertime: edge cases ─────────────────────────
console.log('allocateDayOvertime — edge cases:');

// Zero hours
result = allocateDayOvertime(
  [{ key: 'john', hours: 0 }],
  false, {}, NYC_RULES,
);
assertClose(result[0], 0, '0h → 0 OT');

// Empty entries array
result = allocateDayOvertime([], false, {}, NYC_RULES);
assert(result.length === 0, 'empty entries → empty result');

// No daily threshold (all ST)
const noThresholdRules: OvertimeRules = { ...NYC_RULES, dailyThresholdHours: null };
result = allocateDayOvertime(
  [{ key: 'john', hours: 20 }],
  false, {}, noThresholdRules,
);
assertClose(result[0], 0, '20h with no threshold → 0 OT');

// Weekend overrides daily threshold
result = allocateDayOvertime(
  [{ key: 'john', hours: 4 }],
  true, {}, NYC_RULES,
);
assertClose(result[0], 4, '4h weekend → 4 OT (weekend overrides)');

// ─── splitStOt ───────────────────────────────────────────────
console.log('splitStOt:');

let split = splitStOt(10, false, NYC_RULES);
assertClose(split.st, 8, '10h weekday → 8 ST');
assertClose(split.ot, 2, '10h weekday → 2 OT');

split = splitStOt(6, false, NYC_RULES);
assertClose(split.st, 6, '6h weekday → 6 ST');
assertClose(split.ot, 0, '6h weekday → 0 OT');

split = splitStOt(10, true, NYC_RULES);
assertClose(split.st, 0, '10h weekend → 0 ST');
assertClose(split.ot, 10, '10h weekend → 10 OT');

// ─── to12h / to24h ──────────────────────────────────────────
console.log('to12h / to24h:');

assert(to24h('7:00 AM') === '07:00', 'to24h: 7:00 AM → 07:00');
assert(to24h('11:30 PM') === '23:30', 'to24h: 11:30 PM → 23:30');
assert(to24h('12:00 AM') === '00:00', 'to24h: 12:00 AM → 00:00');
assert(to24h('12:00 PM') === '12:00', 'to24h: 12:00 PM → 12:00');
assert(to24h('07:00') === '07:00', 'to24h: 07:00 → 07:00 (passthrough)');

assert(to12h('07:00') === '7:00 AM', 'to12h: 07:00 → 7:00 AM');
assert(to12h('23:30') === '11:30 PM', 'to12h: 23:30 → 11:30 PM');
assert(to12h('00:00') === '12:00 AM', 'to12h: 00:00 → 12:00 AM');
assert(to12h('12:00') === '12:00 PM', 'to12h: 12:00 → 12:00 PM');

// ─── normalizeEmployeeName ───────────────────────────────────
console.log('normalizeEmployeeName:');

assert(normalizeEmployeeName('  John  Smith  ') === 'john smith', 'trims + lowercases + collapses spaces');
assert(normalizeEmployeeName('JANE DOE') === 'jane doe', 'all caps → lowercase');
assert(normalizeEmployeeName('') === '', 'empty → empty');

// ─── Summary ─────────────────────────────────────────────────
console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failed > 0) process.exit(1);
