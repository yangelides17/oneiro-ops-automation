/**
 * Tests for operational day calculation.
 * Validates against Code.js opDay_() and webapp/src/lib/dateOps.js opDay().
 */
import { opDay, opDayFromIsoTime, isWeekendDate } from '../src/services/overtime.js';
import { opDay as opDayService, opDayFromIsoTime as opDayFromIsoTimeService } from '../src/services/opDay.js';

let passed = 0;
let failed = 0;

function assert(condition: boolean, label: string) {
  if (condition) { passed++; } else { failed++; console.error(`  FAIL: ${label}`); }
}

// ─── opDay (from opDay.ts service) ───────────────────────────
console.log('opDay service:');

// 3 AM with 5 AM cutoff → buckets to PREVIOUS day
const aug24_3am = new Date(2026, 7, 24, 3, 0, 0); // Aug 24 at 3 AM
assert(opDayService(aug24_3am, 5) === '2026-08-23', '3 AM → previous day (cutoff 5)');

// 5 AM with 5 AM cutoff → stays on current day (not strictly before)
const aug24_5am = new Date(2026, 7, 24, 5, 0, 0);
assert(opDayService(aug24_5am, 5) === '2026-08-24', '5 AM → current day (cutoff 5)');

// 6 AM → current day
const aug24_6am = new Date(2026, 7, 24, 6, 0, 0);
assert(opDayService(aug24_6am, 5) === '2026-08-24', '6 AM → current day');

// 11 PM → current day
const aug24_11pm = new Date(2026, 7, 24, 23, 0, 0);
assert(opDayService(aug24_11pm, 5) === '2026-08-24', '11 PM → current day');

// Midnight (0 AM) → previous day
const aug24_midnight = new Date(2026, 7, 24, 0, 0, 0);
assert(opDayService(aug24_midnight, 5) === '2026-08-23', 'midnight → previous day');

// 4:59 AM → previous day
const aug24_459am = new Date(2026, 7, 24, 4, 59, 0);
assert(opDayService(aug24_459am, 5) === '2026-08-23', '4:59 AM → previous day');

// Cutoff = 0 → no rollback ever
const aug24_midnight_cut0 = new Date(2026, 7, 24, 0, 0, 0);
assert(opDayService(aug24_midnight_cut0, 0) === '2026-08-24', 'midnight with cutoff=0 → current day');

// ─── opDayFromIsoTime ────────────────────────────────────────
console.log('opDayFromIsoTime:');

assert(opDayFromIsoTimeService('2026-08-24', '03:00', 5) === '2026-08-23', 'ISO 03:00 → previous day');
assert(opDayFromIsoTimeService('2026-08-24', '07:00', 5) === '2026-08-24', 'ISO 07:00 → current day');
assert(opDayFromIsoTimeService('2026-08-24', '00:00', 5) === '2026-08-23', 'ISO 00:00 → previous day');

// ─── Month boundary ─────────────────────────────────────────
console.log('Month boundary:');

// Aug 1 at 3 AM → July 31
const aug1_3am = new Date(2026, 7, 1, 3, 0, 0);
assert(opDayService(aug1_3am, 5) === '2026-07-31', 'Aug 1 3AM → July 31');

// Jan 1 at 2 AM → Dec 31 of previous year
const jan1_2am = new Date(2026, 0, 1, 2, 0, 0);
assert(opDayService(jan1_2am, 5) === '2025-12-31', 'Jan 1 2AM → Dec 31 prev year');

// ─── Summary ─────────────────────────────────────────────────
console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failed > 0) process.exit(1);
