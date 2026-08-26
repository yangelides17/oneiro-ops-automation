/**
 * Tests for document lifecycle service.
 * Validates doc key generation, parsing, and payroll week helpers.
 */
import {
  buildDocKey, buildPlDocKey, buildMonthEndDocKey, parseDocKey, parsePlDocKey, parseMonthEndDocKey,
  isValidTransition, payrollWeekStart, payrollWeekEnd, payrollMonthIso,
} from '../src/services/docLifecycle.js';

let passed = 0;
let failed = 0;

function assert(condition: boolean, label: string) {
  if (condition) { passed++; } else { failed++; console.error(`  FAIL: ${label}`); }
}

// ─── buildPlDocKey (Production Log — keyed by contractor, NOT contract+region) ──
console.log('buildPlDocKey:');

assert(
  buildPlDocKey('2026-08-24', 'Metro Express') === 'PL_2026-08-24_Metro_Express',
  'PL key: contractor slug with space→underscore',
);
assert(
  buildPlDocKey('2026-08-24', 'Metro Express', 'Bob Smith') === 'PL_2026-08-24_Metro_Express_chief-BobSmith',
  'PL key with crew chief',
);
assert(
  buildPlDocKey('2026-08-24', 'Metro Express', '') === 'PL_2026-08-24_Metro_Express',
  'PL key: blank chief → no suffix',
);
assert(
  buildPlDocKey('2026-08-24', 'Metro Express', null) === 'PL_2026-08-24_Metro_Express',
  'PL key: null chief → no suffix',
);

// ─── buildDocKey (SI/CP — keyed by contract+region) ──────────
console.log('buildDocKey:');

assert(
  buildDocKey('signin', '2026-08-24', '84125MBTP701', 'BK', 'John Smith') === 'SI_2026-08-24_84125MBTP701_BK_chief-JohnSmith',
  'SI key with crew chief',
);
assert(
  buildDocKey('certified_payroll', '2026-08-18', '84125MBTP701', 'BK') === 'CP_2026-08-18_84125MBTP701_BK',
  'CP key',
);
// /EXT suffix stripped
assert(
  buildDocKey('signin', '2026-08-24', '84125MBTP701/EXT', 'BK') === 'SI_2026-08-24_84125MBTP701_BK',
  '/EXT stripped from contract',
);
// Blank crew chief → no suffix
assert(
  buildDocKey('signin', '2026-08-24', '84125MBTP701', 'BK', '') === 'SI_2026-08-24_84125MBTP701_BK',
  'Blank crew chief → no suffix',
);
assert(
  buildDocKey('signin', '2026-08-24', '84125MBTP701', 'BK', null) === 'SI_2026-08-24_84125MBTP701_BK',
  'Null crew chief → no suffix',
);
// PL via buildDocKey → empty (PL uses buildPlDocKey instead)
assert(
  buildDocKey('production_log', '2026-08-24', '84125', 'BK') === '',
  'production_log via buildDocKey → empty (use buildPlDocKey)',
);
// Unknown doc type → empty
assert(
  buildDocKey('unknown_type', '2026-08-24', '84125', 'BK') === '',
  'Unknown doc type → empty string',
);

// ─── buildMonthEndDocKey ─────────────────────────────────────
console.log('buildMonthEndDocKey:');

assert(
  buildMonthEndDocKey('EU', '2026-07', '84125MBTP701', 'BK', 'Metro Express') === 'EU_2026-07_84125MBTP701_BK_MetroExpress',
  'EU month-end key',
);
assert(
  buildMonthEndDocKey('CERT', '2026-07', '84125MBTP701', 'BK', 'Metro Express') === 'CERT_2026-07_84125MBTP701_BK_MetroExpress',
  'CERT month-end key',
);

// ─── parsePlDocKey ───────────────────────────────────────────
console.log('parsePlDocKey:');

let plParsed = parsePlDocKey('PL_2026-08-24_Metro_Express');
assert(plParsed !== null, 'PL key parses');
assert(plParsed!.anchorDate === '2026-08-24', 'PL anchor date');
assert(plParsed!.contractorSlug === 'Metro_Express', 'PL contractor slug');
assert(plParsed!.crewChief === undefined, 'PL no crew chief');

plParsed = parsePlDocKey('PL_2026-08-24_Metro_Express_chief-BobSmith');
assert(plParsed !== null, 'PL key with chief parses');
assert(plParsed!.crewChief === 'BobSmith', 'PL crew chief slug');

assert(parsePlDocKey('SI_2026-08-24_84125_BK') === null, 'SI key → null for PL parser');

// ─── parseDocKey (SI/CP) ─────────────────────────────────────
console.log('parseDocKey:');

let parsed = parseDocKey('SI_2026-08-24_84125MBTP701_BK');
assert(parsed !== null, 'SI key parses');
assert(parsed!.prefix === 'SI', 'SI prefix');
assert(parsed!.anchorDate === '2026-08-24', 'anchor date');
assert(parsed!.contractNum === '84125MBTP701', 'contract num');
assert(parsed!.regionCode === 'BK', 'region code');
assert(parsed!.crewChief === undefined, 'no crew chief');

parsed = parseDocKey('SI_2026-08-24_84125MBTP701_BK_chief-JohnSmith');
assert(parsed !== null, 'SI key with chief parses');
assert(parsed!.prefix === 'SI', 'SI prefix');
assert(parsed!.crewChief === 'JohnSmith', 'crew chief slug');

parsed = parseDocKey('CP_2026-08-18_84125MBTP701_BK');
assert(parsed !== null, 'CP key parses');
assert(parsed!.prefix === 'CP', 'CP prefix');

assert(parseDocKey('PL_2026-08-24_Metro_Express') === null, 'PL key → null for SI/CP parser');
assert(parseDocKey('garbage') === null, 'garbage → null');
assert(parseDocKey('') === null, 'empty → null');

// ─── parseMonthEndDocKey ─────────────────────────────────────
console.log('parseMonthEndDocKey:');

let meParsed = parseMonthEndDocKey('EU_2026-07_84125MBTP701_BK_MetroExpress');
assert(meParsed !== null, 'EU key parses');
assert(meParsed!.key === 'EU', 'EU key');
assert(meParsed!.monthIso === '2026-07', 'month');
assert(meParsed!.contractorSlug === 'MetroExpress', 'contractor slug');

assert(parseMonthEndDocKey('garbage') === null, 'garbage → null');

// ─── isValidTransition ───────────────────────────────────────
console.log('isValidTransition:');

assert(isValidTransition('pending', 'needs_review') === true, 'pending → needs_review');
assert(isValidTransition('needs_review', 'approved') === true, 'needs_review → approved');
assert(isValidTransition('approved', 'archived') === true, 'approved → archived');
assert(isValidTransition('needs_review', 'pending') === true, 'needs_review → pending (re-gen)');
assert(isValidTransition('pending', 'approved') === false, 'pending → approved (invalid)');
assert(isValidTransition('archived', 'pending') === false, 'archived → pending (invalid)');
assert(isValidTransition('approved', 'pending') === false, 'approved → pending (invalid)');

// ─── payrollWeekStart ────────────────────────────────────────
console.log('payrollWeekStart:');

// 2026-08-24 is a Monday → week starts Sunday 2026-08-23
assert(payrollWeekStart('2026-08-24') === '2026-08-23', 'Monday → previous Sunday');
// 2026-08-23 is a Sunday → already the start
assert(payrollWeekStart('2026-08-23') === '2026-08-23', 'Sunday → same day');
// 2026-08-29 is a Saturday → week started 2026-08-23
assert(payrollWeekStart('2026-08-29') === '2026-08-23', 'Saturday → previous Sunday');
// 2026-08-22 is a Saturday → week started 2026-08-16
assert(payrollWeekStart('2026-08-22') === '2026-08-16', 'Saturday → week start');

// ─── payrollWeekEnd ──────────────────────────────────────────
console.log('payrollWeekEnd:');

assert(payrollWeekEnd('2026-08-23') === '2026-08-29', 'Sunday → Saturday');
assert(payrollWeekEnd('2026-08-16') === '2026-08-22', 'Sun Aug 16 → Sat Aug 22');

// ─── payrollMonthIso ─────────────────────────────────────────
console.log('payrollMonthIso:');

// A week ending in August → month is August
assert(payrollMonthIso('2026-08-24') === '2026-08', 'Mid-Aug → 2026-08');
// A date near month boundary: Aug 31 is Monday, week ends Sep 5
assert(payrollMonthIso('2026-08-31') === '2026-09', 'Aug 31 (Mon) → week ends Sep 5 → 2026-09');

// ─── Summary ─────────────────────────────────────────────────
console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failed > 0) process.exit(1);
