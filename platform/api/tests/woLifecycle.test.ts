/**
 * Unit tests for woLifecycle.ts
 *
 * Verifies exact behavioral parity with Code.js:
 *   Status state machine, date derivation, issue appending,
 *   operational day correction, waterblast gate, photos flag.
 *
 * Usage: npx tsx tests/woLifecycle.test.ts
 */
import {
  advanceStatus, deriveWoDates, appendIssues,
  correctOperationalDay, checkWaterblastGate,
  STATUS_SORT_PRIORITY,
} from '../src/services/woLifecycle.js';

let passed = 0;
let failed = 0;
const failures: string[] = [];

function test(label: string, fn: () => void) {
  try {
    fn();
    passed++;
  } catch (err: any) {
    failed++;
    failures.push(`${label}: ${err.message}`);
    console.error(`  ✗ ${label}: ${err.message}`);
  }
}

function assert(condition: boolean, msg: string) {
  if (!condition) throw new Error(msg);
}

function assertEqual(actual: unknown, expected: unknown, label: string) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) throw new Error(`${label}: expected ${e}, got ${a}`);
}

// ═══════════════════════════════════════════════════════════════
// advanceStatus
// ═══════════════════════════════════════════════════════════════
console.log('\nadvanceStatus:');

test('received + complete → completed', () => {
  assertEqual(
    advanceStatus('received', { isComplete: true, hasDispatchDate: false, hasWorkStartDate: false }),
    'completed', 'status');
});

test('dispatched + complete → completed', () => {
  assertEqual(
    advanceStatus('dispatched', { isComplete: true, hasDispatchDate: true, hasWorkStartDate: false }),
    'completed', 'status');
});

test('in_progress + complete → completed', () => {
  assertEqual(
    advanceStatus('in_progress', { isComplete: true, hasDispatchDate: true, hasWorkStartDate: true }),
    'completed', 'status');
});

test('received + workStart → in_progress', () => {
  assertEqual(
    advanceStatus('received', { isComplete: false, hasDispatchDate: true, hasWorkStartDate: true }),
    'in_progress', 'status');
});

test('dispatched + workStart → in_progress', () => {
  assertEqual(
    advanceStatus('dispatched', { isComplete: false, hasDispatchDate: true, hasWorkStartDate: true }),
    'in_progress', 'status');
});

test('received + dispatch only → dispatched', () => {
  assertEqual(
    advanceStatus('received', { isComplete: false, hasDispatchDate: true, hasWorkStartDate: false }),
    'dispatched', 'status');
});

test('received + no dates → stays received', () => {
  assertEqual(
    advanceStatus('received', { isComplete: false, hasDispatchDate: false, hasWorkStartDate: false }),
    'received', 'status');
});

test('in_progress + no complete → stays in_progress', () => {
  assertEqual(
    advanceStatus('in_progress', { isComplete: false, hasDispatchDate: true, hasWorkStartDate: true }),
    'in_progress', 'status');
});

test('completed + no complete → stays completed (never backward)', () => {
  assertEqual(
    advanceStatus('completed', { isComplete: false, hasDispatchDate: true, hasWorkStartDate: true }),
    'completed', 'status');
});

test('returned + no action → stays returned', () => {
  assertEqual(
    advanceStatus('returned', { isComplete: false, hasDispatchDate: false, hasWorkStartDate: false }),
    'returned', 'status');
});

// Handle Title Case input (old app used "In Progress" not "in_progress")
test('handles "In Progress" (Title Case) → in_progress on complete', () => {
  assertEqual(
    advanceStatus('In Progress', { isComplete: true, hasDispatchDate: true, hasWorkStartDate: true }),
    'completed', 'status');
});

test('handles "Received" (Title Case) → dispatched', () => {
  assertEqual(
    advanceStatus('Received', { isComplete: false, hasDispatchDate: true, hasWorkStartDate: false }),
    'dispatched', 'status');
});

// ═══════════════════════════════════════════════════════════════
// deriveWoDates
// ═══════════════════════════════════════════════════════════════
console.log('\nderiveWoDates:');

test('blank dates → both filled from workDate', () => {
  const r = deriveWoDates({}, '2026-08-15', false);
  assertEqual(r.dispatchDate, '2026-08-15', 'dispatch');
  assertEqual(r.workStartDate, '2026-08-15', 'workStart');
  assertEqual(r.workEndDate, null, 'workEnd');
});

test('existing dates preserved', () => {
  const r = deriveWoDates(
    { dispatchDate: '2026-08-10', workStartDate: '2026-08-11', workEndDate: null },
    '2026-08-15',
    false,
  );
  assertEqual(r.dispatchDate, '2026-08-10', 'dispatch');
  assertEqual(r.workStartDate, '2026-08-11', 'workStart');
  assertEqual(r.workEndDate, null, 'workEnd');
});

test('complete sets workEndDate', () => {
  const r = deriveWoDates(
    { dispatchDate: '2026-08-10', workStartDate: '2026-08-11' },
    '2026-08-15',
    true,
  );
  assertEqual(r.workEndDate, '2026-08-15', 'workEnd');
});

test('complete with existing workEndDate overwrites', () => {
  const r = deriveWoDates(
    { dispatchDate: '2026-08-10', workStartDate: '2026-08-11', workEndDate: '2026-08-12' },
    '2026-08-15',
    true,
  );
  assertEqual(r.workEndDate, '2026-08-15', 'workEnd');
});

test('non-complete preserves existing workEndDate', () => {
  const r = deriveWoDates(
    { workEndDate: '2026-08-12' },
    '2026-08-15',
    false,
  );
  assertEqual(r.workEndDate, '2026-08-12', 'workEnd');
});

// ═══════════════════════════════════════════════════════════════
// appendIssues
// ═══════════════════════════════════════════════════════════════
console.log('\nappendIssues:');

test('empty existing + new → just the new issue', () => {
  assertEqual(
    appendIssues('', 'Cracking at curb', '2026-08-15'),
    '2026-08-15: Cracking at curb',
    'result',
  );
});

test('existing + new → appended with newline', () => {
  assertEqual(
    appendIssues('2026-08-10: Pothole', 'Rain delay', '2026-08-15'),
    '2026-08-10: Pothole\n2026-08-15: Rain delay',
    'result',
  );
});

test('existing + empty new → preserved unchanged', () => {
  assertEqual(
    appendIssues('2026-08-10: Pothole', '', '2026-08-15'),
    '2026-08-10: Pothole',
    'result',
  );
});

test('existing + null new → preserved unchanged', () => {
  assertEqual(
    appendIssues('2026-08-10: Pothole', null, '2026-08-15'),
    '2026-08-10: Pothole',
    'result',
  );
});

test('null existing + new → just the new issue', () => {
  assertEqual(
    appendIssues(null, 'First issue', '2026-08-15'),
    '2026-08-15: First issue',
    'result',
  );
});

test('both empty → empty string', () => {
  assertEqual(appendIssues('', '', '2026-08-15'), '', 'result');
});

test('multiple appends accumulate correctly', () => {
  let issues = appendIssues('', 'Day 1 issue', '2026-08-10');
  issues = appendIssues(issues, 'Day 2 issue', '2026-08-11');
  issues = appendIssues(issues, '', '2026-08-12'); // no issue, no change
  issues = appendIssues(issues, 'Day 4 issue', '2026-08-13');
  assertEqual(
    issues,
    '2026-08-10: Day 1 issue\n2026-08-11: Day 2 issue\n2026-08-13: Day 4 issue',
    'result',
  );
});

test('trims whitespace from issues', () => {
  assertEqual(
    appendIssues('  old  ', '  new  ', '2026-08-15'),
    'old\n2026-08-15: new',
    'result',
  );
});

// ═══════════════════════════════════════════════════════════════
// checkWaterblastGate
// ═══════════════════════════════════════════════════════════════
console.log('\ncheckWaterblastGate:');

test('MMA + not confirmed → error', () => {
  const err = checkWaterblastGate('Yes - MMA', 'No');
  assert(err !== null, 'should error');
  assert(err!.includes('Waterblasting not confirmed'), 'message');
});

test('MMA + confirmed → null', () => {
  assertEqual(checkWaterblastGate('Yes - MMA', 'Yes'), null, 'err');
});

test('Thermo → null (no gate)', () => {
  assertEqual(checkWaterblastGate('No - Thermo', 'N/A'), null, 'err');
});

test('blank → null (no gate)', () => {
  assertEqual(checkWaterblastGate('', ''), null, 'err');
});

test('null values → null (no gate)', () => {
  assertEqual(checkWaterblastGate(null, null), null, 'err');
});

test('MMA + empty confirmed → error', () => {
  const err = checkWaterblastGate('Yes - MMA', '');
  assert(err !== null, 'should error');
});

test('MMA + N/A confirmed → error', () => {
  const err = checkWaterblastGate('Yes - MMA', 'N/A');
  assert(err !== null, 'should error');
});

// ═══════════════════════════════════════════════════════════════
// STATUS_SORT_PRIORITY
// ═══════════════════════════════════════════════════════════════
console.log('\nSTATUS_SORT_PRIORITY:');

test('in_progress sorts first', () => {
  const statuses: string[] = ['completed', 'received', 'in_progress', 'dispatched', 'returned'];
  const sorted = statuses.sort((a, b) =>
    (STATUS_SORT_PRIORITY[a] ?? 99) - (STATUS_SORT_PRIORITY[b] ?? 99)
  );
  assertEqual(sorted[0], 'in_progress', 'first');
  assertEqual(sorted[1], 'dispatched', 'second');
  assertEqual(sorted[2], 'received', 'third');
  assertEqual(sorted[3], 'completed', 'fourth');
  assertEqual(sorted[4], 'returned', 'fifth');
});

// ═══════════════════════════════════════════════════════════════
// correctOperationalDay
// ═══════════════════════════════════════════════════════════════
console.log('\ncorrectOperationalDay:');

test('non-today date is never corrected', () => {
  // Even if op-day differs, a manually-picked date stays
  assertEqual(
    correctOperationalDay('2026-01-01', 'America/New_York', 5),
    '2026-01-01',
    'date',
  );
});

// We can't easily test the "3 AM correction" without mocking Date,
// but we can verify the function doesn't crash and returns a valid date.
test('today date returns a valid YYYY-MM-DD', () => {
  const now = new Date();
  const today = now.toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
  const result = correctOperationalDay(today, 'America/New_York', 5);
  assert(/^\d{4}-\d{2}-\d{2}$/.test(result), `should be YYYY-MM-DD, got ${result}`);
});

// ═══════════════════════════════════════════════════════════════
// Summary
// ═══════════════════════════════════════════════════════════════

console.log(`\n=== ${passed} passed, ${failed} failed ===`);
if (failures.length > 0) {
  console.log('\nFAILURES:');
  for (const f of failures) console.log(`  ✗ ${f}`);
}
process.exit(failed > 0 ? 1 : 0);
