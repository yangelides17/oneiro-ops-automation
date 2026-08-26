/**
 * Unit tests for markingItemLogic.ts
 *
 * Verifies exact behavioral parity with Code.js:
 *   CATEGORY_UNITS_, unitForCategory_, expandDirLetters_,
 *   handleCreateMarkingItem_ (validation), handleUpdateMarkingItem_ (rules),
 *   finalizeMarkingStatus_, computeMarkingRollups_.
 *
 * Usage: npx tsx tests/markingItemLogic.test.ts
 */
import {
  deriveUnit, isUnitLocked, validateGridCategory,
  expandDirections, enforceCreateDefaults, applyUpdateRules,
  computeMarkingRollups,
  CATEGORY_UNITS,
} from '../src/services/markingItemLogic.js';

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
// deriveUnit
// ═══════════════════════════════════════════════════════════════
console.log('\nderiveUnit:');

test('HVX Crosswalk → LF', () => assertEqual(deriveUnit('HVX Crosswalk'), 'LF', 'unit'));
test('Stop Msg → EA', () => assertEqual(deriveUnit('Stop Msg'), 'EA', 'unit'));
test('Stop Line → LF', () => assertEqual(deriveUnit('Stop Line'), 'LF', 'unit'));
test('Double Yellow Line → LF', () => assertEqual(deriveUnit('Double Yellow Line'), 'LF', 'unit'));
test('Lane Lines → EA', () => assertEqual(deriveUnit('Lane Lines'), 'EA', 'unit'));
test('Bike Lane → SF', () => assertEqual(deriveUnit('Bike Lane'), 'SF', 'unit'));
test('Bus Lane → SF', () => assertEqual(deriveUnit('Bus Lane'), 'SF', 'unit'));
test('Pedestrian Space → SF', () => assertEqual(deriveUnit('Pedestrian Space'), 'SF', 'unit'));
test('Others → "" (variable)', () => assertEqual(deriveUnit('Others'), '', 'unit'));
test('Empty string → ""', () => assertEqual(deriveUnit(''), '', 'unit'));
test('Unknown category → ""', () => assertEqual(deriveUnit('Nonexistent'), '', 'unit'));
test('Whitespace trimmed', () => assertEqual(deriveUnit('  HVX Crosswalk  '), 'LF', 'unit'));
test('Bike Lane Arrow → EA', () => assertEqual(deriveUnit('Bike Lane Arrow'), 'EA', 'unit'));
test('Old Bike Symbol (w/ rider) → EA', () => assertEqual(deriveUnit('Old Bike Symbol (w/ rider)'), 'EA', 'unit'));
test('Bike Lane Green Bar → SF', () => assertEqual(deriveUnit('Bike Lane Green Bar'), 'SF', 'unit'));
test('Bike Lane Symbol (legacy) → EA', () => assertEqual(deriveUnit('Bike Lane Symbol'), 'EA', 'unit'));

test('isUnitLocked true for locked', () => assert(isUnitLocked('HVX Crosswalk'), 'should be locked'));
test('isUnitLocked false for Others', () => assert(!isUnitLocked('Others'), 'should not be locked'));

// Verify map has same number of entries as Code.js (35 categories)
test('CATEGORY_UNITS has 40 entries (matches Code.js + frontend)', () =>
  assertEqual(Object.keys(CATEGORY_UNITS).length, 40, 'count'));

// ═══════════════════════════════════════════════════════════════
// validateGridCategory
// ═══════════════════════════════════════════════════════════════
console.log('\nvalidateGridCategory:');

test('HVX Crosswalk requires intersection', () => {
  const err = validateGridCategory('HVX Crosswalk', '', 'N');
  assert(err !== null, 'should error');
  assert(err!.includes('Intersection'), 'should mention intersection');
});

test('HVX Crosswalk requires direction', () => {
  const err = validateGridCategory('HVX Crosswalk', '5th Ave', '');
  assert(err !== null, 'should error');
  assert(err!.includes('Direction'), 'should mention direction');
});

test('HVX Crosswalk valid when both provided', () => {
  assertEqual(validateGridCategory('HVX Crosswalk', '5th Ave', 'N'), null, 'err');
});

test('Stop Msg requires intersection', () => {
  assert(validateGridCategory('Stop Msg', '', 'E') !== null, 'should error');
});

test('Stop Line requires intersection', () => {
  assert(validateGridCategory('Stop Line', '', 'W') !== null, 'should error');
});

test('Non-grid category passes without intersection', () => {
  assertEqual(validateGridCategory('Double Yellow Line', '', ''), null, 'err');
});

test('Others passes without intersection', () => {
  assertEqual(validateGridCategory('Others', '', ''), null, 'err');
});

// ═══════════════════════════════════════════════════════════════
// expandDirections
// ═══════════════════════════════════════════════════════════════
console.log('\nexpandDirections:');

test('"North" → ["N"]', () => assertEqual(expandDirections('North'), ['N'], 'dirs'));
test('"East" → ["E"]', () => assertEqual(expandDirections('East'), ['E'], 'dirs'));
test('"South" → ["S"]', () => assertEqual(expandDirections('South'), ['S'], 'dirs'));
test('"West" → ["W"]', () => assertEqual(expandDirections('West'), ['W'], 'dirs'));
test('"EW" → ["E","W"]', () => assertEqual(expandDirections('EW'), ['E', 'W'], 'dirs'));
test('"NS" → ["N","S"]', () => assertEqual(expandDirections('NS'), ['N', 'S'], 'dirs'));
test('"NSEW" → ["N","S","E","W"]', () => assertEqual(expandDirections('NSEW'), ['N', 'S', 'E', 'W'], 'dirs'));
test('"" → []', () => assertEqual(expandDirections(''), [], 'dirs'));
test('null → []', () => assertEqual(expandDirections(null as any), [], 'dirs'));
test('"N" → ["N"]', () => assertEqual(expandDirections('N'), ['N'], 'dirs'));
test('lowercase "north" → ["N"]', () => assertEqual(expandDirections('north'), ['N'], 'dirs'));
test('lowercase "ew" → ["E","W"]', () => assertEqual(expandDirections('ew'), ['E', 'W'], 'dirs'));
test('dedupes "NN" → ["N"]', () => assertEqual(expandDirections('NN'), ['N'], 'dirs'));
test('mixed "NEN" → ["N","E"]', () => assertEqual(expandDirections('NEN'), ['N', 'E'], 'dirs'));
test('non-NSEW chars ignored "NAB" → ["N"]', () => assertEqual(expandDirections('NAB'), ['N'], 'dirs'));

// ═══════════════════════════════════════════════════════════════
// enforceCreateDefaults
// ═══════════════════════════════════════════════════════════════
console.log('\nenforceCreateDefaults:');

test('derives unit from category', () => {
  const result = enforceCreateDefaults({ category: 'HVX Crosswalk' });
  assertEqual(result.unit, 'LF', 'unit');
});

test('forces status to pending', () => {
  const result = enforceCreateDefaults({ category: 'HVX Crosswalk', status: 'completed' });
  assertEqual(result.status, 'pending', 'status');
});

test('defaults addedBy to manual', () => {
  const result = enforceCreateDefaults({ category: 'Others' });
  assertEqual(result.addedBy, 'manual', 'addedBy');
});

test('preserves addedBy=scanner', () => {
  const result = enforceCreateDefaults({ category: 'Others', addedBy: 'scanner' });
  assertEqual(result.addedBy, 'scanner', 'addedBy');
});

test('defaults woSection to manual', () => {
  const result = enforceCreateDefaults({ category: 'Others' });
  assertEqual(result.woSection, 'manual', 'woSection');
});

test('preserves woSection=intersection_grid', () => {
  const result = enforceCreateDefaults({ category: 'HVX Crosswalk', woSection: 'intersection_grid' });
  assertEqual(result.woSection, 'intersection_grid', 'woSection');
});

test('variable category uses client unit', () => {
  const result = enforceCreateDefaults({ category: 'Others', unit: 'SF' });
  assertEqual(result.unit, 'SF', 'unit');
});

test('variable category defaults to EA', () => {
  const result = enforceCreateDefaults({ category: 'Others' });
  assertEqual(result.unit, 'EA', 'unit');
});

test('locked category overrides client unit', () => {
  const result = enforceCreateDefaults({ category: 'HVX Crosswalk', unit: 'EA' });
  assertEqual(result.unit, 'LF', 'unit');
});

// ═══════════════════════════════════════════════════════════════
// applyUpdateRules
// ═══════════════════════════════════════════════════════════════
console.log('\napplyUpdateRules:');

test('category change → auto-derive unit', () => {
  const result = applyUpdateRules({ category: 'Stop Msg' });
  assertEqual(result.unit, 'EA', 'unit');
});

test('category change to locked overrides client unit', () => {
  const result = applyUpdateRules({ category: 'HVX Crosswalk', unit: 'EA' });
  assertEqual(result.unit, 'LF', 'unit');
});

test('category change to variable preserves client unit', () => {
  const result = applyUpdateRules({ category: 'Others', unit: 'SF' });
  assertEqual(result.unit, 'SF', 'unit');
});

test('qty=0 → reverts to pending + clears dateCompleted', () => {
  const result = applyUpdateRules({ quantity: 0 });
  assertEqual(result.status, 'pending', 'status');
  assertEqual(result.dateCompleted, null, 'dateCompleted');
});

test('qty=null → reverts to pending + clears dateCompleted', () => {
  const result = applyUpdateRules({ quantity: null });
  assertEqual(result.status, 'pending', 'status');
  assertEqual(result.dateCompleted, null, 'dateCompleted');
});

test('qty="" (empty string) → reverts to pending', () => {
  const result = applyUpdateRules({ quantity: '' });
  assertEqual(result.status, 'pending', 'status');
});

test('positive qty does NOT change status', () => {
  const result = applyUpdateRules({ quantity: 150 });
  assert(result.status === undefined, 'status should not be set');
  assert(result.dateCompleted === undefined, 'dateCompleted should not be set');
});

test('non-quantity field does not affect status', () => {
  const result = applyUpdateRules({ description: 'test', notes: 'updated' });
  assert(result.status === undefined, 'status should not be set');
});

test('unit on existing locked category gets overridden', () => {
  const result = applyUpdateRules({ unit: 'EA' }, 'HVX Crosswalk');
  assertEqual(result.unit, 'LF', 'unit');
});

test('unit on existing variable category passes through', () => {
  const result = applyUpdateRules({ unit: 'SF' }, 'Others');
  assertEqual(result.unit, 'SF', 'unit');
});

// ═══════════════════════════════════════════════════════════════
// computeMarkingRollups
// ═══════════════════════════════════════════════════════════════
console.log('\ncomputeMarkingRollups:');

test('empty items → blank rollups', () => {
  const r = computeMarkingRollups([]);
  assertEqual(r.markingTypes, '', 'markingTypes');
  assertEqual(r.quantityCompleted, null, 'quantityCompleted');
  assertEqual(r.paintMaterial, '', 'paintMaterial');
});

test('thermo WO: sums LF only, types=N/A', () => {
  const items = [
    { workType: 'Thermo', category: 'HVX Crosswalk', quantity: '120', unit: 'LF', status: 'completed' },
    { workType: 'Thermo', category: 'Stop Msg', quantity: '2', unit: 'EA', status: 'completed' },
    { workType: 'Thermo', category: 'Double Yellow Line', quantity: '300', unit: 'LF', status: 'completed' },
  ];
  const r = computeMarkingRollups(items, 'Thermo');
  assertEqual(r.markingTypes, 'N/A', 'markingTypes');
  assertEqual(r.quantityCompleted, 420, 'quantityCompleted'); // 120 + 300 LF
  assertEqual(r.paintMaterial, 'N/A', 'paintMaterial');
});

test('thermo WO: EA items excluded from sum', () => {
  const items = [
    { workType: 'Thermo', category: 'Stop Msg', quantity: '5', unit: 'EA', status: 'completed' },
  ];
  const r = computeMarkingRollups(items, 'Thermo');
  assertEqual(r.quantityCompleted, null, 'quantityCompleted');
});

test('MMA WO: sums SF only, collects ALL distinct completed categories', () => {
  const items = [
    { workType: 'MMA', category: 'Bike Lane', quantity: '500', unit: 'SF', colorMaterial: 'Green', status: 'completed' },
    { workType: 'MMA', category: 'Bus Lane', quantity: '300', unit: 'SF', colorMaterial: 'Red', status: 'completed' },
    { workType: 'MMA', category: 'Bike Lane Arrow', quantity: '3', unit: 'EA', status: 'completed' },
  ];
  const r = computeMarkingRollups(items, 'MMA');
  // All completed categories listed (EA items included in types, just not in qty sum)
  assertEqual(r.markingTypes, 'Bike Lane, Bike Lane Arrow, Bus Lane', 'markingTypes');
  assertEqual(r.quantityCompleted, 800, 'quantityCompleted'); // 500 + 300 SF only
  assertEqual(r.paintMaterial, 'Green, Red', 'paintMaterial');
});

test('MMA: only completed items contribute to types/materials', () => {
  const items = [
    { workType: 'MMA', category: 'Bike Lane', quantity: '500', unit: 'SF', colorMaterial: 'Green', status: 'completed' },
    { workType: 'MMA', category: 'Bus Lane', quantity: '300', unit: 'SF', colorMaterial: 'Red', status: 'pending' },
  ];
  const r = computeMarkingRollups(items, 'MMA');
  assertEqual(r.markingTypes, 'Bike Lane', 'markingTypes');
  assertEqual(r.paintMaterial, 'Green', 'paintMaterial');
  assertEqual(r.quantityCompleted, 800, 'quantityCompleted'); // both counted in qty sum
});

test('MMA: N/A material excluded', () => {
  const items = [
    { workType: 'MMA', category: 'Bike Lane', quantity: '100', unit: 'SF', colorMaterial: 'N/A', status: 'completed' },
  ];
  const r = computeMarkingRollups(items, 'MMA');
  assertEqual(r.paintMaterial, '', 'paintMaterial');
});

test('mixed: workType from items takes precedence', () => {
  const items = [
    { workType: 'Thermo', category: 'Double Yellow Line', quantity: '100', unit: 'LF', status: 'completed' },
  ];
  // woWorkType is MMA but item says Thermo → thermo path
  const r = computeMarkingRollups(items, 'MMA');
  assertEqual(r.markingTypes, 'N/A', 'markingTypes');
  assertEqual(r.quantityCompleted, 100, 'quantityCompleted');
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
