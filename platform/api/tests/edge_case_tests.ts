/**
 * Edge Case Tests
 *
 * Tests boundary conditions: empty lists, zero quantities, duplicate WO numbers,
 * very long strings, special characters, concurrent-like operations.
 *
 * Usage: npx tsx tests/edge_case_tests.ts
 */
import 'dotenv/config';

const BASE = 'http://localhost:3001';
let COOKIE = '';
let contractorId = '';

let passed = 0;
let failed = 0;
const failures: string[] = [];

function log(msg: string) { console.log(msg); }

async function test(label: string, fn: () => Promise<void>) {
  try {
    await fn();
    passed++;
    console.log(`  ✓ ${label}`);
  } catch (err: any) {
    failed++;
    const msg = `${label}: ${err.message}`;
    failures.push(msg);
    console.error(`  ✗ ${msg}`);
  }
}

function assert(condition: boolean, msg: string) {
  if (!condition) throw new Error(msg);
}

async function api(method: string, path: string, body?: any): Promise<any> {
  const opts: RequestInit = {
    method,
    headers: {
      'Content-Type': 'application/json',
      'Cookie': COOKIE,
    },
  };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(`${BASE}${path}`, opts);
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = { _raw: text }; }
  return { status: res.status, data, ok: res.ok };
}

async function run() {
  log('\n═══ EDGE CASE TESTS ═══\n');

  // Setup
  log('── SETUP ──');
  await test('Login to sandbox account', async () => {
    const res = await fetch(`${BASE}/api/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: 'sandbox@oneiro-test.com', password: 'sandbox123' }),
    });
    const setCookie = res.headers.get('set-cookie') || '';
    COOKIE = setCookie.split(';')[0];
    assert(res.ok, 'Login failed');
  });

  await test('Get/create contractor for tests', async () => {
    const { data } = await api('GET', '/api/settings/contractors');
    if (data?.length > 0) {
      contractorId = data[0].id;
    } else {
      const { data: c } = await api('POST', '/api/settings/contractors', { name: 'Edge Case Co' });
      contractorId = c.id;
    }
    assert(!!contractorId, 'Should have contractor');
  });

  // ═══════════════════════════════════════════════════════════════
  // EMPTY LISTS
  // ═══════════════════════════════════════════════════════════════
  log('\n── EMPTY LISTS ──');

  await test('WO with no markings returns empty array', async () => {
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-EMPTY-001', contractorId,
    });
    const { data, ok } = await api('GET', `/api/wos/${wo.id}/markings`);
    assert(ok, 'Should succeed');
    assert(data.items?.length === 0, `Should have 0 items, got ${data.items?.length}`);
    await api('DELETE', `/api/wos/${wo.id}`);
  });

  await test('Revenue with no matching data returns zeros', async () => {
    const { data, ok } = await api('GET', '/api/revenue?start=1900-01-01&end=1900-01-31');
    assert(ok, 'Should succeed');
    assert(data.totals?.revenue === 0 || data.totals?.items === 0, 'Should return zero for non-existent date range');
  });

  await test('Production with no matching data returns empty', async () => {
    const { data, ok } = await api('GET', '/api/production?start=1900-01-01&end=1900-01-31');
    assert(ok, 'Should succeed');
    assert(data.totals?.SF === 0 || data.totals?.items === 0, 'Should return zero for non-existent range');
  });

  // ═══════════════════════════════════════════════════════════════
  // DUPLICATE WO NUMBERS
  // ═══════════════════════════════════════════════════════════════
  log('\n── DUPLICATE WO NUMBERS ──');

  let dupeWo1Id = '';
  let dupeWo2Id = '';

  await test('Create WO with unique number', async () => {
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-DUPE-001', contractorId,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    dupeWo1Id = data.id;
  });

  await test('Create second WO with same number fails (409 or 500)', async () => {
    const { status } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-DUPE-001', contractorId,
    });
    // Should fail with constraint violation (409) or validation error
    assert(status === 409 || status === 400 || status === 500, `Expected conflict, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // ZERO / EXTREME QUANTITIES
  // ═══════════════════════════════════════════════════════════════
  log('\n── ZERO / EXTREME QUANTITIES ──');

  let zeroMarkingId = '';
  await test('Create marking with zero quantity', async () => {
    const { data, ok } = await api('POST', `/api/wos/${dupeWo1Id}/markings`, {
      category: 'Double Yellow Line', quantity: 0, unit: 'LF',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.quantity === '0.00', `Should store 0, got ${data.quantity}`);
    zeroMarkingId = data.id;
  });

  await test('Create marking with very large quantity', async () => {
    const { data, ok } = await api('POST', `/api/wos/${dupeWo1Id}/markings`, {
      category: 'Single White Line', quantity: 999999.99, unit: 'LF',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.quantity === '999999.99', `Should store large number, got ${data.quantity}`);
    // Cleanup
    await api('DELETE', '/api/markings', { ids: [data.id] });
  });

  await test('Create marking with negative quantity (should accept or reject gracefully)', async () => {
    const { status, data } = await api('POST', `/api/wos/${dupeWo1Id}/markings`, {
      category: 'Lane Arrow', quantity: -5, unit: 'EA',
    });
    // Should either reject (400) or store — not crash
    assert(status !== 500, `Server error with negative quantity (${status})`);
    if (status === 201 && data.id) {
      await api('DELETE', '/api/markings', { ids: [data.id] });
    }
  });

  // ═══════════════════════════════════════════════════════════════
  // SPECIAL CHARACTERS
  // ═══════════════════════════════════════════════════════════════
  log('\n── SPECIAL CHARACTERS ──');

  await test('WO number with special chars', async () => {
    const woNum = 'EDGE-SPEC/001 (Rev.2)';
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: woNum, contractorId,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.woNumber === woNum, 'Should preserve special chars');
    // Verify retrieval
    const { data: fetched } = await api('GET', `/api/wos/${data.id}`);
    assert(fetched.woNumber === woNum, 'Should retrieve with special chars intact');
    await api('DELETE', `/api/wos/${data.id}`);
  });

  await test('WO location with unicode', async () => {
    const location = 'José Martí Blvd & 3rd Straße';
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-UNICODE-001', contractorId, location,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.location === location, 'Should preserve unicode');
    await api('DELETE', `/api/wos/${data.id}`);
  });

  await test('Employee name with apostrophe', async () => {
    const { data, ok } = await api('POST', '/api/settings/employees', {
      name: "O'Brien-González",
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.name === "O'Brien-González", 'Should preserve apostrophe and accents');
    // Verify listed
    const { data: emps } = await api('GET', '/api/settings/employees');
    const found = emps?.find((e: any) => e.name === "O'Brien-González");
    assert(!!found, 'Should find employee with special name');
    // Cleanup
    if (data.id) await api('PATCH', `/api/settings/employees/${data.id}`, { name: 'TO_DELETE' });
  });

  await test('Contractor name with ampersand', async () => {
    const { data, ok } = await api('POST', '/api/settings/contractors', {
      name: 'Smith & Sons Inc.',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.name === 'Smith & Sons Inc.', 'Should preserve ampersand');
    if (data.id) await api('DELETE', `/api/settings/contractors/${data.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // LONG STRINGS
  // ═══════════════════════════════════════════════════════════════
  log('\n── LONG STRINGS ──');

  await test('WO notes with very long text', async () => {
    const longNote = 'A'.repeat(5000);
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-LONG-001', contractorId, notes: longNote,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.notes?.length === 5000, `Should store long note, got length ${data.notes?.length}`);
    await api('DELETE', `/api/wos/${data.id}`);
  });

  await test('WO location with 500 chars', async () => {
    const longLoc = 'L'.repeat(500);
    const { data, status } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-LONG-002', contractorId, location: longLoc,
    });
    // Should either store or reject gracefully
    assert(status !== 500, `Server error with long location (${status})`);
    if (status === 201 && data.id) await api('DELETE', `/api/wos/${data.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // WO LOOKUP EDGE CASES
  // ═══════════════════════════════════════════════════════════════
  log('\n── WO LOOKUP EDGE CASES ──');

  await test('Get WO by number that looks like UUID prefix', async () => {
    // WO number that starts with hex chars but isn't a UUID
    const { data: wo, ok } = await api('POST', '/api/wos', {
      woNumber: 'ABCDEF-001', contractorId,
    });
    assert(ok, `Failed to create: ${JSON.stringify(wo)}`);

    const { data: found, status } = await api('GET', '/api/wos/ABCDEF-001');
    assert(status === 200, `Should find by WO number, got ${status}`);
    assert(found.woNumber === 'ABCDEF-001', 'Should match WO number');
    await api('DELETE', `/api/wos/${wo.id}`);
  });

  await test('Get nonexistent WO returns 404', async () => {
    const { status } = await api('GET', '/api/wos/NONEXISTENT-999');
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Get WO with empty string', async () => {
    // Trailing slash issue
    const { status } = await api('GET', '/api/wos/');
    // Should match the list route and return 200 (WO list)
    assert(status === 200, `Expected 200 for list, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // MARKING ITEM STATE TRANSITIONS
  // ═══════════════════════════════════════════════════════════════
  log('\n── MARKING STATE TRANSITIONS ──');

  await test('Complete → back to pending', async () => {
    if (!zeroMarkingId) return;
    // First complete it
    await api('PATCH', `/api/markings/${zeroMarkingId}`, { status: 'completed', dateCompleted: '2026-01-01' });
    // Then back to pending
    const { data, ok } = await api('PATCH', `/api/markings/${zeroMarkingId}`, { status: 'pending' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.item?.status === 'pending', 'Should revert to pending');
  });

  await test('Skip marking', async () => {
    if (!zeroMarkingId) return;
    const { data, ok } = await api('PATCH', `/api/markings/${zeroMarkingId}`, { status: 'skipped' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.item?.status === 'skipped', 'Should set to skipped');
  });

  // ═══════════════════════════════════════════════════════════════
  // CONCURRENT-LIKE OPERATIONS
  // ═══════════════════════════════════════════════════════════════
  log('\n── CONCURRENT OPERATIONS ──');

  await test('Parallel WO creation with unique numbers', async () => {
    const promises = [1, 2, 3, 4, 5].map(i =>
      api('POST', '/api/wos', { woNumber: `EDGE-PARALLEL-${i}`, contractorId })
    );
    const results = await Promise.all(promises);
    const successes = results.filter(r => r.ok);
    assert(successes.length === 5, `Expected 5 successes, got ${successes.length}`);

    // Cleanup
    for (const r of results) {
      if (r.data?.id) await api('DELETE', `/api/wos/${r.data.id}`);
    }
  });

  await test('Parallel reads of same WO', async () => {
    const promises = Array.from({ length: 10 }, () =>
      api('GET', `/api/wos/${dupeWo1Id}`)
    );
    const results = await Promise.all(promises);
    assert(results.every(r => r.ok), 'All parallel reads should succeed');
    assert(results.every(r => r.data.woNumber === 'EDGE-DUPE-001'), 'All should return same WO');
  });

  await test('Parallel updates to different fields', async () => {
    const promises = [
      api('PATCH', `/api/wos/${dupeWo1Id}`, { notes: 'Updated via parallel 1' }),
      api('PATCH', `/api/wos/${dupeWo1Id}`, { generalRemarks: 'Updated via parallel 2' }),
    ];
    const results = await Promise.all(promises);
    assert(results.every(r => r.ok), 'Parallel field updates should succeed');
  });

  // ═══════════════════════════════════════════════════════════════
  // WORK ORDER STATUS LIFECYCLE
  // ═══════════════════════════════════════════════════════════════
  log('\n── WO STATUS LIFECYCLE ──');

  await test('Full status cycle: received → dispatched → in_progress → completed → returned', async () => {
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-LIFECYCLE-001', contractorId,
    });
    assert(wo.status === 'received', `Initial should be received, got ${wo.status}`);

    const statuses = ['dispatched', 'in_progress', 'completed', 'returned'];
    for (const status of statuses) {
      const { data, ok } = await api('PATCH', `/api/wos/${wo.id}`, { status });
      assert(ok, `Failed to set ${status}: ${JSON.stringify(data)}`);
      assert(data.status === status, `Expected ${status}, got ${data.status}`);
    }

    await api('DELETE', `/api/wos/${wo.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // SIGN-IN EDGE CASES
  // ═══════════════════════════════════════════════════════════════
  log('\n── SIGN-IN EDGE CASES ──');

  await test('Signin with midnight crossing times', async () => {
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-MIDNIGHT-001', contractorId,
    });

    const { status, data } = await api('POST', '/api/signin', {
      rows: [{
        workDate: '2026-09-01',
        woId: wo.id,
        contractorId,
        employeeName: 'Night Shift Worker',
        classification: 'LP',
        timeIn: '10:00 PM',
        timeOut: '6:00 AM',
      }],
    });
    // Should handle gracefully — might compute negative or wrap-around hours
    assert(status !== 500, `Server error with midnight crossing (${status})`);

    await api('DELETE', `/api/wos/${wo.id}`);
  });

  await test('Signin with same time in and out (zero hours)', async () => {
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-ZEROHRS-001', contractorId,
    });

    const { status } = await api('POST', '/api/signin', {
      rows: [{
        workDate: '2026-09-02',
        woId: wo.id,
        contractorId,
        employeeName: 'Zero Hours Worker',
        classification: 'LP',
        timeIn: '8:00 AM',
        timeOut: '8:00 AM',
      }],
    });
    assert(status !== 500, `Server error with zero hours (${status})`);

    await api('DELETE', `/api/wos/${wo.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // PAGINATION / LARGE RESULT SETS
  // ═══════════════════════════════════════════════════════════════
  log('\n── LARGE RESULT SETS ──');

  await test('Create and list many markings on one WO', async () => {
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: 'EDGE-MANY-001', contractorId,
    });

    // Create 20 markings
    const categories = ['Double Yellow Line', 'Single White Line', 'HVX Crosswalk', 'Stop Msg', 'Lane Arrow'];
    const ids: string[] = [];
    for (let i = 0; i < 20; i++) {
      const { data, ok } = await api('POST', `/api/wos/${wo.id}/markings`, {
        category: categories[i % categories.length],
        quantity: (i + 1) * 10,
        unit: i % 3 === 0 ? 'SF' : i % 3 === 1 ? 'LF' : 'EA',
      });
      if (ok) ids.push(data.id);
    }

    // List all
    const { data: list, ok } = await api('GET', `/api/wos/${wo.id}/markings`);
    assert(ok, 'Should succeed');
    assert(list.items?.length === 20, `Should have 20 items, got ${list.items?.length}`);

    // Bulk delete all
    const { data: delResult } = await api('DELETE', '/api/markings', { ids });
    assert(delResult.count === 20, `Should delete 20, got ${delResult.count}`);

    await api('DELETE', `/api/wos/${wo.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // CLEANUP
  // ═══════════════════════════════════════════════════════════════
  log('\n── CLEANUP ──');

  await test('Delete test WOs', async () => {
    if (dupeWo1Id) await api('DELETE', `/api/wos/${dupeWo1Id}`);
  });

  // Summary
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  Edge Cases: ${passed} passed, ${failed} failed`);
  if (failures.length > 0) {
    console.log(`\n  FAILURES:`);
    for (const f of failures) console.log(`    ✗ ${f}`);
  }
  console.log(`${'═'.repeat(60)}`);

  process.exit(failed > 0 ? 1 : 0);
}

run();
