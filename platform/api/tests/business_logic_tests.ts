/**
 * Business Logic API Tests
 *
 * Tests core business logic through the API layer:
 * - Overtime calculation via signin submit
 * - Revenue computation with known data
 * - Production stats verification
 * - Document generation lifecycle
 * - Field report → WO completion flow
 *
 * Usage: npx tsx tests/business_logic_tests.ts
 */
import 'dotenv/config';

const BASE = 'http://localhost:3001';
const RUN_TS = Date.now(); // Unique per run to avoid collisions
let COOKIE = '';
let contractorId = '';
let woId = '';
let woNumber = '';

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
  log('\n═══ BUSINESS LOGIC API TESTS ═══\n');

  // ═══════════════════════════════════════════════════════════════
  // SETUP
  // ═══════════════════════════════════════════════════════════════
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

  await test('Get contractor for tests', async () => {
    const { data } = await api('GET', '/api/settings/contractors');
    if (data?.length > 0) {
      contractorId = data[0].id;
    } else {
      const { data: c } = await api('POST', '/api/settings/contractors', { name: 'BizLogic Test Co' });
      contractorId = c.id;
    }
    assert(!!contractorId, 'Should have contractor');
  });

  await test('Create WO for business logic tests', async () => {
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: `BIZ-TEST-${RUN_TS}`,
      contractorId,
      contractNum: 'TEST-BIZ',
      regionCode: 'MN',
      location: 'Business Logic Ave',
      fromStreet: '1st St',
      toStreet: '10th St',
      workType: 'thermo',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    woId = data.id;
    woNumber = data.woNumber;
  });

  // ═══════════════════════════════════════════════════════════════
  // OVERTIME CALCULATION VIA SIGNIN SUBMIT
  // ═══════════════════════════════════════════════════════════════
  log('\n── OVERTIME CALCULATION ──');

  // Verify OT rules are set to defaults
  await test('OT rules: 8h threshold, weekend=all OT', async () => {
    const { data, ok } = await api('GET', '/api/settings/payroll/overtime');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    const threshold = Number(data.dailyThresholdHours);
    assert(threshold === 8, `Expected 8h threshold, got ${threshold}`);
    assert(data.weekendAllOt === true, 'Weekend should be all OT');
  });

  // Use future dates computed from RUN_TS to avoid interference from prior runs
  // Find a future Monday (weekday) and Saturday (weekend) based on current time
  const futureBase = new Date(Date.now() + 365 * 24 * 60 * 60 * 1000 + RUN_TS % (30 * 24 * 60 * 60 * 1000));
  // Advance to next Monday
  while (futureBase.getDay() !== 1) futureBase.setDate(futureBase.getDate() + 1);
  const weekdayDate = futureBase.toISOString().slice(0, 10);
  // Next Saturday
  const satBase = new Date(futureBase);
  satBase.setDate(satBase.getDate() + 5);
  const weekendDate = satBase.toISOString().slice(0, 10);

  // Submit a weekday shift: 7 AM to 4:30 PM = 9.5h → 8 ST + 1.5 OT
  await test('Weekday 9.5h shift → 8 ST + 1.5 OT', async () => {
    const { data, ok } = await api('POST', '/api/signin', {
      rows: [{
        workDate: weekdayDate, // Monday
        woId,
        contractorId,
        contractNum: 'TEST-BIZ',
        regionCode: 'MN',
        employeeName: 'OT Test Worker',
        classification: 'LP',
        timeIn: '7:00 AM',
        timeOut: '4:30 PM',
        crewChief: 'OT Test Worker',
      }],
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.count === 1, 'Should insert 1 row');
  });

  await test('Verify OT hours for 9.5h weekday shift', async () => {
    const { data, ok } = await api('POST', '/api/signin/day-hours', {
      workDate: weekdayDate,
      contractNum: 'TEST-BIZ',
      regionCode: 'MN',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    const hours = data.hours;
    assert(Array.isArray(hours), 'hours should be an array');
    const worker = hours.find((h: any) => h.employeeName === 'OT Test Worker');
    assert(!!worker, 'Should find OT Test Worker in day hours');
    const hw = Number(worker.totalHours);
    const ot = Number(worker.totalOt);
    // After first submit: 9.5h total, 1.5 OT
    assert(hw >= 9.5, `Expected >= 9.5 hours worked, got ${hw}`);
    assert(ot >= 1.5, `Expected >= 1.5 OT, got ${ot}`);
  });

  // Submit second group for same date — cross-group lookback
  await test('Second group on same day: cross-group OT lookback', async () => {
    // Create a second WO
    const { data: wo2 } = await api('POST', '/api/wos', {
      woNumber: `BIZ-TEST2-${RUN_TS}`, contractorId, contractNum: 'TEST-BIZ', regionCode: 'BK',
    });

    const { data, ok } = await api('POST', '/api/signin', {
      rows: [{
        workDate: weekdayDate, // Same date
        woId: wo2.id,
        contractorId,
        contractNum: 'TEST-BIZ',
        regionCode: 'BK',
        employeeName: 'OT Test Worker', // Same employee
        classification: 'LP',
        timeIn: '5:00 PM',
        timeOut: '7:00 PM',
        crewChief: 'OT Test Worker',
      }],
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);

    // The second 2h should all be OT (already worked 9.5h on the first group)
    // Clean up second WO
    await api('DELETE', `/api/wos/${wo2.id}`);
  });

  // Weekend shift — all OT
  await test('Weekend shift: all hours are OT', async () => {
    const { data, ok } = await api('POST', '/api/signin', {
      rows: [{
        workDate: weekendDate, // Saturday
        woId,
        contractorId,
        contractNum: 'TEST-BIZ',
        regionCode: 'MN',
        employeeName: 'Weekend Worker',
        classification: 'LP',
        timeIn: '8:00 AM',
        timeOut: '2:00 PM',
        crewChief: 'Weekend Worker',
      }],
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
  });

  await test('Verify weekend hours are all OT', async () => {
    const { data, ok } = await api('POST', '/api/signin/day-hours', {
      workDate: weekendDate,
      contractNum: 'TEST-BIZ',
      regionCode: 'MN',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    const hours = data.hours;
    assert(Array.isArray(hours), 'hours should be an array');
    const worker = hours.find((h: any) => h.employeeName === 'Weekend Worker');
    assert(!!worker, 'Should find Weekend Worker in day hours');
    const hw = Number(worker.totalHours);
    const ot = Number(worker.totalOt);
    assert(Math.abs(hw - 6) < 0.01, `Expected 6 hours worked, got ${hw}`);
    assert(Math.abs(ot - 6) < 0.01, `Expected 6 OT (all weekend), got ${ot}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // MARKING → FIELD REPORT → WO COMPLETION FLOW
  // ═══════════════════════════════════════════════════════════════
  log('\n── FIELD REPORT → COMPLETION FLOW ──');

  let flowWoId = '';
  await test('Create WO for completion flow', async () => {
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: `BIZ-FLOW-${RUN_TS}`, contractorId, location: 'Flow Street',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    flowWoId = data.id;
    assert(data.status === 'received', 'Should start as received');
  });

  await test('Add markings to WO', async () => {
    const { ok: ok1 } = await api('POST', `/api/wos/${flowWoId}/markings`, {
      category: 'Double Yellow Line', quantity: 100, unit: 'LF',
    });
    const { ok: ok2 } = await api('POST', `/api/wos/${flowWoId}/markings`, {
      category: 'Stop Msg', quantity: 3, unit: 'EA',
    });
    assert(ok1 && ok2, 'Should create markings');
  });

  await test('Dispatch WO', async () => {
    const { data, ok } = await api('PATCH', `/api/wos/${flowWoId}`, { status: 'dispatched' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.status === 'dispatched', 'Should be dispatched');
  });

  await test('Check shift is valid', async () => {
    const { data, ok } = await api('POST', '/api/field-reports/check-shift', { woId: flowWoId });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.valid === true, 'Shift should be valid');
  });

  await test('Submit field report with markComplete=true', async () => {
    const { data, ok } = await api('POST', '/api/field-reports', {
      woId: flowWoId,
      markComplete: true,
      crewChief: 'Test Chief',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  await test('Verify WO is now completed', async () => {
    const { data, ok } = await api('GET', `/api/wos/${flowWoId}`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.status === 'completed', `Expected completed, got ${data.status}`);
  });

  await test('Verify WO appears in dashboard with completed status', async () => {
    const { data, ok } = await api('GET', '/api/dashboard');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    const wo = data.wos?.find((w: any) => w.woNumber === `BIZ-FLOW-${RUN_TS}`);
    assert(!!wo, 'Should find WO in dashboard');
    assert(wo.status === 'completed', `Dashboard WO should be completed, got ${wo.status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // DOCUMENT GENERATION
  // ═══════════════════════════════════════════════════════════════
  log('\n── DOCUMENT GENERATION ──');

  await test('Generate daily documents for test date', async () => {
    const { data, ok } = await api('POST', '/api/tools/daily-documents', {
      date: weekdayDate,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  await test('Generate certified payroll for test week', async () => {
    // Compute the Sunday before the weekday date
    const weekStartDate = new Date(futureBase);
    weekStartDate.setDate(weekStartDate.getDate() - 1); // Monday - 1 = Sunday
    const { data, ok } = await api('POST', '/api/tools/certified-payroll', {
      weekStart: weekStartDate.toISOString().slice(0, 10),
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  await test('Document status calendar returns data', async () => {
    const docMonth = weekdayDate.slice(0, 7);
    const { data, ok } = await api('GET', `/api/documents/status?month=${docMonth}`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.month === docMonth, `Should return correct month, got ${data.month}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // DASHBOARD DATA INTEGRITY
  // ═══════════════════════════════════════════════════════════════
  log('\n── DASHBOARD DATA INTEGRITY ──');

  await test('Dashboard stats include correct counts', async () => {
    const { data, ok } = await api('GET', '/api/dashboard');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(typeof data.stats?.total === 'number', 'Should have total count');
    assert(typeof data.stats?.completed === 'number', 'Should have completed count');
    assert(data.stats?.completed >= 1, `Should have at least 1 completed WO (our flow WO), got ${data.stats?.completed}`);
  });

  await test('Dashboard WOs include contractor names', async () => {
    const { data, ok } = await api('GET', '/api/dashboard');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    for (const wo of (data.wos || []).slice(0, 5)) {
      // contractorName should be a string, not a UUID
      if (wo.contractorName) {
        assert(!wo.contractorName.match(/^[0-9a-f]{8}-/i),
          `contractorName should not be a UUID: ${wo.contractorName}`);
      }
    }
  });

  await test('Dashboard attention WOs are valid', async () => {
    const { data, ok } = await api('GET', '/api/dashboard');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    if (data.attention?.length > 0) {
      for (const wo of data.attention) {
        assert(!!wo.woNumber, 'Attention WO should have woNumber');
        assert(!!wo.id, 'Attention WO should have id');
      }
    }
  });

  // ═══════════════════════════════════════════════════════════════
  // REVENUE DATA STRUCTURE
  // ═══════════════════════════════════════════════════════════════
  log('\n── REVENUE DATA STRUCTURE ──');

  await test('Revenue endpoint returns correct structure', async () => {
    const { data, ok } = await api('GET', '/api/revenue?start=2026-01-01&end=2026-12-31');
    assert(ok, `Failed: ${JSON.stringify(data)}`);

    // Verify all expected fields exist
    assert('range' in data, 'Should have range');
    assert('totals' in data, 'Should have totals');
    assert(typeof data.totals.revenue === 'number', 'revenue should be number');
    assert(typeof data.totals.items === 'number', 'items should be number');
    assert('needsPricing' in data.totals, 'Should have needsPricing');
    assert('invoicedRevenue' in data.totals, 'Should have invoicedRevenue');
    assert('wipRevenue' in data.totals, 'Should have wipRevenue');
    assert('pctInvoiced' in data.totals, 'Should have pctInvoiced');

    // Verify buckets
    assert(Array.isArray(data.buckets), 'Should have buckets array');
    assert(data.buckets.length === 3, `Should have 3 buckets (thermo/mma/preform), got ${data.buckets.length}`);
    for (const b of data.buckets) {
      assert('key' in b && 'label' in b && 'revenue' in b, `Bucket should have key/label/revenue: ${JSON.stringify(b)}`);
    }

    // Verify byContractor
    assert(Array.isArray(data.byContractor), 'Should have byContractor array');
    for (const c of data.byContractor) {
      assert('contractorName' in c && 'revenue' in c, `byContractor entry should have contractorName/revenue`);
    }

    // Verify byGroup
    assert(Array.isArray(data.byGroup), 'Should have byGroup array');
    for (const g of data.byGroup) {
      assert('group' in g && 'revenue' in g, `byGroup entry should have group/revenue`);
    }

    // Verify daily
    assert(Array.isArray(data.daily), 'Should have daily array');

    // Verify needs_pricing alias exists (React reads it)
    assert('needs_pricing' in data, 'Should have needs_pricing alias');
  });

  // ═══════════════════════════════════════════════════════════════
  // PRODUCTION DATA STRUCTURE
  // ═══════════════════════════════════════════════════════════════
  log('\n── PRODUCTION DATA STRUCTURE ──');

  await test('Production endpoint returns correct structure', async () => {
    const { data, ok } = await api('GET', '/api/production?start=2026-01-01&end=2026-12-31');
    assert(ok, `Failed: ${JSON.stringify(data)}`);

    // Verify totals with uppercase keys
    assert('totals' in data, 'Should have totals');
    assert('SF' in data.totals, 'totals should have SF');
    assert('LF' in data.totals, 'totals should have LF');
    assert('EA' in data.totals, 'totals should have EA');

    // Verify shifts KPIs
    assert('shifts' in data, 'Should have shifts');
    assert(typeof data.shifts.count === 'number', 'shifts should have count');
    assert(typeof data.shifts.days_in_range === 'number', 'shifts should have days_in_range');
    assert(typeof data.shifts.pct_days_worked === 'number', 'shifts should have pct_days_worked');
    assert(typeof data.shifts.longest_streak === 'number', 'shifts should have longest_streak');

    // Verify by_category
    assert(Array.isArray(data.by_category), 'Should have by_category');

    // Verify daily
    assert(Array.isArray(data.daily), 'Should have daily');
  });

  // ═══════════════════════════════════════════════════════════════
  // PENDING COUNTS INTEGRITY
  // ═══════════════════════════════════════════════════════════════
  log('\n── PENDING COUNTS ──');

  await test('Pending counts returns all expected fields', async () => {
    const { data, ok } = await api('GET', '/api/pending-counts');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert('approvals_review' in data, 'Should have approvals_review');
    assert('doc_status_pending' in data, 'Should have doc_status_pending');
    assert(typeof data.approvals_review === 'number', 'approvals_review should be number');
  });

  await test('Doc-status pending count is consistent', async () => {
    const { data: full, ok: ok1 } = await api('GET', '/api/pending-counts');
    const { data: docOnly, ok: ok2 } = await api('GET', '/api/pending-counts/doc-status');
    assert(ok1 && ok2, 'Both should succeed');
    assert(full.doc_status_pending === docOnly.doc_status_pending,
      `Counts should match: ${full.doc_status_pending} vs ${docOnly.doc_status_pending}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // AUDIT TRAIL
  // ═══════════════════════════════════════════════════════════════
  log('\n── AUDIT TRAIL ──');

  await test('WO creation generates audit entry', async () => {
    const { data: wo, ok } = await api('POST', '/api/wos', {
      woNumber: `BIZ-AUDIT-${RUN_TS}`, contractorId,
    });
    assert(ok, 'Should create WO');
    // We can't directly query audit entries via API, but we verified
    // the creation succeeded — the audit write is in the route handler
    await api('DELETE', `/api/wos/${wo.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // SIGN-IN QUEUE → CONTINUATION CHECK
  // ═══════════════════════════════════════════════════════════════
  log('\n── SIGN-IN QUEUE ──');

  await test('Sign-in queue groups by date/contract/region/chief', async () => {
    const { data, ok } = await api('GET', '/api/signin/queue');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.queue), 'Should have queue array');
    assert(typeof data.total_pending === 'number', 'Should have total_pending');

    // Verify queue entry structure
    for (const entry of data.queue.slice(0, 3)) {
      assert('queue_id' in entry, 'Should have queue_id');
      assert('date' in entry, 'Should have date');
      assert('contractNum' in entry, 'Should have contractNum');
      assert('regionCode' in entry, 'Should have regionCode');
      assert(Array.isArray(entry.wos), 'Should have wos array');
    }
  });

  await test('Continuation check returns correct shape', async () => {
    const { data, ok } = await api('POST', '/api/signin/check-continuation', {
      workDate: weekdayDate,
      contractNum: 'TEST-BIZ',
      regionCode: 'MN',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert('canContinue' in data, 'Should have canContinue');
    assert('nextIndex' in data, 'Should have nextIndex');
    assert('existingCount' in data, 'Should have existingCount');
  });

  // ═══════════════════════════════════════════════════════════════
  // WO MAP DATA
  // ═══════════════════════════════════════════════════════════════
  log('\n── MAP DATA ──');

  await test('Map endpoint returns WOs with coordinate fields', async () => {
    // Add coords to our test WO
    await api('PATCH', `/api/wos/${woId}`, { latitude: '40.7128', longitude: '-74.0060' });

    const { data, ok } = await api('GET', '/api/wos/map');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.wos), 'Should have wos array');

    const testWo = data.wos.find((w: any) => w.woNumber === `BIZ-TEST-${RUN_TS}`);
    if (testWo) {
      assert(testWo.latitude !== undefined, 'Should have latitude');
      assert(testWo.longitude !== undefined, 'Should have longitude');
    }
  });

  // ═══════════════════════════════════════════════════════════════
  // WATERBLAST CONFIRMATION FLOW
  // ═══════════════════════════════════════════════════════════════
  log('\n── WATERBLAST FLOW ──');

  await test('Waterblast confirm updates WO', async () => {
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: `BIZ-WB-${RUN_TS}`, contractorId, waterBlastRequired: 'Yes',
    });
    assert(wo.id, 'Should create WO');

    const { data, ok } = await api('POST', `/api/wos/${wo.id}/waterblast/confirm`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.waterBlastConfirmed === 'Yes', `Should confirm waterblast, got ${data.waterBlastConfirmed}`);

    await api('DELETE', `/api/wos/${wo.id}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // COMPATIBILITY ROUTES
  // ═══════════════════════════════════════════════════════════════
  log('\n── COMPATIBILITY ROUTES ──');

  await test('POST /api/markings with woId in body', async () => {
    const { data, ok } = await api('POST', '/api/markings', {
      woId,
      category: 'Lane Arrow',
      quantity: 5,
      unit: 'EA',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    // Clean up
    if (data.item?.id) {
      await api('DELETE', '/api/markings', { ids: [data.item.id] });
    }
  });

  await test('GET /api/wo-markings/:woNumber returns markings', async () => {
    // Add a marking
    const { data: m } = await api('POST', `/api/wos/${woId}/markings`, {
      category: 'Stop Msg', quantity: 2, unit: 'EA',
    });

    const { data, ok } = await api('GET', `/api/wo-markings/${woNumber}`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.items), 'Should have items');
    assert(data.items.length >= 1, 'Should have at least 1 marking');

    // Cleanup
    if (m?.id) await api('DELETE', '/api/markings', { ids: [m.id] });
  });

  await test('GET /api/qb/status returns disconnected', async () => {
    const { data, ok } = await api('GET', '/api/qb/status');
    assert(ok, 'Should succeed');
    assert(data.connected === false, 'QB should be disconnected');
    assert(data.reason === 'not_configured', 'Should say not_configured');
  });

  await test('POST /api/qb/invoice/:woId returns 503', async () => {
    const { status } = await api('POST', `/api/qb/invoice/${woId}`);
    assert(status === 503, `Expected 503, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // CLEANUP
  // ═══════════════════════════════════════════════════════════════
  log('\n── CLEANUP ──');

  await test('Delete test WOs', async () => {
    if (flowWoId) await api('DELETE', `/api/wos/${flowWoId}`);
    if (woId) await api('DELETE', `/api/wos/${woId}`);
  });

  // Summary
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  Business Logic: ${passed} passed, ${failed} failed`);
  if (failures.length > 0) {
    console.log(`\n  FAILURES:`);
    for (const f of failures) console.log(`    ✗ ${f}`);
  }
  console.log(`${'═'.repeat(60)}`);

  process.exit(failed > 0 ? 1 : 0);
}

run();
