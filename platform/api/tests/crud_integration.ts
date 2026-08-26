/**
 * Comprehensive CRUD Integration Test Suite
 *
 * Exercises EVERY write operation in the system against a sandbox account.
 * Tests create → read → update → delete for every entity type.
 *
 * Usage: npx tsx tests/crud_integration.ts
 */
import 'dotenv/config';

const BASE = 'http://localhost:3001';
let COOKIE = '';
let ORG_ID = '';

let passed = 0;
let failed = 0;
let skipped = 0;
const failures: string[] = [];

function log(msg: string) { console.log(msg); }

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
  return { status: res.status, data, ok: res.ok, headers: res.headers };
}

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

async function run() {
  // ═══════════════════════════════════════════════════════════
  // AUTH
  // ═══════════════════════════════════════════════════════════
  log('\n── AUTH ──');

  let loginRes: any;
  await test('Login to sandbox account', async () => {
    const res = await fetch(`${BASE}/api/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: 'sandbox@oneiro-test.com', password: 'sandbox123' }),
    });
    const setCookie = res.headers.get('set-cookie') || '';
    COOKIE = setCookie.split(';')[0]; // Extract session=...
    loginRes = await res.json();
    assert(res.ok, `Login failed: ${JSON.stringify(loginRes)}`);
    assert(loginRes.user?.role === 'owner', 'Should be owner role');
  });

  await test('GET /api/auth/me', async () => {
    const { data } = await api('GET', '/api/auth/me');
    assert(data.user?.email === 'sandbox@oneiro-test.com', 'Should return user');
    assert(data.org?.name === 'Sandbox Test Org', 'Should return org');
    ORG_ID = data.org.id;
  });

  // ═══════════════════════════════════════════════════════════
  // SETTINGS — CRUD for all config entities
  // ═══════════════════════════════════════════════════════════
  log('\n── SETTINGS: Organization ──');

  await test('PATCH /api/settings/org', async () => {
    const { data, ok } = await api('PATCH', '/api/settings/org', {
      address: '123 Test St, New York, NY 10001',
      phone: '555-1234',
      signatoryName: 'Test Principal',
      signatoryTitle: 'President',
      taxId: '123456789',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.address === '123 Test St, New York, NY 10001', 'Address should update');
    assert(data.signatoryName === 'Test Principal', 'Signatory should update');
  });

  // ── Regions ──
  log('\n── SETTINGS: Regions ──');
  let regionId = '';

  await test('GET /api/settings/regions (seeded)', async () => {
    const { data } = await api('GET', '/api/settings/regions');
    assert(Array.isArray(data) && data.length === 5, `Should have 5 seeded regions, got ${data?.length}`);
  });

  await test('POST /api/settings/regions (create)', async () => {
    const { data, ok } = await api('POST', '/api/settings/regions', { code: 'TEST', name: 'Test Region', sortOrder: 99 });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.code === 'TEST', 'Should create region');
    regionId = data.id;
  });

  await test('PATCH /api/settings/regions/:id (update)', async () => {
    const { data, ok } = await api('PATCH', `/api/settings/regions/${regionId}`, { name: 'Updated Region' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.name === 'Updated Region', 'Should update name');
  });

  await test('DELETE /api/settings/regions/:id', async () => {
    const { ok } = await api('DELETE', `/api/settings/regions/${regionId}`);
    assert(ok, 'Should delete');
  });

  // ── Contractors ──
  log('\n── SETTINGS: Contractors ──');
  let contractorId = '';

  await test('POST /api/settings/contractors', async () => {
    const { data, ok } = await api('POST', '/api/settings/contractors', {
      name: 'Test Prime Contractor',
      contactEmail: 'prime@test.com',
      autoGeneratePl: true,
      receivesPl: true,
      receivesCfr: true,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.name === 'Test Prime Contractor', 'Should create contractor');
    contractorId = data.id;
  });

  await test('PATCH /api/settings/contractors/:id', async () => {
    const { data, ok } = await api('PATCH', `/api/settings/contractors/${contractorId}`, { contactPhone: '555-9999' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.contactPhone === '555-9999', 'Should update phone');
  });

  // ── Employees ──
  log('\n── SETTINGS: Employees ──');
  let employeeId = '';

  await test('POST /api/settings/employees', async () => {
    const { data, ok } = await api('POST', '/api/settings/employees', {
      name: 'John Test Worker',
      address: '456 Worker St',
      ssnLast4: '1234',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    employeeId = data.id;
  });

  await test('GET /api/employees (dropdown)', async () => {
    const { data, ok } = await api('GET', '/api/employees');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.employees?.length > 0, 'Should have employees');
    assert(data.employees.some((e: any) => e.name === 'John Test Worker'), 'Should include new employee');
  });

  // ── Contract Pricing ──
  log('\n── SETTINGS: Contract Pricing ──');
  let pricingId = '';

  await test('POST /api/settings/pricing', async () => {
    const { data, ok } = await api('POST', '/api/settings/pricing', {
      contractorId,
      contractNum: 'TEST-001',
      regionCode: 'BK',
      rateLine4: '0.70',
      rateLine12: '1.45',
      ratePreformed: '120.00',
      rateExtruded: '90.00',
      rateColorSurface: '3.40',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    pricingId = data.id;
  });

  // ── Pay Classifications + Rates ──
  log('\n── SETTINGS: Payroll ──');

  await test('GET /api/settings/payroll/classifications', async () => {
    const { data } = await api('GET', '/api/settings/payroll/classifications');
    assert(Array.isArray(data) && data.length === 6, `Should have 6 classifications, got ${data?.length}`);
  });

  await test('GET /api/settings/payroll/rates', async () => {
    const { data } = await api('GET', '/api/settings/payroll/rates');
    assert(Array.isArray(data) && data.length >= 2, 'Should have seeded rates');
  });

  await test('GET /api/settings/payroll/overtime', async () => {
    const { data, ok } = await api('GET', '/api/settings/payroll/overtime');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data?.dailyThresholdHours === '8.00' || data?.dailyThresholdHours === '8', 'Should have 8h threshold');
    assert(data?.weekendAllOt === true, 'Should have weekend OT');
  });

  await test('PATCH /api/settings/payroll/overtime', async () => {
    const { data, ok } = await api('PATCH', '/api/settings/payroll/overtime', { dailyThresholdHours: '10' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    // Reset back
    await api('PATCH', '/api/settings/payroll/overtime', { dailyThresholdHours: '8' });
  });

  // ── Contract Lookup ──
  log('\n── SETTINGS: Contract Lookup ──');

  await test('POST /api/settings/contract-lookup', async () => {
    const { data, ok } = await api('POST', '/api/settings/contract-lookup', {
      contractNum: 'TEST-001',
      regionCode: 'BK',
      regionName: 'Brooklyn',
      contractId: 'REG-12345',
      projectName: 'Test Project',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
  });

  // ── Billing Remaps ──
  log('\n── SETTINGS: Billing Remaps ──');
  let remapId = '';

  await test('POST /api/settings/billing-remaps', async () => {
    const { data, ok } = await api('POST', '/api/settings/billing-remaps', {
      sourceContract: 'TEST-001',
      sourceRegion: 'BK',
      targetContract: 'TEST-001',
      targetRegion: 'MN',
      effectiveDate: '2026-01-01',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    remapId = data.id;
  });

  await test('DELETE /api/settings/billing-remaps/:id', async () => {
    const { ok } = await api('DELETE', `/api/settings/billing-remaps/${remapId}`);
    assert(ok, 'Should delete remap');
  });

  // ── User Management ──
  log('\n── SETTINGS: Users ──');

  await test('GET /api/settings/users', async () => {
    const { data, ok } = await api('GET', '/api/settings/users');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data) && data.length >= 1, 'Should have at least 1 user');
  });

  await test('POST /api/settings/users/invite', async () => {
    const { data, ok } = await api('POST', '/api/settings/users/invite', {
      email: 'invited@test.com',
      role: 'crew',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
  });

  // ═══════════════════════════════════════════════════════════
  // WORK ORDERS — Full CRUD
  // ═══════════════════════════════════════════════════════════
  log('\n── WORK ORDERS ──');
  let woId = '';
  let woNumber = '';

  await test('POST /api/wos (create WO)', async () => {
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: 'RM-TEST-001',
      contractorId,
      contractNum: 'TEST-001',
      regionCode: 'BK',
      location: 'Test Avenue',
      fromStreet: '1st St',
      toStreet: '5th St',
      workType: 'thermo',
      priority: '3 - Schedule',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.woNumber === 'RM-TEST-001', 'Should have WO number');
    assert(data.status === 'received', 'Should be received status');
    woId = data.id;
    woNumber = data.woNumber;
  });

  await test('GET /api/wos (list)', async () => {
    const { data, ok } = await api('GET', '/api/wos');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.wos?.length >= 1, 'Should have at least 1 WO');
    assert(data.wos.some((w: any) => w.woNumber === 'RM-TEST-001'), 'Should include new WO');
  });

  await test('GET /api/wos/:id (by UUID)', async () => {
    const { data, ok } = await api('GET', `/api/wos/${woId}`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.woNumber === 'RM-TEST-001', 'Should return WO');
  });

  await test('GET /api/wos/:woNumber (by WO number)', async () => {
    const { data, ok } = await api('GET', `/api/wos/${woNumber}`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.woNumber === 'RM-TEST-001', 'Should find by WO number');
  });

  await test('PATCH /api/wos/:id (update status)', async () => {
    const { data, ok } = await api('PATCH', `/api/wos/${woId}`, { status: 'dispatched' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.status === 'dispatched', 'Should update status');
  });

  await test('PATCH /api/wos/:id (update coordinates)', async () => {
    const { data, ok } = await api('PATCH', `/api/wos/${woId}`, {
      latitude: '40.6892',
      longitude: '-74.0445',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
  });

  await test('POST /api/wos/:id/waterblast/confirm', async () => {
    const { data, ok } = await api('POST', `/api/wos/${woId}/waterblast/confirm`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.waterBlastConfirmed === 'Yes', 'Should confirm waterblast');
  });

  // ═══════════════════════════════════════════════════════════
  // MARKING ITEMS — Full CRUD
  // ═══════════════════════════════════════════════════════════
  log('\n── MARKING ITEMS ──');
  let markingItemId = '';

  await test('POST /api/wos/:woId/markings (create)', async () => {
    const { data, ok } = await api('POST', `/api/wos/${woId}/markings`, {
      category: 'Double Yellow Line',
      quantity: 250,
      unit: 'LF',
      woSection: 'top_table',
      addedBy: 'manual',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.category === 'Double Yellow Line', 'Should create marking item');
    assert(data.quantity === '250.00', 'Should store quantity');
    markingItemId = data.id;
  });

  let markingItem2Id = '';
  await test('POST /api/wos/:woId/markings (create HVX)', async () => {
    const { data, ok } = await api('POST', `/api/wos/${woId}/markings`, {
      category: 'HVX Crosswalk',
      quantity: 75,
      unit: 'LF',
      woSection: 'intersection_grid',
      intersection: 'Test Ave & 3rd St',
      direction: 'E',
      addedBy: 'scanner',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    markingItem2Id = data.id;
  });

  let markingItem3Id = '';
  await test('POST /api/wos/:woId/markings (create Stop Msg)', async () => {
    const { data, ok } = await api('POST', `/api/wos/${woId}/markings`, {
      category: 'Stop Msg',
      quantity: 2,
      unit: 'EA',
      woSection: 'intersection_grid',
      intersection: 'Test Ave & 3rd St',
      addedBy: 'scanner',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    markingItem3Id = data.id;
  });

  await test('GET /api/wos/:woId/markings (list)', async () => {
    const { data, ok } = await api('GET', `/api/wos/${woId}/markings`);
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.items?.length === 3, `Should have 3 items, got ${data.items?.length}`);
  });

  await test('PATCH /api/markings/:id (complete item)', async () => {
    const { data, ok } = await api('PATCH', `/api/markings/${markingItemId}`, {
      status: 'completed',
      dateCompleted: '2026-08-25',
      quantity: 275,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.item.status === 'completed', 'Should mark completed');
    assert(data.item.dateCompleted === '2026-08-25', 'Should set date');
    assert(data.item.quantity === '275.00', 'Should update quantity');
  });

  await test('DELETE /api/markings (bulk delete)', async () => {
    const { data, ok } = await api('DELETE', '/api/markings', { ids: [markingItem3Id] });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.count === 1, 'Should delete 1 item');
  });

  await test('GET /api/wos/:woId/markings (verify 2 remain)', async () => {
    const { data } = await api('GET', `/api/wos/${woId}/markings`);
    assert(data.items?.length === 2, `Should have 2 items after delete, got ${data.items?.length}`);
  });

  // ═══════════════════════════════════════════════════════════
  // FIELD REPORT — Submit
  // ═══════════════════════════════════════════════════════════
  log('\n── FIELD REPORTS ──');

  await test('POST /api/field-reports/check-shift', async () => {
    const { data, ok } = await api('POST', '/api/field-reports/check-shift', { woId });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.valid === true, 'Should be valid');
  });

  await test('POST /api/field-reports (submit, mark complete)', async () => {
    const { data, ok } = await api('POST', '/api/field-reports', {
      woId,
      markComplete: true,
      crewChief: 'John Test Worker',
      issues: '',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  await test('Verify WO status is completed', async () => {
    const { data } = await api('GET', `/api/wos/${woId}`);
    assert(data.status === 'completed', `Should be completed, got ${data.status}`);
  });

  // ═══════════════════════════════════════════════════════════
  // SIGN-IN — Submit crew data with OT calculation
  // ═══════════════════════════════════════════════════════════
  log('\n── SIGN-IN ──');

  await test('GET /api/signin/queue', async () => {
    const { data, ok } = await api('GET', '/api/signin/queue');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.queue), 'Should have queue array');
  });

  await test('POST /api/signin (submit with OT calc)', async () => {
    const { data, ok } = await api('POST', '/api/signin', {
      rows: [
        {
          workDate: '2026-08-25',
          woId,
          contractorId,
          contractNum: 'TEST-001',
          regionCode: 'BK',
          location: 'Test Avenue',
          employeeName: 'John Test Worker',
          classification: 'LP',
          timeIn: '7:00 AM',
          timeOut: '4:30 PM',
          crewChief: 'John Test Worker',
        },
      ],
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.count === 1, 'Should insert 1 row');
  });

  await test('POST /api/signin/day-hours', async () => {
    const { data, ok } = await api('POST', '/api/signin/day-hours', {
      workDate: '2026-08-25',
      contractNum: 'TEST-001',
      regionCode: 'BK',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
  });

  await test('POST /api/signin/check-continuation', async () => {
    const { data, ok } = await api('POST', '/api/signin/check-continuation', {
      workDate: '2026-08-25',
      contractNum: 'TEST-001',
      regionCode: 'BK',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.canContinue === true, 'Should allow continuation');
  });

  // ═══════════════════════════════════════════════════════════
  // DOCUMENT GENERATION
  // ═══════════════════════════════════════════════════════════
  log('\n── DOCUMENT GENERATION ──');

  await test('POST /api/tools/daily-documents', async () => {
    const { data, ok } = await api('POST', '/api/tools/daily-documents', { date: '2026-08-25' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  await test('POST /api/tools/certified-payroll', async () => {
    const { data, ok } = await api('POST', '/api/tools/certified-payroll', { weekStart: '2026-08-23' });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  await test('POST /api/tools/process-approved', async () => {
    const { data, ok } = await api('POST', '/api/tools/process-approved');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.ok === true, 'Should succeed');
  });

  // ═══════════════════════════════════════════════════════════
  // DOCUMENTS — Approval workflow
  // ═══════════════════════════════════════════════════════════
  log('\n── DOCUMENTS ──');

  await test('GET /api/documents/pending', async () => {
    const { data, ok } = await api('GET', '/api/documents/pending');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.approvals), 'Should have approvals array');
  });

  await test('GET /api/documents/pending/counts', async () => {
    const { data, ok } = await api('GET', '/api/documents/pending/counts');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert('approvals_review' in data, 'Should have approvals_review count');
  });

  await test('GET /api/documents/status (calendar)', async () => {
    const { data, ok } = await api('GET', '/api/documents/status?month=2026-08');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.month === '2026-08', 'Should return correct month');
  });

  // ═══════════════════════════════════════════════════════════
  // DASHBOARD — Read operations
  // ═══════════════════════════════════════════════════════════
  log('\n── DASHBOARD ──');

  await test('GET /api/dashboard', async () => {
    const { data, ok } = await api('GET', '/api/dashboard');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.wos?.length >= 1, 'Should have WOs');
    assert(data.stats?.total >= 1, 'Should have stats');
    // Verify contractor name is included (not just UUID)
    const wo = data.wos[0];
    assert(wo.contractorName, `WO should have contractorName, got: ${JSON.stringify(Object.keys(wo))}`);
  });

  await test('GET /api/revenue', async () => {
    const { data, ok } = await api('GET', '/api/revenue?start=2026-08-01&end=2026-08-31');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert('totals' in data, 'Should have totals');
  });

  await test('GET /api/production', async () => {
    const { data, ok } = await api('GET', '/api/production?start=2026-08-01&end=2026-08-31');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert('totals' in data, 'Should have totals');
    assert('shifts' in data, 'Should have shifts KPIs');
    assert('by_category' in data, 'Should have category breakdown');
  });

  await test('GET /api/pending-counts', async () => {
    const { data, ok } = await api('GET', '/api/pending-counts');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert('approvals_review' in data, 'Should have counts');
  });

  // ═══════════════════════════════════════════════════════════
  // WO SCANNING
  // ═══════════════════════════════════════════════════════════
  log('\n── WO SCANNING ──');

  await test('GET /api/tools/scan-status', async () => {
    const { data, ok } = await api('GET', '/api/tools/scan-status');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.statuses), 'Should have statuses array');
  });

  await test('GET /api/tools/scan-uploads-today', async () => {
    const { data, ok } = await api('GET', '/api/tools/scan-uploads-today');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.uploads), 'Should have uploads array');
  });

  // ═══════════════════════════════════════════════════════════
  // MAP
  // ═══════════════════════════════════════════════════════════
  log('\n── MAP ──');

  await test('GET /api/wos/map', async () => {
    const { data, ok } = await api('GET', '/api/wos/map');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data.wos), 'Should have wos array');
  });

  // ═══════════════════════════════════════════════════════════
  // GEOCODING
  // ═══════════════════════════════════════════════════════════
  log('\n── GEOCODING ──');

  await test('POST /api/geocode', async () => {
    const { data, ok } = await api('POST', '/api/geocode', { address: 'Times Square, New York' });
    // May fail without Google Maps API key configured, that's OK
    if (!ok && data.error?.includes('GOOGLE_MAPS_API_KEY')) {
      skipped++;
      console.log('    (skipped — no Google Maps API key)');
      return;
    }
    assert(ok, `Failed: ${JSON.stringify(data)}`);
  });

  // ═══════════════════════════════════════════════════════════
  // QB STUB
  // ═══════════════════════════════════════════════════════════
  log('\n── QUICKBOOKS (deferred) ──');

  await test('GET /api/qb/status (stub)', async () => {
    const { data, ok } = await api('GET', '/api/qb/status');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(data.connected === false, 'Should be disconnected');
  });

  // ═══════════════════════════════════════════════════════════
  // INTEGRATIONS
  // ═══════════════════════════════════════════════════════════
  log('\n── INTEGRATIONS ──');

  await test('GET /api/integrations', async () => {
    const { data, ok } = await api('GET', '/api/integrations');
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    assert(Array.isArray(data), 'Should be array');
  });

  // ═══════════════════════════════════════════════════════════
  // CLEANUP — Delete test WO
  // ═══════════════════════════════════════════════════════════
  log('\n── CLEANUP ──');

  // Create a second WO to test delete
  let deleteWoId = '';
  await test('POST /api/wos (create WO for deletion)', async () => {
    const { data, ok } = await api('POST', '/api/wos', {
      woNumber: 'RM-DELETE-ME',
      contractorId,
      location: 'Delete Test',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    deleteWoId = data.id;
  });

  await test('DELETE /api/wos/:id', async () => {
    const { ok } = await api('DELETE', `/api/wos/${deleteWoId}`);
    assert(ok, 'Should delete WO');
  });

  await test('GET /api/wos/:id (verify deleted)', async () => {
    const { status } = await api('GET', `/api/wos/${deleteWoId}`);
    assert(status === 404, 'Should be 404 after delete');
  });

  // ═══════════════════════════════════════════════════════════
  // AUTH — Logout
  // ═══════════════════════════════════════════════════════════
  log('\n── AUTH CLEANUP ──');

  await test('POST /api/auth/logout', async () => {
    const { ok } = await api('POST', '/api/auth/logout');
    assert(ok, 'Should logout');
  });

  await test('GET /api/auth/me (after logout, should 401)', async () => {
    const { status } = await api('GET', '/api/auth/me');
    assert(status === 401, 'Should be 401 after logout');
  });

  // ═══════════════════════════════════════════════════════════
  // SUMMARY
  // ═══════════════════════════════════════════════════════════
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  ${passed} passed, ${failed} failed, ${skipped} skipped`);
  if (failures.length > 0) {
    console.log(`\n  FAILURES:`);
    for (const f of failures) console.log(`    ✗ ${f}`);
  }
  console.log(`${'═'.repeat(60)}`);

  process.exit(failed > 0 ? 1 : 0);
}

run();
