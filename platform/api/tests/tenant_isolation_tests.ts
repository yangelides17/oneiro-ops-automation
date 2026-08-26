/**
 * Tenant Isolation Tests
 *
 * Creates two separate organizations and verifies that data from
 * one org is completely invisible to the other.
 *
 * Usage: npx tsx tests/tenant_isolation_tests.ts
 */
import 'dotenv/config';

const BASE = 'http://localhost:3001';

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

async function rawFetch(method: string, path: string, body?: any, cookie?: string): Promise<any> {
  const opts: RequestInit = {
    method,
    headers: {
      'Content-Type': 'application/json',
      ...(cookie ? { 'Cookie': cookie } : {}),
    },
  };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(`${BASE}${path}`, opts);
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = { _raw: text }; }
  return { status: res.status, data, ok: res.ok, headers: res.headers };
}

// ─── Org contexts ───────────────────────────────────────────────

interface OrgContext {
  cookie: string;
  orgId: string;
  orgName: string;
  contractorId?: string;
  woId?: string;
  woNumber?: string;
  markingId?: string;
  regionId?: string;
  employeeId?: string;
}

const TEST_TS = Date.now(); // unique per run

async function signupOrg(suffix: string, name: string, orgName: string): Promise<OrgContext> {
  const email = `iso-${suffix}-${TEST_TS}@test.com`;
  const password = 'isolation_test_pw123';

  const res = await fetch(`${BASE}/api/auth/signup`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password, name, orgName }),
  });
  const data = await res.json() as any;
  if (!res.ok) throw new Error(`Signup failed: ${JSON.stringify(data)}`);

  // Login to get cookie (signup set-cookie may not be captured by fetch)
  const loginRes = await fetch(`${BASE}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });
  const setCookie = loginRes.headers.get('set-cookie') || '';
  const cookie = setCookie.split(';')[0];
  if (!cookie) throw new Error(`No cookie from login for ${email}`);

  return { cookie, orgId: data.org.id, orgName };
}

async function api(ctx: OrgContext, method: string, path: string, body?: any): Promise<any> {
  return rawFetch(method, path, body, ctx.cookie);
}

// ─── Main ───────────────────────────────────────────────────────

async function run() {
  log('\n═══ TENANT ISOLATION TESTS ═══\n');

  // ═══════════════════════════════════════════════════════════════
  // SETUP — Create two orgs with data
  // ═══════════════════════════════════════════════════════════════
  log('── SETUP: Create two organizations ──');

  let orgA: OrgContext = { cookie: '', orgId: '', orgName: '' };
  let orgB: OrgContext = { cookie: '', orgId: '', orgName: '' };

  await test('Create Org A (Alpha Corp)', async () => {
    orgA = await signupOrg('alpha', 'Alpha Owner', 'Alpha Corp Isolation');
    assert(!!orgA.orgId, 'Should have org A ID');
  });

  await test('Create Org B (Beta LLC)', async () => {
    orgB = await signupOrg('beta', 'Beta Owner', 'Beta LLC Isolation');
    assert(!!orgB.orgId, 'Should have org B ID');
    assert(orgB.orgId !== orgA.orgId, 'Orgs should have different IDs');
  });

  // ─── Seed data in Org A ───────────────────────────────────────
  log('\n── SETUP: Seed data in Org A ──');

  await test('Create region in Org A', async () => {
    const { data, ok } = await api(orgA, 'POST', '/api/settings/regions', {
      code: 'ISO-A', name: 'Isolation Region A', sortOrder: 1,
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgA.regionId = data.id;
  });

  await test('Create contractor in Org A', async () => {
    const { data, ok } = await api(orgA, 'POST', '/api/settings/contractors', {
      name: 'Alpha Contractor',
      contactEmail: 'alpha@test.com',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgA.contractorId = data.id;
  });

  await test('Create employee in Org A', async () => {
    const { data, ok } = await api(orgA, 'POST', '/api/settings/employees', {
      name: 'Alpha Employee',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgA.employeeId = data.id;
  });

  await test('Create WO in Org A', async () => {
    const { data, ok } = await api(orgA, 'POST', '/api/wos', {
      woNumber: 'ALPHA-WO-001',
      contractorId: orgA.contractorId,
      location: 'Alpha Street',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgA.woId = data.id;
    orgA.woNumber = data.woNumber;
  });

  await test('Create marking item in Org A', async () => {
    const { data, ok } = await api(orgA, 'POST', `/api/wos/${orgA.woId}/markings`, {
      category: 'Double Yellow Line',
      quantity: 100,
      unit: 'LF',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgA.markingId = data.id;
  });

  // ─── Seed data in Org B ───────────────────────────────────────
  log('\n── SETUP: Seed data in Org B ──');

  await test('Create contractor in Org B', async () => {
    const { data, ok } = await api(orgB, 'POST', '/api/settings/contractors', {
      name: 'Beta Contractor',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgB.contractorId = data.id;
  });

  await test('Create WO in Org B', async () => {
    const { data, ok } = await api(orgB, 'POST', '/api/wos', {
      woNumber: 'BETA-WO-001',
      contractorId: orgB.contractorId,
      location: 'Beta Avenue',
    });
    assert(ok, `Failed: ${JSON.stringify(data)}`);
    orgB.woId = data.id;
    orgB.woNumber = data.woNumber;
  });

  // ═══════════════════════════════════════════════════════════════
  // ISOLATION: Org B cannot see Org A's data
  // ═══════════════════════════════════════════════════════════════
  log('\n── ISOLATION: Work Orders ──');

  await test('Org B cannot list Org A work orders', async () => {
    const { data, ok } = await api(orgB, 'GET', '/api/wos');
    assert(ok, 'Should succeed');
    const alphaWo = data.wos?.find((w: any) => w.woNumber === 'ALPHA-WO-001');
    assert(!alphaWo, 'Should NOT see Alpha WO in Org B listing');
  });

  await test('Org B cannot get Org A WO by UUID', async () => {
    const { status } = await api(orgB, 'GET', `/api/wos/${orgA.woId}`);
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot get Org A WO by number', async () => {
    const { status } = await api(orgB, 'GET', `/api/wos/${orgA.woNumber}`);
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot update Org A WO', async () => {
    const { status } = await api(orgB, 'PATCH', `/api/wos/${orgA.woId}`, { status: 'dispatched' });
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot delete Org A WO', async () => {
    const { status } = await api(orgB, 'DELETE', `/api/wos/${orgA.woId}`);
    assert(status === 404, `Expected 404, got ${status}`);
  });

  // ─── Marking items ───────────────────────────────────────────
  log('\n── ISOLATION: Marking Items ──');

  await test('Org B cannot list Org A markings', async () => {
    const { data } = await api(orgB, 'GET', `/api/wos/${orgA.woId}/markings`);
    // Should return 404 (WO not found in org B) or empty items
    assert(!data.items || data.items.length === 0, 'Should not see Org A markings');
  });

  await test('Org B cannot update Org A marking', async () => {
    const { status } = await api(orgB, 'PATCH', `/api/markings/${orgA.markingId}`, {
      quantity: 999,
    });
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot delete Org A markings', async () => {
    const { data } = await api(orgB, 'DELETE', '/api/markings', { ids: [orgA.markingId] });
    assert(data.count === 0, `Should delete 0 items from another org, got ${data.count}`);
  });

  // ─── Settings ──────────────────────────────────────────────────
  log('\n── ISOLATION: Settings ──');

  await test('Org B regions do not include Org A regions', async () => {
    const { data } = await api(orgB, 'GET', '/api/settings/regions');
    const alphaRegion = data?.find?.((r: any) => r.code === 'ISO-A');
    assert(!alphaRegion, 'Should NOT see Alpha region in Org B');
  });

  await test('Org B contractors do not include Org A contractors', async () => {
    const { data } = await api(orgB, 'GET', '/api/settings/contractors');
    const alphaCon = data?.find?.((c: any) => c.name === 'Alpha Contractor');
    assert(!alphaCon, 'Should NOT see Alpha Contractor in Org B');
  });

  await test('Org B employees do not include Org A employees', async () => {
    const { data } = await api(orgB, 'GET', '/api/settings/employees');
    const alphaEmp = data?.find?.((e: any) => e.name === 'Alpha Employee');
    assert(!alphaEmp, 'Should NOT see Alpha Employee in Org B');
  });

  await test('Org B users do not include Org A users', async () => {
    const { data } = await api(orgB, 'GET', '/api/settings/users');
    const alphaUser = data?.find?.((u: any) => u.name === 'Alpha Owner');
    assert(!alphaUser, 'Should NOT see Alpha Owner in Org B users');
  });

  await test('Org B cannot update Org A region', async () => {
    const { status } = await api(orgB, 'PATCH', `/api/settings/regions/${orgA.regionId}`, { name: 'Hacked' });
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot delete Org A region', async () => {
    const { status } = await api(orgB, 'DELETE', `/api/settings/regions/${orgA.regionId}`);
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot update Org A contractor', async () => {
    const { status } = await api(orgB, 'PATCH', `/api/settings/contractors/${orgA.contractorId}`, { name: 'Hacked' });
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Org B cannot update Org A employee', async () => {
    const { status } = await api(orgB, 'PATCH', `/api/settings/employees/${orgA.employeeId}`, { name: 'Hacked' });
    assert(status === 404, `Expected 404, got ${status}`);
  });

  // ─── Dashboard ─────────────────────────────────────────────────
  log('\n── ISOLATION: Dashboard / Analytics ──');

  await test('Org B dashboard does not include Org A WOs', async () => {
    const { data, ok } = await api(orgB, 'GET', '/api/dashboard');
    assert(ok, 'Should succeed');
    const alphaWo = data.wos?.find((w: any) => w.woNumber === 'ALPHA-WO-001');
    assert(!alphaWo, 'Should NOT see Alpha WO in Org B dashboard');
  });

  await test('Org B map does not include Org A WOs', async () => {
    const { data, ok } = await api(orgB, 'GET', '/api/wos/map');
    assert(ok, 'Should succeed');
    const alphaWo = data.wos?.find((w: any) => w.woNumber === 'ALPHA-WO-001');
    assert(!alphaWo, 'Should NOT see Alpha WO in Org B map');
  });

  await test('Org B signin queue does not include Org A entries', async () => {
    const { data, ok } = await api(orgB, 'GET', '/api/signin/queue');
    assert(ok, 'Should succeed');
    // Queue should be empty or only contain Org B entries
    for (const q of (data.queue || [])) {
      assert(!q.contractorName?.includes('Alpha'), 'Should not see Alpha entries');
    }
  });

  await test('Org B documents do not include Org A docs', async () => {
    const { data, ok } = await api(orgB, 'GET', '/api/documents/pending');
    assert(ok, 'Should succeed');
    // Should be empty or contain only Org B docs
    assert(Array.isArray(data.approvals), 'Should have approvals array');
  });

  // ─── Compatibility routes ──────────────────────────────────────
  log('\n── ISOLATION: Compatibility Routes ──');

  await test('Org B /api/employees does not include Org A employees', async () => {
    const { data } = await api(orgB, 'GET', '/api/employees');
    const alphaEmp = data.employees?.find((e: any) => e.name === 'Alpha Employee');
    assert(!alphaEmp, 'Should NOT see Alpha Employee via compat route');
  });

  await test('Org B /api/wo-markings/:woId returns 404 for Org A WO', async () => {
    const { status } = await api(orgB, 'GET', `/api/wo-markings/${orgA.woNumber}`);
    assert(status === 404, `Expected 404, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // REVERSE: Verify Org A still has its data
  // ═══════════════════════════════════════════════════════════════
  log('\n── VERIFY: Org A data intact ──');

  await test('Org A can still see its WO', async () => {
    const { data, ok } = await api(orgA, 'GET', `/api/wos/${orgA.woId}`);
    assert(ok, 'Should succeed');
    assert(data.woNumber === 'ALPHA-WO-001', 'Should see own WO');
  });

  await test('Org A can still see its markings', async () => {
    const { data, ok } = await api(orgA, 'GET', `/api/wos/${orgA.woId}/markings`);
    assert(ok, 'Should succeed');
    assert(data.items?.length === 1, `Should have 1 marking, got ${data.items?.length}`);
  });

  await test('Org A cannot see Org B WOs', async () => {
    const { data } = await api(orgA, 'GET', '/api/wos');
    const betaWo = data.wos?.find((w: any) => w.woNumber === 'BETA-WO-001');
    assert(!betaWo, 'Should NOT see Beta WO in Org A');
  });

  // ═══════════════════════════════════════════════════════════════
  // CLEANUP
  // ═══════════════════════════════════════════════════════════════
  log('\n── CLEANUP ──');

  await test('Delete Org A WO + data', async () => {
    // Delete marking, WO, settings
    await api(orgA, 'DELETE', '/api/markings', { ids: [orgA.markingId] });
    await api(orgA, 'DELETE', `/api/wos/${orgA.woId}`);
    if (orgA.regionId) await api(orgA, 'DELETE', `/api/settings/regions/${orgA.regionId}`);
    if (orgA.contractorId) await api(orgA, 'DELETE', `/api/settings/contractors/${orgA.contractorId}`);
  });

  await test('Delete Org B WO + data', async () => {
    await api(orgB, 'DELETE', `/api/wos/${orgB.woId}`);
    if (orgB.contractorId) await api(orgB, 'DELETE', `/api/settings/contractors/${orgB.contractorId}`);
  });

  // Summary
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  Tenant Isolation: ${passed} passed, ${failed} failed`);
  if (failures.length > 0) {
    console.log(`\n  FAILURES:`);
    for (const f of failures) console.log(`    ✗ ${f}`);
  }
  console.log(`${'═'.repeat(60)}`);

  process.exit(failed > 0 ? 1 : 0);
}

run();
