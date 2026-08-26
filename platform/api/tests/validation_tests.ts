/**
 * Input Validation Tests
 *
 * Tests malformed payloads, missing required fields, invalid UUIDs,
 * wrong types, SQL injection, and XSS payloads across all endpoints.
 *
 * Usage: npx tsx tests/validation_tests.ts
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

// ─── Main ───────────────────────────────────────────────────────

async function run() {
  log('\n═══ INPUT VALIDATION TESTS ═══\n');

  // Setup — login to sandbox
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

  // Get a valid contractor ID for tests that need one
  await test('Get existing contractor for reference', async () => {
    const { data } = await api('GET', '/api/settings/contractors');
    if (data?.length > 0) contractorId = data[0].id;
    // If no contractors, create one
    if (!contractorId) {
      const { data: c } = await api('POST', '/api/settings/contractors', { name: 'Validation Test Co' });
      contractorId = c.id;
    }
    assert(!!contractorId, 'Should have a contractor ID');
  });

  // ═══════════════════════════════════════════════════════════════
  // AUTH — Signup validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── AUTH: Signup Validation ──');

  await test('Signup: missing email', async () => {
    const { status } = await api('POST', '/api/auth/signup', {
      password: 'testtest123', name: 'Test', orgName: 'Test Org',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signup: invalid email', async () => {
    const { status } = await api('POST', '/api/auth/signup', {
      email: 'not-an-email', password: 'testtest123', name: 'Test', orgName: 'Test',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signup: short password', async () => {
    const { status } = await api('POST', '/api/auth/signup', {
      email: 'new@test.com', password: 'short', name: 'Test', orgName: 'Test',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signup: missing orgName', async () => {
    const { status } = await api('POST', '/api/auth/signup', {
      email: 'new@test.com', password: 'testtest123', name: 'Test',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signup: empty name', async () => {
    const { status } = await api('POST', '/api/auth/signup', {
      email: 'new@test.com', password: 'testtest123', name: '', orgName: 'Test',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signup: duplicate email', async () => {
    const { status } = await api('POST', '/api/auth/signup', {
      email: 'sandbox@oneiro-test.com', password: 'testtest123', name: 'Test', orgName: 'Dup Org',
    });
    assert(status === 409, `Expected 409, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // AUTH — Login validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── AUTH: Login Validation ──');

  await test('Login: missing email', async () => {
    const { status } = await api('POST', '/api/auth/login', { password: 'test' });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Login: wrong password', async () => {
    const { status } = await api('POST', '/api/auth/login', {
      email: 'sandbox@oneiro-test.com', password: 'wrongpassword',
    });
    assert(status === 401, `Expected 401, got ${status}`);
  });

  await test('Login: nonexistent email', async () => {
    const { status } = await api('POST', '/api/auth/login', {
      email: 'nonexistent@test.com', password: 'testtest123',
    });
    assert(status === 401, `Expected 401, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // WORK ORDERS — Create validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── WORK ORDERS: Create Validation ──');

  await test('Create WO: missing woNumber', async () => {
    const { status } = await api('POST', '/api/wos', { contractorId });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Create WO: missing contractorId', async () => {
    const { status } = await api('POST', '/api/wos', { woNumber: 'VAL-001' });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Create WO: invalid contractorId (not UUID)', async () => {
    const { status } = await api('POST', '/api/wos', {
      woNumber: 'VAL-001', contractorId: 'not-a-uuid',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Create WO: empty woNumber', async () => {
    const { status } = await api('POST', '/api/wos', {
      woNumber: '', contractorId,
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Create WO: empty body', async () => {
    const { status } = await api('POST', '/api/wos', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ─── WO status validation ──────────────────────────────────────
  await test('Update WO: invalid status value', async () => {
    // First create a valid WO
    const { data: wo } = await api('POST', '/api/wos', {
      woNumber: 'VAL-STATUS-001', contractorId,
    });
    if (wo.id) {
      const { status } = await api('PATCH', `/api/wos/${wo.id}`, { status: 'invalid_status' });
      assert(status === 400, `Expected 400, got ${status}`);
      // Cleanup
      await api('DELETE', `/api/wos/${wo.id}`);
    }
  });

  // ═══════════════════════════════════════════════════════════════
  // MARKING ITEMS — Validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── MARKING ITEMS: Validation ──');

  await test('Create marking: missing category', async () => {
    // Use a fake WO UUID — the validation should fail before DB lookup
    const fakeWo = '00000000-0000-0000-0000-000000000001';
    const { status } = await api('POST', `/api/wos/${fakeWo}/markings`, {
      quantity: 100, unit: 'LF',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Create marking: invalid unit', async () => {
    const fakeWo = '00000000-0000-0000-0000-000000000001';
    const { status } = await api('POST', `/api/wos/${fakeWo}/markings`, {
      category: 'Test', quantity: 100, unit: 'INVALID',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Update marking: invalid status', async () => {
    const fakeId = '00000000-0000-0000-0000-000000000001';
    const { status } = await api('PATCH', `/api/markings/${fakeId}`, {
      status: 'invalid_status',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Delete markings: missing ids', async () => {
    const { status } = await api('DELETE', '/api/markings', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Delete markings: empty ids array', async () => {
    const { status } = await api('DELETE', '/api/markings', { ids: [] });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Delete markings: non-UUID ids', async () => {
    const { status } = await api('DELETE', '/api/markings', { ids: ['not-a-uuid'] });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // SIGN-IN — Validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── SIGN-IN: Validation ──');

  await test('Submit signin: empty rows', async () => {
    const { status } = await api('POST', '/api/signin', { rows: [] });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Submit signin: missing rows', async () => {
    const { status } = await api('POST', '/api/signin', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Submit signin: row missing employeeName', async () => {
    const { status } = await api('POST', '/api/signin', {
      rows: [{
        workDate: '2026-01-01',
        woId: '00000000-0000-0000-0000-000000000001',
        contractorId: '00000000-0000-0000-0000-000000000001',
        classification: 'LP',
      }],
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Submit signin: row missing classification', async () => {
    const { status } = await api('POST', '/api/signin', {
      rows: [{
        workDate: '2026-01-01',
        woId: '00000000-0000-0000-0000-000000000001',
        contractorId: '00000000-0000-0000-0000-000000000001',
        employeeName: 'Test Worker',
      }],
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signin day-hours: missing workDate', async () => {
    const { status } = await api('POST', '/api/signin/day-hours', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Signin check-continuation: missing workDate', async () => {
    const { status } = await api('POST', '/api/signin/check-continuation', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // FIELD REPORTS — Validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── FIELD REPORTS: Validation ──');

  await test('Submit field report: missing woId', async () => {
    const { status } = await api('POST', '/api/field-reports', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Submit field report: invalid woId', async () => {
    const { status } = await api('POST', '/api/field-reports', { woId: 'not-uuid' });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Submit field report: nonexistent woId', async () => {
    const { status } = await api('POST', '/api/field-reports', {
      woId: '00000000-0000-0000-0000-999999999999',
    });
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('Check shift: missing woId', async () => {
    const { status } = await api('POST', '/api/field-reports/check-shift', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // SETTINGS — Validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── SETTINGS: Validation ──');

  await test('Create region: missing code', async () => {
    const { status } = await api('POST', '/api/settings/regions', { name: 'Test' });
    // Should fail — code is required
    assert(status === 400 || status === 500, `Expected 400/500, got ${status}`);
  });

  await test('Invite user: invalid role', async () => {
    const { status } = await api('POST', '/api/settings/users/invite', {
      email: 'test@test.com', role: 'superadmin',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Invite user: missing email', async () => {
    const { status } = await api('POST', '/api/settings/users/invite', { role: 'crew' });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Invite user: invalid email format', async () => {
    const { status } = await api('POST', '/api/settings/users/invite', {
      email: 'not-email', role: 'crew',
    });
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // TOOLS — Validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── TOOLS: Validation ──');

  await test('Daily documents: missing date', async () => {
    const { status } = await api('POST', '/api/tools/daily-documents', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Certified payroll: missing weekStart', async () => {
    const { status } = await api('POST', '/api/tools/certified-payroll', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // DOCUMENTS — Validation
  // ═══════════════════════════════════════════════════════════════
  log('\n── DOCUMENTS: Validation ──');

  await test('Doc flags: missing id', async () => {
    const { status } = await api('POST', '/api/documents/flags', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('Doc status/flags (old route): missing id', async () => {
    const { status } = await api('POST', '/api/documents/status/flags', {});
    assert(status === 400, `Expected 400, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // SQL INJECTION ATTEMPTS
  // ═══════════════════════════════════════════════════════════════
  log('\n── SECURITY: SQL Injection ──');

  const sqlPayloads = [
    "'; DROP TABLE users; --",
    "1' OR '1'='1",
    "1; SELECT * FROM users --",
    "' UNION SELECT * FROM organizations --",
    "'; UPDATE users SET role='owner' WHERE email='test@test.com'; --",
  ];

  for (const payload of sqlPayloads) {
    await test(`SQL injection in WO number: ${payload.slice(0, 30)}...`, async () => {
      const { status } = await api('GET', `/api/wos/${encodeURIComponent(payload)}`);
      // Should return 404 (not found) — not crash or return unauthorized data
      assert(status === 404 || status === 400, `Expected 404 or 400, got ${status}`);
    });
  }

  await test('SQL injection in query param', async () => {
    const { status } = await api('GET', "/api/documents/status?month=2026-08'; DROP TABLE documents; --");
    // Should not crash
    assert(status !== 500, `Server error — possible SQL injection vulnerability (${status})`);
  });

  await test('SQL injection in JSON body', async () => {
    const { status } = await api('POST', '/api/wos', {
      woNumber: "'; DROP TABLE work_orders; --",
      contractorId,
    });
    // Should either create with the literal string or reject — not crash
    if (status === 201) {
      // Created with literal string — that's safe. Clean up.
      const { data } = await api('GET', '/api/wos');
      const wo = data.wos?.find((w: any) => w.woNumber?.includes('DROP'));
      if (wo) await api('DELETE', `/api/wos/${wo.id}`);
    }
    assert(status !== 500, `Server error — possible SQL injection (${status})`);
  });

  // ═══════════════════════════════════════════════════════════════
  // XSS ATTEMPTS
  // ═══════════════════════════════════════════════════════════════
  log('\n── SECURITY: XSS Payloads ──');

  const xssPayloads = [
    '<script>alert("xss")</script>',
    '<img src=x onerror=alert(1)>',
    '"><script>document.location="http://evil.com"</script>',
    "javascript:alert('xss')",
  ];

  for (const payload of xssPayloads) {
    await test(`XSS in WO number: ${payload.slice(0, 30)}...`, async () => {
      const { data, status } = await api('POST', '/api/wos', {
        woNumber: payload,
        contractorId,
      });
      // Should store as literal text (Drizzle parameterized queries) — not execute
      assert(status !== 500, `Server error with XSS payload (${status})`);
      if (status === 201 && data.id) {
        // Verify stored as literal
        const { data: wo } = await api('GET', `/api/wos/${data.id}`);
        assert(wo.woNumber === payload, 'Should store XSS as literal text');
        await api('DELETE', `/api/wos/${data.id}`);
      }
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // INVALID UUID FORMAT
  // ═══════════════════════════════════════════════════════════════
  log('\n── INVALID UUID HANDLING ──');

  await test('GET /api/wos/not-a-uuid (no crash)', async () => {
    const { status } = await api('GET', '/api/wos/not-a-uuid');
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('PATCH /api/markings/not-a-uuid (no crash)', async () => {
    const { status } = await api('PATCH', '/api/markings/not-a-uuid', { quantity: 1 });
    assert(status !== 500, `Server error with invalid UUID (${status})`);
  });

  await test('GET /api/wos/undefined (no crash)', async () => {
    const { status } = await api('GET', '/api/wos/undefined');
    assert(status === 404, `Expected 404, got ${status}`);
  });

  await test('GET /api/wos/null (no crash)', async () => {
    const { status } = await api('GET', '/api/wos/null');
    assert(status === 404, `Expected 404, got ${status}`);
  });

  // ═══════════════════════════════════════════════════════════════
  // CONTENT TYPE / MALFORMED BODY
  // ═══════════════════════════════════════════════════════════════
  log('\n── MALFORMED REQUEST BODIES ──');

  await test('POST with non-JSON body', async () => {
    const res = await fetch(`${BASE}/api/wos`, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain',
        'Cookie': COOKIE,
      },
      body: 'this is not json',
    });
    // Should reject gracefully
    assert(res.status !== 500, `Server error with non-JSON body (${res.status})`);
  });

  await test('POST with empty body', async () => {
    const { status } = await api('POST', '/api/wos');
    assert(status === 400, `Expected 400, got ${status}`);
  });

  await test('PATCH with null body values', async () => {
    const fakeId = '00000000-0000-0000-0000-000000000001';
    const { status } = await api('PATCH', `/api/wos/${fakeId}`, {
      status: null, location: null,
    });
    // Should not crash
    assert(status !== 500, `Server error with null values (${status})`);
  });

  // ═══════════════════════════════════════════════════════════════
  // EXPIRED / TAMPERED TOKEN
  // ═══════════════════════════════════════════════════════════════
  log('\n── AUTH: Invalid Tokens ──');

  await test('Tampered JWT token', async () => {
    const { status } = await api('GET', '/api/wos');
    // First verify current cookie works
    assert(status === 200, 'Valid cookie should work');

    const tamperedCookie = 'session=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOiJ0YW1wZXJlZCIsIm9yZ0lkIjoiZmFrZSIsInJvbGUiOiJvd25lciJ9.invalid_signature';
    const res = await fetch(`${BASE}/api/wos`, {
      headers: { 'Cookie': tamperedCookie },
    });
    assert(res.status === 401, `Expected 401 for tampered token, got ${res.status}`);
  });

  await test('Malformed cookie value', async () => {
    const res = await fetch(`${BASE}/api/wos`, {
      headers: { 'Cookie': 'session=not-a-jwt-at-all' },
    });
    assert(res.status === 401, `Expected 401, got ${res.status}`);
  });

  // Summary
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  Validation: ${passed} passed, ${failed} failed`);
  if (failures.length > 0) {
    console.log(`\n  FAILURES:`);
    for (const f of failures) console.log(`    ✗ ${f}`);
  }
  console.log(`${'═'.repeat(60)}`);

  process.exit(failed > 0 ? 1 : 0);
}

run();
