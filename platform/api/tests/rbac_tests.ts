/**
 * RBAC Integration Tests
 *
 * Verifies that each role (owner, admin, foreman, crew) gets correct
 * 403 responses on endpoints they shouldn't have access to.
 *
 * Uses direct DB insertion to create test users with specific roles.
 *
 * Usage: npx tsx tests/rbac_tests.ts
 */
import 'dotenv/config';
import { drizzle } from 'drizzle-orm/node-postgres';
import pg from 'pg';
import { eq, and } from 'drizzle-orm';
import bcrypt from 'bcrypt';
import jwt from 'jsonwebtoken';

const { Pool } = pg;

const DATABASE_URL = process.env.DATABASE_URL!;
const JWT_SECRET = process.env.JWT_SECRET!;
const BASE = 'http://localhost:3001';

// ─── Schema imports (inline minimal versions) ──────────────────
// We need the users table reference for direct DB ops
import {
  users, organizations, overtimeRules,
} from '../src/db/schema.js';

const pool = new Pool({ connectionString: DATABASE_URL });
const db = drizzle(pool);

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

// ─── Helpers ────────────────────────────────────────────────────

function signToken(userId: string, orgId: string, role: string): string {
  return jwt.sign({ userId, orgId, role }, JWT_SECRET, { expiresIn: '1h' });
}

function cookieFor(token: string): string {
  return `session=${token}`;
}

async function api(cookie: string, method: string, path: string, body?: any): Promise<any> {
  const opts: RequestInit = {
    method,
    headers: {
      'Content-Type': 'application/json',
      'Cookie': cookie,
    },
  };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(`${BASE}${path}`, opts);
  const text = await res.text();
  let data: any;
  try { data = JSON.parse(text); } catch { data = { _raw: text }; }
  return { status: res.status, data, ok: res.ok };
}

// ─── Test Data ──────────────────────────────────────────────────

const TEST_ORG_NAME = 'RBAC Test Org';
const TEST_PASSWORD = 'rbactest123';
let testOrgId = '';
let ownerUserId = '';
let adminUserId = '';
let foremanUserId = '';
let crewUserId = '';

let ownerCookie = '';
let adminCookie = '';
let foremanCookie = '';
let crewCookie = '';

async function setupTestUsers() {
  const passwordHash = await bcrypt.hash(TEST_PASSWORD, 12);

  // Create org
  const [org] = await db.insert(organizations).values({
    name: TEST_ORG_NAME,
  }).returning();
  testOrgId = org.id;

  // Create default OT rules
  await db.insert(overtimeRules).values({ orgId: testOrgId });

  // Create users with each role
  const roleUsers = [
    { email: 'rbac-owner@test.com', name: 'RBAC Owner', role: 'owner' },
    { email: 'rbac-admin@test.com', name: 'RBAC Admin', role: 'admin' },
    { email: 'rbac-foreman@test.com', name: 'RBAC Foreman', role: 'foreman' },
    { email: 'rbac-crew@test.com', name: 'RBAC Crew', role: 'crew' },
  ];

  for (const u of roleUsers) {
    // Delete if exists from prior run
    await db.delete(users).where(
      and(eq(users.email, u.email), eq(users.orgId, testOrgId))
    );
  }

  const [owner] = await db.insert(users).values({
    orgId: testOrgId, email: 'rbac-owner@test.com', passwordHash, name: 'RBAC Owner', role: 'owner',
  }).returning();
  ownerUserId = owner.id;
  ownerCookie = cookieFor(signToken(owner.id, testOrgId, 'owner'));

  const [admin] = await db.insert(users).values({
    orgId: testOrgId, email: 'rbac-admin@test.com', passwordHash, name: 'RBAC Admin', role: 'admin',
  }).returning();
  adminUserId = admin.id;
  adminCookie = cookieFor(signToken(admin.id, testOrgId, 'admin'));

  const [foreman] = await db.insert(users).values({
    orgId: testOrgId, email: 'rbac-foreman@test.com', passwordHash, name: 'RBAC Foreman', role: 'foreman',
  }).returning();
  foremanUserId = foreman.id;
  foremanCookie = cookieFor(signToken(foreman.id, testOrgId, 'foreman'));

  const [crew] = await db.insert(users).values({
    orgId: testOrgId, email: 'rbac-crew@test.com', passwordHash, name: 'RBAC Crew', role: 'crew',
  }).returning();
  crewUserId = crew.id;
  crewCookie = cookieFor(signToken(crew.id, testOrgId, 'crew'));
}

async function cleanupTestUsers() {
  if (testOrgId) {
    await db.delete(users).where(eq(users.orgId, testOrgId));
    await db.delete(overtimeRules).where(eq(overtimeRules.orgId, testOrgId));
    await db.delete(organizations).where(eq(organizations.id, testOrgId));
  }
}

// ─── RBAC Test Helpers ──────────────────────────────────────────

/** Verify that a role is BLOCKED (403) from an endpoint. */
async function expectForbidden(roleName: string, cookie: string, method: string, path: string, body?: any) {
  await test(`${roleName} → 403 on ${method} ${path}`, async () => {
    const { status } = await api(cookie, method, path, body);
    assert(status === 403, `Expected 403, got ${status}`);
  });
}

/** Verify that a role is ALLOWED (not 403) on an endpoint. */
async function expectAllowed(roleName: string, cookie: string, method: string, path: string, body?: any) {
  await test(`${roleName} → allowed on ${method} ${path}`, async () => {
    const { status } = await api(cookie, method, path, body);
    assert(status !== 403, `Expected non-403, got 403`);
  });
}

// ─── Main ───────────────────────────────────────────────────────

async function run() {
  log('\n═══ RBAC INTEGRATION TESTS ═══\n');

  // Setup
  log('── SETUP ──');
  await test('Create test org + users with all 4 roles', async () => {
    await setupTestUsers();
    assert(!!testOrgId, 'Should have org ID');
    assert(!!ownerCookie, 'Should have owner cookie');
    assert(!!crewCookie, 'Should have crew cookie');
  });

  // Verify each role can authenticate
  log('\n── AUTH VERIFICATION ──');
  for (const [role, cookie] of [['owner', ownerCookie], ['admin', adminCookie], ['foreman', foremanCookie], ['crew', crewCookie]] as const) {
    await test(`${role} can access /api/auth/me`, async () => {
      const { status, data } = await api(cookie, 'GET', '/api/auth/me');
      assert(status === 200, `Expected 200, got ${status}`);
      assert(data.user?.role === role, `Expected role ${role}, got ${data.user?.role}`);
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // SETTINGS — owner-only endpoints
  // ═══════════════════════════════════════════════════════════════
  log('\n── SETTINGS: Owner-Only Endpoints ──');

  // PATCH /api/settings/org — owner only
  await expectAllowed('owner', ownerCookie, 'PATCH', '/api/settings/org', { phone: '555-0000' });
  await expectForbidden('admin', adminCookie, 'PATCH', '/api/settings/org', { phone: '555-0000' });
  await expectForbidden('foreman', foremanCookie, 'PATCH', '/api/settings/org', { phone: '555-0000' });
  await expectForbidden('crew', crewCookie, 'PATCH', '/api/settings/org', { phone: '555-0000' });

  // PATCH /api/settings/users/:id/role — owner only
  const fakeUserId = '00000000-0000-0000-0000-000000000001';
  await expectForbidden('admin', adminCookie, 'PATCH', `/api/settings/users/${fakeUserId}/role`, { role: 'crew' });
  await expectForbidden('foreman', foremanCookie, 'PATCH', `/api/settings/users/${fakeUserId}/role`, { role: 'crew' });
  await expectForbidden('crew', crewCookie, 'PATCH', `/api/settings/users/${fakeUserId}/role`, { role: 'crew' });

  // ═══════════════════════════════════════════════════════════════
  // SETTINGS — owner+admin endpoints
  // ═══════════════════════════════════════════════════════════════
  log('\n── SETTINGS: Owner+Admin Endpoints ──');

  // GET /api/settings/users
  await expectAllowed('owner', ownerCookie, 'GET', '/api/settings/users');
  await expectAllowed('admin', adminCookie, 'GET', '/api/settings/users');
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/settings/users');
  await expectForbidden('crew', crewCookie, 'GET', '/api/settings/users');

  // POST /api/settings/users/invite
  await expectForbidden('foreman', foremanCookie, 'POST', '/api/settings/users/invite', { email: 'x@test.com', role: 'crew' });
  await expectForbidden('crew', crewCookie, 'POST', '/api/settings/users/invite', { email: 'x@test.com', role: 'crew' });

  // Settings CRUD — all restricted to owner+admin
  const settingsWriteEndpoints = [
    ['POST', '/api/settings/regions', { code: 'XX', name: 'Test', sortOrder: 99 }],
    ['POST', '/api/settings/contractors', { name: 'Test Co' }],
    ['POST', '/api/settings/employees', { name: 'Test Emp' }],
    ['POST', '/api/settings/pricing', { contractorId: fakeUserId, contractNum: 'X' }],
    ['POST', '/api/settings/billing-remaps', { sourceContract: 'A', targetContract: 'B' }],
    ['POST', '/api/settings/contract-lookup', { contractNum: 'X', regionCode: 'Y' }],
    ['POST', '/api/settings/categories', { name: 'X' }],
    ['POST', '/api/settings/payroll/classifications', { code: 'XX', name: 'Test' }],
    ['PATCH', '/api/settings/payroll/overtime', { dailyThresholdHours: '8' }],
  ] as const;

  for (const [method, path, body] of settingsWriteEndpoints) {
    await expectForbidden('foreman', foremanCookie, method, path, body);
    await expectForbidden('crew', crewCookie, method, path, body);
  }

  // ═══════════════════════════════════════════════════════════════
  // WORK ORDERS — create/update restricted to owner+admin+foreman
  // ═══════════════════════════════════════════════════════════════
  log('\n── WORK ORDERS: Role Restrictions ──');

  // POST /api/wos — owner, admin, foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'POST', '/api/wos', {
    woNumber: 'RBAC-TEST', contractorId: fakeUserId,
  });

  // PATCH /api/wos/:id — owner, admin, foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'PATCH', `/api/wos/${fakeUserId}`, { status: 'dispatched' });

  // DELETE /api/wos/:id — owner, admin only (NOT foreman, crew)
  await expectForbidden('foreman', foremanCookie, 'DELETE', `/api/wos/${fakeUserId}`);
  await expectForbidden('crew', crewCookie, 'DELETE', `/api/wos/${fakeUserId}`);

  // POST /api/wos/:id/waterblast/confirm — owner, admin, foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'POST', `/api/wos/${fakeUserId}/waterblast/confirm`);

  // GET /api/wos — all roles allowed (no requireRole)
  await expectAllowed('crew', crewCookie, 'GET', '/api/wos');
  await expectAllowed('foreman', foremanCookie, 'GET', '/api/wos');

  // ═══════════════════════════════════════════════════════════════
  // MARKING ITEMS — all roles can read/create/update, delete restricted
  // ═══════════════════════════════════════════════════════════════
  log('\n── MARKING ITEMS: Role Restrictions ──');

  // DELETE /api/markings — owner, admin, foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'DELETE', '/api/markings', { ids: [fakeUserId] });

  // All roles should be allowed to read/create/update markings
  await expectAllowed('crew', crewCookie, 'GET', `/api/wos/${fakeUserId}/markings`);
  await expectAllowed('foreman', foremanCookie, 'GET', `/api/wos/${fakeUserId}/markings`);

  // ═══════════════════════════════════════════════════════════════
  // FIELD REPORTS — finalize restricted to owner+admin
  // ═══════════════════════════════════════════════════════════════
  log('\n── FIELD REPORTS: Role Restrictions ──');

  // POST /api/field-reports/finalize — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'POST', '/api/field-reports/finalize', { docIds: [] });
  await expectForbidden('crew', crewCookie, 'POST', '/api/field-reports/finalize', { docIds: [] });

  // POST /api/field-reports — all roles
  // (would fail with bad data, but shouldn't be 403)
  await expectAllowed('crew', crewCookie, 'POST', '/api/field-reports', { woId: fakeUserId });
  await expectAllowed('foreman', foremanCookie, 'POST', '/api/field-reports', { woId: fakeUserId });

  // ═══════════════════════════════════════════════════════════════
  // SIGN-IN — rows/:docId and header/:docId restricted to owner+admin
  // ═══════════════════════════════════════════════════════════════
  log('\n── SIGN-IN: Role Restrictions ──');

  // GET /api/signin/rows/:docId — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', `/api/signin/rows/${fakeUserId}?workDate=2026-01-01&contractNum=X&regionCode=Y`);
  await expectForbidden('crew', crewCookie, 'GET', `/api/signin/rows/${fakeUserId}?workDate=2026-01-01&contractNum=X&regionCode=Y`);

  // GET /api/signin/header/:docId — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', `/api/signin/header/${fakeUserId}`);
  await expectForbidden('crew', crewCookie, 'GET', `/api/signin/header/${fakeUserId}`);

  // POST /api/signin/rows/:docId/edit — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'POST', `/api/signin/rows/${fakeUserId}/edit`, { edits: {} });
  await expectForbidden('crew', crewCookie, 'POST', `/api/signin/rows/${fakeUserId}/edit`, { edits: {} });

  // GET /api/signin/queue — all roles
  await expectAllowed('crew', crewCookie, 'GET', '/api/signin/queue');
  await expectAllowed('foreman', foremanCookie, 'GET', '/api/signin/queue');

  // ═══════════════════════════════════════════════════════════════
  // DOCUMENTS — most restricted to owner+admin
  // ═══════════════════════════════════════════════════════════════
  log('\n── DOCUMENTS: Role Restrictions ──');

  // GET /api/documents/pending — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/documents/pending');
  await expectForbidden('crew', crewCookie, 'GET', '/api/documents/pending');

  // GET /api/documents/pending/counts — all roles
  await expectAllowed('crew', crewCookie, 'GET', '/api/documents/pending/counts');
  await expectAllowed('foreman', foremanCookie, 'GET', '/api/documents/pending/counts');

  // POST /api/documents/:id/approve — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'POST', `/api/documents/${fakeUserId}/approve`);
  await expectForbidden('crew', crewCookie, 'POST', `/api/documents/${fakeUserId}/approve`);

  // GET /api/documents/status — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/documents/status?month=2026-08');
  await expectForbidden('crew', crewCookie, 'GET', '/api/documents/status?month=2026-08');

  // POST /api/documents/flags — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'POST', '/api/documents/flags', { id: fakeUserId, done: true });
  await expectForbidden('crew', crewCookie, 'POST', '/api/documents/flags', { id: fakeUserId, done: true });

  // ═══════════════════════════════════════════════════════════════
  // TOOLS — mixed access levels
  // ═══════════════════════════════════════════════════════════════
  log('\n── TOOLS: Role Restrictions ──');

  // POST /api/tools/daily-documents — owner+admin+foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'POST', '/api/tools/daily-documents', { date: '2026-01-01' });

  // POST /api/tools/certified-payroll — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'POST', '/api/tools/certified-payroll', { weekStart: '2026-01-01' });
  await expectForbidden('crew', crewCookie, 'POST', '/api/tools/certified-payroll', { weekStart: '2026-01-01' });

  // POST /api/tools/process-approved — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'POST', '/api/tools/process-approved');
  await expectForbidden('crew', crewCookie, 'POST', '/api/tools/process-approved');

  // GET /api/tools/scan-status — owner+admin+foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'GET', '/api/tools/scan-status');
  await expectAllowed('foreman', foremanCookie, 'GET', '/api/tools/scan-status');

  // GET /api/tools/scan-uploads-today — owner+admin+foreman (NOT crew)
  await expectForbidden('crew', crewCookie, 'GET', '/api/tools/scan-uploads-today');

  // ═══════════════════════════════════════════════════════════════
  // DASHBOARD — owner+admin only
  // ═══════════════════════════════════════════════════════════════
  log('\n── DASHBOARD: Role Restrictions ──');

  // GET /api/dashboard — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/dashboard');
  await expectForbidden('crew', crewCookie, 'GET', '/api/dashboard');

  // GET /api/revenue — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/revenue');
  await expectForbidden('crew', crewCookie, 'GET', '/api/revenue');

  // GET /api/production — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/production');
  await expectForbidden('crew', crewCookie, 'GET', '/api/production');

  // GET /api/pending-counts — all roles
  await expectAllowed('crew', crewCookie, 'GET', '/api/pending-counts');
  await expectAllowed('foreman', foremanCookie, 'GET', '/api/pending-counts');

  // GET /api/pending-counts/doc-status — owner+admin only
  await expectForbidden('foreman', foremanCookie, 'GET', '/api/pending-counts/doc-status');
  await expectForbidden('crew', crewCookie, 'GET', '/api/pending-counts/doc-status');

  // ═══════════════════════════════════════════════════════════════
  // PHOTOS — delete restricted to owner+admin+foreman
  // ═══════════════════════════════════════════════════════════════
  log('\n── PHOTOS: Role Restrictions ──');

  // DELETE /api/:id (photos) — the photo router mounts at /api, so delete is /api/:photoId
  // Note: this route pattern is very broad — Express matches it only after all other specific routes
  // The crew role would get 403, but we also need a valid-looking photo UUID that won't match other routes
  // Since this route is at /api/:id, testing it would conflict with other routes. Skip for now.
  // Instead test that the route exists and checks roles when hit directly:
  // (The photos router DELETE /:id is never hit with /api/photos/:id — it's at /api/:id)

  // ═══════════════════════════════════════════════════════════════
  // INTEGRATIONS — owner+admin only
  // ═══════════════════════════════════════════════════════════════
  log('\n── INTEGRATIONS: Role Restrictions ──');

  await expectForbidden('foreman', foremanCookie, 'GET', '/api/integrations');
  await expectForbidden('crew', crewCookie, 'GET', '/api/integrations');
  await expectForbidden('foreman', foremanCookie, 'POST', '/api/integrations/google_drive/connect');
  await expectForbidden('crew', crewCookie, 'POST', '/api/integrations/google_drive/connect');

  // ═══════════════════════════════════════════════════════════════
  // UNAUTHENTICATED — should get 401
  // ═══════════════════════════════════════════════════════════════
  log('\n── UNAUTHENTICATED ACCESS ──');

  const noAuthEndpoints = [
    ['GET', '/api/wos'],
    ['GET', '/api/dashboard'],
    ['GET', '/api/settings/users'],
    ['GET', '/api/signin/queue'],
    ['GET', '/api/documents/pending'],
    ['GET', '/api/pending-counts'],
    ['POST', '/api/wos'],
  ] as const;

  for (const [method, path] of noAuthEndpoints) {
    await test(`No auth → 401 on ${method} ${path}`, async () => {
      const { status } = await api('', method, path);
      assert(status === 401, `Expected 401, got ${status}`);
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // CLEANUP
  // ═══════════════════════════════════════════════════════════════
  log('\n── CLEANUP ──');
  await test('Remove RBAC test org + users', async () => {
    await cleanupTestUsers();
  });

  // Summary
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`  RBAC: ${passed} passed, ${failed} failed`);
  if (failures.length > 0) {
    console.log(`\n  FAILURES:`);
    for (const f of failures) console.log(`    ✗ ${f}`);
  }
  console.log(`${'═'.repeat(60)}`);

  await pool.end();
  process.exit(failed > 0 ? 1 : 0);
}

run();
