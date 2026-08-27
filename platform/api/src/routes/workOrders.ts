import { Router } from 'express';
import { db } from '../db/client.js';
import {
  getWorkOrder, getWorkOrderByNumber, createWorkOrder, updateWorkOrder,
  deleteWorkOrder, listWorkOrdersForMap, getDashboardData,
  listWorkOrdersWithContractor, listMyWork,
} from '../db/queries/workOrders.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { createWoSchema, woStatusSchema } from '../utils/validation.js';
import { createAuditEntry } from '../db/queries/audit.js';
import { detachWdlRows, lookupContractId } from '../services/woLifecycle.js';
import { statusToDisplay, statusToDb } from '../utils/statusFormat.js';
import { z } from 'zod';
import { eq, and, inArray } from 'drizzle-orm';
import { workOrders as woTable, users } from '../db/schema.js';

/** Whitelist schema for WO PATCH — only these fields can be updated. */
const updateWoSchema = z.object({
  status: z.string().optional(),
  location: z.string().optional(),
  fromStreet: z.string().optional(),
  toStreet: z.string().optional(),
  dueDate: z.string().nullable().optional(),
  priority: z.string().optional(),
  workType: z.string().optional(),
  notes: z.string().nullable().optional(),
  generalRemarks: z.string().nullable().optional(),
  waterBlastRequired: z.string().nullable().optional(),
  waterBlastConfirmed: z.string().nullable().optional(),
  waterBlastSqft: z.string().nullable().optional(),
  dispatchDate: z.string().nullable().optional(),
  workStartDate: z.string().nullable().optional(),
  workEndDate: z.string().nullable().optional(),
  issuesReported: z.string().nullable().optional(),
  school: z.string().nullable().optional(),
  prepBy: z.string().nullable().optional(),
  latitude: z.string().nullable().optional(),
  longitude: z.string().nullable().optional(),
  contractorId: z.string().uuid().optional(),
  contractNum: z.string().optional(),
  regionCode: z.string().optional(),
  assignedTo: z.string().uuid().nullable().optional(),
}).strict(); // reject unknown fields

/** Normalize WO status fields from DB snake_case to frontend Title Case. */
function normalizeWoStatus<T extends { status?: string | null }>(wo: T): T {
  return { ...wo, status: statusToDisplay(wo.status) };
}

const router = Router();

/** GET /api/wos — List all WOs with contractor names, sorted by status priority. */
router.get('/', async (req, res) => {
  const wos = await listWorkOrdersWithContractor(db, getOrgId(req));
  res.json({ wos: wos.map(normalizeWoStatus) });
});

/** GET /api/wos/map — WOs with coordinates for map view. */
router.get('/map', async (req, res) => {
  const raw = await listWorkOrdersForMap(db, getOrgId(req));

  // Transform to the shape the frontend NavTab.jsx expects:
  // - woId = WO number (for display + field-report deep-link)
  // - id = UUID (for API calls like /api/wos/:id/files)
  // - Split into mapped (has coords) vs unmapped (needs geocoding)
  const transform = (wo: any) => ({
    ...wo,
    woId: wo.woNumber,                    // NavTab uses woId for display
    lat: wo.latitude ? Number(wo.latitude) : null,
    lng: wo.longitude ? Number(wo.longitude) : null,
    status: statusToDisplay(wo.status),
    marking_item_count: 0,                // TODO: add rollup if needed
    marking_completed_count: 0,
  });

  const all = raw.map(transform);
  const mapped = all.filter(w => w.lat && w.lng);
  const unmapped = all.filter(w => !w.lat || !w.lng);

  res.json({ mapped, unmapped });
});

/** POST /api/wos/assign — Bulk assign WOs to a user (or unassign). */
router.post('/assign', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const { woIds, userId } = req.body;
  if (!Array.isArray(woIds) || woIds.length === 0) {
    return res.status(400).json({ error: 'woIds array required' });
  }

  const orgId = getOrgId(req);

  // Validate userId belongs to this org (if assigning, not unassigning)
  if (userId) {
    const [user] = await db.select({ id: users.id })
      .from(users)
      .where(and(eq(users.id, userId), eq(users.orgId, orgId)))
      .limit(1);
    if (!user) return res.status(400).json({ error: 'User not found in this organization' });
  }

  const updated = await db.update(woTable)
    .set({
      assignedTo: userId || null,
      assignedAt: userId ? new Date() : null,
      updatedAt: new Date(),
    })
    .where(and(eq(woTable.orgId, orgId), inArray(woTable.id, woIds)))
    .returning({ id: woTable.id });

  res.json({ ok: true, count: updated.length });
});

/** GET /api/wos/my-work — WOs assigned to the current user. */
router.get('/my-work', async (req, res) => {
  const orgId = getOrgId(req);
  const userId = req.user!.userId;
  const wos = await listMyWork(db, orgId, userId);

  // Transform same as map endpoint
  const transformed = wos.map((wo: any) => ({
    ...wo,
    status: statusToDisplay(wo.status),
  }));

  res.json({ wos: transformed });
});

/** GET /api/wos/:id — Single WO. Supports lookup by UUID or WO number. */
router.get('/:id', async (req, res) => {
  const orgId = getOrgId(req);
  const param = req.params.id;

  let wo = null;
  if (param.match(/^[0-9a-f]{8}-[0-9a-f]{4}-/i)) {
    wo = await getWorkOrder(db, orgId, param);
  }
  if (!wo) {
    wo = await getWorkOrderByNumber(db, orgId, param);
  }
  if (!wo) return res.status(404).json({ error: 'Work order not found' });
  res.json(normalizeWoStatus(wo));
});

/** POST /api/wos — Create WO. */
router.post('/', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const parsed = createWoSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const orgId = getOrgId(req);

  // Auto-lookup Contract ID from contract_lookup table if not provided
  let contractId = parsed.data.contractId;
  if (!contractId && parsed.data.contractNum && parsed.data.regionCode) {
    contractId = await lookupContractId(db, orgId, parsed.data.contractNum, parsed.data.regionCode) || undefined;
  }

  const wo = await createWorkOrder(db, orgId, {
    ...parsed.data,
    contractId,
    orgId,
  });

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Work Orders',
    action: 'WO Created',
    subject: `WO ${wo.woNumber}`,
    status: 'Created',
  });

  res.status(201).json(normalizeWoStatus(wo));
});

/** PATCH /api/wos/:id — Update WO. Accepts UUID or WO number. */
router.patch('/:id', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  // Validate through whitelist schema — rejects unknown/protected fields
  const parsed = updateWoSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const patchData = { ...parsed.data };
  // Normalize status from frontend Title Case to DB snake_case
  if (patchData.status) {
    patchData.status = statusToDb(patchData.status);
    const statusParsed = woStatusSchema.safeParse(patchData.status);
    if (!statusParsed.success) return res.status(400).json({ error: 'Invalid status' });
  }
  const orgId = getOrgId(req);
  let woId = req.params.id;
  if (!woId.match(/^[0-9a-f]{8}-[0-9a-f]{4}-/i)) {
    const found = await getWorkOrderByNumber(db, orgId, woId);
    if (!found) return res.status(404).json({ error: 'Work order not found' });
    woId = found.id;
  }

  // Read previous state for audit logging and forward-only enforcement
  const before = await getWorkOrder(db, orgId, woId);
  if (!before) return res.status(404).json({ error: 'Work order not found' });

  // Noop check: if only changing status and it's the same, skip the write
  if (patchData.status && patchData.status === before.status && Object.keys(patchData).length === 1) {
    return res.json(normalizeWoStatus(before));
  }

  // Forward-only status enforcement: status can only advance, never go backward.
  // Lifecycle order: received(0) → dispatched(1) → in_progress(2) → completed(3) → returned(4)
  if (patchData.status && before.status) {
    const LIFECYCLE_ORDER: Record<string, number> = {
      received: 0, dispatched: 1, in_progress: 2, completed: 3, returned: 4,
    };
    const oldLifecycle = LIFECYCLE_ORDER[before.status] ?? 0;
    const newLifecycle = LIFECYCLE_ORDER[patchData.status] ?? 0;
    if (newLifecycle < oldLifecycle) {
      return res.status(400).json({
        error: `Cannot move status backward from ${statusToDisplay(before.status)} to ${statusToDisplay(patchData.status)}`,
      });
    }
  }

  const wo = await updateWorkOrder(db, orgId, woId, patchData);
  if (!wo) return res.status(404).json({ error: 'Work order not found' });

  // Audit log on status changes
  if (patchData.status && patchData.status !== before.status) {
    await createAuditEntry(db, orgId, {
      userId: req.user!.userId,
      source: 'Work Orders',
      action: 'Status Changed',
      subject: `WO ${wo.woNumber}`,
      status: `${before.status} → ${wo.status}`,
    });
  }

  res.json(normalizeWoStatus(wo));
});

/** DELETE /api/wos/:id — Delete WO. Accepts UUID or WO number. */
router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  let woId = req.params.id;
  if (!woId.match(/^[0-9a-f]{8}-[0-9a-f]{4}-/i)) {
    const found = await getWorkOrderByNumber(db, orgId, woId);
    if (!found) return res.status(404).json({ error: 'Work order not found' });
    woId = found.id;
  }
  const wo = await getWorkOrder(db, orgId, woId);
  if (!wo) return res.status(404).json({ error: 'Work order not found' });

  // Remove all referencing rows before deleting the WO.
  // marking_items + photos cascade automatically via FK onDelete.
  // These tables have NO ACTION FK and need manual cleanup:
  const { signinEntries, invoices, signatures, documents: docsTable, jobs: jobsTable } = await import('../db/schema.js');
  const { eq: eqOp, and: andOp, sql: sqlOp } = await import('drizzle-orm');
  await db.delete(signinEntries).where(andOp(eqOp(signinEntries.orgId, orgId), eqOp(signinEntries.woId, woId)));
  await db.delete(invoices).where(andOp(eqOp(invoices.orgId, orgId), eqOp(invoices.woId, woId)));
  await db.delete(signatures).where(eqOp(signatures.woId, woId));
  const wdlRemoved = await detachWdlRows(db, orgId, woId);

  // Clean up documents that reference this WO number (text array, not FK).
  // A document's woIds might reference multiple WOs, so only delete if
  // this WO is the only one. Otherwise remove it from the array.
  await db.delete(docsTable).where(andOp(
    eqOp(docsTable.orgId, orgId),
    sqlOp`${docsTable.woIds} @> ARRAY[${wo.woNumber}]::text[]`,
    sqlOp`array_length(${docsTable.woIds}, 1) <= 1`,
  ));
  // For multi-WO documents, remove this WO from the array
  await db.execute(sqlOp`
    UPDATE documents
    SET wo_ids = array_remove(wo_ids, ${wo.woNumber})
    WHERE org_id = ${orgId} AND wo_ids @> ARRAY[${wo.woNumber}]::text[]
  `);

  await deleteWorkOrder(db, orgId, woId);

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Work Orders',
    action: 'WO Deleted',
    subject: `WO ${wo.woNumber}`,
    status: `Deleted (${wdlRemoved} WDL rows removed)`,
  });

  res.json({ ok: true, wdlRemoved });
});

/** POST /api/wos/:id/waterblast/confirm */
router.post('/:id/waterblast/confirm', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const wo = await updateWorkOrder(db, getOrgId(req), req.params.id, { waterBlastConfirmed: 'Yes' });
  if (!wo) return res.status(404).json({ error: 'Work order not found' });
  res.json(normalizeWoStatus(wo));
});

export default router;
