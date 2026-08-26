import { Router } from 'express';
import { db } from '../db/client.js';
import {
  getWorkOrder, getWorkOrderByNumber, createWorkOrder, updateWorkOrder,
  deleteWorkOrder, listWorkOrdersForMap, getDashboardData,
  listWorkOrdersWithContractor,
} from '../db/queries/workOrders.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { createWoSchema, woStatusSchema } from '../utils/validation.js';
import { createAuditEntry } from '../db/queries/audit.js';
import { detachWdlRows, lookupContractId } from '../services/woLifecycle.js';
import { statusToDisplay, statusToDb } from '../utils/statusFormat.js';
import { z } from 'zod';

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
  const wos = await listWorkOrdersForMap(db, getOrgId(req));
  res.json({ wos });
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

  // Detach Work Day Log rows before deleting (FK constraint safety).
  // Old app preserved WDL rows; we remove them since signin_entries
  // holds the permanent payroll records.
  const wdlRemoved = await detachWdlRows(db, orgId, woId);

  // Marking items cascade-delete via FK (schema: onDelete: 'cascade')
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
