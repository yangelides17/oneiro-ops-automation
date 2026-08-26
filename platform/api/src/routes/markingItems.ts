import { Router } from 'express';
import { db } from '../db/client.js';
import {
  listMarkingItems, getMarkingItem, createMarkingItem,
  updateMarkingItem, deleteMarkingItems,
} from '../db/queries/markingItems.js';
import { getWorkOrderByNumber } from '../db/queries/workOrders.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { createMarkingItemSchema } from '../utils/validation.js';
import {
  enforceCreateDefaults, applyUpdateRules, validateGridCategory,
} from '../services/markingItemLogic.js';
import { z } from 'zod';

const updateMarkingItemSchema = z.object({
  category: z.string().optional(),
  quantity: z.number().nullable().optional(),
  unit: z.enum(['SF', 'LF', 'EA']).nullable().optional(),
  status: z.enum(['pending', 'completed', 'skipped']).optional(),
  dateCompleted: z.string().nullable().optional(),
  colorMaterial: z.string().nullable().optional(),
  crewChief: z.string().nullable().optional(),
  intersection: z.string().nullable().optional(),
  direction: z.string().nullable().optional(),
  description: z.string().nullable().optional(),
  notes: z.string().nullable().optional(),
});

const router = Router();

/** Resolve WO ID — could be UUID or WO number. */
async function resolveWoId(orgId: string, woIdOrNumber: string): Promise<string | null> {
  if (woIdOrNumber.match(/^[0-9a-f]{8}-/i)) return woIdOrNumber;
  const wo = await getWorkOrderByNumber(db, orgId, woIdOrNumber);
  return wo?.id ?? null;
}

/** GET /api/wos/:woId/markings */
router.get('/wos/:woId/markings', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const woId = await resolveWoId(orgId, req.params.woId);
  if (!woId) return res.status(404).json({ error: 'Work order not found' });
  const items = await listMarkingItems(db, orgId, woId);
  res.json({ items });
});

/** POST /api/wos/:woId/markings */
router.post('/wos/:woId/markings', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const parsed = createMarkingItemSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const orgId = getOrgId(req);
  const woId = await resolveWoId(orgId, req.params.woId);
  if (!woId) return res.status(404).json({ error: 'Work order not found' });

  // Grid category validation: intersection + direction required
  const gridErr = validateGridCategory(
    parsed.data.category,
    parsed.data.intersection,
    parsed.data.direction,
  );
  if (gridErr) return res.status(400).json({ error: gridErr });

  // Enforce server-side defaults: locked unit, status=pending, addedBy
  const { quantity, ...rest } = enforceCreateDefaults(parsed.data);

  const item = await createMarkingItem(db, orgId, woId, {
    ...rest,
    quantity: quantity != null ? String(quantity) : undefined,
  });
  res.status(201).json(item);
});

/** PATCH /api/markings/:id */
router.patch('/markings/:id', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const parsed = updateMarkingItemSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  // Look up existing item to get its current category (needed for unit enforcement)
  const orgId = getOrgId(req);
  const existing = await getMarkingItem(db, orgId, req.params.id);
  if (!existing) return res.status(404).json({ error: 'Not found' });

  // Apply business rules: category→unit derivation, qty-clearing→status revert
  const rawPatch = { ...parsed.data };
  const rules = applyUpdateRules(rawPatch, existing.category ?? undefined);

  // Convert quantity to string for DB storage
  const { quantity, ...rest } = rules as typeof parsed.data;
  const dbPatch: Record<string, unknown> = { ...rest };
  if (quantity !== undefined) {
    const q = typeof quantity === 'number' ? quantity : parseFloat(String(quantity));
    dbPatch.quantity = (!isNaN(q) && q > 0) ? String(q) : null;
  }

  const item = await updateMarkingItem(db, orgId, req.params.id, dbPatch);
  if (!item) return res.status(404).json({ error: 'Not found' });
  res.json({ item });
});

/** DELETE /api/markings */
router.delete('/markings', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const parsed = z.object({ ids: z.array(z.string().uuid()).min(1) }).safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'ids array of UUIDs required' });
  const count = await deleteMarkingItems(db, getOrgId(req), parsed.data.ids);
  res.json({ ok: true, count });
});

export default router;
