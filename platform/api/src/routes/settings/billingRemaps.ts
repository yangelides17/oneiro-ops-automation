import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listBillingRemaps, createBillingRemap, deleteBillingRemap } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const remapSchema = z.object({
  sourceContract: z.string().min(1),
  sourceRegion: z.string().min(1),
  sourceContractor: z.string().optional(),
  targetContract: z.string().min(1),
  targetRegion: z.string().min(1),
  effectiveDate: z.string(),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listBillingRemaps(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = remapSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createBillingRemap(db, getOrgId(req), { ...parsed.data, orgId: getOrgId(req) });
  res.status(201).json(row);
});

router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deleteBillingRemap(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
