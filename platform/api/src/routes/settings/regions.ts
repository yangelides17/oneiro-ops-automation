import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listRegions, createRegion, updateRegion, deleteRegion } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const regionSchema = z.object({
  code: z.string().min(1),
  name: z.string().min(1),
  sortOrder: z.number().int().optional(),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listRegions(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = regionSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createRegion(db, getOrgId(req), parsed.data);
  res.status(201).json(row);
});

router.patch('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = regionSchema.partial().safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await updateRegion(db, getOrgId(req), req.params.id, parsed.data);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deleteRegion(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
