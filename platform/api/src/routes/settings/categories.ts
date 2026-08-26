import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listCategories, createCategory, updateCategory, deleteCategory, listMultipliers, upsertMultiplier, deleteMultiplier } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const categorySchema = z.object({
  name: z.string().min(1),
  displayName: z.string().optional(),
  unit: z.enum(['SF', 'LF', 'EA']),
  pricingGroup: z.string().optional(),
  formSection: z.enum(['grid', 'mma', 'default']).optional(),
  requiresColor: z.boolean().optional(),
  sortOrder: z.number().int().optional(),
  isActive: z.boolean().optional(),
});

const multiplierSchema = z.object({
  categoryName: z.string().min(1),
  multiplierType: z.enum(['line_width', 'line12', 'extruded_unit', 'preformed_unit']),
  value: z.string().regex(/^\d+(\.\d+)?$/, 'Must be a numeric string'),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listCategories(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = categorySchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createCategory(db, getOrgId(req), { ...parsed.data, orgId: getOrgId(req) });
  res.status(201).json(row);
});

router.patch('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = categorySchema.partial().safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await updateCategory(db, getOrgId(req), req.params.id, parsed.data);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deleteCategory(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

router.get('/multipliers', async (req, res) => {
  res.json(await listMultipliers(db, getOrgId(req)));
});

router.put('/multipliers', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = multiplierSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await upsertMultiplier(db, getOrgId(req), parsed.data);
  res.json(row);
});

router.delete('/multipliers/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deleteMultiplier(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
