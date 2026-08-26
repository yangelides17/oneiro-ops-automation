import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listContractLookup, upsertContractLookup, deleteContractLookup } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const lookupSchema = z.object({
  contractNum: z.string().min(1),
  regionCode: z.string().min(1),
  regionName: z.string().optional(),
  contractId: z.string().optional(),
  projectName: z.string().optional(),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listContractLookup(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = lookupSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await upsertContractLookup(db, getOrgId(req), { ...parsed.data, orgId: getOrgId(req) });
  res.status(201).json(row);
});

router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deleteContractLookup(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
