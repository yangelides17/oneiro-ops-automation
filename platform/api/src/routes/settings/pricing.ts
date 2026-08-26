import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listContractPricing, createContractPricing, updateContractPricing, deleteContractPricing } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const pricingSchema = z.object({
  contractorId: z.string().min(1, 'Select a contractor before adding pricing'),
  contractNum: z.string().min(1),
  regionCode: z.string().optional(),
  effectiveDate: z.string().optional(),
  rateLine4: z.string().optional(),
  rateLine12: z.string().optional(),
  ratePreformed: z.string().optional(),
  rateExtruded: z.string().optional(),
  rateColorSurface: z.string().optional(),
  notes: z.string().optional(),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listContractPricing(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = pricingSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createContractPricing(db, getOrgId(req), { ...parsed.data, orgId: getOrgId(req) });
  res.status(201).json(row);
});

router.patch('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = pricingSchema.partial().safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await updateContractPricing(db, getOrgId(req), req.params.id, parsed.data);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deleteContractPricing(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
