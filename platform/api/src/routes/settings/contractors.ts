import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listContractors, createContractor, updateContractor, deleteContractor } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const contractorSchema = z.object({
  name: z.string().min(1),
  contactName: z.string().optional(),
  contactEmail: z.string().email().optional().or(z.literal('')),
  contactPhone: z.string().optional(),
  address: z.string().optional(),
  autoGeneratePl: z.boolean().optional(),
  receivesPl: z.boolean().optional(),
  receivesCfr: z.boolean().optional(),
  receivesInvoice: z.boolean().optional(),
  receivesCp: z.boolean().optional(),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listContractors(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = contractorSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  // M-25: Default document delivery flags to true so new contractors
  // automatically receive generated documents.
  const data = {
    autoGeneratePl: true,
    receivesPl: true,
    receivesCfr: true,
    receivesInvoice: true,
    receivesCp: true,
    ...parsed.data,
    orgId: getOrgId(req),
  };
  const row = await createContractor(db, getOrgId(req), data);
  res.status(201).json(row);
});

router.patch('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = contractorSchema.partial().safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await updateContractor(db, getOrgId(req), req.params.id, parsed.data);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  try {
    const deleted = await deleteContractor(db, getOrgId(req), req.params.id);
    if (!deleted) return res.status(404).json({ error: 'Not found' });
    res.json({ ok: true });
  } catch (err: any) {
    // M-24: FK constraint violation when WOs reference this contractor
    if (err.code === '23503') {
      return res.status(409).json({
        error: 'Cannot delete contractor — it is referenced by existing work orders. Remove or reassign the work orders first.',
      });
    }
    throw err;
  }
});

export default router;
