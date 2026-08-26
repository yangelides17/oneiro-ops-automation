import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import { listEmployees, createEmployee, updateEmployee } from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const employeeSchema = z.object({
  name: z.string().min(1),
  address: z.string().optional(),
  ssnLast4: z.string().max(4).optional(),
  raceEthnicity: z.string().optional(),
  gender: z.string().optional(),
  isActive: z.boolean().optional(),
});

const router = Router();

router.get('/', async (req, res) => {
  res.json(await listEmployees(db, getOrgId(req)));
});

router.post('/', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = employeeSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createEmployee(db, getOrgId(req), { ...parsed.data, orgId: getOrgId(req) });
  res.status(201).json(row);
});

router.patch('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = employeeSchema.partial().safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await updateEmployee(db, getOrgId(req), req.params.id, parsed.data);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

// M-26: Soft-delete via isActive=false (preserves payroll history)
router.delete('/:id', requireRole('owner', 'admin'), async (req, res) => {
  const row = await updateEmployee(db, getOrgId(req), req.params.id, { isActive: false });
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

export default router;
