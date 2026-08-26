import { Router } from 'express';
import { z } from 'zod';
import { db } from '../../db/client.js';
import {
  listClassifications, createClassification, deleteClassification,
  listPayRates, createPayRate, updatePayRate, deletePayRate,
  getOvertimeRules, updateOvertimeRules,
} from '../../db/queries/settings.js';
import { getOrgId } from '../../middleware/tenant.js';
import { requireRole } from '../../middleware/roles.js';

const classificationSchema = z.object({
  code: z.string().min(1),
  name: z.string().min(1),
  sortOrder: z.number().int().optional(),
});

const payRateSchema = z.object({
  classificationCode: z.string().min(1),
  effectiveDate: z.string(),
  rateSt: z.string(),
  rateOt: z.string(),
  suppSt: z.string().optional(),
  suppOt: z.string().optional(),
  notes: z.string().optional(),
});

const overtimeSchema = z.object({
  dailyThresholdHours: z.string().optional(),
  weeklyThresholdHours: z.string().nullable().optional(),
  weekendAllOt: z.boolean().optional(),
  crossGroupLookback: z.boolean().optional(),
});

const router = Router();

// Classifications
router.get('/classifications', async (req, res) => {
  res.json(await listClassifications(db, getOrgId(req)));
});

router.post('/classifications', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = classificationSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createClassification(db, getOrgId(req), parsed.data);
  res.status(201).json(row);
});

router.delete('/classifications/:id', requireRole('owner', 'admin'), async (req, res) => {
  try {
    const deleted = await deleteClassification(db, getOrgId(req), req.params.id);
    if (!deleted) return res.status(404).json({ error: 'Not found' });
    res.json({ ok: true });
  } catch (err: any) {
    // M-31: FK constraint violation when pay rates reference this classification
    if (err.code === '23503') {
      return res.status(409).json({
        error: 'Cannot delete classification — it has associated pay rates. Remove the rates first.',
      });
    }
    throw err;
  }
});

// Pay rates
router.get('/rates', async (req, res) => {
  res.json(await listPayRates(db, getOrgId(req)));
});

router.post('/rates', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = payRateSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await createPayRate(db, getOrgId(req), { ...parsed.data, orgId: getOrgId(req) });
  res.status(201).json(row);
});

router.patch('/rates/:id', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = payRateSchema.partial().safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });
  const row = await updatePayRate(db, getOrgId(req), req.params.id, parsed.data);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

router.delete('/rates/:id', requireRole('owner', 'admin'), async (req, res) => {
  const deleted = await deletePayRate(db, getOrgId(req), req.params.id);
  if (!deleted) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

// Overtime rules
router.get('/overtime', async (req, res) => {
  const rules = await getOvertimeRules(db, getOrgId(req));
  res.json(rules);
});

router.patch('/overtime', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = overtimeSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  // SET-1: Upsert — create the row if it doesn't exist for this org yet.
  // Fresh orgs may not have an OT rules row (signup creates one, but
  // orgs created via direct DB insert or migration might not).
  const orgId = getOrgId(req);
  let row = await updateOvertimeRules(db, orgId, parsed.data);
  if (!row) {
    const { overtimeRules } = await import('../../db/schema.js');
    const [created] = await db.insert(overtimeRules)
      .values({ orgId, ...parsed.data })
      .returning();
    row = created;
  }
  res.json(row);
});

export default router;
