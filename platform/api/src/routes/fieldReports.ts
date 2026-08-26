import { Router } from 'express';
import { z } from 'zod';
import { eq, and } from 'drizzle-orm';
import { db } from '../db/client.js';
import { getWorkOrder } from '../db/queries/workOrders.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { documents } from '../db/schema.js';
import { buildDocKey } from '../services/docLifecycle.js';
import { submitFieldReport, FieldReportError } from '../services/fieldReportOrch.js';
import { statusToDisplay } from '../utils/statusFormat.js';
import { enqueueFillJob } from '../jobs/producers.js';
import { JOB_TYPES } from '../jobs/types.js';

const submitSchema = z.object({
  woId: z.string().uuid(),
  markComplete: z.boolean().optional(),
  wo_complete: z.enum(['yes', 'no']).optional(),
  crewChief: z.string().optional(),
  issues: z.string().optional(),
  date: z.string().optional(),
  workType: z.string().optional(),
  photos_uploaded: z.string().optional(),
});

const router = Router();

/**
 * POST /api/field-reports — Submit a field report.
 *
 * Thin dispatcher: validates input, calls the orchestration service,
 * then triggers CFR document generation as a follow-up.
 */
router.post('/', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const parsed = submitSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const orgId = getOrgId(req);
  const isComplete = parsed.data.markComplete === true || parsed.data.wo_complete === 'yes';

  try {
    // ── Core orchestration (all business logic lives here) ────
    const result = await submitFieldReport(db, orgId, {
      woId: parsed.data.woId,
      date: parsed.data.date || new Date().toISOString().split('T')[0],
      crewChief: parsed.data.crewChief || '',
      isComplete,
      issues: parsed.data.issues,
      photosUploaded: parsed.data.photos_uploaded,
      workType: parsed.data.workType,
      userId: req.user!.userId,
    });

    // ── CFR document generation (fire-and-forget follow-up) ───
    // Matches the old app's two-phase approach: submit succeeds first,
    // then CFR generation runs. If it fails, the user still sees
    // success — the doc shows up in the approvals queue when it's ready.
    try {
      const wo = await getWorkOrder(db, orgId, result.woId);
      if (wo) {
        const cfrDocKey = buildDocKey(
          'field_report', result.workDate,
          wo.contractNum || '', wo.regionCode || '',
          parsed.data.crewChief,
        );
        if (cfrDocKey) {
          const [doc] = await db.insert(documents).values({
            orgId,
            docType: 'field_report',
            docKey: cfrDocKey,
            anchorDate: result.workDate,
            contractorId: wo.contractorId,
            contractNum: wo.contractNum,
            regionCode: wo.regionCode,
            woIds: [wo.woNumber],
            crewChief: parsed.data.crewChief,
            status: 'pending',
          }).onConflictDoUpdate({
            target: [documents.orgId, documents.docKey],
            set: { status: 'pending', updatedAt: new Date() },
          }).returning();

          await enqueueFillJob(JOB_TYPES.FILL_FIELD_REPORT, {
            orgId,
            documentId: doc.id,
            templateStorageKey: '',
            fillData: {
              _type: 'contractor_field_report',
              wo_number: wo.woNumber,
              contractor: wo.contractorId,
              contract_num: wo.contractNum,
              region_code: wo.regionCode,
              location: wo.location,
              from_street: wo.fromStreet,
              to_street: wo.toStreet,
              school: wo.school,
              prep_by: wo.prepBy,
              date_entered: wo.dateEntered,
              date: result.workDate,
              install_from: wo.workStartDate || result.workDate,
              install_to: result.workDate,
              crew_chief: parsed.data.crewChief,
              issues: wo.issuesReported,
              general_remarks: wo.generalRemarks,
            },
          });
        }
      }
    } catch (docErr: any) {
      // CFR generation failure is non-fatal — log and continue
      console.error('[FieldReport] CFR doc generation failed:', docErr.message);
    }

    res.json({ ...result, status: statusToDisplay(result.status) });
  } catch (err: any) {
    if (err instanceof FieldReportError) {
      return res.status(err.statusCode).json({ error: err.message });
    }
    throw err; // let express-async-errors handle unexpected errors
  }
});

/** POST /api/field-reports/check-shift — Validate shift attribution. */
router.post('/check-shift', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const { woId, date } = req.body;
  if (!woId) return res.status(400).json({ error: 'woId required' });

  const orgId = getOrgId(req);
  const wo = await getWorkOrder(db, orgId, woId);
  if (!wo) return res.status(404).json({ error: 'Work order not found' });

  res.json({
    valid: true,
    contractor: wo.contractorId,
    contractNum: wo.contractNum,
    regionCode: wo.regionCode,
  });
});

/** POST /api/field-reports/finalize — Batch finalize field report docs. */
router.post('/finalize', requireRole('owner', 'admin'), async (req, res) => {
  const { docIds } = req.body;
  if (!Array.isArray(docIds)) return res.status(400).json({ error: 'docIds array required' });

  const orgId = getOrgId(req);
  let count = 0;

  for (const docId of docIds) {
    const [updated] = await db.update(documents)
      .set({ status: 'approved', updatedAt: new Date() })
      .where(and(
        eq(documents.id, docId),
        eq(documents.orgId, orgId),
        eq(documents.status, 'needs_review'),
      ))
      .returning();
    if (updated) count++;
  }

  res.json({ ok: true, count });
});

export default router;
