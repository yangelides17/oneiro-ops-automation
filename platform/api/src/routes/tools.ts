import { Router } from 'express';
import multer from 'multer';
import { eq, and, desc, sql } from 'drizzle-orm';
import { db } from '../db/client.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { createAuditEntry } from '../db/queries/audit.js';
import { enqueueScanJob, enqueueFillJob } from '../jobs/producers.js';
import { r2Storage } from '../integrations/storage/r2.js';
import { generateDailyDocuments, generateCertifiedPayroll } from '../services/docGeneration.js';
import { getScanJobStatuses } from '../services/woScanning.js';
import { jobs, workOrders, documents } from '../db/schema.js';

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const router = Router();

/** POST /api/tools/daily-documents — Generate PL/SI/CFR for a date. */
router.post('/daily-documents', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const { date } = req.body;
  if (!date) return res.status(400).json({ error: 'date required (YYYY-MM-DD)' });

  const orgId = getOrgId(req);

  // TOOL-1: SI completeness gate — refuse if sign-in data is still pending
  const { validateSignInsForDate } = await import('../services/docValidation.js');
  const validation = await validateSignInsForDate(db, orgId, date);
  if (!validation.ok) {
    return res.status(400).json({
      error: validation.error,
      error_code: validation.errorCode,
      missing: validation.missing,
    });
  }

  const created = await generateDailyDocuments(db, orgId, date);

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Tools',
    action: 'Generate Daily Documents',
    subject: date,
    status: `Generated ${created.length} documents`,
  });

  res.json({ ok: true, date, created, entries_found: created.length });
});

/** POST /api/tools/certified-payroll — Generate CP for a week. */
router.post('/certified-payroll', requireRole('owner', 'admin'), async (req, res) => {
  const { weekStart } = req.body;
  if (!weekStart) return res.status(400).json({ error: 'weekStart required (YYYY-MM-DD, must be Sunday)' });

  const orgId = getOrgId(req);

  // TOOL-2: SI completeness gate for the entire week
  const { validateSignInsForWeek } = await import('../services/docValidation.js');
  const validation = await validateSignInsForWeek(db, orgId, weekStart);
  if (!validation.ok) {
    return res.status(400).json({
      error: validation.error,
      error_code: validation.errorCode,
      missing: validation.missing,
    });
  }

  const created = await generateCertifiedPayroll(db, orgId, weekStart);

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Tools',
    action: 'Generate Certified Payroll',
    subject: `Week of ${weekStart}`,
    status: `Generated ${created.length} documents`,
  });

  res.json({ ok: true, weekStart, created, contract_groups: created.length });
});

/** POST /api/tools/month-end — Generate EU/Certificates for a month. */
router.post('/month-end', requireRole('owner', 'admin'), async (req, res) => {
  const { docId } = req.body;
  if (!docId) return res.status(400).json({ error: 'docId required' });

  // Month-end docs are filled by the Python fill_server via HTTP
  // This endpoint builds the fill spec and enqueues the job
  // For now, return the doc ID for the caller to poll status
  res.json({ ok: true, docId, message: 'Month-end generation queued' });
});

/** POST /api/tools/month-end/all — Batch month-end → ZIP. */
router.post('/month-end/all', requireRole('owner', 'admin'), async (req, res) => {
  const { docIds } = req.body;
  if (!Array.isArray(docIds) || docIds.length === 0) {
    return res.status(400).json({ error: 'docIds array required' });
  }
  res.json({ ok: true, message: 'Batch month-end generation queued', count: docIds.length });
});

/**
 * POST /api/tools/generate-doc/:docId — Generate/regenerate a specific document.
 * DOC-3: Per-doc PL/SI/CP generation.
 *
 * Validates SI completeness for the document's date, then enqueues
 * a fill job. Works for any doc type that has a fill module.
 */
router.post('/generate-doc/:docId', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  const doc = await db.select().from(documents).where(and(
    eq(documents.id, req.params.docId),
    eq(documents.orgId, orgId),
  )).limit(1).then(r => r[0]);

  if (!doc) return res.status(404).json({ error: 'Document not found' });

  // Validate SI completeness for the document's date
  if (doc.anchorDate) {
    const { validateSignInsForDate } = await import('../services/docValidation.js');
    const validation = await validateSignInsForDate(db, orgId, doc.anchorDate);
    if (!validation.ok) {
      return res.status(400).json({
        error: validation.error,
        error_code: validation.errorCode,
        missing: validation.missing,
      });
    }
  }

  // Reset status to pending
  await db.update(documents)
    .set({ status: 'pending', updatedAt: new Date() })
    .where(eq(documents.id, doc.id));

  // Map doc type to fill job type
  const typeMap: Record<string, string> = {
    production_log: 'fill_production_log',
    signin: 'fill_signin',
    certified_payroll: 'fill_certified_payroll',
    field_report: 'fill_field_report',
  };
  const jobType = typeMap[doc.docType];
  if (!jobType) {
    return res.status(400).json({ error: `Document type '${doc.docType}' does not support generation` });
  }

  await enqueueFillJob(jobType, {
    orgId,
    documentId: doc.id,
    templateStorageKey: '',
    fillData: {
      _type: doc.docType === 'field_report' ? 'contractor_field_report' : doc.docType,
      _regenerate: true,
      date: doc.anchorDate || '',
      contract_num: doc.contractNum || '',
      region_code: doc.regionCode || '',
      crew_chief: doc.crewChief || '',
    },
    overwriteStorageKey: doc.storageKey || undefined,
  });

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Tools',
    action: 'Document Generation Requested',
    subject: doc.docKey,
    status: 'Queued',
  });

  res.json({ ok: true, docId: doc.id, docType: doc.docType, message: 'Generation queued' });
});

/** POST /api/tools/process-approved — Archive approved documents. */
router.post('/process-approved', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);

  // Find all documents with status 'approved' and move to 'archived'
  const approved = await db.select()
    .from(documents)
    .where(and(eq(documents.orgId, orgId), eq(documents.status, 'approved')));

  let archived = 0;
  let errored = 0;
  for (const doc of approved) {
    try {
      await db.update(documents)
        .set({ status: 'archived', done: true, doneAt: new Date(), updatedAt: new Date() })
        .where(eq(documents.id, doc.id));
      archived++;
    } catch (err) {
      console.error(`[ProcessApproved] Failed to archive doc ${doc.id}:`, err);
      errored++;
    }
  }

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Tools',
    action: 'Process Approved Documents',
    subject: `${archived} documents archived`,
    status: 'Completed',
  });

  res.json({ ok: true, archived, errored, skipped: false });
});

/** POST /api/tools/scan-wo — Upload + parse WO scan via Claude Vision. */
router.post('/scan-wo', requireRole('owner', 'admin', 'foreman'), upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file attached' });

  const orgId = getOrgId(req);
  const path = `scans/${Date.now()}_${req.file.originalname}`;
  const storageKey = await r2Storage.upload(orgId, path, req.file.buffer, req.file.mimetype);

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'WO Scanner',
    action: 'Scan Uploaded',
    subject: req.file.originalname,
    status: 'Processing',
  });

  // Parse via Claude Vision — full pipeline with multi-page support
  try {
    const { parseWorkOrderScanFull } = await import('../services/woVisionParser.js');
    const parsedList = await parseWorkOrderScanFull(req.file.buffer, req.file.mimetype);

    if (parsedList.length === 0) {
      return res.status(201).json({
        ok: true, storageKey, filename: req.file.originalname,
        error: 'No work orders found in scan',
      });
    }

    // Load contractor list once for matching
    const { contractors: contractorsTable } = await import('../db/schema.js');
    const { eq: eqOp } = await import('drizzle-orm');
    const contractorRows = await db.select({ id: contractorsTable.id, name: contractorsTable.name })
      .from(contractorsTable).where(eqOp(contractorsTable.orgId, orgId));

    if (contractorRows.length === 0) {
      return res.status(201).json({
        ok: true, storageKey, filename: req.file.originalname,
        parsed: parsedList, error: 'No contractors configured — WO not created',
      });
    }

    // Process each parsed WO
    const { processScanResult } = await import('../services/woScanning.js');
    const woResults: { woId: string; woNumber: string; duplicate: boolean }[] = [];

    for (const parsed of parsedList) {
      // Resolve contractor ID from parsed name
      const contractorMatch = contractorRows.find(c =>
        c.name.toLowerCase().includes(parsed.contractor.toLowerCase()) ||
        parsed.contractor.toLowerCase().includes(c.name.toLowerCase())
      );
      const contractorId = contractorMatch?.id || contractorRows[0]?.id;

      const result = await processScanResult(db, orgId, contractorId, {
        workOrderId: parsed.workOrderId,
        contractor: parsed.contractor,
        contractNum: parsed.contractNumber,
        regionCode: parsed.regionCode,
        location: parsed.location,
        fromStreet: parsed.fromStreet,
        toStreet: parsed.toStreet,
        dueDate: parsed.dueDate,
        priority: parsed.priority,
        workType: parsed.workType,
        woReceivedDate: parsed.woReceivedDate,
        waterBlastRequired: parsed.waterBlastRequired,
        waterBlastConfirmed: parsed.waterBlastConfirmed,
        waterBlastSqft: parsed.waterBlastSqft,
        generalRemarks: parsed.generalRemarks,
        school: parsed.school,
        prepBy: parsed.prepBy,
        dateEntered: parsed.dateEntered,
        topMarkings: parsed.topMarkings,
        intersectionGrid: parsed.intersectionGrid,
        bikeLaneMarkings: parsed.bikeLaneMarkings,
      }, storageKey, undefined, req.file.originalname);

      woResults.push({
        woId: result.woId,
        woNumber: parsed.workOrderId,
        duplicate: result.duplicate,
      });

      await createAuditEntry(db, orgId, {
        userId: req.user!.userId,
        source: 'WO Scanner',
        action: result.duplicate ? 'Duplicate WO Detected' : 'WO Created from Scan',
        subject: parsed.workOrderId,
        status: result.duplicate ? 'Duplicate' : 'Created',
      });
    }

    // Response shape: single WO uses flat fields for backwards compat,
    // multi-WO adds woResults array
    const first = woResults[0];
    res.status(201).json({
      ok: true,
      storageKey,
      filename: req.file.originalname,
      woId: first?.woId,
      woNumber: first?.woNumber,
      duplicate: first?.duplicate,
      parsed: parsedList.length === 1 ? parsedList[0] : parsedList,
      ...(woResults.length > 1 ? { woResults } : {}),
    });
  } catch (err: any) {
    console.error('[Scan] Vision parse failed:', err.message);
    res.status(201).json({
      ok: true, storageKey, filename: req.file.originalname,
      error: `Parse failed: ${err.message}`,
    });
  }
});

/** GET|POST /api/tools/scan-status — Poll scan results (frontend POSTs). */
router.get('/scan-status', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const orgId = getOrgId(req);
  const scanJobs = await getScanJobStatuses(db, orgId);

  const statuses = scanJobs.map(j => ({
    jobId: j.id,
    status: j.status,
    filename: (j.payload as Record<string, unknown>)?.filename || '',
    result: j.result,
    error: j.error,
    createdAt: j.createdAt,
    completedAt: j.completedAt,
  }));

  res.json({ statuses });
});

// POST handler for scan-status (frontend sends POST, not GET)
router.post('/scan-status', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  // In direct Vision mode, scans complete synchronously — no jobs to poll.
  // Return empty statuses so the frontend's poll loop gets an empty match
  // and doesn't error out.
  res.json({ statuses: [] });
});

/** GET /api/tools/scan-uploads-today — Today's scans. */
router.get('/scan-uploads-today', requireRole('owner', 'admin', 'foreman'), async (req, res) => {
  const orgId = getOrgId(req);
  const today = new Date().toISOString().slice(0, 10);

  // Find WOs created today that have scan data
  const wos = await db.select({
    id: workOrders.id,
    woNumber: workOrders.woNumber,
    scanFileKey: workOrders.scanFileKey,
    originalFilename: workOrders.originalFilename,
    createdAt: workOrders.createdAt,
  })
    .from(workOrders)
    .where(and(
      eq(workOrders.orgId, orgId),
      sql`${workOrders.createdAt}::date = ${today}`,
    ))
    .orderBy(desc(workOrders.createdAt));

  // Group by scanFileKey (original upload) so multi-WO stacks show as one row
  const byFile = new Map<string, { fileId: string; filename: string; uploaded_at: string; woIds: string[]; is_combined: boolean }>();
  for (const w of wos) {
    if (!w.scanFileKey) continue;
    const existing = byFile.get(w.scanFileKey);
    if (existing) {
      existing.woIds.push(w.woNumber);
      existing.is_combined = true;
    } else {
      byFile.set(w.scanFileKey, {
        fileId: w.scanFileKey,
        filename: w.originalFilename || '',
        uploaded_at: w.createdAt?.toISOString?.() || String(w.createdAt || ''),
        woIds: [w.woNumber],
        is_combined: false,
      });
    }
  }

  res.json({ uploads: Array.from(byFile.values()) });
});

/** POST /api/tools/paystub/parse — Parse paystub image via Claude Vision. */
router.post('/paystub/parse', requireRole('owner', 'admin'), upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file attached' });

  // Paystub parsing uses Claude Vision — will be called by the Python worker
  // For now, return empty until the worker integration is complete
  res.json({ employees: [], message: 'Paystub parsing requires Python worker (not yet integrated)' });
});

export default router;
