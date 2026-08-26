import { Router } from 'express';
import { z } from 'zod';
import { db } from '../db/client.js';
import {
  listPendingDocuments, getDocument, updateDocument, setDocumentFlags,
  getDocStatusCalendar, getPendingCounts, listDocumentsForBatch,
} from '../db/queries/documents.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { createAuditEntry } from '../db/queries/audit.js';
import { r2Storage } from '../integrations/storage/r2.js';
import { enqueueFillJob } from '../jobs/producers.js';
import { overlaySignature } from '../services/pdfOverlay.js';
import multer from 'multer';
import { ZipArchive } from 'archiver';

const approveWithSigSchema = z.object({
  signatureB64: z.string().min(1).max(500_000),
  name: z.string().min(1).optional(),
  title: z.string().optional(),
});

const docFlagsSchema = z.object({
  id: z.string().uuid(),
  done: z.boolean().optional(),
  sent: z.boolean().optional(),
});

const batchListSchema = z.object({
  docTypes: z.array(z.string()).optional(),
  contractorId: z.string().uuid().optional(),
  dateStart: z.string().optional(),
  dateEnd: z.string().optional(),
  onlyUnsent: z.boolean().optional(),
});

const router = Router();

/** GET /api/documents/pending — Approval queue. */
router.get('/pending', requireRole('owner', 'admin'), async (req, res) => {
  const docs = await listPendingDocuments(db, getOrgId(req));
  // Enrich with subtitle for display in the approvals list
  const approvals = docs.map(d => ({
    ...d,
    subtitle: [d.contractNum, d.regionCode, d.anchorDate].filter(Boolean).join(' · '),
  }));
  // Count approved docs not yet archived (waiting to be processed)
  const counts = await getPendingCounts(db, getOrgId(req));
  res.json({ approvals, approved_docs_pending: counts.approved_docs_pending });
});

/** GET /api/documents/pending/counts — Badge counts for nav. */
router.get('/pending/counts', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const counts = await getPendingCounts(db, getOrgId(req));
  res.json(counts);
});

/** GET /api/documents/:id/pdf — Stream PDF from storage. */
router.get('/:id/pdf', requireRole('owner', 'admin'), async (req, res) => {
  const doc = await getDocument(db, getOrgId(req), req.params.id);
  if (!doc || !doc.storageKey) return res.status(404).json({ error: 'Document not found' });

  try {
    const data = await r2Storage.download(getOrgId(req), doc.storageKey);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Length', data.length);
    res.setHeader('Cache-Control', 'no-store');
    res.send(data);
  } catch {
    res.status(500).json({ error: 'Failed to retrieve document' });
  }
});

/** GET /api/documents/:id/meta — Document metadata. */
router.get('/:id/meta', requireRole('owner', 'admin'), async (req, res) => {
  const doc = await getDocument(db, getOrgId(req), req.params.id);
  if (!doc) return res.status(404).json({ error: 'Not found' });
  res.json(doc);
});

/** POST /api/documents/:id/approve — Approve a document. */
router.post('/:id/approve', requireRole('owner', 'admin'), async (req, res) => {
  const doc = await updateDocument(db, getOrgId(req), req.params.id, {
    status: 'approved',
  });
  if (!doc) return res.status(404).json({ error: 'Not found' });

  await createAuditEntry(db, getOrgId(req), {
    userId: req.user!.userId,
    source: 'Approvals',
    action: 'Document Approved',
    subject: doc.filename || doc.docKey,
    status: 'Approved',
  });

  res.json({ ok: true, doc });
});

/** POST /api/documents/:id/approve-with-signature — Approve + overlay principal signature. */
router.post('/:id/approve-with-signature', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = approveWithSigSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const orgId = getOrgId(req);
  const doc = await getDocument(db, orgId, req.params.id);
  if (!doc || !doc.storageKey) return res.status(404).json({ error: 'Not found' });

  try {
    // Download current PDF, overlay signature, upload modified version
    const pdfBytes = await r2Storage.download(orgId, doc.storageKey);
    const today = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    const signedPdf = await overlaySignature({
      pdfBytes,
      signatureB64: parsed.data.signatureB64,
      name: parsed.data.name || '',
      title: parsed.data.title || '',
      dateStr: today,
    });

    // Overwrite the existing file in storage
    await r2Storage.upload(orgId, doc.storageKey.replace(`${orgId}/`, ''), signedPdf, 'application/pdf');
  } catch (err) {
    console.error('Signature overlay failed:', err);
    // Continue with approval even if overlay fails — admin can reupload
  }

  const updated = await updateDocument(db, orgId, req.params.id, {
    status: 'approved',
  });

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Approvals',
    action: 'Document Approved (signed)',
    subject: doc.filename || doc.docKey,
    status: 'Approved',
  });

  res.json({ ok: true, doc: updated });
});

/** POST /api/documents/:id/skip-signoff — Approve without signature. */
router.post('/:id/skip-signoff', requireRole('owner', 'admin'), async (req, res) => {
  const doc = await updateDocument(db, getOrgId(req), req.params.id, {
    status: 'approved',
  });
  if (!doc) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true, doc });
});

/** POST /api/documents/:id/regenerate — Regenerate document from current data. */
router.post('/:id/regenerate', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  const doc = await getDocument(db, orgId, req.params.id);
  if (!doc) return res.status(404).json({ error: 'Not found' });

  // Reset status to pending so the fill job recreates it
  await updateDocument(db, orgId, req.params.id, { status: 'pending' });

  // Re-enqueue the fill job based on doc type
  const jobType = doc.docType === 'production_log' ? 'fill_production_log'
    : doc.docType === 'field_report' ? 'fill_field_report'
    : doc.docType === 'signin' ? 'fill_signin'
    : doc.docType === 'certified_payroll' ? 'fill_certified_payroll'
    : null;

  if (jobType) {
    await enqueueFillJob(jobType, {
      orgId,
      documentId: doc.id,
      templateStorageKey: '',
      fillData: { _type: doc.docType, _regenerate: true },
      overwriteStorageKey: doc.storageKey || undefined,
    });
  }

  res.json({ ok: true, message: 'Regeneration queued' });
});

/** POST /api/documents/:id/reupload — Replace PDF in place. */
const reupload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
router.post('/:id/reupload', requireRole('owner', 'admin'), reupload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file attached' });

  const orgId = getOrgId(req);
  const doc = await getDocument(db, orgId, req.params.id);
  if (!doc) return res.status(404).json({ error: 'Not found' });

  // Upload new PDF, overwriting the old storage key if it exists
  const path = doc.storageKey
    ? doc.storageKey.replace(`${orgId}/`, '')
    : `documents/${doc.docType}/${Date.now()}_${req.file.originalname}`;
  const storageKey = await r2Storage.upload(orgId, path, req.file.buffer, 'application/pdf');

  const updated = await updateDocument(db, orgId, req.params.id, {
    storageKey,
    filename: req.file.originalname,
  });

  res.json({ ok: true, doc: updated });
});

/** GET /api/documents/status — Doc status calendar for a month. */
router.get('/status', requireRole('owner', 'admin'), async (req, res) => {
  const month = (req.query.month as string) || new Date().toISOString().slice(0, 7);
  if (!/^\d{4}-\d{2}$/.test(month)) return res.status(400).json({ error: 'month must be YYYY-MM format' });
  const monthStart = `${month}-01`;
  const [y, m] = month.split('-').map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  const monthEnd = `${month}-${String(lastDay).padStart(2, '0')}`;

  const docs = await getDocStatusCalendar(db, getOrgId(req), monthStart, monthEnd);

  // DOC-2: Build structured calendar grouped by date → contractor → doc types
  const dayMap = new Map<string, Map<string, Record<string, { done: boolean; sent: boolean; id: string | null }>>>();
  for (const doc of docs) {
    const date = String(doc.anchorDate || '');
    if (!date) continue;
    if (!dayMap.has(date)) dayMap.set(date, new Map());
    const contractorMap = dayMap.get(date)!;
    const cKey = doc.contractNum || 'unknown';
    if (!contractorMap.has(cKey)) contractorMap.set(cKey, {});
    const docTypes = contractorMap.get(cKey)!;
    docTypes[doc.docType] = { done: doc.done, sent: doc.sent, id: doc.id };
  }

  const days = Array.from(dayMap.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, contractors]) => ({
      date,
      breakdown: Array.from(contractors.entries()).map(([contractNum, docTypes]) => ({
        contractNum,
        ...docTypes,
      })),
    }));

  res.json({ month, days, docs });
});

/** POST /api/documents/status/flags — Toggle done/sent flags. */
router.post('/status/flags', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = docFlagsSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const doc = await setDocumentFlags(db, getOrgId(req), parsed.data.id, {
    done: parsed.data.done,
    sent: parsed.data.sent,
  });
  if (!doc) return res.status(404).json({ error: 'Not found' });
  res.json(doc);
});

/** POST /api/documents/flags — Doc flags. Accepts single or batch format. */
router.post('/flags', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);

  // Batch format: { updates: [{ id, done?, sent? }, ...] }
  if (req.body.updates && Array.isArray(req.body.updates)) {
    const results = [];
    for (const u of req.body.updates) {
      if (u.id) {
        const doc = await setDocumentFlags(db, orgId, u.id, { done: u.done, sent: u.sent });
        if (doc) results.push(doc);
      }
    }
    return res.json({ ok: true, updated: results.length });
  }

  // Single format: { id, done?, sent? }
  const { id, done, sent } = req.body;
  if (!id) return res.status(400).json({ error: 'id required' });
  const doc = await setDocumentFlags(db, orgId, id, { done, sent });
  if (!doc) return res.status(404).json({ error: 'Not found' });
  res.json(doc);
});

/** POST /api/documents/batch/list — List documents for batch download. */
router.post('/batch/list', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = batchListSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const files = await listDocumentsForBatch(db, getOrgId(req), parsed.data);
  res.json({ files, count: files.length });
});

/** POST /api/documents/batch/download — Stream ZIP of documents. */
router.post('/batch/download', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = batchListSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: parsed.error.issues[0].message });

  const orgId = getOrgId(req);
  const files = await listDocumentsForBatch(db, orgId, parsed.data);

  if (files.length === 0) {
    return res.status(404).json({ error: 'No documents match the filters' });
  }

  const archive = new ZipArchive({ zlib: { level: 0 } }); // PDFs already compressed

  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename="documents_${Date.now()}.zip"`);
  archive.pipe(res);

  let appended = 0;
  for (const file of files) {
    if (!file.storageKey) continue;
    try {
      const data = await r2Storage.download(orgId, file.storageKey);
      archive.append(data, { name: file.filename || `doc_${appended}.pdf` });
      appended++;
    } catch (err) {
      console.error(`Failed to fetch ${file.storageKey}:`, err);
    }
  }

  await archive.finalize();
});

export default router;
