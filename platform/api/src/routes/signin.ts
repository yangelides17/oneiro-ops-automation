import { Router } from 'express';
import { z } from 'zod';
import { eq, and, inArray, desc } from 'drizzle-orm';
import { db } from '../db/client.js';
import {
  listWorkDayLogQueue, createSigninEntriesBulk, getDayHoursSummary,
  getSigninRowsForDocument, updateSigninEntry, listSigninByDate,
} from '../db/queries/signin.js';
import { getDocument } from '../db/queries/documents.js';
import { getOrgId } from '../middleware/tenant.js';
import { requireRole } from '../middleware/roles.js';
import { createAuditEntry } from '../db/queries/audit.js';
import {
  allocateDayOvertime, computeRowHours, isWeekendDate, normalizeEmployeeName,
  type OvertimeRules,
} from '../services/overtime.js';
import {
  overtimeRules, contractors as contractorsTable, workOrders as woTable,
  contractLookup, signinEntries,
} from '../db/schema.js';
import { parseDocKey } from '../services/docLifecycle.js';
import { submitSignIn, getDayHoursAllGroups, SigninError } from '../services/signinOrch.js';
import { parseTimeOfDay } from '../services/overtime.js';

const router = Router();

const signinRowSchema = z.object({
  workDate: z.string(),
  woId: z.string().uuid(),
  contractorId: z.string().uuid(),
  contractNum: z.string().optional(),
  regionCode: z.string().optional(),
  location: z.string().optional(),
  employeeName: z.string().min(1),
  classification: z.string().min(1),
  timeIn: z.string().optional(),
  timeOut: z.string().optional(),
  hoursWorked: z.string().optional(),
  otHours: z.string().optional(),
  crewChief: z.string().optional(),
});

const signinEditSchema = z.record(z.string().uuid(), z.object({
  timeIn: z.string().optional(),
  timeOut: z.string().optional(),
  hoursWorked: z.string().optional(),
  otHours: z.string().optional(),
}));

/** Load org's overtime rules. */
async function loadOtRules(orgId: string): Promise<OvertimeRules> {
  const [rules] = await db.select()
    .from(overtimeRules)
    .where(eq(overtimeRules.orgId, orgId))
    .limit(1);
  return {
    dailyThresholdHours: rules?.dailyThresholdHours ? Number(rules.dailyThresholdHours) : 8,
    weeklyThresholdHours: rules?.weeklyThresholdHours ? Number(rules.weeklyThresholdHours) : null,
    weekendAllOt: rules?.weekendAllOt ?? true,
    crossGroupLookback: rules?.crossGroupLookback ?? true,
  };
}

/** GET /api/signin/queue — Pending sign-in entries, grouped by (date, contract, region, chief). */
router.get('/queue', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const orgId = getOrgId(req);
  const rawQueue = await listWorkDayLogQueue(db, orgId);

  // Load contractor details (name, contact, address)
  const contractorRows = await db.select({
    id: contractorsTable.id,
    name: contractorsTable.name,
    contactName: contractorsTable.contactName,
    address: contractorsTable.address,
  })
    .from(contractorsTable).where(eq(contractorsTable.orgId, orgId));
  const cMap = new Map(contractorRows.map(c => [c.id, c]));

  // SI-7: Load contract lookup for project_name and contractId
  // contractLookup imported at top of file
  const clRows = await db.select({
    contractNum: contractLookup.contractNum,
    regionCode: contractLookup.regionCode,
    contractId: contractLookup.contractId,
    projectName: contractLookup.projectName,
  })
    .from(contractLookup).where(eq(contractLookup.orgId, orgId));
  const clMap = new Map(clRows.map(r => [`${r.contractNum}|${r.regionCode}`, r]));

  // Load WO numbers for display
  const woIds = [...new Set(rawQueue.map(r => r.woId))];
  let woMap = new Map<string, { woNumber: string; location: string | null }>();
  if (woIds.length > 0) {
    const woRows = await db.select({ id: woTable.id, woNumber: woTable.woNumber, location: woTable.location })
      .from(woTable).where(inArray(woTable.id, woIds));
    woMap = new Map(woRows.map(w => [w.id, { woNumber: w.woNumber, location: w.location }]));
  }

  // Group by (date, contractNum, regionCode, crewChief)
  const groups = new Map<string, {
    date: string;
    contractNum: string;
    regionCode: string;
    contractorName: string;
    contractorId: string;
    primeContractor: string;
    address: string;
    contractId: string;
    projectName: string;
    crewChief: string;
    wos: { id: string; woNumber: string; location: string }[];
  }>();

  for (const row of rawQueue) {
    const key = `${row.workDate}|${row.contractNum || ''}|${row.regionCode || ''}|${row.crewChief || ''}`;
    if (!groups.has(key)) {
      const contractor = cMap.get(row.contractorId);
      const cn = String(row.contractNum || '').split('/')[0].trim();
      const cl = clMap.get(`${cn}|${row.regionCode || ''}`);
      groups.set(key, {
        date: row.workDate,
        contractNum: row.contractNum || '',
        regionCode: row.regionCode || '',
        contractorName: contractor?.name || '',
        contractorId: row.contractorId,
        primeContractor: contractor?.contactName || contractor?.name || '',
        address: contractor?.address || '',
        contractId: cl?.contractId || '',
        projectName: cl?.projectName || '',
        crewChief: row.crewChief || '',
        wos: [],
      });
    }
    const wo = woMap.get(row.woId);
    if (wo) {
      const g = groups.get(key)!;
      if (!g.wos.some(w => w.id === row.woId)) {
        g.wos.push({ id: wo.woNumber, woNumber: wo.woNumber, location: wo.location || '' });
      }
    }
  }

  const queue = Array.from(groups.entries()).map(([key, g]) => ({
    queue_id: key,
    ...g,
  }));

  res.json({ queue, total_pending: queue.length });
});

/**
 * POST /api/signin — Submit sign-in data.
 *
 * Accepts two formats:
 *   1. New format: { rows: [{ workDate, woId, contractorId, ... }] }
 *   2. Orchestrated format: { date, contractNum, regionCode, contractorId, woIds, crew, crewChief }
 *
 * Format 2 goes through the full orchestration service (billing remap,
 * WDL status update, etc). Format 1 is the legacy per-row insert path.
 */
router.post('/', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  // ── Format 2: Orchestrated submit (preferred) ───────────────
  if (req.body.crew && Array.isArray(req.body.crew)) {
    try {
      // SI-5: Decode base64 PDF if source='uploaded'
      const source = req.body.source === 'uploaded' ? 'uploaded' as const : 'generated' as const;
      const uploadBlob = source === 'uploaded' && req.body.upload_blob_b64
        ? Buffer.from(req.body.upload_blob_b64, 'base64')
        : undefined;

      const result = await submitSignIn(db, getOrgId(req), {
        date: req.body.date,
        contractNum: req.body.contractNum || req.body.contract_number,
        regionCode: req.body.regionCode || req.body.borough,
        contractorId: req.body.contractorId,
        contractorName: req.body.contractorName || req.body.contractor,
        crewChief: req.body.crewChief || req.body.crew_chief || '',
        woIds: req.body.woIds || req.body.wo_ids || [],
        crew: req.body.crew.map((m: any) => ({
          name: m.name || m.employee_name || '',
          classification: m.classification || '',
          timeIn: m.timeIn || m.time_in || '',
          timeOut: m.timeOut || m.time_out || '',
          hoursWorked: m.hoursWorked || m.hours || '',
        })),
        userId: req.user!.userId,
        submitId: req.body.submit_id || req.body.submitId,
        source,
        uploadBlob,
        uploadFilename: req.body.upload_filename || req.body.uploadFilename,
      });
      return res.status(201).json(result);
    } catch (err: any) {
      if (err instanceof SigninError) {
        return res.status(err.statusCode).json({ error: err.message });
      }
      throw err;
    }
  }

  // ── Format 1: Legacy per-row insert ─────────────────────────
  const parsed = z.object({ rows: z.array(signinRowSchema).min(1) }).safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: parsed.error.issues[0].message });
  }

  const orgId = getOrgId(req);
  const otRules = await loadOtRules(orgId);

  const rows = parsed.data.rows.map(r => {
    let hours = r.hoursWorked ? Number(r.hoursWorked) : 0;
    if (!r.hoursWorked && r.timeIn && r.timeOut) {
      hours = computeRowHours(r.timeIn, r.timeOut);
    }
    return { ...r, hours, key: normalizeEmployeeName(r.employeeName) };
  });

  const workDate = rows[0].workDate;
  const isWeekend = isWeekendDate(workDate);

  const priorHoursByEmp: Record<string, number> = {};
  if (otRules.crossGroupLookback) {
    const existingEntries = await listSigninByDate(db, orgId, workDate);
    for (const entry of existingEntries) {
      const key = normalizeEmployeeName(entry.employeeName);
      priorHoursByEmp[key] = (priorHoursByEmp[key] || 0) + (Number(entry.hoursWorked) || 0);
    }
  }

  const otValues = allocateDayOvertime(
    rows.map(r => ({ key: r.key, hours: r.hours })),
    isWeekend,
    priorHoursByEmp,
    otRules,
  );

  const finalRows = rows.map((r, i) => ({
    orgId,
    workDate: r.workDate,
    woId: r.woId,
    contractorId: r.contractorId,
    contractNum: r.contractNum,
    regionCode: r.regionCode,
    location: r.location,
    employeeName: r.employeeName,
    classification: r.classification,
    timeIn: r.timeIn,
    timeOut: r.timeOut,
    hoursWorked: String(r.hours.toFixed(2)),
    otHours: String(otValues[i].toFixed(2)),
    crewChief: r.crewChief,
  }));

  const entries = await createSigninEntriesBulk(db, orgId, finalRows);

  await createAuditEntry(db, orgId, {
    userId: req.user!.userId,
    source: 'Sign-In',
    action: 'Sign-In Submitted',
    subject: `${entries.length} employees`,
    status: 'Submitted',
  });

  res.status(201).json({ ok: true, count: entries.length });
});

/**
 * POST /api/signin/check-continuation — Check if shift continues from prior entry.
 *
 * SI-8: Port of Code.js handleCheckSignInContinuation_ (lines 5986–6069).
 * Scans existing sign-in data for any employee's time-out that falls
 * within 60 minutes before the new time-in on a DIFFERENT date.
 * If found, returns continuation=true so the UI can prompt the user.
 */
router.post('/check-continuation', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const { workDate, time_in, contractNum, regionCode, crewChief } = req.body;
  if (!workDate) return res.status(400).json({ error: 'workDate required' });

  // Also return existing count for this group (backwards compat)
  const existing = await getSigninRowsForDocument(db, getOrgId(req), {
    workDate,
    contractNum: contractNum || '',
    regionCode: regionCode || '',
    crewChief,
  });

  // If no time_in provided, skip the gap detection (just return count)
  if (!time_in) {
    return res.json({
      continuation: false,
      canContinue: true,
      nextIndex: existing.length,
      existingCount: existing.length,
    });
  }

  // Parse the target time-in into a datetime for gap comparison
  const targetTime = parseTimeOfDay(time_in);
  if (!targetTime) {
    return res.json({ continuation: false, canContinue: true, nextIndex: existing.length, existingCount: existing.length });
  }

  const [dy, dm, dd] = workDate.split('-').map(Number);
  const targetDt = new Date(dy, dm - 1, dd, targetTime.hours, targetTime.minutes);

  // Search recent sign-in entries for a time-out within 60 min of this time-in
  // Only look at entries from DIFFERENT dates (same-date doesn't need continuation prompt)
  const orgId = getOrgId(req);
  const recentEntries = await db.select({
    workDate: signinEntries.workDate,
    timeIn: signinEntries.timeIn,
    timeOut: signinEntries.timeOut,
    contractNum: signinEntries.contractNum,
    regionCode: signinEntries.regionCode,
  })
    .from(signinEntries)
    .where(eq(signinEntries.orgId, orgId))
    .orderBy(desc(signinEntries.workDate))
    .limit(200); // recent entries only

  let bestMatch: { previous_date: string; previous_contract: string; previous_time_out: string; gap_minutes: number } | null = null;

  for (const entry of recentEntries) {
    if (!entry.timeOut || entry.workDate === workDate) continue;
    const tout = parseTimeOfDay(entry.timeOut);
    if (!tout) continue;

    const [ey, em, ed] = (entry.workDate || '').split('-').map(Number);
    if (!ey) continue;

    // If time-out <= time-in (same day), shift crossed midnight
    const tin = entry.timeIn ? parseTimeOfDay(entry.timeIn) : null;
    const tinMins = tin ? tin.hours * 60 + tin.minutes : 0;
    const toutMins = tout.hours * 60 + tout.minutes;
    const cross = tin && toutMins <= tinMins;

    const toutDate = new Date(ey, em - 1, ed + (cross ? 1 : 0), tout.hours, tout.minutes);
    const gapMin = (targetDt.getTime() - toutDate.getTime()) / 60000;

    if (gapMin < 0 || gapMin > 60) continue;
    if (bestMatch && gapMin >= bestMatch.gap_minutes) continue;

    bestMatch = {
      previous_date: entry.workDate || '',
      previous_contract: [entry.contractNum, entry.regionCode].filter(Boolean).join(' / '),
      previous_time_out: entry.timeOut || '',
      gap_minutes: Math.round(gapMin),
    };
  }

  res.json({
    continuation: !!bestMatch,
    canContinue: true,
    nextIndex: existing.length,
    existingCount: existing.length,
    ...(bestMatch || {}),
  });
});

/**
 * POST /api/signin/day-hours — Hours summary for a date.
 *
 * Returns { totals: { "Employee Name": hours } } to match old app shape.
 * Scoped to ALL sign-in groups on this date (not narrowed to one contract).
 */
router.post('/day-hours', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const workDate = req.body.workDate || req.body.date;
  if (!workDate) return res.status(400).json({ error: 'workDate required' });

  const totals = await getDayHoursAllGroups(db, getOrgId(req), workDate);
  res.json({ totals });
});

/** GET /api/signin/rows/:docId — Sign-in rows for a document (used in approval editing). */
router.get('/rows/:docId', requireRole('owner', 'admin'), async (req, res) => {
  const { workDate, contractNum, regionCode, crewChief } = req.query as Record<string, string>;
  if (!workDate || !contractNum || !regionCode) {
    return res.status(400).json({ error: 'workDate, contractNum, regionCode required as query params' });
  }
  const rows = await getSigninRowsForDocument(db, getOrgId(req), {
    workDate, contractNum, regionCode, crewChief,
  });
  res.json({ rows });
});

/** GET /api/signin/header/:docId — Sign-in header info for approval display. */
router.get('/header/:docId', requireRole('owner', 'admin'), async (req, res) => {
  const orgId = getOrgId(req);
  const doc = await getDocument(db, orgId, req.params.docId);
  if (!doc) return res.status(404).json({ error: 'Document not found' });

  const parsed = parseDocKey(doc.docKey);
  const contractNum = parsed?.contractNum || doc.contractNum || '';
  const regionCode = parsed?.regionCode || doc.regionCode || '';

  // SI-10: Enrich header with contractor contacts + contract lookup
  let primeContractor = '';
  let address = '';
  let contractId = '';
  let projectName = '';
  const wos: { woNumber: string; location: string }[] = [];

  if (doc.contractorId) {
    const [contractor] = await db.select({
      name: contractorsTable.name,
      contactName: contractorsTable.contactName,
      address: contractorsTable.address,
    })
      .from(contractorsTable)
      .where(eq(contractorsTable.id, doc.contractorId))
      .limit(1);
    if (contractor) {
      primeContractor = contractor.contactName || contractor.name || '';
      address = contractor.address || '';
    }
  }

  // Contract lookup
  const cn = contractNum.split('/')[0].trim();
  if (cn && regionCode) {
    // contractLookup imported at top of file
    const [cl] = await db.select({
      contractId: contractLookup.contractId,
      projectName: contractLookup.projectName,
    })
      .from(contractLookup)
      .where(and(eq(contractLookup.orgId, orgId), eq(contractLookup.contractNum, cn), eq(contractLookup.regionCode, regionCode)))
      .limit(1);
    if (cl) {
      contractId = cl.contractId || '';
      projectName = cl.projectName || '';
    }
  }

  // WO list from doc.woIds — single query instead of N+1
  if (doc.woIds?.length) {
    const woRows = await db.select({ woNumber: woTable.woNumber, location: woTable.location })
      .from(woTable)
      .where(and(eq(woTable.orgId, orgId), inArray(woTable.woNumber, doc.woIds)));
    for (const wo of woRows) {
      wos.push({ woNumber: wo.woNumber, location: wo.location || '' });
    }
  }

  res.json({
    date: parsed?.anchorDate || doc.anchorDate || '',
    contractNum,
    regionCode,
    crewChief: parsed?.crewChief || doc.crewChief || '',
    primeContractor,
    address,
    contractId,
    projectName,
    wos,
  });
});

/** POST /api/signin/rows/:docId/edit — Edit hours in an approved sign-in. */
router.post('/rows/:docId/edit', requireRole('owner', 'admin'), async (req, res) => {
  const parsed = signinEditSchema.safeParse(req.body.edits);
  if (!parsed.success) {
    return res.status(400).json({ error: 'edits must be { entryId: { timeIn?, timeOut?, hoursWorked?, otHours? } }' });
  }

  const orgId = getOrgId(req);

  for (const [entryId, fields] of Object.entries(parsed.data)) {
    const updates: Record<string, string | undefined> = { ...fields };

    if (fields.timeIn && fields.timeOut) {
      const hours = computeRowHours(fields.timeIn, fields.timeOut);
      updates.hoursWorked = hours.toFixed(2);
    }

    await updateSigninEntry(db, orgId, entryId, updates);
  }

  // Recompute OT for the entire day
  const doc = await getDocument(db, orgId, req.params.docId);
  if (doc?.anchorDate) {
    const otRules = await loadOtRules(orgId);
    const dayEntries = await listSigninByDate(db, orgId, doc.anchorDate);
    const isWeekend = isWeekendDate(doc.anchorDate);

    const otValues = allocateDayOvertime(
      dayEntries.map(e => ({
        key: normalizeEmployeeName(e.employeeName),
        hours: Number(e.hoursWorked) || 0,
      })),
      isWeekend,
      {},
      otRules,
    );

    for (let i = 0; i < dayEntries.length; i++) {
      await updateSigninEntry(db, orgId, dayEntries[i].id, {
        otHours: otValues[i].toFixed(2),
      });
    }
  }

  res.json({ ok: true });
});

export default router;
