/**
 * Sign-In Orchestration Service
 *
 * Single entry point for sign-in submission. Replaces the inline
 * orchestration in Code.js handleSubmitSignIn_ (lines 5609–5964).
 *
 * Execution order (matches old app exactly):
 *   1. Validate required fields
 *   2. Apply billing remap to (contractNum, regionCode)
 *   3. Compute OT with cross-group lookback
 *   4. Insert sign-in entries
 *   5. Mark Work Day Log rows as "Submitted" (clears queue cards)
 *   6. Create audit entry
 *   7. Return result
 *
 * Sign-in document generation (JSON for fill worker) is a separate
 * concern handled by the document generation service.
 */

import { eq, and, asc, inArray } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { workDayLog, workOrders, contractors, documents, organizations } from '../db/schema.js';
import {
  createSigninEntriesBulk, listSigninByDate,
} from '../db/queries/signin.js';
import { createAuditEntry } from '../db/queries/audit.js';
import {
  allocateDayOvertime, computeRowHours, isWeekendDate,
  normalizeEmployeeName, to12h, type OvertimeRules,
} from './overtime.js';
import { billingRemapAsOf, billingRemap, type RemapRule } from './billingRemap.js';
import { listBillingRemaps } from '../db/queries/settings.js';
import { overtimeRules as otRulesTable } from '../db/schema.js';
import { buildDocKey } from './docLifecycle.js';
import { enqueueFillJob } from '../jobs/producers.js';
import { JOB_TYPES } from '../jobs/types.js';

// ─── Input / Output Types ─────────────────────────────────────

export interface SigninCrewMember {
  name: string;
  classification: string;
  timeIn?: string;
  timeOut?: string;
  hoursWorked?: string;
}

export interface SigninSubmitInput {
  date: string;                    // YYYY-MM-DD (already op-day corrected by FR submit)
  contractNum: string;
  regionCode: string;
  contractorId: string;
  contractorName?: string;
  crewChief: string;
  woIds: string[];                 // WO UUIDs
  crew: SigninCrewMember[];
  userId: string;
  submitId?: string;               // OTH-3: idempotency key from client
  source?: 'generated' | 'uploaded'; // SI-5: 'uploaded' = manual PDF
  uploadBlob?: Buffer;             // SI-5: raw PDF bytes when source='uploaded'
  uploadFilename?: string;         // SI-5: original filename
}

export interface SigninSubmitResult {
  ok: true;
  count: number;                   // rows inserted
  billedContractNum: string;       // after remap
  billedRegionCode: string;        // after remap
  duplicate?: boolean;             // true if this was a cached replay
}

// ─── Idempotency Cache ────────────────────────────────────────
// OTH-3: Port of Code.js idempotency logic (lines 5617–5627).
// Client sends a stable submitId per draft; we cache the response
// for 10 minutes. A network retry returns the cached result.

const IDEMPOTENCY_TTL_MS = 10 * 60 * 1000; // 10 minutes
const idempotencyCache = new Map<string, { result: SigninSubmitResult; expiresAt: number }>();

/** Periodically prune expired entries (runs on every submit, cheap). */
function pruneIdempotencyCache() {
  const now = Date.now();
  for (const [key, entry] of idempotencyCache) {
    if (entry.expiresAt < now) idempotencyCache.delete(key);
  }
}

// ─── Load OT Rules ────────────────────────────────────────────

async function loadOtRules(db: Db, orgId: string): Promise<OvertimeRules> {
  const [rules] = await db.select()
    .from(otRulesTable)
    .where(eq(otRulesTable.orgId, orgId))
    .limit(1);
  return {
    dailyThresholdHours: rules?.dailyThresholdHours ? Number(rules.dailyThresholdHours) : 8,
    weeklyThresholdHours: rules?.weeklyThresholdHours ? Number(rules.weeklyThresholdHours) : null,
    weekendAllOt: rules?.weekendAllOt ?? true,
    crossGroupLookback: rules?.crossGroupLookback ?? true,
  };
}

// ─── Orchestration ────────────────────────────────────────────

export async function submitSignIn(
  db: Db,
  orgId: string,
  input: SigninSubmitInput,
): Promise<SigninSubmitResult> {

  // ── 0. Idempotency check ─────────────────────────────────────
  pruneIdempotencyCache();
  const submitId = String(input.submitId || '').trim();
  if (submitId) {
    const cacheKey = `${orgId}:${submitId}`;
    const cached = idempotencyCache.get(cacheKey);
    if (cached && cached.expiresAt > Date.now()) {
      return { ...cached.result, duplicate: true };
    }
  }

  // ── 1. Validate ─────────────────────────────────────────────
  if (!input.date) throw new SigninError('Missing date', 400);
  if (!input.contractNum) throw new SigninError('Missing contractNum', 400);
  if (!input.regionCode) throw new SigninError('Missing regionCode', 400);
  if (!input.woIds?.length) throw new SigninError('woIds must be a non-empty array', 400);
  if (!input.crew?.length) throw new SigninError('At least one crew member is required', 400);

  const crewChief = String(input.crewChief || '').trim();

  // ── 2. Billing remap ────────────────────────────────────────
  // From the cutover date onward, the queue already sends billing tuples.
  // A stale client or pre-cutover overdue card may send raw — remap it
  // here, gated on the work date. Contractor is never remapped.
  const remapRows = await listBillingRemaps(db, orgId);
  const remapRules: RemapRule[] = remapRows.map(r => ({
    sourceContract: r.sourceContract,
    sourceRegion: r.sourceRegion,
    sourceContractor: r.sourceContractor,
    targetContract: r.targetContract,
    targetRegion: r.targetRegion,
    effectiveDate: r.effectiveDate,
  }));

  const billed = billingRemapAsOf(
    remapRules,
    input.date,
    input.contractNum,
    input.regionCode,
    input.contractorName || '',
  );

  // ── 3. Compute OT with cross-group lookback ─────────────────
  const otRules = await loadOtRules(db, orgId);
  const isWeekend = isWeekendDate(input.date);

  // Build per-crew hours
  const crewWithHours = input.crew.map(m => {
    let hours = m.hoursWorked ? Number(m.hoursWorked) : 0;
    if (!m.hoursWorked && m.timeIn && m.timeOut) {
      hours = computeRowHours(m.timeIn, m.timeOut);
    }
    return {
      ...m,
      hours,
      key: normalizeEmployeeName(m.name),
    };
  });

  // Cross-group lookback: get prior hours from existing sign-in entries
  const priorHoursByEmp: Record<string, number> = {};
  if (otRules.crossGroupLookback) {
    const existingEntries = await listSigninByDate(db, orgId, input.date);
    for (const entry of existingEntries) {
      const key = normalizeEmployeeName(entry.employeeName);
      priorHoursByEmp[key] = (priorHoursByEmp[key] || 0) + (Number(entry.hoursWorked) || 0);
    }
  }

  const otValues = allocateDayOvertime(
    crewWithHours.map(m => ({ key: m.key, hours: m.hours })),
    isWeekend,
    priorHoursByEmp,
    otRules,
  );

  // ── 4. Insert sign-in entries ───────────────────────────────
  // Resolve WO numbers for the comma-list (display field)
  const woRows = input.woIds.length > 0
    ? await db.select({ id: workOrders.id, woNumber: workOrders.woNumber, location: workOrders.location })
        .from(workOrders)
        .where(and(eq(workOrders.orgId, orgId), inArray(workOrders.id, input.woIds)))
    : [];
  const woMap = new Map(woRows.map(w => [w.id, w]));
  const woNumbers = input.woIds.map(id => woMap.get(id)?.woNumber || id);

  const finalRows = crewWithHours.map((m, i) => ({
    orgId,
    workDate: input.date,
    woId: input.woIds[0],         // Primary WO (first in list)
    contractorId: input.contractorId,
    contractNum: billed.contractNum,   // BILLING identity
    regionCode: billed.regionCode,     // BILLING identity
    location: input.woIds
      .map(id => woMap.get(id)?.location || '')
      .filter(Boolean)
      .join('; '),
    employeeName: String(m.name || '').trim(),
    classification: String(m.classification || '').trim(),
    timeIn: m.timeIn ? to12h(m.timeIn) : undefined,
    timeOut: m.timeOut ? to12h(m.timeOut) : undefined,
    hoursWorked: String(m.hours.toFixed(2)),
    otHours: String(otValues[i].toFixed(2)),
    crewChief,
  }));

  const entries = await createSigninEntriesBulk(db, orgId, finalRows);

  // ── 5. Mark WDL rows as "Submitted" ─────────────────────────
  // Single UPDATE with all filters — no N+1 loop, race-safe.
  // Filtering by crewChief is critical for multi-crew days: Crew A's
  // submit must not clear Crew B's queue card.
  const wdlResult = await db.update(workDayLog)
    .set({ status: 'submitted' })
    .where(and(
      eq(workDayLog.orgId, orgId),
      eq(workDayLog.status, 'pending'),
      inArray(workDayLog.woId, input.woIds),
      eq(workDayLog.crewChief, crewChief),
    ))
    .returning({ id: workDayLog.id });
  const wdlUpdated = wdlResult.length;

  // ── 6. Audit entry ──────────────────────────────────────────
  await createAuditEntry(db, orgId, {
    userId: input.userId,
    source: 'Sign-In',
    action: 'Sign-In Submitted',
    subject: woNumbers.join(', '),
    status: 'Submitted',
    details: {
      date: input.date,
      crewCount: entries.length,
      woCount: input.woIds.length,
      crewChief,
      billedContract: billed.contractNum,
      billedRegion: billed.regionCode,
      wdlCardsCleared: wdlUpdated,
    },
  });

  // ── 7. Generate sign-in document ─────────────────────────────
  try {
    if (input.source === 'uploaded' && input.uploadBlob) {
      // SI-5: Manual PDF upload — store the PDF directly as needs_review.
      // Port of Code.js handleSubmitSignIn_ (lines 5831–5838).
      await uploadManualSignIn(db, orgId, {
        date: input.date,
        contractNum: billed.contractNum,
        regionCode: billed.regionCode,
        contractorId: input.contractorId,
        crewChief,
        woNumbers,
        pdfBytes: input.uploadBlob,
        filename: input.uploadFilename || 'signin_manual.pdf',
      });
    } else {
      // SI-4: Build the fill spec and enqueue a fill job.
      // Port of Code.js handleSubmitSignIn_ (lines 5840–5932).
      await generateSignInDocument(db, orgId, {
        date: input.date,
        contractNum: billed.contractNum,
        regionCode: billed.regionCode,
        contractorId: input.contractorId,
        contractorName: input.contractorName || '',
        crewChief,
        woNumbers,
        woLocations: input.woIds.map(id => woMap.get(id)?.location || ''),
        crew: crewWithHours.map((m) => ({
          name: String(m.name || '').trim(),
          classification: String(m.classification || '').trim(),
          timeIn: m.timeIn ? to12h(m.timeIn) : '',
          timeOut: m.timeOut ? to12h(m.timeOut) : '',
        })),
        remapRules,
      });
    }
  } catch (err: any) {
    // Document generation failure is non-fatal
    console.error('[SignIn] Document generation failed:', err.message);
  }

  const result: SigninSubmitResult = {
    ok: true,
    count: entries.length,
    billedContractNum: billed.contractNum,
    billedRegionCode: billed.regionCode,
  };

  // Cache result for idempotency replays
  if (submitId) {
    idempotencyCache.set(`${orgId}:${submitId}`, {
      result,
      expiresAt: Date.now() + IDEMPOTENCY_TTL_MS,
    });
  }

  return result;
}

// ─── Day Hours (All Groups) ───────────────────────────────────

/**
 * Get day-hours totals across ALL sign-in groups for a date.
 * Returns the old-app response shape: { totals: { "Employee Name": hours } }
 *
 * Port of Code.js handleListSignInDayHours_ (lines 6115–6143).
 * The old app returned ALL employees' hours for the entire date
 * (not scoped to a single contract/region group).
 */
export async function getDayHoursAllGroups(
  db: Db,
  orgId: string,
  workDate: string,
): Promise<Record<string, number>> {
  const entries = await listSigninByDate(db, orgId, workDate);
  const totals: Record<string, number> = {};
  for (const entry of entries) {
    const name = String(entry.employeeName || '').trim();
    if (!name) continue;
    totals[name] = (totals[name] || 0) + (Number(entry.hoursWorked) || 0);
  }
  return totals;
}

// ─── Borough Name Map ─────────────────────────────────────────

const BOROUGH_NAMES: Record<string, string> = {
  M: 'Manhattan', BX: 'Bronx', BK: 'Brooklyn', QU: 'Queens', SI: 'Staten Island',
};

// ─── Manual Sign-In Upload ────────────────────────────────────

/**
 * Upload a manually-filled sign-in PDF directly to R2.
 * Port of Code.js handleSubmitSignIn_ (lines 5831–5838).
 *
 * The PDF goes straight to status='needs_review' — no fill job needed
 * because the crew already filled out the form by hand.
 * Filename gets a _MANUAL suffix to distinguish from machine-filled PDFs.
 */
async function uploadManualSignIn(
  db: Db,
  orgId: string,
  params: {
    date: string;
    contractNum: string;
    regionCode: string;
    contractorId: string;
    crewChief: string;
    woNumbers: string[];
    pdfBytes: Buffer;
    filename: string;
  },
): Promise<void> {
  const { r2Storage } = await import('../integrations/storage/r2.js');

  // Build doc key with _MANUAL suffix
  const siDocKey = buildDocKey('signin', params.date, params.contractNum, params.regionCode, params.crewChief);
  if (!siDocKey) return;
  const manualDocKey = `${siDocKey}_MANUAL`;

  // Upload PDF to R2
  const r2Path = `documents/signin/${manualDocKey}.pdf`;
  const storageKey = await r2Storage.upload(orgId, r2Path, params.pdfBytes, 'application/pdf');

  // Build a clean filename
  const manualFilename = `${manualDocKey}.pdf`;

  // Create document record — goes directly to needs_review (no fill needed)
  await db.insert(documents).values({
    orgId,
    docType: 'signin',
    docKey: manualDocKey,
    anchorDate: params.date,
    contractorId: params.contractorId,
    contractNum: params.contractNum,
    regionCode: params.regionCode,
    woIds: params.woNumbers,
    crewChief: params.crewChief,
    status: 'needs_review',
    storageKey,
    filename: manualFilename,
  }).onConflictDoUpdate({
    target: [documents.orgId, documents.docKey],
    set: {
      status: 'needs_review',
      storageKey,
      filename: manualFilename,
      updatedAt: new Date(),
    },
  });
}

// ─── Sign-In Document Generation ──────────────────────────────

/**
 * Generate the sign-in fill spec and enqueue a fill job.
 * Port of Code.js handleSubmitSignIn_ (lines 5840–5932).
 */
async function generateSignInDocument(
  db: Db,
  orgId: string,
  params: {
    date: string;
    contractNum: string;
    regionCode: string;
    contractorId: string;
    contractorName: string;
    crewChief: string;
    woNumbers: string[];
    woLocations: string[];
    crew: { name: string; classification: string; timeIn: string; timeOut: string }[];
    remapRules: RemapRule[];
  },
): Promise<void> {
  const siDocKey = buildDocKey('signin', params.date, params.contractNum, params.regionCode, params.crewChief);
  if (!siDocKey) return;

  // Contractor contact lookup
  const [contractor] = await db.select({
    name: contractors.name,
    contactName: contractors.contactName,
    address: contractors.address,
  })
    .from(contractors)
    .where(eq(contractors.id, params.contractorId))
    .limit(1);

  const primeContractor = contractor?.contactName || contractor?.name || params.contractorName;
  const address = contractor?.address || '';

  // Org name for subcontractor field
  const [org] = await db.select({ name: organizations.name })
    .from(organizations).where(eq(organizations.id, orgId)).limit(1);

  // PDF content uses billing identity (always remapped)
  const mapped = billingRemap(
    params.remapRules, params.contractNum, params.regionCode, params.contractorName,
  );
  const boroughName = BOROUGH_NAMES[mapped.regionCode] || mapped.regionCode;
  const contractLabel = boroughName
    ? `${mapped.contractNum} - ${boroughName}`
    : mapped.contractNum;

  // Date format for PDF: M/D/YY
  const dm = params.date.match(/^(\d{4})-(\d{2})-(\d{2})/);
  const dateFmt = dm
    ? `${parseInt(dm[2], 10)}/${parseInt(dm[3], 10)}/${dm[1].slice(-2)}`
    : params.date;

  // WO label
  const woLabel = params.woNumbers.length > 1
    ? `${params.woNumbers.join(', ')} (${params.woNumbers.length})`
    : params.woNumbers[0] || '';

  // Project name: pair each WO with its location
  const projectName = params.woNumbers
    .map((wo, i) => {
      const loc = (params.woLocations[i] || '').trim();
      return loc ? `${wo} | ${loc}` : wo;
    })
    .join('; ');

  // Build fill payload (matches fill_signin.py expected shape)
  const fillData: Record<string, unknown> = {
    _type: 'signin',
    date: dateFmt,
    prime_contractor: primeContractor,
    subcontractor: org?.name || '',
    contract_number: contractLabel,
    address,
    agency: 'DOT',
    project_name: projectName,
    crew_chief: params.crewChief,
    wo_ids: params.woNumbers,
    locations: params.woLocations,
    wo_label: woLabel,
    wo_id: params.woNumbers[0] || '',
    crew: params.crew.map(m => ({
      name: m.name,
      classification: m.classification,
      time_in: m.timeIn,
      time_out: m.timeOut,
      sig_in_b64: '',
      sig_out_b64: '',
    })),
    contractor_name: '',
    contractor_title: '',
    date_signed: '',
    contractor_signature_b64: '',
  };

  // Create document record
  const [doc] = await db.insert(documents).values({
    orgId,
    docType: 'signin',
    docKey: siDocKey,
    anchorDate: params.date,
    contractorId: params.contractorId,
    contractNum: params.contractNum,
    regionCode: params.regionCode,
    woIds: params.woNumbers,
    crewChief: params.crewChief,
    status: 'pending',
  }).onConflictDoUpdate({
    target: [documents.orgId, documents.docKey],
    set: { status: 'pending', updatedAt: new Date() },
  }).returning();

  await enqueueFillJob(JOB_TYPES.FILL_SIGNIN, {
    orgId,
    documentId: doc.id,
    templateStorageKey: '',
    fillData,
  });
}

// ─── Error Type ───────────────────────────────────────────────

export class SigninError extends Error {
  constructor(
    message: string,
    public readonly statusCode: number,
  ) {
    super(message);
    this.name = 'SigninError';
  }
}
