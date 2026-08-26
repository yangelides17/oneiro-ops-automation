/**
 * Field Report Orchestration Service
 *
 * Single entry point for field report submission. Replaces the inline
 * orchestration in Code.js handleSubmitFieldReport_ (lines 11734–11946).
 *
 * Execution order (matches old app exactly):
 *   1. Validate required fields (woId, date, crewChief)
 *   2. Apply operational-day correction to the submitted date
 *   3. Load the WO from DB
 *   4. MMA waterblast gate check
 *   5. Derive dates (dispatchDate, workStartDate, workEndDate)
 *   6. Advance status via state machine
 *   7. Append issues (date-prefixed, never overwrite)
 *   8. Finalize marking items (promote pending→completed, revert qty=0)
 *   9. Compute marking rollups
 *   10. Write all WO updates in one DB call
 *   11. Create Work Day Log entry
 *   12. Create audit entry
 *   13. Return result
 *
 * CFR document generation is triggered by the CALLER (route handler)
 * after this function succeeds, matching the old app's two-phase
 * approach (submit → then finalize_field_report_docs).
 */

import { eq } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { organizations } from '../db/schema.js';
import { getWorkOrder, updateWorkOrder } from '../db/queries/workOrders.js';
import { listMarkingItems } from '../db/queries/markingItems.js';
import { createWorkDayLogEntry } from '../db/queries/signin.js';
import { createAuditEntry } from '../db/queries/audit.js';
import {
  advanceStatus, deriveWoDates, appendIssues,
  correctOperationalDay, checkWaterblastGate,
} from './woLifecycle.js';
import { finalizeMarkingStatus, computeMarkingRollups } from './markingItemLogic.js';

// ─── Input / Output Types ─────────────────────────────────────

export interface FieldReportInput {
  woId: string;
  date: string;                    // YYYY-MM-DD from the client
  crewChief: string;
  isComplete: boolean;             // wo_complete === 'yes' or markComplete === true
  issues?: string;
  photosUploaded?: string | boolean;
  workType?: string;
  userId: string;                  // for audit trail
}

export interface FieldReportResult {
  ok: true;
  woId: string;
  woNumber: string;
  status: string;
  workDate: string;                // the corrected operational day
  markingsFinalized: { promoted: number; reverted: number };
  rollups: { markingTypes: string; quantityCompleted: number | null; paintMaterial: string };
}

// ─── Orchestration ────────────────────────────────────────────

export async function submitFieldReport(
  db: Db,
  orgId: string,
  input: FieldReportInput,
): Promise<FieldReportResult> {

  // ── 1. Validate required fields ─────────────────────────────
  if (!input.woId) throw new FieldReportError('Missing woId', 400);
  if (!input.date) throw new FieldReportError('Missing date', 400);
  if (!input.crewChief?.trim()) {
    throw new FieldReportError(
      'Missing crew chief — pick the Crew Chief before submitting.', 400,
    );
  }
  const crewChief = input.crewChief.trim();

  // ── 2. Operational-day correction ───────────────────────────
  const [org] = await db.select({
    timezone: organizations.timezone,
    cutoff: organizations.opDayCutoffHour,
  })
    .from(organizations)
    .where(eq(organizations.id, orgId))
    .limit(1);

  const timezone = org?.timezone || 'America/New_York';
  const cutoff = org?.cutoff ?? 5;
  const workDate = correctOperationalDay(input.date, timezone, cutoff);

  // ── 3. Load WO ──────────────────────────────────────────────
  const wo = await getWorkOrder(db, orgId, input.woId);
  if (!wo) throw new FieldReportError('Work order not found', 404);

  // ── 4. MMA waterblast gate ──────────────────────────────────
  const gateErr = checkWaterblastGate(wo.waterBlastRequired, wo.waterBlastConfirmed);
  if (gateErr) throw new FieldReportError(gateErr, 400);

  // ── 5. Derive dates ─────────────────────────────────────────
  const dates = deriveWoDates(wo, workDate, input.isComplete);

  // ── 6. Advance status ───────────────────────────────────────
  const newStatus = advanceStatus(wo.status || 'received', {
    isComplete: input.isComplete,
    hasDispatchDate: !!dates.dispatchDate,
    hasWorkStartDate: !!dates.workStartDate,
  });

  // ── 7. Append issues ────────────────────────────────────────
  const newIssues = appendIssues(wo.issuesReported, input.issues, workDate);

  // ── 8. Finalize marking items ───────────────────────────────
  // Called on EVERY submit (not just complete) — this is how partial-day
  // work gets committed. Any item with qty > 0 → Completed with this
  // day's date. Items with qty=0 → reverted to Pending.
  const markResult = await finalizeMarkingStatus(db, orgId, wo.id, workDate, crewChief);

  // ── 9. Compute marking rollups ──────────────────────────────
  const items = await listMarkingItems(db, orgId, wo.id);
  const rollups = computeMarkingRollups(items, wo.workType || undefined);

  // ── 10. Write WO updates ────────────────────────────────────
  await updateWorkOrder(db, orgId, wo.id, {
    status: newStatus,
    dispatchDate: dates.dispatchDate,
    workStartDate: dates.workStartDate,
    workEndDate: dates.workEndDate,
    issuesReported: newIssues || undefined,
  });

  // ── 11. Create Work Day Log entry ───────────────────────────
  await createWorkDayLogEntry(db, orgId, {
    orgId,
    workDate,
    woId: wo.id,
    contractorId: wo.contractorId,
    contractNum: wo.contractNum ?? undefined,
    regionCode: wo.regionCode ?? undefined,
    location: wo.location ?? undefined,
    crewChief,
    frSubmittedAt: new Date(),
  });

  // ── 12. Audit entry ─────────────────────────────────────────
  const actionNote = input.isComplete
    ? 'WO marked COMPLETE — review for invoicing, field report, and production log'
    : '';

  await createAuditEntry(db, orgId, {
    userId: input.userId,
    source: 'Field Reports',
    action: 'Field Report Submitted',
    subject: `WO ${wo.woNumber}`,
    status: newStatus,
    details: {
      workDate,
      crewChief,
      isComplete: input.isComplete,
      markingsPromoted: markResult.promoted,
      markingsReverted: markResult.reverted,
      ...(actionNote ? { note: actionNote } : {}),
    },
  });

  // ── 13. Return result ───────────────────────────────────────
  return {
    ok: true,
    woId: wo.id,
    woNumber: wo.woNumber,
    status: newStatus,
    workDate,
    markingsFinalized: markResult,
    rollups,
  };
}

// ─── Error Type ───────────────────────────────────────────────

export class FieldReportError extends Error {
  constructor(
    message: string,
    public readonly statusCode: number,
  ) {
    super(message);
    this.name = 'FieldReportError';
  }
}
