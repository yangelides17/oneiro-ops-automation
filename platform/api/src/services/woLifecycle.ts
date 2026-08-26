/**
 * Work Order Lifecycle Service
 *
 * Manages WO state transitions, date derivation, issue tracking, and
 * rollup computation. Extracted from the inline orchestration logic in
 * Code.js handleSubmitFieldReport_ (lines 11734–11943), handleUpdateWOStatus_
 * (lines 11013–11048), and handleDeleteWO_ (lines 11058–11117).
 *
 * Every function is pure or database-scoped — no HTTP concerns.
 */

import { eq, and, sql } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { workOrders, workDayLog, markingItems, contractLookup } from '../db/schema.js';
import { opToday } from './opDay.js';

// ─── Status State Machine ─────────────────────────────────────

/**
 * WO status values in progression order.
 * Status only moves FORWARD, never backward.
 * Exact port of Code.js (lines 11824–11831).
 */
export type WoStatus = 'received' | 'dispatched' | 'in_progress' | 'completed' | 'returned';

/** Sort priority for WO lists: active states first, completed/returned last. */
export const STATUS_SORT_PRIORITY: Record<string, number> = {
  'in_progress': 0,
  'dispatched':  1,
  'received':    2,
  'completed':   3,
  'returned':    4,
};

/**
 * Advance WO status based on the current state and the field report action.
 *
 * Rules (exact port of Code.js lines 11824–11831):
 *   - wo_complete → 'completed'
 *   - workStart now set AND status is received/dispatched → 'in_progress'
 *   - dispatch now set AND status is received → 'dispatched'
 *   - otherwise → unchanged
 *
 * Status never moves backward.
 */
export function advanceStatus(
  currentStatus: string,
  opts: {
    isComplete: boolean;
    hasDispatchDate: boolean;
    hasWorkStartDate: boolean;
  },
): WoStatus {
  const status = currentStatus.toLowerCase().replace(/\s+/g, '_') as WoStatus;

  if (opts.isComplete) return 'completed';

  if (opts.hasWorkStartDate && (status === 'received' || status === 'dispatched')) {
    return 'in_progress';
  }

  if (opts.hasDispatchDate && status === 'received') {
    return 'dispatched';
  }

  return status;
}

// ─── Date Derivation ──────────────────────────────────────────

export interface DerivedDates {
  dispatchDate: string;
  workStartDate: string;
  workEndDate: string | null;
}

/**
 * Auto-derive WO date fields from the field report's work date.
 *
 * Rules (exact port of Code.js lines 11818–11821):
 *   - dispatchDate: first FR's date fills it if blank
 *   - workStartDate: first FR's date fills it if blank
 *   - workEndDate: only set when WO is marked complete
 */
export function deriveWoDates(
  wo: {
    dispatchDate?: string | null;
    workStartDate?: string | null;
    workEndDate?: string | null;
  },
  workDate: string,
  isComplete: boolean,
): DerivedDates {
  return {
    dispatchDate: wo.dispatchDate || workDate,
    workStartDate: wo.workStartDate || workDate,
    workEndDate: isComplete ? workDate : (wo.workEndDate || null),
  };
}

// ─── Issue Tracking ───────────────────────────────────────────

/**
 * Append new issues with a date prefix, preserving existing history.
 *
 * Exact port of Code.js lines 11834–11838:
 *   newIssues = currentIssues + '\n' + date + ': ' + issues.trim()
 *
 * Only appends if newIssues is non-empty. Returns the combined string.
 */
export function appendIssues(
  existing: string | null | undefined,
  newIssues: string | null | undefined,
  date: string,
): string {
  const current = String(existing || '').trim();
  const addition = String(newIssues || '').trim();

  if (!addition) return current;

  const issueLine = `${date}: ${addition}`;
  return current ? `${current}\n${issueLine}` : issueLine;
}

// ─── Operational Day ──────────────────────────────────────────

/**
 * Apply operational-day correction to a submitted date.
 *
 * Exact port of Code.js lines 11756–11762:
 *   If the submitted date equals calendar-today AND the operational day
 *   differs (e.g. crew submitting at 3 AM), correct to yesterday.
 *   If the user explicitly picked a different date, don't touch it.
 */
export function correctOperationalDay(
  submittedDate: string,
  timezone: string = 'America/New_York',
  cutoffHour: number = 5,
): string {
  // Calendar today in the org's timezone
  const now = new Date();
  const calendarToday = now.toLocaleDateString('en-CA', { timeZone: timezone }); // YYYY-MM-DD

  // Only correct if the submitted date IS calendar-today (the default).
  // If the user explicitly picked a different date, leave it alone.
  if (submittedDate !== calendarToday) return submittedDate;

  const opDay = opToday(timezone, cutoffHour);
  return opDay !== calendarToday ? opDay : submittedDate;
}

// ─── MMA Waterblast Gate ──────────────────────────────────────

/**
 * Check whether this WO requires waterblast confirmation before submit.
 * Returns an error message if blocked, null if OK.
 *
 * Exact port of Code.js lines 11789–11797.
 */
export function checkWaterblastGate(
  waterBlastRequired: string | null | undefined,
  waterBlastConfirmed: string | null | undefined,
): string | null {
  if (String(waterBlastRequired || '') === 'Yes - MMA' &&
      String(waterBlastConfirmed || '') !== 'Yes') {
    return 'Waterblasting not confirmed for this MMA work order — toggle ' +
           '"Waterblasting Confirmed" at the top of the Field Report before submitting.';
  }
  return null;
}

// ─── Contract ID Lookup ───────────────────────────────────────

/**
 * Look up the Contract ID / Reg # from the contract_lookup table.
 * Used during WO creation to auto-fill contractId.
 *
 * Port of Code.js handleWriteWO_ (lines 6681–6696).
 */
export async function lookupContractId(
  db: Db,
  orgId: string,
  contractNum: string,
  regionCode: string,
): Promise<string | null> {
  if (!contractNum || !regionCode) return null;

  // Strip /EXT suffix for lookup
  const cn = String(contractNum).split('/')[0].trim();

  const [row] = await db.select({ contractId: contractLookup.contractId })
    .from(contractLookup)
    .where(and(
      eq(contractLookup.orgId, orgId),
      eq(contractLookup.contractNum, cn),
      eq(contractLookup.regionCode, regionCode),
    ))
    .limit(1);

  return row?.contractId ?? null;
}

// ─── WO Delete Safety ─────────────────────────────────────────

/**
 * Remove WDL rows before deleting the WO.
 * The old app preserved WDL rows (payroll/time-card history) but our
 * FK is NOT NULL with RESTRICT, so we must remove them first.
 * Signin entries (the permanent payroll records) are separate and unaffected.
 */
export async function detachWdlRows(
  db: Db,
  orgId: string,
  woId: string,
): Promise<number> {
  const deleted = await db.delete(workDayLog)
    .where(and(eq(workDayLog.orgId, orgId), eq(workDayLog.woId, woId)))
    .returning({ id: workDayLog.id });
  return deleted.length;
}

// Photos are tracked via the photos table (not a flag on the WO).
// The old app had a sticky "Photos Uploaded? = Yes/No" column;
// the new architecture uses a relational photos table as the
// source of truth — query photos.count(woId) to check.
