/**
 * Document Generation Validation Service
 *
 * Pre-flight checks before generating daily documents or certified payroll.
 * Port of Code.js _validateAllSIsDoneForContractorDay_ (line 9143)
 * and _validateAllCPForWeek_ (line 9280).
 *
 * The old app checked that all sign-in sheets for a date/week were marked
 * "Done" in the doc lifecycle log. Our simplified version checks that all
 * Work Day Log entries for the date have been submitted (status != 'pending'),
 * meaning sign-in data has been entered for every crew that worked.
 */

import { eq, and, between, sql } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { workDayLog } from '../db/schema.js';

export interface ValidationResult {
  ok: boolean;
  errorCode?: string;
  error?: string;
  missing?: string[];
}

/**
 * Validate that all sign-in data has been submitted for a given date.
 * Checks that no WDL entries are still in 'pending' status.
 *
 * Port of Code.js _validateAllSIsDoneForContractorDay_ (simplified).
 */
export async function validateSignInsForDate(
  db: Db,
  orgId: string,
  date: string,
): Promise<ValidationResult> {
  const pendingEntries = await db.select({
    contractNum: workDayLog.contractNum,
    regionCode: workDayLog.regionCode,
    crewChief: workDayLog.crewChief,
  })
    .from(workDayLog)
    .where(and(
      eq(workDayLog.orgId, orgId),
      eq(workDayLog.workDate, date),
      eq(workDayLog.status, 'pending'),
    ));

  if (pendingEntries.length === 0) {
    return { ok: true };
  }

  // Build a human-readable list of what's missing
  const seen = new Set<string>();
  const missing: string[] = [];
  for (const entry of pendingEntries) {
    const key = `${entry.contractNum}|${entry.regionCode}|${entry.crewChief || ''}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const label = [entry.contractNum, entry.regionCode, entry.crewChief]
      .filter(Boolean)
      .join(' · ');
    missing.push(label);
  }

  return {
    ok: false,
    errorCode: 'SI_NOT_DONE',
    error: `${missing.length} sign-in group${missing.length === 1 ? '' : 's'} still pending for ${date}. Submit sign-in data before generating documents.`,
    missing,
  };
}

/**
 * Validate that all sign-in data has been submitted for a payroll week.
 * Checks every day from weekStart (Sunday) through weekStart + 6 (Saturday).
 *
 * Port of Code.js _validateAllCPForWeek_ (simplified).
 */
export async function validateSignInsForWeek(
  db: Db,
  orgId: string,
  weekStart: string,
): Promise<ValidationResult> {
  // Compute week end (Saturday = weekStart + 6 days)
  const start = new Date(weekStart);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  const weekEnd = end.toISOString().slice(0, 10);

  const pendingEntries = await db.select({
    workDate: workDayLog.workDate,
    contractNum: workDayLog.contractNum,
    regionCode: workDayLog.regionCode,
    crewChief: workDayLog.crewChief,
  })
    .from(workDayLog)
    .where(and(
      eq(workDayLog.orgId, orgId),
      between(workDayLog.workDate, weekStart, weekEnd),
      eq(workDayLog.status, 'pending'),
    ));

  if (pendingEntries.length === 0) {
    return { ok: true };
  }

  const seen = new Set<string>();
  const missing: string[] = [];
  for (const entry of pendingEntries) {
    const key = `${entry.workDate}|${entry.contractNum}|${entry.regionCode}|${entry.crewChief || ''}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const label = [entry.workDate, entry.contractNum, entry.regionCode, entry.crewChief]
      .filter(Boolean)
      .join(' · ');
    missing.push(label);
  }

  return {
    ok: false,
    errorCode: 'SI_NOT_DONE',
    error: `${missing.length} sign-in group${missing.length === 1 ? '' : 's'} still pending for week of ${weekStart}. Submit all sign-in data before generating certified payroll.`,
    missing,
  };
}
