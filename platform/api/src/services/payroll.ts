/**
 * Certified payroll computation service.
 *
 * Computes per-employee hours, rates, and gross/net pay for a payroll week.
 * Ported from Code.js generateCertifiedPayroll (line 3123) and
 * computePayPeriodGrossForEmployee_ (line 3074).
 *
 * A payroll week runs Sunday–Saturday. Rates are resolved from the
 * pay_rates table using the week's end date (Saturday) as the effective
 * date boundary.
 */
import { eq, and, between, asc } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { signinEntries, payRates, contractors } from '../db/schema.js';
import { payrollWeekEnd } from './docLifecycle.js';
import { billingRemapAsOf, type RemapRule } from './billingRemap.js';
import { normalizeEmployeeName } from './overtime.js';

export interface PayrollEmployee {
  employeeName: string;
  classification: string;
  hoursByDay: Record<string, number>;
  otByDay: Record<string, number>;
  totalSt: number;
  totalOt: number;
  rateSt: number;
  rateOt: number;
  suppSt: number;
  suppOt: number;
  grossPay: number;
  /** Total gross across ALL contracts for this employee in this week. */
  allWorkGross: number;
}

export interface PayrollGroup {
  contractNum: string;
  regionCode: string;
  contractorId: string;
  contractorName: string;
  employees: PayrollEmployee[];
}

const DAY_KEYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'] as const;

/**
 * Compute payroll for a week, grouped by BILLING (contract, region).
 * Applies billing remap to group sign-in entries correctly.
 *
 * @param weekStart - Sunday of the payroll week (YYYY-MM-DD)
 * @param remapRules - Billing remap rules from the billing_remaps table
 */
export async function computeWeekPayroll(
  db: Db,
  orgId: string,
  weekStart: string,
  remapRules: RemapRule[],
): Promise<PayrollGroup[]> {
  const weekEnd = payrollWeekEnd(weekStart);

  // Load ALL sign-in entries for the week (needed for allWorkGross)
  const allEntries = await db.select()
    .from(signinEntries)
    .where(and(
      eq(signinEntries.orgId, orgId),
      between(signinEntries.workDate, weekStart, weekEnd),
    ))
    .orderBy(asc(signinEntries.workDate), asc(signinEntries.employeeName));

  if (allEntries.length === 0) return [];

  // Load pay rates
  const rates = await db.select()
    .from(payRates)
    .where(eq(payRates.orgId, orgId))
    .orderBy(asc(payRates.classificationCode), asc(payRates.effectiveDate));

  // Load contractor names
  const contractorRows = await db.select({ id: contractors.id, name: contractors.name })
    .from(contractors).where(eq(contractors.orgId, orgId));
  const contractorNameMap = new Map(contractorRows.map(c => [c.id, c.name]));

  // Pre-compute allWorkGross per employee (across ALL contracts)
  // Port of Code.js computePayPeriodGrossForEmployee_
  const allWorkGrossByEmployee = computeAllWorkGross(allEntries, rates, weekEnd);

  // Group entries by BILLING identity (contract, region) using billing remap
  const groupMap = new Map<string, typeof allEntries>();
  for (const entry of allEntries) {
    const contractorName = contractorNameMap.get(entry.contractorId) || '';
    const billed = billingRemapAsOf(
      remapRules,
      entry.workDate,
      entry.contractNum || '',
      entry.regionCode || '',
      contractorName,
    );
    const key = `${billed.contractNum}|${billed.regionCode}|${entry.contractorId}`;
    const group = groupMap.get(key) || [];
    group.push(entry);
    groupMap.set(key, group);
  }

  const results: PayrollGroup[] = [];

  for (const [key, groupEntries] of groupMap) {
    const [contractNum, regionCode, contractorId] = key.split('|');

    // Group by (employee, classification) — same employee with two classifications
    // gets TWO rows on the CP
    const empMap = new Map<string, { classification: string; entries: typeof groupEntries }>();
    for (const e of groupEntries) {
      const empKey = `${e.employeeName}|${e.classification}`;
      const existing = empMap.get(empKey) || { classification: e.classification, entries: [] };
      existing.entries.push(e);
      empMap.set(empKey, existing);
    }

    const employees: PayrollEmployee[] = [];

    for (const [empKey, empData] of empMap) {
      const employeeName = empKey.split('|')[0];
      const classification = empData.classification;

      const rate = resolvePayRate(rates, classification, weekEnd);

      // Compute hours by day of week
      const hoursByDay: Record<string, number> = { sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 };
      const otByDay: Record<string, number> = { sun: 0, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0 };

      for (const e of empData.entries) {
        const dayIdx = getDayIndex(e.workDate, weekStart);
        if (dayIdx < 0 || dayIdx > 6) continue;
        const dayKey = DAY_KEYS[dayIdx];
        hoursByDay[dayKey] += Number(e.hoursWorked) || 0;
        otByDay[dayKey] += Number(e.otHours) || 0;
      }

      const totalHours = Object.values(hoursByDay).reduce((a, b) => a + b, 0);
      const totalOt = Object.values(otByDay).reduce((a, b) => a + b, 0);
      const totalSt = Math.max(0, totalHours - totalOt);

      const rateSt = rate?.rateSt ?? 0;
      const rateOt = rate?.rateOt ?? 0;
      const suppSt = rate?.suppSt ?? 0;
      const suppOt = rate?.suppOt ?? 0;
      const grossPay = round2(
        totalSt * (rateSt + suppSt) + totalOt * (rateOt + suppOt),
      );

      // Look up allWorkGross for this employee
      const normName = normalizeEmployeeName(employeeName);
      const allWorkGross = allWorkGrossByEmployee.get(normName) ?? grossPay;

      employees.push({
        employeeName,
        classification,
        hoursByDay,
        otByDay,
        totalSt: round2(totalSt),
        totalOt: round2(totalOt),
        rateSt,
        rateOt,
        suppSt,
        suppOt,
        grossPay,
        allWorkGross,
      });
    }

    // Sort alphabetically by name then classification
    employees.sort((a, b) => {
      const nameCmp = a.employeeName.localeCompare(b.employeeName);
      return nameCmp !== 0 ? nameCmp : a.classification.localeCompare(b.classification);
    });

    results.push({
      contractNum,
      regionCode,
      contractorId,
      contractorName: contractorNameMap.get(contractorId) || '',
      employees,
    });
  }

  return results;
}

/**
 * Compute total gross pay across ALL contracts for each employee in the week.
 * Port of Code.js computePayPeriodGrossForEmployee_ (line 3074).
 *
 * Groups by (employee, classification), resolves rates, sums gross.
 * Returns a map of normalizedEmployeeName → total gross.
 */
function computeAllWorkGross(
  allEntries: { employeeName: string; classification: string; hoursWorked: string | null; otHours: string | null }[],
  rates: { classificationCode: string; effectiveDate: string; rateSt: string; rateOt: string; suppSt: string; suppOt: string }[],
  weekEnd: string,
): Map<string, number> {
  // Group by (normalizedName, classification) across ALL contracts
  const byEmpClass = new Map<string, { st: number; ot: number; classification: string }>();

  for (const e of allEntries) {
    const normName = normalizeEmployeeName(e.employeeName);
    const cls = e.classification;
    const key = `${normName}|||${cls}`;
    const hours = Number(e.hoursWorked) || 0;
    const ot = Number(e.otHours) || 0;
    const st = Math.max(0, hours - ot);

    const existing = byEmpClass.get(key) || { st: 0, ot: 0, classification: cls };
    existing.st += st;
    existing.ot += ot;
    byEmpClass.set(key, existing);
  }

  // Sum gross per employee (across all classifications)
  const result = new Map<string, number>();
  for (const [key, data] of byEmpClass) {
    const normName = key.split('|||')[0];
    const rate = resolvePayRate(rates, data.classification, weekEnd);
    if (!rate) continue;

    const gross = data.st * (rate.rateSt + rate.suppSt) + data.ot * (rate.rateOt + rate.suppOt);
    const prev = result.get(normName) || 0;
    result.set(normName, round2(prev + gross));
  }

  return result;
}

/**
 * Resolve pay rate for a classification as of a given date.
 * Most recent effectiveDate <= targetDate wins.
 */
function resolvePayRate(
  allRates: { classificationCode: string; effectiveDate: string; rateSt: string; rateOt: string; suppSt: string; suppOt: string }[],
  classification: string,
  targetDate: string,
): { rateSt: number; rateOt: number; suppSt: number; suppOt: number } | null {
  const candidates = allRates
    .filter(r => r.classificationCode === classification && r.effectiveDate <= targetDate)
    .sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate));

  if (candidates.length === 0) return null;
  const best = candidates[0];
  return {
    rateSt: Number(best.rateSt),
    rateOt: Number(best.rateOt),
    suppSt: Number(best.suppSt),
    suppOt: Number(best.suppOt),
  };
}

function getDayIndex(dateStr: string, weekStart: string): number {
  const d = new Date(dateStr);
  const ws = new Date(weekStart);
  return Math.round((d.getTime() - ws.getTime()) / (24 * 60 * 60 * 1000));
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}
