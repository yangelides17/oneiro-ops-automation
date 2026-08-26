/**
 * Document generation orchestration service.
 *
 * Orchestrates the creation of document fill jobs:
 * - Daily documents (Production Log, Sign-In, CFR)
 * - Certified Payroll (weekly)
 * - Month-end documents (EU, Certificates)
 *
 * This service queries the data needed for each document type,
 * builds the JSON payload the Python filler expects, creates
 * a document record, and enqueues a fill job.
 *
 * Replaces Code.js generateDailyDocuments (line 681),
 * generateProductionLog_ (line 1037), and generateCertifiedPayroll (line 3123).
 */
import { eq, and } from 'drizzle-orm';
import type { Db } from '../db/client.js';
import { documents, contractors, workOrders, organizations, billingRemaps } from '../db/schema.js';
import { listSigninByDate } from '../db/queries/signin.js';
import { listMarkingItems } from '../db/queries/markingItems.js';
import { buildDocKey, buildPlDocKey, buildMonthEndDocKey, payrollWeekStart, payrollWeekEnd } from './docLifecycle.js';
import { aggregateForProductionLog } from './markingAggregation.js';
import { enqueueFillJob } from '../jobs/producers.js';
import { JOB_TYPES } from '../jobs/types.js';
import { computeWeekPayroll } from './payroll.js';
import { listBillingRemaps } from '../db/queries/settings.js';
import type { RemapRule } from './billingRemap.js';

/**
 * Generate daily documents for a given date.
 * Creates Production Log and Sign-In JSON payloads for each
 * (contractor, contract, region, crew_chief) that had activity.
 *
 * Returns the list of document records created.
 */
export async function generateDailyDocuments(
  db: Db,
  orgId: string,
  targetDate: string,
): Promise<{ docType: string; docKey: string; id: string }[]> {
  const created: { docType: string; docKey: string; id: string }[] = [];

  // Get all sign-in entries for this date
  const signinRows = await listSigninByDate(db, orgId, targetDate);
  if (signinRows.length === 0) return created;

  // Group by (contractorId, contractNum, regionCode, crewChief)
  const groups = new Map<string, typeof signinRows>();
  for (const row of signinRows) {
    const key = `${row.contractorId}|${row.contractNum || ''}|${row.regionCode || ''}|${row.crewChief || ''}`;
    const existing = groups.get(key) || [];
    existing.push(row);
    groups.set(key, existing);
  }

  // For each group, check if contractor has auto_generate_pl enabled
  for (const [key, rows] of groups) {
    const [contractorId, contractNum, regionCode, crewChief] = key.split('|');

    // Get contractor config
    const [contractor] = await db.select()
      .from(contractors)
      .where(and(eq(contractors.id, contractorId), eq(contractors.orgId, orgId)))
      .limit(1);
    if (!contractor) continue;

    // Generate Production Log if enabled for this contractor
    if (contractor.autoGeneratePl) {
      const plDocKey = buildPlDocKey(targetDate, contractor.name, crewChief || null);
      if (!plDocKey) continue;

      // Get marking items for each WO in this group
      const woIds = [...new Set(rows.map(r => r.woId))];
      const plPayload: Record<string, unknown> = {
        _type: 'production_log',
        date: targetDate,
        contractor: contractor.name,
        crew_chief: crewChief || undefined,
        work_orders: [],
      };

      for (const woId of woIds) {
        const items = await listMarkingItems(db, orgId, woId);
        const aggregated = aggregateForProductionLog(
          items.map(i => ({
            category: i.category,
            quantity: i.quantity,
            unit: i.unit,
            colorMaterial: i.colorMaterial,
            dateCompleted: i.dateCompleted,
            status: i.status,
            crewChief: i.crewChief,
          })),
          targetDate,
          crewChief || null,
        );

        const [wo] = await db.select()
          .from(workOrders)
          .where(and(eq(workOrders.id, woId), eq(workOrders.orgId, orgId)))
          .limit(1);
        if (!wo) continue;

        (plPayload.work_orders as unknown[]).push({
          wo_number: wo.woNumber,
          borough: wo.regionCode,
          location: wo.location,
          markings: aggregated.markings,
          sqft: aggregated.sqft,
          paint: aggregated.paint,
        });
      }

      // Create document record
      const [doc] = await db.insert(documents).values({
        orgId,
        docType: 'production_log',
        docKey: plDocKey,
        anchorDate: targetDate,
        contractorId,
        contractNum,
        regionCode,
        woIds: woIds.map(id => id),
        crewChief: crewChief || undefined,
        status: 'pending',
      }).onConflictDoUpdate({
        target: [documents.orgId, documents.docKey],
        set: { status: 'pending', updatedAt: new Date() },
      }).returning();

      // Enqueue fill job
      await enqueueFillJob(JOB_TYPES.FILL_PRODUCTION_LOG, {
        orgId,
        documentId: doc.id,
        templateStorageKey: '', // Resolved by worker from template cache
        fillData: plPayload,
      });

      created.push({ docType: 'production_log', docKey: plDocKey, id: doc.id });
    }
  }

  return created;
}

/**
 * Generate certified payroll for a payroll week.
 * Creates one CP document per (contract, region) that had sign-in activity.
 *
 * The actual payroll computation (hours, rates, gross/net) is done by
 * the payroll service — this function orchestrates the flow.
 */
export async function generateCertifiedPayroll(
  db: Db,
  orgId: string,
  weekStartDate: string,
): Promise<{ docType: string; docKey: string; id: string }[]> {
  const created: { docType: string; docKey: string; id: string }[] = [];
  const weekStart = payrollWeekStart(weekStartDate);
  const weekEnd = payrollWeekEnd(weekStart);

  // Load org config for employer info
  const [org] = await db.select()
    .from(organizations)
    .where(eq(organizations.id, orgId))
    .limit(1);
  if (!org) return created;

  // Load billing remap rules
  const remapRows = await listBillingRemaps(db, orgId);
  const remapRules: RemapRule[] = remapRows.map(r => ({
    sourceContract: r.sourceContract,
    sourceRegion: r.sourceRegion,
    sourceContractor: r.sourceContractor,
    targetContract: r.targetContract,
    targetRegion: r.targetRegion,
    effectiveDate: r.effectiveDate,
  }));

  // Compute payroll grouped by billing (contract, region)
  const payrollGroups = await computeWeekPayroll(db, orgId, weekStart, remapRules);

  // For each group, build CP JSON and enqueue fill job
  const DAY_LABELS = ['S', 'M', 'T', 'W', 'R', 'F', 'S'];

  for (const group of payrollGroups) {
    if (group.employees.length === 0) continue;

    const cpDocKey = buildDocKey('certified_payroll', weekStart, group.contractNum, group.regionCode);
    if (!cpDocKey) continue;

    // Build day headers
    const daysJson = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(weekStart);
      d.setDate(d.getDate() + i);
      const month = d.getMonth() + 1;
      const day = String(d.getDate()).padStart(2, '0');
      daysJson.push({ label: DAY_LABELS[i], date: `${month}/${day}` });
    }

    // Build workers JSON for filler
    const workersJson = group.employees.map(emp => ({
      name: emp.employeeName,
      address: '',  // Resolved from employee registry by worker
      ssn4: '',     // Resolved from employee registry by worker
      trade: emp.classification,
      journeyperson: true,
      st_hours: ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'].map(
        d => String(emp.hoursByDay[d] - (emp.otByDay[d] || 0) || 0),
      ),
      ot_hours: ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'].map(
        d => String(emp.otByDay[d] || 0),
      ),
      total_st: String(emp.totalSt),
      total_ot: String(emp.totalOt),
      rate_st: (emp.rateSt + emp.suppSt).toFixed(2),
      rate_ot: (emp.rateOt + emp.suppOt).toFixed(2),
      supp_st: emp.suppSt.toFixed(2),
      supp_ot: emp.suppOt.toFixed(2),
      gross_pay: emp.grossPay.toFixed(2),
      total_gross_pay: emp.allWorkGross.toFixed(2),
      net_pay: '',
      deductions: '',
      annualized_rate: '',
    }));

    // Week ending date formatted for the form
    const weDate = new Date(weekEnd);
    const weekEndingFormatted = `${String(weDate.getMonth() + 1).padStart(2, '0')}/${String(weDate.getDate()).padStart(2, '0')}/${weDate.getFullYear()}`;

    const cpPayload = {
      _type: 'certified_payroll',
      payroll_number: '',
      week_ending: weekEndingFormatted,
      employer: {
        name: org.name,
        address: org.address || '',
        email: org.email || '',
        phone: org.phone || '',
        tax_id: org.taxId || '',
      },
      prime_contractor: org.signatoryName || '',
      contract_registration: group.contractNum,
      agency: '',
      agency_pin: group.contractNum,
      project_address: '',
      project_name: '',
      pla: false,
      days: daysJson,
      workers: workersJson,
    };

    // Create document record
    const [doc] = await db.insert(documents).values({
      orgId,
      docType: 'certified_payroll',
      docKey: cpDocKey,
      anchorDate: weekStart,
      contractorId: group.contractorId,
      contractNum: group.contractNum,
      regionCode: group.regionCode,
      status: 'pending',
    }).onConflictDoUpdate({
      target: [documents.orgId, documents.docKey],
      set: { status: 'pending', updatedAt: new Date() },
    }).returning();

    // Enqueue fill job
    await enqueueFillJob(JOB_TYPES.FILL_CERTIFIED_PAYROLL, {
      orgId,
      documentId: doc.id,
      templateStorageKey: '',
      fillData: cpPayload,
    });

    created.push({ docType: 'certified_payroll', docKey: cpDocKey, id: doc.id });
  }

  return created;
}
