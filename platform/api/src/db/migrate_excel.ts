/**
 * Migration script: Import Oneiro data from Excel export into PostgreSQL.
 *
 * Usage: npx tsx src/db/migrate_excel.ts <path-to-xlsx> <org-id>
 */
import 'dotenv/config';
import { db } from './client.js';
import {
  organizations, contractors, employees, contractLookup, contractPricing,
  payRates, payClassifications, workOrders, markingItems, signinEntries,
  workDayLog, documents, payrollEntries, invoices, auditLog, overtimeRules,
} from './schema.js';
import { eq } from 'drizzle-orm';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const XLSX_PATH = process.argv[2];
const ORG_ID = process.argv[3];

if (!XLSX_PATH || !ORG_ID) {
  console.error('Usage: npx tsx src/db/migrate_excel.ts <path-to-xlsx> <org-id>');
  process.exit(1);
}

function log(msg: string) { console.log(`[migrate] ${msg}`); }

function parseDate(v: any): string | null {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  // Handle "2026-05-25 00:00:00" format
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  return null;
}

function parseTimestamp(v: any): Date | null {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function str(v: any): string { return String(v || '').trim(); }
function numOrNull(v: any): string | null {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return isNaN(n) ? null : String(n);
}
function boolFromYes(v: any): boolean { return String(v || '').toLowerCase() === 'yes'; }

async function migrate() {
  log(`Reading ${XLSX_PATH}...`);
  const workbook = XLSX.readFile(XLSX_PATH, { cellDates: true });

  function sheet(name: string): Record<string, any>[] {
    const ws = workbook.Sheets[name];
    if (!ws) { log(`  Sheet "${name}" not found, skipping`); return []; }
    return XLSX.utils.sheet_to_json(ws, { defval: null });
  }

  // ─── 1. Update Org Settings ──────────────────────────────
  log('Updating organization settings...');
  await db.update(organizations).set({
    address: '435 South Ave Apt 220, Garwood, NJ 07027',
    phone: '+1 917-620-9809',
    email: 'anthe@theoneiro.com',
    taxId: '0450359615',
    signatoryName: 'Marianthy Angelides',
    signatoryTitle: 'Manager / President',
    updatedAt: new Date(),
  }).where(eq(organizations.id, ORG_ID));

  // ─── 2. Contractors ──────────────────────────────────────
  log('Migrating contractors...');
  const ccRows = sheet('Contractor Contacts');
  const contractorMap = new Map<string, string>(); // name → uuid

  for (const r of ccRows) {
    const name = str(r['Prime Contractor']);
    if (!name) continue;
    const [row] = await db.insert(contractors).values({
      orgId: ORG_ID,
      name,
      contactName: str(r['Contact Name']),
      contactEmail: str(r['Email']),
      contactPhone: str(r['Phone']),
      address: str(r['Address']),
      autoGeneratePl: name === 'Metro Express',
      receivesPl: boolFromYes(r['Receives Production Log?']),
      receivesCfr: boolFromYes(r['Receives Field Report?']),
      receivesInvoice: boolFromYes(r['Receives Invoice?']),
      receivesCp: boolFromYes(r['Receives Certified Payroll?']),
    }).onConflictDoNothing().returning();
    if (row) contractorMap.set(name, row.id);
  }
  // Also create contractors from WO data that might not be in contacts
  const woRows = sheet('Work Order Tracker');
  const woContractorNames = new Set(woRows.map((r: any) => str(r['Prime Contractor'])).filter(Boolean));
  for (const name of woContractorNames) {
    if (!contractorMap.has(name)) {
      const [row] = await db.insert(contractors).values({
        orgId: ORG_ID, name,
      }).onConflictDoNothing().returning();
      if (row) contractorMap.set(name, row.id);
    }
  }
  log(`  ${contractorMap.size} contractors`);

  // If onConflictDoNothing returned nothing (already existed), load from DB
  if (contractorMap.size === 0) {
    const existing = await db.select().from(contractors).where(eq(contractors.orgId, ORG_ID));
    for (const c of existing) contractorMap.set(c.name, c.id);
    log(`  Loaded ${contractorMap.size} existing contractors from DB`);
  }

  // ─── 3. Employees ────────────────────────────────────────
  log('Migrating employees...');
  const empRows = sheet('Employee Registry');
  let empCount = 0;
  for (const r of empRows) {
    const name = str(r['Full Name']);
    if (!name) continue;
    await db.insert(employees).values({
      orgId: ORG_ID,
      name,
      address: str(r['Address']),
      ssnLast4: str(r['Last 4 SSN']),
      raceEthnicity: str(r['Race']),
    }).onConflictDoNothing();
    empCount++;
  }
  log(`  ${empCount} employees`);

  // ─── 4. Contract Lookup ──────────────────────────────────
  log('Migrating contract lookup...');
  const clRows = sheet('Contract Lookup');
  let clCount = 0;
  for (const r of clRows) {
    const cn = str(r['Contract Number']);
    const rc = str(r['Borough Code']);
    if (!cn || !rc) continue;
    await db.insert(contractLookup).values({
      orgId: ORG_ID,
      contractNum: cn,
      regionCode: rc,
      regionName: str(r['Borough Full Name']),
      contractId: str(r['Contract ID / Registration #']),
      projectName: str(r['Project Name / Description']),
    }).onConflictDoNothing();
    clCount++;
  }
  log(`  ${clCount} contract lookup entries`);

  // ─── 5. Contract Pricing ─────────────────────────────────
  log('Migrating contract pricing...');
  const cpRows = sheet('Contract Pricing');
  let cpCount = 0;
  for (const r of cpRows) {
    const contractorName = str(r['Prime Contractor']);
    const cid = contractorMap.get(contractorName);
    if (!cid) { log(`  Warning: contractor "${contractorName}" not found, skipping pricing row`); continue; }
    await db.insert(contractPricing).values({
      orgId: ORG_ID,
      contractorId: cid,
      contractNum: str(r['Contract #']),
      regionCode: str(r['Borough']),
      effectiveDate: parseDate(r['Effective Date']),
      rateLine4: numOrNull(r['4" Line $/LF']),
      rateLine12: numOrNull(r['12" Line $/LF']),
      ratePreformed: numOrNull(r['Preformed L&S $/Unit']),
      rateExtruded: numOrNull(r['Extruded L&S $/Unit']),
      rateColorSurface: numOrNull(r['Color Surface $/SF']),
      notes: str(r['Notes']),
    });
    cpCount++;
  }
  log(`  ${cpCount} contract pricing rows`);

  // ─── 6. Payroll Rates (replace seeded) ───────────────────
  log('Migrating payroll rates...');
  // Clear seeded rates and replace with actual
  await db.delete(payRates).where(eq(payRates.orgId, ORG_ID));
  const prRows = sheet('Payroll Rates');
  let prCount = 0;
  for (const r of prRows) {
    const cls = str(r['Classification']);
    const ed = parseDate(r['Effective Date']);
    if (!cls || !ed) continue;
    await db.insert(payRates).values({
      orgId: ORG_ID,
      classificationCode: cls,
      effectiveDate: ed,
      rateSt: numOrNull(r['ST Rate ($/hr)']) || '0',
      rateOt: numOrNull(r['OT Rate ($/hr)']) || '0',
      suppSt: numOrNull(r['ST Supplemental ($/hr)']) || '0',
      suppOt: numOrNull(r['OT Supplemental ($/hr)']) || '0',
      notes: str(r['Notes']),
    }).onConflictDoNothing();
    prCount++;
  }
  log(`  ${prCount} payroll rate rows`);

  // ─── 7. Work Orders ──────────────────────────────────────
  log('Migrating work orders...');
  const woMap = new Map<string, string>(); // wo_number → uuid
  let woCount = 0;

  for (const r of woRows) {
    const woNum = str(r['Work Order #']);
    if (!woNum) continue;
    const contractorName = str(r['Prime Contractor']);
    const cid = contractorMap.get(contractorName);
    if (!cid) { log(`  Warning: contractor "${contractorName}" not found for WO ${woNum}`); continue; }

    const [wo] = await db.insert(workOrders).values({
      orgId: ORG_ID,
      woNumber: woNum,
      contractorId: cid,
      contractNum: str(r['Contract Number']),
      regionCode: str(r['Borough']),
      contractId: str(r['Contract ID / Reg #']),
      location: str(r['Location']),
      fromStreet: str(r['From Street']),
      toStreet: str(r['To Street']),
      dueDate: parseDate(r['Due Date']),
      priority: str(r['Priority Level']),
      workType: str(r['Pavement Work Type']),
      woReceivedDate: parseDate(r['WO Received Date']),
      waterBlastRequired: str(r['Water Blast Required?']),
      waterBlastConfirmed: str(r['Water Blast Confirmed?']),
      waterBlastSqft: numOrNull(r['Water Blast SQFT']),
      status: (str(r['Status']) || 'received').toLowerCase().replace(/ /g, '_'),
      dispatchDate: parseDate(r['Dispatch Date']),
      workStartDate: parseDate(r['Work Start Date']),
      workEndDate: parseDate(r['Work End Date']),
      issuesReported: str(r['Issues Reported']),
      notes: str(r['Notes']),
      dateEntered: parseDate(r['Date Entered']),
      school: str(r['School']),
      prepBy: str(r['Prep By']),
      generalRemarks: str(r['General Remarks']),
      latitude: numOrNull(r['Latitude']),
      longitude: numOrNull(r['Longitude']),
      geocodeWarning: str(r['Geocode Warning']),
      geocodedAt: parseTimestamp(r['Geocoded At']),
      scanFileKey: str(r['Scan File ID']),  // Drive ID → stored as reference
      originalFilename: str(r['Original Filename']),
    }).onConflictDoNothing().returning();

    if (wo) {
      woMap.set(woNum, wo.id);
      woCount++;
    }
  }
  log(`  ${woCount} work orders`);

  // ─── 8. Marking Items ────────────────────────────────────
  log('Migrating marking items...');
  const miRows = sheet('Marking Items');
  let miCount = 0;
  for (const r of miRows) {
    const woNum = str(r['Work Order #']);
    const woId = woMap.get(woNum);
    if (!woId) continue; // WO not migrated (missing contractor, etc.)

    await db.insert(markingItems).values({
      orgId: ORG_ID,
      woId,
      workType: str(r['Work Type']),
      woSection: (str(r['WO Section']) || 'manual').toLowerCase().replace(/ /g, '_'),
      category: str(r['Marking Type']),
      intersection: str(r['Intersection']),
      direction: str(r['Direction']),
      description: str(r['Description']),
      quantity: numOrNull(r['Quantity Completed']),
      unit: str(r['Unit']),
      colorMaterial: str(r['Color/Material']),
      dateCompleted: parseDate(r['Date Completed']),
      status: (str(r['Status']) || 'pending').toLowerCase(),
      crewChief: str(r['Crew Chief']),
      addedBy: str(r['Added By']) || 'scanner',
      notes: str(r['Notes']),
    });
    miCount++;
  }
  log(`  ${miCount} marking items`);

  // ─── 9. Daily Sign-In Data ───────────────────────────────
  log('Migrating sign-in entries...');
  const siRows = sheet('Daily Sign-In Data');
  let siCount = 0;
  for (const r of siRows) {
    const woNum = str(r['Work Order #']);
    // Handle comma-separated WO numbers — use the first one
    const firstWo = woNum.split(',')[0].trim();
    const woId = woMap.get(firstWo);
    const contractorName = str(r['Prime Contractor']);
    const cid = contractorMap.get(contractorName);
    if (!woId || !cid) continue;

    await db.insert(signinEntries).values({
      orgId: ORG_ID,
      workDate: parseDate(r['Date']) || '',
      woId,
      contractorId: cid,
      contractNum: str(r['Contract #']),
      regionCode: str(r['Borough']),
      location: str(r['Location']),
      employeeName: str(r['Employee Name']),
      classification: str(r['Classification']),
      timeIn: str(r['Time In']),
      timeOut: str(r['Time Out']),
      hoursWorked: numOrNull(r['Hours Worked']),
      otHours: numOrNull(r['Overtime Hours']),
      crewChief: str(r['Crew Chief']),
      adminReviewed: boolFromYes(r['Admin Reviewed?']),
      reviewNotes: str(r['Review Notes']),
    });
    siCount++;
  }
  log(`  ${siCount} sign-in entries`);

  // ─── 10. Work Day Log ────────────────────────────────────
  log('Migrating work day log...');
  const wdlRows = sheet('Work Day Log');
  let wdlCount = 0;
  for (const r of wdlRows) {
    const woNum = str(r['Work Order #']);
    const woId = woMap.get(woNum);
    const contractorName = str(r['Prime Contractor']);
    const cid = contractorMap.get(contractorName);
    if (!woId || !cid) continue;

    await db.insert(workDayLog).values({
      orgId: ORG_ID,
      workDate: parseDate(r['Date']) || '',
      woId,
      contractorId: cid,
      contractNum: str(r['Contract #']),
      regionCode: str(r['Borough']),
      location: str(r['Location']),
      crewChief: str(r['Crew Chief']),
      frSubmittedAt: parseTimestamp(r['Field Report Submitted At']),
      status: (str(r['Sign-In Status']) || 'pending').toLowerCase(),
    });
    wdlCount++;
  }
  log(`  ${wdlCount} work day log entries`);

  // ─── 11. Doc Lifecycle Log ───────────────────────────────
  log('Migrating doc lifecycle...');
  const dlRows = sheet('Doc Lifecycle Log');
  let dlCount = 0;
  for (const r of dlRows) {
    const docId = str(r['Doc ID']);
    if (!docId) continue;

    // Map doc type names to our enum
    const typeMap: Record<string, string> = {
      'Production Log': 'production_log',
      'Sign-In': 'signin',
      'Certified Payroll': 'certified_payroll',
      'Employee Utilization': 'employee_utilization',
      'Certificates': 'certificates',
    };
    const docType = typeMap[str(r['Doc Type'])] || 'production_log';
    const contractorName = str(r['Prime Contractor']);
    const cid = contractorMap.get(contractorName);

    await db.insert(documents).values({
      orgId: ORG_ID,
      docType,
      docKey: docId,
      anchorDate: parseDate(r['Anchor Date']),
      contractorId: cid || undefined,
      contractNum: str(r['Contract #']),
      regionCode: str(r['Borough']),
      woIds: str(r['WO IDs']) ? str(r['WO IDs']).split(',').map((s: string) => s.trim()) : [],
      done: boolFromYes(r['Done']),
      sent: boolFromYes(r['Sent']),
      status: boolFromYes(r['Done']) ? 'archived' : 'needs_review',
      doneAt: parseTimestamp(r['Done At']),
      sentAt: parseTimestamp(r['Sent At']),
      notes: str(r['Notes']),
    }).onConflictDoNothing();
    dlCount++;
  }
  log(`  ${dlCount} doc lifecycle entries`);

  // ─── 12. Certified Payroll Tracker ───────────────────────
  log('Migrating payroll entries...');
  const cpTracker = sheet('Certified Payroll Tracker');
  let peCount = 0;
  for (const r of cpTracker) {
    const empName = str(r['Employee Name']);
    if (!empName) continue;

    await db.insert(payrollEntries).values({
      orgId: ORG_ID,
      weekStart: parseDate(r['Payroll Week (Start)']) || '',
      weekEnd: parseDate(r['Payroll Week (End)']) || '',
      contractNum: str(r['Contract Number']),
      regionCode: str(r['Borough']),
      contractId: str(r['Contract ID / Reg #']),
      projectName: str(r['Project Name']),
      employeeName: empName,
      classification: str(r['Classification']),
      hoursByDay: {
        sun: Number(r['Sun Hrs']) || 0,
        mon: Number(r['Mon Hrs']) || 0,
        tue: Number(r['Tue Hrs']) || 0,
        wed: Number(r['Wed Hrs']) || 0,
        thu: Number(r['Thu Hrs']) || 0,
        fri: Number(r['Fri Hrs']) || 0,
        sat: Number(r['Sat Hrs']) || 0,
      },
      totalSt: numOrNull(r['Total ST Hours']) || '0',
      totalOt: numOrNull(r['Total OT Hours']) || '0',
      rateSt: numOrNull(r['ST Rate']),
      rateOt: numOrNull(r['OT Rate']),
      grossPay: numOrNull(r['Gross Pay']),
      allWorkGross: numOrNull(r['Total Gross (All Work)']),
      deductions: numOrNull(r['Deductions']),
      netPay: numOrNull(r['Net Pay']),
      suppSt: numOrNull(r['ST Supplemental']),
      suppOt: numOrNull(r['OT Supplemental']),
      matchStatus: str(r['Match Status']),
      sentStatus: str(r['Sent?']),
    });
    peCount++;
  }
  log(`  ${peCount} payroll entries`);

  // ─── 13. Invoices ────────────────────────────────────────
  log('Migrating invoices...');
  const invRows = sheet('Invoices & AR');
  let invCount = 0;
  for (const r of invRows) {
    const invNum = str(r['Invoice #']);
    if (!invNum) continue;
    const woNum = str(r['Work Order #']);
    const woId = woMap.get(woNum);
    const contractorName = str(r['Prime Contractor']);
    const cid = contractorMap.get(contractorName);

    await db.insert(invoices).values({
      orgId: ORG_ID,
      invoiceNumber: invNum,
      invoiceDate: parseDate(r['Invoice Date']) || '',
      dueDate: parseDate(r['Due Date']),
      contractorId: cid || undefined,
      contractNum: str(r['Contract #']),
      regionCode: str(r['Borough']),
      woId: woId || undefined,
      description: str(r['Description']),
      sqft: numOrNull(r['SQFT']),
      rate: numOrNull(r['Rate']),
      amount: numOrNull(r['Amount']),
      status: (str(r['Status']) || 'draft').toLowerCase(),
      paymentReceived: boolFromYes(r['Payment Received?']),
      paymentDate: parseDate(r['Payment Date']),
    }).onConflictDoNothing();
    invCount++;
  }
  log(`  ${invCount} invoices`);

  log('\nMigration complete!');
  log(`Summary:
  Organization: updated with company info
  Contractors: ${contractorMap.size}
  Employees: ${empCount}
  Contract Lookup: ${clCount}
  Contract Pricing: ${cpCount}
  Payroll Rates: ${prCount}
  Work Orders: ${woCount}
  Marking Items: ${miCount}
  Sign-In Entries: ${siCount}
  Work Day Log: ${wdlCount}
  Doc Lifecycle: ${dlCount}
  Payroll Entries: ${peCount}
  Invoices: ${invCount}
  `);
}

migrate().then(() => process.exit(0)).catch((e) => {
  console.error('Migration failed:', e);
  process.exit(1);
});
