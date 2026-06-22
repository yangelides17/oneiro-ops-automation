/**
 * ═══════════════════════════════════════════════════════════════
 * ONEIRO COLLECTION LLC — OPERATIONS AUTOMATION ENGINE
 * Google Apps Script (runs free inside Google Workspace)
 * ═══════════════════════════════════════════════════════════════
 * 
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Name it "Oneiro Ops Automation"
 * 3. Paste this entire file into Code.gs
 * 4. Run setupAutomation() once to create folder structure + triggers
 * 5. When prompted, authorize the script to access Drive, Sheets, Gmail
 * 
 * CONFIGURATION: Update the CONFIG object below with your actual IDs
 *
 * CHANGELOG:
 *  v1.2 — 2026-04-10 — Removed daily summary email trigger (sendDailySummary)
 */

// ── CONFIGURATION ──────────────────────────────────────────────
const CONFIG = {
  // After uploading the spreadsheet to Google Drive and converting to Sheets,
  // paste the Spreadsheet ID here (from the URL: docs.google.com/spreadsheets/d/{THIS_ID}/edit)
  SPREADSHEET_ID: '1gYV3qk7I1wON-ehebt6AB3X5HdDA4IPO70F8nuqrWzw',
  
  // These will be auto-populated by setupAutomation()
  // but you can also set them manually if the folders already exist
  ROOT_FOLDER_ID: '',
  SCAN_INBOX_ID: '',
  NEEDS_REVIEW_ID: '',
  APPROVED_SENT_ID: '',
  ARCHIVE_ID: '',
  TEMPLATES_ID: '',
  
  // Email addresses for document delivery
  // (looked up from Contractor Contacts sheet, but fallback defaults here)
  ADMIN_EMAIL: 'yangelides17@gmail.com',  // test email — 2026-04-10

  // Employer info stamped into the Certified Payroll JSON export
  EMPLOYER: {
    name:    'Oneiro Collection LLC',
    address: '435 South Avenue Apt 220, Garwood, NJ 07027',
    email:   'anthe@theoneiro.com',
    phone:   '',
    tax_id:  '0450359615'
  },

  // Signatory for certified payroll
  SIGNATORY: {
    name:  'Marianthy Angelides',
    title: 'Principal'
  },

  // Claude API key for OCR parsing (set this when you add n8n)
  // For now, the scan inbox just logs files for manual data entry
  CLAUDE_API_KEY: '',

  // Timezone
  TIMEZONE: 'America/New_York',

  // Prime contractors that get daily production logs generated. Each
  // listed contractor gets ONE production log per day (filtered to
  // their own WOs only). Other primes' WOs are silently skipped — they
  // don't require production logs today. Add a name here AND drop a
  // matching template in workers/templates + register it in the
  // Python worker's _TEMPLATE_BY_CONTRACTOR dict to enable a new prime.
  PRODUCTION_LOG_CONTRACTORS: ['Metro Express'],

  // Operational-day cutoff hour (local time). Any timestamp with hour
  // strictly less than this value is bucketed to the PREVIOUS calendar
  // day for accounting purposes (Daily Sign-In Data Date column, OT
  // rule day-of-week, "Completed Today" filters, etc). 5 covers the
  // typical road-striping night shift that wraps up between 4–5 AM
  // (FR submitted just after work ends still buckets to the previous
  // evening). Edge case: a true 3–5 AM day shift would mis-bucket
  // backward — use the Sign-In tab kebab to override on those.
  OPERATIONAL_DAY_CUTOFF_HOUR: 5
};


/**
 * Operational-day bucket. Returns YYYY-MM-DD (local time, no UTC shift).
 * If the given Date's local hour is strictly below CONFIG cutoff,
 * the date returned is the calendar day BEFORE — so 02:30 on Tue
 * with cutoff=4 returns Mon's date.
 *
 * Used everywhere a date represents an "accounting day" (when did this
 * shift effectively start) rather than a literal calendar day. The
 * canonical write-points are handleSubmitFieldReport_ and
 * handleSubmitSignIn_; downstream readers (cert payroll, production
 * log, "Completed Today") just consume the resulting Date column.
 */
function opDay_(date, cutoffHour) {
  const cutoff = (typeof cutoffHour === 'number')
    ? cutoffHour
    : (CONFIG.OPERATIONAL_DAY_CUTOFF_HOUR || 4);
  const d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  const target = (d.getHours() < cutoff)
    ? new Date(d.getFullYear(), d.getMonth(), d.getDate() - 1)
    : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const yyyy = target.getFullYear();
  const mm   = String(target.getMonth() + 1).padStart(2, '0');
  const dd   = String(target.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Combine an ISO YYYY-MM-DD calendar date with an HH:MM time of day
 * (24-hour) into a local-parts Date, then opDay_() it. This is the
 * authoritative shift-start date when both the calendar date and the
 * Time In are known (e.g. Sign-In submit).
 */
function opDayFromIsoTime_(isoDate, hhmm, cutoffHour) {
  const md = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  const tm = String(hhmm   || '').match(/^(\d{1,2}):(\d{2})/);
  if (!md || !tm) return String(isoDate || '');
  const dt = new Date(
    Number(md[1]), Number(md[2]) - 1, Number(md[3]),
    Number(tm[1]), Number(tm[2])
  );
  return opDay_(dt, cutoffHour);
}

/** Today's operational day in the script timezone. */
function opToday_(cutoffHour) {
  return opDay_(new Date(), cutoffHour);
}


// ═══════════════════════════════════════════════════════════════
// 1. INITIAL SETUP — Run once to create folder structure + triggers
// ═══════════════════════════════════════════════════════════════

function setupAutomation() {
  Logger.log('🚀 Setting up Oneiro Operations automation...');
  
  // Create folder structure
  const rootFolder = createFolderStructure_();
  
  // Set up time-based triggers
  setupTriggers_();
  
  // Log completion
  Logger.log('✅ Setup complete!');
  Logger.log('Root folder ID: ' + rootFolder.getId());
  Logger.log('Update CONFIG.ROOT_FOLDER_ID with this value');
  
  // Store folder IDs in script properties for persistence
  const props = PropertiesService.getScriptProperties();
  Logger.log('\n📋 Folder IDs (saved to script properties):');
  Logger.log('ROOT_FOLDER_ID: ' + props.getProperty('ROOT_FOLDER_ID'));
  Logger.log('SCAN_INBOX_ID: ' + props.getProperty('SCAN_INBOX_ID'));
  Logger.log('NEEDS_REVIEW_ID: ' + props.getProperty('NEEDS_REVIEW_ID'));
  Logger.log('APPROVED_SENT_ID: ' + props.getProperty('APPROVED_SENT_ID'));
  Logger.log('ARCHIVE_ID: ' + props.getProperty('ARCHIVE_ID'));
  Logger.log('ARCHIVE_ERRORS_ID: ' + props.getProperty('ARCHIVE_ERRORS_ID'));
  Logger.log('TEMPLATES_ID: ' + props.getProperty('TEMPLATES_ID'));
}

function createFolderStructure_() {
  const root = DriveApp.createFolder('📁 Oneiro Operations');
  const props = PropertiesService.getScriptProperties();
  
  // Top-level folders
  const scanInbox = root.createFolder('📥 Scan Inbox');
  const needsReview = root.createFolder('📋 Needs Review');
  const approvedSent = root.createFolder('✅ Approved Docs');
  const archive = root.createFolder('🗂️ Archive');
  const reports = root.createFolder('📊 Reports');
  const templates = root.createFolder('⚙️ Templates');

  // Needs Review subfolders
  needsReview.createFolder('Production Logs');
  needsReview.createFolder('Field Reports');
  needsReview.createFolder('Invoices');
  needsReview.createFolder('Certified Payroll');

  // Archive subfolders by contractor
  const contractors = ['Metro Express', 'Denville', 'Delan'];
  contractors.forEach(contractor => {
    archive.createFolder(contractor);
  });

  // Fallback bucket for files the approved-docs cron couldn't archive
  // (bad filename, missing tracker row, etc).  Kept inside Archive so
  // admins can triage it alongside the rest of the archived content.
  const archiveErrors = archive.createFolder('⚠️ Archive Errors');
  
  // Reports subfolders
  reports.createFolder('Weekly Payroll Reports');
  reports.createFolder('Monthly Workforce Utilization');
  
  // Save folder IDs
  props.setProperty('ROOT_FOLDER_ID', root.getId());
  props.setProperty('SCAN_INBOX_ID', scanInbox.getId());
  props.setProperty('NEEDS_REVIEW_ID', needsReview.getId());
  props.setProperty('APPROVED_SENT_ID', approvedSent.getId());
  props.setProperty('ARCHIVE_ID', archive.getId());
  props.setProperty('ARCHIVE_ERRORS_ID', archiveErrors.getId());
  props.setProperty('TEMPLATES_ID', templates.getId());
  
  return root;
}

function setupTriggers_() {
  // Remove existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // Check scan inbox every 15 minutes
  ScriptApp.newTrigger('checkScanInbox')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  // Check for approved documents every 10 minutes
  ScriptApp.newTrigger('processApprovedDocuments')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  // Daily summary email — DISABLED 2026-04-10 (removed per request)
  // ScriptApp.newTrigger('sendDailySummary')
  //   .timeBased()
  //   .atHour(18)
  //   .everyDays(1)
  //   .inTimezone(CONFIG.TIMEZONE)
  //   .create();

  // Installable onOpen trigger — required for standalone scripts so the
  // custom menu appears automatically when the spreadsheet is opened.
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
    .onOpen()
    .create();

  Logger.log('⏰ Triggers set up: Scan Inbox (15min), Approved Docs (10min), onOpen menu');
}


// ═══════════════════════════════════════════════════════════════
// MARKING ITEMS — per-item marking completion schema (MMA + Thermo)
// ═══════════════════════════════════════════════════════════════
//
// `Marking Items` captures one row per discrete piece of marking work
// on a WO — top-table items, per-intersection per-crosswalk entries, or
// manually added items. The parser seeds rows at scan time; the Field
// Report UI loads them for the crew to enter SF + material; rollups
// flow back to WO Tracker cols 19-21.
//
// Run once from the custom menu after each schema change. Idempotent.
//
// For multi-crew support, run setupMultiCrewSchema() AFTER this to
// add the Crew Chief column to both Marking Items (col 16) and Daily
// Sign-In Data (col 13). Doing it as a separate step lets re-runs of
// setupMarkingItems stay forward-compatible with existing crew-chief-
// tagged data on prod (this function only sets headers + validations
// for the original schema — it doesn't touch col 16).

function setupMarkingItems() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const MAX_ROWS = 2000;

  // ── 1. Marking Items tab ──────────────────────────────────────
  //
  // Ordering: rows are written to the sheet contiguously per WO (all
  // seed rows from a scan are in one setValues call; manual rows get
  // appended at submit time). Filtering by WO# preserves insertion
  // order. No "Sort Order" column needed — eliminates an arithmetic
  // collision risk at ~80+ intersections or 1000+ manual adds.
  let markingSheet = ss.getSheetByName('Marking Items');
  const createdNew = !markingSheet;
  if (createdNew) {
    markingSheet = ss.insertSheet('Marking Items');
  } else {
    // Migration from earlier schemas — detect and drop legacy cols.
    const colEHeader = String(markingSheet.getRange(1, 5).getValue() || '').toLowerCase();
    if (colEHeader.indexOf('sort order') !== -1) {
      markingSheet.deleteColumns(5, 1);
      Logger.log('↻ Dropped legacy "Sort Order" column — ordering now implicit');
    }
    const colIHeader = String(markingSheet.getRange(1, 9).getValue() || '').toLowerCase();
    if (colIHeader.indexOf('planned') !== -1) {
      markingSheet.deleteColumns(9, 1);
      Logger.log('↻ Dropped legacy "Planned" column — redundant with Category + Direction');
    }
  }

  const mHeaders = [
    'Item ID', 'Work Order #', 'Work Type', 'WO Section',
    'Marking Type', 'Intersection', 'Direction', 'Description',
    'Quantity Completed', 'Unit', 'Color/Material',
    'Date Completed', 'Status', 'Added By', 'Notes'
  ];
  markingSheet.getRange(1, 1, 1, mHeaders.length).setValues([mHeaders]);
  markingSheet.getRange(1, 1, 1, mHeaders.length).setFontWeight('bold');
  markingSheet.setFrozenRows(1);

  // Canonical marking categories — covers Page 1 (scan time) + Page 2
  // (Contractor Field Report) + MMA types. `strict: false` turns the
  // dropdown into a picker/suggestion list so manually added items can
  // introduce new categories without being rejected.
  const MARKING_CATEGORIES = [
    // WO Page 1 — Top Table
    'Double Yellow Line', 'Lane Lines', 'Gores', 'Messages', 'Arrows',
    'Solid Lines', 'Rail Road X/Diamond', 'Others',
    // WO Page 1 — Intersection Grid
    'HVX Crosswalk', 'Stop Msg', 'Stop Line',
    // Page 2 — detailed lines
    '4" Line', '6" Line', '8" Line', '12" Line', '16" Line', '24" Line',
    // Page 2 — messages
    'Only Msg', 'Bus Msg', 'Bump Msg', 'Custom Msg', '20 MPH Msg',
    // Page 2 — railroad
    'Railroad (RR)', 'Railroad (X)',
    // Page 2 — arrows
    'L/R Arrow', 'Straight Arrow', 'Combination Arrow', 'Combination Arrow (L/R)',
    // Page 2 — miscellaneous
    'Speed Hump Markings', 'Shark Teeth 12x18', 'Shark Teeth 24x36',
    // Page 2 — bike lane (old/new preform bike symbol; both map to the
    // same Bike Symbol CFR/PL field and sum, but price differently)
    'Bike Lane Arrow', 'Old Bike Symbol (w/ rider)', 'New Bike Symbol (just bike)',
    'Bike Lane Green Bar',
    // Preform pedestrian symbol
    'Pedestrian Men',
    // MMA
    'Bike Lane', 'Pedestrian Space', 'Bus Lane',
  ];

  // Column indices (1-based) for the 15-col schema:
  //   A  1 Item ID         E  5 Marking Type    I   9 Quantity Completed
  //   B  2 Work Order #    F  6 Intersection    J  10 Unit
  //   C  3 Work Type       G  7 Direction       K  11 Color/Material
  //   D  4 WO Section      H  8 Description     L  12 Date Completed
  //                                              M  13 Status
  //                                              N  14 Added By
  //                                              O  15 Notes

  // Clear ALL existing data validations on rows 2+ first — otherwise
  // stale validators left over from prior schemas (e.g. an old SF/LF/EA
  // dropdown hanging on what is now the Quantity column) will reject
  // legitimate input after a column swap or header change.
  markingSheet.getRange(2, 1, MAX_ROWS - 1, mHeaders.length).clearDataValidations();

  const mDropdowns = [
    { col:  3, values: ['MMA', 'Thermo'],                              strict: true  },
    { col:  4, values: ['Top Table', 'Intersection Grid', 'Manual'],   strict: true  }, // WO Section
    { col:  5, values: MARKING_CATEGORIES,                             strict: false }, // Marking Type
    { col:  7, values: ['N', 'E', 'S', 'W'],                           strict: true  }, // Direction
    { col: 10, values: ['SF', 'LF', 'EA'],                             strict: true  }, // Unit
    { col: 13, values: ['Pending', 'Completed', 'Skipped'],            strict: true  }, // Status
    { col: 14, values: ['Scanner', 'Manual'],                          strict: true  }, // Added By
  ];
  mDropdowns.forEach(({ col, values, strict }) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(!strict)
      .build();
    markingSheet.getRange(2, col, MAX_ROWS - 1, 1).setDataValidation(rule);
  });

  Logger.log(createdNew
    ? '✅ Marking Items sheet created with 15-col header + validations'
    : '↻ Marking Items header/validations reconciled');

  // ── 2. Daily Sign-In Data → 14-col schema ─────────────────────
  // Old schema (18 cols) had SQFT Completed / Paint/Material /
  // WO Complete? / Issues-Notes at cols 13-16. Those four columns
  // migrate to Marking Items (per-item) and WO Tracker rollups.
  const signInSheet = ss.getSheetByName('Daily Sign-In Data');
  if (!signInSheet) {
    Logger.log('⚠️  Daily Sign-In Data sheet not found — skipping schema update');
  } else {
    // Detect old schema by the col-13 header; drop those 4 cols if present.
    const col13Header = String(signInSheet.getRange(1, 13).getValue() || '').toLowerCase();
    if (col13Header.indexOf('sqft') !== -1) {
      signInSheet.deleteColumns(13, 4);
      Logger.log('↻ Dropped legacy cols 13-16 (SQFT/Paint/WO Complete/Issues) from Daily Sign-In Data');
    }

    const sHeaders = [
      'Date', 'Work Order #', 'Prime Contractor', 'Contract #', 'Borough',
      'Location', 'Employee Name', 'Classification', 'Time In', 'Time Out',
      'Hours Worked', 'Overtime Hours', 'Admin Reviewed?', 'Review Notes'
    ];
    signInSheet.getRange(1, 1, 1, sHeaders.length).setValues([sHeaders]);
    signInSheet.getRange(1, 1, 1, sHeaders.length).setFontWeight('bold');
    signInSheet.setFrozenRows(1);

    // Re-apply Admin Reviewed? dropdown on new col 13 (was col 17 before drop).
    const adminRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .build();
    signInSheet.getRange(2, 13, MAX_ROWS - 1, 1).setDataValidation(adminRule);

    Logger.log('✅ Daily Sign-In Data reshaped to 14-col schema + Admin Reviewed? dropdown reapplied');
  }

  Logger.log('🚀 setupMarkingItems() complete');
  try {
    SpreadsheetApp.getUi().alert(
      createdNew
        ? 'Marking Items sheet created + Daily Sign-In Data reshaped. See Logs for details.'
        : 'Schema reconciled. See Logs for details.'
    );
  } catch (_) {
    // No UI context (e.g. running from script editor without active sheet) — skip alert.
  }
}


/**
 * One-shot, idempotent schema migration for multi-crew support. Adds a
 * `Crew Chief` column to:
 *   - Work Day Log         → col 8 (between "Field Report Submitted At"
 *                                    and "Sign-In Status")
 *   - Daily Sign-In Data   → col 13 (after "Overtime Hours")
 *   - Marking Items        → col 16 (after "Notes")
 *
 * Run once from the Apps Script editor. Re-runs are safe (per-sheet
 * check skips columns already present). Existing rows stay blank; the
 * blank value is first-class throughout — legacy data behaves exactly
 * like today, multi-crew tagging only kicks in for new submissions.
 */
function setupMultiCrewSchema() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const HEADER = 'Crew Chief';
  const log = [];

  const addColumn = (sheetName, insertBeforeCol1) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      log.push(`⚠️ ${sheetName}: sheet missing — skip`);
      return;
    }
    const lastCol = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headerRow.indexOf(HEADER) !== -1) {
      log.push(`↻ ${sheetName}: Crew Chief already present — skip`);
      return;
    }
    sheet.insertColumnBefore(insertBeforeCol1);
    sheet.getRange(1, insertBeforeCol1).setValue(HEADER).setFontWeight('bold');
    log.push(`✅ ${sheetName}: inserted Crew Chief at col ${insertBeforeCol1}`);
  };

  // WDL: insert at col 8 (existing col 8 "Sign-In Status" shifts to col 9).
  addColumn('Work Day Log', 8);

  // Daily Sign-In Data: insert at col 13 (after Overtime Hours at col 12).
  // The sheet may already have admin/review cols past 12 — insertColumnBefore
  // shifts them right, which is what we want.
  addColumn('Daily Sign-In Data', 13);

  // Marking Items: insert at col 16 (after Notes at col 15).
  addColumn('Marking Items', 16);

  Logger.log('🚀 setupMultiCrewSchema():\n  ' + log.join('\n  '));

  try {
    SpreadsheetApp.getUi().alert('Multi-crew schema:\n\n' + log.join('\n'));
  } catch (_) { /* no UI context */ }
}


/**
 * Idempotently creates the Contract Pricing sheet that the Revenue
 * dashboard + future invoice generator both read from. Keyed on
 * (Prime Contractor, Contract #, Borough) with an Effective Date
 * column so contract amendments don't lose history.
 *
 * Schema (1-indexed cols):
 *   A Prime Contractor   F 12" Line $/LF   (covers HVX Crosswalk + Stop Line —
 *   B Contract #                            wording matches contractor expectations)
 *   C Borough            G Preformed L&S $/Unit
 *   D Effective Date     H Extruded L&S $/Unit
 *   E 4" Line $/LF       I Color Surface $/SF
 *                        J Notes
 *
 * Effective Date semantics: lookup picks the most recent row with
 * Effective Date <= item's Date Completed. A blank Effective Date is
 * treated as "effective forever" — only used when no dated row matches.
 */
function setupContractPricing() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Contract Pricing');
  const createdNew = !sheet;
  if (createdNew) sheet = ss.insertSheet('Contract Pricing');

  const headers = [
    'Prime Contractor', 'Contract #', 'Borough', 'Effective Date',
    '4" Line $/LF', '12" Line $/LF', 'Preformed L&S $/Unit',
    'Extruded L&S $/Unit', 'Color Surface $/SF', 'Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Date validator on col D (Effective Date). Allow blank.
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, 999, 1).setDataValidation(dateRule);

  // A2 is intentionally left blank for the user's first contract row;
  // documenting the blank-date rule lives in the Notes col instead so
  // it doesn't conflict with row data.
  Logger.log(createdNew
    ? '✅ Contract Pricing sheet created'
    : '↻ Contract Pricing schema reconciled');
}


/**
 * Payroll Rates: prevailing-wage schedule keyed by employee
 * classification with an effective-date column. Mirrors the Contract
 * Pricing pattern — newest dated row whose Effective Date <= the
 * payroll week-end date wins (see _resolvePayrollRate_).
 *
 * Seeded with the LP and SAT rates effective 2025-07-01. When the
 * City publishes new rates, append a new row with the future
 * Effective Date — no code changes needed.
 *
 * Idempotent: if the sheet already exists, leaves it alone (only
 * reconciles headers). Safe to run anytime.
 */
function setupPayrollRates() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Payroll Rates');
  const createdNew = !sheet;
  if (createdNew) sheet = ss.insertSheet('Payroll Rates');

  const headers = [
    'Classification', 'Effective Date',
    'ST Rate ($/hr)', 'OT Rate ($/hr)',
    'ST Supplemental ($/hr)', 'OT Supplemental ($/hr)',
    'Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Classification dropdown (LP/SAT) on col A.
  const classRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['LP', 'SAT'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 1, 999, 1).setDataValidation(classRule);

  // Date validator on col B.
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, 999, 1).setDataValidation(dateRule);

  if (createdNew) {
    // Seed the current schedule (effective 2025-07-01, last day 2026-06-20).
    // OT rate = 1.5 × ST rate per the prevailing-wage spec.
    const seedDate = new Date(2025, 6, 1); // July is month index 6
    sheet.getRange(2, 1, 2, headers.length).setValues([
      ['LP',  seedDate, 46.00, 69.00, 21.17, 22.87, 'Line Person — schedule effective through 2026-06-20'],
      ['SAT', seedDate, 40.00, 60.00, 21.17, 22.87, 'Striping Assistant — schedule effective through 2026-06-20'],
    ]);
    sheet.getRange(2, 2, 999, 1).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(2, 3, 999, 4).setNumberFormat('$#,##0.00');
  }

  Logger.log(createdNew
    ? '✅ Payroll Rates sheet created + seeded with LP/SAT @ 2025-07-01'
    : '↻ Payroll Rates schema reconciled');
}


// ═══════════════════════════════════════════════════════════════
// 2. SCAN INBOX WATCHER — Detects new WO PDFs dropped in folder
// ═══════════════════════════════════════════════════════════════

function checkScanInbox() {
  const props = PropertiesService.getScriptProperties();
  const inboxId = props.getProperty('SCAN_INBOX_ID');
  if (!inboxId) { Logger.log('❌ Scan Inbox folder ID not set'); return; }
  
  const inbox = DriveApp.getFolderById(inboxId);
  const files = inbox.getFiles();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const logSheet = ss.getSheetByName('Automation Log');
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Skip already-processed files (marked with ✅ prefix)
    if (fileName.startsWith('✅')) continue;
    
    // Log the detection
    logSheet.appendRow([
      new Date(),
      'Scan Inbox Watcher',
      'New file detected',
      fileName,
      'Pending manual data entry',
      'Detected',
      '',
      'Yes — Enter WO data into tracker'
    ]);
    
    // Rename file to mark as detected (prevents re-processing)
    file.setName('✅ ' + fileName);
    
    // Try to extract Work Order number from filename for smart filing
    const woMatch = fileName.match(/PT[-_]?(\d+)/i);
    if (woMatch) {
      Logger.log('📄 Detected WO: PT-' + woMatch[1] + ' from file: ' + fileName);
    }
    
    // Send notification email to admin
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: '📄 New Work Order Scanned: ' + fileName,
      htmlBody: `
        <h3>New Work Order Detected</h3>
        <p>A new file has been added to the Scan Inbox:</p>
        <p><strong>${fileName}</strong></p>
        <p><a href="https://drive.google.com/file/d/${file.getId()}/view">View File</a></p>
        <p><strong>Action Required:</strong> Enter this work order's data into the 
        <a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit">Work Order Tracker</a>.</p>
        <hr>
        <p style="color: #666; font-size: 12px;">Oneiro Operations Automation</p>
      `
    });
    
    Logger.log('📧 Notification sent for: ' + fileName);
  }
}


// ═══════════════════════════════════════════════════════════════
// 3. DOCUMENT GENERATORS — Create docs from tracker data
// ═══════════════════════════════════════════════════════════════

/**
 * Generate all daily documents for a given date.
 * Call this from the spreadsheet via a custom menu or manually.
 *
 * @param {string} dateStr - Date in MM/DD/YYYY format. Defaults to the
 *   current OPERATIONAL day, not raw calendar today — so a 3 AM run
 *   correctly bucketed to yesterday's shift will pick up yesterday's
 *   Daily Sign-In Data rows.
 */
function generateDailyDocuments(dateStr, opts) {
  let targetDate;
  if (dateStr) {
    targetDate = new Date(dateStr);
  } else {
    // Construct a local-parts Date from opToday_() so .toDateString()
    // comparisons against Daily Sign-In Data's Date column work
    // regardless of timezone shifts.
    const iso = opToday_();
    const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})/);
    targetDate = m
      ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0)
      : new Date();
  }
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Get all sign-in data for this date
  const signInSheet = ss.getSheetByName('Daily Sign-In Data');
  const signInData = signInSheet.getDataRange().getValues();
  const headers = signInData[0];

  const todaysEntries = signInData.slice(1).filter(row => {
    if (!row[0]) return false;
    const rowDate = new Date(row[0]);
    return rowDate.toDateString() === targetDate.toDateString();
  });

  if (todaysEntries.length === 0) {
    Logger.log('No entries found for ' + targetDate.toDateString());
    return { generated: [] };
  }

  // Group entries by Work Order. WO# column may be a comma-list when one
  // sign-in covers multiple WOs in the same shift — fan each row out into
  // every WO it references so each gets the same crew context.
  const byWorkOrder = {};
  todaysEntries.forEach(row => {
    _splitWOIds_(row[1]).forEach(woId => {
      if (!byWorkOrder[woId]) byWorkOrder[woId] = [];
      byWorkOrder[woId].push(row);
    });
  });

  // Get Work Order details from tracker
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData = woSheet.getDataRange().getValues();
  const woHeaders = woData[0];

  Logger.log(`📋 Processing ${Object.keys(byWorkOrder).length} work orders for ${targetDate.toDateString()}`);

  // Generate Production Log (Metro Express format)
  const plOpts = (opts && opts.contractorFilter) ? { contractorFilter: opts.contractorFilter } : undefined;
  const plFiles = generateProductionLog_(targetDate, todaysEntries, byWorkOrder, woData, ss, plOpts) || [];

  // Per-doc generation (PL only) skips Field Reports + Invoices — those
  // are tied to WO completion events, not the per-day PL flow.
  if (!opts || !opts.skipFieldReportsAndInvoices) {
    Object.entries(byWorkOrder).forEach(([woId, entries]) => {
      const isComplete = entries.some(e => String(e[14]).toLowerCase() === 'yes');
      if (isComplete) {
        generateFieldReport_(woId, entries, woData, ss);
        generateInvoice_(woId, entries, woData, ss);
      }
    });
  }

  return { generated: plFiles.map(f => f.getName()) };
}

/** Lazy-cached spreadsheet timezone, so repeated formatTime_ calls
 *  don't hit SpreadsheetApp on every row. */
let _ssTzCache_ = null;
function _getSsTz_() {
  if (_ssTzCache_) return _ssTzCache_;
  try {
    _ssTzCache_ = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSpreadsheetTimeZone();
  } catch (e) {
    _ssTzCache_ = CONFIG.TIMEZONE;   // fallback
  }
  return _ssTzCache_;
}


/**
 * Normalize any time-ish cell value to "h:mm AM/PM".
 *
 * Sheets stores a time-only value as a Date pinned to 1899-12-30. When
 * we read via getValues() that comes back as a Date object, and raw
 * JSON-stringifying it produces "1899-12-30T07:00:00.000Z" — ugly on
 * the downstream PDF. This helper handles all four shapes we see:
 *   Date object (1899-12-30 epoch) → formatted in script TZ
 *   "7:00 AM" / "12:30 PM"         → normalized (leading zeros stripped)
 *   "07:00" / "13:30" (24h)        → converted
 *   anything else                  → pass through
 */
function formatTime_(val) {
  if (val == null || val === '') return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    // Format in the SPREADSHEET's TZ, not the script's. Sheets parses
    // time strings against its own TZ when storing them, so reading
    // back and rendering in that same TZ gives us the string the user
    // actually typed. Cached per-execution for cheap reuse.
    const tz = _getSsTz_();
    return Utilities.formatDate(val, tz, 'h:mm a');
  }
  const s = String(val).trim();
  if (!s) return '';
  // Already "h:mm AM/PM" — strip leading zero on the hour
  let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if (m) return `${parseInt(m[1], 10)}:${m[2]} ${m[3].toUpperCase()}`;
  // 24-hour "HH:MM" or "HH:MM:SS"
  m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) {
    let h = parseInt(m[1], 10);
    const mm = m[2];
    const ampm = h < 12 ? 'AM' : 'PM';
    if (h === 0) h = 12;
    else if (h > 12) h -= 12;
    return `${h}:${mm} ${ampm}`;
  }
  return s;
}


/**
 * Parse a time string like "9:15 AM" or "1:15 PM" (or the Sheets
 * epoch-Date form) into minutes since midnight. Used for correct
 * earliest-in / latest-out comparisons.
 */
function parseTimeToMinutes_(timeStr) {
  if (timeStr == null || timeStr === '') return 0;
  if (timeStr instanceof Date && !isNaN(timeStr.getTime())) {
    // Same TZ rationale as formatTime_ — compute hours/minutes in the
    // spreadsheet's TZ so 'earliest in' / 'latest out' reflects what
    // the user actually sees in the cell.
    const tz = _getSsTz_();
    const hhmm = Utilities.formatDate(timeStr, tz, 'H:mm');
    const [h, m] = hhmm.split(':').map(n => parseInt(n, 10));
    return (isNaN(h) ? 0 : h) * 60 + (isNaN(m) ? 0 : m);
  }
  const s = String(timeStr).trim().toUpperCase();
  const match = s.match(/(\d+):(\d+)\s*(AM|PM)/);
  if (!match) {
    // 24-hour fallback
    const m24 = s.match(/^(\d{1,2}):(\d{2})/);
    if (m24) return parseInt(m24[1], 10) * 60 + parseInt(m24[2], 10);
    return 0;
  }
  let h = parseInt(match[1]);
  const m = parseInt(match[2]);
  const ampm = match[3];
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return h * 60 + m;
}

/**
 * Marking Items category → PL template printed label. Used by
 * aggregateMarkingItemsForPL_ to map crew-entered Marking Items rows
 * into the strings the Production Log PDF filler expects in its
 * per-WO `markings` dict.
 *
 * Categories not present in this map are intentionally NOT rendered in
 * the PL grid:
 *  - Parent-only categories (Gores, Messages, Arrows, Solid Lines,
 *    Rail Road X/Diamond) never carry quantities.
 *  - Ops-review items (Ped Stop, Shark Teeth 12x18, Bike Lane Green Bar,
 *    Custom Msg) will get a home if/when the ops manager assigns them.
 *  - MMA SF categories (Bike Lane / Bus Lane / Pedestrian Space) are
 *    aggregated separately into the PL's Color Surface Treatment
 *    rows via the `sqft` + `paint` fields — NOT into this marking grid.
 *
 * Two special cases are handled outside this map:
 *  - 'HVX Crosswalk' + 'Stop Line' LF → sum into 'CrossWalks/Stop Lines'.
 *  - MMA SF → sqft + paint.
 */
const PL_CATEGORY_MAP_ = {
  // LF
  'Double Yellow Line':  'Double Yellow Line (Center Line)',
  'Lane Lines':          'Lane Lines 4" (Skips)',
  '4" Line':             '4" Lines',
  '6" Line':             '6" Lines',
  '8" Line':             '8" Lines',
  '12" Line':            '12" Lines (Gore)',
  '16" Line':            '16" Lines',
  '24" Line':            '24" Lines',
  // EA
  'Stop Msg':            'Stop Message',
  'Only Msg':            'Message Only',
  'Bus Msg':             'Bus Message',
  'Bump Msg':            'Bump',
  '20 MPH Msg':          '20 MPH Message',
  'Railroad (RR)':       'Railroad - RR',
  'Railroad (X)':        'Railroad - X',
  'L/R Arrow':           'Left & or Right Arrows',
  'Straight Arrow':      'Straight Arrow',
  'Combination Arrow':   'Combination Arrow',
  // Combination Arrow (L/R) rolls up into the same PL row as
  // Combination Arrow; aggregator at aggregateMarkingItemsForPL_ sums
  // them naturally via the shared label.
  'Combination Arrow (L/R)': 'Combination Arrow',
  'Speed Hump Markings': 'Speed Hump Marking',
  'Shark Teeth 24x36':   'Sharks Teeth 24" 36"',
  'Bike Lane Arrow':     'Bicycle Lane Arrow',
  'Bike Lane Symbol':    'Bicycle Lane Symbol',   // legacy alias
  // Old + new preform bike symbols share the PL row; aggregator sums them.
  'Old Bike Symbol (w/ rider)':  'Bicycle Lane Symbol',
  'New Bike Symbol (just bike)': 'Bicycle Lane Symbol',
};


/**
 * Aggregate all Completed Marking Items for a single WO into the shape
 * the Production Log filler expects on a per-WO-column basis:
 *
 *   {
 *     markings: { '<PL row label>': <qty>, ... },   // grid cells
 *     sqft:     <total MMA SF across this WO, or ''>, // Color Surface Treatment 1
 *     paint:    '<color or comma-joined colors>',    // Color Surface Treatment 2
 *   }
 *
 * Only rows with Status='Completed' and quantity > 0 are counted.
 * Categories not in PL_CATEGORY_MAP_ are ignored (except HVX Crosswalk,
 * Stop Line, and the MMA SF trio, which have their own paths).
 */
function aggregateMarkingItemsForPL_(ss, woId, targetDateIso, crewChief) {
  const out = { markings: {}, sqft: '', paint: '' };
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return out;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return out;

  // Col indices (post-multi-crew schema):
  //   1 WO#, 4 Marking Type, 5 Intersection, 6 Direction,
  //   8 Quantity, 9 Unit, 10 Color/Material,
  //  11 Date Completed (set by finalizeMarkingStatus_ at FR submit),
  //  12 Status, 15 Crew Chief (tagged at completion time)
  const tgt = String(targetDateIso || '').slice(0, 10);
  const matchesDate = (cell) => {
    if (!tgt) return true;          // no target date filter (legacy callers)
    if (cell instanceof Date && !isNaN(cell.getTime())) {
      return Utilities.formatDate(cell, CONFIG.TIMEZONE, 'yyyy-MM-dd') === tgt;
    }
    return String(cell || '').slice(0, 10) === tgt;
  };
  // Strict equality on chief: blank crewChief matches only blank-tagged
  // (legacy) items; populated crewChief matches only items tagged by
  // that crew. Per-crew PLs see only their own portion of a WO that was
  // worked across multiple shifts on the same day.
  const chiefFilter = String(crewChief || '').trim();
  const matchesChief = (cell) => {
    return String(cell || '').trim() === chiefFilter;
  };
  const woItems = data.slice(1).filter(r =>
    String(r[1]  || '').trim() === woId &&
    String(r[12] || '').toLowerCase() === 'completed' &&
    matchesDate(r[11]) &&
    matchesChief(r[15])
  );
  if (woItems.length === 0) return out;

  let sfSum = 0;
  const colorsSet = {};
  let crosswalkSum = 0;
  let stopLineSum  = 0;
  let pedMenSum    = 0;

  woItems.forEach(r => {
    const category = String(r[4] || '').trim();
    const qty      = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;
    const unit     = String(r[9] || '').toUpperCase();

    // MMA SF → Color Surface Treatment 1/2, not the grid.
    // Bike Lane Green Bar joins the same SF rollup — it's a colored
    // surface treatment by definition. Color/Material is optional in
    // the UI but admins should set it to "Green" (or whatever specific
    // shade) so the Color Surface Treatment 2 cell lists it.
    if (unit === 'SF' &&
        (category === 'Bike Lane' || category === 'Bus Lane' ||
         category === 'Pedestrian Space' || category === 'Bike Lane Green Bar')) {
      sfSum += qty;
      const color = String(r[10] || '').trim();
      if (color && color.toLowerCase() !== 'n/a') colorsSet[color] = true;
      return;
    }

    // HVX Crosswalk + Stop Line share the same row on the production
    // log but are tracked SEPARATELY so the cell can render as
    // "<crosswalk LF> / <stopline LF>" — admin can see both numbers
    // at a glance instead of an opaque sum.
    if (category === 'HVX Crosswalk') { crosswalkSum += qty; return; }
    if (category === 'Stop Line')     { stopLineSum  += qty; return; }

    // Pedestrian Men → the PL's currently-unused "PED X-ING Message" row,
    // written as "{n} PED MEN" so it's unmistakably PED MEN, not PED XING.
    if (category === 'Pedestrian Men') { pedMenSum += qty; return; }

    // Standard rename; unmapped categories drop silently
    const plLabel = PL_CATEGORY_MAP_[category];
    if (!plLabel) return;
    out.markings[plLabel] = (out.markings[plLabel] || 0) + qty;
  });

  if (crosswalkSum > 0 || stopLineSum > 0) {
    out.markings['CrossWalks/Stop Lines'] = `${crosswalkSum}/${stopLineSum}`;
  }
  if (pedMenSum > 0) {
    out.markings['PED X-ING Message'] = `${pedMenSum} PED MEN`;
  }
  if (sfSum > 0) {
    out.sqft  = sfSum;
    out.paint = Object.keys(colorsSet).sort().join(', ');
  }
  return out;
}


/**
 * Generate daily Production Log JSONs.
 *
 * - One JSON per contractor in CONFIG.PRODUCTION_LOG_CONTRACTORS that
 *   had ANY work-order activity on targetDate (Status doesn't matter
 *   — an in-progress multi-day WO appears on each day's log it was
 *   worked).
 * - Marking-item LF totals are filtered to items whose Date Completed
 *   equals targetDate, so cumulative WO totals don't leak into other
 *   days' logs.
 * - Returns an array of created Drive files (or empty array if no
 *   enabled contractor had work that day).
 */
function generateProductionLog_(targetDate, allEntries, byWorkOrder, woData, ss, opts) {
  const props = PropertiesService.getScriptProperties();
  const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
  const targetDayStr = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  let enabled = (CONFIG.PRODUCTION_LOG_CONTRACTORS || [])
    .map(s => String(s).trim()).filter(Boolean);
  if (opts && opts.contractorFilter) {
    const want = String(opts.contractorFilter).trim();
    enabled = enabled.filter(c => c === want);
    if (enabled.length === 0) {
      Logger.log('Production Log: contractor "' + want + '" not enabled — nothing to generate.');
      return [];
    }
  }
  if (enabled.length === 0) {
    Logger.log('No contractors enabled for production logs — skipping.');
    return [];
  }

  // ── Group every sign-in row by (contractor, crew_chief) ───────
  // Each DSID row carries its own crew chief at col 12 (post-multi-crew
  // schema). One WO can appear under MULTIPLE chiefs on the same date
  // when two crews handed off shifts on it — each chief gets its own
  // PL bucket showing only that crew's portion of the WO via the
  // chief-filtered Marking Items aggregation.
  //
  // Blank chief is a first-class value: legacy DSID rows (pre-migration)
  // group into a `(contractor, '')` bucket that produces today's
  // single-crew PL filename / doc_id / first-LP fallback chief.
  const byContractorChief = {};
  Object.entries(byWorkOrder).forEach(([woId, entries]) => {
    const woRow = woData.find(r => String(r[0]) === String(woId));
    if (!woRow) return;
    const contractor = String(woRow[1] || '').trim();
    if (!contractor) return;
    entries.forEach(row => {
      const chief = String(row[12] || '').trim();
      const key = contractor + '|||' + chief;
      if (!byContractorChief[key]) {
        byContractorChief[key] = { contractor, chief, wos: {} };
      }
      if (!byContractorChief[key].wos[woId]) {
        byContractorChief[key].wos[woId] = [];
      }
      byContractorChief[key].wos[woId].push(row);
    });
  });

  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  const subFolder    = getOrCreateSubfolder_(reviewFolder, 'Production Logs');
  const dateFormatted = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'MM/dd/yyyy');
  const logSheet = ss.getSheetByName('Automation Log');

  const generated = [];

  const buckets = Object.values(byContractorChief).filter(b => enabled.includes(b.contractor));
  if (buckets.length === 0) {
    Logger.log(`No enabled-contractor sign-ins on ${targetDayStr} — no Production Logs generated.`);
    return generated;
  }

  buckets.forEach(({ contractor, chief, wos: wosForBucket }) => {
    if (!wosForBucket || Object.keys(wosForBucket).length === 0) {
      return;
    }

    // Crew is derived from THIS (contractor, chief) bucket's rows only.
    const relevantEntries = [];
    Object.values(wosForBucket).forEach(rows => relevantEntries.push(...rows));

    const employees = {};
    relevantEntries.forEach(row => {
      const name = row[6];
      const timeInMins  = parseTimeToMinutes_(row[8]);
      const timeOutMins = parseTimeToMinutes_(row[9]);
      if (!employees[name]) {
        employees[name] = {
          timeIn: formatTime_(row[8]), timeOut: formatTime_(row[9]),
          timeInMins, timeOutMins,
          classification: row[7],
        };
      } else {
        if (timeInMins  < employees[name].timeInMins)  {
          employees[name].timeIn = formatTime_(row[8]);
          employees[name].timeInMins = timeInMins;
        }
        if (timeOutMins > employees[name].timeOutMins) {
          employees[name].timeOut = formatTime_(row[9]);
          employees[name].timeOutMins = timeOutMins;
        }
      }
    });

    // ── Per-WO payload ────────────────────────────────────────
    const workOrdersJson = Object.entries(wosForBucket).map(([woId, entries]) => {
      const woRow    = woData.find(r => String(r[0]) === String(woId));
      const borough  = woRow ? String(woRow[3]).toUpperCase() : '';
      const location = woRow ? String(woRow[5]).toUpperCase()
                             : String(entries[0][5] || '').toUpperCase();
      const status   = woRow ? String(woRow[15] || '').trim().toLowerCase() : '';

      // Marking-item LF totals filtered to items completed ON targetDate
      // AND tagged with this bucket's crew chief (so two crews handing
      // off a WO on the same day each see only their own portion).
      const agg = aggregateMarkingItemsForPL_(ss, woId, targetDayStr, chief);

      return {
        wo_number:    String(woId),
        borough:      borough,
        location:     location,
        sqft:         agg.sqft  !== '' ? String(agg.sqft)  : '',
        paint:        agg.paint !== '' ? String(agg.paint) : '',
        complete:     status === 'completed' ? 'Y' : 'N',
        layout_yn:    '',
        layout_hours: '',
        markings:     agg.markings,
      };
    });

    const sortedNames = Object.keys(employees);
    // Prefer the explicit crew chief from the bucket; fall back to the
    // "first LP" heuristic only when chief is blank (legacy data).
    const crewChiefName = chief && employees[chief]
      ? chief
      : (chief
          ? chief  // chief set but didn't sign in himself — still show as chief on form
          : (sortedNames.find(n => employees[n].classification === 'LP') || sortedNames[0]));
    const crewMemberNames = sortedNames.filter(n => n !== crewChiefName);
    const chiefTimes = employees[crewChiefName] || { timeIn: '', timeOut: '' };

    const logJson = {
      _type:             'production_log',
      contractor:        contractor,
      date:              dateFormatted,
      crew_number:       '',
      truck_number:      '',
      inspector_present: '',
      gas_tank_refilled: '',
      materials: {
        thermo_white_bags:  '',
        thermo_yellow_bags: '',
        beads_bags:         '',
        paint_cans:         '',
      },
      crew_chief: crewChiefName ? {
        name:     crewChiefName,
        time_in:  chiefTimes.timeIn,
        time_out: chiefTimes.timeOut,
      } : { name: '', time_in: '', time_out: '' },
      crew: crewMemberNames.map(n => ({
        name:     n,
        time_in:  employees[n].timeIn,
        time_out: employees[n].timeOut,
      })),
      work_orders: workOrdersJson,
    };

    // Filename includes a contractor slug + (when this bucket carries a
    // crew chief) a chief slug. The archive parser uses both to route
    // the file into the right contractor folder; the chief slug
    // disambiguates two crews from the same contractor working the
    // same source job on the same day. Blank chief → no slug → matches
    // today's legacy filename so existing parsers keep working.
    const slug = contractor.replace(/\s+/g, '_');
    const chiefSlug = chief ? '_chief-' + chief.replace(/[^A-Za-z0-9]/g, '') : '';
    const jsonFileName = `Production_Log_${targetDayStr}_${slug}${chiefSlug}.json`;
    const jsonFile = subFolder.createFile(
      jsonFileName, JSON.stringify(logJson, null, 2), MimeType.PLAIN_TEXT);
    generated.push(jsonFile);

    if (logSheet) {
      logSheet.appendRow([
        new Date(), 'Production Log Generator', 'Daily trigger',
        `${Object.keys(wosForBucket).length} WO(s) for ${contractor}${chief ? ' (' + chief + ')' : ''} on ${dateFormatted}`,
        jsonFileName, 'Generated',
        '', 'Yes — review in Approvals tab',
      ]);
    }
    Logger.log('✅ Production log JSON exported: ' + jsonFileName);
  });

  if (generated.length === 0) {
    Logger.log(`No enabled contractors had WOs on ${targetDayStr} — no Production Log generated.`);
  }
  return generated;
}


/**
 * DEBUG / one-off: run from the Apps Script editor to generate a
 * Production Log for any target date (default = today). Useful for
 * verifying the Marking Items aggregation + the completion filter
 * without waiting for the daily trigger. Logs the grouped WO list
 * and the count of items in each WO column.
 */
function debugGenerateProductionLogForToday() {
  const DATE_OVERRIDE = '';   // ← optional 'YYYY-MM-DD'; blank = today

  const targetDate = DATE_OVERRIDE ? new Date(DATE_OVERRIDE + 'T12:00:00') : new Date();
  const ss         = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const signIn     = ss.getSheetByName('Daily Sign-In Data').getDataRange().getValues();
  const woData     = ss.getSheetByName('Work Order Tracker').getDataRange().getValues();

  const todaysEntries = signIn.slice(1).filter(row => {
    if (!row[0]) return false;
    return new Date(row[0]).toDateString() === targetDate.toDateString();
  });
  Logger.log(`Target date: ${targetDate.toDateString()} — ${todaysEntries.length} sign-in row(s).`);

  const byWorkOrder = {};
  todaysEntries.forEach(row => {
    _splitWOIds_(row[1]).forEach(woId => {
      if (!byWorkOrder[woId]) byWorkOrder[woId] = [];
      byWorkOrder[woId].push(row);
    });
  });
  Logger.log(`WOs with sign-in activity: ${Object.keys(byWorkOrder).join(', ') || '(none)'}`);

  generateProductionLog_(targetDate, todaysEntries, byWorkOrder, woData, ss);
}

/**
 * Generate Contractor Field Report for a completed Work Order
 */
function generateFieldReport_(woId, entries, woData, ss) {
  const props = PropertiesService.getScriptProperties();
  const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
  
  const woRow = woData.find(r => r[0] === woId);
  if (!woRow) { Logger.log('❌ WO not found in tracker: ' + woId); return; }
  
  const contractor = woRow[1];
  const contractNum = woRow[2];
  const borough = woRow[3];
  const location = woRow[5];
  const fromSt = woRow[6];
  const toSt = woRow[7];
  const sqft = entries[0][12];
  const paint = entries[0][13];
  const workDate = entries[0][0];
  const issues = entries.map(e => e[15]).filter(Boolean).join('; ');
  
  // Find crew chief
  const crewChief = entries.find(e => e[7] === 'LP');
  const crewChiefName = crewChief ? crewChief[6] : 'N/A';
  
  const dateFormatted = workDate instanceof Date 
    ? Utilities.formatDate(workDate, CONFIG.TIMEZONE, 'MM/dd/yyyy')
    : String(workDate);
  
  let report = `CONTRACTOR FIELD REPORT\n`;
  report += `${'═'.repeat(60)}\n\n`;
  report += `Work Order:      ${woId}\n`;
  report += `Contractor:      ${contractor}\n`;
  report += `Contract #:      ${contractNum}\n`;
  report += `Borough:         ${borough}\n`;
  report += `Location:        ${location}\n`;
  report += `From:            ${fromSt}\n`;
  report += `To:              ${toSt}\n`;
  report += `\nInstallation Date: ${dateFormatted}\n`;
  report += `\nPAVEMENT MARKINGS:\n`;
  report += `${'─'.repeat(40)}\n`;
  report += `Color Surface Treatment:  ${sqft} SQFT\n`;
  report += `Paint / Material:         ${paint}\n`;
  report += `\nGeneral Remarks: ${issues || 'None'}\n`;
  report += `\nCrew Chief: ${crewChiefName}\n`;
  report += `Contractor Notes: ONEIRO / WBSE\n`;
  report += `\n${'═'.repeat(60)}\n`;
  report += `⚠️  NEEDS REVIEW — Verify all fields before sending to ${contractor}\n`;
  report += `    Contractor signature required on final version\n`;
  
  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  const subFolder = getOrCreateSubfolder_(reviewFolder, 'Field Reports');
  const fileName = `Field_Report_${woId}_${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd')}.txt`;
  subFolder.createFile(fileName, report, MimeType.PLAIN_TEXT);
  
  const logSheet = ss.getSheetByName('Automation Log');
  logSheet.appendRow([
    new Date(), 'Field Report Generator', 'WO Completed',
    woId, fileName, 'Generated',
    '', 'Yes — Review, sign, and send to ' + contractor
  ]);
  
  Logger.log('✅ Field report generated: ' + fileName);
}

/**
 * Generate Invoice for a completed Work Order
 */
function generateInvoice_(woId, entries, woData, ss) {
  const props = PropertiesService.getScriptProperties();
  const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
  
  const woRow = woData.find(r => r[0] === woId);
  if (!woRow) return;
  
  const contractor = woRow[1];
  // Apply billing remap so sub-prime work on a contract they didn't
  // win is billed under their own contract. Source WO Tracker stays
  // raw — only the invoice content + AR row + rate lookup shift.
  const _invMapped = _billingRemap_(woRow[2], woRow[3], contractor);
  const contractNum = _invMapped.contractNum;
  const borough = _invMapped.borough;
  const location = woRow[5];
  const sqft = Number(entries[0][12]) || 0;

  // Look up rate from Contract Lookup
  const clSheet = ss.getSheetByName('Contract Lookup');
  const clData = clSheet.getDataRange().getValues();
  const contractRow = clData.find(r =>
    String(r[0]).includes(String(contractNum).split('/')[0]) &&
    String(r[1]) === String(borough)
  );
  const rate = contractRow ? Number(contractRow[6]) : 3.40; // Default rate
  const amount = sqft * rate;
  
  // Get next invoice number from AR sheet
  const arSheet = ss.getSheetByName('Invoices & AR');
  const arData = arSheet.getDataRange().getValues();
  const lastInvoice = arData.slice(1).reduce((max, row) => {
    const num = Number(row[0]);
    return num > max ? num : max;
  }, 1100);
  const nextInvoiceNum = lastInvoice + 1;
  
  // Get contractor billing info from Contacts sheet
  const ccSheet = ss.getSheetByName('Contractor Contacts');
  const ccData = ccSheet.getDataRange().getValues();
  const billingContact = ccData.find(r => r[0] === contractor && String(r[8]).toLowerCase() === 'yes');
  const billTo = billingContact ? billingContact[1] : contractor;
  const billAddress = billingContact ? billingContact[5] : '';
  
  const invoiceDate = new Date();
  const dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + 30);
  
  const workDate = entries[0][0] instanceof Date 
    ? Utilities.formatDate(entries[0][0], CONFIG.TIMEZONE, 'MM/dd/yy')
    : String(entries[0][0]);
  
  let invoice = `
                                          Oneiro Collection LLC
                                          435 South Ave apt 220
                                          Garwood, NJ 07027 US
                                          +1 9176209809
                                          anthe@theoneiro.com

INVOICE

BILL TO                                   INVOICE    ${nextInvoiceNum}
${billTo}                                 DATE       ${Utilities.formatDate(invoiceDate, CONFIG.TIMEZONE, 'MM/dd/yyyy')}
${contractor}                             TERMS      Net 30
${billAddress}                            DUE DATE   ${Utilities.formatDate(dueDate, CONFIG.TIMEZONE, 'MM/dd/yyyy')}

${'─'.repeat(72)}
DATE        DESCRIPTION                                  QTY      RATE    AMOUNT
${'─'.repeat(72)}
            ${woId}    Contract ${contractNum}/${getBoroughName_(borough)}
            ${location} - Completed ${workDate}
                                                      ${sqft.toLocaleString().padStart(7)}    ${rate.toFixed(2)}  ${amount.toLocaleString('en-US', {style: 'currency', currency: 'USD'})}
${'─'.repeat(72)}

Contact Oneiro Collection LLC to pay.     BALANCE DUE           ${amount.toLocaleString('en-US', {style: 'currency', currency: 'USD'})}
`;

  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  const subFolder = getOrCreateSubfolder_(reviewFolder, 'Invoices');
  const fileName = `Invoice_${nextInvoiceNum}_${woId}.txt`;
  subFolder.createFile(fileName, invoice, MimeType.PLAIN_TEXT);
  
  // Also add to the AR tracker
  arSheet.appendRow([
    nextInvoiceNum, invoiceDate, dueDate,
    contractor, contractNum, borough,
    woId, `${location} - Completed ${workDate}`,
    sqft, rate, amount,
    'Draft', '', '',
    '', 'Net 30 — Auto-generated'
  ]);
  
  const logSheet = ss.getSheetByName('Automation Log');
  logSheet.appendRow([
    new Date(), 'Invoice Generator', 'WO Completed',
    woId, `Invoice #${nextInvoiceNum} — $${amount.toFixed(2)}`,
    'Generated', '', 'Yes — Review and send to ' + contractor
  ]);
  
  Logger.log(`✅ Invoice #${nextInvoiceNum} generated: $${amount.toFixed(2)} for ${woId}`);
}


// ═══════════════════════════════════════════════════════════════
// 4. REVIEW → APPROVE → SEND → FILE WORKFLOW
// ═══════════════════════════════════════════════════════════════

/**
 * Processes documents that the admin has moved to the Approved Docs folder.
 * Automatically emails them to the correct contractor contact and files them.
 *
 * Wrapped in a script-wide lock so a manual run from the editor can't
 * race the 10-min cron. Without the lock, two concurrent invocations
 * each call file.makeCopy on the same approved doc before either one
 * trashes it — producing duplicate archive copies in every WO folder.
 */
function processApprovedDocuments() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('ℹ️ processApprovedDocuments: another invocation holds the lock — skipping this run.');
    return { archived: 0, errored: 0, skipped: true };
  }
  try {
    return _processApprovedDocumentsImpl_();
  } finally {
    try { lock.releaseLock(); } catch (e) { /* never held */ }
  }
}

function _processApprovedDocumentsImpl_() {
  const props = PropertiesService.getScriptProperties();
  const approvedId = props.getProperty('APPROVED_SENT_ID');
  if (!approvedId) return { archived: 0, errored: 0 };

  const approvedFolder = DriveApp.getFolderById(approvedId);
  const files = approvedFolder.getFiles();

  // Cheap pre-flight: if the folder is empty (or every file is already
  // 📨-prefixed), bail before we open the spreadsheet or run any of the
  // heavier setup. The Python worker pokes this every ~15s; the vast
  // majority of those calls have no work to do, and skipping the
  // SpreadsheetApp.openById + Logger noise keeps the GAS execution log
  // readable when something does go wrong.
  let hasWork = false;
  const queued = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().startsWith('📨')) continue;
    hasWork = true;
    queued.push(f);
  }
  if (!hasWork) return { archived: 0, errored: 0 };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Counts returned to the caller (Tools menu → toast feedback).
  let archivedCount = 0;
  let erroredCount  = 0;

  for (let qi = 0; qi < queued.length; qi++) {
    const file = queued[qi];
    const fileName = file.getName();

    // Determine document type and extract WO info from filename.
    // WO # covers the three prefixes the parser supports (PT/RM/PM).
    const WO_REGEX = /(PT|RM|PM)[-_]?\d+/i;
    let docType = 'Unknown';
    let woId = '';

    if (fileName.includes('Production_Log')) {
      docType = 'Production Log';
    } else if (
      // Two filename conventions are in play:
      //   "Contractor_Field_Report_<WO>_<date>_FILLED.pdf" (legacy)
      //   "CFR_<WO>_<date>_FILLED.pdf"                     (current — the
      //     Python worker now reuses the source JSON's stem, and Apps
      //     Script names the CFR JSON `CFR_<WO>_<date>.json`).
      fileName.includes('Field_Report') || fileName.startsWith('CFR_')
    ) {
      docType = 'Field Report';
      const match = fileName.match(WO_REGEX);
      if (match) woId = match[0];
    } else if (fileName.includes('Invoice')) {
      docType = 'Invoice';
      const match = fileName.match(WO_REGEX);
      if (match) woId = match[0];
    } else if (fileName.includes('Certified_Payroll')) {
      docType = 'Certified Payroll';
    } else if (fileName.includes('SignIn') || fileName.includes('Sign_In') || fileName.includes('Sign-In')) {
      docType = 'Sign-In';
      // New multi-WO pattern: SignIn_<contractNum>_<borough>_<YYYY-MM-DD>[_MANUAL].pdf.
      // The contract number (e.g. "PT-11930") matches WO_REGEX too, so we
      // explicitly detect the multi-WO shape and skip the regex extraction
      // there — archiveDocument_ falls back to contract+date lookup when
      // woId is empty.
      const isMultiWOPattern = /^SignIn_[^_]+_[^_]+_\d{4}-\d{2}-\d{2}/.test(fileName);
      if (!isMultiWOPattern) {
        const match = fileName.match(WO_REGEX);
        if (match) woId = match[0];
      }
    }
    
    // Per-doc emailing has been retired in favor of the batch
    // Download Documents flow on the dashboard. The cron is now a
    // pure archival pipeline. getRecipientsForDoc_ stays in the file
    // for future automation iterations but is no longer called here.
    //
    // Archive the file.  archiveDocument_ never throws — it returns a
    // result so we can branch on failure and preserve the file rather
    // than trashing it into the void.
    const archiveResult = archiveDocument_(file, docType, woId, ss);

    const logSheet = ss.getSheetByName('Automation Log');

    if (archiveResult.success) {
      const archiveNote = docType === 'Production Log' || docType === 'Certified Payroll'
        ? `Archived to master folder + duplicated into WO subfolder(s)`
        : woId
          ? `Archived to WO folder: ${woId}`
          : `Archived (doc type: ${docType})`;

      // Flip the matching Done? flag(s) on the WO Tracker. archiveDocument_
      // returns the doc_type and the list of WOs the archive operation
      // covered (multiple for Production Log / Sign-In / CP, one for
      // CFR / Invoice). The helper is a no-op for unknown doc types.
      _setDocLifecycleFlags_(ss, archiveResult.wo_ids || [], archiveResult.doc_type, { done: true });

      logSheet.appendRow([
        new Date(), 'Approve & Send', 'File moved to Approved Docs folder',
        fileName,
        `${archiveNote} | Done flag set on ${(archiveResult.wo_ids || []).length} WO(s) — Sent flag is set on batch download`,
        'Completed', '', 'No'
      ]);

      // Delete from Approved Docs — archive is now the single source of truth
      file.setTrashed(true);
      Logger.log(`🗑️ Deleted from Approved Docs: ${fileName}`);
      archivedCount += 1;
    } else {
      // Archive failed — preserve the file in ⚠️ Archive Errors so admin
      // can diagnose + re-file manually, rather than silently trashing it.
      const errorsFolder = getOrCreateArchiveErrorsFolder_();
      if (errorsFolder) {
        try {
          file.moveTo(errorsFolder);
          Logger.log(`⚠️ Archive failed for ${fileName} — moved to ⚠️ Archive Errors (reason: ${archiveResult.reason})`);
        } catch (moveErr) {
          // Last-resort: couldn't even move it. Leave it where it is so the
          // next tick retries; make the failure loud in the log.
          Logger.log(`❌ Could not move ${fileName} to Archive Errors: ${moveErr && moveErr.stack || moveErr}`);
        }
      } else {
        Logger.log(`❌ Archive failed for ${fileName} AND Archive Errors folder unavailable — leaving in Approved Docs for retry. Reason: ${archiveResult.reason}`);
      }

      logSheet.appendRow([
        new Date(), 'Approve & Send', 'Archive failed — moved to Archive Errors',
        fileName,
        `Reason: ${archiveResult.reason}`,
        'Error', '', 'Yes'
      ]);
      erroredCount += 1;
    }
  }

  // Archive flips Done flags on the WO Tracker (the dashboard surfaces
  // these). Invalidate the dashboard cache so the next /api/dashboard
  // rebuilds and reflects the new flag state instead of serving up to
  // 60s of stale data.
  if (archivedCount > 0) {
    _invalidateCacheKeys_(['dashboard_v1']);
  }

  return { archived: archivedCount, errored: erroredCount };
}

function getRecipientsForDoc_(docType, woId, ss) {
  const ccSheet = ss.getSheetByName('Contractor Contacts');
  const ccData = ccSheet.getDataRange().getValues();
  const headers = ccData[0];
  const recipients = [];
  
  // Map doc type to the "Receives X?" column
  const docTypeColMap = {
    'Production Log': 6,    // Column G
    'Field Report': 7,      // Column H
    'Invoice': 8,           // Column I
    'Certified Payroll': 9, // Column J
  };
  
  const colIdx = docTypeColMap[docType];
  if (!colIdx) return recipients;
  
  // If we have a WO ID, look up the contractor
  let contractor = '';
  if (woId) {
    const woSheet = ss.getSheetByName('Work Order Tracker');
    const woData = woSheet.getDataRange().getValues();
    const woRow = woData.find(r => r[0] === woId);
    if (woRow) contractor = woRow[1];
  }
  
  ccData.slice(1).forEach(row => {
    if (String(row[colIdx]).toLowerCase() === 'yes') {
      if (!contractor || row[0] === contractor) {
        if (row[3]) { // Has email
          recipients.push({ name: row[1], email: row[3] });
        }
      }
    }
  });
  
  return recipients;
}

// ── Doc lifecycle column maps ─────────────────────────────────
// Single source of truth for which WO Tracker column tracks each
// doc type's Done/Sent state. Used by _setDocLifecycleFlags_ +
// list_documents_for_batch + set_docs_sent.
//
// Doc-type strings match the values archiveDocument_ returns.
// 0-idx column numbers (sheet col = idx+1).
// Per-WO storage covers only CFR + Invoice now. PL/SI/CP moved to the
// Doc Lifecycle Log (per-doc storage) — those entries were removed
// from the WO Tracker by migrateRemoveObsoleteDocColumns.
const DOC_TYPE_DONE_COL_ = Object.freeze({
  'Field Report': 25,   // legacy — Field Report Done?
  'Invoice':      44,   // Invoice Done? — shifted left after the obsolete-col trim
});
const DOC_TYPE_SENT_COL_ = Object.freeze({
  'Field Report': 43,   // CFR Sent?
  'Invoice':      29,   // legacy — Invoice Sent?
});

/**
 * Bulk-flip Done? and/or Sent? flags on the WO Tracker for a list of
 * WO ids and one doc type. `flags` is `{done?: boolean, sent?: boolean}`;
 * pass true to write 'Yes', false to write 'No', omit to leave alone.
 *
 * Idempotent. Doesn't read the row first — just writes. Safe to call
 * with an empty woIds array (no-op).
 *
 * Single sheet read of all rows up front so we can map woId → row
 * number once without N round-trips.
 */
function _setDocLifecycleFlags_(ss, woIds, docType, flags) {
  if (!Array.isArray(woIds) || woIds.length === 0) return;
  if (!flags || (flags.done == null && flags.sent == null)) return;

  const doneCol = DOC_TYPE_DONE_COL_[docType];
  const sentCol = DOC_TYPE_SENT_COL_[docType];
  if (doneCol == null && sentCol == null) {
    Logger.log('⚠️ _setDocLifecycleFlags_: unknown doc_type "' + docType + '" — skipping');
    return;
  }

  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) return;
  ensureWoTrackerExtraCols_(woSheet);

  const data = woSheet.getDataRange().getValues();
  const rowByWoId = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (id) rowByWoId[id] = i + 1;   // 1-indexed sheet row
  }

  const doneVal = flags.done == null ? null : (flags.done ? 'Yes' : 'No');
  const sentVal = flags.sent == null ? null : (flags.sent ? 'Yes' : 'No');

  woIds.forEach(woId => {
    const row = rowByWoId[String(woId).trim()];
    if (!row) return;
    if (doneVal != null && doneCol != null) {
      woSheet.getRange(row, doneCol + 1).setValue(doneVal);
    }
    if (sentVal != null && sentCol != null) {
      woSheet.getRange(row, sentCol + 1).setValue(sentVal);
    }
  });
}

// ── action: set_docs_sent ─────────────────────────────────────
//
// Bulk-flip Done? and/or Sent? flags on the WO Tracker. Used by the
// batch download endpoint after a successful zip stream, and by the
// dashboard's manual toggle UI.
//
// Body shape:
//   { updates: [{ wo_id, doc_type, sent?: bool, done?: bool }, ...] }
//
// Groups updates by (doc_type, flag combination) to minimize sheet
// writes. Returns counts per flag for caller telemetry.
function handleSetDocsSent_(body) {
  const d = body.data || {};
  const updates = Array.isArray(d.updates) ? d.updates : [];
  if (updates.length === 0) return jsonResponse_({ updated: 0 });

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // doc_type values from the listing handler / webapp use friendly names
  // (CFR, Production Log, Sign-In, Certified Payroll, Invoice). The
  // column maps DOC_TYPE_DONE_COL_ / DOC_TYPE_SENT_COL_ key on
  // archiveDocument_'s internal names — only CFR ≠ Field Report.
  // Without this translation, set_docs_sent silently skipped every
  // CFR row (the helper logs "unknown doc_type" but the caller doesn't
  // see it), so the batch download's mark-sent step did nothing for
  // CFRs. Accept either friendly or internal here so manual flips
  // (Phase 5) can use whichever is convenient.
  const toInternal = {
    'CFR':               'Field Report',
    'Field Report':      'Field Report',
    'Production Log':    'Production Log',
    'Sign-In':           'Sign-In',
    'Certified Payroll': 'Certified Payroll',
    'Invoice':           'Invoice',
  };

  // PL / SI / CP storage moved to the Doc Lifecycle Log (per-doc rows
  // keyed by Doc ID). Per-WO updates for those types no longer make
  // sense — caller should route to set_doc_status with a doc_id. Skip
  // them here with a warning so a stale caller doesn't silently flip
  // unrelated rows.
  const PER_DOC_TYPES = { 'Production Log': 1, 'Sign-In': 1, 'Certified Payroll': 1 };

  // Group woIds by (doc_type, done?, sent?) so each unique flag combo
  // is one batch write call. Most calls will collapse to one bucket.
  const buckets = {};
  updates.forEach(u => {
    const raw = String(u.doc_type || '').trim();
    const docType = toInternal[raw];
    if (!docType) {
      Logger.log('⚠️ handleSetDocsSent_: skipping unknown doc_type "' + raw + '"');
      return;
    }
    if (PER_DOC_TYPES[docType]) {
      Logger.log('⚠️ handleSetDocsSent_: per-WO update for "' + docType + '" ignored — use set_doc_status with doc_id instead');
      return;
    }
    const woId = String(u.wo_id || '').trim();
    if (!woId) return;
    const doneKey = (u.done === true) ? 'Y' : (u.done === false ? 'N' : '_');
    const sentKey = (u.sent === true) ? 'Y' : (u.sent === false ? 'N' : '_');
    if (doneKey === '_' && sentKey === '_') return;
    const k = docType + '|' + doneKey + '|' + sentKey;
    if (!buckets[k]) buckets[k] = { docType, flags: {}, woIds: [] };
    if (doneKey !== '_') buckets[k].flags.done = (doneKey === 'Y');
    if (sentKey !== '_') buckets[k].flags.sent = (sentKey === 'Y');
    buckets[k].woIds.push(woId);
  });

  let total = 0;
  Object.values(buckets).forEach(b => {
    _setDocLifecycleFlags_(ss, b.woIds, b.docType, b.flags);
    total += b.woIds.length;
  });

  if (total > 0) _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ updated: total });
}

// ── action: set_doc_status ────────────────────────────────────
//
// Per-doc lifecycle update for time-anchored docs (PL / SI / CP) that
// live in the Doc Lifecycle Log. Body shape:
//
//   { updates: [{ doc_id, done?, sent? }, ...] }
//
// Each update target is a single row in the Log (1-to-1 with Doc ID).
// Returns count of rows touched.
function handleSetDocStatus_(body) {
  const d = body.data || {};
  const updates = Array.isArray(d.updates) ? d.updates : [];
  if (updates.length === 0) return jsonResponse_({ updated: 0 });

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let total = 0;
  updates.forEach(u => {
    const docId = String(u.doc_id || '').trim();
    if (!docId) return;
    const flags = {};
    if (u.done === true)  flags.done = true;
    if (u.done === false) flags.done = false;
    if (u.sent === true)  flags.sent = true;
    if (u.sent === false) flags.sent = false;
    if (flags.done == null && flags.sent == null) return;
    _setDocLifecycleStatus_(ss, docId, flags);
    total++;
  });

  if (total > 0) _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ updated: total });
}

/**
 * Archive a document using the correct folder structure:
 *
 *   Archive / [Contractor] / [ContractNum - Borough] /
 *     ├── PT-XXXXX - [Location] /   ← WO folder (Field Reports, Invoices, Sign-Ins filed directly here)
 *     │     └── Photos/             ← only subfolder inside a WO
 *     ├── Production Logs/          ← master copy; also duplicated into each WO folder
 *     ├── Sign-Ins/                 ← master copy; also duplicated into each WO folder
 *     └── Certified Payroll/        ← master copy; also duplicated into each WO folder
 *
 * Returns one of:
 *   - On success: { success: true,  doc_type: <one of "Field Report" |
 *                   "Invoice" | "Production Log" | "Sign-In" |
 *                   "Certified Payroll">,
 *                   wo_ids: [<id>, ...] }
 *     `wo_ids` lists every Work Order this archive operation touched —
 *     one element for per-WO docs (CFR/Invoice), many for the
 *     master-copy types (Production Log/Sign-In/CP). The caller in
 *     _processApprovedDocumentsImpl_ uses these to flip the matching
 *     "<DocType> Done?" flag on each WO Tracker row.
 *   - On failure: { success: false, reason: <string> }
 *     The caller is responsible for preserving the file (moveTo
 *     Archive Errors) rather than trashing it, so nothing is ever
 *     lost to a bad filename, missing tracker row, or transient
 *     Drive error.
 */
function archiveDocument_(file, docType, woId, ss) {
  try {
    const props = PropertiesService.getScriptProperties();
    const archiveId = props.getProperty('ARCHIVE_ID');
    if (!archiveId) {
      return { success: false, reason: 'ARCHIVE_ID not set' };
    }

    const archiveRoot = DriveApp.getFolderById(archiveId);
    const cleanName = file.getName().replace('📨 ', '');

    if (docType === 'Field Report') {
      // Field Report PDFs from the worker now embed the FULL WO document
      // (original scan pages + the freshly-rendered CFR page). They
      // REPLACE the prior WO doc in the archive folder — atomic swap so
      // there's never a window where the WO is missing.
      //
      // Exception: filenames containing `_Updated_<yyyy-mm-dd>_` come
      // from the admin "save as new CFR" path (punch-order rework).
      // Those archive ALONGSIDE the canonical WO doc — the original
      // record stays put. Both files in the folder are full WO+CFR
      // merges; only the CFR-page content differs.
      if (!woId) {
        return { success: false, reason: 'CFR has no WO # in filename' };
      }
      const woFolder = getWOFolder_(archiveRoot, woId, ss);
      if (!woFolder) {
        return { success: false, reason: `Could not resolve WO folder for CFR ${woId} — is it missing from the Work Order Tracker?` };
      }
      const SAVE_AS_NEW_RX = /_Updated_\d{4}-\d{2}-\d{2}_FILLED\.pdf$/i;
      if (SAVE_AS_NEW_RX.test(cleanName)) {
        file.makeCopy(cleanName, woFolder);
        Logger.log(`📁 Archived save-as-new CFR → WO folder ${woId} (preserved canonical WO doc)`);
      } else {
        _replaceArchivedWODoc_(woFolder, file, woId);
        Logger.log(`📁 Archived merged CFR → WO folder ${woId} (replaced prior WO doc)`);
      }
      return { success: true, doc_type: 'Field Report', wo_ids: [woId] };

    } else if (docType === 'Invoice') {
      // Invoice files alongside the WO doc — no replacement.
      if (!woId) {
        return { success: false, reason: `Invoice has no WO # in filename` };
      }
      const woFolder = getWOFolder_(archiveRoot, woId, ss);
      if (!woFolder) {
        return { success: false, reason: `Could not resolve WO folder for Invoice ${woId} — is it missing from the Work Order Tracker?` };
      }
      file.makeCopy(cleanName, woFolder);
      Logger.log(`📁 Archived Invoice: ${cleanName} → WO folder ${woId}`);
      return { success: true, doc_type: 'Invoice', wo_ids: [woId] };

    } else if (docType === 'Sign-In') {
      // New multi-WO sign-ins land as SignIn_<contractNum>_<borough>_<YYYY-MM-DD>[_<ContractorSlug>][_chief-<ChiefSlug>][_MANUAL].pdf.
      // We look up every WO that worked that contract on that day and copy
      // the same PDF into each WO folder (same pattern as Production Logs).
      // Optional <ContractorSlug> appears when a billing remap split the
      // file by sub-prime; optional `_chief-<ChiefSlug>` appears when
      // multiple crews from the same prime worked the same source job.
      // Both are used to route the file to the right lifecycle row.
      const newPat = cleanName.match(
        /^SignIn_([^_]+)_([^_]+)_(\d{4}-\d{2}-\d{2})(?:_(?!chief-)([A-Za-z0-9]+))?(?:_chief-([A-Za-z0-9]+))?/
      );
      if (newPat) {
        const [, contractNum, borough, dateStr, fileContractorSlug, fileChiefSlug] = newPat;
        const slugMatches = (co) => String(co || '').replace(/[^A-Za-z0-9]/g, '') === fileContractorSlug;
        const isSuffixMarker = fileContractorSlug === 'MANUAL' || fileContractorSlug === 'FILLED';
        const useFileContractor = fileContractorSlug && !isSuffixMarker;
        const crewChief = fileChiefSlug || '';
        const logDate = new Date(dateStr + 'T12:00:00');
        const wosByContract = getWOsForDate_(logDate, ss);
        const matchKey = Object.keys(wosByContract).find(k => {
          const parts = k.split('|');
          if (parts[1] !== contractNum || parts[2] !== borough) return false;
          if (useFileContractor && !slugMatches(parts[0])) return false;
          return true;
        });
        if (!matchKey) {
          return {
            success: false,
            reason: `No WOs found in Daily Sign-In Data for ${contractNum}/${borough}${useFileContractor ? ' (' + fileContractorSlug + ')' : ''} on ${dateStr} — sign-in submitted before any matching Work Day Log row?`
          };
        }
        const [contractor] = matchKey.split('|');
        const wos = wosByContract[matchKey];
        const contractFolder = getOrCreateSubfolder_(
          getOrCreateSubfolder_(archiveRoot, contractor),
          `${contractNum} - ${getBoroughName_(borough)}`
        );
        // Master copy at contract level — mirrors Production Logs and
        // Certified Payroll. Lets the listing handler read Sign-Ins
        // from a single place instead of de-duping across WO folders.
        const masterCopy = file.makeCopy(cleanName, getOrCreateSubfolder_(contractFolder, 'Sign-Ins'));
        wos.forEach(wo => {
          const woFolder = getOrCreateSubfolder_(contractFolder, `${wo.id} - ${wo.location}`);
          file.makeCopy(cleanName, woFolder);
        });
        Logger.log(`📁 Archived Sign-In → ${contractor}/${contractNum}${crewChief ? ' (' + crewChief + ')' : ''} + master + ${wos.length} WO folder(s)`);

        // Doc Lifecycle Log: upsert SI row with done=true. Crew chief
        // (from filename slug) routes to the correct pending row when
        // multiple crews share the same source job. Blank chief matches
        // legacy single-crew rows.
        try {
          _upsertDocLifecycleRow_(ss, {
            doc_id:       _docLifecycleId_('Sign-In', dateStr, contractNum, borough, crewChief),
            doc_type:     'Sign-In',
            anchor:       dateStr,
            contractor:   contractor,
            contract_num: contractNum,
            borough:      borough,
            crew_chief:   crewChief,
            wo_ids:       wos.map(w => w.id),
            done:         true,
            file_id:      masterCopy.getId(),
          });
        } catch (e) {
          Logger.log('⚠️ archiveDocument_ Sign-In: Doc Lifecycle Log upsert failed: ' + e.message);
        }

        return { success: true, doc_type: 'Sign-In', wo_ids: wos.map(w => w.id) };
      }
      // Legacy single-WO Sign-In (pre-multi-WO refactor) — keep working.
      if (!woId) {
        return { success: false, reason: `Sign-In has no contract/WO info in filename` };
      }
      const woFolder = getWOFolder_(archiveRoot, woId, ss);
      if (!woFolder) {
        return { success: false, reason: `Could not resolve WO folder for Sign-In ${woId} — is it missing from the Work Order Tracker?` };
      }
      file.makeCopy(cleanName, woFolder);
      Logger.log(`📁 Archived Sign-In (legacy): ${cleanName} → WO folder ${woId}`);
      return { success: true, doc_type: 'Sign-In', wo_ids: [woId] };

    } else if (docType === 'Production Log') {
      // Filename patterns:
      //   New (per-contractor):       Production_Log_<YYYY-MM-DD>_<Contractor_Slug>_FILLED.pdf
      //   Multi-crew (per-chief):     Production_Log_<YYYY-MM-DD>_<Contractor_Slug>_chief-<ChiefSlug>_FILLED.pdf
      //   Legacy (single combined):   Production_Log_<YYYY-MM-DD>_FILLED.pdf
      const dateMatch = cleanName.match(/(\d{4}-\d{2}-\d{2})/);
      if (!dateMatch) {
        return { success: false, reason: 'Could not parse date from Production Log filename' };
      }
      const logDate = new Date(dateMatch[1] + 'T12:00:00');

      // Pull contractor slug + optional chief slug between the date and
      // "_FILLED" (or end). Empty contractor = legacy filename → file
      // under every contractor that had work that day. Empty chief =
      // single-crew (legacy) behavior.
      const slugs = (() => {
        const m = cleanName.match(
          /Production_Log_\d{4}-\d{2}-\d{2}_(.+?)(?:_chief-([A-Za-z0-9]+))?(?:_FILLED)?\.pdf$/);
        return m
          ? { contractor: m[1] || '', chief: m[2] || '' }
          : { contractor: '', chief: '' };
      })();
      const contractorTarget = slugs.contractor.replace(/_/g, ' ').trim();
      const crewChief = slugs.chief;

      // Group WOs worked that day by contractor/contract/borough
      const wosByContract = getWOsForDate_(logDate, ss);
      if (Object.keys(wosByContract).length === 0) {
        return { success: false, reason: `No WOs found in Daily Sign-In Data for ${dateMatch[1]} — nothing to archive the Production Log against` };
      }

      const matchingEntries = contractorTarget
        ? Object.entries(wosByContract).filter(([key]) =>
            String(key.split('|')[0] || '').trim() === contractorTarget)
        : Object.entries(wosByContract);

      if (matchingEntries.length === 0) {
        return { success: false,
          reason: `No WOs in Daily Sign-In Data for contractor "${contractorTarget}" on ${dateMatch[1]}` };
      }

      const allCoveredWoIds = [];
      const plAnchorIso = dateMatch[1];
      let firstMasterFileId = '';
      let plContractor = '';
      matchingEntries.forEach(([key, wos]) => {
        const [contractor, contractNum, borough] = key.split('|');
        if (!plContractor) plContractor = contractor;
        const contractFolder = getOrCreateSubfolder_(
          getOrCreateSubfolder_(archiveRoot, contractor),
          `${contractNum} - ${getBoroughName_(borough)}`
        );
        // Master copy at contract level — every (contract, borough) gets
        // its own master so downstream contract-level browsing still
        // finds the PL.
        const masterCopy = file.makeCopy(cleanName, getOrCreateSubfolder_(contractFolder, 'Production Logs'));
        if (!firstMasterFileId) firstMasterFileId = masterCopy.getId();
        wos.forEach(wo => {
          const woFolder = getOrCreateSubfolder_(contractFolder, `${wo.id} - ${wo.location}`);
          file.makeCopy(cleanName, woFolder);
          allCoveredWoIds.push(wo.id);
        });
        Logger.log(`📁 Archived Production Log → ${contractor}/${contractNum}${crewChief ? ' (' + crewChief + ')' : ''} + ${wos.length} WO folder(s)`);
      });

      // Doc Lifecycle Log: ONE row per (date, contractor, crew_chief).
      // Multi-crew shifts give each chief their own pending row, and
      // each PL file flips its row Done independently.
      try {
        _upsertDocLifecycleRow_(ss, {
          doc_id:     _plDocId_(plAnchorIso, plContractor, crewChief),
          doc_type:   'Production Log',
          anchor:     plAnchorIso,
          contractor: plContractor,
          crew_chief: crewChief,
          // contract_num + borough deliberately blank — PL spans contracts.
          wo_ids:     allCoveredWoIds,
          done:       true,
          file_id:    firstMasterFileId,
        });
      } catch (e) {
        Logger.log('⚠️ archiveDocument_ Prod Log: Doc Lifecycle Log upsert failed: ' + e.message);
      }
      return { success: true, doc_type: 'Production Log', wo_ids: allCoveredWoIds };

    } else if (docType === 'Certified Payroll') {
      // Parse from filename: Certified_Payroll_[contractNum]_[borough]_YYYY-MM-DD[_<ContractorSlug>][_FILLED].pdf
      // Optional <ContractorSlug> after the date appears when a billing
      // remap split the CP by sub-prime — used to pick the right
      // contractor folder when multiple primes worked the same source job.
      const match = cleanName.match(/Certified_Payroll_([^_]+)_([^_]+)_(\d{4}-\d{2}-\d{2})(?:_([A-Za-z0-9]+))?/);
      if (!match) {
        return { success: false, reason: 'Could not parse contract info from Certified Payroll filename' };
      }
      const [, contractNum, borough, weekStartStr, fileContractorSlug] = match;
      const _cpIsSuffixMarker = fileContractorSlug === 'FILLED' || fileContractorSlug === 'MANUAL';
      const _cpUseFileContractor = fileContractorSlug && !_cpIsSuffixMarker;
      const _cpSlugMatches = (co) => String(co || '').replace(/[^A-Za-z0-9]/g, '') === fileContractorSlug;
      const weekStart = new Date(weekStartStr + 'T12:00:00');
      const wos = getWOsForPayrollWeek_(contractNum, borough, weekStart, ss);
      let contractor = '';
      if (_cpUseFileContractor) {
        // Narrow to the WOs matching the filename's contractor slug —
        // when both primes worked the same source job, picking the
        // first wo's contractor would archive to the wrong folder.
        const filtered = wos.filter(w => _cpSlugMatches(w.contractor));
        if (filtered.length > 0) contractor = filtered[0].contractor || '';
      }
      if (!contractor) contractor = wos.length > 0 ? (wos[0].contractor || '') : '';
      if (!contractor) {
        const clSheet = ss.getSheetByName('Contract Lookup');
        const clRow = clSheet
          ? clSheet.getDataRange().getValues()
              .find(r => String(r[0]).includes(contractNum) && String(r[1]) === borough)
          : null;
        if (clRow) contractor = String(clRow[5] || '').split(',')[0].trim();
      }
      if (!contractor) {
        return {
          success: false,
          reason: `No contractor found — no Field Reports logged for contract ${contractNum}/${borough} during week of ${weekStartStr}, and Contract Lookup has no matching row either. Add a Field Report for the week or a Contract Lookup entry.`
        };
      }

      const contractFolder = getOrCreateSubfolder_(
        getOrCreateSubfolder_(archiveRoot, contractor),
        `${contractNum} - ${getBoroughName_(borough)}`
      );
      // Master copy at contract level
      const cpMasterCopy = file.makeCopy(cleanName, getOrCreateSubfolder_(contractFolder, 'Certified Payroll'));
      // Duplicate into each WO folder worked during that payroll week for this contract
      wos.forEach(wo => {
        const woFolder = getOrCreateSubfolder_(contractFolder, `${wo.id} - ${wo.location}`);
        file.makeCopy(cleanName, woFolder);
      });
      Logger.log(`📁 Archived Certified Payroll → ${contractor}/${contractNum} + ${wos.length} WO folder(s)`);

      // Doc Lifecycle Log: one row per (week, contract, borough).
      try {
        _upsertDocLifecycleRow_(ss, {
          doc_id:       _docLifecycleId_('Certified Payroll', weekStartStr, contractNum, borough),
          doc_type:     'Certified Payroll',
          anchor:       weekStartStr,
          contractor:   contractor,
          contract_num: contractNum,
          borough:      borough,
          wo_ids:       wos.map(w => w.id),
          done:         true,
          file_id:      cpMasterCopy.getId(),
        });
      } catch (e) {
        Logger.log('⚠️ archiveDocument_ CP: Doc Lifecycle Log upsert failed: ' + e.message);
      }

      return { success: true, doc_type: 'Certified Payroll', wo_ids: wos.map(w => w.id) };

    } else {
      return { success: false, reason: `Unknown docType '${docType}' — no archive path defined` };
    }
  } catch (err) {
    const stack = err && err.stack || String(err);
    Logger.log(`❌ archiveDocument_ threw for ${file.getName()} (${docType}/${woId}): ${stack}`);
    return { success: false, reason: `Exception: ${err && err.message || err}` };
  }
}

/**
 * Safety net for the approved-docs pipeline: returns the Drive folder
 * where files land when archiveDocument_ can't file them normally
 * (missing tracker row, unparsable filename, Drive hiccup, etc). Sits
 * inside the Archive folder itself so admins can open it alongside the
 * rest of the archive to triage.
 *
 * Creates the folder + caches the ID the first time it's called, so
 * old deployments pick it up without needing a re-run of
 * createFolderStructure_.
 */
function getOrCreateArchiveErrorsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('ARCHIVE_ERRORS_ID');
  if (cached) {
    try { return DriveApp.getFolderById(cached); } catch (e) {
      Logger.log(`⚠️ Cached ARCHIVE_ERRORS_ID is stale (${cached}) — recreating`);
    }
  }
  const archiveId = props.getProperty('ARCHIVE_ID');
  if (!archiveId) {
    Logger.log('❌ Cannot create Archive Errors folder — ARCHIVE_ID is not set');
    return null;
  }
  const errors = getOrCreateSubfolder_(DriveApp.getFolderById(archiveId), '⚠️ Archive Errors');
  props.setProperty('ARCHIVE_ERRORS_ID', errors.getId());
  return errors;
}

/**
 * Lazy-create the Work Day Log sheet — the queue source for Sign-In tab.
 * Each Field Report submit appends one row here. The Sign-In tab reads
 * rows where Status='Pending' and groups them by (date, contract, borough)
 * so the user can file a single sign-in covering all WOs touched in a
 * shift.
 */
function _getOrCreateWorkDayLogSheet_(ss) {
  let sheet = ss.getSheetByName('Work Day Log');
  if (sheet) return sheet;

  sheet = ss.insertSheet('Work Day Log');
  const headers = [
    'Date', 'Work Order #', 'Prime Contractor', 'Contract #', 'Borough',
    'Location', 'Field Report Submitted At', 'Crew Chief', 'Sign-In Status'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Status dropdown — keeps human-edited corrections sane. Status moved
  // to col 9 (1-idx) when Crew Chief was inserted at col 8.
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Submitted', 'Skipped'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 9, 1000, 1).setDataValidation(statusRule);

  return sheet;
}

/**
 * Filename prefixes that identify an "auxiliary" archived doc (sign-ins,
 * invoices, etc) sitting alongside the WO doc in a WO archive folder.
 * Anything in the WO folder whose name does NOT start with one of these
 * is treated as the WO document itself — i.e. either the original scan
 * (with whatever filename the user uploaded under) or the canonical
 * `WO_<woId>.pdf` produced by a prior merged-CFR archive run.
 */
const _AUX_DOC_PREFIXES_ = Object.freeze([
  'SignIn_', 'Invoice_', 'CFR_', 'Contractor_Field_Report_',
  'Production_Log_', 'Certified_Payroll_'
]);

function _isAuxDocName_(name) {
  return _AUX_DOC_PREFIXES_.some(p => String(name || '').startsWith(p));
}

/**
 * Atomic-swap the WO doc in a WO archive folder. Trashes any existing
 * "WO doc" file (canonical name OR a non-auxiliary PDF — i.e. the
 * original scan upload) and copies `newFile` in as `WO_<woId>.pdf`.
 * Aux docs (sign-ins, invoices, etc) are left untouched.
 */
function _replaceArchivedWODoc_(woFolder, newFile, woId) {
  const canonical = `WO_${woId}.pdf`;
  const files = woFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== 'application/pdf') continue;
    const name = f.getName();
    const isWODoc = (name === canonical) || !_isAuxDocName_(name);
    if (isWODoc) f.setTrashed(true);
  }
  newFile.makeCopy(canonical, woFolder);
}

/**
 * action: lookup_archived_wo_pdf
 *
 * Finds the current "WO document" in the WO's archive folder and
 * returns its Drive file ID + filename so the Python worker can fetch
 * it via Drive API and merge a freshly-rendered CFR page into it.
 *
 * body.data: { wo_id }
 * response: { found: bool, file_id?, filename? }
 *
 * Lookup precedence:
 *   1. Canonical name `WO_<woId>.pdf` (produced by a prior merged-CFR
 *      archive run).
 *   2. Any PDF in the folder whose name doesn't match a known aux-doc
 *      prefix (the original scan, before its first CFR was merged).
 */
function handleLookupArchivedWOPdf_(body) {
  const d = body.data || {};
  const woId = String(d.wo_id || '').trim();
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);

  const props       = PropertiesService.getScriptProperties();
  const archiveId   = props.getProperty('ARCHIVE_ID');
  if (!archiveId) return jsonResponse_({ error: 'ARCHIVE_ID not set' }, 500);
  const archiveRoot = DriveApp.getFolderById(archiveId);
  const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const woFolder = getWOFolder_(archiveRoot, woId, ss);
  if (!woFolder) return jsonResponse_({ found: false });

  const canonical = `WO_${woId}.pdf`;
  let canonicalHit = null;
  let fallbackHit  = null;
  const files = woFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== 'application/pdf') continue;
    const name = f.getName();
    if (name === canonical) { canonicalHit = f; break; }
    if (!_isAuxDocName_(name) && !fallbackHit) fallbackHit = f;
  }

  const target = canonicalHit || fallbackHit;
  if (!target) return jsonResponse_({ found: false });
  return jsonResponse_({
    found:    true,
    file_id:  target.getId(),
    filename: target.getName(),
  });
}

/**
 * action: list_metro_completed_docs
 *
 * One-off bundling helper. Returns metadata for:
 *   - every Status='Completed' Metro Express WO's archived PDF (the
 *     merged WO/CFR doc), and
 *   - every Metro Express production log archived under
 *     Archive/Metro Express/<contract-borough>/Production Logs/
 *     (deduped by filename — the same daily log lives under each
 *     contract folder it covers).
 *
 * Caller fetches each file's bytes via the existing
 * get_drive_file_bytes action.
 */
function handleListMetroCompletedDocs_(body) {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData  = woSheet.getDataRange().getValues();

  const props       = PropertiesService.getScriptProperties();
  const archiveId   = props.getProperty('ARCHIVE_ID');
  if (!archiveId) return jsonResponse_({ error: 'ARCHIVE_ID not set' }, 500);
  const archiveRoot = DriveApp.getFolderById(archiveId);

  // ── Collect every completed Metro Express WO doc ───────────────
  const wos = [];
  woData.slice(1).forEach(r => {
    if (!r[0]) return;
    const contractor = String(r[1] || '').trim();
    const status     = String(r[15] || '').trim().toLowerCase();
    if (contractor !== 'Metro Express') return;
    if (status     !== 'completed')     return;

    const woId = String(r[0]);
    const woFolder = getWOFolder_(archiveRoot, woId, ss);
    if (!woFolder) {
      wos.push({ wo_id: woId, error: 'WO folder not found' });
      return;
    }
    // Find the WO doc — canonical first, else any non-aux PDF.
    const canonical = `WO_${woId}.pdf`;
    let canonicalHit = null, fallbackHit = null;
    const files = woFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getMimeType() !== 'application/pdf') continue;
      const name = f.getName();
      if (name === canonical) { canonicalHit = f; break; }
      if (!_isAuxDocName_(name) && !fallbackHit) fallbackHit = f;
    }
    const target = canonicalHit || fallbackHit;
    if (!target) {
      wos.push({ wo_id: woId, error: 'WO PDF not found in folder' });
      return;
    }

    const workEnd = r[18];
    let workEndIso = '';
    if (workEnd instanceof Date && !isNaN(workEnd.getTime())) {
      workEndIso = Utilities.formatDate(workEnd, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    } else if (workEnd) {
      const d = new Date(workEnd);
      if (!isNaN(d.getTime())) {
        workEndIso = Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      }
    }

    wos.push({
      wo_id:         woId,
      file_id:       target.getId(),
      filename:      target.getName(),
      work_end_date: workEndIso,
      contract_num:  String(r[2] || '').split('/')[0],
      borough:       String(r[3] || ''),
      location:      String(r[5] || ''),
    });
  });

  // ── Collect every Metro production log (master copies) ─────────
  // Path: Archive / Metro Express / <contract-borough> / Production Logs /
  const productionLogs = [];
  const metroIt = archiveRoot.getFoldersByName('Metro Express');
  if (metroIt.hasNext()) {
    const metroFolder = metroIt.next();
    const seenNames = {};
    const contractFolders = metroFolder.getFolders();
    while (contractFolders.hasNext()) {
      const cf = contractFolders.next();
      const plIt = cf.getFoldersByName('Production Logs');
      if (!plIt.hasNext()) continue;
      const plFolder = plIt.next();
      const plFiles = plFolder.getFiles();
      while (plFiles.hasNext()) {
        const f = plFiles.next();
        if (f.getMimeType() !== 'application/pdf') continue;
        const name = f.getName();
        if (seenNames[name]) continue;
        seenNames[name] = true;
        const m = name.match(/(\d{4}-\d{2}-\d{2})/);
        productionLogs.push({
          date:     m ? m[1] : '',
          file_id:  f.getId(),
          filename: name,
        });
      }
    }
  }

  return jsonResponse_({ wos, production_logs: productionLogs });
}

/** Get or create the WO-level subfolder: Archive/Contractor/Contract-Borough/WO#-Location */
function getWOFolder_(archiveRoot, woId, ss) {
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData  = woSheet.getDataRange().getValues();
  // 0-based row index (in `woData`) of the matching WO; the actual
  // sheet row is `idx + 1` because `woData` includes the header at [0].
  const idx = woData.findIndex(r => r[0] === woId);
  if (idx < 1) return null;
  const woRow       = woData[idx];
  const contractor  = woRow[1] || 'General';
  const contractNum = String(woRow[2]).split('/')[0];
  const borough     = woRow[3];
  const location    = woRow[5];
  const contractFolder = getOrCreateSubfolder_(
    getOrCreateSubfolder_(archiveRoot, contractor),
    `${contractNum} - ${getBoroughName_(borough)}`
  );
  const woFolder = getOrCreateSubfolder_(contractFolder, `${woId} - ${location}`);

  // Persist the URL on the WO row (col 43 / 0-idx 42) the first time
  // we resolve it. Subsequent dashboard refreshes read it from the sheet
  // instead of walking Drive — see handleGetDashboardData_. Wrapped in
  // try/catch because URL persistence must never break archiving.
  try {
    if (!woRow[42]) {
      ensureWoTrackerExtraCols_(woSheet);
      woSheet.getRange(idx + 1, 43).setValue(woFolder.getUrl());
    }
  } catch (e) {
    Logger.log('⚠️ getWOFolder_: failed to persist Archive Folder URL for ' + woId + ': ' + e.message);
  }

  return woFolder;
}

/**
 * Read-only counterpart to getWOFolder_. Walks the same path but never
 * creates anything — returns the folder if every level already exists,
 * `null` otherwise. Use this in dashboard / read paths so we don't
 * silently spawn empty folders for every WO on every refresh.
 *
 * Optional `woRowsCache` is the WO Tracker's getDataRange().getValues()
 * — pass it in when calling repeatedly to avoid re-fetching the sheet.
 */
function findWOFolder_(archiveRoot, woId, ss, woRowsCache) {
  const woData = woRowsCache || ss.getSheetByName('Work Order Tracker').getDataRange().getValues();
  const woRow  = woData.find(r => r[0] === woId);
  if (!woRow) return null;
  const contractor  = woRow[1] || 'General';
  const contractNum = String(woRow[2]).split('/')[0];
  const borough     = woRow[3];
  const location    = woRow[5];

  const cIt = archiveRoot.getFoldersByName(contractor);
  if (!cIt.hasNext()) return null;
  const contractorFolder = cIt.next();

  const cnIt = contractorFolder.getFoldersByName(`${contractNum} - ${getBoroughName_(borough)}`);
  if (!cnIt.hasNext()) return null;
  const contractFolder = cnIt.next();

  const woIt = contractFolder.getFoldersByName(`${woId} - ${location}`);
  if (!woIt.hasNext()) return null;
  return woIt.next();
}


// ── Geocoding helpers ─────────────────────────────────────────
//
// Server-side geocoding via Google's Geocoding API. Runs at scan
// intake (archiveWOFile_) and on demand from the backfill functions.
// Coords (cols 46–47) + warning (col 48) + timestamp (col 49) drive
// the Nav tab map.

// Borough → state suffix nudges Google toward NYC results when the
// query is ambiguous (e.g., "5 AVE" exists in multiple boroughs).
const _NYC_BOROUGH_TO_AREA_ = {
  'MN': 'Manhattan, New York, NY',
  'BK': 'Brooklyn, NY',
  'BX': 'Bronx, NY',
  'QN': 'Queens, NY',
  'SI': 'Staten Island, NY',
  // Long forms (in case the Tracker has either)
  'Manhattan':     'Manhattan, New York, NY',
  'Brooklyn':      'Brooklyn, NY',
  'Bronx':         'Bronx, NY',
  'Queens':        'Queens, NY',
  'Staten Island': 'Staten Island, NY',
};

function _boroughArea_(borough) {
  const key = String(borough || '').trim();
  return _NYC_BOROUGH_TO_AREA_[key] || (key ? `${key}, NY` : 'New York, NY');
}

// NYC bounding box (SW → NE corners) used as a viewport bias on every
// geocode call. Google will *prefer* results inside this box but can
// still return points outside it — `components=administrative_area:NY`
// is the hard restriction that keeps us from drifting into NJ or CT.
// Box covers all five boroughs with a small buffer.
const _NYC_GEOCODE_BOUNDS_ = '40.477,-74.259|40.917,-73.700';

// Single Google Geocoding API call. Returns { lat, lng } on success
// (status === 'OK' and at least one result), null otherwise.
function _geocodeOne_(addr) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GOOGLE_MAPS_API_KEY');
  if (!key) throw new Error('GOOGLE_MAPS_API_KEY script property not set');

  // `bounds` biases toward NYC viewport. `components` restricts to NY
  // state — combined they kill the "5 AVE in Mt Vernon" / "MAIN ST,
  // Long Island" matches that were causing wild cluster spreads.
  const url = 'https://maps.googleapis.com/maps/api/geocode/json'
    + '?address=' + encodeURIComponent(addr)
    + '&bounds=' + encodeURIComponent(_NYC_GEOCODE_BOUNDS_)
    + '&components=' + encodeURIComponent('country:US|administrative_area:NY')
    + '&region=us'
    + '&key=' + encodeURIComponent(key);

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('⚠️ _geocodeOne_ HTTP ' + code + ' for ' + addr);
      return null;
    }
    const body = JSON.parse(resp.getContentText());
    if (body.status === 'ZERO_RESULTS') return null;
    if (body.status !== 'OK' || !Array.isArray(body.results) || body.results.length === 0) {
      Logger.log('⚠️ _geocodeOne_ status=' + body.status + ' for ' + addr);
      return null;
    }
    const loc = body.results[0].geometry && body.results[0].geometry.location;
    if (!loc || typeof loc.lat !== 'number' || typeof loc.lng !== 'number') return null;
    return { lat: loc.lat, lng: loc.lng };
  } catch (e) {
    Logger.log('⚠️ _geocodeOne_ error for ' + addr + ': ' + e.message);
    return null;
  }
}

// Great-circle distance in miles between two {lat, lng} points.
function _haversineMiles_(a, b) {
  if (!a || !b) return Infinity;
  const R = 3958.8;  // Earth radius in miles
  const toRad = (d) => d * Math.PI / 180;
  const dLat = toRad(b.lat - a.lat);
  const dLng = toRad(b.lng - a.lng);
  const lat1 = toRad(a.lat);
  const lat2 = toRad(b.lat);
  const h = Math.sin(dLat / 2) * Math.sin(dLat / 2)
          + Math.sin(dLng / 2) * Math.sin(dLng / 2) * Math.cos(lat1) * Math.cos(lat2);
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(h)));
}

// Orchestrates: build primary + validation queries, geocode them,
// cluster-check, return one of:
//   { lat, lng }           — confident pin, no warning
//   { warning: '...' }     — could not produce a confident pin
//
// `d` shape (matches archiveWOFile_'s synthetic):
//   { work_order_id, location, from_street, to_street, borough }
function geocodeWO_(d, ss) {
  const wo       = String(d.work_order_id || '').trim();
  const location = String(d.location    || '').trim();
  const from     = String(d.from_street || '').trim();
  const to       = String(d.to_street   || '').trim();
  const borough  = String(d.borough     || '').trim();
  const area     = _boroughArea_(borough);

  if (!location || !from) {
    return { warning: 'Missing location or from_street — cannot geocode' };
  }

  // Primary: start of job
  const primaryAddr = `${location} and ${from}, ${area}, USA`;
  const primary = _geocodeOne_(primaryAddr);

  // Validation queries. Every query is a full intersection of the form
  // "{location} and {crossStreet}" — the Tracker stores `location` as
  // the single road work is being done on, so a cross-street alone
  // ("13 AV") would geocode to a random point along that avenue.
  //
  // Sources:
  //   1. `to_street`   — always tried if present.
  //   2. Marking Items — `Intersection Grid` rows for this WO; col 5
  //      holds the cross-street (one per intersection on the run).
  //      Other WO Sections (Top Table / Manual) don't have meaningful
  //      cross-streets and are skipped. Up to 3 distinct picks.
  const validationAddrs = [];
  const norm = (s) => String(s || '').trim().toUpperCase();
  const seen = {};
  if (from) seen[norm(from)] = true;  // primary already covers from_street
  if (to) {
    seen[norm(to)] = true;
    validationAddrs.push(`${location} and ${to}, ${area}, USA`);
  }

  try {
    const miSheet = ss.getSheetByName('Marking Items');
    if (miSheet) {
      const data = miSheet.getDataRange().getValues();
      const intersections = [];
      for (let i = 1; i < data.length && intersections.length < 6; i++) {
        if (String(data[i][1] || '').trim() !== wo) continue;
        if (String(data[i][4] || '').trim() !== 'Intersection Grid') continue;
        const crossSt = String(data[i][5] || '').trim();
        if (!crossSt) continue;
        const key = norm(crossSt);
        if (seen[key]) continue;
        seen[key] = true;
        intersections.push(crossSt);
      }
      // Shuffle + take up to 3
      intersections.sort(() => Math.random() - 0.5);
      intersections.slice(0, 3).forEach(crossSt => {
        validationAddrs.push(`${location} and ${crossSt}, ${area}, USA`);
      });
    }
  } catch (e) {
    Logger.log('⚠️ geocodeWO_ marking-items read failed for ' + wo + ': ' + e.message);
  }

  // Resolve each validation address to coords, keeping the address
  // string so we can log diagnostics when the cluster check fails.
  const validationResolved = validationAddrs.map(addr => ({
    addr,
    point: _geocodeOne_(addr),
  })).filter(v => v.point);

  // Resolve
  if (!primary && validationResolved.length === 0) {
    return { warning: 'Geocoding failed: no results for any address candidate' };
  }
  if (!primary) {
    return { warning: 'Primary intersection unresolvable — admin must set coords manually' };
  }

  // Distances from primary, sorted ascending. With 3+ validation
  // points we drop the single worst outlier — Google occasionally
  // returns a wildly wrong match for ambiguous street names and we
  // don't want one bad candidate vetoing an otherwise-tight cluster.
  const distances = validationResolved
    .map(v => ({ ...v, dist: _haversineMiles_(primary, v.point) }))
    .sort((a, b) => a.dist - b.dist);
  const considered = distances.length >= 3 ? distances.slice(0, -1) : distances;
  const maxSpread = considered.reduce((m, v) => Math.max(m, v.dist), 0);

  if (maxSpread > 1.0) {
    // Log every candidate (address + coords + dist) so the admin can
    // see whether the WO data is wrong, the geocoder is mis-matching,
    // or the WO genuinely spans > 1 mi and needs manual placement.
    Logger.log(`⚠️ geocodeWO_ ${wo} cluster spread ${maxSpread.toFixed(2)}mi`);
    Logger.log(`   primary: ${primaryAddr}`);
    Logger.log(`            → ${primary.lat.toFixed(5)}, ${primary.lng.toFixed(5)}`);
    distances.forEach(v => {
      const flagged = considered.includes(v) ? '  ' : '✂ ';  // ✂ = dropped outlier
      Logger.log(`   ${flagged}${v.dist.toFixed(2)}mi  ${v.addr}`);
      Logger.log(`        → ${v.point.lat.toFixed(5)}, ${v.point.lng.toFixed(5)}`);
    });
    return { warning: `Cluster spread ${maxSpread.toFixed(2)} miles between candidates — verify pin` };
  }

  return { lat: primary.lat, lng: primary.lng };
}

// Write the geocode result to the Tracker row + log warnings to
// Automation Log. Idempotent — overwrites any prior values for that
// WO. Wrapped in try/catch by callers so a geocoding failure never
// breaks the scan-intake pipeline.
function _persistGeocode_(ss, woId, result) {
  const woSheet = ss.getSheetByName('Work Order Tracker');
  ensureWoTrackerExtraCols_(woSheet);
  const data = woSheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === woId) { rowIdx = i; break; }
  }
  if (rowIdx === -1) {
    Logger.log('⚠️ _persistGeocode_: WO not found in Tracker: ' + woId);
    return;
  }

  const sheetRow = rowIdx + 1;
  // Cols 46 (Latitude), 47 (Longitude), 48 (Geocode Warning), 49 (Geocoded At)
  if (result.lat != null && result.lng != null) {
    woSheet.getRange(sheetRow, 46).setValue(result.lat);
    woSheet.getRange(sheetRow, 47).setValue(result.lng);
    woSheet.getRange(sheetRow, 48).setValue('');
  } else {
    // Don't wipe an existing good pin on a re-attempt that returns
    // only a warning — only update the warning column.
    woSheet.getRange(sheetRow, 48).setValue(String(result.warning || ''));
  }
  woSheet.getRange(sheetRow, 49).setValue(new Date());

  if (result.warning) {
    const logSheet = ss.getSheetByName('Automation Log');
    if (logSheet) {
      logSheet.appendRow([
        new Date(), 'Geocoder', 'Geocode warning',
        woId, result.warning, 'Warning', '', 'No',
      ]);
    }
  }
}


/**
 * Build a fast WO# → Location map from the Work Order Tracker. Used by
 * the WO-listing helpers below so multi-WO sign-in rows (where the WO#
 * column is a comma-list "RM-1, RM-2") still resolve each WO's actual
 * location for archive-folder paths.
 */
function _buildLocationByWOMap_(ss) {
  const woData = ss.getSheetByName('Work Order Tracker').getDataRange().getValues();
  const byId = {};
  woData.slice(1).forEach(r => {
    if (r[0]) byId[String(r[0])] = String(r[5] || '');
  });
  return byId;
}

/** Split a Daily Sign-In Data WO# cell into a list (handles comma-list values). */
function _splitWOIds_(cell) {
  return String(cell || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

/** Return WOs worked on a given date, grouped by "contractor|contractNum|borough" */
function getWOsForDate_(date, ss) {
  const data = ss.getSheetByName('Daily Sign-In Data').getDataRange().getValues();
  const locByWO = _buildLocationByWOMap_(ss);
  const wosByContract = {};
  const seen = new Set();
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    if (new Date(row[0]).toDateString() !== date.toDateString()) return;
    const key = `${row[2]}|${String(row[3]).split('/')[0]}|${row[4]}`;
    if (!wosByContract[key]) wosByContract[key] = [];
    _splitWOIds_(row[1]).forEach(woId => {
      if (seen.has(woId)) return;
      seen.add(woId);
      wosByContract[key].push({ id: woId, location: locByWO[woId] || row[5] || '' });
    });
  });
  return wosByContract;
}

/**
 * Return unique WOs for a contract+borough during a payroll week.
 * Each entry: { id, location, contractor } — contractor pulled from
 * Daily Sign-In Data's Prime Contractor column so the Cert Payroll
 * archive path can be derived without depending on Contract Lookup.
 */
function getWOsForPayrollWeek_(contractNum, borough, weekStart, ss) {
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  const data = ss.getSheetByName('Daily Sign-In Data').getDataRange().getValues();
  const locByWO = _buildLocationByWOMap_(ss);
  const wos = [];
  const seen = new Set();
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[0]);
    if (d < weekStart || d > weekEnd) return;
    if (String(row[3]).split('/')[0] !== contractNum || row[4] !== borough) return;
    _splitWOIds_(row[1]).forEach(woId => {
      if (seen.has(woId)) return;
      seen.add(woId);
      wos.push({
        id:         woId,
        location:   locByWO[woId] || row[5] || '',
        contractor: String(row[2] || '').trim(),
      });
    });
  });
  return wos;
}


// ═══════════════════════════════════════════════════════════════
// 5. SIGN-IN SHEET DATA VALIDATION
// ═══════════════════════════════════════════════════════════════

/**
 * Validates sign-in data for common errors.
 * Run after data entry to flag issues before document generation.
 */
function validateSignInData(dateStr) {
  const targetDate = dateStr ? new Date(dateStr) : new Date();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const signInSheet = ss.getSheetByName('Daily Sign-In Data');
  const data = signInSheet.getDataRange().getValues();
  
  const issues = [];
  
  data.slice(1).forEach((row, idx) => {
    if (!row[0]) return;
    const rowDate = new Date(row[0]);
    if (rowDate.toDateString() !== targetDate.toDateString()) return;
    
    const rowNum = idx + 2;
    const woId = row[1];
    const empName = row[6];
    const timeIn = row[8];
    const timeOut = row[9];
    const hours = row[10];
    
    // Check: Missing employee name
    if (!empName) issues.push(`Row ${rowNum}: Missing employee name for ${woId}`);
    
    // Check: Missing time in/out
    if (!timeIn) issues.push(`Row ${rowNum}: Missing Time In for ${empName} on ${woId}`);
    if (!timeOut) issues.push(`Row ${rowNum}: Missing Time Out for ${empName} on ${woId}`);
    
    // Check: Hours > 12 (likely error)
    if (Number(hours) > 12) issues.push(`Row ${rowNum}: ${empName} logged ${hours} hours — verify this is correct`);
    
    // Check: Hours = 0
    if (Number(hours) === 0) issues.push(`Row ${rowNum}: ${empName} has 0 hours on ${woId}`);
    
    // Check: Missing SQFT
    if (!row[12]) issues.push(`Row ${rowNum}: Missing SQFT for ${woId}`);
    
    // Check: Missing paint/material
    if (!row[13]) issues.push(`Row ${rowNum}: Missing paint/material for ${woId}`);
    
    // Check: WO marked complete but missing classification
    if (!row[7]) issues.push(`Row ${rowNum}: Missing classification for ${empName}`);
  });
  
  // Check cross-WO consistency: same employee shouldn't have overlapping times
  // (This is a simplified check — could be made more sophisticated)
  
  if (issues.length > 0) {
    Logger.log('⚠️  VALIDATION ISSUES FOUND:\n' + issues.join('\n'));
    
    // Email admin with issues
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: `⚠️ Sign-In Data Issues — ${Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'MM/dd/yyyy')}`,
      htmlBody: `
        <h3>Sign-In Data Validation Issues</h3>
        <p>The following issues were found for ${Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'MM/dd/yyyy')}:</p>
        <ul>${issues.map(i => `<li>${i}</li>`).join('')}</ul>
        <p><a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit#gid=${signInSheet.getSheetId()}">Fix in Spreadsheet</a></p>
      `
    });
  } else {
    Logger.log('✅ All sign-in data validates cleanly');
  }
  
  return issues;
}


// ═══════════════════════════════════════════════════════════════
// 6. CERTIFIED PAYROLL GENERATION
// ═══════════════════════════════════════════════════════════════

/**
 * Generate certified payroll data for a given week.
 * Groups hours by Contract ID and creates separate entries per contract.
 * 
 * @param {string} weekStartStr - Sunday date in MM/DD/YYYY format
 */
/**
 * YTD gross pay for an employee across ALL projects, up to and
 * including the payroll week end date. Trusts the per-row col 11
 * (Overtime Hours) written by handleSubmitSignIn_, which already
 * applies the day-of-week OT rule with same-day cross-group lookback.
 * Deriving OT here would silently disagree with col 11 if anyone
 * hand-edits a row's OT contribution (e.g. correcting historical
 * cross-contract rows the old per-shift logic mis-split).
 *
 * Rates are resolved per pay-period week (Sun–Sat) AND per
 * classification using the week's end-date — so YTD stays correct
 * across rate-schedule boundaries AND across mid-year classification
 * changes (an employee who worked some weeks as LP and others as SAT
 * gets each bucket priced at the matching rate). Supplementals are
 * paid the same hours as wages (ST hours pair with ST supp, OT hours
 * pair with OT supp).
 *
 * signInData:   full Daily Sign-In Data getValues() (incl. header row).
 * empName:      employee name as written in Sign-In Data col 6.
 * payrollRates: result of _loadPayrollRates_(ss). Empty array → 0.
 * weekEnd:      Date object marking end of the payroll week (YTD cutoff).
 */
function computeYtdGrossForEmployee_(signInData, empName, payrollRates, weekEnd) {
  const normName = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const target   = normName(empName);
  if (!target) return 0;
  if (!Array.isArray(payrollRates) || payrollRates.length === 0) return 0;

  const yearStart = new Date(weekEnd.getFullYear(), 0, 1, 0, 0, 0);

  // Bucket rows into Sun–Sat weeks so we can apply the rate effective on
  // each week's end date. JS getDay(): Sun=0, Mon=1, ..., Sat=6.
  const weekEndKey = (d) => {
    // Days until upcoming Saturday (inclusive). For Saturday itself, 0.
    const daysToSat = 6 - d.getDay();
    const we = new Date(d.getFullYear(), d.getMonth(), d.getDate() + daysToSat, 23, 59, 59);
    return { key: Utilities.formatDate(we, CONFIG.TIMEZONE, 'yyyy-MM-dd'), date: we };
  };

  // Key by (weekEnd, classification) so a week with mixed classifications
  // gets one rate-resolved bucket per classification.
  const byWeekClass = {};
  signInData.slice(1).forEach(row => {
    if (!row[0]) return;
    const rowDate = new Date(row[0]);
    if (isNaN(rowDate.getTime())) return;
    if (rowDate < yearStart || rowDate > weekEnd) return;
    if (normName(row[6]) !== target) return;
    const cls = String(row[7] || '').trim();
    if (!cls) return;
    const hours = Number(row[10]) || 0;
    if (hours <= 0) return;
    const ot = Number(row[11]) || 0;
    const st = Math.max(0, hours - ot);
    const { key, date: we } = weekEndKey(rowDate);
    const wkClsKey = key + '|||' + cls;
    if (!byWeekClass[wkClsKey]) {
      byWeekClass[wkClsKey] = { weekEnd: we, classification: cls, st: 0, ot: 0 };
    }
    byWeekClass[wkClsKey].st += st;
    byWeekClass[wkClsKey].ot += ot;
  });

  let ytd = 0;
  Object.values(byWeekClass).forEach(({ weekEnd: we, classification: cls, st, ot }) => {
    const rate = _resolvePayrollRate_(payrollRates, cls, we);
    if (!rate) return; // No rate effective for this (week, classification) — skip silently in YTD
    ytd += st * ((rate.st_rate || 0) + (rate.st_supp || 0))
         + ot * ((rate.ot_rate || 0) + (rate.ot_supp || 0));
  });
  return ytd;
}


function generateCertifiedPayroll(weekStartStr, opts) {
  // Parse MM/DD/YYYY explicitly to avoid UTC-shift issues
  const parts = weekStartStr.trim().split('/');
  if (parts.length !== 3) {
    Logger.log('❌ Invalid date format. Use MM/DD/YYYY. Got: ' + weekStartStr);
    return 0;
  }
  const weekStart = new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]), 0, 0, 0);
  const weekEnd   = new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]) + 6, 23, 59, 59); // end of Sunday
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const signInSheet = ss.getSheetByName('Daily Sign-In Data');
  const data = signInSheet.getDataRange().getValues();
  const empSheet = ss.getSheetByName('Employee Registry');
  const empData = empSheet.getDataRange().getValues();
  const clSheet = ss.getSheetByName('Contract Lookup');
  const clData = clSheet.getDataRange().getValues();
  const cpSheet = ss.getSheetByName('Certified Payroll Tracker');

  // Prevailing-wage rates by classification (LP/SAT/...) with
  // effective-date semantics — single source of truth replacing the
  // per-employee rate cols that used to live on Employee Registry.
  const payrollRates = _loadPayrollRates_(ss);
  if (payrollRates.length === 0) {
    Logger.log('❌ Certified Payroll: Payroll Rates sheet is empty or missing. Run setupPayrollRates() first.');
    return 0;
  }

  // CP Tracker had legacy "ER Fringe" / "EMP Fringe" headers in cols
  // 25 / 26 (1-indexed). Those cells now carry ST/OT Supplemental
  // values per the classification rate table — relabel idempotently
  // so the sheet docs match what's written. Existing historical row
  // values are left untouched; only the header row changes.
  if (cpSheet) {
    const headerRow = cpSheet.getRange(1, 1, 1, cpSheet.getLastColumn()).getValues()[0];
    if (headerRow[24] === 'ER Fringe') cpSheet.getRange(1, 25).setValue('ST Supplemental');
    if (headerRow[25] === 'EMP Fringe') cpSheet.getRange(1, 26).setValue('OT Supplemental');
  }
  
  // Filter to this week's entries
  const weekEntries = data.slice(1).filter(row => {
    if (!row[0]) return false;
    const d = new Date(row[0]);
    return d >= weekStart && d <= weekEnd;
  });
  
  if (weekEntries.length === 0) {
    Logger.log('⚠️ No sign-in entries found for week of ' + weekStartStr + ' (' + weekStart.toDateString() + ' – ' + weekEnd.toDateString() + ')');
    Logger.log('    Check that Daily Sign-In Data has rows with dates in that range.');
    return 0;
  }
  
  // Group by Contract # + Borough (which maps to Contract ID)
  const byContract = {};
  weekEntries.forEach(row => {
    const contractNum = String(row[3]).split('/')[0]; // Strip suffix
    const borough = row[4];
    const key = `${contractNum}|${borough}`;
    if (!byContract[key]) byContract[key] = [];
    byContract[key].push(row);
  });

  // Per-doc generation: filter byContract to a single (contract, borough) tuple.
  if (opts && opts.contractFilter) {
    const wantKey = `${opts.contractFilter.contractNum}|${opts.contractFilter.borough}`;
    Object.keys(byContract).forEach(k => { if (k !== wantKey) delete byContract[k]; });
    if (!byContract[wantKey]) {
      Logger.log('⚠️ Certified Payroll: no Daily Sign-In rows for ' + wantKey +
                 ' in week of ' + weekStartStr + ' — nothing to generate.');
      return 0;
    }
  }
  
  // Full classification names — add new codes here as needed
  const CLASSIFICATION_NAMES = {
    'LP':   'Line Person',
    'SAT':  'Striping Assistant',
    'OP':   'Operator',
    'LAB':  'Laborer',
    'FGL':  'Flagger',
    'SUP':  'Supervisor',
  };
  const classificationName = code => CLASSIFICATION_NAMES[String(code).trim().toUpperCase()] || String(code).trim();

  // For each contract group, generate certified payroll entries.
  //
  // Sub-prime billing remap: when a (raw contract, borough) has a
  // remap rule (sub-prime work on a contract the prime didn't win),
  // split this bucket by contractor so each sub-prime gets their own
  // CP billed under their own contract. Without remap, the bucket
  // runs once for all contractors combined (existing behavior — one
  // CP per (contract, borough)).
  let cpSubBucketsGenerated = 0;
  Object.entries(byContract).forEach(([key, rawEntries]) => {
    const [rawContractNum, rawBorough] = key.split('|');

    const subBuckets = _hasBillingRemap_(rawContractNum, rawBorough)
      ? Object.entries(rawEntries.reduce((acc, r) => {
          const co = String(r[2] || '').trim();
          (acc[co] = acc[co] || []).push(r);
          return acc;
        }, {})).map(([contractor, entries]) => {
          const m = _billingRemap_(rawContractNum, rawBorough, contractor);
          return { contractor, contractNum: m.contractNum, borough: m.borough, entries };
        })
      : [{ contractor: null, contractNum: rawContractNum, borough: rawBorough, entries: rawEntries }];

  subBuckets.forEach(({ contractor: bucketContractor, contractNum, borough, entries }) => {
    cpSubBucketsGenerated++;

    // Look up Contract ID from lookup table (uses MAPPED contract+borough
    // when remap applied, so the right contract's registration/project
    // populate the form).
    const clRow = clData.find(r =>
      String(r[0]).includes(contractNum) && String(r[1]) === borough
    );
    const contractId = clRow ? clRow[3] : '⚠️ MISSING — CHECK LOOKUP TABLE';
    const projectName = clRow ? clRow[4] : '';

    // Group by (employee, classification). An employee who worked two
    // classifications in the same week (e.g. LP three days, SAT two
    // days) gets two entries on the CP — one per classification with
    // that classification's hours and rates — so the math on the form
    // is transparent and the rate column matches the work performed.
    // ST/OT per row come from cols 10 (Hours) and 11 (Overtime), which
    // handleSubmitSignIn_ already split with cross-group same-day
    // lookback; the OT correctly lands on whichever row crossed the 8h
    // limit, regardless of which classification that row was under.
    const byEmployeeClass = {};
    entries.forEach(row => {
      const emp = String(row[6] || '').trim();
      const cls = String(row[7] || '').trim();
      if (!emp || !cls) return;
      const key = emp + '|||' + cls;
      if (!byEmployeeClass[key]) {
        byEmployeeClass[key] = {
          name: emp, classification: cls,
          days: {}, stByDay: {}, otByDay: {}, totalHours: 0,
        };
      }
      const dayOfWeek = new Date(row[0]).getDay(); // 0=Sun, 1=Mon...6=Sat
      const hours = Number(row[10]) || 0;
      const ot    = Number(row[11]) || 0;
      const st    = Math.max(0, hours - ot);
      byEmployeeClass[key].days[dayOfWeek]    = (byEmployeeClass[key].days[dayOfWeek]    || 0) + hours;
      byEmployeeClass[key].stByDay[dayOfWeek] = (byEmployeeClass[key].stByDay[dayOfWeek] || 0) + st;
      byEmployeeClass[key].otByDay[dayOfWeek] = (byEmployeeClass[key].otByDay[dayOfWeek] || 0) + ot;
      byEmployeeClass[key].totalHours += hours;
    });
    
    // JS getDay() → form day index (Mon=0 … Sun=6)
    // Sun=0 … Sat=6  (matches form column order S M T W R F S)
    const jsDayToFormDay = { 0:0, 1:1, 2:2, 3:3, 4:4, 5:5, 6:6 };
    const DAY_LABELS = ['S','M','T','W','R','F','S'];

    // Build days array for JSON (Sun–Sat covering the payroll week)
    const daysJson = [];
    for (let fi = 0; fi < 7; fi++) {
      const d = new Date(weekStart);
      d.setDate(d.getDate() + fi);
      const month = d.getMonth() + 1;
      const day   = String(d.getDate()).padStart(2, '0');
      daysJson.push({ label: DAY_LABELS[fi], date: `${month}/${day}` });
    }

    const workersJson = [];

    // Pre-build a normalized full-name lookup for this contract group so
    // each employee match is O(1) and — crucially — doesn't collide when
    // two employees share a first name (previous version substring-matched
    // on the first name only).
    const normName = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
    const empByName = {};
    empData.slice(1).forEach(r => {
      const n = normName(r[1]);
      if (n) empByName[n] = r;
    });

    // Write to Certified Payroll Tracker. Sort alphabetically by name
    // and then classification so an employee's two entries appear
    // adjacent on the form rather than scattered by insertion order.
    const sortedEntries = Object.values(byEmployeeClass).sort((a, b) => {
      const nameCmp = a.name.localeCompare(b.name);
      return nameCmp !== 0 ? nameCmp : a.classification.localeCompare(b.classification);
    });
    sortedEntries.forEach(info => {
      const empName = info.name;
      // Address + SSN4 still come from Employee Registry (personal info).
      // Rates no longer do — they come from Payroll Rates keyed by
      // classification + week-end date.
      const empRow = empByName[normName(empName)] || null;
      if (!empRow) {
        Logger.log(`⚠️ Certified Payroll: no Employee Registry row matches ${JSON.stringify(empName)} — address/SSN will be blank.`);
      }
      const empAddr = empRow ? String(empRow[2] || '') : '';
      const empSsn4 = empRow ? String(empRow[3] || '') : '';

      // Per-week rate resolution: use the rate effective on this
      // payroll week's end date (Sunday). Skip the worker (with a
      // loud warning) if no rate row applies — happens when the
      // sign-in row's classification (info.classification) isn't in
      // the Payroll Rates table.
      const rateRow = _resolvePayrollRate_(payrollRates, info.classification, weekEnd);
      if (!rateRow) {
        Logger.log(`⚠️ Certified Payroll: no Payroll Rates row for classification ${JSON.stringify(info.classification)} effective on or before ${weekEnd.toDateString()} — skipping ${empName}.`);
        return;
      }
      const stRate = rateRow.st_rate || 0;
      const otRate = rateRow.ot_rate || 0;
      const stSupp = rateRow.st_supp || 0;
      const otSupp = rateRow.ot_supp || 0;

      // YTD gross across all projects this year. Used to fill the
      // "Total Gross Pay (All Work)" column on the certified payroll
      // form. Resolves rates per-week AND per-classification so a YTD
      // spanning a rate-schedule boundary OR a mid-year classification
      // change reflects the actual rates paid. Same value goes on each
      // of an employee's CP entries (per-person YTD, not per-class).
      const ytdGross = computeYtdGrossForEmployee_(
        data, empName, payrollRates, weekEnd
      );

      // Sum this contract's per-day ST/OT from the col-10/col-11 split
      // already computed at submission time. Weekend (all OT) and
      // >8/day rules are enforced upstream in handleSubmitSignIn_;
      // here we just total what's in the rows for this contract.
      let totalST = 0;
      let totalOT = 0;
      const stHours = ['0','0','0','0','0','0','0'];
      const otHours = ['0','0','0','0','0','0','0'];

      for (let dow = 0; dow < 7; dow++) {
        const st = info.stByDay[dow] || 0;
        const ot = info.otByDay[dow] || 0;
        if (st === 0 && ot === 0) continue;
        const fi = jsDayToFormDay[dow];
        if (st > 0) stHours[fi] = String(st);
        if (ot > 0) otHours[fi] = String(ot);
        totalST += st;
        totalOT += ot;
      }

      // Gross pay = ST hours × (st_rate + st_supp) + OT hours × (ot_rate + ot_supp).
      // Supplementals are paid per-hour-worked, ST hours always pair
      // with ST supp, OT hours always pair with OT supp.
      const grossPay = (totalST * (stRate + stSupp)) + (totalOT * (otRate + otSupp));

      cpSheet.appendRow([
        weekStart, weekEnd,
        contractNum, borough, contractId,
        projectName,
        empName, info.classification,
        info.days[0] || 0, // Sun
        info.days[1] || 0, // Mon
        info.days[2] || 0, // Tue
        info.days[3] || 0, // Wed
        info.days[4] || 0, // Thu
        info.days[5] || 0, // Fri
        info.days[6] || 0, // Sat
        totalST, totalOT,
        stRate, otRate, grossPay,
        '', '', '', // Total gross, withholdings, net — need payroll software data
        stSupp, otSupp,
        'Pending Verification', '', // Match status
        'No', ''   // Sent status
      ]);

      workersJson.push({
        name:            empName,
        address:         empAddr,
        ssn4:            empSsn4,
        trade:           classificationName(info.classification),
        journeyperson:   true,
        st_hours:        stHours,
        ot_hours:        otHours,
        total_st:        String(totalST),
        total_ot:        String(totalOT),
        // "Hourly Rate of Pay" on the form folds the supplemental
        // (fringe) rate into the cash rate per NYC convention — total
        // prevailing-wage rate per hour. The "Hourly Contributions to
        // Benefit Funds or Accounts" column still shows the supp
        // breakdown via supp_st / supp_ot below.
        rate_st:         (stRate + stSupp).toFixed(2),
        rate_ot:         (otRate + otSupp).toFixed(2),
        supp_st:         stSupp.toFixed(2),
        supp_ot:         otSupp.toFixed(2),
        gross_pay:       grossPay.toFixed(2),
        total_gross_pay: ytdGross.toFixed(2),   // YTD across all projects
        net_pay:         '',
        deductions:      '',
        annualized_rate: ''
      });
    });

    // ── JSON export for local PDF filler ─────────────────────────────────────
    const weekEndFormatted = Utilities.formatDate(weekEnd, CONFIG.TIMEZONE, 'MM/dd/yyyy');

    const cpContractNumber = clRow ? String(clRow[0] || '') : '';

    // Note: no `signatory` block here. The OFFICER/TITLE/DATE/YEAR fields
    // are intentionally left blank at fill time — they're populated when
    // the admin clicks "Approve & Sign" in the Approvals tab, which runs
    // POST /api/approvals/:fileId/approve-cert-payroll and overlays the
    // principal signature + name + title + today's date via pdf-lib.
    const cpJson = {
      _type:                 'certified_payroll',
      payroll_number:        String(_payrollWeekNumber_(weekStart)),
      week_ending:           weekEndFormatted,
      employer:              CONFIG.EMPLOYER,
      prime_contractor:      CONFIG.SIGNATORY.name,
      contract_registration: contractId,
      agency:                'NYC DOT',
      agency_pin:            cpContractNumber,
      project_address:       '',
      project_name:          projectName,
      pla:                   false,
      days:                  daysJson,
      workers:               workersJson,
    };
    const cpProps = PropertiesService.getScriptProperties();
    const cpFolder = getOrCreateSubfolder_(
      DriveApp.getFolderById(cpProps.getProperty('NEEDS_REVIEW_ID')), 'Certified Payroll'
    );
    // Filename uses the week START date (Sunday) — reads naturally as
    // "week of 2026-04-19" and matches the semantic in the archive code.
    // The form still prints "WEEK ENDING DATE" (Saturday) in its header
    // because that's the standard payroll convention on the form itself.
    //
    // Filename keeps RAW (contract, borough) so archive parsers still
    // match against WO Tracker (also raw). When a split happened
    // (sub-prime billing remap), append a contractor slug to
    // disambiguate the two PDFs that share the same raw key — without
    // the slug they'd collide in Drive and one would clobber the other.
    const _cpRawContract = bucketContractor ? rawContractNum : contractNum;
    const _cpRawBorough  = bucketContractor ? rawBorough     : borough;
    const _cpContractorSlug = bucketContractor
      ? '_' + bucketContractor.replace(/[^A-Za-z0-9]/g, '')
      : '';
    const cpJsonName = `Certified_Payroll_${_cpRawContract}_${_cpRawBorough}_${Utilities.formatDate(weekStart, CONFIG.TIMEZONE, 'yyyy-MM-dd')}${_cpContractorSlug}.json`;
    cpFolder.createFile(cpJsonName, JSON.stringify(cpJson, null, 2), MimeType.PLAIN_TEXT);
    Logger.log(`✅ Certified payroll JSON exported: ${cpJsonName}`);
    // ─────────────────────────────────────────────────────────────────────────

    Logger.log(`✅ Certified payroll entries created for ${contractNum}/${borough}${bucketContractor ? ' (' + bucketContractor + ')' : ''}: ${sortedEntries.length} entries (one per employee × classification)`);
  });
  });

  const contractCount = cpSubBucketsGenerated;

  // Flag for human review
  const logSheet = ss.getSheetByName('Automation Log');
  logSheet.appendRow([
    new Date(), 'Certified Payroll Generator', 'Weekly trigger',
    `Week of ${weekStartStr}`,
    `${contractCount} contract groups processed`,
    'Generated',
    'Total Gross Pay, Withholdings, and Net Pay need payroll software verification',
    'Yes — Cross-check with payrollforconstruction.com and complete missing fields'
  ]);

  return contractCount;
}


// ═══════════════════════════════════════════════════════════════
// 7. DAILY SUMMARY EMAIL
// ═══════════════════════════════════════════════════════════════

function sendDailySummary() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData = woSheet.getDataRange().getValues();
  
  // Count by status
  const statusCounts = {};
  woData.slice(1).forEach(row => {
    const status = row[15] || 'Unknown';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });
  
  // Pending items
  const received = woData.slice(1).filter(r => r[15] === 'Received');
  const inProgress = woData.slice(1).filter(r => r[15] === 'In Progress');
  
  // AR summary
  const arSheet = ss.getSheetByName('Invoices & AR');
  const arData = arSheet.getDataRange().getValues();
  const unpaidInvoices = arData.slice(1).filter(r => r[11] !== 'Paid' && r[10]);
  const totalOutstanding = unpaidInvoices.reduce((sum, r) => sum + (Number(r[10]) || 0), 0);
  
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'EEEE, MMMM d, yyyy');
  
  const html = `
    <div style="font-family: Arial, sans-serif; max-width: 600px;">
      <h2 style="color: #1F2937;">Oneiro Operations — Daily Summary</h2>
      <p style="color: #6B7280;">${today}</p>
      
      <h3 style="color: #374151;">Work Order Status</h3>
      <table style="border-collapse: collapse; width: 100%;">
        ${Object.entries(statusCounts).map(([status, count]) => `
          <tr>
            <td style="padding: 6px 12px; border-bottom: 1px solid #E5E7EB;">${status}</td>
            <td style="padding: 6px 12px; border-bottom: 1px solid #E5E7EB; font-weight: bold;">${count}</td>
          </tr>
        `).join('')}
      </table>
      
      ${received.length > 0 ? `
        <h3 style="color: #B45309;">⚠️ Received — Needs Dispatch</h3>
        <ul>${received.map(r => `<li>${r[0]} — ${r[5]} (Due: ${r[8]})</li>`).join('')}</ul>
      ` : ''}
      
      <h3 style="color: #374151;">Accounts Receivable</h3>
      <p><strong>Total Outstanding: $${totalOutstanding.toLocaleString('en-US', {minimumFractionDigits: 2})}</strong></p>
      <p>${unpaidInvoices.length} unpaid invoice(s)</p>
      
      <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 20px 0;">
      <p style="color: #9CA3AF; font-size: 12px;">
        <a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit">Open Operations Hub</a>
      </p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: `📊 Oneiro Daily Summary — ${today}`,
    htmlBody: html
  });
}


// ═══════════════════════════════════════════════════════════════
// 8. CUSTOM MENU (added to spreadsheet)
// ═══════════════════════════════════════════════════════════════

/**
 * One-off helper — run this from the Apps Script editor if the
 * "🔧 Oneiro Automation" menu isn't appearing in the spreadsheet.
 *
 * Standalone scripts occasionally lose the installable onOpen trigger
 * after an OAuth re-grant. The symptom is "From spreadsheet" no longer
 * appearing as an event source in the Triggers UI dropdown, because
 * the script's binding to the target spreadsheet was cleared.
 *
 * This function deletes any stale onOpen triggers, then creates a new
 * one programmatically via ScriptApp.newTrigger — which also
 * re-establishes the script→spreadsheet binding and prompts for OAuth
 * approval on first run.  Does NOT touch other triggers (scan inbox
 * poll, approved docs cron) or the folder structure.
 */
function reinstallMenuTrigger() {
  // Wipe any existing onOpen triggers on this project
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(t);
      removed += 1;
    }
  });

  // Recreate, bound to the configured spreadsheet
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
    .onOpen()
    .create();

  Logger.log(`✅ onOpen trigger reinstalled (removed ${removed} stale trigger(s)). Close + reopen the spreadsheet to see the menu.`);
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 Oneiro Automation')
    .addItem('📋 Generate Daily Documents (Today)', 'generateDailyDocumentsToday')
    .addItem('📋 Generate Daily Documents (Custom Date)...', 'promptGenerateDaily')
    .addSeparator()
    .addItem('✅ Validate Sign-In Data (Today)', 'validateSignInToday')
    .addSeparator()
    .addItem('📑 Generate Certified Payroll...', 'promptCertifiedPayroll')
    .addSeparator()
    .addItem('📧 Send Daily Summary Now', 'sendDailySummary')
    .addItem('🔍 Check Scan Inbox Now', 'checkScanInbox')
    .addSeparator()
    .addItem('⚙️ Run Initial Setup', 'setupAutomation')
    .addItem('⚙️ Set up Marking Items schema', 'setupMarkingItems')
    .addToUi();
}

function generateDailyDocumentsToday() {
  generateDailyDocuments();
  SpreadsheetApp.getUi().alert('✅ Daily documents generated! Check the "Needs Review" folder in Drive.');
}

function promptGenerateDaily() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter date (MM/DD/YYYY):');
  if (response.getSelectedButton() === ui.Button.OK) {
    generateDailyDocuments(response.getResponseText());
    ui.alert('✅ Documents generated for ' + response.getResponseText());
  }
}

function validateSignInToday() {
  const issues = validateSignInData();
  const ui = SpreadsheetApp.getUi();
  if (issues.length === 0) {
    ui.alert('✅ All sign-in data validates cleanly!');
  } else {
    ui.alert('⚠️ Issues Found:\n\n' + issues.join('\n'));
  }
}

function promptCertifiedPayroll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter week start date (Sunday, MM/DD/YYYY):');
  if (response.getSelectedButton() === ui.Button.OK) {
    const count = generateCertifiedPayroll(response.getResponseText());
    if (count > 0) {
      ui.alert(`✅ Certified payroll generated for ${count} contract group(s)!\n\nCheck:\n• "Docs Needing Review → Certified Payroll" in Drive for the filled PDF\n• "Certified Payroll Tracker" tab to cross-verify with payrollforconstruction.com`);
    } else {
      ui.alert('⚠️ No sign-in entries found for that week.\n\nCheck that Daily Sign-In Data has rows with dates in that range, then try again.\n\nSee the Apps Script Logs (Extensions → Apps Script → Executions) for details.');
    }
  }
}


// ═══════════════════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function getOrCreateSubfolder_(parent, name) {
  // Fast path: if the folder already exists, return it without taking
  // the lock. This is the common case and needs no synchronization.
  let folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();

  // Slow path: we think we need to create it. Take a script-wide lock
  // so concurrent doPost calls don't race to create duplicate folders
  // (happens when N photo uploads fire in parallel and the Photos
  // subfolder doesn't exist yet — all N check simultaneously, all N
  // see "no folder", all N create one).
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);  // up to 10s; folder creation is <100ms
  } catch (e) {
    // Couldn't acquire — fall through and risk the duplicate rather
    // than erroring out the caller. Better a double-create than a
    // broken upload.
    Logger.log('⚠️ getOrCreateSubfolder_: could not acquire lock for ' + name);
  }
  try {
    // Re-check inside the lock — another request may have created it
    // while we waited.
    folders = parent.getFoldersByName(name);
    if (folders.hasNext()) return folders.next();
    return parent.createFolder(name);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* never held */ }
  }
}

function getBoroughName_(code) {
  const map = { 'M': 'Manhattan', 'BX': 'Bronx', 'BK': 'Brooklyn', 'QU': 'Queens', 'SI': 'Staten Island' };
  return map[code] || code;
}

function getFolderId_(name) {
  return PropertiesService.getScriptProperties().getProperty(name) || CONFIG[name] || '';
}


/**
 * Sheet write with per-column diagnostics.
 *
 * Attempts a single batched setValues (fast path). If that fails —
 * typically because a cell has data validation that rejects the value —
 * retries column by column to identify exactly which column + value was
 * the culprit, then re-throws with a specific message.
 *
 * @param sheet       the Sheet
 * @param row         1-indexed row
 * @param startCol    1-indexed starting column
 * @param values      array of values to write (one per column)
 * @param labels      human-readable column names, parallel to values
 * @param sheetLabel  short label for the sheet (used in error messages)
 */
function writeRowWithProbing_(sheet, row, startCol, values, labels, sheetLabel) {
  try {
    sheet.getRange(row, startCol, 1, values.length).setValues([values]);
    // CRITICAL: Apps Script's setValues is DEFERRED — the write gets
    // buffered and validation fires on the next read (e.g. getLastRow()
    // on any sheet). Without flush(), a dropdown violation in this
    // write would escape this try/catch and surface as a bare "Invalid
    // Entry" error at some later unrelated sheet read. flush() forces
    // the write to commit now so any validation error lands in our
    // catch block where the probe can identify the culprit column.
    SpreadsheetApp.flush();
  } catch (batchErr) {
    // Probe per-column to identify which one failed
    for (let i = 0; i < values.length; i++) {
      try {
        sheet.getRange(row, startCol + i).setValue(values[i]);
        SpreadsheetApp.flush();  // per-cell flush for same reason as above
      } catch (colErr) {
        const label = labels[i] || `col ${startCol + i}`;
        const val   = JSON.stringify(values[i]);
        throw new Error(
          `${sheetLabel} → "${label}" rejected value ${val}. ` +
          `Cell validation: ${colErr.message}`
        );
      }
    }
    // Probe couldn't reproduce per-cell — still attach the sheet label so the
    // caller never sees a bare "Invalid Entry" message with no context.
    throw new Error(
      `${sheetLabel} → batch write failed (per-cell probe could not isolate). ` +
      `Original: ${batchErr.message}`
    );
  }
}


/**
 * Append a row, but if the batched append fails (e.g. a cell has dropdown
 * validation that rejected the value), probe per-column by writing cells
 * one at a time to identify which column was rejected, then re-throw with
 * a specific message. The probe row is deleted afterwards so we leave no
 * partial data behind.
 */
function appendRowWithProbing_(sheet, values, labels, sheetLabel) {
  // Catch-all safety wrapper — guarantees the sheet label + phase tag is
  // attached to any error that escapes the probe logic, so a caller never
  // sees a bare Google validation message without context.
  const ctx = { phase: 'init' };
  try {
    return appendRowWithProbingImpl_(sheet, values, labels, sheetLabel, ctx);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    if (msg.indexOf(sheetLabel) !== -1) throw err;  // already tagged
    throw new Error(`${sheetLabel} (phase=${ctx.phase}) → ${msg}`);
  }
}

function appendRowWithProbingImpl_(sheet, values, labels, sheetLabel, ctx) {
  ctx.phase = 'trim';
  // Trim trailing empty cells — defensive no-op with setValues, but kept
  // so the probe's range width matches what the caller actually intended.
  let trimLen = values.length;
  while (trimLen > 0 && (values[trimLen - 1] === '' || values[trimLen - 1] == null)) {
    trimLen--;
  }
  if (trimLen < values.length) values = values.slice(0, trimLen);

  ctx.phase = 'getLastRow';
  const targetRow = sheet.getLastRow() + 1;

  // Prefer setValues over appendRow. appendRow validates the FULL row
  // range against all data-validation rules — including dropdown columns
  // we aren't writing. setValues is scoped to the specific range we're
  // writing, so only those columns' validators run.
  ctx.phase = 'setValues';
  try {
    sheet.getRange(targetRow, 1, 1, values.length).setValues([values]);
    // Force deferred validation to fire now — see writeRowWithProbing_
    // for the full explanation. Without flush(), a dropdown violation
    // escapes this try and surfaces as a bare error at the next sheet
    // read somewhere else in the handler.
    SpreadsheetApp.flush();
    return;
  } catch (batchErr) {
    ctx.phase = 'per-cell probe';
    Logger.log(`❌ setValues threw on ${sheetLabel}: ${batchErr.message}`);
    // Keep the variable name so the rest of the probe/diagnostic code
    // below continues to make sense — it's still the "next row after
    // last data" we just tried to fill.
    let culprit = null;
    try {
      for (let i = 0; i < values.length; i++) {
        try {
          sheet.getRange(targetRow, i + 1).setValue(values[i]);
          SpreadsheetApp.flush();  // per-cell flush so validation fires now
        } catch (colErr) {
          const label = labels[i] || `col ${i + 1}`;
          const val   = JSON.stringify(values[i]);
          culprit = new Error(
            `${sheetLabel} → "${label}" rejected value ${val}. ` +
            `Cell validation: ${colErr.message}`
          );
          break;
        }
      }
    } finally {
      // Remove the probe row whether we identified the culprit or not —
      // we don't want partial data left behind on validation failure, and
      // we don't want a duplicate if the probe somehow succeeded.
      // IMPORTANT: wrap in its own try/catch — if deleteRow throws, its
      // exception would replace any pending `culprit` throw below and
      // masquerade as the submit's failure (we've seen exactly that).
      try {
        if (sheet.getLastRow() >= targetRow) {
          sheet.deleteRow(targetRow);
        }
      } catch (delErr) {
        Logger.log(`⚠️ deleteRow(probe) for ${sheetLabel} failed: ${delErr.message}`);
      }
    }
    if (culprit) throw culprit;

    ctx.phase = 'rule inspection';

    // Per-cell probe didn't reproduce. That happens when appendRow validates
    // the row atomically but per-cell setValue bypasses validation (e.g.
    // dropdown ranges applied to specific rows, cross-cell rules). Scan
    // validation rules on both row 2 and the intended append row, merge
    // them, and check every list/range rule against its incoming value.
    const DV = SpreadsheetApp.DataValidationCriteria;
    const mergedRules = new Array(values.length).fill(null);
    const inspectRows = [2, targetRow, Math.max(targetRow - 1, 3)];
    inspectRows.forEach(rowNum => {
      if (rowNum < 2) return;
      try {
        const rules = sheet.getRange(rowNum, 1, 1, values.length).getDataValidations()[0];
        for (let i = 0; i < rules.length; i++) {
          if (!mergedRules[i] && rules[i]) mergedRules[i] = rules[i];
        }
      } catch (err) { /* ignore row read errors */ }
    });

    // Check each value against list/range rules — this is the deterministic
    // identifier when per-cell probe can't reproduce.
    for (let i = 0; i < mergedRules.length; i++) {
      const rule = mergedRules[i];
      if (!rule) continue;
      const t       = rule.getCriteriaType();
      const crArgs  = rule.getCriteriaValues() || [];
      let allowed   = null;
      if (t === DV.VALUE_IN_LIST) {
        allowed = (Array.isArray(crArgs[0]) ? crArgs[0] : [crArgs[0]]).map(String);
      } else if (t === DV.VALUE_IN_RANGE && crArgs[0]) {
        try { allowed = crArgs[0].getValues().flat().filter(v => v !== '').map(String); }
        catch (err) { allowed = null; }
      }
      if (!allowed) continue;
      const raw   = values[i];
      const asStr = (raw == null) ? '' : String(raw);
      if (asStr === '') continue;
      if (allowed.indexOf(asStr) === -1) {
        throw new Error(
          `${sheetLabel} → "${labels[i] || 'col ' + (i + 1)}" value ` +
          `${JSON.stringify(raw)} not in allowed list [${allowed.join(', ')}].`
        );
      }
    }

    // Still couldn't isolate. Summarize what we DID find on the sheet so
    // the next paste-back tells us which columns actually have validation
    // (and of what type) — that narrows it down even when the rule type
    // isn't a plain list/range dropdown.
    const ruleSummary = mergedRules.map((r, i) => {
      if (!r) return null;
      const t    = r.getCriteriaType();
      const args = r.getCriteriaValues() || [];
      let desc   = String(t || 'UNKNOWN');
      if (t === DV.VALUE_IN_LIST) {
        const items = Array.isArray(args[0]) ? args[0] : [args[0]];
        desc += ` [${items.map(String).join(',')}]`;
      } else if (t === DV.VALUE_IN_RANGE && args[0]) {
        try { desc += ` [range ${args[0].getA1Notation()}]`; } catch (_) {}
      }
      return `${labels[i]}: ${desc}`;
    }).filter(Boolean);

    const summary = values
      .map((v, i) => `${labels[i] || 'col' + (i + 1)}=${JSON.stringify(v)}`)
      .join(' | ');
    throw new Error(
      `${sheetLabel} → batch append failed. Rules found: {${ruleSummary.join(' ; ')}}. ` +
      `Row values: [${summary}]. Original: ${batchErr.message}`
    );
  }
}


/**
 * Append many rows to a sheet in a single setValues call (one
 * round-trip + one flush instead of N). Falls back to per-row
 * appendRowWithProbing_ if the batch fails, preserving the full
 * probe-and-diagnose behavior for validation errors.
 *
 * All rows must be the same length (will be right-padded with '' if
 * not). Use this whenever the handler appends 2+ rows with the same
 * shape — typical savings: 1-2 seconds per additional row on crowded
 * sheets.
 */
function appendRowsWithProbing_(sheet, rows, labels, sheetLabel) {
  if (!rows || rows.length === 0) return;
  if (rows.length === 1) {
    appendRowWithProbing_(sheet, rows[0], labels, sheetLabel);
    return;
  }

  // Pad to a consistent width so setValues accepts them
  const width = Math.max(...rows.map(r => r.length));
  const padded = rows.map(r => {
    if (r.length === width) return r.slice();
    const copy = r.slice();
    while (copy.length < width) copy.push('');
    return copy;
  });

  const targetRow = sheet.getLastRow() + 1;
  try {
    sheet.getRange(targetRow, 1, padded.length, width).setValues(padded);
    SpreadsheetApp.flush();   // force deferred validation to fire now
    return;
  } catch (batchErr) {
    Logger.log(`⚠️ Batched setValues failed on ${sheetLabel} (${rows.length} rows): `
               + `${batchErr.message}. Falling back to per-row probing.`);
    // Clean up any partial batch that may have been written before the error
    try {
      const lastWritten = sheet.getLastRow();
      if (lastWritten >= targetRow) {
        sheet.getRange(targetRow, 1, lastWritten - targetRow + 1, width).clearContent();
      }
    } catch (cleanupErr) {
      Logger.log(`⚠️ post-batch cleanup failed: ${cleanupErr.message}`);
    }
    rows.forEach(row => appendRowWithProbing_(sheet, row, labels, sheetLabel));
  }
}


// ═══════════════════════════════════════════════════════════════
// UPLOAD PROXY — Receives filled PDFs from the Railway worker
// ═══════════════════════════════════════════════════════════════
//
// WHY THIS EXISTS:
//   Google Drive service accounts have no personal storage quota,
//   so they cannot create files in regular (non-Shared) Drive folders.
//   The Railway worker fills PDFs but cannot upload them directly.
//   Instead it POSTs the bytes here, and Apps Script saves the file
//   as the real Drive-owning Google account — no quota issues ever.
//
// ONE-TIME SETUP (do this after deploying the Apps Script project):
//   Step 1 — Set the upload secret:
//     Extensions → Apps Script → Project Settings → Script Properties
//     Add:  UPLOAD_SECRET = <any strong random string you choose>
//
//   Step 2 — Deploy as a Web App:
//     Deploy → New deployment → Type: Web App
//     Execute as:     Me  (your Google account)
//     Who has access: Anyone
//     → Copy the deployment URL
//
//   Step 3 — Add two env vars to Railway:
//     APPS_SCRIPT_UPLOAD_URL = <the deployment URL from Step 2>
//     APPS_SCRIPT_UPLOAD_KEY = <the same secret from Step 1>

/**
 * HTTP POST handler — called by the Railway worker to save a filled PDF.
 *
 * Request body (JSON):
 *   { key, filename, folder_id, data }   ← data is base64-encoded PDF bytes
 *
 * Response JSON:
 *   { success: true, file_id, file_url, filename }
 *   { error: "...", _status: 400|401|500 }
 */
/**
 * doPost — central HTTP handler for Railway worker requests.
 *
 * Dispatches on body.action:
 *   "upload_pdf"  — save a filled PDF to a Drive folder (original proxy)
 *   "write_wo"    — write a parsed Work Order to the WO Tracker + archive PDF
 *
 * All requests must include body.key matching the UPLOAD_SECRET Script Property.
 */
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('UPLOAD_SECRET');

    if (!secret || body.key !== secret) {
      return jsonResponse_({ error: 'unauthorized' }, 401);
    }

    const action = body.action || 'upload_pdf';

    if (action === 'upload_pdf') {
      return handleUploadPdf_(body);
    } else if (action === 'write_wo') {
      return handleWriteWO_(body);
    } else if (action === 'get_active_wos') {
      return handleGetActiveWOs_();
    } else if (action === 'submit_field_report') {
      return handleSubmitFieldReport_(body);
    } else if (action === 'check_fr_shift_attribution') {
      return handleCheckFrShiftAttribution_(body);
    } else if (action === 'finalize_field_report_docs') {
      return handleFinalizeFieldReportDocs_(body);
    } else if (action === 'get_dashboard_data') {
      return handleGetDashboardData_();
    } else if (action === 'get_revenue_data') {
      return handleGetRevenueData_(body);
    } else if (action === 'get_production_data') {
      return handleGetProductionData_(body);
    } else if (action === 'upload_photo') {
      return handleUploadPhoto_(body);
    } else if (action === 'list_wo_photos') {
      return handleListWOPhotos_(body);
    } else if (action === 'get_wo_photo_content') {
      return handleGetWOPhotoContent_(body);
    } else if (action === 'reverse_geocode') {
      return handleReverseGeocode_(body);
    } else if (action === 'upload_signature') {
      return handleUploadSignature_(body);
    } else if (action === 'get_marking_items') {
      return handleGetMarkingItems_(body);
    } else if (action === 'create_marking_item') {
      return handleCreateMarkingItem_(body);
    } else if (action === 'update_marking_item') {
      return handleUpdateMarkingItem_(body);
    } else if (action === 'delete_marking_items') {
      return handleDeleteMarkingItems_(body);
    } else if (action === 'trash_file') {
      return handleTrashFile_(body);
    } else if (action === 'upload_wo_scan') {
      return handleUploadWOScan_(body);
    } else if (action === 'get_scan_status') {
      return handleGetScanStatus_(body);
    } else if (action === 'get_scan_uploads_today') {
      return handleGetScanUploadsToday_(body);
    } else if (action === 'log_wo_scan_failure') {
      return handleLogWOScanFailure_(body);
    } else if (action === 'list_pending_approvals') {
      return handleListPendingApprovals_(body);
    } else if (action === 'get_pending_counts') {
      return handleGetPendingCounts_(body);
    } else if (action === 'get_doc_status_pending_count') {
      return handleGetDocStatusPendingCount_(body);
    } else if (action === 'update_wo_status') {
      return handleUpdateWOStatus_(body);
    } else if (action === 'delete_wo') {
      return handleDeleteWO_(body);
    } else if (action === 'edit_completed_wo') {
      return handleEditCompletedWO_(body);
    } else if (action === 'update_wo_coordinates') {
      return handleUpdateWOCoordinates_(body);
    } else if (action === 'get_wo_map_data') {
      return handleGetWOMapData_(body);
    } else if (action === 'get_drive_file_bytes') {
      return handleGetDriveFileBytes_(body);
    } else if (action === 'approve_doc') {
      return handleApproveDoc_(body);
    } else if (action === 'approve_signin_with_bytes') {
      return handleApproveSignInWithBytes_(body);
    } else if (action === 'set_waterblast_confirmed') {
      return handleSetWaterblastConfirmed_(body);
    } else if (action === 'generate_daily_documents') {
      return handleGenerateDailyDocuments_(body);
    } else if (action === 'generate_certified_payroll') {
      return handleGenerateCertifiedPayroll_(body);
    } else if (action === 'generate_pl_for_doc') {
      return handleGeneratePLForDoc_(body);
    } else if (action === 'generate_cp_for_doc') {
      return handleGenerateCPForDoc_(body);
    } else if (action === 'process_approved_documents') {
      return handleProcessApprovedDocuments_(body);
    } else if (action === 'list_employees') {
      return handleListEmployees_(body);
    } else if (action === 'list_signin_queue') {
      return handleListSignInQueue_(body);
    } else if (action === 'submit_signin') {
      return handleSubmitSignIn_(body);
    } else if (action === 'check_signin_continuation') {
      return handleCheckSignInContinuation_(body);
    } else if (action === 'list_signin_day_hours') {
      return handleListSignInDayHours_(body);
    } else if (action === 'list_signin_rows_for_file') {
      return handleListSignInRowsForFile_(body);
    } else if (action === 'save_signin_row_edits') {
      return handleSaveSignInRowEdits_(body);
    } else if (action === 'approve_doc_skip_signoff') {
      return handleApproveDocSkipSignoff_(body);
    } else if (action === 'lookup_archived_wo_pdf') {
      return handleLookupArchivedWOPdf_(body);
    } else if (action === 'list_metro_completed_docs') {
      return handleListMetroCompletedDocs_(body);
    } else if (action === 'set_docs_sent') {
      return handleSetDocsSent_(body);
    } else if (action === 'set_doc_status') {
      return handleSetDocStatus_(body);
    } else if (action === 'get_doc_status_calendar') {
      return handleGetDocStatusCalendar_(body);
    } else if (action === 'list_documents_for_batch') {
      return handleListDocumentsForBatch_(body);
    } else if (action === 'get_qb_invoice_payload') {
      return handleGetQbInvoicePayload_(body);
    } else if (action === 'record_qb_invoice') {
      return handleRecordQbInvoice_(body);
    } else if (action === 'clear_qb_invoice') {
      return handleClearQbInvoice_(body);
    } else if (action === 'get_qb_refresh_token') {
      return handleGetQbRefreshToken_();
    } else if (action === 'set_qb_refresh_token') {
      return handleSetQbRefreshToken_(body);
    } else if (action === 'get_qb_customer_id') {
      return handleGetQbCustomerId_(body);
    } else if (action === 'set_qb_customer_id') {
      return handleSetQbCustomerId_(body);
    } else {
      return jsonResponse_({ error: 'Unknown action: ' + action }, 400);
    }

  } catch (err) {
    Logger.log('❌ doPost error: ' + err.toString());
    if (err.stack) Logger.log('Stack trace:\n' + err.stack);
    // Include the stack in the JSON response so the React app / network
    // inspector can show it without needing access to Apps Script Executions
    // (whose list entries aren't clickable in the current UI).
    return jsonResponse_({ error: err.toString(), stack: err.stack || '' }, 500);
  }
}


// ── action: upload_pdf ────────────────────────────────────────────────────────

function handleUploadPdf_(body) {
  const fileName = body.filename;
  const folderId = body.folder_id;
  const encoded  = body.data;

  if (!fileName || !folderId || !encoded) {
    return jsonResponse_({ error: 'Missing required fields: filename, folder_id, data' }, 400);
  }

  const pdfBytes = Utilities.base64Decode(encoded);
  const blob     = Utilities.newBlob(pdfBytes, 'application/pdf', fileName);
  const folder   = DriveApp.getFolderById(folderId);
  const file     = folder.createFile(blob);

  Logger.log('📄 Upload proxy saved: ' + file.getName() + ' → folder ' + folderId);

  return jsonResponse_({
    success:  true,
    file_id:  file.getId(),
    file_url: file.getUrl(),
    filename: file.getName()
  });
}


// ── action: upload_wo_scan ───────────────────────────────────────
//
// Webapp-initiated WO scan upload. Writes the file into the Drive
// "Scan Inbox" folder so the existing Railway watcher picks it up,
// parses with Claude Vision, and calls write_wo exactly like a
// drag-dropped file would. Archive path + all downstream behavior
// (tracker row, Marking Items seed, archive move, Automation Log)
// is unchanged.
//
// body.data:
//   filename  — original file name (e.g. "RM-43281 scan.pdf")
//   mime_type — MIME type (e.g. "application/pdf", "image/jpeg")
//   data      — base64-encoded file bytes

function handleUploadWOScan_(body) {
  const d = body.data || {};
  const { filename, mime_type, data } = d;
  if (!filename || !data) {
    return jsonResponse_({ error: 'Missing required fields: filename, data' }, 400);
  }

  const props = PropertiesService.getScriptProperties();
  const scanInboxId = props.getProperty('SCAN_INBOX_ID');
  if (!scanInboxId) {
    return jsonResponse_({ error: 'SCAN_INBOX_ID not configured in Script Properties' }, 500);
  }

  const bytes = Utilities.base64Decode(data);
  const blob  = Utilities.newBlob(bytes, mime_type || 'application/pdf', filename);
  const folder = DriveApp.getFolderById(scanInboxId);
  const file   = folder.createFile(blob);

  Logger.log('📥 WO scan uploaded via webapp: ' + file.getName() + ' (' + file.getId() + ')');
  return jsonResponse_({
    success:  true,
    file_id:  file.getId(),
    filename: file.getName(),
    url:      file.getUrl()
  });
}


// ── action: get_scan_status ───────────────────────────────────

/**
 * Returns per-file status for WO scans the webapp uploaded. Called by
 * the Scan WO page's poll loop to transition each item from
 * "Parsing…" → "Done (RM-xxx)" or "Error".
 *
 * body.data = { file_ids: [<string>, ...] }
 *
 * For each file_id, response entry shape:
 *   { file_id,
 *     status: 'done' | 'pending' | 'error' | 'unknown',
 *     wo_ids?: [<WO#>, ...]   // present when status='done'
 *     message?: <string>       // present when status='error'
 *   }
 *
 * Matching rules:
 *   done    — WO Tracker has at least one row where
 *             col 38 (Scan File ID) === file_id   OR
 *             col 39 (Combined Scan File ID) === file_id
 *   error   — Automation Log has a row where Source='WO Scan' AND
 *             Action='Parse failed' AND Details contains the file_id
 *   pending — file is still present in the Scan Inbox folder
 *   unknown — file absent from Scan Inbox, no tracker match, no error
 *             log (rare — treat as error in the UI)
 *
 * Sheets are read once and reused for the whole batch; Drive lookup
 * is scoped to the Scan Inbox folder only (1 query per file_id).
 */
function handleGetScanStatus_(body) {
  const d = body.data || {};
  const fileIds = Array.isArray(d.file_ids) ? d.file_ids.filter(Boolean).map(String) : [];
  if (fileIds.length === 0) return jsonResponse_({ statuses: [] });

  const ss       = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet  = ss.getSheetByName('Work Order Tracker');
  const woData   = woSheet.getDataRange().getValues();
  const logSheet = ss.getSheetByName('Automation Log');
  const logData  = logSheet ? logSheet.getDataRange().getValues() : [];

  // Index WO Tracker rows by col 38 (Scan File ID) and col 39 (Combined).
  const byScanId     = {};   // file_id → [wo_id, ...]
  const byCombinedId = {};
  woData.slice(1).forEach(r => {
    const woId     = String(r[0]  || '').trim();
    const scanId   = String(r[38] || '').trim();
    const combined = String(r[39] || '').trim();
    if (!woId) return;
    if (scanId)   (byScanId[scanId]       = byScanId[scanId]       || []).push(woId);
    if (combined) (byCombinedId[combined] = byCombinedId[combined] || []).push(woId);
  });

  // Index Automation Log scan failures by file_id (substring match in Details col).
  // Only consider the most-recent failure per file_id so a post-retry success
  // doesn't keep surfacing the old error.
  const errorsByFileId = {};
  logData.slice(1).forEach(r => {
    const source  = String(r[1] || '').trim();
    const action  = String(r[2] || '').trim();
    const details = String(r[4] || '');
    if (source !== 'WO Scan' || action !== 'Parse failed') return;
    fileIds.forEach(fid => {
      if (details.indexOf(fid) !== -1) errorsByFileId[fid] = details;
    });
  });

  // Look up Scan Inbox presence once (one query per file_id).
  const props       = PropertiesService.getScriptProperties();
  const scanInboxId = props.getProperty('SCAN_INBOX_ID');
  const stillInInbox = {};
  if (scanInboxId) {
    fileIds.forEach(fid => {
      try {
        const f = DriveApp.getFileById(fid);
        if (!f.isTrashed()) {
          // Confirm the file is in Scan Inbox (not already archived)
          const parents = f.getParents();
          while (parents.hasNext()) {
            if (parents.next().getId() === scanInboxId) {
              stillInInbox[fid] = true;
              break;
            }
          }
        }
      } catch (e) {
        // Not found or inaccessible → falls through to unknown/error
      }
    });
  }

  const statuses = fileIds.map(fid => {
    const direct   = byScanId[fid]     || [];
    const combined = byCombinedId[fid] || [];
    const ids      = direct.concat(combined);
    if (ids.length > 0) {
      return { file_id: fid, status: 'done', wo_ids: ids };
    }
    if (errorsByFileId[fid]) {
      return { file_id: fid, status: 'error', message: errorsByFileId[fid] };
    }
    if (stillInInbox[fid]) {
      return { file_id: fid, status: 'pending' };
    }
    return { file_id: fid, status: 'unknown' };
  });

  return jsonResponse_({ statuses });
}


// ── action: get_scan_uploads_today ────────────────────────────

/**
 * Returns today's scan-originated WO Tracker rows, grouped by the
 * upload that produced them. The Scan WO page uses this as its
 * source of truth so the queue reflects the tracker (cross-device,
 * survives browser clears, respects admin deletions).
 *
 * Grouping key = col 39 (Combined Scan File ID) if set, else col 38
 * (Scan File ID). All splits from one multi-WO upload share the
 * same Combined ID, so they roll up to one queue item.
 *
 * Response shape:
 *   { uploads: [{
 *       file_id:     <grouping key — what the webapp originally uploaded>,
 *       filename:    <Original Filename from col 41, or the WO# as fallback>,
 *       uploaded_at: <ISO timestamp from col 40>,
 *       is_combined: <true if this upload produced multiple tracker rows>,
 *       wo_ids:      [<WO#>, ...]   // sorted by WO#
 *     }, ...]
 *   }
 *
 * "Today" is defined in the spreadsheet's timezone (not UTC) so the
 * boundary matches what the admin sees in the sheet.
 */
function handleGetScanUploadsToday_(body) {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData  = woSheet.getDataRange().getValues();

  const tz    = ss.getSpreadsheetTimeZone() || CONFIG.TIMEZONE;
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const asDate = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    if (v) {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  };

  const groups = {};  // groupKey → aggregated shape
  woData.slice(1).forEach(r => {
    const woId       = String(r[0]  || '').trim();
    const scanId     = String(r[38] || '').trim();
    const combinedId = String(r[39] || '').trim();
    if (!woId) return;
    if (!scanId && !combinedId) return;   // not a scan-originated row

    const tsRaw   = r[40];
    const tsDate  = asDate(tsRaw);
    if (!tsDate) return;
    if (Utilities.formatDate(tsDate, tz, 'yyyy-MM-dd') !== today) return;

    const groupKey = combinedId || scanId;
    const filename = String(r[41] || '').trim();
    if (!groups[groupKey]) {
      groups[groupKey] = {
        file_id:     groupKey,
        filename:    filename || woId,
        uploaded_at: tsDate.toISOString(),
        is_combined: !!combinedId,
        wo_ids:      [],
      };
    } else {
      // Keep earliest upload timestamp in the group (closest to when
      // the user actually hit "Upload")
      if (tsDate.toISOString() < groups[groupKey].uploaded_at) {
        groups[groupKey].uploaded_at = tsDate.toISOString();
      }
    }
    groups[groupKey].wo_ids.push(woId);
  });

  // Sort wo_ids within each group, and sort uploads by most-recent first
  const uploads = Object.values(groups).map(g => ({
    ...g,
    wo_ids: g.wo_ids.slice().sort(),
  })).sort((a, b) => a.uploaded_at < b.uploaded_at ? 1 : -1);

  return jsonResponse_({ uploads });
}


// ═══════════════════════════════════════════════════════════════
// DOCUMENT APPROVALS — webapp "Approvals" tab backend
// ═══════════════════════════════════════════════════════════════

// Map webapp doc_type → Needs-Review subfolder name.
const APPROVAL_SUBFOLDERS_ = {
  signin:            'Sign-In Logs',
  production_log:    'Production Logs',
  field_report:      'Field Reports',
  certified_payroll: 'Certified Payroll',
};

// Inverse lookup helper — given a subfolder name, return the doc_type.
function _docTypeFromSubfolderName_(name) {
  for (const key of Object.keys(APPROVAL_SUBFOLDERS_)) {
    if (APPROVAL_SUBFOLDERS_[key] === name) return key;
  }
  return null;
}

// Best-effort subtitle extraction from filename. Returns a short string
// the webapp list shows under the doc_type badge.
function _approvalSubtitleFromFilename_(docType, filename) {
  const s = String(filename || '');
  // WO # pattern, used by signin + CFR + any single-WO doc
  const woMatch = s.match(/(PT|PM|RM)-\d+/i);
  if (docType === 'signin' || docType === 'field_report') {
    return woMatch ? woMatch[0] : s.replace(/\.pdf$/i, '');
  }
  if (docType === 'production_log') {
    // e.g. "Production_Log_2026-04-21_FILLED.pdf" → "2026-04-21"
    const d = s.match(/\d{4}-\d{2}-\d{2}/);
    return d ? d[0] : s.replace(/\.pdf$/i, '');
  }
  if (docType === 'certified_payroll') {
    // e.g. "Certified_Payroll_84125MBTP701_BK_2026-04-21_FILLED.pdf"
    // Or with billing-remap split: "..._2026-04-21_MetroExpress_FILLED.pdf"
    const m = s.match(/Certified_Payroll_([^_]+)_([^_]+)_(\d{4}-\d{2}-\d{2})(?:_([A-Za-z0-9]+))?/);
    if (m) {
      const [, cn, bor, date, suffix] = m;
      const isMarker = suffix === 'FILLED' || suffix === 'MANUAL';
      const tag = (suffix && !isMarker) ? ` (${suffix})` : '';
      return `${cn}-${bor}${tag} · ${date}`;
    }
    return s.replace(/\.pdf$/i, '');
  }
  return s.replace(/\.pdf$/i, '');
}


// ── action: list_pending_approvals ────────────────────────────
//
// Returns every PDF currently sitting in Docs Needing Review's
// four subfolders (Sign-In Logs, Production Logs, Field Reports,
// Certified Payroll). Sorted newest-first. Webapp's Approvals page
// uses this as its master list.
function handleListPendingApprovals_(body) {
  const props          = PropertiesService.getScriptProperties();
  const needsReviewId  = props.getProperty('NEEDS_REVIEW_ID');
  if (!needsReviewId) return jsonResponse_({ error: 'NEEDS_REVIEW_ID not set' }, 500);

  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  const approvals    = [];

  Object.entries(APPROVAL_SUBFOLDERS_).forEach(([docType, subName]) => {
    const sub = reviewFolder.getFoldersByName(subName);
    if (!sub.hasNext()) return;
    const subFolder = sub.next();
    const files     = subFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      // Only PDFs — we intentionally ignore any .json source files that
      // may still be in-flight. Worker trashes them after the filler
      // runs, but be defensive.
      if (f.getMimeType() !== 'application/pdf') continue;
      const filename = f.getName();
      approvals.push({
        file_id:    f.getId(),
        doc_type:   docType,
        filename:   filename,
        subtitle:   _approvalSubtitleFromFilename_(docType, filename),
        created_at: f.getDateCreated().toISOString(),
        size:       f.getSize(),
      });
    }
  });

  // FIFO: oldest first so the longest-waiting reviews surface at the top.
  approvals.sort((a, b) => a.created_at < b.created_at ? -1 : 1);
  // approved_docs_pending = files sitting in the Approved Docs folder
  // waiting for processApprovedDocuments to pick them up. Excludes the
  // 📨-prefixed already-sent rows. The Approvals page surfaces this number
  // next to the "Process Approved Docs Now" worker button.
  let approved_docs_pending = null;
  try { approved_docs_pending = _countApprovedDocsPending_(); } catch (e) {}
  return jsonResponse_({ approvals, approved_docs_pending });
}


// ── helpers: pending-count primitives ─────────────────────────
//
// Each helper does one focused count, reusing the same iteration the
// matching list_* action does. Wrapped individually so a misconfigured
// prop in one folder doesn't sink the whole counts response.

// Count PDFs sitting in the four Docs Needing Review subfolders. Same
// scoping as handleListPendingApprovals_ but skips the per-file payload
// build for speed.
function _countDocsNeedingReview_() {
  const props = PropertiesService.getScriptProperties();
  const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
  if (!needsReviewId) throw new Error('NEEDS_REVIEW_ID not set');
  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  let n = 0;
  Object.entries(APPROVAL_SUBFOLDERS_).forEach(([_docType, subName]) => {
    const sub = reviewFolder.getFoldersByName(subName);
    if (!sub.hasNext()) return;
    const files = sub.next().getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      if (f.getMimeType() !== 'application/pdf') continue;
      n++;
    }
  });
  return n;
}

// Count files in the Approved Docs folder waiting for the worker. The
// 📨 prefix marks already-sent files (processApprovedDocuments rewrites
// the name on success), so they're excluded.
function _countApprovedDocsPending_() {
  const props = PropertiesService.getScriptProperties();
  const approvedId = props.getProperty('APPROVED_SENT_ID');
  if (!approvedId) throw new Error('APPROVED_SENT_ID not set');
  const folder = DriveApp.getFolderById(approvedId);
  let n = 0;
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.isTrashed()) continue;
    if (f.getName().startsWith('📨')) continue;
    n++;
  }
  return n;
}

// Count outstanding (date, contract, borough) sign-in groups from Work
// Day Log. Same Pending-status filter and grouping as
// handleListSignInQueue_ — just returns the group count, not the full
// payload.
function _countPendingSignins_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Work Day Log');
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  const seen = new Set();
  data.slice(1).forEach(row => {
    if (String(row[7] || '').trim() !== 'Pending') return;
    const dateStr     = (row[0] instanceof Date && !isNaN(row[0].getTime()))
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (String(row[0] || '').match(/^(\d{4}-\d{2}-\d{2})/) || [])[1] || '';
    const contractNum = String(row[3] || '').split('/')[0].trim();
    const borough     = String(row[4] || '').trim();
    if (!dateStr || !contractNum || !borough) return;
    seen.add(`${dateStr}|${contractNum}|${borough}`);
  });
  return seen.size;
}


// ── action: get_pending_counts ────────────────────────────────
//
// Cold-start endpoint for the webapp's nav badges. Returns three cheap
// counts in one round-trip so a fresh visitor on any non-queue page
// (Field Report, Scan WO, etc.) still sees how much work is queued.
//
// Intentionally OMITS the Doc Status pending count — that requires
// _buildDocStatusPayload_ which is expensive. The Doc Status tab badge
// instead populates from the regular /api/doc-status fetch when the
// user actually visits Dashboard.
//
// Each count is wrapped in try/catch so a misconfigured Drive prop only
// nulls out the matching field, not the whole response. UI hides any
// `null` field.
function handleGetPendingCounts_(_body) {
  let approvals_review = null;
  let approved_docs_pending = null;
  let signins_pending = null;
  try { approvals_review      = _countDocsNeedingReview_(); }   catch (e) {}
  try { approved_docs_pending = _countApprovedDocsPending_(); } catch (e) {}
  try { signins_pending       = _countPendingSignins_(); }      catch (e) {}
  return jsonResponse_({
    approvals_review,
    approved_docs_pending,
    signins_pending,
  });
}


// ── action: get_doc_status_pending_count ──────────────────────
//
// Heavier cousin of get_pending_counts split out so cold-start can
// fetch it in parallel from the webapp without making the fast counts
// wait. Builds the same payload _buildDocStatusPayload_ produces for
// the current month and returns just the pending-list length. The
// pending list itself is all-time (oldest first); the month arg only
// scopes the calendar cells, which we discard.
function handleGetDocStatusPendingCount_(_body) {
  try {
    const monthIso = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');
    const payload  = _buildDocStatusPayload_(monthIso);
    const n = Array.isArray(payload && payload.pending) ? payload.pending.length : 0;
    return jsonResponse_({ doc_status_pending: n });
  } catch (e) {
    Logger.log('get_doc_status_pending_count error: ' + e.message);
    return jsonResponse_({ doc_status_pending: null, error: e.message });
  }
}


// ── action: get_drive_file_bytes ──────────────────────────────
//
// Returns a Drive file's raw bytes as base64 so the webapp can pipe
// it into react-pdf. Scoped to files inside NEEDS_REVIEW_ID — admin
// shouldn't be able to download arbitrary Drive files via this.
function handleGetDriveFileBytes_(body) {
  const d = body.data || {};
  const fileId = String(d.file_id || '').trim();
  if (!fileId) return jsonResponse_({ error: 'Missing file_id' }, 400);

  // Trust model: file_ids reach this handler only after coming through
  // list_pending_approvals, which already scoped its scan to the four
  // Docs Needing Review subfolders. We previously walked the file's
  // parents tree to re-verify that, but on Shared Drives that walk
  // throws "Service error: Drive" repeatedly even on files that are
  // visibly in the right folder. Drop the redundant check.

  let step = 'init';
  try {
    step = 'getFileById';
    const file = _withDriveRetry_('getFileById preview', () => DriveApp.getFileById(fileId));

    step = 'isTrashed';
    if (file.isTrashed()) return jsonResponse_({ error: 'File is trashed' }, 404);

    step = 'getBlob/getBytes';
    const result = _withDriveRetry_('getBlob preview', () => {
      const blob = file.getBlob();
      return {
        filename:  file.getName(),
        mime_type: blob.getContentType(),
        size:      file.getSize(),
        data:      Utilities.base64Encode(blob.getBytes()),
      };
    });
    return jsonResponse_(result);
  } catch (err) {
    Logger.log(`❌ handleGetDriveFileBytes_ failed at step=${step}: ${err}\n${err && err.stack || ''}`);
    return jsonResponse_({ error: `[step=${step}] ${err && err.message || err}` }, 500);
  }
}

// Helper: is `file` anywhere under the folder with id `ancestorId`?
// Walks up via getParents(). Cheap since Drive folders rarely have
// deep nesting here (Docs Needing Review → 1 subfolder → file).
function _isUnderParent_(file, ancestorId) {
  if (!ancestorId) return false;
  const seen = {};
  let frontier = [];
  const parents = file.getParents();
  while (parents.hasNext()) frontier.push(parents.next());
  while (frontier.length) {
    const f = frontier.shift();
    const id = f.getId();
    if (id === ancestorId) return true;
    if (seen[id]) continue;
    seen[id] = true;
    const p = f.getParents();
    while (p.hasNext()) frontier.push(p.next());
  }
  return false;
}

/**
 * Retry a Drive op up to N times with linear backoff. The Apps Script
 * Drive service throws transient "Service error: Drive" exceptions for
 * eventual-consistency races (e.g. moveTo on a file that was created
 * milliseconds ago) and brief backend hiccups — both are retryable.
 * `label` shows up in Logger.log for diagnosis.
 */
function _withDriveRetry_(label, fn, attempts) {
  const maxAttempts = attempts || 3;
  let lastErr;
  for (let i = 0; i < maxAttempts; i++) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      const msg = String(err && err.message || err);
      Logger.log(`⚠️ Drive op "${label}" attempt ${i + 1}/${maxAttempts} failed: ${msg}`);
      if (i < maxAttempts - 1) Utilities.sleep(500 * (i + 1));
    }
  }
  throw lastErr;
}


// ── action: approve_doc ───────────────────────────────────────
//
// Moves a pending-approval PDF from Docs Needing Review/{type} into
// Approved Docs. The existing processApprovedDocuments cron sees the
// file on its next tick and handles email + archive — same path as if
// the admin had dragged it in Drive manually.
function handleApproveDoc_(body) {
  const d = body.data || {};
  const fileId = String(d.file_id || '').trim();
  if (!fileId) return jsonResponse_({ error: 'Missing file_id' }, 400);

  const props      = PropertiesService.getScriptProperties();
  const approvedId = props.getProperty('APPROVED_SENT_ID');
  if (!approvedId) return jsonResponse_({ error: 'APPROVED_SENT_ID not set' }, 500);

  // Trust model: file_ids reach approve handlers only after coming
  // through list_pending_approvals (the only way the webapp knows
  // them), which scopes its scan to Docs Needing Review subfolders.
  // The previous _isUnderParent_ re-check threw "Service error: Drive"
  // on Shared Drives because the upward parent walk is unreliable
  // there. Drop the redundant check.

  let step = 'init';
  try {
    step = 'getFileById';
    const file = _withDriveRetry_('getFileById', () => DriveApp.getFileById(fileId));

    step = 'isTrashed';
    if (file.isTrashed()) return jsonResponse_({ error: 'File is trashed' }, 404);

    step = 'moveTo Approved Docs';
    _withDriveRetry_('moveTo approve', () => {
      file.moveTo(DriveApp.getFolderById(approvedId));
    });

    // Log it — processApprovedDocuments will log its own "emailed" row
    // 0-10 min later when it picks this up.
    try {
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        .getSheetByName('Automation Log')
        .appendRow([
          new Date(), 'Approvals', 'Approved', file.getName(),
          'Moved to ✅ Approved Docs via webapp', 'Pending email',
          '', 'Cron will email + archive within 10 min'
        ]);
    } catch (logErr) {
      Logger.log('⚠️ Automation Log write failed on approve: ' + logErr);
    }

    return jsonResponse_({ success: true, file_id: fileId });
  } catch (err) {
    Logger.log(`❌ handleApproveDoc_ failed at step=${step}: ${err}\n${err && err.stack || ''}`);
    return jsonResponse_({ error: `[step=${step}] ${err && err.message || err}` }, 500);
  }
}


// ── action: approve_signin_with_bytes ─────────────────────────
//
// Doc-type-agnostic "approve with patched bytes" handler: takes a PDF
// already patched server-side (signature image + flattened text fields)
// and atomically uploads it to Approved Docs while trashing the original.
//
// Used by both:
//   - /api/approvals/:fileId/approve-signin       (Sign-In sheets)
//   - /api/approvals/:fileId/approve-cert-payroll (Certified Payroll)
//
// Express patches the PDF via pdf-lib and POSTs the resulting bytes
// here. This handler does upload + trash + log in one atomic move so
// the cron sees the signed PDF immediately. The legacy name is kept
// for backward compatibility — it's the "signed bytes" flow regardless
// of the source doc type.
//
// body.data = { file_id, filename, bytes_b64 }
//   file_id   — the ORIGINAL (unsigned) PDF's Drive file ID
//   filename  — what to name the new signed file
//   bytes_b64 — the patched PDF bytes, base64-encoded
function handleApproveSignInWithBytes_(body) {
  const d        = body.data || {};
  const fileId   = String(d.file_id  || '').trim();
  const filename = String(d.filename || '').trim();
  const b64      = String(d.bytes_b64 || '');
  if (!fileId || !filename || !b64) {
    return jsonResponse_({ error: 'Missing required fields: file_id, filename, bytes_b64' }, 400);
  }

  const props          = PropertiesService.getScriptProperties();
  const approvedId     = props.getProperty('APPROVED_SENT_ID');
  const needsReviewId  = props.getProperty('NEEDS_REVIEW_ID');
  if (!approvedId)    return jsonResponse_({ error: 'APPROVED_SENT_ID not set' }, 500);
  if (!needsReviewId) return jsonResponse_({ error: 'NEEDS_REVIEW_ID not set' }, 500);

  try {
    // Safety gate: make sure the original lives under Docs Needing Review
    const original = DriveApp.getFileById(fileId);
    if (original.isTrashed()) return jsonResponse_({ error: 'File is trashed' }, 404);
    if (!_isUnderParent_(original, needsReviewId)) {
      return jsonResponse_({ error: 'File not in Docs Needing Review' }, 403);
    }

    // Write signed bytes into Approved Docs
    const approvedFolder = DriveApp.getFolderById(approvedId);
    const bytes          = Utilities.base64Decode(b64);
    const blob           = Utilities.newBlob(bytes, 'application/pdf', filename);
    const newFile        = approvedFolder.createFile(blob);

    // Trash the original unsigned PDF
    original.setTrashed(true);

    // Log — cron will log its own "emailed" row 0-10 min later
    try {
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        .getSheetByName('Automation Log')
        .appendRow([
          new Date(), 'Approvals', 'Signed + Approved', newFile.getName(),
          'Sign-In signed via webapp; moved to ✅ Approved Docs',
          'Pending email', '',
          'Cron will email + archive within 10 min'
        ]);
    } catch (logErr) {
      Logger.log('⚠️ Automation Log write failed on signed approve: ' + logErr);
    }

    return jsonResponse_({
      success:  true,
      file_id:  newFile.getId(),
      filename: newFile.getName(),
    });
  } catch (err) {
    return jsonResponse_({ error: String(err) }, 500);
  }
}


// ═══════════════════════════════════════════════════════════════
// SCRIPT CACHE HELPERS — used by dashboard + revenue endpoints
// ═══════════════════════════════════════════════════════════════
//
// Apps Script CacheService gives us a 6-hour, 100KB-per-key, in-memory
// cache shared across executions. The dashboard handler does heavy
// per-request work (sheet reads + Drive walks for any WO whose
// Archive Folder URL hasn't been backfilled yet); caching its JSON
// payload for 60 seconds turns most refreshes into ~50ms hits.
//
// Cache invalidation: WO/Marking-Item mutators call
// _invalidateCacheKeys_(['dashboard_v1']) after writing. That also
// bumps a global token used by parameterized cache keys (revenue,
// which keys per date range), so we don't have to enumerate variants.

/**
 * Wrap a producer function with CacheService. Returns the cached value
 * on hit, runs `fn()` and caches the JSON-stringified result on miss.
 * Falls through to `fn()` directly if the result exceeds the 100KB
 * per-entry cap (logs a warning rather than throwing).
 */
function _withScriptCache_(key, ttlSec, fn) {
  const cache = CacheService.getScriptCache();
  try {
    const hit = cache.get(key);
    if (hit) {
      try { return JSON.parse(hit); }
      catch (e) { /* corrupt entry — fall through to recompute */ }
    }
  } catch (e) {
    // CacheService can throw on rare backend failures — never let it
    // break the calling handler.
    Logger.log('⚠️ _withScriptCache_ get failed for ' + key + ': ' + e.message);
  }

  const value = fn();
  try {
    const payload = JSON.stringify(value);
    if (payload.length <= 100 * 1024) {
      cache.put(key, payload, ttlSec);
    } else {
      Logger.log('⚠️ _withScriptCache_ skip put for ' + key +
                 ' (size=' + payload.length + ' bytes exceeds 100KB cap)');
    }
  } catch (e) {
    Logger.log('⚠️ _withScriptCache_ put failed for ' + key + ': ' + e.message);
  }
  return value;
}

/**
 * Invalidate the given cache keys AND bump the global cache token used
 * by parameterized cache keys (revenue endpoint keys per date range).
 * Bumping the token instead of enumerating variants means we never
 * have to keep an exhaustive list of revenue cache keys to evict.
 */
function _invalidateCacheKeys_(keys) {
  const cache = CacheService.getScriptCache();
  try {
    if (Array.isArray(keys) && keys.length) cache.removeAll(keys);
  } catch (e) {
    Logger.log('⚠️ _invalidateCacheKeys_ removeAll failed: ' + e.message);
  }
  try {
    // Token is a monotonic value; any cached payload keyed with the
    // previous token instantly becomes unreachable.
    cache.put('cache_token', String(Date.now()), 21600);
  } catch (e) {
    Logger.log('⚠️ _invalidateCacheKeys_ token bump failed: ' + e.message);
  }
}

/**
 * Read the current cache token (used when constructing parameterized
 * cache keys). Initializes lazily so the first reader doesn't see null.
 */
function _getCacheToken_() {
  const cache = CacheService.getScriptCache();
  let t = null;
  try { t = cache.get('cache_token'); } catch (e) { /* fall through */ }
  if (!t) {
    t = String(Date.now());
    try { cache.put('cache_token', t, 21600); } catch (e) { /* best effort */ }
  }
  return t;
}


// ═══════════════════════════════════════════════════════════════
// SIGN-IN TAB — webapp "Sign-In" page backend
// ═══════════════════════════════════════════════════════════════

// ── action: list_employees ────────────────────────────────────
//
// Returns the list of employee names (Employee Registry col B) used
// by the Sign-In tab's crew-row dropdown. Cached for 5 min in
// CacheService so a 10-row sign-in form doesn't hit the sheet 10×.
function handleListEmployees_(body) {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('signin_employees_v1');
  if (cached) {
    return jsonResponse_({ employees: JSON.parse(cached) });
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Employee Registry');
  if (!sheet) return jsonResponse_({ employees: [] });

  const data = sheet.getDataRange().getValues();
  const seen = new Set();
  const employees = [];
  data.slice(1).forEach(r => {
    const name = String(r[1] || '').trim();
    if (!name || seen.has(name)) return;
    seen.add(name);
    employees.push({ name });
  });
  employees.sort((a, b) => a.name.localeCompare(b.name));

  cache.put('signin_employees_v1', JSON.stringify(employees), 300);
  return jsonResponse_({ employees });
}

// ── action: list_signin_queue ─────────────────────────────────
//
// Returns outstanding (date, contract) groups built from the Work Day
// Log. Each group lists the WOs that had Field Reports submitted that
// day on that contract. The Sign-In tab uses this as its master list.
function handleListSignInQueue_(body) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Work Day Log');
  if (!sheet) return jsonResponse_({ queue: [] });

  const data = sheet.getDataRange().getValues();

  // Contract Lookup gives us contract_id + project name for the right pane.
  const clSheet = ss.getSheetByName('Contract Lookup');
  const clData  = clSheet ? clSheet.getDataRange().getValues() : [];
  const lookupContract = (contractNum, borough) => {
    const row = clData.slice(1).find(r =>
      String(r[0] || '').includes(contractNum) && String(r[1] || '') === borough
    );
    return row
      ? { contract_id: String(row[3] || ''), project_name: String(row[4] || '') }
      : { contract_id: '', project_name: '' };
  };

  // Contractor Contacts → the prime contractor's Contact Name (col 1) and
  // Address (col 5) for the sign-in header, keyed by Contractor (col 0).
  // Read once into a map; mirrors the lookup the sign-in PDF builder uses.
  const ccSheet = ss.getSheetByName('Contractor Contacts');
  const ccByContractor = {};
  if (ccSheet) {
    ccSheet.getDataRange().getValues().forEach(r => {
      const name = String(r[0] || '').trim();
      if (name) ccByContractor[name] = {
        contact: String(r[1] || '').trim(),
        address: String(r[5] || '').trim(),
      };
    });
  }

  // ISO-format a Sheet date cell (Date object | string) to YYYY-MM-DD
  // in the local timezone — same logic the rest of the app uses to
  // avoid UTC-shift drift.
  const toIsoDate = (cell) => {
    if (cell instanceof Date && !isNaN(cell.getTime())) {
      const y = cell.getFullYear();
      const m = String(cell.getMonth() + 1).padStart(2, '0');
      const d = String(cell.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    const s = String(cell || '');
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : s;
  };

  const groups = new Map();
  data.slice(1).forEach(row => {
    // WDL post-multi-crew schema:
    //   0 Date, 1 WO#, 2 Contractor, 3 Contract#, 4 Borough, 5 Location,
    //   6 FR Submitted At, 7 Crew Chief, 8 Sign-In Status.
    if (String(row[8] || '').trim() !== 'Pending') return;
    const dateStr     = toIsoDate(row[0]);
    const woId        = String(row[1] || '').trim();
    const contractor  = String(row[2] || '').trim();
    const contractNum = String(row[3] || '').split('/')[0].trim();
    const borough     = String(row[4] || '').trim();
    const location    = String(row[5] || '').trim();
    const crewChief   = String(row[7] || '').trim();
    if (!dateStr || !woId || !contractNum || !borough) return;

    // Key includes contractor AND crew chief. Per the multi-crew model,
    // each crew (identified by chief) gets its own queue card → its own
    // SI submission → its own filename/lifecycle row. Without the chief
    // in the key, two crews on the same source job would collapse into
    // one card and one prime's WOs would silently merge with the other's.
    // Blank chief is a first-class value (legacy single-crew rows) and
    // groups with itself just like a named chief would.
    const key = `${dateStr}|${contractor}|${contractNum}|${borough}|${crewChief}`;
    if (!groups.has(key)) {
      const lookup = lookupContract(contractNum, borough);
      // Billing remap: the sign-in sheet the user signs is labeled with
      // the billing identity (e.g. Denville/BK billed as QU). Source data
      // stays raw — we just surface the remapped tuple alongside it so
      // the queue UI matches what gets printed on the doc.
      const billed = _billingRemap_(contractNum, borough, contractor);
      groups.set(key, {
        queue_id:             key,
        date:                 dateStr,
        contract_number:      contractNum,
        borough:              borough,
        bill_contract_number: billed.contractNum,
        bill_borough:         billed.borough,
        contractor:           contractor,          // raw WO-Tracker contractor (logic/remap)
        // "Prime Contractor" on the sign-in sheet = the Contact Name from
        // Contractor Contacts (col 1); fall back to the contractor company.
        prime_contractor:     (ccByContractor[contractor] && ccByContractor[contractor].contact) || '',
        subcontractor:        CONFIG.EMPLOYER.name, // always Oneiro
        address:              (ccByContractor[contractor] && ccByContractor[contractor].address) || '',
        crew_chief:           crewChief,
        contract_id:          lookup.contract_id,
        project_name:         lookup.project_name,
        wos:                  [],
      });
    }
    const grp = groups.get(key);
    if (!grp.wos.find(w => w.id === woId)) {
      grp.wos.push({ id: woId, location: location });
    }
  });

  const queue = Array.from(groups.values());
  // FIFO: oldest work day on top so overdue sign-ins clear first.
  queue.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return a.contract_number.localeCompare(b.contract_number);
  });

  return jsonResponse_({ queue });
}

// ── action: submit_signin ─────────────────────────────────────
//
// Body shape:
//   data: {
//     queue_id, date (YYYY-MM-DD, shift START date),
//     contract_number, borough, contractor,
//     wo_ids: ["RM-1", "RM-2", ...],
//     crew: [{name, classification, time_in, time_out, hours, overtime,
//             sig_in_b64, sig_out_b64}, ...],
//     contractor_name, contractor_title, date_signed, contractor_signature_b64,
//     source: 'generated' | 'uploaded',
//     upload_blob_b64?, upload_filename?  (when source='uploaded')
//   }
//
// Behavior:
//   1. Append crew rows to Daily Sign-In Data — one row per crew member
//      with the WO# column being a comma-list of every WO this sign-in
//      covers. OT recalculated server-side from the date's day-of-week.
//   2. Mark matching Work Day Log rows as Status='Submitted' so they
//      drop off the queue.
//   3. Generate path → write Sign-In JSON to Drive for the Python filler.
//      Upload path → drop the original PDF directly with _MANUAL suffix.
//   4. Append to Automation Log.
function handleSubmitSignIn_(body) {
  const d = body.data || {};

  // Idempotency cache. Client sends a stable `submit_id` per draft;
  // we cache the response under that id for 10 minutes. A retry after
  // a perceived failure (network blip, slow response, etc) hits the
  // cache and returns the original success — no duplicate rows in
  // Daily Sign-In Data.
  const submitId = String(d.submit_id || '').trim();
  const _cache = submitId ? CacheService.getScriptCache() : null;
  if (_cache) {
    const cached = _cache.get('signin_submit_' + submitId);
    if (cached) {
      try {
        const obj = JSON.parse(cached);
        return jsonResponse_({ ...obj, duplicate: true });
      } catch (e) { /* fall through, recompute */ }
    }
  }


  if (!d.date)            return jsonResponse_({ error: 'Missing date' }, 400);
  if (!d.contract_number) return jsonResponse_({ error: 'Missing contract_number' }, 400);
  if (!d.borough)         return jsonResponse_({ error: 'Missing borough' }, 400);
  if (!Array.isArray(d.wo_ids) || d.wo_ids.length === 0) {
    return jsonResponse_({ error: 'wo_ids must be a non-empty array' }, 400);
  }
  if (!Array.isArray(d.crew) || d.crew.length === 0) {
    return jsonResponse_({ error: 'At least one crew member is required' }, 400);
  }
  // Crew chief comes from the queue card the user picked. Required at
  // FR submit; here we accept blank to keep legacy queue cards (no
  // chief) working — blank chief is the legacy single-crew default.
  const crewChief = String(d.crew_chief || '').trim();
  const source = String(d.source || 'generated');
  if (source !== 'generated' && source !== 'uploaded') {
    return jsonResponse_({ error: 'source must be "generated" or "uploaded"' }, 400);
  }
  if (source === 'uploaded' && !d.upload_blob_b64) {
    return jsonResponse_({ error: 'upload_blob_b64 required when source=uploaded' }, 400);
  }

  let step = 'init';
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // ── Resolve the authoritative shift-start date ────────────────
    // The client sends d.date as the SHIFT START day — already
    // operational-day-correct, either because the FR's Work Day Log
    // row was opDay-bucketed at FR submit time, or because the user
    // explicitly kebab-edited the shift start date in the Sign-In tab.
    // We trust it verbatim. (Re-running opDay here would double-correct
    // since Time In is interpreted on the shift start date — a 02:00
    // Time In on a Monday-shift would wrongly slip back to Sunday.)
    const effectiveDate = String(d.date);

    // ── Determine OT rule by start-day DOW ────────────────────────
    const dowOfWork = (() => {
      const m = String(effectiveDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!m) return -1;
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay();
    })();
    const isWeekend = (dowOfWork === 0 || dowOfWork === 6);

    // ── Daily Sign-In Data append (one row per crew member) ───────
    step = 'Daily Sign-In Data append';
    const signInSheet = ss.getSheetByName('Daily Sign-In Data');
    const woListNorm  = d.wo_ids.map(s => String(s).trim()).filter(Boolean);
    const woListStr   = woListNorm.join(', ');

    // Look up each WO's Location from the WO Tracker so the Daily
    // Sign-In Data Location column mirrors the WO# column (a parallel
    // list). Both are display-only — archive routing always reads
    // per-WO location from the tracker directly.
    const woSheet  = ss.getSheetByName('Work Order Tracker');
    const woData   = woSheet.getDataRange().getValues();
    const allLocations = woListNorm.map(woId => {
      const r = woData.find(rr => String(rr[0] || '') === woId);
      return r ? String(r[5] || '').trim() : '';
    });
    const displayLocation = allLocations.filter(Boolean).join('; ');

    // ── Per-row OT split with same-day lookback ───────────────────
    // OT col reflects each row's *contribution* to the daily 8-hour
    // ST allowance. A worker may have rows on this date from an
    // earlier Sign-In Group (different contract/borough/contractor)
    // OR appear twice in this same submission. Both cases must feed
    // the OT split so the second row carries the OT once the daily
    // total crosses 8. Without this, two independent 4-hour shifts
    // would each record 0 OT despite totaling 8+; cf. the multi-
    // group bug where 6h + 3.75h split as 9.75 ST / 0 OT instead of
    // 8 ST / 1.75 OT.
    const normNameKey = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
    const normDateKey = v => {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      const s = String(v);
      const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) return `${m[1]}-${m[2]}-${m[3]}`;
      const dt = new Date(s);
      return isNaN(dt.getTime()) ? s : Utilities.formatDate(dt, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    };
    const todayKey = normDateKey(effectiveDate);
    const priorHoursByEmp = {};
    const signInData = signInSheet.getDataRange().getValues();
    for (let i = 1; i < signInData.length; i++) {
      const row = signInData[i];
      if (normDateKey(row[0]) !== todayKey) continue;
      const n = normNameKey(row[6]);
      if (!n) continue;
      priorHoursByEmp[n] = (priorHoursByEmp[n] || 0) + (Number(row[10]) || 0);
    }

    // Per-row OT via the shared allocator — single source of truth with
    // handleSaveSignInRowEdits_ (the Approvals-page hours editor) so the
    // two can never diverge. priorHoursByEmp seeds hours already on this
    // opDay from earlier Sign-In Groups so the 8h ST cap is shared.
    const otByRow = _allocateDayOvertime_(
      d.crew.map(m => ({ key: normNameKey(m.name), hours: parseFloat(m.hours) || 0 })),
      isWeekend,
      priorHoursByEmp
    );

    const crewRows = d.crew.map((member, idx) => {
      const hours    = parseFloat(member.hours) || 0;
      const overtime = otByRow[idx];
      return [
        effectiveDate,                                //  0  Date (shift START, opDay-derived)
        woListStr,                                    //  1  Work Order # (comma-list)
        String(d.contractor || '').trim(),            //  2  Prime Contractor
        String(d.contract_number || '').trim(),       //  3  Contract #
        String(d.borough || '').trim(),               //  4  Borough
        displayLocation,                              //  5  Location
        String(member.name || '').trim(),             //  6  Employee Name
        String(member.classification || '').trim(),   //  7  Classification
        _fmt24to12_(member.time_in),                  //  8  Time In  (h:mm AM/PM)
        _fmt24to12_(member.time_out),                 //  9  Time Out (h:mm AM/PM)
        hours,                                        // 10  Hours Worked
        overtime,                                     // 11  Overtime Hours
        crewChief                                     // 12  Crew Chief (per-crew tagging)
      ];
    });

    appendRowsWithProbing_(
      signInSheet,
      crewRows,
      ['Date', 'Work Order #', 'Prime Contractor', 'Contract #', 'Borough',
       'Location', 'Employee Name', 'Classification', 'Time In', 'Time Out',
       'Hours Worked', 'Overtime Hours', 'Crew Chief'],
      'Daily Sign-In Data'
    );

    // ── Mark Work Day Log rows for these WOs as Submitted ─────────
    // Match by (WO# in this submission, Status='Pending', Crew Chief).
    // Filtering by crew chief is critical for multi-crew days: Crew A's
    // submit must not clear Crew B's queue card on the same source job.
    // We DO NOT filter by date because the user may have kebab-overridden
    // the shift start date, OR opDay may have shifted the effective date
    // away from the FR's recorded date. WDL post-multi-crew schema:
    //   7 Crew Chief, 8 Sign-In Status.
    step = 'Work Day Log update';
    const wdlSheet = ss.getSheetByName('Work Day Log');
    if (wdlSheet) {
      const wdlData = wdlSheet.getDataRange().getValues();
      const woSet = new Set(woListNorm);
      for (let i = 1; i < wdlData.length; i++) {
        const row = wdlData[i];
        if (!woSet.has(String(row[1] || '').trim())) continue;
        if (String(row[8] || '').trim() !== 'Pending') continue;
        if (String(row[7] || '').trim() !== crewChief) continue;
        wdlSheet.getRange(i + 1, 9).setValue('Submitted');
      }
    }

    // ── Drive write: JSON for filler OR original upload PDF ───────
    step = 'Drive write';
    const props         = PropertiesService.getScriptProperties();
    const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
    if (!needsReviewId) throw new Error('NEEDS_REVIEW_ID not set');
    const reviewFolder = DriveApp.getFolderById(needsReviewId);
    const signInFolder = getOrCreateSubfolder_(reviewFolder, 'Sign-In Logs');

    const isoDate = String(effectiveDate).slice(0, 10);
    // Filename keeps RAW (contract, borough) so the archive parser
    // matches WO Tracker. Two suffixes can append after the date:
    //   - `_<ContractorSlug>` when the sub-prime billing remap applies
    //     (Brooklyn → Manhattan/Queens) — disambiguates two primes
    //     submitting against the same source job.
    //   - `_chief-<ChiefSlug>` when this submission carries a crew chief
    //     — disambiguates two crews from the SAME prime submitting
    //     against the same source job on the same day.
    // Blank chief omits the suffix entirely → legacy filenames preserved.
    const _siRemapApplies = _hasBillingRemap_(d.contract_number, d.borough) &&
      _billingRemap_(d.contract_number, d.borough, d.contractor).borough !== String(d.borough || '').trim();
    const _siContractorSlug = _siRemapApplies
      ? '_' + String(d.contractor || '').replace(/[^A-Za-z0-9]/g, '')
      : '';
    const _siChiefSlug = crewChief
      ? '_chief-' + crewChief.replace(/[^A-Za-z0-9]/g, '')
      : '';
    const baseName = `SignIn_${d.contract_number}_${d.borough}_${isoDate}${_siContractorSlug}${_siChiefSlug}`;

    let writtenName;

    if (source === 'uploaded') {
      const fileName = `${baseName}_MANUAL.pdf`;
      const existing = signInFolder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);
      const bytes = Utilities.base64Decode(d.upload_blob_b64);
      const blob  = Utilities.newBlob(bytes, 'application/pdf', fileName);
      signInFolder.createFile(blob);
      writtenName = fileName;
    } else {
      // Build the multi-WO Sign-In JSON for fill_signin.py.
      const dateFmt = (iso) => {
        const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (!m) return String(iso);
        const [, yyyy, mm, dd] = m;
        return `${parseInt(mm, 10)}/${parseInt(dd, 10)}/${yyyy.slice(-2)}`;
      };

      const locations = woListNorm.map(woId => {
        const r = woData.find(rr => String(rr[0] || '') === woId);
        return r ? String(r[5] || '') : '';
      });

      // Look up the Contact Name (col 1) + Address (col 5) from Contractor
      // Contacts. The sign-in "Prime Contractor" field is the Contact Name.
      let primeAddress = '';
      let primeContact = '';
      try {
        const ccSheet = ss.getSheetByName('Contractor Contacts');
        if (ccSheet) {
          const ccData = ccSheet.getDataRange().getValues();
          const ccRow  = ccData.find(r => String(r[0] || '').trim() === String(d.contractor || '').trim());
          if (ccRow) {
            primeContact = String(ccRow[1] || '').trim();
            primeAddress = String(ccRow[5] || '').trim();
          }
        }
      } catch (err) {
        Logger.log(`⚠️ Contractor Contacts lookup failed: ${err}`);
      }

      // Apply billing remap for the PDF content: sub-prime work on a
      // contract they didn't win shows the billed-under contract here.
      // Filename / Daily Sign-In Data row stay raw so the archive cron
      // matches WO Tracker (also raw) and the source job stays one
      // folder.
      const _siMapped = _billingRemap_(d.contract_number, d.borough, d.contractor);
      const siContractNum = _siMapped.contractNum;
      const siBorough = _siMapped.borough;
      const boroughName = getBoroughName_(siBorough);
      const contractLabel = boroughName
        ? `${siContractNum} - ${boroughName}`
        : siContractNum;

      // Pre-formatted display string for the "Work Order #" cell on the
      // sign-in template — saves the filler from having to format it.
      const woLabel = woListNorm.length > 1
        ? `${woListNorm.join(', ')} (${woListNorm.length})`
        : woListNorm[0];

      // Project Name / Location field on the sign-in PDF. We pair every
      // WO with its OWN location ("<WO> | <Loc>"); Contract Lookup's
      // project_name is intentionally not used here — that field is
      // sometimes empty and when it isn't, it's the contract's umbrella
      // description (e.g. "Manhattan Pavement Marking"), not the
      // worksite-specific text the form expects.
      const woLocPairs = woListNorm.map((wo, i) => {
        const loc = (locations[i] || '').trim();
        return loc ? `${wo} | ${loc}` : wo;
      });
      const projectName = woLocPairs.join('; ');

      const payload = {
        _type:              'signin',
        date:               dateFmt(effectiveDate),
        prime_contractor:   primeContact || String(d.contractor || ''),
        subcontractor:      CONFIG.EMPLOYER.name,
        contract_number:    contractLabel,
        address:            primeAddress,
        agency:             'DOT',
        project_name:       projectName,
        crew_chief:         crewChief,
        wo_ids:             woListNorm,
        locations:          locations,
        wo_label:           woLabel,
        // Back-compat: older filler builds may still read wo_id (single).
        // Set it to the first WO so they don't crash; new builds prefer
        // wo_label / wo_ids.
        wo_id:              woListNorm[0],
        crew: d.crew.map(m => ({
          name:           String(m.name || '').trim(),
          classification: String(m.classification || '').trim(),
          time_in:        _fmt24to12_(m.time_in),
          time_out:       _fmt24to12_(m.time_out),
          sig_in_b64:     m.sig_in_b64  || '',
          sig_out_b64:    m.sig_out_b64 || '',
        })),
        // Principal sign-off block stays blank — filled by Approvals flow
        // (PrincipalSignModal → pdf-lib in Express).
        contractor_name:          '',
        contractor_title:         '',
        date_signed:              '',
        contractor_signature_b64: '',
      };

      const fileName = `${baseName}.json`;
      const existing = signInFolder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);
      signInFolder.createFile(fileName, JSON.stringify(payload, null, 2), MimeType.PLAIN_TEXT);
      writtenName = fileName;
    }

    // ── Automation Log ────────────────────────────────────────────
    step = 'Automation Log';
    appendRowWithProbing_(
      ss.getSheetByName('Automation Log'),
      [
        new Date(), 'Sign-In Tab', 'Sign-In submitted', writtenName,
        `${d.crew.length} crew × ${woListNorm.length} WO(s) on ${effectiveDate} (${source})`,
        'Pending review', '', ''
      ],
      ['Timestamp', 'Source', 'Action', 'Related', 'Details', 'Status', 'User', 'Next Steps'],
      'Automation Log'
    );

    const result = { success: true, filename: writtenName, source };
    if (_cache) {
      try { _cache.put('signin_submit_' + submitId, JSON.stringify(result), 600); }
      catch (e) { /* cache write best-effort */ }
    }
    return jsonResponse_(result);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const wrapped = new Error(`[step=${step}] ${msg}`);
    wrapped.stack = err && err.stack ? err.stack : wrapped.stack;
    throw wrapped;
  }
}

// ── action: check_signin_continuation ──────────────────────────
//
// Sanity check before a Sign-In submit. If the user just entered
// Time In = X on date D, scan Daily Sign-In Data for a previous row
// whose Time-Out datetime falls within 60 minutes BEFORE D+X. If one
// is found AND it's on a different operational day, return it as a
// suggested "continuation" — the client prompts the user to confirm
// they want to bucket this sign-in under the previous shift's date
// (matters most for shifts that cross weekend or week boundaries).
//
// body.data: { time_in: "HH:MM", default_date: "YYYY-MM-DD" }
// response: { continuation: bool, previous_date?, previous_contract?,
//             previous_time_out?, gap_minutes? }
function handleCheckSignInContinuation_(body) {
  const d = body.data || {};
  const targetTimeIn = String(d.time_in     || '').trim();
  const targetDate   = String(d.default_date || '').trim();
  if (!targetTimeIn || !targetDate) {
    return jsonResponse_({ continuation: false });
  }

  const tParts = _parseSignInTimeOfDay_(targetTimeIn);
  const dParts = targetDate.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!tParts || !dParts) return jsonResponse_({ continuation: false });
  const targetDt = new Date(
    Number(dParts[1]), Number(dParts[2]) - 1, Number(dParts[3]),
    tParts.hours, tParts.minutes
  );

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Daily Sign-In Data');
  if (!sheet) return jsonResponse_({ continuation: false });
  const data = sheet.getDataRange().getValues();

  let bestMatch = null;
  let bestGap   = Infinity;

  data.slice(1).forEach(row => {
    const rowDateRaw = row[0];
    const rowTimeIn  = row[8];
    const rowTimeOut = row[9];
    if (!rowDateRaw || !rowTimeOut) return;

    // Parse row date
    let rowDateObj;
    if (rowDateRaw instanceof Date && !isNaN(rowDateRaw.getTime())) {
      rowDateObj = new Date(
        rowDateRaw.getFullYear(), rowDateRaw.getMonth(), rowDateRaw.getDate()
      );
    } else {
      const m = String(rowDateRaw).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!m) return;
      rowDateObj = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }

    const tin  = _parseSignInTimeOfDay_(rowTimeIn);
    const tout = _parseSignInTimeOfDay_(rowTimeOut);
    if (!tout) return;

    // If Time Out total-mins <= Time In total-mins, the shift crossed
    // midnight — so the actual Time-Out datetime is on row.date + 1.
    const tinMins  = tin ? (tin.hours * 60 + tin.minutes) : 0;
    const toutMins = tout.hours * 60 + tout.minutes;
    const cross = tin && toutMins <= tinMins;
    const toutDate = new Date(
      rowDateObj.getFullYear(),
      rowDateObj.getMonth(),
      rowDateObj.getDate() + (cross ? 1 : 0),
      tout.hours, tout.minutes
    );

    const gapMin = (targetDt.getTime() - toutDate.getTime()) / 60000;
    if (gapMin < 0)   return;   // row is in the future / concurrent
    if (gapMin > 60)  return;   // too far in the past
    if (gapMin >= bestGap) return;

    const rowDateIso = Utilities.formatDate(rowDateObj, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    if (rowDateIso === targetDate) return;   // same day; no continuation prompt needed

    bestMatch = {
      previous_date:     rowDateIso,
      previous_contract: String(row[3] || '').trim() +
        (row[4] ? ' / ' + String(row[4]).trim() : ''),
      previous_time_out: String(rowTimeOut).trim(),
    };
    bestGap = gapMin;
  });

  if (!bestMatch) return jsonResponse_({ continuation: false });
  return jsonResponse_({
    continuation:      true,
    previous_date:     bestMatch.previous_date,
    previous_contract: bestMatch.previous_contract,
    previous_time_out: bestMatch.previous_time_out,
    gap_minutes:       Math.round(bestGap),
  });
}

// Convert "HH:MM" (24-hour, what HTML <input type="time"> emits) into
// "h:mm AM/PM" for storage in Daily Sign-In Data and the Sign-In JSON.
// Accepts already-12-hour input as a no-op so the helper is idempotent.
function _fmt24to12_(s) {
  const t = String(s || '').trim();
  if (!t) return '';
  if (/\b(AM|PM)\b/i.test(t)) return t;
  const m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return t;
  const hh24 = Number(m[1]);
  const mm   = m[2];
  const ampm = hh24 >= 12 ? 'PM' : 'AM';
  const hh12 = (hh24 % 12) || 12;
  return `${hh12}:${mm} ${ampm}`;
}

// Parse Daily Sign-In Data Time In/Out cells. Accepts both 24-hour
// "HH:MM" (what the new Sign-In form posts) and 12-hour "h:mm AM/PM"
// (legacy data from the old Field Report flow).
function _parseSignInTimeOfDay_(s) {
  const t = String(s || '').trim();
  if (!t) return null;
  let m = t.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]) % 12;
    if (m[3].toUpperCase() === 'PM') h += 12;
    return { hours: h, minutes: Number(m[2]) };
  }
  m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return { hours: Number(m[1]), minutes: Number(m[2]) };
  return null;
}


// ── action: list_signin_day_hours ──────────────────────────────
//
// Sums hours already on Daily Sign-In Data for a given operational day.
// The Sign-In tab uses this to display "Shift Totals" — what each
// employee will end up with across ALL sign-ins on this date once the
// in-progress one is submitted, so the user can sanity-check OT before
// submitting the second contract's sheet.
//
// body.data = { date: "YYYY-MM-DD" }
// response  = { totals: { "<employee name>": <hours number>, ... } }
function handleListSignInDayHours_(body) {
  const d = body.data || {};
  const targetDate = String(d.date || '').trim();
  if (!targetDate) return jsonResponse_({ totals: {} });

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Daily Sign-In Data');
  if (!sheet) return jsonResponse_({ totals: {} });
  const data = sheet.getDataRange().getValues();

  const totals = {};
  data.slice(1).forEach(row => {
    const rowDateRaw = row[0];
    let rowDateIso;
    if (rowDateRaw instanceof Date && !isNaN(rowDateRaw.getTime())) {
      rowDateIso = Utilities.formatDate(rowDateRaw, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    } else {
      const m = String(rowDateRaw || '').match(/^(\d{4}-\d{2}-\d{2})/);
      rowDateIso = m ? m[1] : '';
    }
    if (rowDateIso !== targetDate) return;
    const empName = String(row[6] || '').trim();
    const hours   = Number(row[10]) || 0;
    if (!empName) return;
    totals[empName] = (totals[empName] || 0) + hours;
  });

  return jsonResponse_({ totals });
}


// ── Sign-In OT + row helpers (shared submit + admin-edit) ─────

// Allocate per-row overtime across one operational day. `entries` are
// ordered the way hours should fill each employee's 8-hour straight-time
// bucket; `prior` seeds hours already counted earlier that day (other
// sheets / earlier rows in the same submission). Weekend → all OT.
// Returns OT values aligned to `entries`. Single source of truth so
// handleSubmitSignIn_ and handleSaveSignInRowEdits_ can never diverge.
function _allocateDayOvertime_(entries, isWeekend, prior) {
  const counted = Object.assign({}, prior || {});
  return entries.map(e => {
    const hours = Number(e.hours) || 0;
    const key   = e.key;
    let ot;
    if (isWeekend) {
      ot = hours;
    } else {
      const p          = counted[key] || 0;
      const combinedST = Math.min(p + hours, 8);
      const priorST    = Math.min(p, 8);
      ot = Math.max(0, hours - (combinedST - priorST));
    }
    counted[key] = (counted[key] || 0) + hours;
    return ot;
  });
}

function _normDateKey_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const s = String(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? s : Utilities.formatDate(dt, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}
function _normNameKey_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function _slugify_(s)     { return String(s || '').replace(/[^A-Za-z0-9]/g, ''); }
function _round2_(n)      { return Math.round((Number(n) || 0) * 100) / 100; }
function _isWeekendDateStr_(dateStr) {
  const m = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return false;
  const dow = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay();
  return dow === 0 || dow === 6;
}

// Hours between two Sign-In time cells, mirroring the client calcHours:
// cross-midnight (out <= in) rolls to the next day; rounded to 2dp.
function _signInRowHours_(timeIn, timeOut) {
  const a = _parseSignInTimeOfDay_(timeIn);
  const b = _parseSignInTimeOfDay_(timeOut);
  if (!a || !b) return 0;
  let mins = (b.hours * 60 + b.minutes) - (a.hours * 60 + a.minutes);
  if (mins <= 0) mins += 24 * 60;
  return _round2_(mins / 60);
}

// Parse a Sign-In PDF filename into its routing parts. Mirrors the
// archive matcher regex (archiveDocument_ Sign-In branch). Returns null
// for legacy/non-matching names. contractorSlug is suppressed when it is
// really a _MANUAL/_FILLED suffix marker rather than a contractor.
function _parseSignInFilename_(filename) {
  const cleanName = String(filename || '').replace(/\.pdf$/i, '');
  const m = cleanName.match(
    /^SignIn_([^_]+)_([^_]+)_(\d{4}-\d{2}-\d{2})(?:_(?!chief-)([A-Za-z0-9]+))?(?:_chief-([A-Za-z0-9]+))?/
  );
  if (!m) return null;
  const [, contractNum, borough, dateStr, rawContractor, chiefSlug] = m;
  const isSuffixMarker = rawContractor === 'MANUAL' || rawContractor === 'FILLED';
  return {
    contractNum,
    borough,
    dateStr,
    contractorSlug: (rawContractor && !isSuffixMarker) ? rawContractor : '',
    chiefSlug: chiefSlug || '',
  };
}

// True when a Daily Sign-In Data row belongs to the sign-in sheet
// described by `f` (a _parseSignInFilename_ result). Matches by date +
// contract + borough, plus contractor/chief slug when the filename
// carries them (billing-remap / multi-crew disambiguation).
function _signInRowMatchesFile_(row, f) {
  if (_normDateKey_(row[0]) !== f.dateStr) return false;
  if (String(row[3] || '').trim() !== f.contractNum) return false;
  if (String(row[4] || '').trim() !== f.borough) return false;
  if (f.contractorSlug && _slugify_(row[2]) !== f.contractorSlug) return false;
  if (f.chiefSlug && _slugify_(row[12]) !== f.chiefSlug) return false;
  return true;
}

function _signInRowView_(row, sheetRowIndex) {
  return {
    row_index:      sheetRowIndex,
    name:           String(row[6]  || '').trim(),
    classification: String(row[7]  || '').trim(),
    time_in:        String(row[8]  || '').trim(),
    time_out:       String(row[9]  || '').trim(),
    hours:          Number(row[10]) || 0,
    overtime:       Number(row[11]) || 0,
    crew_chief:     String(row[12] || '').trim(),
  };
}


// ── action: list_signin_rows_for_file ─────────────────────────
//
// Powers the Approvals-page hours editor. Given a pending sign-in's
// file_id + filename, returns each Daily Sign-In Data row for that sheet
// (with its 1-based sheet row index so edits can target exact cells),
// the other-sheet hours per employee for that day (for the shift-totals
// strip), and an `ambiguous` flag when a legacy chief-less filename
// can't be safely attributed to one of several crews on that day.
function handleListSignInRowsForFile_(body) {
  const d        = body.data || {};
  const filename = String(d.filename || '').trim();
  const f = _parseSignInFilename_(filename);
  if (!f) return jsonResponse_({ error: 'Could not parse sign-in filename: ' + filename }, 400);

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Daily Sign-In Data');
  if (!sheet) return jsonResponse_({ error: 'Daily Sign-In Data not found' }, 500);
  const data = sheet.getDataRange().getValues();

  const rows       = [];
  const fileRowSet = {};          // 1-based sheet row index → true
  for (let i = 1; i < data.length; i++) {
    if (!_signInRowMatchesFile_(data[i], f)) continue;
    fileRowSet[i + 1] = true;
    rows.push(_signInRowView_(data[i], i + 1));
  }

  // Ambiguity guard — only when the filename carries no chief slug. If
  // the (date, contract, borough) bucket holds rows from 2+ distinct
  // crew chiefs, we can't attribute the file to one crew → refuse edit.
  let ambiguous = false;
  if (!f.chiefSlug) {
    const chiefs = {};
    rows.forEach(r => { const c = _slugify_(r.crew_chief); if (c) chiefs[c] = true; });
    if (Object.keys(chiefs).length >= 2) ambiguous = true;
  }

  // Other-sheet hours per employee for this date (everything on the date
  // NOT belonging to this file) → feeds the shift-totals strip.
  const otherHours = {};
  for (let i = 1; i < data.length; i++) {
    if (fileRowSet[i + 1]) continue;
    if (_normDateKey_(data[i][0]) !== f.dateStr) continue;
    const name = String(data[i][6] || '').trim();
    if (!name) continue;
    otherHours[name] = (otherHours[name] || 0) + (Number(data[i][10]) || 0);
  }

  return jsonResponse_({
    rows,
    other_hours: otherHours,
    meta: {
      date:            f.dateStr,
      contract:        f.contractNum,
      borough:         f.borough,
      contractor_slug: f.contractorSlug,
      chief_slug:      f.chiefSlug,
      ambiguous,
    },
  });
}


// ── action: save_signin_row_edits ─────────────────────────────
//
// Admin corrections to a submitted sign-in. Writes Classification /
// Time In / Time Out / Hours for the edited rows, then recomputes
// Overtime for EVERY row on that operational day — OT caps are per
// employee-per-day across ALL sheets, so editing one crew can shift a
// shared employee's split on another sheet. Daily Sign-In Data is the
// payroll source of truth; the signed PDF is never touched.
function handleSaveSignInRowEdits_(body) {
  const d        = body.data || {};
  const filename = String(d.filename || '').trim();
  const editsIn  = Array.isArray(d.rows) ? d.rows : [];
  const f = _parseSignInFilename_(filename);
  if (!f) return jsonResponse_({ error: 'Could not parse sign-in filename: ' + filename }, 400);
  if (!editsIn.length) return jsonResponse_({ error: 'No rows to save' }, 400);

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return jsonResponse_({ error: 'Could not acquire lock — try again' }, 503);
  }
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Daily Sign-In Data');
    if (!sheet) return jsonResponse_({ error: 'Daily Sign-In Data not found' }, 500);
    const data    = sheet.getDataRange().getValues();
    const lastRow = data.length;

    // Validate every edited row still matches this file before writing
    // (guards a stale row_index after concurrent appends / re-sorts).
    const editByRow = {};
    for (const e of editsIn) {
      const ri = Number(e.row_index);
      if (!ri || ri < 2 || ri > lastRow) {
        return jsonResponse_({ error: `Row ${e.row_index} out of range — refresh and retry` }, 409);
      }
      if (!_signInRowMatchesFile_(data[ri - 1], f)) {
        return jsonResponse_({ error: `Row ${ri} no longer matches this sheet — refresh and retry` }, 409);
      }
      editByRow[ri] = e;
    }

    // 1) Apply edits (class/times/hours) to the in-memory working copy.
    Object.keys(editByRow).forEach(riStr => {
      const row = data[Number(riStr) - 1];
      const e   = editByRow[riStr];
      const cls = (e.classification === 'LP' || e.classification === 'SAT')
        ? e.classification : String(row[7] || '').trim();
      const tin  = String(e.time_in  || '').trim();
      const tout = String(e.time_out || '').trim();
      row[7]  = cls;
      row[8]  = _fmt24to12_(tin);
      row[9]  = _fmt24to12_(tout);
      row[10] = _signInRowHours_(tin, tout);
    });

    // 2) Recompute OT for the whole day in sheet order (across all sheets).
    const dayRowIdx = [];                       // 0-based data indices for the date
    for (let i = 1; i < data.length; i++) {
      if (_normDateKey_(data[i][0]) === f.dateStr) dayRowIdx.push(i);
    }
    const beforeOt = {};
    dayRowIdx.forEach(i => { beforeOt[i] = Number(data[i][11]) || 0; });
    const otVals = _allocateDayOvertime_(
      dayRowIdx.map(i => ({ key: _normNameKey_(data[i][6]), hours: Number(data[i][10]) || 0 })),
      _isWeekendDateStr_(f.dateStr), {}
    );
    dayRowIdx.forEach((i, k) => { data[i][11] = _round2_(otVals[k]); });

    // 3a) Edited rows: write Classification..Overtime (sheet cols 8–12).
    let fileTouched = 0, otherTouched = 0;
    Object.keys(editByRow).forEach(riStr => {
      const i = Number(riStr) - 1;
      sheet.getRange(i + 1, 8, 1, 5)
        .setValues([[data[i][7], data[i][8], data[i][9], data[i][10], data[i][11]]]);
      fileTouched++;
    });
    // 3b) Non-edited day rows: only OT (col 12) can have shifted.
    dayRowIdx.forEach(i => {
      if (editByRow[i + 1]) return;
      if ((Number(data[i][11]) || 0) === beforeOt[i]) return;
      sheet.getRange(i + 1, 12).setValue(data[i][11]);
      if (_signInRowMatchesFile_(data[i], f)) fileTouched++; else otherTouched++;
    });

    // Refreshed file-row view for the UI.
    const updatedFileRows = [];
    for (let i = 1; i < data.length; i++) {
      if (_signInRowMatchesFile_(data[i], f)) updatedFileRows.push(_signInRowView_(data[i], i + 1));
    }

    _logAutomation_(
      'Sign-In Edit',
      'Admin edited hours',
      `${f.contractNum}/${f.borough} ${f.dateStr}${f.chiefSlug ? ' (' + f.chiefSlug + ')' : ''}`,
      `Edited ${Object.keys(editByRow).length} row(s); recomputed OT on ${otherTouched} other-sheet row(s). File: ${filename}`,
      'Success', 'No'
    );

    return jsonResponse_({
      success: true,
      rows: updatedFileRows,
      other_sheet_updates: otherTouched,
    });
  } finally {
    lock.releaseLock();
  }
}


// ── action: approve_doc_skip_signoff ───────────────────────────
//
// Same as approve_doc but the Approvals UI uses this for manually-
// uploaded sign-in PDFs the principal already wet-signed by hand. The
// move to Approved Docs happens with no pdf-lib overlay; the cron
// archive + email path is identical.
function handleApproveDocSkipSignoff_(body) {
  const d = body.data || {};
  const fileId = String(d.file_id || '').trim();
  if (!fileId) return jsonResponse_({ error: 'Missing file_id' }, 400);

  const props      = PropertiesService.getScriptProperties();
  const approvedId = props.getProperty('APPROVED_SENT_ID');
  if (!approvedId) return jsonResponse_({ error: 'APPROVED_SENT_ID not set' }, 500);

  // Same trust model as handleApproveDoc_ — see comments there.

  let step = 'init';
  try {
    step = 'getFileById';
    const file = _withDriveRetry_('getFileById', () => DriveApp.getFileById(fileId));

    step = 'isTrashed';
    if (file.isTrashed()) return jsonResponse_({ error: 'File is trashed' }, 404);

    step = 'moveTo Approved Docs';
    _withDriveRetry_('moveTo skip-signoff', () => {
      file.moveTo(DriveApp.getFolderById(approvedId));
    });

    try {
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        .getSheetByName('Automation Log')
        .appendRow([
          new Date(), 'Approvals', 'Approved (sign-off skipped)', file.getName(),
          'Moved to ✅ Approved Docs without principal signature overlay (already wet-signed)',
          'Pending email', '',
          'Cron will email + archive within 10 min'
        ]);
    } catch (logErr) {
      Logger.log('⚠️ Automation Log write failed on skip-signoff: ' + logErr);
    }

    return jsonResponse_({ success: true, file_id: fileId });
  } catch (err) {
    Logger.log(`❌ handleApproveDocSkipSignoff_ failed at step=${step}: ${err}\n${err && err.stack || ''}`);
    return jsonResponse_({ error: `[step=${step}] ${err && err.message || err}` }, 500);
  }
}


// ── action: log_wo_scan_failure ───────────────────────────────

/**
 * The Python worker calls this when a scan can't be parsed (Pass 1
 * empty, Pass 2 returns _parse_error, etc.). Writes a single row to
 * Automation Log with a shape handleGetScanStatus_ can key on:
 *   col 1 (Source)  = 'WO Scan'
 *   col 2 (Action)  = 'Parse failed'
 *   col 4 (Details) contains the file_id substring so the webapp can
 *                   surface the error to the user who uploaded it.
 *
 * body.data = { file_id, filename, error }
 */
function handleLogWOScanFailure_(body) {
  const d = body.data || {};
  const fileId   = String(d.file_id  || '').trim();
  const filename = String(d.filename || '').trim();
  const err      = String(d.error    || 'unknown error').trim();
  if (!fileId) return jsonResponse_({ error: 'Missing file_id' }, 400);

  try {
    const ss  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const log = ss.getSheetByName('Automation Log');
    log.appendRow([
      new Date(),
      'WO Scan',
      'Parse failed',
      filename,
      // Include file_id verbatim so handleGetScanStatus_ can substring-match on it
      `file_id=${fileId} | ${err}`,
      'Error',
      '',
      'Webapp will show this to the uploader for retry'
    ]);
    return jsonResponse_({ success: true });
  } catch (e) {
    Logger.log('⚠️ handleLogWOScanFailure_ write failed: ' + e);
    return jsonResponse_({ error: String(e) }, 500);
  }
}


// ── action: write_wo ──────────────────────────────────────────────────────────

/**
 * Writes a parsed Work Order to the WO Tracker sheet and archives the source PDF.
 *
 * body.file_id  — Drive file ID of the original scanned WO PDF
 * body.data     — normalized dict from parse_work_order.normalize_wo_data()
 *
 * WO Tracker columns (0-indexed, 42 total):
 *  0  Work Order #          11  WO Received Date       22  Issues Reported
 *  1  Prime Contractor      12  Water Blast Required?  23  Photos Uploaded?
 *  2  Contract Number       13  Water Blast Confirmed? 24  Production Log Done?
 *  3  Borough               14  Water Blast SQFT       25  Field Report Done?
 *  4  Contract ID / Reg #   15  Status                 26  Invoice #
 *  5  Location              16  Dispatch Date          27  Invoice Date
 *  6  From Street           17  Work Start Date        28  Invoice Amount
 *  7  To Street             18  Work End Date          29  Invoice Sent?
 *  8  Due Date              19  Marking Types          30  Payment Received?
 *  9  Priority Level        20  Quantity Completed     31  Payment Date
 * 10  Pavement Work Type    21  Paint / Material Used  32  Certified Payroll Week
 *                                                      33  Filed?
 *                                                      34  Notes
 *                                                      35  Date Entered           (from WO scan, for CFR)
 *                                                      36  School                 (from WO scan, default "NA")
 *                                                      37  Prep By                (from WO scan, for CFR)
 *                                                      38  Scan File ID           (source PDF / split in Scan Inbox)
 *                                                      39  Combined Scan File ID  (only set for multi-WO splits — shared by all splits from same combined PDF)
 *                                                      40  Scan Upload Timestamp  (Date — lets the Scan WO page query "today's uploads")
 *                                                      41  Original Filename      (filename the user picked in the webapp; all splits share it)
 */
function handleWriteWO_(body) {
  const fileId = body.file_id;
  const d      = body.data || {};

  if (!d.work_order_id) {
    return jsonResponse_({ error: 'Missing work_order_id in data' }, 400);
  }

  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  ensureWoTrackerCFRCols_(woSheet);
  const allRows = woSheet.getDataRange().getValues();

  // Breadcrumb — verify the 3 CFR fields arrived from the Vision parser
  Logger.log('CFR scan fields → date_entered=' + JSON.stringify(d.date_entered)
             + ', school=' + JSON.stringify(d.school)
             + ', prep_by=' + JSON.stringify(d.prep_by));

  // ── Duplicate check ───────────────────────────────────────────
  const isDuplicate = allRows.slice(1).some(r => String(r[0]) === String(d.work_order_id));
  if (isDuplicate) {
    Logger.log('⚠️ WO already in tracker: ' + d.work_order_id + ' — deleting from Scan Inbox');
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
    }
    return jsonResponse_({ success: true, duplicate: true, work_order_id: d.work_order_id });
  }

  // ── Contract ID lookup ────────────────────────────────────────
  // Strip the suffix (e.g. /SFT, /PRE) from both the incoming contract number
  // and each row in Contract Lookup before comparing, so they always match.
  const contractNumStripped = String(d.contract_number || '').split('/')[0].trim();
  const clSheet  = ss.getSheetByName('Contract Lookup');
  const clData   = clSheet ? clSheet.getDataRange().getValues() : [];
  const contractRow = clData.slice(1).find(r => {
    const clNum     = String(r[0] || '').split('/')[0].trim();
    const clBorough = String(r[1] || '').trim();
    return clNum === contractNumStripped && clBorough === String(d.borough || '').trim();
  });

  // Contract Lookup col 3 holds the Contract ID / Reg #
  // Columns: 0=Contract Number, 1=Borough Code, 2=Borough Full Name, 3=Contract ID/Reg #
  const CONTRACT_ID_COL = 3;
  const contractId = contractRow ? String(contractRow[CONTRACT_ID_COL] || '').trim() : '';
  const contractIdMissing = !contractId;

  // ── Build row (35 columns) ────────────────────────────────────
  // Columns 0-15 come from the WO scan.
  // Columns 16-34 are operational — filled later from the web app and generated docs.
  // Yes/No tracking columns default to "No" so status is trackable from day one.
  const row = [
    d.work_order_id      || '',  //  0  Work Order #
    d.prime_contractor   || '',  //  1  Prime Contractor
    d.contract_number    || '',  //  2  Contract Number
    d.borough            || '',  //  3  Borough
    contractId,                  //  4  Contract ID / Reg # (from Contract Lookup; blank if not found)
    d.location           || '',  //  5  Location
    d.from_street        || '',  //  6  From Street
    d.to_street          || '',  //  7  To Street
    d.due_date           || '',  //  8  Due Date
    d.priority_level     || '',  //  9  Priority Level
    d.pavement_work_type || '',  // 10  Pavement Work Type
    d.wo_received_date   || '',  // 11  WO Received Date
    d.water_blast_required  || '',  // 12  Water Blast Required?
    d.water_blast_confirmed || '',  // 13  Water Blast Confirmed?
    d.water_blast_sqft   || '',     // 14  Water Blast SQFT
    'Received',                  // 15  Status

    // ── Operational columns (filled from web app / generated docs) ──
    '', '', '',                  // 16-18  Dispatch / Work Start / Work End dates
    '',                          // 19  Marking Types    ← from crew web app
    '', '', '',                  // 20-22  Quantity Completed, Paint/Material, Issues
    'No', 'No', 'No',            // 23-25  Photos, Prod Log, Field Report Done
    '', '', '', 'No',            // 26-29  Invoice #, Date, Amount, Sent
    'No', '',                    // 30-31  Payment Received, Date
    '',                          // 32  Certified Payroll Week
    'No',                        // 33  Filed
    '',                          // 34  Notes             ← from crew web app
    d.date_entered || '',        // 35  Date Entered      ← from WO scan (for CFR)
    d.school       || 'NA',      // 36  School            ← from WO scan (default NA)
    d.prep_by      || '',        // 37  Prep By           ← from WO scan (for CFR)
    fileId         || '',        // 38  Scan File ID      ← source PDF (or split) in Scan Inbox
    body.combined_file_id || '', // 39  Combined Scan File ID ← only for multi-WO stack splits; blank otherwise
    new Date(),                  // 40  Scan Upload Timestamp  ← lets the Scan WO page query "today's uploads"
    body.original_filename || '' // 41  Original Filename       ← what the user picked in the webapp (all splits from one combined PDF share this)
  ];

  woSheet.appendRow(row);
  Logger.log('✅ WO added to tracker: ' + d.work_order_id
             + (contractIdMissing ? ' (Contract ID not found in lookup)' : ''));

  // ── Seed Marking Items ────────────────────────────────────────
  // The parser hands us d.top_markings and d.intersection_grid — expand
  // them into per-crosswalk / per-direction rows so the Field Report UI
  // can load them later and the crew enters SF per item. No-op if the
  // Marking Items sheet doesn't exist (i.e. setupMarkingItems not run).
  const seededCount = seedMarkingItems_(ss, d);
  Logger.log(`   📋 Seeded ${seededCount} Marking Items for ${d.work_order_id}`);

  // ── Archive the source PDF ────────────────────────────────────
  if (fileId) archiveWOFile_(fileId, d);

  // ── Automation Log ────────────────────────────────────────────
  const actionNote = contractIdMissing
    ? 'Action Required: Contract ID / Reg # not found in Contract Lookup — request from prime contractor and add to both WO Tracker (col E) and Contract Lookup sheet'
    : 'WO intake complete — review extracted fields for accuracy';

  ss.getSheetByName('Automation Log').appendRow([
    new Date(),
    'Scan Inbox Parser',
    'WO Scan Processed',
    d.work_order_id,
    (d.prime_contractor || '') + ' / ' + (d.contract_number || '') + ' / ' + (d.location || ''),
    'Added to Tracker',
    '',
    actionNote
  ]);

  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ success: true, work_order_id: d.work_order_id });
}


/**
 * Archive the original WO PDF scan into:
 *   Archive / [Contractor] / [ContractNum - Borough] / [WO# - Location] /
 * Then trash the original in Scan Inbox so the archive is the single
 * source of truth. Non-fatal: the WO tracker row is already written
 * before this runs, so an archive failure never rolls that back.
 *
 * Failures + successes are logged to the Automation Log sheet (not just
 * Logger.log) so admins have a visible trail — a silent try/catch was
 * previously leaving files stuck in Scan Inbox with no user-facing
 * signal that anything had gone wrong.
 */
function archiveWOFile_(fileId, d) {
  const ss       = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const logSheet = ss.getSheetByName('Automation Log');
  const woId     = d.work_order_id || '(unknown)';

  try {
    const props      = PropertiesService.getScriptProperties();
    const archiveId  = props.getProperty('ARCHIVE_ID');
    if (!archiveId) {
      const msg = 'ARCHIVE_ID script property is not set';
      Logger.log('⚠️ archiveWOFile_: ' + msg);
      if (logSheet) logSheet.appendRow([
        new Date(), 'Scan Inbox Parser', 'WO PDF archive failed', woId, msg,
        'Error', '', 'Yes — file is still in Scan Inbox'
      ]);
      return;
    }
    if (!fileId) {
      const msg = 'No fileId provided by the worker — cannot archive';
      Logger.log('⚠️ archiveWOFile_: ' + msg);
      if (logSheet) logSheet.appendRow([
        new Date(), 'Scan Inbox Parser', 'WO PDF archive failed', woId, msg,
        'Error', '', 'Yes'
      ]);
      return;
    }

    const file        = DriveApp.getFileById(fileId);
    const archiveRoot = DriveApp.getFolderById(archiveId);
    const contractor  = d.prime_contractor || 'Unknown';
    const contractNum = String(d.contract_number || '').split('/')[0];
    const borough     = d.borough || '';
    const location    = d.location || d.work_order_id;

    // Archive / Contractor / ContractNum - Borough / WO# - Location /
    const contractorFolder = getOrCreateSubfolder_(archiveRoot, contractor);
    const contractFolder   = getOrCreateSubfolder_(contractorFolder,
                               contractNum + (borough ? ' - ' + getBoroughName_(borough) : ''));
    const woFolder         = getOrCreateSubfolder_(contractFolder,
                               d.work_order_id + ' - ' + location);

    file.makeCopy(file.getName(), woFolder);
    file.setTrashed(true);  // delete from Scan Inbox — archive is the single source of truth

    // Cache the WO folder URL on the Tracker row (col 43 / 0-idx 42)
    // so the dashboard + Field Report panel can render the "View WO"
    // link without paying a Drive walk. getWOFolder_ has the same
    // side-effect for downstream archives, but the scan-upload path
    // bypasses it and historically left col 43 blank until a
    // downstream doc was filed. Wrapped in try/catch — cache failure
    // must never break the upload pipeline.
    try {
      const woSheet = ss.getSheetByName('Work Order Tracker');
      ensureWoTrackerExtraCols_(woSheet);
      const woData  = woSheet.getDataRange().getValues();
      const rowIdx  = woData.findIndex(r => String(r[0] || '').trim() === d.work_order_id);
      if (rowIdx > 0 && !woData[rowIdx][42]) {
        woSheet.getRange(rowIdx + 1, 43).setValue(woFolder.getUrl());
      }
    } catch (e) {
      Logger.log('⚠️ archiveWOFile_: failed to cache Archive Folder URL for ' + woId + ': ' + e.message);
    }

    // Geocode the WO and cache lat/lng to cols 46-49 so the Nav tab's
    // map can render a pin for it. Smart cluster validation against
    // up to 3 Marking-Items intersections; > 1 mile spread = no pin
    // (admin can set manually via the Nav tab's Edit Coordinates flow).
    // Wrapped in try/catch — geocoding failure is non-fatal to scan
    // intake (the WO row is already written).
    try {
      const result = geocodeWO_(d, ss);
      _persistGeocode_(ss, d.work_order_id, result);
    } catch (e) {
      Logger.log('⚠️ archiveWOFile_: geocoding failed for ' + woId + ': ' + e.message);
      // Best-effort: still write a warning so admin sees something
      // in the Tracker.
      try {
        _persistGeocode_(ss, d.work_order_id, {
          warning: 'Geocoder threw: ' + (e.message || String(e))
        });
      } catch (_) { /* nothing left to do */ }
    }

    const pathNote = `${contractor}/${contractNum}${borough ? ' - ' + getBoroughName_(borough) : ''}/${d.work_order_id} - ${location}`;
    Logger.log('📁 WO PDF archived: ' + woId + ' → ' + pathNote);
    if (logSheet) logSheet.appendRow([
      new Date(), 'Scan Inbox Parser', 'WO PDF archived', woId, pathNote,
      'Completed', '', 'No'
    ]);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    Logger.log('⚠️ Could not archive WO file: ' + msg + (err && err.stack ? '\n' + err.stack : ''));
    if (logSheet) logSheet.appendRow([
      new Date(), 'Scan Inbox Parser', 'WO PDF archive failed', woId, `fileId=${fileId || '(none)'} | ${msg}`,
      'Error', '', 'Yes — run retryArchiveStuckWO(\'' + woId + '\') to retry'
    ]);
    // Non-fatal — WO row was already written to tracker
  }
}


/**
 * One-off helper — run from the Apps Script editor to retry archiving
 * a WO whose source PDF is still stuck in Scan Inbox because the
 * original archive step threw. Finds the WO in the tracker, reads its
 * Scan File ID (col 39 / 0-idx 38), reconstructs the archive context
 * from the same row, and reuses archiveWOFile_.
 *
 * Usage:
 *   - From the editor's Run button (no args): edit DEFAULT_WO_TO_RETRY
 *     below to the stuck WO # and hit Run.
 *   - From code / console: retryArchiveStuckWO('RM-43316')
 */
function retryArchiveStuckWO(woId) {
  // Fallback WO # for one-click runs from the Apps Script editor — edit
  // this line when you need to rerun the helper on a different stuck WO.
  const DEFAULT_WO_TO_RETRY = 'RM-43316';

  woId = woId || DEFAULT_WO_TO_RETRY;
  if (!woId) throw new Error('retryArchiveStuckWO requires a WO # (e.g. "RM-43316")');
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const rows    = woSheet.getDataRange().getValues();
  const row     = rows.slice(1).find(r => String(r[0]) === String(woId));
  if (!row) throw new Error('WO not found in tracker: ' + woId);

  const fileId = String(row[38] || '');
  if (!fileId) throw new Error('No Scan File ID on row for ' + woId + ' — nothing to archive');

  const d = {
    work_order_id:    String(row[0]),
    prime_contractor: String(row[1]),
    contract_number:  String(row[2]),
    borough:          String(row[3]),
    location:         String(row[5]),
  };
  Logger.log('↻ Retrying archive for ' + woId + ' (fileId=' + fileId + ')');
  archiveWOFile_(fileId, d);
  Logger.log('✅ retryArchiveStuckWO done — check Automation Log for result');
}


/**
 * Idempotently ensures the WO Tracker has every extra column we've
 * added over time: the 3 CFR cols (Date Entered, School, Prep By),
 * the 3 scan-tracking cols, and the persisted Archive Folder URL
 * that lets the dashboard skip per-WO Drive walks.
 * Safe to call on every scan intake — only writes headers that are
 * actually missing.
 *
 * Column layout (1-indexed): 36 = Date Entered, 37 = School,
 *   38 = Prep By, 39 = Scan File ID, 40 = Combined Scan File ID,
 *   41 = Scan Upload Timestamp, 42 = Original Filename,
 *   43 = Archive Folder URL.
 */
function ensureWoTrackerExtraCols_(woSheet) {
  const EXTRA_HEADERS = [
    'Date Entered',            // col 36 / 0-idx 35
    'School',                  // col 37 / 0-idx 36
    'Prep By',                 // col 38 / 0-idx 37
    'Scan File ID',            // col 39 / 0-idx 38
    'Combined Scan File ID',   // col 40 / 0-idx 39
    'Scan Upload Timestamp',   // col 41 / 0-idx 40 — for "today's uploads" query
    'Original Filename',       // col 42 / 0-idx 41 — filename the user picked in the webapp
    'Archive Folder URL',      // col 43 / 0-idx 42 — written once by getWOFolder_; read by dashboard
    // Per-WO doc lifecycle flags — only CFR + Invoice live here now.
    // PL / SI / CP were trimmed by migrateRemoveObsoleteDocColumns and
    // moved to the Doc Lifecycle Log (per-doc storage). Legacy Field
    // Report Done? (0-idx 25) and Invoice Sent? (0-idx 29) stay put.
    'CFR Sent?',               // col 44 / 0-idx 43 — pairs with col 26/0-idx 25 Field Report Done?
    'Invoice Done?',           // col 45 / 0-idx 44 — pairs with col 30/0-idx 29 Invoice Sent?
    // Geocoding pins — written by geocodeWO_ at scan intake or by
    // backfillWOGeocode_batch* / manual update_wo_coordinates. Drives
    // the Nav tab's map view.
    'Latitude',                // col 46 / 0-idx 45 — pin lat (start of job)
    'Longitude',               // col 47 / 0-idx 46 — pin lng
    'Geocode Warning',         // col 48 / 0-idx 47 — non-empty when cluster check or API call failed
    'Geocoded At',             // col 49 / 0-idx 48 — timestamp of last successful geocode (or last attempt)
  ];
  const START_COL = 36;
  const N = EXTRA_HEADERS.length;
  const lastCol = woSheet.getLastColumn();
  if (lastCol >= START_COL + N - 1) {
    const hdrs = woSheet.getRange(1, START_COL, 1, N).getValues()[0];
    const allSet = EXTRA_HEADERS.every((h, i) => String(hdrs[i]).trim() === h);
    if (allSet) return;
  }
  woSheet.getRange(1, START_COL, 1, N).setValues([EXTRA_HEADERS]);
  woSheet.getRange(1, START_COL, 1, N).setFontWeight('bold');
}

/**
 * Back-compat alias — older code paths still call ensureWoTrackerCFRCols_.
 * Keep both names live so we don't break any callers during rollout.
 */
function ensureWoTrackerCFRCols_(woSheet) { return ensureWoTrackerExtraCols_(woSheet); }

/**
 * One-shot bootstrap that extends the WO Tracker with the new doc
 * lifecycle columns (CFR/Prod Log/Sign-In/CP/Invoice — Done? + Sent?).
 * Run once from the Apps Script editor after deploying.
 *
 * Public name (no trailing underscore) so it appears in the editor's
 * function-runner dropdown. Idempotent — safe to re-run; it only writes
 * headers that aren't already present.
 *
 * Avoid using setupAutomation for this — that function reinstalls
 * triggers and creates folder structure, both of which are out of
 * scope for a column rename / add.
 */
function setupDocLifecycleColumns() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) {
    Logger.log('❌ setupDocLifecycleColumns: Work Order Tracker sheet not found');
    return;
  }
  ensureWoTrackerExtraCols_(woSheet);
  Logger.log('✅ setupDocLifecycleColumns: WO Tracker schema reconciled with Done/Sent columns');
}


// ═══════════════════════════════════════════════════════════════
// DOC LIFECYCLE LOG — per-doc tracking for time-anchored docs
// ═══════════════════════════════════════════════════════════════
//
// Production Logs, Sign-Ins, and Certified Payrolls don't fit a
// per-WO storage shape — a WO that worked 3 days has 3 PLs, not 1.
// The legacy WO Tracker columns can only track "at least one of these
// is done"; once flipped from one occurrence's archive, subsequent
// occurrences are invisible to the data model.
//
// The Doc Lifecycle Log holds one row per actual doc instance, keyed
// by a synthetic Doc ID:  <DOCTYPE>_<ANCHOR>_<CONTRACTNUM>_<BOROUGH>
//   PL_2026-04-13_84125MBTP701_BK
//   SI_2026-04-13_84125MBTP701_BK
//   CP_2026-04-12_84125MBTP701_BK   (week-start anchor)
//
// CFR + Invoice stay on the WO Tracker (their data shape is genuinely
// per-WO).

const DOC_LIFECYCLE_LOG_SHEET_ = 'Doc Lifecycle Log';
const DOC_LIFECYCLE_HEADERS_ = [
  'Doc ID',           // 0  primary key (synthetic)
  'Doc Type',         // 1  Production Log | Sign-In | Certified Payroll
  'Anchor Date',      // 2  YYYY-MM-DD (work date for PL/SI; week-start for CP)
  'Prime Contractor', // 3
  'Contract #',       // 4
  'Borough',          // 5  M / BX / BK / QU / SI
  'WO IDs',           // 6  comma-separated
  'Done',             // 7  Yes / blank
  'Sent',             // 8  Yes / blank (always blank for SI; not tracked)
  'Drive File ID',    // 9  set by archiveDocument_
  'Done At',          // 10 timestamp
  'Sent At',          // 11 timestamp
  'Notes',            // 12 free-form
];

/**
 * Idempotently create the Doc Lifecycle Log sheet. Public name (no
 * underscore) so it appears in the Apps Script editor dropdown — the
 * admin runs this once after deploying.
 */
function setupDocLifecycleLog() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(DOC_LIFECYCLE_LOG_SHEET_);
  const createdNew = !sheet;
  if (createdNew) sheet = ss.insertSheet(DOC_LIFECYCLE_LOG_SHEET_);

  sheet.getRange(1, 1, 1, DOC_LIFECYCLE_HEADERS_.length).setValues([DOC_LIFECYCLE_HEADERS_]);
  sheet.getRange(1, 1, 1, DOC_LIFECYCLE_HEADERS_.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Done / Sent dropdown validators on rows 2+ for visual cue.
  const yesNo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes'], true)   // blank also valid; Yes is the only filled value
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 8, 9999, 1).setDataValidation(yesNo);   // Done
  sheet.getRange(2, 9, 9999, 1).setDataValidation(yesNo);   // Sent

  Logger.log(createdNew
    ? '✅ setupDocLifecycleLog: created sheet "' + DOC_LIFECYCLE_LOG_SHEET_ + '"'
    : '↻ setupDocLifecycleLog: sheet already exists; headers reconciled');
}

/**
 * Build the synthetic Doc ID from natural identifiers. Deterministic
 * — used both at archive time (upsert) and at backfill time, so the
 * same logical doc always resolves to the same key.
 *
 *   PL_2026-04-13_84125MBTP701_BK
 *   SI_2026-04-13_84125MBTP701_BK
 *   CP_2026-04-12_84125MBTP701_BK
 *
 * docType maps:
 *   'Production Log'    → PL
 *   'Sign-In'           → SI
 *   'Certified Payroll' → CP
 */
const _DOC_TYPE_PREFIX_ = Object.freeze({
  'Production Log':    'PL',
  'Sign-In':           'SI',
  'Certified Payroll': 'CP',
});
function _docLifecycleId_(docType, anchorIso, contractNum, borough, crewChief) {
  const prefix = _DOC_TYPE_PREFIX_[docType];
  if (!prefix) return '';
  // Strip any '/EXT' suffix from contract number (matches the rest of
  // the codebase — see split('/')[0] in archiveDocument_ etc.).
  const cn = String(contractNum || '').split('/')[0].trim();
  const base = prefix + '_' + String(anchorIso).trim() + '_' + cn + '_' + String(borough).trim();
  // Optional per-crew suffix. When two crews work the same source job
  // on the same day, each gets its own lifecycle row → Doc Status shows
  // both pending entries. Blank chief omits the suffix → matches legacy
  // doc IDs so already-archived rows keep matching.
  const chiefSlug = String(crewChief || '').replace(/[^A-Za-z0-9]/g, '');
  return chiefSlug ? (base + '_chief-' + chiefSlug) : base;
}

/**
 * Read the entire Doc Lifecycle Log into an array of objects keyed by
 * Doc ID. Single sheet read; callers can pass the result to multiple
 * helpers without re-fetching.
 */
function _readDocLifecycle_(ss) {
  const sheet = ss.getSheetByName(DOC_LIFECYCLE_LOG_SHEET_);
  if (!sheet) return { rows: [], byId: {} };
  const data = sheet.getDataRange().getValues();
  const rows = [];
  const byId = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const id = String(r[0] || '').trim();
    if (!id) continue;
    const obj = {
      sheet_row:    i + 1,           // 1-indexed
      doc_id:       id,
      doc_type:     String(r[1] || '').trim(),
      anchor:       (r[2] instanceof Date)
                      ? Utilities.formatDate(r[2], CONFIG.TIMEZONE, 'yyyy-MM-dd')
                      : String(r[2] || '').trim(),
      contractor:   String(r[3] || '').trim(),
      contract_num: String(r[4] || '').trim(),
      borough:      String(r[5] || '').trim(),
      wo_ids:       String(r[6] || '').split(',').map(s => s.trim()).filter(Boolean),
      done:         String(r[7] || '').toLowerCase() === 'yes',
      sent:         String(r[8] || '').toLowerCase() === 'yes',
      file_id:      String(r[9] || '').trim(),
      done_at:      r[10] || '',
      sent_at:      r[11] || '',
      notes:        String(r[12] || '').trim(),
    };
    rows.push(obj);
    byId[id] = obj;
  }
  return { rows, byId };
}

/**
 * Insert or update a row in the Doc Lifecycle Log. Doc ID is the
 * primary key. If the row exists, mutates the supplied fields and
 * stamps Done At / Sent At when those flags first flip blank → Yes.
 *
 * Idempotent: re-archiving a doc that's already in the log just
 * updates the Drive File ID + done timestamp without duplicating.
 */
function _upsertDocLifecycleRow_(ss, payload) {
  const sheet = ss.getSheetByName(DOC_LIFECYCLE_LOG_SHEET_);
  if (!sheet) {
    Logger.log('⚠️ _upsertDocLifecycleRow_: Doc Lifecycle Log not found — run setupDocLifecycleLog');
    return;
  }
  const docId = String(payload.doc_id || '').trim();
  if (!docId) return;

  const data = sheet.getDataRange().getValues();
  let foundRow = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === docId) { foundRow = i + 1; break; }
  }

  const now = new Date();
  const woIdsCsv = Array.isArray(payload.wo_ids)
    ? payload.wo_ids.filter(Boolean).join(', ')
    : (payload.wo_ids || '');

  if (foundRow === 0) {
    // Insert
    const row = [
      docId,
      payload.doc_type || '',
      payload.anchor   || '',
      payload.contractor   || '',
      payload.contract_num || '',
      payload.borough      || '',
      woIdsCsv,
      payload.done ? 'Yes' : '',
      payload.sent ? 'Yes' : '',
      payload.file_id || '',
      payload.done ? now : '',
      payload.sent ? now : '',
      payload.notes || '',
    ];
    sheet.appendRow(row);
    return;
  }

  // Update existing row in place. Only mutate explicitly-supplied fields.
  const r = data[foundRow - 1];
  const updates = [];
  if (payload.doc_type   !== undefined) updates.push({ col: 2,  val: payload.doc_type });
  if (payload.anchor     !== undefined) updates.push({ col: 3,  val: payload.anchor });
  if (payload.contractor !== undefined) updates.push({ col: 4,  val: payload.contractor });
  if (payload.contract_num !== undefined) updates.push({ col: 5, val: String(payload.contract_num).split('/')[0].trim() });
  if (payload.borough    !== undefined) updates.push({ col: 6,  val: payload.borough });
  if (payload.wo_ids     !== undefined) updates.push({ col: 7,  val: woIdsCsv });
  if (payload.done === true) {
    updates.push({ col: 8, val: 'Yes' });
    if (String(r[7] || '').toLowerCase() !== 'yes') updates.push({ col: 11, val: now });
  } else if (payload.done === false) {
    updates.push({ col: 8, val: '' });
    updates.push({ col: 11, val: '' });
  }
  if (payload.sent === true) {
    updates.push({ col: 9, val: 'Yes' });
    if (String(r[8] || '').toLowerCase() !== 'yes') updates.push({ col: 12, val: now });
  } else if (payload.sent === false) {
    updates.push({ col: 9, val: '' });
    updates.push({ col: 12, val: '' });
  }
  if (payload.file_id    !== undefined) updates.push({ col: 10, val: payload.file_id });
  if (payload.notes      !== undefined) updates.push({ col: 13, val: payload.notes });

  updates.forEach(u => sheet.getRange(foundRow, u.col).setValue(u.val));
}

/**
 * Flip Done and/or Sent on a row by Doc ID. If the row doesn't exist
 * yet (manual Mark Done on a fresh work-day item that hasn't been
 * archived or backfilled), derive the row's metadata — doc_type,
 * anchor, contractor, contract_num, borough, wo_ids — from the doc_id
 * + Work Day Log so we don't leave a phantom skeleton row in the Log.
 */
function _setDocLifecycleStatus_(ss, docId, flags) {
  if (!docId || !flags || (flags.done == null && flags.sent == null)) return;
  const payload = {
    doc_id: docId,
    done:   flags.done,
    sent:   flags.sent,
  };
  const { byId } = _readDocLifecycle_(ss);
  if (!byId[docId]) {
    const meta = _resolveDocLifecycleMetadata_(ss, docId);
    if (meta) Object.assign(payload, meta);
  }
  _upsertDocLifecycleRow_(ss, payload);
}

/**
 * Reverse-engineer doc-row metadata from a synthetic Doc ID + the
 * Work Day Log. Used when a manual Mark Done flips a doc whose Log
 * row doesn't exist yet — without this helper the upsert would
 * append a row with only Doc ID / Done / timestamp populated.
 *
 * Returns null on unparseable doc_id.
 */
function _resolveDocLifecycleMetadata_(ss, docId) {
  // PL — PL_<date>_<contractor_slug>
  const plParsed = _parsePLDocId_(docId);
  if (plParsed) {
    const day = _scanWorkDayLogForDate_(ss, plParsed.anchor);
    const woIds = Array.from(new Set(
      day.filter(r => r.contractor === plParsed.contractor).map(r => r.wo_id).filter(Boolean)
    ));
    return {
      doc_type:   'Production Log',
      anchor:     plParsed.anchor,
      contractor: plParsed.contractor,
      wo_ids:     woIds,
    };
  }
  // SI / CP — <prefix>_<date-or-week>_<contractnum>_<borough>
  const stdParsed = _parseDocLifecycleId_(docId);
  if (!stdParsed) return null;
  if (stdParsed.prefix === 'SI') {
    const matches = _scanWorkDayLogForDate_(ss, stdParsed.anchor)
      .filter(r => r.contract_num === stdParsed.contractNum && r.borough === stdParsed.borough);
    return {
      doc_type:     'Sign-In',
      anchor:       stdParsed.anchor,
      contractor:   matches.length > 0 ? matches[0].contractor : '',
      contract_num: stdParsed.contractNum,
      borough:      stdParsed.borough,
      wo_ids:       Array.from(new Set(matches.map(r => r.wo_id).filter(Boolean))),
    };
  }
  if (stdParsed.prefix === 'CP') {
    const weekStart = new Date(stdParsed.anchor + 'T12:00:00');
    const wos = getWOsForPayrollWeek_(stdParsed.contractNum, stdParsed.borough, weekStart, ss);
    return {
      doc_type:     'Certified Payroll',
      anchor:       stdParsed.anchor,
      contractor:   wos.length > 0 ? (wos[0].contractor || '') : '',
      contract_num: stdParsed.contractNum,
      borough:      stdParsed.borough,
      wo_ids:       wos.map(w => w.id),
    };
  }
  return null;
}

/**
 * Read every Work Day Log row matching the given ISO date.
 * Returns [{ wo_id, contractor, contract_num, borough }, ...].
 */
function _scanWorkDayLogForDate_(ss, dateIso) {
  const sheet = ss.getSheetByName('Work Day Log');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const mm = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return mm ? mm[1] : '';
  };
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (ymd(r[0]) !== dateIso) continue;
    out.push({
      wo_id:        String(r[1] || '').trim(),
      contractor:   String(r[2] || '').trim(),
      contract_num: String(r[3] || '').split('/')[0].trim(),
      borough:      String(r[4] || '').trim(),
    });
  }
  return out;
}

/**
 * Inverse of _docLifecycleId_. Pulls (prefix, anchor, contractNum, borough)
 * out of a synthetic Doc ID. Returns null on malformed input. Used for
 * SI and CP IDs only — PL rows use _plDocId_ / _parsePLDocId_ since
 * their natural unit is (date, contractor), not (date, contract, borough).
 *
 *   "SI_2026-04-13_84125MBTP701_BK" → { prefix:'SI', anchor:'2026-04-13',
 *                                       contractNum:'84125MBTP701', borough:'BK' }
 */
function _parseDocLifecycleId_(docId) {
  // Trailing `_chief-<slug>` is optional — present on multi-crew rows,
  // absent on legacy and CP rows. Capture and return so callers can
  // route by crew.
  const m = String(docId || '').match(
    /^(PL|SI|CP)_(\d{4}-\d{2}-\d{2})_(.+)_([A-Z]{1,2})(?:_chief-([A-Za-z0-9]+))?$/
  );
  if (!m) return null;
  return {
    prefix:      m[1],
    anchor:      m[2],
    contractNum: m[3],
    borough:     m[4],
    crew_chief:  m[5] || '',
  };
}

/**
 * Build a per-(date, contractor) PL Doc ID. The PL file is per-contractor-
 * per-day (one PL covers all that contractor's contracts), so the Log row
 * shape mirrors that. Slug matches the existing PL filename slug used by
 * generateProductionLog_ so the ID round-trips with archive/file paths.
 *
 *   _plDocId_('2026-05-08', 'Metro Express') → 'PL_2026-05-08_Metro_Express'
 */
function _plDocId_(anchorIso, contractor, crewChief) {
  const slug = String(contractor || '').trim().replace(/\s+/g, '_');
  const base = 'PL_' + String(anchorIso).trim() + '_' + slug;
  // Optional per-crew suffix. Multi-crew shifts get one PL per chief.
  // Blank chief omits the suffix → matches legacy doc IDs.
  const chiefSlug = String(crewChief || '').replace(/[^A-Za-z0-9]/g, '');
  return chiefSlug ? (base + '_chief-' + chiefSlug) : base;
}

/**
 * Inverse of _plDocId_. Returns null on malformed input. Optional
 * trailing `_chief-<slug>` is captured separately so the contractor
 * name doesn't accidentally pick up the suffix.
 *
 *   "PL_2026-05-08_Metro_Express" → { anchor:'2026-05-08', contractor:'Metro Express', crew_chief:'' }
 *   "PL_2026-05-08_Metro_Express_chief-BobSmith" → { …, crew_chief:'BobSmith' }
 */
function _parsePLDocId_(docId) {
  const s = String(docId || '');
  const chiefMatch = s.match(/^(PL_\d{4}-\d{2}-\d{2}_.+)_chief-([A-Za-z0-9]+)$/);
  const base = chiefMatch ? chiefMatch[1] : s;
  const crewChief = chiefMatch ? chiefMatch[2] : '';
  const m = base.match(/^PL_(\d{4}-\d{2}-\d{2})_(.+)$/);
  if (!m) return null;
  return { anchor: m[1], contractor: m[2].replace(/_/g, ' '), crew_chief: crewChief };
}

/**
 * Detector: is this an OLD-format PL doc_id (PL_<date>_<cn>_<bor>)?
 * Used by the migration to distinguish stale rows from new-format rows.
 * Strips optional `_chief-<slug>` before matching so multi-crew variants
 * of the old format are also flagged for migration.
 */
function _isOldFormatPLDocId_(docId) {
  const stripped = String(docId || '').replace(/_chief-[A-Za-z0-9]+$/, '');
  return /^PL_\d{4}-\d{2}-\d{2}_.+_[A-Z]{1,2}$/.test(stripped);
}

/**
 * Look up the contractor for a (date, contractNum, borough) tuple
 * via Work Day Log. Returns the contractor string or '' if no match.
 *
 * WDL columns (0-idx): 0=Date, 1=WO, 2=Contractor, 3=Contract #, 4=Borough.
 */
function _lookupContractorForDateContractBorough_(ss, dateIso, contractNum, borough) {
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (!wdlSheet) return '';
  const data = wdlSheet.getDataRange().getValues();
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const mm = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return mm ? mm[1] : '';
  };
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (ymd(r[0]) !== dateIso) continue;
    const cn = String(r[3] || '').split('/')[0].trim();
    if (cn !== contractNum) continue;
    if (String(r[4] || '').trim() !== borough) continue;
    return String(r[2] || '').trim();
  }
  return '';
}

/**
 * Validate that the Sign-In sheet for a single (date, contract, borough)
 * tuple is marked Done in the Doc Lifecycle Log. Returns:
 *   { ok: true }
 *   { ok: false, error_code: 'SI_NOT_DONE', error: '...' }
 */
function _validateSignInDoneForDay_(ss, dateIso, contractNum, borough) {
  const siId = _docLifecycleId_('Sign-In', dateIso, contractNum, borough);
  const { byId } = _readDocLifecycle_(ss);
  const row = byId[siId];
  if (row && row.done) return { ok: true };
  return {
    ok: false,
    error_code: 'SI_NOT_DONE',
    error: 'Sign-In sheet is not done for ' + dateIso + ' · ' + contractNum + ' · ' + borough +
           '. Ensure the Sign-In sheet is complete, then retry.',
  };
}

/**
 * Sign-In lifecycle rows are keyed per crew chief (…_chief-<name>); a
 * blank chief omits the suffix (legacy single-crew rows). Returns true
 * iff THIS crew's Sign-In row for the tuple is marked Done. The CP/PL
 * gates must check per crew, not one no-chief id per date — otherwise a
 * sheet submitted by a named crew chief is never found and generation is
 * wrongly blocked.
 */
function _signInRowDone_(byId, dateIso, contractNum, borough, crewChief) {
  const id  = _docLifecycleId_('Sign-In', dateIso, contractNum, borough, crewChief);
  const row = byId[id];
  return !!(row && row.done);
}

/**
 * Validate that every Sign-In sheet for a (week, contract, borough) is Done.
 * Walks Work Day Log to find every (date) within [weekStart, weekStart+6]
 * that worked for this contract+borough, then checks each one's SI row.
 *
 * Returns:
 *   { ok: true }
 *   { ok: false, error_code: 'SI_NOT_DONE', error: '...', missing_dates: [...] }
 */
function _validateSignInDoneForWeek_(ss, weekStartIso, contractNum, borough) {
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (!wdlSheet) {
    return { ok: false, error_code: 'NO_WORK_DAY_LOG', error: 'Work Day Log sheet missing.' };
  }
  const data = wdlSheet.getDataRange().getValues();
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const mm = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return mm ? mm[1] : '';
  };
  const weekEndIso = (function () {
    const [y, m, d] = weekStartIso.split('-').map(Number);
    const end = new Date(y, m - 1, d + 6);
    return Utilities.formatDate(end, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  })();
  // Collect every (date, crew chief) that worked this contract+borough in
  // the week. Sign-In rows are keyed per crew chief, so each crew's sheet
  // is validated — not one no-chief id per date.
  const workedCrews = new Map();   // `${date}|${chief}` → { date, chief }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const dateIso = ymd(r[0]);
    if (!dateIso || dateIso < weekStartIso || dateIso > weekEndIso) continue;
    const cn = String(r[3] || '').split('/')[0].trim();
    if (cn !== contractNum) continue;
    if (String(r[4] || '').trim() !== borough) continue;
    const chief = String(r[7] || '').trim();
    workedCrews.set(dateIso + '|' + chief, { date: dateIso, chief });
  }
  if (workedCrews.size === 0) {
    return {
      ok: false,
      error_code: 'NO_WORK',
      error: 'No work logged for ' + contractNum + ' · ' + borough +
             ' during week of ' + weekStartIso + '. Nothing to put on a Certified Payroll.',
    };
  }
  const { byId } = _readDocLifecycle_(ss);
  const missing = [];
  Array.from(workedCrews.values())
    .sort((a, b) => (a.date + '|' + a.chief).localeCompare(b.date + '|' + b.chief))
    .forEach(({ date, chief }) => {
      if (!_signInRowDone_(byId, date, contractNum, borough, chief)) {
        missing.push(chief ? date + ' (' + chief + ')' : date);
      }
    });
  if (missing.length === 0) return { ok: true };
  return {
    ok: false,
    error_code: 'SI_NOT_DONE',
    error: 'Sign-In sheets not done for: ' + missing.join(', ') +
           ' (' + contractNum + ' · ' + borough + '). Ensure each Sign-In sheet is complete, then retry.',
    missing_dates: missing,
  };
}

/**
 * One-time backfill: walks Work Day Log to enumerate every (date,
 * contract, borough) tuple where work happened. For each tuple, inserts:
 *   - one Production Log row in Doc Lifecycle Log
 *   - one Sign-In row
 * Plus, for each unique (week_start, contract, borough) tuple,
 * one Certified Payroll row.
 *
 * Done and Sent columns are LEFT BLANK. The admin works through them
 * manually via the Doc Status tab's pending list — automated status
 * inference would pollute results given the user has test docs.
 *
 * Idempotent: skips Doc IDs already present in the Log. Safe to re-run.
 *
 * Public name (no underscore) so the editor function-runner picks it up.
 */
function backfillDocLifecycleLog() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const logSheet = ss.getSheetByName(DOC_LIFECYCLE_LOG_SHEET_);
  if (!logSheet) {
    Logger.log('❌ backfillDocLifecycleLog: Doc Lifecycle Log sheet not found — run setupDocLifecycleLog first');
    return;
  }
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (!wdlSheet) {
    Logger.log('❌ backfillDocLifecycleLog: Work Day Log not found — backfill source unavailable');
    return;
  }

  // Set of existing Doc IDs so re-runs skip duplicates.
  const { byId: existingById } = _readDocLifecycle_(ss);

  // Walk WDL once. SI groups by (date, contractor, contract, borough);
  // PL groups by (date, contractor) since the file is per-contractor;
  // CP weeks group by (week_start, contract, borough).
  const wdlData = wdlSheet.getDataRange().getValues();
  const siGroups = {};    // key: date|contractor|contractNum|borough → { contractor, contractNum, borough, anchor, wo_ids: Set }
  const plGroups = {};    // key: date|contractor → { contractor, anchor, wo_ids: Set }
  const weekTuples = {};  // key: weekStartIso|contractNum|borough → { contractor, contractNum, borough, anchor, wo_ids: Set }

  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };
  // Sunday-anchored week start for a given Date.
  const weekStartIsoFor = (date) => {
    const d = new Date(date);
    const dow = d.getDay();   // 0 = Sunday
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate() - dow);
    return Utilities.formatDate(start, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  };

  for (let i = 1; i < wdlData.length; i++) {
    const r = wdlData[i];
    const dateIso = ymd(r[0]);
    if (!dateIso) continue;
    const woId        = String(r[1] || '').trim();
    const contractor  = String(r[2] || '').trim();
    const contractNum = String(r[3] || '').split('/')[0].trim();
    const borough     = String(r[4] || '').trim();
    // WDL post-multi-crew schema: col 7 = Crew Chief. Blank for legacy
    // rows (pre-migration); groups with itself per the fallback rules.
    const crewChief   = String(r[7] || '').trim();
    if (!woId || !contractor || !contractNum || !borough) continue;

    // SI key includes crew chief so two crews on the same source job
    // produce two separate pending SI lifecycle rows (Doc Status shows
    // each as its own card).
    const siKey = dateIso + '|' + contractor + '|' + contractNum + '|' + borough + '|' + crewChief;
    if (!siGroups[siKey]) {
      siGroups[siKey] = {
        anchor: dateIso, contractor, contractNum, borough, crewChief, wo_ids: new Set(),
      };
    }
    siGroups[siKey].wo_ids.add(woId);

    // PL key includes chief so a contractor running two crews on the
    // same date produces two separate pending PL lifecycle rows.
    const plKey = dateIso + '|' + contractor + '|' + crewChief;
    if (!plGroups[plKey]) {
      plGroups[plKey] = {
        anchor: dateIso, contractor, crewChief, wo_ids: new Set(),
      };
    }
    plGroups[plKey].wo_ids.add(woId);

    // Week tuple uses Sunday anchor of the day's date. CP stays single-
    // crew per (week, contract, borough) by design — workers are
    // reported per (employee, classification), not per crew.
    const wkStart = weekStartIsoFor(new Date(dateIso + 'T12:00:00'));
    const weekKey = wkStart + '|' + contractNum + '|' + borough;
    if (!weekTuples[weekKey]) {
      weekTuples[weekKey] = {
        anchor: wkStart, contractor, contractNum, borough, wo_ids: new Set(),
      };
    }
    weekTuples[weekKey].wo_ids.add(woId);
  }

  let plInserted = 0, siInserted = 0, cpInserted = 0, skipped = 0;

  Object.values(plGroups).forEach(g => {
    const plId = _plDocId_(g.anchor, g.contractor, g.crewChief);
    if (!existingById[plId]) {
      _upsertDocLifecycleRow_(ss, {
        doc_id: plId, doc_type: 'Production Log', anchor: g.anchor,
        contractor: g.contractor,
        crew_chief: g.crewChief,
        wo_ids: Array.from(g.wo_ids),
      });
      plInserted++;
    } else skipped++;
  });

  Object.values(siGroups).forEach(g => {
    const siId = _docLifecycleId_('Sign-In', g.anchor, g.contractNum, g.borough, g.crewChief);
    if (!existingById[siId]) {
      _upsertDocLifecycleRow_(ss, {
        doc_id: siId, doc_type: 'Sign-In', anchor: g.anchor,
        contractor: g.contractor, contract_num: g.contractNum, borough: g.borough,
        crew_chief: g.crewChief,
        wo_ids: Array.from(g.wo_ids),
      });
      siInserted++;
    } else skipped++;
  });

  Object.values(weekTuples).forEach(w => {
    const woArr = Array.from(w.wo_ids);
    const cpId = _docLifecycleId_('Certified Payroll', w.anchor, w.contractNum, w.borough);
    if (!existingById[cpId]) {
      _upsertDocLifecycleRow_(ss, {
        doc_id: cpId, doc_type: 'Certified Payroll', anchor: w.anchor,
        contractor: w.contractor, contract_num: w.contractNum, borough: w.borough,
        wo_ids: woArr,
      });
      cpInserted++;
    } else skipped++;
  });

  _invalidateCacheKeys_([]);   // bump cache token so any UI re-fetches see the new rows
  Logger.log('✅ backfillDocLifecycleLog: PL=' + plInserted +
             ', SI=' + siInserted +
             ', CP=' + cpInserted +
             ', skippedExisting=' + skipped);
}

/**
 * One-time migration: consolidate old-format PL Doc Lifecycle Log rows
 * (one per (date, contract, borough)) into new-format rows (one per
 * (date, contractor)). The PL file has always been per-contractor —
 * the Log just had the wrong shape, copy-pasted from SI/CP.
 *
 * Existing Done/Sent state is preserved via OR-merge:
 *   merged.done = ANY old row's done === true
 *   merged.sent = ANY old row's sent === true
 *   merged.done_at = earliest non-empty
 *   merged.sent_at = earliest non-empty
 *   merged.file_id = first non-empty (all old rows for one PL share file_id)
 *   merged.wo_ids  = union, deduped
 *   merged.notes   = first non-empty
 *
 * Idempotent: if a new-format row already exists for (date, contractor),
 * it's OR-merged with the consolidated state from the old rows so a
 * second run never downgrades anything.
 *
 * Public name (no underscore) — surfaces in the Apps Script editor.
 */
function migrateConsolidatePLRows() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DOC_LIFECYCLE_LOG_SHEET_);
  if (!sheet) {
    Logger.log('❌ migrateConsolidatePLRows: Doc Lifecycle Log sheet not found');
    return;
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('ℹ️ migrateConsolidatePLRows: Log is empty — nothing to do.');
    return;
  }

  // Group old-format PL rows by (anchor, contractor). Track sheet row
  // numbers so we can delete them after consolidation.
  // Columns (1-idx): A=Doc ID, B=Doc Type, C=Anchor, D=Contractor,
  //                  E=Contract#, F=Borough, G=WO IDs, H=Done, I=Sent,
  //                  J=File ID, K=Done At, L=Sent At, M=Notes.
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };
  const earlierDate = (a, b) => {
    if (!a) return b;
    if (!b) return a;
    const ta = (a instanceof Date) ? a.getTime() : new Date(a).getTime();
    const tb = (b instanceof Date) ? b.getTime() : new Date(b).getTime();
    return ta <= tb ? a : b;
  };

  const groups = {};                    // key: anchor|contractor → merge state
  const oldRowsToDelete = [];           // sheet row numbers (1-idx, with header)

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const docId    = String(r[0] || '').trim();
    const docType  = String(r[1] || '').trim();
    if (docType !== 'Production Log') continue;
    if (!_isOldFormatPLDocId_(docId)) continue;

    const anchor     = ymd(r[2]);
    const contractor = String(r[3] || '').trim();
    if (!anchor || !contractor) continue;

    const k = anchor + '|' + contractor;
    if (!groups[k]) {
      groups[k] = {
        anchor, contractor,
        done: false, sent: false,
        done_at: '', sent_at: '',
        file_id: '',
        wo_ids: new Set(),
        notes: '',
      };
    }
    const g = groups[k];
    if (String(r[7] || '').toLowerCase() === 'yes') g.done = true;
    if (String(r[8] || '').toLowerCase() === 'yes') g.sent = true;
    if (r[10]) g.done_at = earlierDate(g.done_at, r[10]);
    if (r[11]) g.sent_at = earlierDate(g.sent_at, r[11]);
    if (!g.file_id && r[9]) g.file_id = String(r[9]).trim();
    String(r[6] || '').split(',').map(s => s.trim()).filter(Boolean).forEach(id => g.wo_ids.add(id));
    if (!g.notes && r[12]) g.notes = String(r[12]).trim();

    oldRowsToDelete.push(i + 1);   // 1-idx with header
  }

  // Also OR-merge with existing new-format rows so re-runs preserve
  // anything the admin toggled after a partial migration.
  const { byId: existingById } = _readDocLifecycle_(ss);
  let written = 0, skipped = 0;
  Object.values(groups).forEach(g => {
    const newId = _plDocId_(g.anchor, g.contractor);
    const existing = existingById[newId];
    if (existing) {
      if (existing.done) g.done = true;
      if (existing.sent) g.sent = true;
      if (existing.done_at) g.done_at = earlierDate(g.done_at, existing.done_at);
      if (existing.sent_at) g.sent_at = earlierDate(g.sent_at, existing.sent_at);
      if (!g.file_id && existing.file_id) g.file_id = existing.file_id;
      (existing.wo_ids || []).forEach(id => g.wo_ids.add(id));
      if (!g.notes && existing.notes) g.notes = existing.notes;
    }
    _upsertDocLifecycleRow_(ss, {
      doc_id:     newId,
      doc_type:   'Production Log',
      anchor:     g.anchor,
      contractor: g.contractor,
      // contract_num + borough deliberately blank — PL row spans them.
      wo_ids:     Array.from(g.wo_ids),
      done:       g.done,
      sent:       g.sent,
      file_id:    g.file_id,
      notes:      g.notes,
    });
    // _upsertDocLifecycleRow_ stamps Done At / Sent At only on blank →
    // Yes transitions, so explicitly write the earliest preserved
    // timestamps after the upsert.
    if (g.done_at || g.sent_at) {
      const written_id = newId;
      const fresh = _readDocLifecycle_(ss).byId[written_id];
      if (fresh && fresh.sheet_row) {
        if (g.done_at) sheet.getRange(fresh.sheet_row, 11).setValue(g.done_at);
        if (g.sent_at) sheet.getRange(fresh.sheet_row, 12).setValue(g.sent_at);
      }
    }
    written++;
  });

  // Delete old-format rows from the bottom up so row indices don't shift.
  oldRowsToDelete.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));

  _invalidateCacheKeys_([]);
  Logger.log('✅ migrateConsolidatePLRows: groups=' + Object.keys(groups).length +
             ', new-format rows written=' + written +
             ', old rows deleted=' + oldRowsToDelete.length);
}

/**
 * One-time migration: trim the now-obsolete PL/SI/CP Done/Sent columns
 * from the WO Tracker, since per-doc tracking moved to the Doc
 * Lifecycle Log. After this runs, Google Sheets shifts trailing
 * columns left and Invoice Done? lands at its new position — the
 * follow-up deploy must update DOC_TYPE_DONE_COL_['Invoice'] to match.
 *
 * **Run prerequisites:**
 *   1. setupDocLifecycleLog has been run.
 *   2. backfillDocLifecycleLog has been run (so existing PL/SI/CP
 *      state is captured before we lose the WO Tracker columns).
 *   3. You're prepared to follow up with a code deploy that updates
 *      DOC_TYPE_DONE_COL_['Invoice'] + trims ensureWoTrackerExtraCols_.
 *
 * Without #3, the Invoice Done chip on the WO Tracker tab will read
 * the wrong column until the deploy lands.
 *
 * Public name (no underscore) so it shows in the editor dropdown.
 */
function migrateRemoveObsoleteDocColumns() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) {
    Logger.log('❌ migrateRemoveObsoleteDocColumns: Work Order Tracker not found');
    return;
  }
  const logSheet = ss.getSheetByName(DOC_LIFECYCLE_LOG_SHEET_);
  if (!logSheet) {
    Logger.log('❌ migrateRemoveObsoleteDocColumns: Doc Lifecycle Log missing — run setupDocLifecycleLog + backfillDocLifecycleLog first');
    return;
  }
  if (logSheet.getLastRow() < 2) {
    Logger.log('❌ migrateRemoveObsoleteDocColumns: Doc Lifecycle Log is empty — run backfillDocLifecycleLog first');
    return;
  }

  // Header-driven trim: find whichever obsolete columns exist by name
  // and delete them right-to-left. The live block has historically
  // drifted from EXTRA_HEADERS's commented positions, so we don't
  // hardcode 1-idx values here.
  const OBSOLETE = new Set([
    'Prod Log Sent?',
    'Sign-In Done?',
    'Sign-In Sent?',
    'CP Done?',
    'CP Sent?',
  ]);
  const lastCol = woSheet.getLastColumn();
  const allHeaders = woSheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());

  const toDelete = [];
  for (let i = 0; i < allHeaders.length; i++) {
    if (OBSOLETE.has(allHeaders[i])) {
      toDelete.push({ col: i + 1, header: allHeaders[i] });
    }
  }
  if (toDelete.length === 0) {
    Logger.log('ℹ️ migrateRemoveObsoleteDocColumns: no obsolete columns found — nothing to delete.');
    return;
  }

  toDelete.sort((a, b) => b.col - a.col);
  toDelete.forEach(t => woSheet.deleteColumn(t.col));

  const newHeaders = woSheet.getRange(1, 1, 1, woSheet.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim());
  const newInvoiceIdx = newHeaders.indexOf('Invoice Done?');

  Logger.log('✅ migrateRemoveObsoleteDocColumns: deleted ' + toDelete.length +
             ' col(s): ' + toDelete.map(t => '"' + t.header + '" @ 1-idx ' + t.col).join(', ') +
             '. Invoice Done? now at 0-idx ' + newInvoiceIdx +
             '. Follow up with a code deploy that sets DOC_TYPE_DONE_COL_[\'Invoice\']' +
             ' = ' + newInvoiceIdx + ' and trims ensureWoTrackerExtraCols_ EXTRA_HEADERS.');

  _invalidateCacheKeys_(['dashboard_v1']);
}

/**
 * One-time backfill: walks Drive (read-only, via findWOFolder_) to
 * populate WO Tracker col 43 (Archive Folder URL) for any row where
 * it's currently empty. Run once from the Apps Script editor after
 * deploying the col-43 schema change. Subsequent archiving for new
 * WOs writes the URL inline in getWOFolder_.
 *
 * No trailing underscore on the name so the Apps Script editor
 * shows it in the function-runner dropdown.
 */
function backfillArchiveFolderUrls() {
  const ss        = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet   = ss.getSheetByName('Work Order Tracker');
  ensureWoTrackerExtraCols_(woSheet);

  const archiveId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_ID');
  if (!archiveId) {
    Logger.log('❌ backfillArchiveFolderUrls: ARCHIVE_ID property not set');
    return;
  }
  const archiveRoot = DriveApp.getFolderById(archiveId);

  const data = woSheet.getDataRange().getValues();
  let filled  = 0;
  let missing = 0;
  let already = 0;
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const woId = String(row[0] || '').trim();
    if (!woId) continue;
    if (row[42]) { already++; continue; }
    const folder = findWOFolder_(archiveRoot, woId, ss, data);
    if (folder) {
      woSheet.getRange(i + 1, 43).setValue(folder.getUrl());
      filled++;
    } else {
      missing++;
    }
  }
  Logger.log('✅ backfillArchiveFolderUrls: filled=' + filled +
             ', missing=' + missing + ', alreadyHad=' + already);
}


/**
 * One-shot backfill: geocodes existing WO Tracker rows that don't yet
 * have lat/lng populated. Two batches because Apps Script's 6-min
 * execution limit is generous but we'd rather be safe with the API's
 * sequential geocoding calls (~1.5s each × 25 = ~38s).
 *
 * Usage: in the Apps Script editor, pick `backfillWOGeocode_batch1`
 * from the function dropdown, hit Run. When it finishes, run
 * `backfillWOGeocode_batch2` for the next batch. The shared
 * `_backfillWOGeocode_` walks the Tracker top-down and processes only
 * rows where Latitude (col 46) is currently blank — so re-running
 * either batch is idempotent (already-geocoded rows are skipped).
 */
function backfillWOGeocode_batch1() { _backfillWOGeocode_(25); }
function backfillWOGeocode_batch2() { _backfillWOGeocode_(25); }
// Catch-all for any remaining unmapped WOs after the two fixed-size
// batches. Safe to re-run — _backfillWOGeocode_ skips rows that
// already have a Latitude. Cap of 100 keeps us well under the 6-minute
// Apps Script execution limit (100 WOs × ~2s/each ≈ 3.5 min).
function backfillWOGeocode_remaining() { _backfillWOGeocode_(100); }

function _backfillWOGeocode_(limit) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  ensureWoTrackerExtraCols_(woSheet);
  const data = woSheet.getDataRange().getValues();

  let processed = 0;
  let succeeded = 0;
  let warned    = 0;
  for (let i = 1; i < data.length && processed < limit; i++) {
    const r = data[i];
    const woId = String(r[0] || '').trim();
    if (!woId) continue;
    // Skip rows that already have a Latitude — geocodeWO_ side of
    // _persistGeocode_ leaves existing coords intact on warning, so
    // any non-empty col 46 means we've successfully placed this WO.
    if (r[45] !== '' && r[45] != null) continue;

    const d = {
      work_order_id: woId,
      location:      String(r[5]  || ''),
      from_street:   String(r[6]  || ''),
      to_street:     String(r[7]  || ''),
      borough:       String(r[3]  || ''),
    };
    try {
      const result = geocodeWO_(d, ss);
      _persistGeocode_(ss, woId, result);
      if (result.lat != null) {
        succeeded++;
        Logger.log(`✅ ${woId} → ${result.lat.toFixed(5)}, ${result.lng.toFixed(5)}`);
      } else {
        warned++;
        Logger.log(`⚠️ ${woId} → ${result.warning || 'no result'}`);
      }
    } catch (e) {
      warned++;
      Logger.log(`❌ ${woId} threw: ${e.message}`);
      try {
        _persistGeocode_(ss, woId, { warning: 'Backfill threw: ' + e.message });
      } catch (_) { /* nothing left to do */ }
    }
    processed++;
  }

  Logger.log(`Backfill batch done: processed=${processed}, succeeded=${succeeded}, warned=${warned}`);
}


/**
 * One-shot backfill: walks each WO's archive folder and writes
 * Done='Yes' on every doc-type lifecycle column whose corresponding
 * file already exists in Drive. Run once after deploying the new
 * Done/Sent schema.
 *
 * Detects each doc by filename prefix inside the WO folder (CFR_,
 * Invoice_, SignIn_) and inside the contract-level master subfolders
 * (Production Logs/, Certified Payroll/). Sent flags are NOT touched
 * — there's no reliable way to recover past send history; admin can
 * use the Download Documents modal to mark the unsent backlog.
 *
 * Public name (no trailing underscore) so it appears in the editor's
 * function-runner dropdown.
 */
function backfillDocLifecycleFlags() {
  const ss        = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet   = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) {
    Logger.log('❌ backfillDocLifecycleFlags: Work Order Tracker not found');
    return;
  }
  ensureWoTrackerExtraCols_(woSheet);

  const archiveId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_ID');
  if (!archiveId) {
    Logger.log('❌ backfillDocLifecycleFlags: ARCHIVE_ID property not set');
    return;
  }
  const archiveRoot = DriveApp.getFolderById(archiveId);

  // Cache contract-level master files (Production Logs / CP) so we
  // don't re-walk those folders for every WO. Keyed by
  // <contractor>|<contractNum>|<borough>.
  const masterCache = {};
  const getMaster = (contractor, contractNum, borough) => {
    const k = contractor + '|' + contractNum + '|' + borough;
    if (masterCache[k] != null) return masterCache[k];
    let prod = false, cp = false;
    try {
      const cIt = archiveRoot.getFoldersByName(contractor);
      if (!cIt.hasNext()) return (masterCache[k] = { prod: false, cp: false });
      const cFolder = cIt.next();
      const cnIt = cFolder.getFoldersByName(`${contractNum} - ${getBoroughName_(borough)}`);
      if (!cnIt.hasNext()) return (masterCache[k] = { prod: false, cp: false });
      const ctFolder = cnIt.next();
      const plIt = ctFolder.getFoldersByName('Production Logs');
      if (plIt.hasNext()) {
        const f = plIt.next().getFiles();
        if (f.hasNext()) prod = true;
      }
      const cpIt = ctFolder.getFoldersByName('Certified Payroll');
      if (cpIt.hasNext()) {
        const f = cpIt.next().getFiles();
        if (f.hasNext()) cp = true;
      }
    } catch (e) {
      Logger.log('⚠️ getMaster failed for ' + k + ': ' + e.message);
    }
    masterCache[k] = { prod, cp };
    return masterCache[k];
  };

  const data = woSheet.getDataRange().getValues();
  let touched = 0, scanned = 0, skipped = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const woId = String(row[0] || '').trim();
    if (!woId) continue;
    scanned++;

    const contractor  = String(row[1] || '').trim();
    const contractNum = String(row[2] || '').split('/')[0].trim();
    const borough     = String(row[3] || '').trim();
    if (!contractor || !contractNum || !borough) { skipped++; continue; }

    const folder = findWOFolder_(archiveRoot, woId, ss, data);
    if (!folder) { skipped++; continue; }

    // Detect per-WO docs in this WO's folder.
    //   - CFR: archived as the merged WO doc named `WO_<wo_id>.pdf`
    //     (or, for legacy folders that haven't had a CFR merged yet,
    //     ANY non-aux PDF — original scan). Mirrors the precedence in
    //     lookup_archived_wo_pdf.
    //   - Invoice: filename starts with Invoice_.
    //   - Sign-In: filename starts with SignIn_.
    const canonical = `WO_${woId}.pdf`;
    let cfrFound = false;
    let invFound = false;
    let signInFound = false;
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      if (name === canonical) {
        cfrFound = true;
      } else if (/^Invoice_/i.test(name)) {
        invFound = true;
      } else if (/^SignIn_/i.test(name)) {
        signInFound = true;
      } else if (f.getMimeType() === 'application/pdf' && !_isAuxDocName_(name)) {
        // Non-aux PDF = legacy WO doc (pre-merged-CFR) — counts as CFR.
        cfrFound = true;
      }
    }
    const masters = getMaster(contractor, contractNum, borough);

    const sheetRow = i + 1;
    const set = (col, val) => woSheet.getRange(sheetRow, col + 1).setValue(val);
    // PL/SI/CP per-WO columns no longer exist — those flags live in the
    // Doc Lifecycle Log now. This backfill stays narrow: just CFR + INV.
    if (cfrFound) set(DOC_TYPE_DONE_COL_['Field Report'], 'Yes');
    if (invFound) set(DOC_TYPE_DONE_COL_['Invoice'],      'Yes');
    if (cfrFound || invFound) touched++;
  }

  _invalidateCacheKeys_(['dashboard_v1']);
  Logger.log('✅ backfillDocLifecycleFlags: scanned=' + scanned +
             ', touched=' + touched + ', skipped=' + skipped);
}


/**
 * One-shot backfill: walks every contractor / contract folder under
 * Archive, scans the per-WO subfolders for SignIn_* PDFs, and copies
 * each unique one (by filename) up to a new contract-level
 * "Sign-Ins/" master folder. Brings existing data into the same
 * shape that archiveDocument_ now produces going forward.
 *
 * Public name (no trailing underscore) so it appears in the editor
 * function-runner dropdown. Idempotent — the per-filename dedupe
 * means re-running won't create duplicate masters.
 */
function backfillSignInMasters() {
  const props = PropertiesService.getScriptProperties();
  const archiveId = props.getProperty('ARCHIVE_ID');
  if (!archiveId) {
    Logger.log('❌ backfillSignInMasters: ARCHIVE_ID property not set');
    return;
  }
  const archiveRoot = DriveApp.getFolderById(archiveId);

  let contractorsScanned = 0;
  let contractsScanned   = 0;
  let mastersCreated     = 0;
  let copied             = 0;
  let skippedExisting    = 0;

  const contractorIt = archiveRoot.getFolders();
  while (contractorIt.hasNext()) {
    const contractorFolder = contractorIt.next();
    contractorsScanned++;
    const contractIt = contractorFolder.getFolders();
    while (contractIt.hasNext()) {
      const contractFolder = contractIt.next();
      // Skip the Archive Errors folder — never has WO subfolders.
      if (/Archive Errors/i.test(contractFolder.getName())) continue;
      contractsScanned++;

      // Build dedupe set from existing master (if any) so re-runs
      // don't duplicate.
      let masterFolder = null;
      const existingIt = contractFolder.getFoldersByName('Sign-Ins');
      if (existingIt.hasNext()) {
        masterFolder = existingIt.next();
      }
      const existing = new Set();
      if (masterFolder) {
        const exFiles = masterFolder.getFiles();
        while (exFiles.hasNext()) existing.add(exFiles.next().getName());
      }

      // Walk every WO subfolder of this contract; collect one copy per
      // unique SignIn_* filename. Skip the master folder itself + any
      // other non-WO containers (Production Logs, Certified Payroll).
      const woIt = contractFolder.getFolders();
      const filesToCopy = {};   // filename → File
      while (woIt.hasNext()) {
        const sub = woIt.next();
        const name = sub.getName();
        if (name === 'Sign-Ins' || name === 'Production Logs' || name === 'Certified Payroll') continue;
        const fIt = sub.getFiles();
        while (fIt.hasNext()) {
          const f = fIt.next();
          const fname = f.getName();
          if (!/^SignIn_/i.test(fname)) continue;
          if (existing.has(fname)) { skippedExisting++; continue; }
          if (filesToCopy[fname])  continue;   // dedupe across WO folders
          filesToCopy[fname] = f;
        }
      }

      const filenames = Object.keys(filesToCopy);
      if (filenames.length === 0) continue;

      if (!masterFolder) {
        masterFolder = contractFolder.createFolder('Sign-Ins');
        mastersCreated++;
      }
      filenames.forEach(fn => {
        try {
          filesToCopy[fn].makeCopy(fn, masterFolder);
          copied++;
        } catch (e) {
          Logger.log('⚠️ backfillSignInMasters: failed to copy ' + fn + ': ' + e.message);
        }
      });
    }
  }

  Logger.log('✅ backfillSignInMasters: contractors=' + contractorsScanned +
             ', contracts=' + contractsScanned +
             ', mastersCreated=' + mastersCreated +
             ', copied=' + copied +
             ', skippedExisting=' + skippedExisting);
}


/**
 * DEBUG / one-off: run this manually from the Apps Script editor to add
 * the 3 CFR columns to the WO Tracker immediately, without waiting for
 * a fresh WO scan.
 */
function addCFRColumnsNow() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Work Order Tracker');
  ensureWoTrackerCFRCols_(sheet);
  Logger.log('✅ CFR columns added/verified on Work Order Tracker.');
}


/**
 * DEBUG / one-off: run this manually to force the CFR export for a given
 * WO, bypassing handleSubmitFieldReport_ entirely. Edit WO_ID below, then
 * click Run. Will exercise aggregateMarkingItemsForCFR_ + generate the
 * JSON and tell you if anything throws.
 */
function debugGenerateCFRForWO() {
  const WO_ID = 'RM-43304';   // ← edit me

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const all = woSheet.getDataRange().getValues();
  const row = all.slice(1).find(r => String(r[0]) === String(WO_ID));
  if (!row) {
    Logger.log('❌ WO not found: ' + WO_ID);
    return;
  }

  const fakeD = {
    wo_id: WO_ID,
    date: new Date().toISOString().slice(0, 10),
    wo_complete: true,
  };

  // Aggregate first so we can see the shape
  const agg = aggregateMarkingItemsForCFR_(ss, WO_ID);
  Logger.log('Aggregated: ' + JSON.stringify(agg, null, 2));

  // Issues: pull from the tracker directly for this debug path
  const issues = String(row[22] || '').trim();

  try {
    generateContractorFieldReportJson_(fakeD, row, ss, issues);
    Logger.log('✅ CFR JSON exported for ' + WO_ID);
  } catch (err) {
    Logger.log('❌ CFR export failed: ' + err);
  }
}


/**
 * Flip "Water Blast Confirmed?" (WO Tracker col 14 / 0-idx 13) for a WO.
 * Used by the Field Report page to gate submission on MMA jobs — crew
 * toggles Yes when waterblasting has been completed, No to undo.
 *
 * body: { key, data: { wo_id, confirmed: boolean } }
 * Writes 'Yes' or 'No' to col 14 and logs to Automation Log.
 */
function handleSetWaterblastConfirmed_(body) {
  // The webapp's callAppsScript helper nests payloads under body.data;
  // fall back to top-level so direct curl / editor tests still work.
  const d         = body.data || body;
  const woId      = String(d.wo_id || '').trim();
  const confirmed = !!d.confirmed;
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);

  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();
  const idx     = allRows.findIndex((r, i) => i > 0 && String(r[0]) === woId);
  if (idx === -1) return jsonResponse_({ error: 'Work Order not found: ' + woId }, 404);

  const wbRequired = String(allRows[idx][12] || '');
  // Only MMA WOs have a meaningful confirm state. For Thermo we'd be stomping
  // the "N/A" sentinel, which would confuse the UI + downstream logic.
  if (wbRequired !== 'Yes - MMA') {
    return jsonResponse_({
      error: 'Waterblasting confirmation is only valid for MMA work orders (got water_blast_required="' + wbRequired + '")'
    }, 400);
  }

  const newValue = confirmed ? 'Yes' : 'No';
  // Sheet row is 1-indexed (header is row 1). WO Tracker col 13 (0-indexed)
  // is sheet col 14 (1-indexed).
  woSheet.getRange(idx + 1, 14).setValue(newValue);

  const logSheet = ss.getSheetByName('Automation Log');
  if (logSheet) {
    logSheet.appendRow([
      new Date(), 'Field Report', 'Waterblasting confirmation toggled',
      woId,
      `Water Blast Confirmed → ${newValue}`,
      'Completed', '', 'No'
    ]);
  }

  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({
    success: true,
    wo_id: woId,
    water_blast_confirmed: newValue,
  });
}


/**
 * Trigger the daily-document generator for a given date.  Used by the
 * webapp's Dashboard "Tools" menu now that the standalone script's
 * custom menu isn't firing reliably.
 *
 * body.data = { date: 'MM/DD/YYYY' | 'YYYY-MM-DD' | '' }
 *   Empty string → today.
 *
 * Returns { success, date_used, entries_found }.
 */
function handleGenerateDailyDocuments_(body) {
  const d       = body.data || {};
  const dateStr = String(d.date || '').trim();
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const targetDate = dateStr ? new Date(dateStr) : new Date();
    const dateIso = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');

    // SI validation — refuse if any PL-eligible contractor that worked
    // this day has any (contract, borough) tuple whose Sign-In isn't Done.
    // Same gate as the per-doc Generate button on the Doc Status tab.
    const v = _validateAllPLEligibleForDay_(ss, dateIso);
    if (!v.ok) {
      return jsonResponse_({ error: v.error, error_code: v.error_code, missing: v.missing || [] }, 400);
    }

    // generateDailyDocuments(dateStr) accepts anything Date parses
    // (MM/DD/YYYY, YYYY-MM-DD, empty → today).  No return value today,
    // so fetch the entry count ourselves for the UI.
    generateDailyDocuments(dateStr);
    const data = ss.getSheetByName('Daily Sign-In Data').getDataRange().getValues();
    let entries = 0;
    data.slice(1).forEach(row => {
      if (!row[0]) return;
      if (new Date(row[0]).toDateString() === targetDate.toDateString()) entries += 1;
    });
    return jsonResponse_({
      success:        true,
      date_used:      Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'MM/dd/yyyy'),
      entries_found:  entries,
    });
  } catch (err) {
    return jsonResponse_({ error: String(err && err.message || err) }, 500);
  }
}


/**
 * Manually trigger the approved-docs email+archive cron from the
 * webapp Dashboard's Tools menu.  Wraps processApprovedDocuments()
 * which already uses a script-wide lock, so this is safe to run
 * concurrently with the 10-min time-based trigger — whichever
 * grabs the lock first does the work; the other returns 0.
 *
 * Returns { success, archived, errored } for the toast summary.
 */
function handleProcessApprovedDocuments_(body) {
  try {
    const result = processApprovedDocuments() || {};
    return jsonResponse_({
      success:   true,
      archived:  result.archived || 0,
      errored:   result.errored || 0,
      skipped:   !!result.skipped,
    });
  } catch (err) {
    return jsonResponse_({ error: String(err && err.message || err) }, 500);
  }
}


/**
 * Trigger the certified-payroll generator for a given week-start
 * (Sunday, MM/DD/YYYY).  Webapp Dashboard "Tools" menu surface.
 *
 * body.data = { week_start: 'MM/DD/YYYY' }
 *
 * Returns { success, week_start, contract_groups }.
 */
function handleGenerateCertifiedPayroll_(body) {
  const d         = body.data || {};
  const weekStart = String(d.week_start || '').trim();
  if (!weekStart) return jsonResponse_({ error: 'Missing week_start' }, 400);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // Convert MM/DD/YYYY → YYYY-MM-DD for the validator
    const parts = weekStart.split('/');
    const weekStartIso = (parts.length === 3)
      ? Utilities.formatDate(
          new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]), 12, 0, 0),
          CONFIG.TIMEZONE, 'yyyy-MM-dd')
      : weekStart;

    // SI validation — refuse if any (contract, borough) that worked this
    // week has any day's Sign-In incomplete. Mirrors the per-doc gate.
    const v = _validateAllCPForWeek_(ss, weekStartIso);
    if (!v.ok) {
      return jsonResponse_({ error: v.error, error_code: v.error_code, missing: v.missing || [] }, 400);
    }

    const count = generateCertifiedPayroll(weekStart);
    return jsonResponse_({
      success:          true,
      week_start:       weekStart,
      contract_groups:  count,
    });
  } catch (err) {
    return jsonResponse_({ error: String(err && err.message || err) }, 500);
  }
}


/**
 * Per-doc Production Log generation. Triggered from a Doc Status tab
 * pending-list "Generate" button. Validates that EVERY (contract,
 * borough) the contractor worked that anchor day has its Sign-In Done
 * before firing — the PL file covers all of those contracts in one
 * document, so any missing SI means missing crew data for some WOs.
 *
 * body.data = { doc_id: 'PL_YYYY-MM-DD_<contractor_slug>' }
 */
function handleGeneratePLForDoc_(body) {
  const d     = body.data || {};
  const docId = String(d.doc_id || '').trim();
  if (!docId) return jsonResponse_({ error: 'Missing doc_id' }, 400);
  const parsed = _parsePLDocId_(docId);
  if (!parsed) {
    return jsonResponse_({ error: 'Malformed PL doc_id: ' + docId }, 400);
  }

  // Defense in depth — refuse if contractor isn't on the PL-required
  // list. Stale clients shouldn't surface a Generate button for them
  // anyway, but lock the door at the backend regardless.
  const enabled = new Set(
    (CONFIG.PRODUCTION_LOG_CONTRACTORS || []).map(s => String(s).trim()).filter(Boolean)
  );
  if (!enabled.has(parsed.contractor)) {
    return jsonResponse_({
      error: parsed.contractor + ' is not currently configured to require Production Logs.',
      error_code: 'PL_NOT_REQUIRED',
    }, 400);
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Validate every (contract, borough) the contractor worked that day
  // has its Sign-In Done. The PL covers them all, so any missing SI =
  // incomplete data on the resulting PL file.
  const v = _validateAllSIsDoneForContractorDay_(ss, parsed.contractor, parsed.anchor);
  if (!v.ok) {
    return jsonResponse_({
      error: v.error, error_code: v.error_code,
      missing: v.missing || [],
    }, 400);
  }

  try {
    // Convert ISO → MM/DD/YYYY (generateDailyDocuments parses both, but
    // the Date parser is more reliable on the MM/DD form).
    const [yyyy, mm, dd] = parsed.anchor.split('-');
    const dateMmDd = `${mm}/${dd}/${yyyy}`;
    const result = generateDailyDocuments(dateMmDd, {
      contractorFilter:           parsed.contractor,
      skipFieldReportsAndInvoices: true,
    });
    const files = (result && result.generated) || [];
    if (files.length === 0) {
      return jsonResponse_({
        error:      'No Production Log file was created. ' + parsed.contractor +
                    ' may not be enabled for PL generation, or no WOs were found for ' +
                    parsed.anchor + '.',
        error_code: 'NO_OUTPUT',
      }, 400);
    }
    return jsonResponse_({
      success: true,
      message: 'Production Log JSON queued for ' + parsed.contractor + ' on ' + parsed.anchor +
               '. Review the filled PDF in the Approvals tab — Done flips automatically when archived.',
      files,
    });
  } catch (err) {
    return jsonResponse_({ error: String(err && err.message || err) }, 500);
  }
}

/**
 * Validate every (contract, borough) the contractor worked that day
 * has its Sign-In marked Done in the Doc Lifecycle Log. Returns:
 *   { ok: true,  tuples: [{contractNum, borough}, ...] }
 *   { ok: false, error_code: 'SI_NOT_DONE', error: '...', missing: [...] }
 *   { ok: false, error_code: 'NO_WORK', error: '...' }
 */
function _validateAllSIsDoneForContractorDay_(ss, contractor, dateIso) {
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (!wdlSheet) {
    return { ok: false, error_code: 'NO_WORK_DAY_LOG', error: 'Work Day Log sheet missing.' };
  }
  const data = wdlSheet.getDataRange().getValues();
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const mm = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return mm ? mm[1] : '';
  };
  const tupleSet = new Set();
  const crews    = new Map();   // `${cn}|${bor}|${chief}` → { cn, bor, chief }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (ymd(r[0]) !== dateIso) continue;
    if (String(r[2] || '').trim() !== contractor) continue;
    const cn  = String(r[3] || '').split('/')[0].trim();
    const bor = String(r[4] || '').trim();
    if (!cn || !bor) continue;
    const chief = String(r[7] || '').trim();
    tupleSet.add(cn + '|' + bor);
    crews.set(cn + '|' + bor + '|' + chief, { cn, bor, chief });
  }
  if (tupleSet.size === 0) {
    return {
      ok: false,
      error_code: 'NO_WORK',
      error: contractor + ' has no Work Day Log entries for ' + dateIso + '.',
    };
  }
  const tuples = Array.from(tupleSet).map(s => {
    const [contractNum, borough] = s.split('|');
    return { contractNum, borough };
  });
  const { byId } = _readDocLifecycle_(ss);
  const missing = [];
  Array.from(crews.values()).forEach(({ cn, bor, chief }) => {
    if (!_signInRowDone_(byId, dateIso, cn, bor, chief)) {
      missing.push(cn + ' · ' + bor + (chief ? ' (' + chief + ')' : ''));
    }
  });
  if (missing.length === 0) return { ok: true, tuples };
  return {
    ok: false,
    error_code: 'SI_NOT_DONE',
    error: 'Sign-In sheet not complete for: ' + missing.join('; ') +
           ' on ' + dateIso + '. ' + contractor +
           '’s Production Log would be missing crew data — ensure each Sign-In is complete, then retry.',
    missing,
  };
}

/**
 * Fanout validator for the Tools-dropdown Production Log run.
 * Walks every PL-eligible contractor that worked this day, validates
 * each one's (contract, borough) Sign-Ins are Done. Refuses the entire
 * run if any are missing — matches the per-doc strict gate.
 */
function _validateAllPLEligibleForDay_(ss, dateIso) {
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (!wdlSheet) {
    return { ok: false, error_code: 'NO_WORK_DAY_LOG', error: 'Work Day Log sheet missing.' };
  }
  const enabled = new Set(
    (CONFIG.PRODUCTION_LOG_CONTRACTORS || []).map(s => String(s).trim()).filter(Boolean)
  );
  if (enabled.size === 0) return { ok: true, missing: [] };

  const data = wdlSheet.getDataRange().getValues();
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const mm = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return mm ? mm[1] : '';
  };
  // contractor → Map `${cn}|${bor}|${chief}` → { cn, bor, chief }. Sign-In
  // rows are keyed per crew chief, so collect the chief dimension too.
  const crewsByContractor = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (ymd(r[0]) !== dateIso) continue;
    const contractor = String(r[2] || '').trim();
    if (!enabled.has(contractor)) continue;
    const cn  = String(r[3] || '').split('/')[0].trim();
    const bor = String(r[4] || '').trim();
    if (!cn || !bor) continue;
    const chief = String(r[7] || '').trim();
    if (!crewsByContractor[contractor]) crewsByContractor[contractor] = new Map();
    crewsByContractor[contractor].set(cn + '|' + bor + '|' + chief, { cn, bor, chief });
  }
  if (Object.keys(crewsByContractor).length === 0) {
    // No PL-eligible contractor worked this day — nothing to validate.
    return { ok: true, missing: [] };
  }

  const { byId } = _readDocLifecycle_(ss);
  const missing = [];
  Object.keys(crewsByContractor).forEach(contractor => {
    Array.from(crewsByContractor[contractor].values()).forEach(({ cn, bor, chief }) => {
      if (!_signInRowDone_(byId, dateIso, cn, bor, chief)) {
        missing.push(contractor + ' · ' + cn + ' · ' + bor + (chief ? ' (' + chief + ')' : ''));
      }
    });
  });
  if (missing.length === 0) return { ok: true, missing: [] };
  return {
    ok: false,
    error_code: 'SI_NOT_DONE',
    error: 'Sign-In sheets not complete for ' + dateIso + ': ' + missing.join('; ') +
           '. Ensure each Sign-In is complete, then retry — the affected Production Logs would otherwise be missing crew data.',
    missing,
  };
}

/**
 * Fanout validator for the Tools-dropdown Certified Payroll run.
 * Walks every (contract, borough) that worked this week, validates
 * every worked day's Sign-In is Done. Refuses the entire run if any
 * are missing — matches the per-doc strict gate.
 */
function _validateAllCPForWeek_(ss, weekStartIso) {
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (!wdlSheet) {
    return { ok: false, error_code: 'NO_WORK_DAY_LOG', error: 'Work Day Log sheet missing.' };
  }
  const data = wdlSheet.getDataRange().getValues();
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const mm = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return mm ? mm[1] : '';
  };
  const weekEndIso = (function () {
    const [y, m, dd] = weekStartIso.split('-').map(Number);
    const end = new Date(y, m - 1, dd + 6);
    return Utilities.formatDate(end, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  })();
  // (cn|bor) → Map `${date}|${chief}` → { date, chief }. Sign-In rows are
  // keyed per crew chief, so collect the chief dimension too.
  const tupleCrews = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const dateIso = ymd(r[0]);
    if (!dateIso || dateIso < weekStartIso || dateIso > weekEndIso) continue;
    const cn  = String(r[3] || '').split('/')[0].trim();
    const bor = String(r[4] || '').trim();
    if (!cn || !bor) continue;
    const chief = String(r[7] || '').trim();
    const k = cn + '|' + bor;
    if (!tupleCrews[k]) tupleCrews[k] = new Map();
    tupleCrews[k].set(dateIso + '|' + chief, { date: dateIso, chief });
  }
  if (Object.keys(tupleCrews).length === 0) {
    // No work this week — nothing to validate, generation will simply produce 0.
    return { ok: true, missing: [] };
  }
  const { byId } = _readDocLifecycle_(ss);
  const missing = [];
  Object.keys(tupleCrews).forEach(k => {
    const [cn, bor] = k.split('|');
    Array.from(tupleCrews[k].values())
      .sort((a, b) => (a.date + '|' + a.chief).localeCompare(b.date + '|' + b.chief))
      .forEach(({ date, chief }) => {
        if (!_signInRowDone_(byId, date, cn, bor, chief)) {
          missing.push(cn + ' · ' + bor + ' · ' + (chief ? date + ' (' + chief + ')' : date));
        }
      });
  });
  if (missing.length === 0) return { ok: true, missing: [] };
  return {
    ok: false,
    error_code: 'SI_NOT_DONE',
    error: 'Sign-In sheets not complete for week of ' + weekStartIso + ': ' + missing.join('; ') +
           '. Ensure each Sign-In is complete, then retry — the affected Certified Payrolls would otherwise be missing hours.',
    missing,
  };
}


/**
 * Per-doc Certified Payroll generation. Triggered from a Doc Status tab
 * pending-list "Generate" button. Validates that every Sign-In sheet
 * for the week's worked days (this contract+borough only) is Done
 * before firing — refuses with a list of missing dates otherwise.
 *
 * body.data = { doc_id: 'CP_YYYY-MM-DD_<contractnum>_<borough>' }
 */
function handleGenerateCPForDoc_(body) {
  const d     = body.data || {};
  const docId = String(d.doc_id || '').trim();
  if (!docId) return jsonResponse_({ error: 'Missing doc_id' }, 400);
  const parsed = _parseDocLifecycleId_(docId);
  if (!parsed || parsed.prefix !== 'CP') {
    return jsonResponse_({ error: 'Malformed CP doc_id: ' + docId }, 400);
  }
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Validate that every worked day's SI is done
  const v = _validateSignInDoneForWeek_(ss, parsed.anchor, parsed.contractNum, parsed.borough);
  if (!v.ok) {
    return jsonResponse_({ error: v.error, error_code: v.error_code, missing_dates: v.missing_dates || [] }, 400);
  }

  try {
    const [yyyy, mm, dd] = parsed.anchor.split('-');
    const weekStartMmDd = `${mm}/${dd}/${yyyy}`;
    const groupCount = generateCertifiedPayroll(weekStartMmDd, {
      contractFilter: { contractNum: parsed.contractNum, borough: parsed.borough },
    });
    if (!groupCount) {
      return jsonResponse_({
        error:      'Certified Payroll generator produced 0 contract groups for ' +
                    parsed.contractNum + ' · ' + parsed.borough +
                    ' (week of ' + parsed.anchor + '). Check Daily Sign-In Data has rows in that range.',
        error_code: 'NO_OUTPUT',
      }, 400);
    }
    return jsonResponse_({
      success: true,
      message: 'Certified Payroll JSON queued for ' + parsed.contractNum + ' · ' +
               parsed.borough + ' (week of ' + parsed.anchor +
               '). Review the filled PDF in the Approvals tab — Done flips automatically when archived.',
      contract_groups: groupCount,
    });
  } catch (err) {
    return jsonResponse_({ error: String(err && err.message || err) }, 500);
  }
}


/**
 * Trash a Drive file by ID. Used by the Python worker after it fills a PDF
 * from a JSON payload — trashing the source JSON prevents the poll loop
 * from re-processing it if the worker's local `.processed_files.json`
 * state gets lost (e.g. Railway container restart with ephemeral disk).
 *
 * body: { key, file_id }
 */
function handleTrashFile_(body) {
  // Accept both the legacy flat shape (Python worker posts {key, file_id}
  // direct) and the Express-proxy shape (callAppsScript wraps data under
  // body.data). Without this fallback the photo-delete button in the
  // Field Report UI fires `Missing file_id` because the proxy nests its
  // payload one level deeper than the original Python caller did.
  const fileId = (body.data && body.data.file_id) || body.file_id;
  if (!fileId) return jsonResponse_({ error: 'Missing file_id' }, 400);
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return jsonResponse_({ success: true, file_id: fileId });
  } catch (err) {
    return jsonResponse_({ error: String(err) }, 500);
  }
}


/** Wraps an object as a JSON ContentService response. */
function jsonResponse_(obj, statusCode) {
  if (statusCode && statusCode !== 200) obj._status = statusCode;
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Append a row to the Automation Log sheet. Soft-fails if the sheet is
 * missing or append throws — logging must never take down the caller.
 * Column order matches the rest of the codebase:
 *   Timestamp | Source | Action | Subject | Details | Status | User | Action Needed
 */
function _logAutomation_(source, action, subject, details, status, actionNeeded) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Automation Log');
    if (!sheet) return;
    sheet.appendRow([
      new Date(),
      source   || '',
      action   || '',
      subject  || '',
      details  || '',
      status   || '',
      '',                         // User column — no auth yet
      actionNeeded || 'No'
    ]);
  } catch (err) {
    Logger.log('⚠️ _logAutomation_ failed: ' + (err && err.message || err));
  }
}


// ═══════════════════════════════════════════════════════════════
// MARKING ITEMS — seed on scan, read for field report
// ═══════════════════════════════════════════════════════════════

/**
 * Expand a Stop Msg / Stop Line cell value into a list of single-letter
 * directions. Handles:
 *   - Full English words: "North"/"East"/"South"/"West" → ["N"|"E"|"S"|"W"]
 *   - Letter strings:     "EW" → ["E","W"]; "NSEW" → ["N","S","E","W"]
 *   - Mixed/malformed:    filters to just N/S/E/W chars and de-dupes.
 * Returns [] for empty input.
 */
// ── Category → Unit map ──────────────────────────────────────
// Each Marking Type has a fixed unit of measure. Keep this in sync
// with webapp/src/lib/markingCategories.js CATEGORY_UNITS.
// Categories intentionally omitted (e.g. "Others") accept any unit.
const CATEGORY_UNITS_ = {
  // SF (square feet) — MMA area work
  'Bike Lane':           'SF',
  'Bus Lane':            'SF',
  'Pedestrian Space':    'SF',

  // LF (linear feet) — lines, crosswalks, stop lines
  'Double Yellow Line':  'LF',
  'Lane Lines':          'LF',
  'Solid Lines':         'LF',
  '4" Line':             'LF',
  '6" Line':             'LF',
  '8" Line':             'LF',
  '12" Line':            'LF',
  '16" Line':            'LF',
  '24" Line':            'LF',
  'Gores':               'LF',
  'HVX Crosswalk':       'LF',
  'Stop Line':           'LF',

  // EA (each / count) — messages, arrows, misc
  'Messages':            'EA',
  'Stop Msg':            'EA',
  'Only Msg':            'EA',
  'Bus Msg':             'EA',
  'Bump Msg':            'EA',
  'Custom Msg':          'EA',
  '20 MPH Msg':          'EA',
  'Railroad (RR)':       'EA',
  'Railroad (X)':        'EA',
  'Rail Road X/Diamond': 'EA',
  'Arrows':              'EA',
  'L/R Arrow':           'EA',
  'Straight Arrow':      'EA',
  'Combination Arrow':   'EA',
  'Combination Arrow (L/R)': 'EA',
  'Speed Hump Markings': 'EA',
  'Shark Teeth 12x18':   'EA',
  'Shark Teeth 24x36':   'EA',
  'Bike Lane Arrow':     'EA',
  'Bike Lane Symbol':    'EA',   // legacy — retired from picker, kept for old rows
  'Old Bike Symbol (w/ rider)':  'EA',
  'New Bike Symbol (just bike)': 'EA',
  'Pedestrian Men':      'EA',
  'Bike Lane Green Bar': 'SF',
  // "Others" is intentionally variable (user picks).
};

function unitForCategory_(category) {
  return CATEGORY_UNITS_[String(category || '').trim()] || '';
}


// ═══════════════════════════════════════════════════════════════
// PRICING ENGINE — Marking Type → revenue
// ═══════════════════════════════════════════════════════════════
//
// Lookup tables here are the single source of truth for pricing
// math. Mirror the SAME data (without the helpers) in
// webapp/src/lib/pricing.js so the React side can reference group
// names, unit counts, etc. for display purposes — but actual revenue
// math always runs server-side off these tables.
//
// Pricing groups:
//   line4         — base $/LF rate; multiplied by LINE_WIDTH_MULTIPLIER
//                   for non-4" widths and double-yellow.
//   line12        — flat $/LF; reserved for HVX Crosswalk + Stop Line
//                   (typically a bulk-discounted rate).
//   preformed     — $/Unit × PREFORMED_UNIT_COUNT (preformed thermo).
//   extruded      — $/Unit × EXTRUDED_UNIT_COUNT (standard contractor
//                   unit-count table — same across all contractors).
//   color_surface — flat $/SF for green/red surface treatments.
//   unpriced      — categories that always require manual pricing
//                   (Custom Msg, Others, parent categories like
//                   "Messages"/"Arrows" that should never carry a
//                   quantity in the first place).

const PRICING_GROUP_BY_CATEGORY_ = Object.freeze({
  // line4 (multiplied by width ratio)
  '4" Line':            'line4',
  '6" Line':            'line4',
  '8" Line':            'line4',
  '12" Line':           'line4',
  '16" Line':           'line4',
  '24" Line':           'line4',
  'Lane Lines':         'line4',
  'Double Yellow Line': 'line4',

  // line12 (Crosswalk + Stop Line bulk rate)
  'HVX Crosswalk':      'line12',
  'Stop Line':          'line12',

  // preformed thermoplastic
  'Bike Lane Symbol':   'preformed',   // legacy alias for Old Bike Symbol
  'Old Bike Symbol (w/ rider)':  'preformed',
  'New Bike Symbol (just bike)': 'preformed',
  'Pedestrian Men':     'preformed',

  // extruded thermoplastic — priced by unit count
  'Stop Msg':            'extruded',
  'Only Msg':            'extruded',
  'Bus Msg':             'extruded',
  'Bump Msg':            'extruded',
  '20 MPH Msg':          'extruded',
  'Railroad (RR)':       'extruded',
  'Railroad (X)':        'extruded',
  'L/R Arrow':           'extruded',
  'Straight Arrow':      'extruded',
  'Combination Arrow':   'extruded',
  'Combination Arrow (L/R)': 'extruded',
  'Speed Hump Markings': 'extruded',
  'Shark Teeth 12x18':   'extruded',
  'Shark Teeth 24x36':   'extruded',
  'Bike Lane Arrow':     'extruded',

  // color surface treatment ($/SF)
  'Bike Lane':           'color_surface',
  'Bus Lane':            'color_surface',
  'Pedestrian Space':    'color_surface',
  'Bike Lane Green Bar': 'color_surface',

  // always unpriced — flagged in Needs Pricing bucket
  'Custom Msg':          'unpriced',
  'Others':              'unpriced',
  'Solid Lines':         'unpriced',
  'Gores':               'unpriced',
  'Messages':            'unpriced',
  'Arrows':              'unpriced',
  'Rail Road X/Diamond': 'unpriced',
});

// Standard width ratios — line width / 4. Set in stone across
// contractors per the user's pricing convention.
const LINE_WIDTH_MULTIPLIER_ = Object.freeze({
  '4" Line':            1.0,
  '6" Line':            1.5,
  '8" Line':            2.0,
  '12" Line':           3.0,
  '16" Line':           4.0,
  '24" Line':           6.0,
  'Lane Lines':         1.0,
  'Double Yellow Line': 2.0,
});

// Standard NYC DOT thermo unit table. `null` = unit count not yet
// known — items will surface in the Needs Pricing bucket with
// reason='no_unit_count' until the table arrives.
//
// Message values are sums of the per-letter 8' Letters & Numbers unit
// counts (S=0.37, T=0.25, O=0.39, P=0.34, etc.). Worked examples for
// each message are shown alongside so a future edit can sanity-check
// against the same DOT table.
//   STOP   = S(0.37)+T(0.25)+O(0.39)+P(0.34)              = 1.35
//   ONLY   = O(0.39)+N(0.46)+L(0.25)+Y(0.25)              = 1.35
//   BUMP   = B(0.46)+U(0.36)+M(0.48)+P(0.34)              = 1.64
//   BUS    = B(0.46)+U(0.36)+S(0.37)                      = 1.19
//   20 MPH = 2(0.37)+0(0.39)+M(0.48)+P(0.34)+H(0.39)      = 1.97
//   RR     = R(0.41)+R(0.41)                              = 0.82
//   X      = X(0.31)                                      = 0.31
// Symbol/arrow values come from the Symbols (Extruded) table:
//   Turn Arrow            = 1.00   → mapped to 'L/R Arrow'
//   Through (Straight)    = 0.81   → mapped to 'Straight Arrow'
//   Combo Arrow           = 1.65   → mapped to 'Combination Arrow'
//   Bicycle Facility Arr  = 0.29   → mapped to 'Bike Lane Arrow'
//   Speed Hump Marking    = 0.78
//   Sharks Teeth 12x18    = 0.05
//   Sharks Teeth 24x36    = 0.19
const EXTRUDED_UNIT_COUNT_ = Object.freeze({
  'Stop Msg':            1.35,
  'Only Msg':            1.35,
  'Bus Msg':             1.19,
  'Bump Msg':            1.64,
  '20 MPH Msg':          1.97,
  'Railroad (RR)':       0.82,
  'Railroad (X)':        0.31,
  'L/R Arrow':           1.00,
  'Straight Arrow':      0.81,
  'Combination Arrow':   1.65,
  'Combination Arrow (L/R)': 1.74,
  'Speed Hump Markings': 0.78,
  'Shark Teeth 12x18':   0.05,
  'Shark Teeth 24x36':   0.19,
  'Bike Lane Arrow':     0.29,
});

const PREFORMED_UNIT_COUNT_ = Object.freeze({
  'Bike Lane Symbol': 0.91,   // legacy alias for Old Bike Symbol
  'Old Bike Symbol (w/ rider)':  0.91,
  'New Bike Symbol (just bike)': 0.97,
  'Pedestrian Men':   0.84,
});

// Line12-group multiplier. HVX Crosswalk is priced at the base 12"
// rate; Stop Line is a 24" stripe and bills at 2× the base rate.
const LINE12_MULTIPLIER_ = Object.freeze({
  'HVX Crosswalk': 1.0,
  'Stop Line':     2.0,
});

/**
 * Read the Contract Pricing sheet once and return parsed rate rows.
 * Each row has its rate fields coerced to Numbers (or null when blank
 * — distinguishes "no rate set" from "rate set to 0"). Effective Date
 * is normalized to a Date instance or null.
 */
function _loadContractPricing_(ss) {
  const sheet = ss.getSheetByName('Contract Pricing');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const contractor  = String(r[0] || '').trim();
    const contractNum = String(r[1] || '').trim();
    const borough     = String(r[2] || '').trim();
    if (!contractor || !contractNum || !borough) continue;
    const effRaw = r[3];
    let effDate = null;
    if (effRaw instanceof Date && !isNaN(effRaw.getTime())) effDate = effRaw;
    else if (typeof effRaw === 'string' && effRaw.trim()) {
      const d = new Date(effRaw);
      if (!isNaN(d.getTime())) effDate = d;
    }
    const num = (v) => {
      if (v === '' || v == null) return null;
      const n = Number(v);
      return isNaN(n) ? null : n;
    };
    rows.push({
      contractor,
      contract_num: contractNum.split('/')[0].trim(),
      borough,
      effective_date: effDate,
      rates: {
        line4:         num(r[4]),
        line12:        num(r[5]),
        preformed:     num(r[6]),
        extruded:      num(r[7]),
        color_surface: num(r[8]),
      },
      notes: String(r[9] || '').trim(),
    });
  }
  return rows;
}

/**
 * Pick the rate row that applies to a given (contractor, contract,
 * borough) on the supplied date. Dated rows beat blank-date rows when
 * both could match. Returns null when nothing applies — caller should
 * surface a `no_rate` reason in that case.
 */
function _resolveRateRow_(rates, contractor, contractNum, borough, dateIso) {
  if (!Array.isArray(rates) || rates.length === 0) return null;
  const cn = String(contractNum || '').split('/')[0].trim();
  const itemDate = (() => {
    if (!dateIso) return null;
    const d = new Date(dateIso);
    return isNaN(d.getTime()) ? null : d;
  })();

  const candidates = rates.filter(r =>
    r.contractor === contractor &&
    r.contract_num === cn &&
    r.borough === borough
  );
  if (candidates.length === 0) return null;

  // Dated candidates whose Effective Date <= item date.
  const dated = candidates.filter(r => r.effective_date != null);
  const datedApplicable = (itemDate
    ? dated.filter(r => r.effective_date.getTime() <= itemDate.getTime())
    : dated.slice()
  ).sort((a, b) => b.effective_date.getTime() - a.effective_date.getTime());
  if (datedApplicable.length > 0) return datedApplicable[0];

  // Fall back to a blank-date row (effective forever).
  const blank = candidates.find(r => r.effective_date == null);
  return blank || null;
}


/**
 * Load the Payroll Rates sheet into a simple list. Mirrors
 * _loadContractPricing_'s shape: one parsed row per sheet row, with
 * effective_date as a Date (or null when the cell is blank).
 *
 * Returns: [{
 *   classification: 'LP'|'SAT',
 *   effective_date: Date|null,
 *   st_rate, ot_rate, st_supp, ot_supp: number,
 *   notes: string,
 * }]
 */
function _loadPayrollRates_(ss) {
  const sheet = ss.getSheetByName('Payroll Rates');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const classification = String(r[0] || '').trim().toUpperCase();
    if (!classification) continue;
    const effRaw = r[1];
    let effDate = null;
    if (effRaw instanceof Date && !isNaN(effRaw.getTime())) effDate = effRaw;
    else if (typeof effRaw === 'string' && effRaw.trim()) {
      const d = new Date(effRaw);
      if (!isNaN(d.getTime())) effDate = d;
    }
    const num = (v) => {
      if (v === '' || v == null) return null;
      const n = Number(v);
      return isNaN(n) ? null : n;
    };
    rows.push({
      classification,
      effective_date: effDate,
      st_rate: num(r[2]),
      ot_rate: num(r[3]),
      st_supp: num(r[4]),
      ot_supp: num(r[5]),
      notes:   String(r[6] || '').trim(),
    });
  }
  return rows;
}

/**
 * Pick the rate row that applies to a given classification on the
 * supplied date (a payroll week-end date — per-week resolution).
 * Dated rows beat blank-date rows when both could match. Returns null
 * when nothing applies — caller should warn and skip the worker.
 */
function _resolvePayrollRate_(rates, classification, dateIso) {
  if (!Array.isArray(rates) || rates.length === 0) return null;
  const cls = String(classification || '').trim().toUpperCase();
  if (!cls) return null;
  const itemDate = (() => {
    if (!dateIso) return null;
    const d = (dateIso instanceof Date) ? dateIso : new Date(dateIso);
    return isNaN(d.getTime()) ? null : d;
  })();

  const candidates = rates.filter(r => r.classification === cls);
  if (candidates.length === 0) return null;

  const dated = candidates.filter(r => r.effective_date != null);
  const datedApplicable = (itemDate
    ? dated.filter(r => r.effective_date.getTime() <= itemDate.getTime())
    : dated.slice()
  ).sort((a, b) => b.effective_date.getTime() - a.effective_date.getTime());
  if (datedApplicable.length > 0) return datedApplicable[0];

  const blank = candidates.find(r => r.effective_date == null);
  return blank || null;
}


// ─── Billing remap (sub-prime work) ─────────────────────────────
// When a sub-prime is performing work on a contract they did NOT win
// (because the original prime is failing to perform), the work is
// billed under a contract the sub-prime DID win. Source data (WO
// Tracker, Daily Sign-In Data, Work Day Log) keeps the raw source
// contract+borough; billing-doc generation applies this remap so the
// Invoice / Sign-In Sheet / Certified Payroll content (and the
// Invoices & AR / CP Tracker rows) reflect the billing identity.
//
// Drive archive folders and Doc Lifecycle keys stay keyed by RAW so
// source-job artifacts (CFR, Production Log) co-locate with the
// billing docs in one folder per source job.
//
// Bandaid for now — if this becomes routine, lift into a sheet tab + UI.
// Borough values are stored as abbreviations everywhere in source data
// (WO Tracker, Daily Sign-In Data, Work Day Log, Contract Pricing,
// Contract Lookup) — see getBoroughName_ for the conversion table.
// Match those abbreviations exactly here, otherwise the remap silently
// no-ops and the bandaid does nothing.
const _BILLING_REMAP_ = [
  { contractNum: '84125MBTP701', borough: 'BK', contractor: 'Metro Express',
    bill_as: { contractNum: '84125MBTP701', borough: 'M' } },
  { contractNum: '84125MBTP701', borough: 'BK', contractor: 'Denville',
    bill_as: { contractNum: '84125MBTP701', borough: 'QU' } },
];

function _billingRemap_(contractNum, borough, contractor) {
  const cn = String(contractNum || '').trim();
  const br = String(borough || '').trim();
  const co = String(contractor || '').trim();
  const hit = _BILLING_REMAP_.find(r =>
    r.contractNum === cn && r.borough === br && r.contractor === co);
  return hit ? { ...hit.bill_as } : { contractNum: cn, borough: br };
}

// True when at least one remap rule targets this (raw contract, borough).
// Used by the CP generator to decide whether to split a bucket by
// contractor (each sub-prime gets their own CP billed under their own
// contract). Clean buckets just produce one CP as before.
function _hasBillingRemap_(contractNum, borough) {
  const cn = String(contractNum || '').trim();
  const br = String(borough || '').trim();
  return _BILLING_REMAP_.some(r => r.contractNum === cn && r.borough === br);
}


/**
 * Payroll week number per the rule "the Sun–Sat week containing Jan 1
 * is week 1, the next week is week 2, …". A week that straddles year-
 * end belongs to whichever year's Jan 1 it contains — so the week
 * Sun Dec 27, 2026 – Sat Jan 2, 2027 is week 1 of 2027, not week 53
 * of 2026.
 *
 * Implementation: try the *next* year's Jan 1 first (handles late-Dec
 * weeks whose Saturday falls in the new year), then the current year's.
 * Whichever year's "week 1 Sunday" is the latest one not after
 * weekStart owns this week.
 */
function _payrollWeekNumber_(weekStart) {
  const tryYear = (Y) => {
    const jan1 = new Date(Y, 0, 1);
    const dow  = jan1.getDay();         // 0=Sun … 6=Sat
    // Days back to the Sunday of Jan 1's Sun–Sat week.
    const w1Sun = new Date(Y, 0, 1 - dow);
    if (w1Sun.getTime() > weekStart.getTime()) return null;
    return Math.floor((weekStart.getTime() - w1Sun.getTime()) / 86400000 / 7) + 1;
  };
  const Y = weekStart.getFullYear();
  const next = tryYear(Y + 1);
  if (next != null) return next;
  const cur = tryYear(Y);
  if (cur != null) return cur;
  return 1;
}


/**
 * Compute revenue for one Marking Item. Returns
 *   { revenue: <number>, group: <string>, reason: <string|null> }
 *
 * `reason` semantics (precedence):
 *   'unit_migration'    — Bike Lane Green Bar entered as EA (legacy);
 *                         needs manual re-entry as SF before pricing.
 *   'unpriced_category' — category is in the always-manual bucket
 *                         (Custom Msg / Others / parent categories).
 *   'no_rate'           — no Contract Pricing row matches.
 *   'no_unit_count'     — extruded category whose unit count isn't
 *                         in the table yet.
 *   null                — priced cleanly. revenue is the dollar value.
 *
 * `item` shape: { category, quantity, unit, ... } from Marking Items.
 * `woMeta` shape: { contractor, contract_num, borough }.
 */
function priceMarkingItem_(item, woMeta, rates) {
  const cat = String(item.category || '').trim();
  const qty = Number(item.quantity);
  if (!cat || isNaN(qty) || qty <= 0) {
    return { revenue: 0, group: 'unpriced', reason: 'unpriced_category' };
  }

  // Bike Lane Green Bar legacy unit migration: anything still entered
  // as EA can't be priced as a $/SF surface treatment. Flag it instead
  // of producing a wrong number.
  if (cat === 'Bike Lane Green Bar' && String(item.unit || '').toUpperCase() === 'EA') {
    return { revenue: 0, group: 'color_surface', reason: 'unit_migration' };
  }

  const group = PRICING_GROUP_BY_CATEGORY_[cat] || 'unpriced';
  if (group === 'unpriced') {
    return { revenue: 0, group, reason: 'unpriced_category' };
  }

  const rateRow = _resolveRateRow_(
    rates,
    woMeta.contractor,
    woMeta.contract_num,
    woMeta.borough,
    item.date_completed
  );
  if (!rateRow) return { revenue: 0, group, reason: 'no_rate' };

  switch (group) {
    case 'line4': {
      const base = rateRow.rates.line4;
      if (base == null) return { revenue: 0, group, reason: 'no_rate' };
      const mult = LINE_WIDTH_MULTIPLIER_[cat];
      if (mult == null) return { revenue: 0, group, reason: 'no_unit_count' };
      return { revenue: qty * base * mult, group, reason: null };
    }
    case 'line12': {
      const base = rateRow.rates.line12;
      if (base == null) return { revenue: 0, group, reason: 'no_rate' };
      const mult = LINE12_MULTIPLIER_[cat];
      if (mult == null) return { revenue: 0, group, reason: 'no_unit_count' };
      return { revenue: qty * base * mult, group, reason: null };
    }
    case 'preformed': {
      const base  = rateRow.rates.preformed;
      const units = PREFORMED_UNIT_COUNT_[cat];
      if (base == null) return { revenue: 0, group, reason: 'no_rate' };
      if (units == null) return { revenue: 0, group, reason: 'no_unit_count' };
      return { revenue: qty * base * units, group, reason: null };
    }
    case 'extruded': {
      const base  = rateRow.rates.extruded;
      const units = EXTRUDED_UNIT_COUNT_[cat];
      if (base == null) return { revenue: 0, group, reason: 'no_rate' };
      if (units == null) return { revenue: 0, group, reason: 'no_unit_count' };
      return { revenue: qty * base * units, group, reason: null };
    }
    case 'color_surface': {
      const base = rateRow.rates.color_surface;
      if (base == null) return { revenue: 0, group, reason: 'no_rate' };
      return { revenue: qty * base, group, reason: null };
    }
    default:
      return { revenue: 0, group: 'unpriced', reason: 'unpriced_category' };
  }
}


function expandDirLetters_(val) {
  const s = String(val || '').trim();
  if (!s) return [];
  const FULL_WORD = { 'NORTH': 'N', 'EAST': 'E', 'SOUTH': 'S', 'WEST': 'W' };
  const upper = s.toUpperCase();
  if (FULL_WORD[upper]) return [FULL_WORD[upper]];
  // Treat as concatenated letters. Dedupe while preserving order.
  const seen = {};
  const out  = [];
  upper.split('').forEach(c => {
    if ('NSEW'.indexOf(c) === -1) return;
    if (seen[c]) return;
    seen[c] = true;
    out.push(c);
  });
  return out;
}


/**
 * Expand parsed WO top_markings + intersection_grid arrays into rows on
 * the Marking Items sheet. Called from handleWriteWO_ after the Tracker
 * row is appended.
 *
 * Row schema (17 cols) — set up by setupMarkingItems():
 *   A Item ID  B Work Order #  C Work Type  D Section  E Sort Order
 *   F Category  G Intersection  H Direction  I Description
 *   J Planned  K Unit  L Quantity Completed  M Color/Material
 *   N Date Completed  O Status  P Added By  Q Notes
 *
 * Expansion rules (per user requirement: one row per discrete marking):
 *   - Top Table: one row per non-empty {category, description}
 *   - Intersection Grid HVX columns (N/E/S/W): each non-empty cell = 1 row
 *   - Stop Msg / Stop Lines: split into multiple rows via expandDirLetters_
 *     ("EW" → 2 rows, "NSEW" → 4 rows, "West" → 1 row)
 *
 * @returns number of rows inserted (0 if nothing to seed or sheet missing).
 */
function seedMarkingItems_(ss, d) {
  const topMarkings = d.top_markings        || [];
  const grid        = d.intersection_grid   || [];
  const bikeMarks   = d.bike_lane_markings  || [];
  // PT-XXXXX = Paint → Color Surface / MMA WO; gets one auto-seeded Color
  // Surface item even when no other markings were parsed off the form.
  const isPaintWO   = String(d.work_order_id || '').toUpperCase().startsWith('PT-');
  if (topMarkings.length === 0 && grid.length === 0 && bikeMarks.length === 0 && !isPaintWO) return 0;

  const markingSheet = ss.getSheetByName('Marking Items');
  if (!markingSheet) {
    Logger.log('⚠️ Marking Items sheet not found — skip seeding. ' +
               'Run setupMarkingItems() from the menu first.');
    return 0;
  }

  const woId     = d.work_order_id;
  const workType = d.work_type || '';
  const pad3     = (x) => String(x).padStart(3, '0');
  const rows     = [];
  let n = 1;

  // ── Top Table items ───────────────────────────────────────────
  // Ordering is implicit from the order rows are pushed to the `rows`
  // array below → single setValues write → sheet row order. No sort
  // column needed.
  topMarkings.forEach((m) => {
    if (!m || !m.category || !m.description) return;
    const cat = String(m.category).trim();
    rows.push([
      `${woId}-${pad3(n++)}`,        // A  Item ID
      woId,                           // B  Work Order #
      workType,                       // C  Work Type
      'Top Table',                    // D  WO Section
      cat,                            // E  Marking Type
      '',                             // F  Intersection
      '',                             // G  Direction
      String(m.description).trim(),   // H  Description
      '',                             // I  Quantity Completed
      unitForCategory_(cat) || 'EA',  // J  Unit (derived from category)
      '',                             // K  Color/Material
      '',                             // L  Date Completed
      'Pending',                      // M  Status
      'Scanner',                      // N  Added By
      ''                              // O  Notes
    ]);
  });

  // ── Intersection Grid items ───────────────────────────────────
  const DIR_ORDER = ['n', 'e', 's', 'w', 'stop_msg', 'stop_lines'];
  grid.forEach((ig) => {
    if (!ig || !ig.intersection) return;
    DIR_ORDER.forEach((key) => {
      const raw = String(ig[key] || '').trim();
      if (!raw) return;

      let category, directions;
      if (key === 'stop_msg') {
        category   = 'Stop Msg';
        directions = expandDirLetters_(raw);
      } else if (key === 'stop_lines') {
        category   = 'Stop Line';
        directions = expandDirLetters_(raw);
      } else {
        category   = 'HVX Crosswalk';
        directions = [key.toUpperCase()];   // n→N, e→E, etc.
      }

      directions.forEach((dir) => {
        rows.push([
          `${woId}-${pad3(n++)}`,       // A  Item ID
          woId,                          // B  Work Order #
          workType,                      // C  Work Type
          'Intersection Grid',           // D  WO Section
          category,                      // E  Marking Type
          String(ig.intersection).trim(),// F  Intersection
          dir,                           // G  Direction
          '',                            // H  Description (blank for grid items)
          '',                            // I  Quantity Completed
          unitForCategory_(category) || 'EA',  // J  Unit (derived from category)
          '',                            // K  Color/Material
          '',                            // L  Date Completed
          'Pending',                     // M  Status
          'Scanner',                     // N  Added By
          ''                             // O  Notes
        ]);
      });
    });
  });

  // ── Bike-lane / pedestrian symbols ────────────────────────────
  // From the WO's General Remarks or the "Bike Lane Work (NEW)" row (per
  // the vision parse, d.bike_lane_markings). These have no structured
  // top-table cell, so they're seeded as Pending items the crew completes.
  // Quantity is left blank (crew enters the completed count); the WO's
  // stated count goes in the Description. Bike Symbols always map to the
  // OLD style — an admin switches the rare new ones by hand.
  const BIKE_TYPE_TO_CATEGORY = {
    'Bike Symbol':    'Old Bike Symbol (w/ rider)',
    'Bike Arrow':     'Bike Lane Arrow',
    'Pedestrian Men': 'Pedestrian Men',
  };
  bikeMarks.forEach((bm) => {
    if (!bm || !bm.type) return;
    const category = BIKE_TYPE_TO_CATEGORY[String(bm.type).trim()];
    if (!category) return;   // unknown type — skip
    const qty = bm.quantity;
    const hasQty = qty != null && String(qty).trim() !== '' && !isNaN(qty);
    rows.push([
      `${woId}-${pad3(n++)}`,                // A  Item ID
      woId,                                   // B  Work Order #
      workType,                               // C  Work Type
      'Top Table',                            // D  WO Section
      category,                               // E  Marking Type
      '',                                     // F  Intersection
      '',                                     // G  Direction
      hasQty ? `Per WO: ${qty}` : 'Per WO',   // H  Description (count noted here)
      '',                                     // I  Quantity Completed (blank — crew fills)
      unitForCategory_(category) || 'EA',     // J  Unit (derived from category → EA)
      '',                                     // K  Color/Material
      '',                                     // L  Date Completed
      'Pending',                              // M  Status
      'Scanner',                              // N  Added By
      ''                                      // O  Notes
    ]);
  });

  // ── Color Surface auto-seed (Paint / PT WOs) ──────────────────
  // A PT- WO is a Paint / Color Surface (MMA) job. Seed one Color Surface
  // item so the crew just confirms the sub-type and fills SQFT + color.
  // Defaults to 'Bike Lane' (an MMA category, unit SF) — the crew switches
  // to Bus Lane / Pedestrian Space if needed. Color/Material is left blank
  // and is required before the item can be marked Completed.
  if (isPaintWO) {
    rows.push([
      `${woId}-${pad3(n++)}`,                 // A  Item ID
      woId,                                    // B  Work Order #
      workType,                                // C  Work Type (= MMA for PT)
      'Top Table',                             // D  WO Section
      'Bike Lane',                             // E  Marking Type (Color Surface default)
      '',                                      // F  Intersection
      '',                                      // G  Direction
      'Color Surface — confirm type & color',  // H  Description
      '',                                      // I  Quantity Completed (blank — crew fills SQFT)
      unitForCategory_('Bike Lane') || 'SF',   // J  Unit (= SF)
      '',                                      // K  Color/Material (crew fills — required)
      '',                                      // L  Date Completed
      'Pending',                               // M  Status
      'Scanner',                               // N  Added By
      ''                                       // O  Notes
    ]);
  }

  if (rows.length === 0) return 0;

  // Write all rows at once via setValues (never appendRow — see memory
  // feedback_apps_script_appendrow_empty.md). Flush forces any deferred
  // validation error to fire here instead of at the next sheet read.
  const startRow = markingSheet.getLastRow() + 1;
  markingSheet
    .getRange(startRow, 1, rows.length, rows[0].length)
    .setValues(rows);
  SpreadsheetApp.flush();

  return rows.length;
}


/**
 * Apply completion updates to existing Marking Items rows. Writes cols
 * I (Quantity), J (Unit), K (Color/Material), L (Date Completed),
 * M (Status), O (Notes) per item. Zero/blank quantity → Status =
 * Pending (so a user can un-mark a row by clearing the number).
 *
 * @returns {number} rows actually touched (missing item_ids are skipped).
 */
function applyMarkingUpdates_(ss, updates, dateOfWork) {
  if (!updates || updates.length === 0) return 0;

  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) throw new Error('Marking Items sheet missing — run setupMarkingItems() first');

  const data = sheet.getDataRange().getValues();
  const idxById = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (id) idxById[id] = i + 1;  // 1-indexed sheet row
  }

  let updated = 0;
  updates.forEach(u => {
    const rowNum = idxById[u.item_id];
    if (!rowNum) {
      Logger.log(`⚠️ Marking Item ${u.item_id} not found — skipping update`);
      return;
    }
    const qty    = parseFloat(u.quantity);
    const hasQty = !isNaN(qty) && qty > 0;

    sheet.getRange(rowNum, 9).setValue(hasQty ? qty : '');                       // I Quantity
    if (u.unit) sheet.getRange(rowNum, 10).setValue(String(u.unit).trim());      // J Unit
    sheet.getRange(rowNum, 11).setValue(String(u.color_material || '').trim()); // K Color/Material
    sheet.getRange(rowNum, 12).setValue(hasQty ? dateOfWork : '');               // L Date Completed
    sheet.getRange(rowNum, 13).setValue(hasQty ? 'Completed' : 'Pending');       // M Status
    if (u.notes !== undefined) {
      sheet.getRange(rowNum, 15).setValue(String(u.notes || '').trim());         // O Notes
    }
    updated++;
  });

  SpreadsheetApp.flush();  // surface deferred validation errors in-place
  return updated;
}


/**
 * Append manually-added Marking Items rows (items the WO scan didn't
 * cover, or corrections the crew added in the field).
 *
 * @returns {number} rows appended.
 */
function applyMarkingNew_(ss, woId, workType, newItems, dateOfWork) {
  if (!newItems || newItems.length === 0) return 0;

  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) throw new Error('Marking Items sheet missing — run setupMarkingItems() first');

  // Find the highest item number already used for this WO so the new
  // IDs continue the sequence.
  const data = sheet.getDataRange().getValues();
  let maxN = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    if (id.indexOf(woId + '-') === 0) {
      const n = parseInt(id.split('-').pop(), 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
  }

  const pad3 = (x) => String(x).padStart(3, '0');
  const rows = newItems.map((item, idx) => {
    const qty    = parseFloat(item.quantity);
    const hasQty = !isNaN(qty) && qty > 0;
    return [
      `${woId}-${pad3(maxN + idx + 1)}`,           // A Item ID
      woId,                                         // B Work Order #
      workType,                                     // C Work Type
      'Manual',                                     // D WO Section
      String(item.category || '').trim(),           // E Marking Type
      String(item.intersection || '').trim(),       // F Intersection
      String(item.direction || '').trim(),          // G Direction
      String(item.description || '').trim(),        // H Description
      hasQty ? qty : '',                            // I Quantity
      String(item.unit || 'SF').trim(),             // J Unit
      String(item.color_material || '').trim(),     // K Color/Material
      hasQty ? dateOfWork : '',                     // L Date Completed
      hasQty ? 'Completed' : 'Pending',             // M Status
      'Manual',                                     // N Added By
      ''                                            // O Notes
    ];
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  return rows.length;
}


/**
 * Recompute WO Tracker cols 19-21 rollups from the Marking Items sheet.
 *
 * MMA:    marking_types     = distinct Category joined by ", "
 *         paint_material    = distinct Color/Material joined by ", "
 *         quantity_completed = SUM(Quantity WHERE Unit = 'SF')   ← MMA → SF
 *
 * Thermo: marking_types     = "N/A"   (too many categories to rollup)
 *         paint_material    = "N/A"   (Thermo doesn't record material)
 *         quantity_completed = SUM(Quantity WHERE Unit = 'LF')   ← Thermo → LF
 *
 * EA-unit items are always excluded (counts of messages, arrows, etc.
 * don't roll up into a square-foot or linear-foot total).
 *
 * Only COMPLETED items contribute to marking_types / paint_material.
 * Quantity Completed includes all items with a quantity (regardless of
 * status) so the tracker reflects what's been measured even if status
 * is still Pending.
 */
function computeMarkingRollups_(ss, woId, preloadedData) {
  const blank = { marking_types: '', quantity_completed: '', paint_material: '' };
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return blank;

  // Accept data preloaded by an earlier call in the same handler (e.g.
  // finalizeMarkingStatus_) so we don't re-scan the full sheet when
  // both run back-to-back.
  const data = preloadedData || sheet.getDataRange().getValues();
  if (data.length < 2) return blank;

  const woItems = data.slice(1).filter(r => String(r[1] || '').trim() === woId);
  if (woItems.length === 0) return blank;

  const anyThermo = woItems.some(r => String(r[2] || '').toLowerCase() === 'thermo');
  const targetUnit = anyThermo ? 'LF' : 'SF';

  // Column indices (0-based) under the 15-col schema:
  //   0 Item ID     4 Marking Type   8 Quantity      12 Status
  //   1 WO#         5 Intersection   9 Unit          13 Added By
  //   2 Work Type   6 Direction      10 Material     14 Notes
  //   3 WO Section  7 Description    11 Date Completed
  let qtySum = 0;
  let hasQty = false;
  woItems.forEach(r => {
    const unit = String(r[9] || '').toUpperCase();
    if (unit !== targetUnit) return;   // skips EA + the off-type unit
    const qty = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;
    qtySum += qty;
    hasQty = true;
  });

  if (anyThermo) {
    return {
      marking_types:      'N/A',
      quantity_completed: hasQty ? qtySum : '',
      paint_material:     'N/A',
    };
  }

  const cats = {}, mats = {};
  woItems.forEach(r => {
    const status = String(r[12] || '').toLowerCase();
    if (status !== 'completed') return;
    const cat = String(r[4]  || '').trim();
    const mat = String(r[10] || '').trim();
    if (cat) cats[cat] = true;
    if (mat && mat.toLowerCase() !== 'n/a') mats[mat] = true;
  });

  return {
    marking_types:      Object.keys(cats).sort().join(', '),
    quantity_completed: hasQty ? qtySum : '',
    paint_material:     Object.keys(mats).sort().join(', '),
  };
}


/**
 * One-shot migration — recompute Quantity Completed (col 20) for every
 * WO in the Tracker using the updated SF/LF branch logic. Run once from
 * the Apps Script editor after the SQFT → Quantity rename. Idempotent —
 * safe to re-run; the rollup is deterministic per WO.
 *
 * Also rewrites the col 20 header to "Quantity Completed" so the sheet
 * label matches the new value semantics.
 */
function migrateQuantityCompleted() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) throw new Error('Work Order Tracker sheet not found');

  // Step 1 — rename the col 20 header. Col 21 is the 1-indexed position.
  woSheet.getRange(1, 21).setValue('Quantity Completed');

  // Step 2 — preload Marking Items once so each WO's recompute is cheap.
  const miSheet = ss.getSheetByName('Marking Items');
  const miData  = miSheet ? miSheet.getDataRange().getValues() : null;

  // Step 3 — walk every WO row, recompute, write back to col 20.
  const data = woSheet.getDataRange().getValues();
  let processed = 0, updated = 0, skipped = 0;
  for (let i = 1; i < data.length; i++) {
    const woId = String(data[i][0] || '').trim();
    if (!woId) { skipped += 1; continue; }
    processed += 1;
    const rollups = computeMarkingRollups_(ss, woId, miData);
    const newVal  = rollups.quantity_completed;
    const oldVal  = data[i][20];
    if (String(oldVal) !== String(newVal)) {
      woSheet.getRange(i + 1, 21).setValue(newVal);  // col 21 (1-indexed) = col 20 (0-indexed)
      updated += 1;
    }
  }

  Logger.log(`migrateQuantityCompleted: processed=${processed}, updated=${updated}, skipped=${skipped}`);
  return { processed, updated, skipped };
}


/**
 * HTTP handler: return all Marking Items rows for a given WO, in sheet
 * row order (which is the insertion order — scan-seeded items first,
 * then manual items as they were added). The Field Report UI calls this
 * on WO selection to pre-populate its per-item SF input list.
 *
 * body: { action: 'get_marking_items', key, wo_id }
 * response: { items: [...] }
 */
function handleGetMarkingItems_(body) {
  // Express proxy wraps args under `data`; fall back to top-level for
  // direct callers.
  const payload = body.data || body;
  const woId = String(payload.wo_id || '').trim();
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return jsonResponse_({ items: [] });

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse_({ items: [] });

  // Column indices (0-based) under the 15-col schema:
  //   0 Item ID     4 Marking Type   8 Quantity      12 Status
  //   1 WO#         5 Intersection   9 Unit          13 Added By
  //   2 Work Type   6 Direction      10 Material     14 Notes
  //   3 WO Section  7 Description    11 Date Completed
  // .filter() preserves array order → sheet row order → insertion order.
  const items = data.slice(1)
    .filter(r => String(r[1] || '').trim() === woId)
    .map(r => ({
      item_id:        String(r[0]  || ''),
      work_order_id:  String(r[1]  || ''),
      work_type:      String(r[2]  || ''),
      section:        String(r[3]  || ''),
      category:       String(r[4]  || ''),
      intersection:   String(r[5]  || ''),
      direction:      String(r[6]  || ''),
      description:    String(r[7]  || ''),
      quantity:       r[8]  === '' || r[8]  == null ? null : Number(r[8]),
      unit:           String(r[9]  || ''),
      color_material: String(r[10] || ''),
      date_completed: r[11] instanceof Date
                         ? Utilities.formatDate(r[11], CONFIG.TIMEZONE, 'yyyy-MM-dd')
                         : String(r[11] || ''),
      status:         String(r[12] || ''),
      added_by:       String(r[13] || ''),
      notes:          String(r[14] || '')
    }));

  return jsonResponse_({ items });
}


// ═══════════════════════════════════════════════════════════════
// MARKING ITEMS — per-row CRUD (live Drive sync from the UI)
// ═══════════════════════════════════════════════════════════════

/**
 * Read a single row by Item ID and return the same shape as
 * handleGetMarkingItems_ entries (so the UI can spot-update state
 * without a refetch). Returns null if the ID is not found.
 */
function readMarkingItemById_(sheet, itemId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() !== itemId) continue;
    const r = data[i];
    return {
      item_id:        String(r[0]  || ''),
      work_order_id:  String(r[1]  || ''),
      work_type:      String(r[2]  || ''),
      section:        String(r[3]  || ''),
      category:       String(r[4]  || ''),
      intersection:   String(r[5]  || ''),
      direction:      String(r[6]  || ''),
      description:    String(r[7]  || ''),
      quantity:       r[8]  === '' || r[8]  == null ? null : Number(r[8]),
      unit:           String(r[9]  || ''),
      color_material: String(r[10] || ''),
      date_completed: r[11] instanceof Date
                         ? Utilities.formatDate(r[11], CONFIG.TIMEZONE, 'yyyy-MM-dd')
                         : String(r[11] || ''),
      status:         String(r[12] || ''),
      added_by:       String(r[13] || ''),
      notes:          String(r[14] || '')
    };
  }
  return null;
}


/**
 * Append ONE manually-added row. Status is always 'Pending' regardless
 * of whether the client supplied a quantity — crews only confirm
 * completion via the Field Report submit, not at item creation.
 *
 * body.data: { wo_id, category, description?, intersection?, direction?,
 *   unit?, quantity?, color_material?, notes?, work_type? }
 * response:  { item: <full row object> }
 */
function handleCreateMarkingItem_(body) {
  const d    = body.data || {};
  const woId = String(d.wo_id || '').trim();
  const cat  = String(d.category || '').trim();
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (!cat)  return jsonResponse_({ error: 'Missing category' }, 400);

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return jsonResponse_({
    error: 'Marking Items sheet missing — run setupMarkingItems() first'
  }, 500);

  // Derive work_type: prefer any existing row for this WO, then client hint.
  const allRows = sheet.getDataRange().getValues();
  let workType = '';
  let maxN     = 0;
  for (let i = 1; i < allRows.length; i++) {
    const id = String(allRows[i][0] || '');
    if (id.indexOf(woId + '-') !== 0) continue;
    if (!workType) workType = String(allRows[i][2] || '');
    const n = parseInt(id.split('-').pop(), 10);
    if (!isNaN(n) && n > maxN) maxN = n;
  }
  if (!workType) workType = String(d.work_type || '').trim();

  const qty = parseFloat(d.quantity);
  const hasQty = !isNaN(qty) && qty > 0;
  const pad3 = (x) => String(x).padStart(3, '0');
  const newId = `${woId}-${pad3(maxN + 1)}`;

  // Unit: when the category has a fixed unit, use the map; otherwise
  // (e.g. "Others") accept whatever the client sent, defaulting to 'EA'.
  const lockedUnit = unitForCategory_(cat);
  const finalUnit  = lockedUnit || String(d.unit || 'EA').trim();

  const row = [
    newId,                                      // A Item ID
    woId,                                       // B Work Order #
    workType,                                   // C Work Type
    'Manual',                                   // D WO Section
    cat,                                        // E Marking Type
    String(d.intersection || '').trim(),        // F Intersection
    String(d.direction    || '').trim(),        // G Direction
    String(d.description  || '').trim(),        // H Description
    hasQty ? qty : '',                          // I Quantity
    finalUnit,                                  // J Unit (derived from category)
    String(d.color_material || '').trim(),      // K Color/Material
    '',                                         // L Date Completed — always blank until submit
    'Pending',                                  // M Status — always Pending until submit
    'Manual',                                   // N Added By
    String(d.notes || '').trim()                // O Notes
  ];

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();

  const item = readMarkingItemById_(sheet, newId);
  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ item });
}


/**
 * Patch one or more editable fields on an existing row.
 *
 * Status / Date Completed rule: when `quantity` is in the patch and the
 * new value is 0/null/empty, force Status back to 'Pending' and clear
 * Date Completed (so a previously-Completed row can't look done with no
 * measurement). When `quantity` is set to > 0, leave Status and Date
 * Completed untouched — promotion to Completed is a submit-time job.
 *
 * body.data: { item_id, <any patchable field> }
 *   patchable: work_type, section, category, intersection, direction,
 *              description, quantity, unit, color_material, notes.
 * response:  { item: <full updated row> }
 */
function handleUpdateMarkingItem_(body) {
  const d      = body.data || {};
  const itemId = String(d.item_id || '').trim();
  if (!itemId) return jsonResponse_({ error: 'Missing item_id' }, 400);

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return jsonResponse_({
    error: 'Marking Items sheet missing — run setupMarkingItems() first'
  }, 500);

  const data = sheet.getDataRange().getValues();
  let rowNum = 0;
  let currentRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === itemId) {
      rowNum = i + 1;
      currentRow = data[i];
      break;
    }
  }
  if (!rowNum) return jsonResponse_({ error: 'item_id not found: ' + itemId }, 404);

  // Column indices. 1-based for setValue; 0-based (IDX) for currentRow reads.
  const COL = {
    work_type:     3,
    section:       4,
    category:      5,
    intersection:  6,
    direction:     7,
    description:   8,
    quantity:      9,
    unit:          10,
    color_material:11,
    date_completed:12,
    status:        13,
    notes:         15,
  };
  const IDX = {
    work_type: 2, section: 3, category: 4, intersection: 5, direction: 6,
    description: 7, quantity: 8, unit: 9, color_material: 10, status: 12,
    notes: 14,
  };

  const wasCompleted   = String(currentRow[IDX.status] || '') === 'Completed';
  let   anyFieldChanged = false;

  // If the patch changes the category, auto-derive the new unit from
  // the CATEGORY_UNITS_ map. If the new category is variable (e.g.
  // "Others"), honor the client-supplied unit or leave existing.
  // This runs BEFORE the string-field writes so the unit cell ends up
  // consistent with the (new) category even when the client didn't
  // explicitly patch unit.
  if (d.category !== undefined) {
    const newCat     = String(d.category || '').trim();
    const lockedUnit = unitForCategory_(newCat);
    if (lockedUnit) {
      // Force unit into the patch — overrides anything the client sent.
      d.unit = lockedUnit;
    }
  }

  // Write each patchable string field present in the request. Track
  // whether any value actually changed — used below to decide if a
  // previously-Completed row should reopen to Pending.
  ['work_type','section','category','intersection','direction','description',
   'unit','color_material','notes'].forEach(key => {
    if (d[key] === undefined) return;
    const newVal = String(d[key] || '').trim();
    const curVal = String(currentRow[IDX[key]] || '').trim();
    if (newVal !== curVal) anyFieldChanged = true;
    sheet.getRange(rowNum, COL[key]).setValue(newVal);
  });

  if (d.quantity !== undefined) {
    const q      = parseFloat(d.quantity);
    const hasQty = !isNaN(q) && q > 0;
    const cur    = currentRow[IDX.quantity];
    const curStr = (cur === '' || cur == null) ? '' : String(cur);
    const newStr = hasQty ? String(q) : '';
    if (curStr !== newStr) anyFieldChanged = true;
    sheet.getRange(rowNum, COL.quantity).setValue(hasQty ? q : '');

    // Clearing a qty flips a previously-Completed row back to Pending
    // immediately and clears Date Completed. Setting a positive qty
    // alone would otherwise leave Status as-is — but see the
    // "anyFieldChanged on Completed" rule below, which also covers
    // the set-qty-to-a-different-number case.
    if (!hasQty) {
      sheet.getRange(rowNum, COL.status).setValue('Pending');
      sheet.getRange(rowNum, COL.date_completed).setValue('');
    }
  }

  // Reopen a Completed row whenever the user actually changes something
  // — crews need to re-confirm at next submit. No-op if the row was
  // already Pending, or if the patch didn't actually change any values
  // (e.g. user opened Edit modal and hit Confirm without touching
  // anything).
  //
  // EXCEPTION: the admin "edit completed WO" flow on the Field Report
  // page passes `preserve_completion: true` so inline edits don't wipe
  // Date Completed (which would re-bucket the work to a different day
  // in Production Log + Dashboard rollups). Clearing qty to 0/blank
  // still flips to Pending above — that's an explicit "this didn't
  // get done" signal that overrides the preserve flag.
  if (wasCompleted && anyFieldChanged && !d.preserve_completion) {
    sheet.getRange(rowNum, COL.status).setValue('Pending');
    sheet.getRange(rowNum, COL.date_completed).setValue('');
  }

  SpreadsheetApp.flush();
  const item = readMarkingItemById_(sheet, itemId);
  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ item });
}


/**
 * Delete one or more Marking Items rows by Item ID.
 *
 * body.data: { item_ids: ['<id1>', '<id2>', ...] }
 * response:  { deleted: [<ids that were found and removed>] }
 */
function handleDeleteMarkingItems_(body) {
  const d   = body.data || {};
  const ids = Array.isArray(d.item_ids) ? d.item_ids.map(String) : [];
  if (ids.length === 0) return jsonResponse_({ error: 'item_ids is empty' }, 400);

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return jsonResponse_({
    error: 'Marking Items sheet missing — run setupMarkingItems() first'
  }, 500);

  const data = sheet.getDataRange().getValues();
  const idSet = {};
  ids.forEach(id => { idSet[String(id).trim()] = true; });

  // Collect 1-based row numbers to delete. Sort descending so deleting
  // doesn't shift earlier row indices out from under us.
  const rowNums = [];
  const deleted = [];
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim();
    if (idSet[rowId]) {
      rowNums.push(i + 1);
      deleted.push(rowId);
    }
  }
  rowNums.sort((a, b) => b - a);
  rowNums.forEach(r => sheet.deleteRow(r));
  SpreadsheetApp.flush();

  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ deleted });
}


/**
 * Submit-time status promotion. Walks every Marking Items row for `woId`:
 *
 *   qty > 0 AND Status != Completed  → promote to Completed + set Date = dateOfWork
 *   qty > 0 AND already Completed    → leave untouched (preserves original Date Completed)
 *   qty empty/0                      → force Pending, clear Date
 *
 * Preserving Date Completed across re-submissions is important: when a
 * WO spans multiple days and the crew submits a second Field Report on
 * Day 2, any rows already marked Completed on Day 1 must keep their
 * Day 1 Date Completed — otherwise the audit trail gets rewritten.
 */
function finalizeMarkingStatus_(ss, woId, dateOfWork, crewChief) {
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return { touched: 0, data: null };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { touched: 0, data };

  // Crew Chief lives at col 16 (1-idx) / 15 (0-idx) — see setupMultiCrewSchema.
  // When promoting Pending → Completed, tag the row with this submission's
  // chief so per-crew production aggregation (multi-shift handoff on the
  // same WO) can attribute each item to whoever finished it.
  const chief = String(crewChief || '').trim();

  let touched = 0;
  data.forEach((r, idx) => {
    if (idx === 0) return;  // header
    if (String(r[1] || '').trim() !== woId) return;

    const q              = parseFloat(r[8]);
    const hasQty         = !isNaN(q) && q > 0;
    const currentStatus  = String(r[12] || '');
    const rowNum         = idx + 1;

    if (hasQty) {
      // Promote only if not already Completed. If already Completed,
      // both Status and Date Completed are preserved — this is the
      // bug fix that keeps Day 1's date intact across Day 2 submits.
      if (currentStatus !== 'Completed') {
        sheet.getRange(rowNum, 13).setValue('Completed');
        sheet.getRange(rowNum, 12).setValue(dateOfWork);
        if (chief) sheet.getRange(rowNum, 16).setValue(chief);
        // Keep the in-memory copy in sync so downstream rollup readers
        // don't have to re-fetch the sheet.
        r[12] = 'Completed';
        r[11] = dateOfWork;
        if (chief) r[15] = chief;
        touched += chief ? 3 : 2;
      }
    } else {
      // qty empty/0 — row cannot be Completed. Revert to Pending and
      // clear Date if needed (defensive: covers admins zeroing a qty
      // directly on the sheet after a prior Completed submit).
      if (currentStatus !== 'Pending') {
        sheet.getRange(rowNum, 13).setValue('Pending');
        r[12] = 'Pending';
        touched++;
      }
      const currentDate = r[11] instanceof Date
        ? Utilities.formatDate(r[11], CONFIG.TIMEZONE, 'yyyy-MM-dd')
        : String(r[11] || '');
      if (currentDate !== '') {
        sheet.getRange(rowNum, 12).setValue('');
        r[11] = '';
        touched++;
      }
    }
  });
  if (touched) SpreadsheetApp.flush();
  return { touched, data };
}


// ═══════════════════════════════════════════════════════════════
// 9. WEB APP — Crew Field Report
// ═══════════════════════════════════════════════════════════════

/**
 * Serves the crew field report single-page web app.
 *
 * Deploy settings (Apps Script → Deploy → New Deployment → Web App):
 *   Execute as:  Me (script owner)
 *   Who has access: Anyone with the link
 *
 * The Apps Script URL is the app's only security layer.
 * The UPLOAD_SECRET key is injected at serve time so the browser
 * can authenticate its doPost calls without it appearing in source.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('FieldReport');
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.apiKey    = PropertiesService.getScriptProperties()
                         .getProperty('UPLOAD_SECRET') || '';
  return template.evaluate()
    .setTitle('Oneiro — Field Report')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
}


// ── action: get_active_wos ────────────────────────────────────

/**
 * Returns all non-complete Work Orders from the WO Tracker.
 * Used to populate the WO dropdown in the field report form.
 * Sorted: In Progress → Dispatched → Received.
 */
function handleGetActiveWOs_() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();

  // Completed + Returned WOs are now returned too so the Field Report
  // dropdown can open them (admin edit / re-generate CFR via the kebab
  // on the WO Tracker). They sort to the bottom of the dropdown.
  const ORDER = {
    'in progress': 0,
    'dispatched':  1,
    'received':    2,
    'completed':   3,
    'returned':    4,
  };

  const wos = allRows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id:                      String(r[0]),
      contractor:              String(r[1]),
      contract_number:         String(r[2]),
      borough:                 String(r[3]),
      location:                String(r[5]),
      from_street:             String(r[6]),
      to_street:               String(r[7]),
      due_date:                String(r[8]),
      priority:                String(r[9]),
      work_type:               String(r[10]),
      // cols 12/13 drive the Waterblasting Confirmed gate in the Field Report UI.
      // water_blast_required values: 'Yes - MMA' | 'No - Thermo' | '' — set at scan time.
      // water_blast_confirmed values: 'Yes' | 'No' | 'N/A' | '' — flipped by the toggle.
      water_blast_required:    String(r[12] || ''),
      water_blast_confirmed:   String(r[13] || ''),
      status:                  String(r[15]),
      dispatch_date:           String(r[16] || ''),
      work_start_date:         String(r[17] || ''),
      // Drive folder URL — same source the WO Tracker uses for its
      // 📁 link. Field Report page surfaces this as a "View WO in
      // Drive" link so crew leads can pull up the scanned WO + plan
      // docs without navigating away.
      folder_url:              String(r[42] || '').trim() || null,
    }))
    .sort((a, b) => {
      const aO = ORDER[a.status.toLowerCase()] ?? 99;
      const bO = ORDER[b.status.toLowerCase()] ?? 99;
      return aO - bO;
    });

  return jsonResponse_({ wos });
}


// ── action: update_wo_status ──────────────────────────────────
//
// Manual status change from the Dashboard kebab. Skips the one-way
// Received → Dispatched → In Progress → Completed state machine so
// admins can flip to any status (e.g. Returned for WOs handed back to
// the prime). Audit row appended to Automation Log so the change is
// traceable.
const WO_STATUS_VALUES_ = ['Received', 'Dispatched', 'In Progress', 'Completed', 'Returned'];

function handleUpdateWOStatus_(body) {
  const d = body.data || {};
  const woId   = String(d.wo_id || '').trim();
  const status = String(d.status || '').trim();
  if (!woId)   return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (!status) return jsonResponse_({ error: 'Missing status' }, 400);
  if (WO_STATUS_VALUES_.indexOf(status) === -1) {
    return jsonResponse_({ error: 'Invalid status: ' + status }, 400);
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Work Order Tracker');
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === woId) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return jsonResponse_({ error: 'WO not found: ' + woId }, 404);

  const previousStatus = String(data[rowIdx][15] || '');
  if (previousStatus === status) {
    return jsonResponse_({ success: true, noop: true });
  }
  sheet.getRange(rowIdx + 1, 16).setValue(status);

  const logSheet = ss.getSheetByName('Automation Log');
  if (logSheet) {
    logSheet.appendRow([
      new Date(), 'WO Tracker', 'Status changed manually',
      woId, `${previousStatus || '(blank)'} → ${status}`,
      'Completed', '', 'No',
    ]);
  }
  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ success: true, previous: previousStatus, status });
}


// ── action: delete_wo ─────────────────────────────────────────
//
// Hard-delete a WO from the Tracker + every Marking Items row keyed to
// it. Preserves the audit trail elsewhere (Work Day Log, Sign-In Data,
// Doc Lifecycle Log, Drive archive folder) so historical reporting
// stays intact. Logs a row to Automation Log with what was removed +
// preserved counts.
function handleDeleteWO_(body) {
  const d = body.data || {};
  const woId = String(d.wo_id || '').trim();
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData = woSheet.getDataRange().getValues();
  let woRowIdx = -1;
  for (let i = 1; i < woData.length; i++) {
    if (String(woData[i][0] || '').trim() === woId) { woRowIdx = i; break; }
  }
  if (woRowIdx === -1) return jsonResponse_({ error: 'WO not found: ' + woId }, 404);

  // Delete Marking Items rows bottom-up so indices stay valid.
  const miSheet = ss.getSheetByName('Marking Items');
  let miDeleted = 0;
  if (miSheet) {
    const miData = miSheet.getDataRange().getValues();
    for (let i = miData.length - 1; i >= 1; i--) {
      if (String(miData[i][1] || '').trim() === woId) {
        miSheet.deleteRow(i + 1);
        miDeleted++;
      }
    }
  }

  // Count Work Day Log rows so the audit entry tells admin what's left
  // behind (we don't delete them — payroll / time-card history).
  let wdlCount = 0;
  const wdlSheet = ss.getSheetByName('Work Day Log');
  if (wdlSheet) {
    const wdl = wdlSheet.getDataRange().getValues();
    for (let i = 1; i < wdl.length; i++) {
      if (String(wdl[i][1] || '').trim() === woId) wdlCount++;
    }
  }

  // Drop the Tracker row last so a partial failure earlier doesn't
  // orphan the WO row with no marking items.
  woSheet.deleteRow(woRowIdx + 1);

  const logSheet = ss.getSheetByName('Automation Log');
  if (logSheet) {
    logSheet.appendRow([
      new Date(), 'WO Tracker', 'WO deleted',
      woId,
      `Removed Tracker row + ${miDeleted} Marking Items row(s). ` +
      `${wdlCount} Work Day Log row(s) preserved. ` +
      `Drive archive folder preserved.`,
      'Completed', '', 'No',
    ]);
  }
  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({
    success: true,
    marking_items_deleted: miDeleted,
    work_day_log_preserved: wdlCount,
  });
}


// ── action: edit_completed_wo ─────────────────────────────────
//
// Admin-driven update to an already-Completed WO. Three CFR-handling
// modes:
//   data_only:    skip CFR JSON regen entirely (typo / silent fix)
//   replace_cfr:  regen JSON with original filename — Python merges
//                 new CFR into archived WO PDF; on approve the
//                 canonical WO_<woId>.pdf is replaced
//   new_cfr:      regen JSON with _Updated_<asOfDate> filename suffix
//                 — Python produces a separate filled PDF; archive
//                 saves alongside (canonical preserved)
//
// Marking items are assumed already-edited via the live PATCH endpoint
// (with preserve_completion=true so original Date Completed stays put).
// This handler just (re-)writes the rollups and queues the CFR JSON if
// requested. Optionally appends a Work Day Log row so the punch-order
// rework counts toward sign-in / production-log queues for a given date.
function handleEditCompletedWO_(body) {
  const d = body.data || {};
  const woId  = String(d.wo_id || '').trim();
  const mode  = String(d.regen_mode || 'data_only').trim();
  if (!woId)  return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (['data_only', 'replace_cfr', 'new_cfr'].indexOf(mode) === -1) {
    return jsonResponse_({ error: 'Invalid regen_mode: ' + mode }, 400);
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData = woSheet.getDataRange().getValues();
  let woRowIdx = -1;
  for (let i = 1; i < woData.length; i++) {
    if (String(woData[i][0] || '').trim() === woId) { woRowIdx = i; break; }
  }
  if (woRowIdx === -1) return jsonResponse_({ error: 'WO not found: ' + woId }, 404);
  const woRow = woData[woRowIdx];

  const asOfDate = String(d.as_of_date || '').trim();
  const includeInProduction = !!d.include_in_production;
  const productionDate = String(d.production_date || '').trim();

  // ── Batched marking-item changes ─────────────────────────────
  // Frontend buffers ALL marking edits while the admin is in edit mode
  // (no live PATCHes). On save mode click, the buffer is shipped here
  // as three lists. Apply in order: deletes → edits → adds. This keeps
  // index shifts predictable and ensures every newly-added row picks
  // up an ID after the deletes have happened.
  const editsIn   = Array.isArray(d.marking_edits)   ? d.marking_edits   : [];
  const addsIn    = Array.isArray(d.marking_adds)    ? d.marking_adds    : [];
  const deletesIn = Array.isArray(d.marking_deletes) ? d.marking_deletes : [];

  const miSheet = ss.getSheetByName('Marking Items');
  if (!miSheet) return jsonResponse_({ error: 'Marking Items sheet missing' }, 500);

  // Column indices (1-indexed for setValue):
  //   E5 Category, F6 Intersection, G7 Direction, H8 Description,
  //   I9 Quantity, J10 Unit, K11 Color/Material, L12 Date Completed,
  //   M13 Status, O15 Notes.
  const MI_COL = {
    work_type: 3, section: 4, category: 5, intersection: 6, direction: 7,
    description: 8, quantity: 9, unit: 10, color_material: 11,
    date_completed: 12, status: 13, notes: 15,
  };

  // 1) DELETES — bottom-up so row indices stay stable. Defensive: skip
  //    rows that don't belong to this WO.
  const deletedIds = [];
  if (deletesIn.length > 0) {
    const deleteSet = {};
    deletesIn.forEach(id => { deleteSet[String(id).trim()] = true; });
    const miData = miSheet.getDataRange().getValues();
    for (let i = miData.length - 1; i >= 1; i--) {
      const itemId = String(miData[i][0] || '').trim();
      if (!deleteSet[itemId]) continue;
      if (String(miData[i][1] || '').trim() !== woId) continue;  // defensive
      miSheet.deleteRow(i + 1);
      deletedIds.push(itemId);
    }
    if (deletedIds.length) SpreadsheetApp.flush();
  }

  // 2) EDITS — per-field writes. Skipping the reopen-on-change rule
  //    `handleUpdateMarkingItem_` applies in normal mode: in the
  //    edit-completed context Date Completed semantics come from the
  //    production toggle below, not from this loop.
  const editedIds = [];
  if (editsIn.length > 0) {
    const miData = miSheet.getDataRange().getValues();
    const rowByItemId = {};
    for (let i = 1; i < miData.length; i++) {
      const id = String(miData[i][0] || '').trim();
      if (id && String(miData[i][1] || '').trim() === woId) {
        rowByItemId[id] = i + 1;
      }
    }
    editsIn.forEach(patch => {
      const itemId = String(patch && patch.item_id || '').trim();
      if (!itemId) return;
      const sheetRow = rowByItemId[itemId];
      if (!sheetRow) return;
      // Write each changed column. Unit/category coupling (unitForCategory_)
      // is enforced client-side via MarkingFormModal already.
      ['work_type','section','category','intersection','direction',
       'description','unit','color_material','notes'].forEach(key => {
        if (patch[key] === undefined) return;
        miSheet.getRange(sheetRow, MI_COL[key]).setValue(String(patch[key] || '').trim());
      });
      if (patch.quantity !== undefined) {
        const q = parseFloat(patch.quantity);
        const hasQty = !isNaN(q) && q > 0;
        miSheet.getRange(sheetRow, MI_COL.quantity).setValue(hasQty ? q : '');
        // Zeroing qty out IS an explicit "didn't get done" — fall back
        // to the original revert-to-Pending behavior so rollups stay
        // honest. preserve_completion has no effect on this branch.
        if (!hasQty) {
          miSheet.getRange(sheetRow, MI_COL.status).setValue('Pending');
          miSheet.getRange(sheetRow, MI_COL.date_completed).setValue('');
        }
      }
      editedIds.push(itemId);
    });
    if (editedIds.length) SpreadsheetApp.flush();
  }

  // 3) ADDS — append new rows. Status=Pending, Date Completed=blank;
  //    `finalizeMarkingStatus_` below promotes them to Completed using
  //    finalizeDate when Quantity > 0. New IDs continue the WO's
  //    existing item-number sequence.
  const addedIds = [];
  if (addsIn.length > 0) {
    const miData = miSheet.getDataRange().getValues();
    let maxN = 0;
    let woWorkTypeFromExisting = '';
    for (let i = 1; i < miData.length; i++) {
      const id = String(miData[i][0] || '');
      if (id.indexOf(woId + '-') !== 0) continue;
      if (!woWorkTypeFromExisting) woWorkTypeFromExisting = String(miData[i][2] || '');
      const n = parseInt(id.split('-').pop(), 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
    const pad3 = (x) => String(x).padStart(3, '0');
    const rows = addsIn.map((item, idx) => {
      const qty    = parseFloat(item.quantity);
      const hasQty = !isNaN(qty) && qty > 0;
      const cat    = String(item.category || '').trim();
      const lockedUnit = unitForCategory_(cat);
      const finalUnit  = lockedUnit || String(item.unit || 'EA').trim();
      const newId = `${woId}-${pad3(maxN + idx + 1)}`;
      addedIds.push(newId);
      return [
        newId,                                            // A Item ID
        woId,                                             // B Work Order #
        String(item.work_type || woWorkTypeFromExisting || ''),  // C Work Type
        String(item.section || 'Manual'),                 // D WO Section
        cat,                                              // E Marking Type
        String(item.intersection || '').trim(),           // F Intersection
        String(item.direction || '').trim(),              // G Direction
        String(item.description || '').trim(),            // H Description
        hasQty ? qty : '',                                // I Quantity
        finalUnit,                                        // J Unit
        String(item.color_material || '').trim(),         // K Color/Material
        '',                                               // L Date Completed (finalize fills in)
        'Pending',                                        // M Status (finalize promotes if hasQty)
        'Manual',                                         // N Added By
        String(item.notes || '').trim(),                  // O Notes
      ];
    });
    if (rows.length > 0) {
      const startRow = miSheet.getLastRow() + 1;
      miSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      SpreadsheetApp.flush();
    }
  }

  // Derive touchedIds for the production-date override below — edited
  // existing items + newly-added items (the ones whose date should
  // bucket into production_date if checked).
  const touchedIds = editedIds.concat(addedIds);

  // finalizeDate drives Date Completed for items that get promoted
  // Pending → Completed during this update (newly-added items above
  // start Pending), the prefix on appended Issues lines, and the
  // touched-already-Completed override below.
  //
  //   include_in_production checked → production_date (rework day)
  //   else                          → WO's work_end / work_start
  //                                   (original timeframe)
  //
  // Falls back to as_of_date / today if all WO date columns are blank.
  const woWorkEnd   = _normalizeTrackerDate_(woRow[18]);
  const woWorkStart = _normalizeTrackerDate_(woRow[17]);
  const finalizeDate = (includeInProduction && productionDate)
    ? productionDate
    : (woWorkEnd || woWorkStart || asOfDate);

  // ── Recompute marking rollups (cols 19-21) from current sheet state.
  // No state-machine bump — completed WOs stay Completed regardless of
  // what the marking rows now look like. Status changes go through the
  // separate update_wo_status action.
  const { data: markingData } = finalizeMarkingStatus_(ss, woId, finalizeDate);

  // If include_in_production is checked, the admin is logging this work
  // as production for production_date. Overwrite Date Completed on
  // every touched, already-Completed item so they bucket into that day
  // — finalize only handles the Pending → Completed transition; items
  // already Completed pre-edit kept their original date.
  // Items NOT in touchedIds (= un-touched existing items) are left
  // alone — admin didn't edit them, no reason to re-bucket.
  if (includeInProduction && productionDate && touchedIds.length > 0) {
    const miData2 = markingData || miSheet.getDataRange().getValues();
    const touchedSet = {};
    touchedIds.forEach(id => { touchedSet[id] = true; });
    for (let i = 1; i < miData2.length; i++) {
      const r = miData2[i];
      if (String(r[1] || '').trim() !== woId) continue;
      if (!touchedSet[String(r[0] || '').trim()]) continue;
      if (String(r[12] || '') !== 'Completed') continue;
      miSheet.getRange(i + 1, 12).setValue(productionDate);
      r[11] = productionDate;
    }
    SpreadsheetApp.flush();
  }
  const rollups = computeMarkingRollups_(ss, woId, markingData);

  // Append new issues with a date prefix consistent with finalizeDate
  // so reading the Issues log later tells you when each note's work
  // actually happened.
  const newIssuesText = String(d.issues || '').trim();
  const currentIssues = String(woRow[22] || '').trim();
  let issuesOut = currentIssues;
  if (newIssuesText) {
    const prefix = finalizeDate ? finalizeDate + ': ' : '';
    issuesOut = currentIssues
      ? currentIssues + '\n' + prefix + newIssuesText
      : prefix + newIssuesText;
  }

  const woValues = [
    String(woRow[15] || 'Completed'),   // 15 Status (unchanged)
    _normalizeTrackerDate_(woRow[16]),  // 16 Dispatch (unchanged)
    _normalizeTrackerDate_(woRow[17]),  // 17 Work Start (unchanged)
    _normalizeTrackerDate_(woRow[18]),  // 18 Work End (unchanged)
    rollups.marking_types,              // 19
    rollups.quantity_completed,         // 20
    rollups.paint_material,             // 21
    issuesOut,                          // 22
    String(woRow[23] || ''),            // 23 Photos? (unchanged)
  ];
  writeRowWithProbing_(woSheet, woRowIdx + 1, 16, woValues,
    ['Status', 'Dispatch Date', 'Work Start Date', 'Work End Date',
     'Marking Types', 'Quantity Completed', 'Paint/Material',
     'Issues Reported', 'Photos Uploaded?'],
    'WO Tracker (edit completed)'
  );

  // Optionally log a new Work Day Log row so the work counts in
  // sign-in / production-log queues for the given production_date.
  let wdlAppended = false;
  if (includeInProduction && productionDate) {
    const prodDate = productionDate;
    const wdlSheet = _getOrCreateWorkDayLogSheet_(ss);
    appendRowWithProbing_(
      wdlSheet,
      [
        prodDate,                                                       // 0 Date
        woId,                                                           // 1 WO#
        String(woRow[1] || ''),                                         // 2 Contractor
        String(woRow[2] || ''),                                         // 3 Contract #
        String(woRow[3] || ''),                                         // 4 Borough
        String(woRow[5] || ''),                                         // 5 Location
        new Date(),                                                     // 6 Submitted At
        'Pending',                                                      // 7 Status
      ],
      ['Date', 'Work Order #', 'Prime Contractor', 'Contract #',
       'Borough', 'Location', 'Submitted At', 'Status'],
      'Work Day Log (edit completed)'
    );
    wdlAppended = true;
  }

  // CFR JSON queue. data_only → no work here.
  let cfrFilename = null;
  if (mode === 'replace_cfr' || mode === 'new_cfr') {
    const originalDate = _normalizeTrackerDate_(woRow[18]) || _normalizeTrackerDate_(woRow[17]) || asOfDate;
    const filenameSuffix = mode === 'new_cfr' && asOfDate ? `_Updated_${asOfDate}` : '';
    cfrFilename = `CFR_${woId}_${originalDate}${filenameSuffix}.json`;
    // Build the `d` shape generateContractorFieldReportJson_ expects.
    // For replace mode: install_to stays the original work-end date so
    // the regenerated CFR mirrors the original timeline. For new_cfr:
    // install_to advances to the punch-rework date.
    const synthD = {
      wo_id: woId,
      date: (mode === 'new_cfr' && asOfDate) ? asOfDate : originalDate,
      issues: newIssuesText || '',
      photos_uploaded: String(woRow[23] || '').toLowerCase() === 'yes',
    };
    generateContractorFieldReportJson_(synthD, woRow, ss, '', { filename: cfrFilename });
  }

  // Audit
  const logSheet = ss.getSheetByName('Automation Log');
  if (logSheet) {
    const note =
      `mode=${mode}` +
      `, marking_changes=edits:${editedIds.length}/adds:${addedIds.length}/deletes:${deletedIds.length}` +
      (wdlAppended ? `, WDL row appended for ${productionDate}` : '') +
      (cfrFilename ? `, CFR JSON queued: ${cfrFilename}` : '');
    logSheet.appendRow([
      new Date(), 'WO Tracker', 'Completed WO edited',
      woId, note, 'Completed', '', 'No',
    ]);
  }

  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({
    success: true,
    mode,
    wdl_appended: wdlAppended,
    cfr_filename: cfrFilename,
    marking_changes: {
      edited:  editedIds.length,
      added:   addedIds.length,
      deleted: deletedIds.length,
    },
  });
}


// ── action: update_wo_coordinates ─────────────────────────────
//
// Manual lat/lng entry from the Nav tab — used to fix flagged
// geocode warnings and to give admins an escape hatch when Google's
// geocoder picked the wrong intersection. Validates the values,
// writes them to the Tracker, clears any existing warning, logs the
// change. Idempotent.
function handleUpdateWOCoordinates_(body) {
  const d = body.data || {};
  const woId = String(d.wo_id || '').trim();
  const lat  = parseFloat(d.lat);
  const lng  = parseFloat(d.lng);
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (isNaN(lat) || lat < -90  || lat > 90)  return jsonResponse_({ error: 'Invalid lat (must be -90..90)' }, 400);
  if (isNaN(lng) || lng < -180 || lng > 180) return jsonResponse_({ error: 'Invalid lng (must be -180..180)' }, 400);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  ensureWoTrackerExtraCols_(woSheet);
  const data = woSheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === woId) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return jsonResponse_({ error: 'WO not found: ' + woId }, 404);

  const sheetRow = rowIdx + 1;
  woSheet.getRange(sheetRow, 46).setValue(lat);          // Latitude
  woSheet.getRange(sheetRow, 47).setValue(lng);          // Longitude
  woSheet.getRange(sheetRow, 48).setValue('');           // clear warning
  woSheet.getRange(sheetRow, 49).setValue(new Date());   // Geocoded At

  const logSheet = ss.getSheetByName('Automation Log');
  if (logSheet) {
    logSheet.appendRow([
      new Date(), 'Geocoder', 'Coords set manually',
      woId, `${lat.toFixed(5)}, ${lng.toFixed(5)}`,
      'Completed', '', 'No',
    ]);
  }

  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ success: true, lat, lng });
}


// ── action: get_wo_map_data ───────────────────────────────────
//
// Returns active WOs (Received / Dispatched / In Progress) for the
// Nav tab map. Two arrays: `mapped` (lat/lng populated) and
// `unmapped` (coords missing — surfaced in the "Needs coords" panel
// so admin can set them manually).
function handleGetWOMapData_(_body) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData  = woSheet.getDataRange().getValues();

  // Marking-items count per WO — same shape as _buildDashboardPayload_
  // computes. Total + completed so the popover can show progress.
  const miCounts = {};
  const miSheet = ss.getSheetByName('Marking Items');
  if (miSheet) {
    const miData = miSheet.getDataRange().getValues();
    for (let i = 1; i < miData.length; i++) {
      const woId = String(miData[i][1] || '').trim();
      if (!woId) continue;
      const bucket = miCounts[woId] || (miCounts[woId] = { total: 0, completed: 0 });
      bucket.total += 1;
      if (String(miData[i][12] || '').toLowerCase() === 'completed') bucket.completed += 1;
    }
  }

  const ACTIVE = { 'received': 1, 'dispatched': 1, 'in progress': 1 };

  const mapped = [];
  const unmapped = [];
  for (let i = 1; i < woData.length; i++) {
    const r = woData[i];
    const woId = String(r[0] || '').trim();
    if (!woId) continue;
    const status = String(r[15] || '');
    if (!ACTIVE[status.toLowerCase()]) continue;

    const lat = r[45];
    const lng = r[46];
    const counts = miCounts[woId] || { total: 0, completed: 0 };

    const base = {
      wo_id:                  woId,
      contractor:             String(r[1]  || ''),
      contract_num:           String(r[2]  || ''),
      borough:                String(r[3]  || ''),
      location:               String(r[5]  || ''),
      from_street:            String(r[6]  || ''),
      to_street:              String(r[7]  || ''),
      due_date:               r[8] instanceof Date
                                 ? Utilities.formatDate(r[8], CONFIG.TIMEZONE, 'yyyy-MM-dd')
                                 : String(r[8] || ''),
      priority:               String(r[9]  || ''),
      status:                 status,
      folder_url:             String(r[42] || '').trim() || null,
      geocode_warning:        String(r[47] || ''),
      marking_item_count:     counts.total,
      marking_completed_count: counts.completed,
    };

    if (typeof lat === 'number' && typeof lng === 'number' && !isNaN(lat) && !isNaN(lng)) {
      mapped.push({ ...base, lat, lng });
    } else {
      unmapped.push({ ...base, lat: null, lng: null });
    }
  }

  return jsonResponse_({ mapped, unmapped });
}


// ── action: submit_field_report ───────────────────────────────

/**
 * Writes a crew field report:
 *   1. Appends one row per crew member to Daily Sign-In Data
 *   2. Updates WO Tracker operational columns (Status, Dispatch, Start/End
 *      dates, Marking Types, SQFT, Paint/Material, Issues, Photos)
 *
 * Expected body.data fields:
 *   wo_id           — Work Order # (e.g. "PT-11930")
 *   date            — Date of work (YYYY-MM-DD). Used to auto-set Dispatch Date,
 *                     Work Start Date (if blank), and Work End Date (if complete).
 *   wo_complete     — boolean — marks WO complete, sets Work End Date = date
 *   marking_types       — string ("Crosswalk: 500 SF, Stop Bar: 10 LF")
 *   quantity_completed  — number (SF for MMA, LF for Thermo)
 *   paint_material      — string
 *   issues          — string (appended to existing issues with date prefix)
 *   photos_uploaded — boolean
 *   crew            — [{name, classification, time_in, time_out, hours}]
 *                     (currently ignored by the server — sign-in lives in
 *                     its own tab and writes Daily Sign-In Data directly)
 *
 * Daily Sign-In Data columns (0-indexed, 18 total):
 *   0  Date                8  Time In           16  Admin Reviewed?
 *   1  Work Order #        9  Time Out           17  Review Notes
 *   2  Prime Contractor   10  Hours Worked
 *   3  Contract #         11  Overtime Hours
 *   4  Borough            12  SQFT Completed
 *   5  Location           13  Paint/Material
 *   6  Employee Name      14  WO Complete?
 *   7  Classification     15  Issues/Notes
 *
 * WO Tracker columns updated (0-indexed):
 *   15  Status             19  Marking Types      23  Photos Uploaded?
 *   16  Dispatch Date      20  Quantity Completed
 *   17  Work Start Date    21  Paint/Material Used
 *   18  Work End Date      22  Issues Reported
 */

// Normalize whatever's in a Tracker date cell down to a YYYY-MM-DD
// string. Handles:
//   - Date objects (Sheets stores dates this way when a cell is typed
//     to "Date" — typically because someone manually typed e.g. "5/5"
//     and Sheets auto-coerced it).
//   - "Tue May 05 2026 00:00:00 GMT-…" long-strings (the artifact of
//     a previous submit that did `String(dateObject)` and persisted
//     the result). The dayname/monthname/year regex anchors so we only
//     touch strings that smell like Date.toString() output — anything
//     else passes through.
//   - Already-canonical YYYY-MM-DD strings — passed through verbatim.
//   - Empty / falsy — returns ''.
function _normalizeTrackerDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^(Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+\w{3}\s+\d{1,2}\s+\d{4}\b/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
  }
  return s;
}

/**
 * Pre-submit gate for the Field Report form: warn the user when a
 * Crew Chief appears to still be wrapping up last night's overnight
 * shift but the form's date is set to today.
 *
 * The overnight-shift cutoff (CONFIG.OPERATIONAL_DAY_CUTOFF_HOUR = 5)
 * auto-buckets pre-5am submits to yesterday's operational day. But a
 * chief who finishes a night shift and submits at 7am sees the form
 * default to today's calendar date — silently mis-attributing the
 * shift to a day that the crew didn't actually work.
 *
 * This handler detects the at-risk window:
 *   - work_date == calendar today (NYC)
 *   - submit time is between cutoff (5am) and cutoff + 4hr (9am)
 *   - chief has at least one WDL row dated yesterday whose FR Submitted
 *     At timestamp is later than yesterday 11pm (i.e. the row was
 *     filed in the late-evening / overnight window characteristic of
 *     an overnight crew, NOT a day-shift end-of-day submit).
 *
 * Returns { should_confirm, prior_date, reason }. The webapp shows a
 * soft-warn modal asking the user to pick last-night vs today; the
 * decision is the user's, this handler only flags.
 *
 * body.data = { crew_chief, work_date: 'YYYY-MM-DD' }
 */
function handleCheckFrShiftAttribution_(body) {
  const d         = body.data || {};
  const crewChief = String(d.crew_chief || '').trim();
  const workDate  = String(d.work_date  || '').trim();
  const noop = { should_confirm: false, prior_date: null, reason: '' };

  if (!crewChief || !workDate) return jsonResponse_(noop);

  const now = new Date();
  const calendarToday = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  if (workDate !== calendarToday) return jsonResponse_(noop);

  // Window: from cutoff (5am NYC) to cutoff + 4hr (9am NYC). Before
  // cutoff, opDay correction already buckets the submit to yesterday
  // — no warning needed. After 9am, the chief is almost certainly
  // starting a new day; warning would be noise.
  const cutoffHour = CONFIG.OPERATIONAL_DAY_CUTOFF_HOUR || 5;
  const nyHourStr  = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH');
  const nyHour     = parseInt(nyHourStr, 10);
  if (isNaN(nyHour) || nyHour < cutoffHour || nyHour >= (cutoffHour + 4)) {
    return jsonResponse_(noop);
  }

  // Yesterday in NYC calendar terms.
  const todayParts = calendarToday.split('-').map(Number);
  const yDate = new Date(todayParts[0], todayParts[1] - 1, todayParts[2] - 1);
  const yesterdayIso = Utilities.formatDate(yDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // Threshold: yesterday 23:00 NYC (the late-evening / overnight
  // signal). Anything submitted before this on yesterday is most
  // likely a day-shift end-of-shift filing, not a night shift.
  const threshold = new Date(yDate.getFullYear(), yDate.getMonth(), yDate.getDate(), 23, 0, 0);

  const wdlSheet = ss => ss.getSheetByName('Work Day Log');
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = wdlSheet(ss);
  if (!sheet) return jsonResponse_(noop);

  const data = sheet.getDataRange().getValues();
  // WDL post-multi-crew schema:
  //   0 Date, 2 Contractor, 3 Contract#, 4 Borough,
  //   6 FR Submitted At, 7 Crew Chief, 8 Sign-In Status.
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };
  const chiefNorm = crewChief.toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (ymd(r[0]) !== yesterdayIso) continue;
    const rowChief = String(r[7] || '').trim().toLowerCase();
    if (rowChief !== chiefNorm) continue;
    const submittedAt = r[6];
    if (!(submittedAt instanceof Date)) continue;
    if (submittedAt.getTime() >= threshold.getTime()) {
      return jsonResponse_({
        should_confirm: true,
        prior_date:     yesterdayIso,
        reason:         `${crewChief} filed FRs dated ${yesterdayIso} late evening / overnight — likely an overnight shift still being wrapped up.`,
      });
    }
  }

  return jsonResponse_(noop);
}


function handleSubmitFieldReport_(body) {
  const d = body.data || {};

  if (!d.wo_id) return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (!d.date)  return jsonResponse_({ error: 'Missing date' }, 400);
  // Crew roster lives on the Sign-In tab. The Field Report captures the
  // Crew Chief name as a per-shift crew identifier — threaded through
  // WDL, Marking Items completion, DSID, and the Sign-In queue grouping
  // so multiple crews from the same contractor on the same source job
  // produce separate queue cards / PLs. Required to support multi-crew
  // dispatch and multi-shift handoffs on the same WO.
  const crewChief = String(d.crew_chief || '').trim();
  if (!crewChief) {
    return jsonResponse_({ error: 'Missing crew_chief — pick the Crew Chief before submitting.' }, 400);
  }

  // Operational-day correction. If the client sent calendar-today (the
  // default) but submission is happening before the cutoff hour (e.g.
  // crew finishing at 3:30 AM), bucket to yesterday's operational day
  // instead — that's the day the shift actually started. We DON'T
  // touch d.date when the client sent a date other than calendar-today
  // (that's the user explicitly overriding from the picker).
  const _calendarToday = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  if (d.date === _calendarToday) {
    const opTodayIso = opToday_();
    if (opTodayIso && opTodayIso !== _calendarToday) {
      d.date = opTodayIso;
    }
  }

  // Breadcrumb: updated before each risky operation so any uncaught exception
  // inside this handler surfaces with "[step=<phase>]" attached — otherwise
  // doPost's catch returns only the raw exception text (e.g. "Invalid Entry")
  // and there's no way to tell which phase blew up.
  let step = 'init';

  try {

  step = 'open spreadsheet / find WO row';
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();

  // Find WO row — allRows[0] is header (sheet row 1), data starts at allRows[1]
  const woRowIdx = allRows.findIndex((r, i) => i > 0 && String(r[0]) === String(d.wo_id));
  if (woRowIdx === -1) {
    return jsonResponse_({ error: 'Work Order not found: ' + d.wo_id }, 404);
  }
  const woRow = allRows[woRowIdx];

  // ── MMA waterblasting gate ──────────────────────────────────
  // col 12 "Water Blast Required?" → "Yes - MMA" means this is MMA work and
  // col 13 "Water Blast Confirmed?" must be "Yes" before we accept a report.
  // The UI greys out + disables Submit for unconfirmed MMA WOs, but we also
  // guard server-side so a hand-crafted POST can't slip past.
  step = 'mma waterblasting gate';
  const wbRequired  = String(woRow[12] || '');
  const wbConfirmed = String(woRow[13] || '');
  if (wbRequired === 'Yes - MMA' && wbConfirmed !== 'Yes') {
    return jsonResponse_({
      error: 'Waterblasting not confirmed for this MMA work order — toggle ' +
             '"Waterblasting Confirmed" at the top of the Field Report before submitting.'
    }, 400);
  }

  // ── Derive updated WO Tracker values ─────────────────────────

  // Status: Received → Dispatched → In Progress → Complete
  const currentStatus    = String(woRow[15] || 'Received');
  // Read existing date columns through a normalizer so a stale "Date
  // object" cell (typical when someone manually typed a date in Sheets
  // and Sheets auto-coerced it to a date type) doesn't get written
  // back as "Tue May 05 2026 00:00:00 GMT-…". The normalizer accepts
  // either a Date object or a long-format String, and emits YYYY-MM-DD.
  // Self-healing: any row touched by a future Field Report submit
  // gets its dispatch/work_start/work_end repaired in place.
  const currentDispatch  = _normalizeTrackerDate_(woRow[16]);
  const currentWorkStart = _normalizeTrackerDate_(woRow[17]);
  const currentWorkEnd   = _normalizeTrackerDate_(woRow[18]);
  const currentIssues    = String(woRow[22] || '').trim();

  // Auto-derive dates from Date of Work when not already recorded.
  // Dispatch Date and Work Start Date are treated as the same — both default
  // to the first Field Report's Date of Work if blank in the tracker.
  const newDispatch  = currentDispatch  || d.date;
  const newWorkStart = currentWorkStart || d.date;
  // Work End Date is only written when the WO is marked complete on this submission.
  const newWorkEnd   = d.wo_complete ? d.date : currentWorkEnd;

  // Progress status forward (never backward)
  let newStatus = currentStatus;
  if (d.wo_complete) {
    newStatus = 'Completed';
  } else if (newWorkStart && (currentStatus === 'Received' || currentStatus === 'Dispatched')) {
    newStatus = 'In Progress';
  } else if (newDispatch && currentStatus === 'Received') {
    newStatus = 'Dispatched';
  }

  // Append new issues with date prefix; preserve existing
  let newIssues = currentIssues;
  if (d.issues && d.issues.trim()) {
    const issueLine = d.date + ': ' + d.issues.trim();
    newIssues = currentIssues ? currentIssues + '\n' + issueLine : issueLine;
  }

  // Photos: once "Yes", stays "Yes"
  const newPhotos = (d.photos_uploaded || String(woRow[23]).toLowerCase() === 'yes')
    ? 'Yes' : 'No';

  // ── Finalize Marking Items + compute rollups ─────────────────
  // Marking Items rows are already live-persisted via the per-row CRUD
  // endpoints (create/update/delete) — no need to re-apply anything here.
  // Submit time does two things:
  //  (1) Promote Status: any row with qty > 0 → Completed, else Pending.
  //      (Status was held at Pending during data entry; this is when the
  //      crew formally confirms the day's work.)
  //  (2) Recompute WO Tracker cols 19-21 rollups from final sheet state.
  step = 'marking finalize status';
  const { touched: nFinalized, data: markingData } = finalizeMarkingStatus_(ss, d.wo_id, d.date, crewChief);

  step = 'marking rollup';
  // Reuse the values finalizeMarkingStatus_ already read (with its
  // in-memory updates applied). Saves a full-sheet getDataRange call.
  const rollups = computeMarkingRollups_(ss, d.wo_id, markingData);
  Logger.log(`✅ Marking Items: ${nFinalized} status cells updated; ` +
             `rollup → types=${JSON.stringify(rollups.marking_types)}, ` +
             `quantity=${rollups.quantity_completed}, ` +
             `material=${JSON.stringify(rollups.paint_material)}`);

  // ── Write WO Tracker cols 15–23 (0-indexed) ──────────────────
  // Col 15 (0-indexed) = col 16 (1-indexed); 9 columns → cols 16–24 (1-indexed)
  const woValues = [
    newStatus,                          // col 15: Status
    newDispatch,                        // col 16: Dispatch Date
    newWorkStart,                       // col 17: Work Start Date
    newWorkEnd,                         // col 18: Work End Date
    rollups.marking_types,              // col 19: Marking Types (rollup)
    rollups.quantity_completed,         // col 20: Quantity Completed (rollup)
    rollups.paint_material,             // col 21: Paint/Material Used (rollup)
    newIssues,                          // col 22: Issues Reported
    newPhotos                           // col 23: Photos Uploaded?
  ];
  const woLabels = [
    'Status', 'Dispatch Date', 'Work Start Date', 'Work End Date',
    'Marking Types', 'Quantity Completed', 'Paint/Material',
    'Issues Reported', 'Photos Uploaded?'
  ];
  step = 'WO Tracker write';
  writeRowWithProbing_(woSheet, woRowIdx + 1, 16, woValues, woLabels, 'WO Tracker');

  Logger.log('✅ WO Tracker updated: ' + d.wo_id + ' → ' + newStatus);

  // ── Write Work Day Log row ─────────────────────────────────────
  // Each Field Report submission creates exactly one row here. The
  // Sign-In tab queue reads this sheet (filtered to Status='Pending')
  // and groups rows by (date, contract # + borough) so the user can
  // file a single sign-in covering every WO worked that night for a
  // given contract.
  step = 'Work Day Log append';
  const wdlSheet = _getOrCreateWorkDayLogSheet_(ss);
  appendRowWithProbing_(
    wdlSheet,
    [
      d.date,                              // 0 Date
      d.wo_id,                             // 1 Work Order #
      String(woRow[1] || ''),              // 2 Prime Contractor
      String(woRow[2] || ''),              // 3 Contract #
      String(woRow[3] || ''),              // 4 Borough
      String(woRow[5] || ''),              // 5 Location
      new Date(),                          // 6 Field Report Submitted At
      crewChief,                           // 7 Crew Chief
      'Pending'                            // 8 Sign-In Status
    ],
    [
      'Date', 'Work Order #', 'Prime Contractor', 'Contract #', 'Borough',
      'Location', 'Field Report Submitted At', 'Crew Chief', 'Sign-In Status'
    ],
    'Work Day Log'
  );

  Logger.log('✅ Work Day Log: queued ' + d.wo_id + ' for sign-in on ' + d.date);

  // CFR JSON generation moves through a separate action
  // (`finalize_field_report_docs`) the client fires AFTER it gets this
  // success response. Sign-In JSON generation is no longer triggered
  // here — sign-ins are filed from the Sign-In tab once the user is
  // ready to gather a whole night's worth of WOs into one sheet.

  // ── Automation Log ────────────────────────────────────────────
  const actionNote = d.wo_complete
    ? 'WO marked COMPLETE — review for invoicing, field report, and production log'
    : '';

  step = 'Automation Log write';
  appendRowWithProbing_(
    ss.getSheetByName('Automation Log'),
    [
      new Date(),
      'Field Report Web App',
      'Field report submitted',
      d.wo_id,
      'Queued for sign-in on ' + d.date,
      newStatus,
      '',
      actionNote
    ],
    ['Timestamp', 'Source', 'Action', 'Related', 'Details', 'Status', 'User', 'Next Steps'],
    'Automation Log'
  );

  _invalidateCacheKeys_(['dashboard_v1']);
  return jsonResponse_({ success: true, wo_id: d.wo_id, status: newStatus });

  } catch (err) {
    // Attach the current phase so the React caller sees e.g.
    //   "[step=Automation Log write] Automation Log → "Status" rejected value …"
    // instead of a bare "Invalid Entry" with no hint.
    const msg = (err && err.message) ? err.message : String(err);
    const wrapped = new Error(`[step=${step}] ${msg}`);
    wrapped.stack = err && err.stack ? err.stack : wrapped.stack;
    throw wrapped;
  }
}


/**
 * Background JSON generation for a Field Report submit. Writes the
 * Contractor Field Report JSON when `wo_complete === true` so the
 * Railway worker can produce the filled CFR PDF.
 *
 * The client fires this as a separate POST right after the main
 * `submit_field_report` returns success. Failures go to Automation Log
 * — they never bubble back to the user because by this point the
 * report data is already safely persisted in the spreadsheet.
 *
 * Sign-In JSON generation moved out of the Field Report flow entirely:
 * sign-ins are filed from the Sign-In tab so a single sheet can cover
 * every WO worked that night for a contract.
 */
function handleFinalizeFieldReportDocs_(body) {
  const d = body.data || {};
  if (!d.wo_id) return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (!d.date)  return jsonResponse_({ error: 'Missing date' }, 400);

  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();
  const woRowIdx = allRows.findIndex((r, i) => i > 0 && String(r[0]) === String(d.wo_id));
  if (woRowIdx === -1) return jsonResponse_({ error: 'Work Order not found: ' + d.wo_id }, 404);
  const woRow = allRows[woRowIdx];
  const issuesAggregate = String(woRow[22] || '').trim();

  const result = { success: true, wo_id: d.wo_id };

  // CFR JSON — only when this submit marked the WO complete
  if (d.wo_complete) {
    try {
      generateContractorFieldReportJson_(d, woRow, ss, issuesAggregate);
      result.cfr = 'ok';
    } catch (err) {
      Logger.log('⚠️ CFR JSON export failed: ' + err);
      result.cfr = 'failed';
      try {
        ss.getSheetByName('Automation Log').appendRow([
          new Date(), 'CFR JSON Export', 'Failed', d.wo_id,
          String(err), 'Error', '', 'Check logs — Contractor Field Report PDF will not be generated'
        ]);
      } catch (logErr) {
        Logger.log('⚠️ Could not write CFR failure to Automation Log: ' + logErr);
      }
    }
  } else {
    result.cfr = 'skipped';
  }

  return jsonResponse_(result);
}


// ── Contractor Field Report export ────────────────────────────

/**
 * Map a Marking Items category name to the CFR template's top-table label.
 * Most categories are 1:1; the few renames below handle plural/singular and
 * "Arrow" → "Arrows" drift between the two surfaces.
 * Returns null for categories that don't appear in the CFR top table
 * (e.g. HVX Crosswalk / Stop Line / Stop Msg — those land in the grid).
 */
function mapCategoryToCFR_(category) {
  const CFR_RENAMES = {
    'Lane Lines':         'Lane Line',
    'L/R Arrow':          'L/R Arrows',
    'Straight Arrow':     'Straight Arrows',
    'Combination Arrow':  'Combination Arrows',
    'Combination Arrow (L/R)': 'Combination Arrows',
    'Bike Lane Arrow':    'Bike Lane Arrows',
    'Bike Lane Symbol':   'Bike Lane Symbols',   // legacy alias
    // Both old and new preform bike symbols land in the same CFR cell;
    // the aggregator sums them via this shared label.
    'Old Bike Symbol (w/ rider)':  'Bike Lane Symbols',
    'New Bike Symbol (just bike)': 'Bike Lane Symbols',
  };
  // No top-table cell in the CFR template for these. HVX Crosswalk /
  // Stop Line land in the grid; Pedestrian Men has no CFR field at all
  // and is surfaced in General Remarks ("PED MEN: {n}") instead.
  // Stop Msg used to live here but the CFR has BOTH a top-table roll-up
  // cell (page0_field23) AND a per-intersection grid cell, so it gets
  // aggregated into the top_table dict and additionally split
  // per-intersection in the grid loop below.
  const GRID_ONLY = { 'HVX Crosswalk': 1, 'Stop Line': 1, 'Pedestrian Men': 1 };

  const c = String(category || '').trim();
  if (!c) return null;
  if (GRID_ONLY[c]) return null;
  return CFR_RENAMES[c] || c;
}


/**
 * Aggregate all Marking Items for a WO into the CFR's top-table and grid
 * payloads. Only rows with Status='Completed' are counted.
 *
 * Returns:
 *   {
 *     top_table: { 'Double Yellow Line': 260, 'Lane Line': 180, ... },
 *     grid: [
 *       { intersection: '5 AV', n: '', e: '', s: '', w: 60,
 *         stop_msg: '', sch8: '', sch10: '', st_line: '' },
 *       ...
 *     ]
 *   }
 *
 * Intersection rows appear in the order they first show up in the Marking
 * Items sheet (which is insertion order — scan-seeded first, then manual).
 * School Msg 8'/10' columns are always blank (no source category today).
 */
function aggregateMarkingItemsForCFR_(ss, woId) {
  const out = { top_table: {}, grid: [] };
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return out;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return out;

  // Col indices (same as computeMarkingRollups_): 1 WO#, 4 Marking Type,
  // 5 Intersection, 6 Direction, 8 Quantity, 9 Unit, 12 Status.
  const woItems = data.slice(1).filter(r =>
    String(r[1] || '').trim() === woId &&
    String(r[12] || '').toLowerCase() === 'completed'
  );
  if (woItems.length === 0) return out;

  // ── Top table: sum Quantity by mapped CFR category ──
  woItems.forEach(r => {
    const category = String(r[4] || '').trim();
    const cfrLabel = mapCategoryToCFR_(category);
    if (!cfrLabel) return;
    const qty = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;
    out.top_table[cfrLabel] = (out.top_table[cfrLabel] || 0) + qty;
  });

  // ── Intersection grid: build per-intersection rows ──
  // Only HVX Crosswalk / Stop Msg / Stop Line items land in the grid.
  // "In between" picks from the Field Report's intersection combobox are
  // stored in the same column as real intersections but formatted as
  // "A – B" (en-dash, spaces) by FieldReport.jsx — they must be excluded
  // so the CFR grid contains only true intersection rows.
  // Preserve first-seen order via an index map. The row is only created
  // after we've classified both the category and the destination column,
  // so a top-table-only category (e.g. Lane Line) with an intersection
  // attached can never produce an empty grid row that shifts the layout.
  const rowByIntersection = {};
  const order = [];
  const isBetweenLocation = (s) => s.indexOf(' – ') !== -1;

  woItems.forEach(r => {
    const intersection = String(r[5] || '').trim();
    if (!intersection || isBetweenLocation(intersection)) return;
    const category = String(r[4] || '').trim();
    const direction = String(r[6] || '').trim().toUpperCase();
    const qty = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;

    let colKey = null;
    if (category === 'HVX Crosswalk' && ['N','E','S','W'].indexOf(direction) !== -1) {
      colKey = direction.toLowerCase();
    } else if (category === 'Stop Msg') {
      colKey = 'stop_msg';
    } else if (category === 'Stop Line') {
      colKey = 'st_line';
    }
    if (!colKey) return;
    // School Msg 8'/10': no source category → stays blank.

    if (!rowByIntersection[intersection]) {
      rowByIntersection[intersection] = {
        intersection: intersection,
        order: '',   // not tracked in Marking Items
        n: '', e: '', s: '', w: '',
        stop_msg: '', sch8: '', sch10: '', st_line: ''
      };
      order.push(intersection);
    }
    const gridRow = rowByIntersection[intersection];
    gridRow[colKey] = (parseFloat(gridRow[colKey]) || 0) + qty;
  });

  out.grid = order.slice(0, 10).map(i => rowByIntersection[i]);
  return out;
}


/**
 * Build the General Remarks segments for completed markings that have NO
 * dedicated CFR field, in priority order:
 *   1. Color Surface — "Color Surface Completed: {SF} SQFT {Color}, …"
 *      (per-color breakdown across Bike Lane / Bus Lane / Pedestrian Space)
 *   2. PED MEN       — "PED MEN: {n}"  (Pedestrian Men symbol count)
 *   3. Other Markings— "Other Markings Completed: {desc} {qty} {unit}, …"
 *      (every completed 'Others' item)
 * Returns an array of non-empty segment strings; the caller joins them
 * with " | " ahead of the Reported Issues. Scope mirrors the CFR
 * aggregation: all completed items for the WO (every date/crew).
 *
 * Col indices: 1 WO#, 4 Marking Type, 7 Description, 8 Quantity, 9 Unit,
 * 10 Color/Material, 12 Status.
 */
function buildMarkingRemarksSegments_(ss, woId) {
  const segments = [];
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return segments;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return segments;

  const items = data.slice(1).filter(r =>
    String(r[1] || '').trim() === woId &&
    String(r[12] || '').toLowerCase() === 'completed'
  );
  if (items.length === 0) return segments;

  // Trim trailing ".0" so 12.0 → "12" but 12.5 stays.
  const fmt = (n) => {
    const v = Math.round((Number(n) + Number.EPSILON) * 100) / 100;
    return String(v);
  };

  // Color Surface = the SF color-surface treatments that have no CFR field.
  // (Bike Lane Green Bar is excluded — it has its own CFR cell, Text_20.)
  const COLOR_SURFACE = { 'Bike Lane': 1, 'Bus Lane': 1, 'Pedestrian Space': 1 };
  const sfByColor = {};          // color -> summed SF
  const sfColorOrder = [];       // first-seen color order
  let   sfNoColor = 0;           // SF with no color recorded
  let   pedMen = 0;
  const others = [];

  items.forEach(r => {
    const category = String(r[4] || '').trim();
    const qty      = parseFloat(r[8]);
    const unit     = String(r[9] || '').trim().toUpperCase();

    if (COLOR_SURFACE[category] && unit === 'SF') {
      if (isNaN(qty) || qty <= 0) return;
      const color = String(r[10] || '').trim();
      if (color && color.toLowerCase() !== 'n/a') {
        if (!(color in sfByColor)) sfColorOrder.push(color);
        sfByColor[color] = (sfByColor[color] || 0) + qty;
      } else {
        sfNoColor += qty;
      }
      return;
    }
    if (category === 'Pedestrian Men') {
      if (!isNaN(qty) && qty > 0) pedMen += qty;
      return;
    }
    if (category === 'Others') {
      if (isNaN(qty) || qty <= 0) return;
      const desc = String(r[7] || '').trim();
      others.push(`${desc ? desc + ' ' : ''}${fmt(qty)}${unit ? ' ' + unit : ''}`.trim());
      return;
    }
  });

  // 1. Color Surface — per-color breakdown
  const colorParts = sfColorOrder.map(c => `${fmt(sfByColor[c])} SQFT ${c}`);
  if (sfNoColor > 0) colorParts.push(`${fmt(sfNoColor)} SQFT`);
  if (colorParts.length) {
    segments.push(`Color Surface Completed: ${colorParts.join(', ')}`);
  }
  // 2. PED MEN
  if (pedMen > 0) segments.push(`PED MEN: ${fmt(pedMen)}`);
  // 3. Other Markings
  if (others.length) {
    segments.push(`Other Markings Completed: ${others.join(', ')}`);
  }
  return segments;
}


/**
 * Build a Contractor Field Report JSON payload from the submitted field
 * report + the WO Tracker row + aggregated Marking Items + the freshly
 * computed (not yet written) issues string. Writes the JSON into the
 * Needs Review / Field Reports Drive folder. The Railway worker
 * (watch_and_fill.py) polls that folder and fills the PDF.
 *
 * Called from handleSubmitFieldReport_ ONLY when d.wo_complete === true.
 *
 * aggregatedIssues: the full Issues Reported string across every submit
 * for this WO (includes the current submit's issues). Built by
 * handleSubmitFieldReport_ as newIssues.
 */
function generateContractorFieldReportJson_(d, woRow, ss, aggregatedIssues, opts) {
  const props          = PropertiesService.getScriptProperties();
  const needsReviewId  = props.getProperty('NEEDS_REVIEW_ID');
  if (!needsReviewId) throw new Error('NEEDS_REVIEW_ID not set');

  const reviewFolder   = DriveApp.getFolderById(needsReviewId);
  const fieldRptFolder = getOrCreateSubfolder_(reviewFolder, 'Field Reports');

  // WO Tracker cols: 0 WO#, 1 Prime Contractor, 2 Contract #, 3 Boro,
  // 5 Location, 6 From, 7 To, 17 Work Start, 18 Work End,
  // 35 Date Entered, 36 School, 37 Prep By.
  const workOrder       = String(woRow[0]  || '').trim();
  const primeContractor = String(woRow[1]  || '').trim();
  const contractNum     = String(woRow[2]  || '').trim();
  const boro            = String(woRow[3]  || '').trim();
  const location        = String(woRow[5]  || '').trim();
  const fromStreet      = String(woRow[6]  || '').trim();
  const toStreet        = String(woRow[7]  || '').trim();
  const workStart       = String(woRow[17] || '').trim();
  const school          = String(woRow[36] || 'NA').trim() || 'NA';
  const prepBy          = String(woRow[37] || '').trim();

  // Date Entered — Sheets auto-parses "Wednesday, January 28, 2026" into
  // a Date object, so stringifying it yields
  //   "Wed Jan 28 2026 03:00:00 GMT-0500 (Eastern Standard Time)"
  // which overflows the CFR field. Reformat Date objects back to long
  // English; otherwise use the raw string the user typed.
  const rawDateEntered = woRow[35];
  let dateEntered = '';
  if (rawDateEntered instanceof Date && !isNaN(rawDateEntered.getTime())) {
    dateEntered = Utilities.formatDate(rawDateEntered, Session.getScriptTimeZone(),
                                       'EEEE, MMMM d, yyyy');
  } else if (rawDateEntered && typeof rawDateEntered.getTime === 'function'
             && !isNaN(rawDateEntered.getTime())) {
    // Fallback: duck-type for cross-realm Date objects where `instanceof Date` fails
    dateEntered = Utilities.formatDate(new Date(rawDateEntered.getTime()),
                                       Session.getScriptTimeZone(),
                                       'EEEE, MMMM d, yyyy');
  } else {
    // String path — strip an appended " HH:MM:SS GMT..." tail if one exists
    let s = String(rawDateEntered || '').trim();
    s = s.replace(/\s+\d{1,2}:\d{2}:\d{2}\s+GMT.*$/, '').trim();
    dateEntered = s;
  }

  // install_from = existing Work Start (first-submit date, already in tracker)
  // install_to   = today's submit date (WO is being completed now)
  //
  // Accepts multiple input shapes and returns M/D/YYYY:
  //   "2026-01-28"                    → "1/28/2026"
  //   "1/28/2026" / "01/28/2026"      → "1/28/2026"
  //   "Wednesday, January 28, 2026"   → "1/28/2026"   (WO scan format)
  //   "January 28, 2026"              → "1/28/2026"
  //   Date object                     → "1/28/2026"
  //   anything else                   → pass through unchanged
  const dateFmt = (val) => {
    if (!val) return '';
    if (val instanceof Date && !isNaN(val.getTime())) {
      return `${val.getMonth() + 1}/${val.getDate()}/${val.getFullYear()}`;
    }
    const s = String(val).trim();
    if (!s) return '';
    // ISO yyyy-mm-dd
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${parseInt(m[2], 10)}/${parseInt(m[3], 10)}/${m[1]}`;
    // Already M/D/YYYY — normalize leading zeros
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      const yyyy = m[3].length === 2 ? '20' + m[3] : m[3];
      return `${parseInt(m[1], 10)}/${parseInt(m[2], 10)}/${yyyy}`;
    }
    // Long English — let Date parse it (handles "Wednesday, January 28, 2026",
    // "January 28, 2026", etc.)
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    }
    return s;
  };

  const aggregated = aggregateMarkingItemsForCFR_(ss, d.wo_id);

  // General Remarks composition. The CFR field renders single-line, so we
  // join every segment with " | ". Priority order (highest first):
  //   1. Color Surface  2. PED MEN  3. Other Markings  4. Reported Issues.
  // The first three come from completed markings that have no dedicated
  // CFR field; the issues are the accumulated per-submit notes.
  const markingSegments = buildMarkingRemarksSegments_(ss, d.wo_id);
  const issueSegments = String(aggregatedIssues || '')
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(Boolean);
  const flatRemarks = [...markingSegments, ...issueSegments].join(' | ');

  const payload = {
    _type:             'contractor_field_report',
    wo_id:             d.wo_id,
    // Preserve exactly as scanned from the WO (e.g. "Wednesday, January 28, 2026")
    date_entered:      dateEntered,
    work_order:        workOrder,
    contractor:        primeContractor,
    contract_number:   contractNum,
    boro:              boro,
    location:          location,
    school:            school,
    from:              fromStreet,
    to:                toStreet,
    install_from:      dateFmt(workStart) || workStart,
    install_to:        dateFmt(d.date),
    general_remarks:   flatRemarks,
    markings:          aggregated.top_table,
    grid:              aggregated.grid,
    prep_by:           prepBy,
    crew_chief:        'Stamati Angelides',
    contractor_notes:  'Oneiro Collection - WBE',
  };

  const isoDate  = (d.date || '').slice(0, 10) || 'unknown';
  // Default filename matches what handleFinalizeFieldReportDocs_ emits.
  // Callers from the admin "edit completed WO" path pass a custom name
  // — used for save-as-new mode to add a `_Updated_<date>` suffix so
  // the regenerated PDF doesn't replace the canonical archived WO doc.
  const fileName = (opts && opts.filename) ? opts.filename : `CFR_${d.wo_id}_${isoDate}.json`;

  // Overwrite any existing file with the same name (re-submit of the
  // completion day is idempotent — same payload, same filename).
  const existing = fieldRptFolder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  fieldRptFolder.createFile(
    fileName,
    JSON.stringify(payload, null, 2),
    MimeType.PLAIN_TEXT
  );

  Logger.log('✅ CFR JSON exported: ' + fileName);
}


// ── action: get_dashboard_data ────────────────────────────────

/**
 * Returns all WO Tracker rows + summary stats for the React dashboard.
 * Called by the Express backend (/api/dashboard) which proxies here.
 *
 * Wrapped in _withScriptCache_ (60s TTL) — full rebuild reads two
 * sheets and only walks Drive for WOs whose Archive Folder URL hasn't
 * been backfilled yet. Mutating handlers call _invalidateCacheKeys_
 * to bust the cache on writes.
 */
function handleGetDashboardData_() {
  const payload = _withScriptCache_('dashboard_v1', 60, _buildDashboardPayload_);
  return jsonResponse_(payload);
}

function _buildDashboardPayload_() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();

  // ── Marking Items counts: read once, group by WO #. ────────────────
  // Marking Items col 1 = WO #, col 12 = Status.
  const miCounts = {};   // { woId: { total, completed } }
  const miSheet  = ss.getSheetByName('Marking Items');
  if (miSheet) {
    const miData = miSheet.getDataRange().getValues();
    miData.slice(1).forEach(r => {
      const woId = String(r[1] || '').trim();
      if (!woId) return;
      const bucket = miCounts[woId] || (miCounts[woId] = { total: 0, completed: 0 });
      bucket.total += 1;
      if (String(r[12] || '').toLowerCase() === 'completed') bucket.completed += 1;
    });
  }

  // ── Drive folder URLs: prefer the persisted Archive Folder URL on
  // col 43; fall back to a read-only Drive walk only when the cell is
  // empty (newly-created WOs that haven't archived a doc yet, or
  // pre-backfill rows). The walk path is the ~150ms-per-WO cost we're
  // designing this whole change to avoid; getWOFolder_ writes the URL
  // back the first time it's resolved so subsequent refreshes stay fast.
  const archiveId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_ID');
  const archiveRoot = archiveId ? DriveApp.getFolderById(archiveId) : null;

  // Format a Date cell as YYYY-MM-DD, leaving non-date values untouched.
  const fmtDate = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    return v == null ? '' : String(v);
  };

  // Lookup so the fallback Drive-walk path can write the resolved URL
  // back to the WO row. The map below loses the original allRows index;
  // this precomputes it so we can find the right row for the setValue.
  const rowIdxByWoId = {};
  for (let i = 1; i < allRows.length; i++) {
    const id = String(allRows[i][0] || '').trim();
    if (id) rowIdxByWoId[id] = i;
  }

  const wos = allRows.slice(1)
    .filter(r => r[0])   // skip blank rows
    .map(r => {
      const woId      = String(r[0]);
      const workType  = String(r[10] || '');
      const isMMA     = workType.toUpperCase() === 'MMA';
      const counts    = miCounts[woId] || { total: 0, completed: 0 };
      let folderUrl = String(r[42] || '').trim() || null;
      if (!folderUrl && archiveRoot) {
        const folder = findWOFolder_(archiveRoot, woId, ss, allRows);
        if (folder) {
          folderUrl = folder.getUrl();
          // Self-heal: cache the URL on the row so we don't pay the
          // ~150ms Drive walk again on the next dashboard refresh.
          // Wrapped in try/catch — cache failure must never break the
          // dashboard read.
          try {
            const rowIdx = rowIdxByWoId[woId];
            if (rowIdx != null) {
              woSheet.getRange(rowIdx + 1, 43).setValue(folderUrl);
            }
          } catch (e) {
            Logger.log('⚠️ dashboard folder-url cache write failed for ' + woId + ': ' + e.message);
          }
        }
      }
      // Per-doc-type lifecycle flags. Only CFR + Invoice still live
      // on the WO row. PL/SI/CP moved to the Doc Lifecycle Log; their
      // entries in the `docs` payload are kept (always false) so the
      // shape stays back-compat for the Download Documents modal.
      const yes = (col) => String(r[col] || '').toLowerCase() === 'yes';
      const docs = {
        cfr:               { done: yes(DOC_TYPE_DONE_COL_['Field Report']), sent: yes(DOC_TYPE_SENT_COL_['Field Report']) },
        production_log:    { done: false, sent: false },
        signin:            { done: false, sent: false },
        certified_payroll: { done: false, sent: false },
        invoice:           { done: yes(DOC_TYPE_DONE_COL_['Invoice']),      sent: yes(DOC_TYPE_SENT_COL_['Invoice']) },
      };
      return {
        id:                  woId,
        contractor:          String(r[1]),
        contract_num:        String(r[2]),
        borough:             String(r[3]),
        contract_id:         String(r[4]),
        location:            String(r[5]),
        from_street:         String(r[6]),
        to_street:           String(r[7]),
        due_date:            fmtDate(r[8]),
        priority:            String(r[9]  || ''),
        work_type:           workType,
        wo_received:         fmtDate(r[11]),
        water_blast:         String(r[12] || ''),
        status:              String(r[15] || 'Received'),
        dispatch_date:       fmtDate(r[16]),
        work_start:          fmtDate(r[17]),
        work_end:            fmtDate(r[18]),
        marking_types:       String(r[19] || ''),
        // Quantity Completed (col 20) — SF total for MMA, LF total for Thermo.
        quantity:            r[20] != null ? String(r[20]) : '',
        quantity_unit:       isMMA ? 'SF' : 'LF',
        paint:               String(r[21] || ''),
        issues:              String(r[22] || ''),
        photos:              String(r[23] || ''),
        // Legacy flat keys — kept for one release, drop in follow-up.
        prod_log:            String(r[24] || ''),
        field_report:        String(r[25] || ''),
        invoice_sent:        String(r[29] || ''),
        payment_recv:        String(r[30] || ''),
        markings_total:      counts.total,
        markings_completed:  counts.completed,
        folder_url:          folderUrl,
        docs,
        // QB invoice fields — populated by recordQbInvoice_ after a
        // successful POST to QuickBooks Online. invoice_doc_number is
        // the human-readable invoice # (QB DocNumber); qb_invoice_id is
        // Intuit's internal Id used to build the view URL. The webapp
        // decorates this payload with invoice_view_url before sending
        // to React (since the base URL differs between sandbox + prod).
        invoice_doc_number:  String(r[26] || '').trim(),
        invoice_date:        fmtDate(r[27]),
        invoice_amount:      r[28] !== '' && r[28] != null ? Number(r[28]) : null,
        qb_invoice_id:       String(r[49] || '').trim(),
      };
    });

  // Pipeline summary counts
  const count = (statusStr) =>
    wos.filter(w => w.status.toLowerCase() === statusStr.toLowerCase()).length;

  const stats = {
    total:       wos.length,
    received:    count('Received'),
    dispatched:  count('Dispatched'),
    in_progress: count('In Progress'),
    complete:    count('Completed')
  };

  // Contractor breakdown
  const byContractor = {};
  wos.forEach(w => {
    const c = w.contractor || 'Unknown';
    byContractor[c] = (byContractor[c] || 0) + 1;
  });

  // WOs needing attention (issues reported, incomplete docs)
  const attention = wos.filter(w =>
    w.status.toLowerCase() !== 'completed' && (
      (w.issues && w.issues.trim()) ||
      (w.status.toLowerCase() === 'in progress' && w.photos.toLowerCase() !== 'yes')
    )
  ).map(w => w.id);

  return { wos, stats, byContractor, attention };
}


// ── action: get_revenue_data ──────────────────────────────────

/**
 * Aggregates priced Marking Items into the Revenue dashboard payload.
 * Body:
 *   { start: 'YYYY-MM-DD', end: 'YYYY-MM-DD' }   // both inclusive
 *
 * Cache key folds in the rotating cache_token, so any WO/Marking-Item
 * mutation that calls _invalidateCacheKeys_ instantly invalidates
 * every revenue cache variant — no need to enumerate date-range keys.
 */
function handleGetRevenueData_(body) {
  const d = body.data || {};
  const start = String(d.start || '').trim();
  const end   = String(d.end   || '').trim();
  // Default to month-to-date when called without an explicit range.
  const today = new Date();
  const ymd   = (dt) => Utilities.formatDate(dt, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const startEff = /^\d{4}-\d{2}-\d{2}$/.test(start)
    ? start
    : ymd(new Date(today.getFullYear(), today.getMonth(), 1));
  const endEff = /^\d{4}-\d{2}-\d{2}$/.test(end) ? end : ymd(today);

  const token = _getCacheToken_();
  const key   = 'revenue_v1_' + token + '_' + startEff + '_' + endEff;
  const payload = _withScriptCache_(key, 60, () =>
    _buildRevenuePayload_(startEff, endEff)
  );
  return jsonResponse_(payload);
}

function _buildRevenuePayload_(startIso, endIso) {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const miSheet = ss.getSheetByName('Marking Items');
  if (!woSheet || !miSheet) {
    return {
      range: { start: startIso, end: endIso },
      totals: { revenue: 0, items: 0, unpriced_items: 0 },
      daily: [], by_contractor: [], by_group: [], top_wos: [], needs_pricing: [],
    };
  }

  // ── WO metadata index (woId → { contractor, contract_num, borough, location })
  // Apply billing remap so the dashboard's pricing lookup matches what
  // QB will actually invoice — sub-prime work on a contract they didn't
  // win prices under their own contract. Without this, those items would
  // hit an unpriced fallback whenever the raw (Brooklyn) pricing rows
  // are absent and silently undercount revenue.
  const woRows = woSheet.getDataRange().getValues();
  const woById = {};
  for (let i = 1; i < woRows.length; i++) {
    const id = String(woRows[i][0] || '').trim();
    if (!id) continue;
    const contractor = String(woRows[i][1] || '').trim();
    const _revMapped = _billingRemap_(
      String(woRows[i][2] || '').split('/')[0].trim(),
      String(woRows[i][3] || '').trim(),
      contractor
    );
    woById[id] = {
      contractor,
      contract_num: _revMapped.contractNum,
      borough:      _revMapped.borough,
      location:     String(woRows[i][5] || '').trim(),
    };
  }

  const rates = _loadContractPricing_(ss);

  // Format any Date / string Date Completed cell to YYYY-MM-DD in local tz.
  const fmt = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    if (!v) return '';
    const m = String(v).match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };

  // ── Walk Marking Items, score each completed-in-range row.
  const miData = miSheet.getDataRange().getValues();
  const allGroups = ['line4', 'line12', 'preformed', 'extruded', 'color_surface'];

  const dailyMap        = {};   // dateIso → { revenue, by_group: { ... } }
  const contractorMap   = {};   // contractor → { revenue, items }
  const groupMap        = {};   // group → { revenue, items }
  const woMap           = {};   // woId → { contractor, location, revenue, items }
  const needsPricing    = [];

  let totalRevenue   = 0;
  let totalItems     = 0;
  let unpricedCount  = 0;

  for (let i = 1; i < miData.length; i++) {
    const r = miData[i];
    const status = String(r[12] || '').toLowerCase();
    if (status !== 'completed') continue;
    const dateIso = fmt(r[11]);
    if (!dateIso || dateIso < startIso || dateIso > endIso) continue;

    const woId = String(r[1] || '').trim();
    const meta = woById[woId];
    if (!meta) continue;   // orphaned item — skip silently

    const item = {
      item_id:        String(r[0] || '').trim(),
      category:       String(r[4] || '').trim(),
      quantity:       r[8],
      unit:           String(r[9] || '').trim(),
      date_completed: dateIso,
    };

    const result = priceMarkingItem_(item, meta, rates);
    totalItems += 1;

    if (result.reason !== null) {
      unpricedCount += 1;
      needsPricing.push({
        wo_id:        woId,
        item_id:      item.item_id,
        category:     item.category,
        qty:          Number(item.quantity) || 0,
        unit:         item.unit,
        reason:       result.reason,
        // Contractor / contract / borough are what the pricing lookup
        // actually used (i.e. post-billing-remap when applicable) so
        // the user can find the right row to add in Contract Pricing.
        contractor:   meta.contractor,
        contract_num: meta.contract_num,
        borough:      meta.borough,
      });
      continue;
    }

    const rev = Number(result.revenue) || 0;
    totalRevenue += rev;

    // daily
    let day = dailyMap[dateIso];
    if (!day) {
      day = { date: dateIso, revenue: 0, by_group: {} };
      allGroups.forEach(g => { day.by_group[g] = 0; });
      dailyMap[dateIso] = day;
    }
    day.revenue += rev;
    day.by_group[result.group] = (day.by_group[result.group] || 0) + rev;

    // contractor
    const cKey = meta.contractor || 'Unknown';
    let cBucket = contractorMap[cKey];
    if (!cBucket) cBucket = contractorMap[cKey] = { contractor: cKey, revenue: 0, items: 0 };
    cBucket.revenue += rev;
    cBucket.items   += 1;

    // group
    let gBucket = groupMap[result.group];
    if (!gBucket) gBucket = groupMap[result.group] = { group: result.group, revenue: 0, items: 0 };
    gBucket.revenue += rev;
    gBucket.items   += 1;

    // wo
    let wBucket = woMap[woId];
    if (!wBucket) {
      wBucket = woMap[woId] = {
        wo_id:      woId,
        contractor: meta.contractor,
        location:   meta.location,
        revenue:    0,
        items:      0,
      };
    }
    wBucket.revenue += rev;
    wBucket.items   += 1;
  }

  const daily = Object.keys(dailyMap).sort().map(k => dailyMap[k]);
  const byContractor = Object.keys(contractorMap)
    .map(k => contractorMap[k])
    .sort((a, b) => b.revenue - a.revenue);
  const byGroup = allGroups.map(g => groupMap[g] || { group: g, revenue: 0, items: 0 });
  const topWos = Object.keys(woMap)
    .map(k => woMap[k])
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, 25);

  // ── Walk Daily Sign-In Data → fully-loaded labor cost per day, split by
  // work type (MMA vs Thermo). Work type comes from the WO-number prefix:
  // PT- = Paint/MMA, RM- = Road Markings/Thermo. A crew never mixes the two
  // in a single shift (different vehicles/material), so the shift's first
  // PT/RM token decides the bucket. Cost uses the same fully-loaded formula
  // as Certified Payroll: ST×(st_rate+st_supp) + OT×(ot_rate+ot_supp), with
  // ST = HoursWorked − Overtime. Per-shift ST/OT here (cols 10/11) can differ
  // slightly from the CP's weekly-40h/YTD OT allocation — fine for a trend.
  const siSheet     = ss.getSheetByName('Daily Sign-In Data');
  const payrollRates = _loadPayrollRates_(ss);
  const laborDailyMap = {};   // dateIso → { mma, thermo }
  let laborMmaTotal    = 0;
  let laborThermoTotal = 0;

  if (siSheet && Array.isArray(payrollRates) && payrollRates.length) {
    const siData = siSheet.getDataRange().getValues();
    for (let i = 1; i < siData.length; i++) {
      const row = siData[i];
      if (!row[0]) continue;
      const dateIso = fmt(row[0]);
      if (!dateIso || dateIso < startIso || dateIso > endIso) continue;

      const m = String(row[1] || '').match(/(PT|RM)-/i);
      if (!m) continue;   // no recognizable work-type prefix — skip
      const workType = m[1].toUpperCase() === 'PT' ? 'mma' : 'thermo';

      const cls  = String(row[7] || '').trim().toUpperCase();
      const rate = _resolvePayrollRate_(payrollRates, cls, dateIso);
      if (!rate) continue;   // no rate effective for this (date, class) — skip

      const hours = Number(row[10]) || 0;
      if (hours <= 0) continue;
      const ot = Number(row[11]) || 0;
      const st = Math.max(0, hours - ot);
      const cost = st * ((rate.st_rate || 0) + (rate.st_supp || 0))
                 + ot * ((rate.ot_rate || 0) + (rate.ot_supp || 0));
      if (cost <= 0) continue;

      let bucket = laborDailyMap[dateIso];
      if (!bucket) bucket = laborDailyMap[dateIso] = { date: dateIso, mma: 0, thermo: 0 };
      bucket[workType] += cost;
      if (workType === 'mma') laborMmaTotal += cost;
      else                    laborThermoTotal += cost;
    }
  }

  const laborDaily = Object.keys(laborDailyMap).sort().map(k => laborDailyMap[k]);

  return {
    range: { start: startIso, end: endIso },
    totals: {
      revenue:        totalRevenue,
      items:          totalItems,
      unpriced_items: unpricedCount,
    },
    daily,
    by_contractor: byContractor,
    by_group:      byGroup,
    top_wos:       topWos,
    needs_pricing: needsPricing,
    labor_daily:   laborDaily,
    labor_totals:  {
      mma:    laborMmaTotal,
      thermo: laborThermoTotal,
      total:  laborMmaTotal + laborThermoTotal,
    },
  };
}


// ── action: list_documents_for_batch ──────────────────────────
//
// Returns the metadata for every archived doc that matches the
// caller's filters. The Express batch-download endpoint takes this
// list, fetches each file's bytes via get_drive_file_bytes, and
// streams them into a zip. No bytes are returned here.
//
// Body shape:
//   {
//     mode: 'unsent' | 'wo_numbers' | 'date_range',
//     contractors: ['Metro Express', ...],     // empty/missing = all
//     doc_types:   ['CFR','Production Log','Sign-In','Certified Payroll','Invoice'],
//     wo_ids:      ['RM-43101', ...],          // mode=wo_numbers
//     date_start:  'YYYY-MM-DD',               // mode=date_range
//     date_end:    'YYYY-MM-DD',
//   }
//
// Returns:
//   {
//     files: [{
//       file_id, filename, mime_type, size,
//       contractor, contract_num, borough,
//       doc_type, wo_ids: [...], work_date: 'YYYY-MM-DD',
//       done: bool, sent: bool,
//     }],
//     counts:    { total, by_doc_type: {...}, by_contractor: {...} },
//     missing:   [{ wo_id, doc_type, reason }],
//     warnings:  [strings],
//     truncated: bool,
//   }
//
// User-facing doc_type values are CFR / Production Log / Sign-In /
// Certified Payroll / Invoice. Internally we translate "CFR" → the
// "Field Report" key used by archiveDocument_ + the lifecycle column
// maps so we don't have to rename those.
const _DOC_TYPE_FRIENDLY_TO_INTERNAL_ = Object.freeze({
  'CFR':               'Field Report',
  'Production Log':    'Production Log',
  'Sign-In':           'Sign-In',
  'Certified Payroll': 'Certified Payroll',
  'Invoice':           'Invoice',
});
const _DOC_TYPE_INTERNAL_TO_FRIENDLY_ = Object.freeze({
  'Field Report':      'CFR',
  'Production Log':    'Production Log',
  'Sign-In':           'Sign-In',
  'Certified Payroll': 'Certified Payroll',
  'Invoice':           'Invoice',
});

const MAX_BATCH_FILES_ = 500;

function handleListDocumentsForBatch_(body) {
  const d = body.data || {};
  const mode = String(d.mode || '').trim();
  const validModes = { unsent: 1, wo_numbers: 1, date_range: 1 };
  if (!validModes[mode]) {
    return jsonResponse_({ error: 'mode must be one of unsent | wo_numbers | date_range' }, 400);
  }

  const contractorsFilter = Array.isArray(d.contractors)
    ? d.contractors.map(s => String(s).trim()).filter(Boolean)
    : [];
  const docTypesFriendly = Array.isArray(d.doc_types) && d.doc_types.length
    ? d.doc_types.map(s => String(s).trim())
    : Object.keys(_DOC_TYPE_FRIENDLY_TO_INTERNAL_);
  const docTypes = docTypesFriendly
    .map(f => _DOC_TYPE_FRIENDLY_TO_INTERNAL_[f])
    .filter(Boolean);
  if (docTypes.length === 0) {
    return jsonResponse_({ error: 'doc_types had no recognized values' }, 400);
  }

  const woIdsFilter = Array.isArray(d.wo_ids)
    ? d.wo_ids.map(s => String(s).trim().toUpperCase()).filter(Boolean)
    : [];
  const dateStart = String(d.date_start || '').trim();
  const dateEnd   = String(d.date_end   || '').trim();
  const includeSIsWithCP = !!d.include_sis_with_cp;   // unsent + CP only
  const includePhotos    = !!d.include_photos;         // wo_numbers only
  if (mode === 'wo_numbers' && woIdsFilter.length === 0) {
    return jsonResponse_({ error: 'wo_numbers mode requires wo_ids' }, 400);
  }
  if (mode === 'date_range' && (!/^\d{4}-\d{2}-\d{2}$/.test(dateStart) ||
                                !/^\d{4}-\d{2}-\d{2}$/.test(dateEnd))) {
    return jsonResponse_({ error: 'date_range mode requires date_start + date_end (YYYY-MM-DD)' }, 400);
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) return jsonResponse_({ error: 'Work Order Tracker not found' }, 500);
  ensureWoTrackerExtraCols_(woSheet);

  const archiveId = PropertiesService.getScriptProperties().getProperty('ARCHIVE_ID');
  if (!archiveId) return jsonResponse_({ error: 'ARCHIVE_ID property not set' }, 500);
  const archiveRoot = DriveApp.getFolderById(archiveId);

  const data = woSheet.getDataRange().getValues();

  // Doc Lifecycle Log → authoritative state for PL / SI / CP. The
  // legacy per-WO Done/Sent columns on the WO Tracker for those types
  // are no longer maintained, so we look up done/sent by Doc ID here.
  const docLog = _readDocLifecycle_(ss);
  const docStateForMaster = (docType, anchorIso, contractNum, borough, contractor, crewChief) => {
    // PL keyed by (date, contractor, crew_chief); SI keyed by
    // (date, contract, borough, crew_chief); CP by (week, contract, borough).
    // Multi-crew PL/SI files carry a `_chief-<slug>` suffix in their name and
    // their lifecycle row is keyed per crew — so the chief MUST be included
    // or the lookup misses and the file is wrongly treated as not-done /
    // not-sent (skipped on unsent, or re-downloaded because never marked sent).
    const id = (docType === 'Production Log')
      ? _plDocId_(anchorIso, contractor, crewChief)
      : (docType === 'Sign-In')
        ? _docLifecycleId_('Sign-In', anchorIso, contractNum, borough, crewChief)
        : _docLifecycleId_(docType, anchorIso, contractNum, borough);
    const row = docLog.byId[id];
    if (!row) return { done: false, sent: false, doc_id: id };
    return { done: !!row.done, sent: !!row.sent, doc_id: id };
  };

  // Date helper — match the same fmtDate used by other handlers.
  const fmtDate = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    if (!v) return '';
    const m = String(v).match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };

  // Build a parallel array of WO row metadata that we'll filter against.
  const allWos = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const woId = String(row[0] || '').trim();
    if (!woId) continue;
    allWos.push({
      wo_id:        woId,
      contractor:   String(row[1] || '').trim(),
      contract_num: String(row[2] || '').split('/')[0].trim(),
      borough:      String(row[3] || '').trim(),
      location:     String(row[5] || '').trim(),
      work_end:     fmtDate(row[18]),
      // Per-doc Done/Sent flags. PL/SI/CP no longer have per-WO cols —
      // those statuses are looked up against the Doc Lifecycle Log
      // elsewhere in the batch flow. Defaulting them to false here.
      done: {
        'Field Report':      String(row[DOC_TYPE_DONE_COL_['Field Report']] || '').toLowerCase() === 'yes',
        'Production Log':    false,
        'Sign-In':           false,
        'Certified Payroll': false,
        'Invoice':           String(row[DOC_TYPE_DONE_COL_['Invoice']]      || '').toLowerCase() === 'yes',
      },
      sent: {
        'Field Report':      String(row[DOC_TYPE_SENT_COL_['Field Report']] || '').toLowerCase() === 'yes',
        'Production Log':    false,
        'Sign-In':           false,
        'Certified Payroll': false,
        'Invoice':           String(row[DOC_TYPE_SENT_COL_['Invoice']]      || '').toLowerCase() === 'yes',
      },
    });
  }
  const woById = {};
  allWos.forEach(w => { woById[w.wo_id] = w; });

  // Mode-specific WO scope.
  let wosInScope;
  if (mode === 'wo_numbers') {
    wosInScope = woIdsFilter.map(id => woById[id]).filter(Boolean);
  } else if (mode === 'date_range') {
    wosInScope = allWos.filter(w => w.work_end >= dateStart && w.work_end <= dateEnd && w.work_end);
  } else {
    // unsent — every WO is potentially in scope; per-doc filtering happens below.
    wosInScope = allWos.slice();
  }
  if (contractorsFilter.length) {
    wosInScope = wosInScope.filter(w => contractorsFilter.indexOf(w.contractor) !== -1);
  }

  const warnings = [];
  if (mode === 'wo_numbers') {
    woIdsFilter.forEach(id => {
      if (!woById[id]) warnings.push('WO not found in tracker: ' + id);
    });
  }

  // Master-copy folder cache, contractor folder cache, WO folder cache —
  // each Drive walk is the slowest part of this handler so we memoize.
  const wofolderCache = {};
  const masterCache = {};   // key contractor|contractNum|borough → { prodFolder, cpFolder }
  const getWoFolder = (w) => {
    const k = w.wo_id;
    if (wofolderCache[k] !== undefined) return wofolderCache[k];
    const f = findWOFolder_(archiveRoot, k, ss, data);
    wofolderCache[k] = f || null;
    return wofolderCache[k];
  };
  const getMasters = (contractor, contractNum, borough) => {
    const k = contractor + '|' + contractNum + '|' + borough;
    if (masterCache[k] !== undefined) return masterCache[k];
    let prod = null, cp = null, signin = null;
    try {
      const cIt = archiveRoot.getFoldersByName(contractor);
      if (cIt.hasNext()) {
        const cFolder = cIt.next();
        const cnIt = cFolder.getFoldersByName(`${contractNum} - ${getBoroughName_(borough)}`);
        if (cnIt.hasNext()) {
          const ctFolder = cnIt.next();
          const plIt = ctFolder.getFoldersByName('Production Logs');
          if (plIt.hasNext()) prod = plIt.next();
          const cpIt = ctFolder.getFoldersByName('Certified Payroll');
          if (cpIt.hasNext()) cp = cpIt.next();
          const siIt = ctFolder.getFoldersByName('Sign-Ins');
          if (siIt.hasNext()) signin = siIt.next();
        }
      }
    } catch (e) {
      warnings.push('getMasters failed for ' + k + ': ' + e.message);
    }
    masterCache[k] = { prod, cp, signin };
    return masterCache[k];
  };

  // Output buckets
  const files = [];
  const missing = [];
  const seenFileIds = new Set();
  // Master docs (PL/SI/CP) are also deduped by their lifecycle doc_id: a PL
  // that covers WOs across multiple (contract, borough) groups is copied into
  // EACH group's master "Production Logs" folder, so the same logical PL is
  // otherwise encountered once per group → duplicate downloads.
  const seenMasterDocIds = new Set();
  const pushFile = (f) => {
    if (seenFileIds.has(f.file_id)) return;
    seenFileIds.add(f.file_id);
    files.push(f);
  };

  // Helper: determine if the WO row qualifies for the requested doc
  // type given the mode. Returns { include, missingReason? }.
  const qualifies = (w, internalType) => {
    if (mode === 'unsent') {
      if (!w.done[internalType])  return { include: false };               // never generated
      if (w.sent[internalType])   return { include: false };               // already sent
      return { include: true };
    }
    if (mode === 'wo_numbers') {
      if (!w.done[internalType]) {
        return { include: false, missingReason: internalType + ' Done? = No' };
      }
      return { include: true };
    }
    // date_range
    if (!w.done[internalType]) {
      return { include: false, missingReason: internalType + ' Done? = No' };
    }
    return { include: true };
  };

  // ── CFR (Field Report) — archived as the merged WO doc.
  //
  // archiveDocument_ runs _replaceArchivedWODoc_ on Field Report files,
  // which renames the file to canonical `WO_<wo_id>.pdf` and trashes
  // any prior WO doc. So the file in the WO folder is NOT named CFR_*.
  // Match precedence mirrors lookup_archived_wo_pdf (Code.js:1818):
  //   1. Canonical `WO_<wo_id>.pdf`.
  //   2. Any non-aux PDF (the original scan, before its first CFR was
  //      merged) as a legacy fallback.
  if (docTypes.indexOf('Field Report') !== -1) {
    wosInScope.forEach(w => {
      if (files.length >= MAX_BATCH_FILES_) return;
      const q = qualifies(w, 'Field Report');
      if (!q.include) {
        if (q.missingReason) {
          missing.push({ wo_id: w.wo_id, doc_type: 'CFR', reason: q.missingReason });
        }
        return;
      }
      const folder = getWoFolder(w);
      if (!folder) {
        missing.push({ wo_id: w.wo_id, doc_type: 'CFR', reason: 'archive folder not found' });
        return;
      }
      const canonical = `WO_${w.wo_id}.pdf`;
      let canonicalHit = null;
      let fallbackHit  = null;
      const fIter = folder.getFiles();
      while (fIter.hasNext()) {
        const f = fIter.next();
        if (f.getMimeType() !== 'application/pdf') continue;
        const name = f.getName();
        if (name === canonical) { canonicalHit = f; break; }
        if (!_isAuxDocName_(name) && !fallbackHit) fallbackHit = f;
      }
      const target = canonicalHit || fallbackHit;
      if (!target) {
        missing.push({ wo_id: w.wo_id, doc_type: 'CFR', reason: 'Done? = Yes but no merged WO doc found in folder' });
        return;
      }
      pushFile({
        file_id:      target.getId(),
        filename:     target.getName(),
        mime_type:    target.getMimeType(),
        size:         target.getSize(),
        contractor:   w.contractor,
        contract_num: w.contract_num,
        borough:      w.borough,
        location:     w.location,
        doc_type:     'CFR',
        wo_ids:       [w.wo_id],
        work_date:    w.work_end,
        done:         true,
        sent:         w.sent['Field Report'],
      });
    });
  }

  // ── Invoice — Invoice_<num>_<wo_id>.<ext> in the WO folder.
  if (docTypes.indexOf('Invoice') !== -1) {
    wosInScope.forEach(w => {
      if (files.length >= MAX_BATCH_FILES_) return;
      const q = qualifies(w, 'Invoice');
      if (!q.include) {
        if (q.missingReason) {
          missing.push({ wo_id: w.wo_id, doc_type: 'Invoice', reason: q.missingReason });
        }
        return;
      }
      const folder = getWoFolder(w);
      if (!folder) {
        missing.push({ wo_id: w.wo_id, doc_type: 'Invoice', reason: 'archive folder not found' });
        return;
      }
      let foundAny = false;
      const fIter = folder.getFiles();
      while (fIter.hasNext()) {
        const f = fIter.next();
        const name = f.getName();
        if (!/^Invoice_/i.test(name)) continue;
        // Restrict to invoices that mention this WO id, so we don't
        // grab a sibling WO's invoice if more than one was filed.
        if (name.toUpperCase().indexOf(w.wo_id.toUpperCase()) === -1) continue;
        pushFile({
          file_id:      f.getId(),
          filename:     name,
          mime_type:    f.getMimeType(),
          size:         f.getSize(),
          contractor:   w.contractor,
          contract_num: w.contract_num,
          borough:      w.borough,
          location:     w.location,
          doc_type:     'Invoice',
          wo_ids:       [w.wo_id],
          work_date:    w.work_end,
          done:         true,
          sent:         w.sent['Invoice'],
        });
        foundAny = true;
      }
      if (!foundAny) {
        missing.push({ wo_id: w.wo_id, doc_type: 'Invoice', reason: 'Done? = Yes but no Invoice_*.pdf found in folder' });
      }
    });
  }

  // ── Master-copy docs: Production Log + Certified Payroll
  // Walk each contract's master subfolder once; filter by date range
  // (date_range mode) or include all (other modes filtered by WO).
  // For each file, resolve the WOs it covers via getWOsForDate_ /
  // getWOsForPayrollWeek_ so the manifest can list them.
  const groupKey = (w) => w.contractor + '|' + w.contract_num + '|' + w.borough;
  const woGroups = {};   // group key → array of WO rows
  wosInScope.forEach(w => {
    const k = groupKey(w);
    if (!woGroups[k]) woGroups[k] = [];
    woGroups[k].push(w);
  });

  const masterDocs = docTypes.filter(t => t === 'Production Log' || t === 'Certified Payroll' || t === 'Sign-In');
  masterDocs.forEach(docType => {
    Object.keys(woGroups).forEach(gk => {
      if (files.length >= MAX_BATCH_FILES_) return;
      const groupWos = woGroups[gk];
      const [contractor, contractNum, borough] = gk.split('|');
      const masters = getMasters(contractor, contractNum, borough);
      const masterFolder = docType === 'Production Log' ? masters.prod
                         : docType === 'Sign-In'        ? masters.signin
                                                        : masters.cp;
      if (!masterFolder) return;

      // Filename prefix per doc type — guards against accidentally
      // picking up unrelated PDFs that happened to land in the master.
      const masterPrefix = docType === 'Production Log' ? /^Production_Log_/i
                         : docType === 'Sign-In'        ? /^SignIn_/i
                                                        : /^Certified_Payroll_/i;

      // For master-docs (PL / SI / CP), we no longer use the per-WO
      // qualifies() check (those columns are stale post-refactor).
      // Instead, the Doc Lifecycle Log is the source of truth for
      // Done/Sent state per (date, contract, borough) or (week,
      // contract, borough). We still need the scope check though:
      // ensure at least one of this group's WOs is in the user's
      // selected scope (wo_numbers / date_range / unsent semantics
      // applied below per-file via the Log lookup).

      const fIter = masterFolder.getFiles();
      while (fIter.hasNext()) {
        if (files.length >= MAX_BATCH_FILES_) break;
        const f = fIter.next();
        const name = f.getName();
        if (!masterPrefix.test(name)) continue;
        const dm = name.match(/(\d{4}-\d{2}-\d{2})/);
        if (!dm) continue;
        const fileDate = dm[1];

        // Apply date_range filter — natural date per doc type:
        //   PL/SI: production date / sign-in date (file's own anchor)
        //   CP:    week_start..week_start+6 must overlap [dateStart, dateEnd]
        if (mode === 'date_range') {
          if (docType === 'Certified Payroll') {
            const [y, m, dd] = fileDate.split('-').map(Number);
            const weekEnd = Utilities.formatDate(
              new Date(y, m - 1, dd + 6), CONFIG.TIMEZONE, 'yyyy-MM-dd');
            if (weekEnd < dateStart || fileDate > dateEnd) continue;
          } else {
            if (fileDate < dateStart || fileDate > dateEnd) continue;
          }
        }

        // Resolve covered WOs (used for the manifest + the wo_numbers
        // overlap gate; not used to gate date_range any more — that's
        // the file's own date that matters).
        let coveredWoIds;
        if (docType === 'Production Log' || docType === 'Sign-In') {
          const map = getWOsForDate_(new Date(fileDate + 'T12:00:00'), ss);
          const keyMatch = Object.keys(map).find(k => {
            const parts = k.split('|');
            return parts[0] === contractor && parts[1] === contractNum && parts[2] === borough;
          });
          coveredWoIds = keyMatch ? map[keyMatch].map(x => x.id) : [];
        } else {
          const wos = getWOsForPayrollWeek_(contractNum, borough, new Date(fileDate + 'T12:00:00'), ss);
          coveredWoIds = wos.map(x => x.id);
        }
        if (coveredWoIds.length === 0) continue;

        // wo_numbers mode only: at least one covered WO must overlap
        // with the user's selected list. date_range and unsent gate
        // the file via its own date / Log state, not via the WO list.
        if (mode === 'wo_numbers') {
          const inScopeIds = new Set(groupWos.map(w => w.wo_id));
          if (!coveredWoIds.some(id => inScopeIds.has(id))) continue;
        }

        // Look up doc state in the Log. Anchor for PL/SI is the file's
        // shift date; for CP it's the week-start (which is what's
        // already in the filename).
        // Multi-crew PL/SI files carry a `_chief-<slug>` suffix and their
        // lifecycle row is keyed per crew — pull the chief out of the
        // filename so the state lookup (and mark-sent) hit the right row.
        // CP filenames never carry a chief.
        let fileChief = '';
        if (docType === 'Production Log' || docType === 'Sign-In') {
          const cm = name.match(/_chief-([A-Za-z0-9]+)/);
          fileChief = cm ? cm[1] : '';
        }
        const logState = docStateForMaster(docType, fileDate, contractNum, borough, contractor, fileChief);

        // unsent mode: include only if Done=Yes and Sent=No (or for
        // SI, Done=Yes — SI Sent isn't tracked, so missing == done false).
        if (mode === 'unsent') {
          if (!logState.done) continue;
          if (docType !== 'Sign-In' && logState.sent) continue;
          // SI has no separate Sent state — once Done it never becomes
          // "unsent." Keep including SI rows so they show up in batches.
        }

        // Collapse the same logical doc encountered once per (contract,
        // borough) group — a PL spans contracts so it's copied into each
        // group's master folder, and repeated archives can leave extra
        // same-identity copies. One zip entry + one mark-sent per doc_id.
        if (seenMasterDocIds.has(logState.doc_id)) continue;
        seenMasterDocIds.add(logState.doc_id);

        const allSent = !!logState.sent;

        pushFile({
          file_id:      f.getId(),
          filename:     name,
          mime_type:    f.getMimeType(),
          size:         f.getSize(),
          contractor,
          contract_num: contractNum,
          borough,
          doc_type:     _DOC_TYPE_INTERNAL_TO_FRIENDLY_[docType],
          wo_ids:       coveredWoIds,
          work_date:    fileDate,
          doc_id:       logState.doc_id,
          done:         true,
          sent:         allSent,
        });

        // CP + SI bundle: when admin checked "Include matching Sign-Ins"
        // in unsent mode, attach every Done SI for this CP's week into
        // a sibling folder so the recipient sees one zip with the CP
        // at the root and "Sign-Ins for CP …/" alongside.
        if (docType === 'Certified Payroll' && includeSIsWithCP && mode === 'unsent') {
          const [yy, mm, ddd] = fileDate.split('-').map(Number);
          const weekStartDt = new Date(yy, mm - 1, ddd);
          const weekEndIso = Utilities.formatDate(
            new Date(yy, mm - 1, ddd + 6), CONFIG.TIMEZONE, 'yyyy-MM-dd'
          );
          const siFolder = masters.signin;
          if (siFolder) {
            const bundleDir = 'Sign-Ins for CP ' + contractNum + ' ' + borough + ' ' + fileDate;
            const siIter = siFolder.getFiles();
            while (siIter.hasNext()) {
              if (files.length >= MAX_BATCH_FILES_) break;
              const sf = siIter.next();
              const sname = sf.getName();
              if (!/^SignIn_/i.test(sname)) continue;
              const sm = sname.match(/(\d{4}-\d{2}-\d{2})/);
              if (!sm) continue;
              const sDate = sm[1];
              if (sDate < fileDate || sDate > weekEndIso) continue;
              // Only include SIs that are Done in the Log.
              const siState = docStateForMaster('Sign-In', sDate, contractNum, borough, contractor);
              if (!siState.done) continue;
              pushFile({
                file_id:      sf.getId(),
                filename:     sname,
                zip_path:     bundleDir + '/' + sname,
                mime_type:    sf.getMimeType(),
                size:         sf.getSize(),
                contractor,
                contract_num: contractNum,
                borough,
                doc_type:     'Sign-In',
                wo_ids:       [],
                work_date:    sDate,
                done:         true,
                sent:         false,        // bundled — not part of mark-sent flip
                bundled:      true,         // signals "do not flip flags on download"
              });
            }
          }
        }
      }
    });
  });

  // Sign-In is now handled by the master-doc loop above (it lives in
  // its own contract-level "Sign-Ins/" master folder, mirroring
  // Production Logs and Certified Payroll).

  // Photos — only for wo_numbers mode when admin checked "Include Photos."
  // Walks each WO's <wo>/Photos/ subfolder and bundles every image into
  // a per-WO subdirectory in the zip. Counts toward MAX_BATCH_FILES_.
  if (includePhotos && mode === 'wo_numbers') {
    wosInScope.forEach(w => {
      if (files.length >= MAX_BATCH_FILES_) return;
      const folder = getWoFolder(w);
      if (!folder) return;
      let photosFolder = null;
      try {
        const it = folder.getFoldersByName('Photos');
        if (it.hasNext()) photosFolder = it.next();
      } catch (e) { /* ignore */ }
      if (!photosFolder) return;
      const bundleDir = w.wo_id + '/Photos';
      const fIt = photosFolder.getFiles();
      while (fIt.hasNext()) {
        if (files.length >= MAX_BATCH_FILES_) break;
        const f = fIt.next();
        pushFile({
          file_id:      f.getId(),
          filename:     f.getName(),
          zip_path:     bundleDir + '/' + f.getName(),
          mime_type:    f.getMimeType(),
          size:         f.getSize(),
          contractor:   w.contractor,
          contract_num: w.contract_num,
          borough:      w.borough,
          location:     w.location,
          doc_type:     'Photo',
          wo_ids:       [w.wo_id],
          work_date:    w.work_end,
          done:         true,
          sent:         false,
          bundled:      true,
        });
      }
    });
  }

  // Counts
  const counts = { total: files.length, by_doc_type: {}, by_contractor: {} };
  files.forEach(f => {
    counts.by_doc_type[f.doc_type]    = (counts.by_doc_type[f.doc_type]    || 0) + 1;
    counts.by_contractor[f.contractor]= (counts.by_contractor[f.contractor]|| 0) + 1;
  });

  return jsonResponse_({
    files,
    counts,
    missing,
    warnings,
    truncated: files.length >= MAX_BATCH_FILES_,
  });
}


// ── action: get_production_data ───────────────────────────────

/**
 * Aggregates Marking Items quantities by unit (SF / LF / EA) for the
 * Production dashboard tab. Same shape philosophy as
 * handleGetRevenueData_ but no pricing math — pure quantity rollup.
 *
 * Body:
 *   { start: 'YYYY-MM-DD', end: 'YYYY-MM-DD' }   // both inclusive
 *
 * Cache key folds in the rotating cache_token, so Marking Item
 * mutations (Field Report submits, manual edits, etc) bust every
 * production cache variant via _invalidateCacheKeys_.
 */
function handleGetProductionData_(body) {
  const d = body.data || {};
  const start = String(d.start || '').trim();
  const end   = String(d.end   || '').trim();
  const today = new Date();
  const ymd   = (dt) => Utilities.formatDate(dt, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const startEff = /^\d{4}-\d{2}-\d{2}$/.test(start)
    ? start
    : ymd(new Date(today.getFullYear(), today.getMonth(), 1));
  const endEff = /^\d{4}-\d{2}-\d{2}$/.test(end) ? end : ymd(today);

  const token = _getCacheToken_();
  const key   = 'production_v1_' + token + '_' + startEff + '_' + endEff;
  const payload = _withScriptCache_(key, 60, () =>
    _buildProductionPayload_(startEff, endEff)
  );
  return jsonResponse_(payload);
}

function _buildProductionPayload_(startIso, endIso) {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const miSheet = ss.getSheetByName('Marking Items');
  if (!woSheet || !miSheet) {
    return {
      range: { start: startIso, end: endIso },
      totals: { SF: 0, LF: 0, EA: 0, items: 0 },
      daily: [], by_contractor: [], by_category: [], top_wos: [],
    };
  }

  // ── WO metadata index
  const woRows = woSheet.getDataRange().getValues();
  const woById = {};
  for (let i = 1; i < woRows.length; i++) {
    const id = String(woRows[i][0] || '').trim();
    if (!id) continue;
    woById[id] = {
      contractor: String(woRows[i][1] || '').trim(),
      borough:    String(woRows[i][3] || '').trim(),
      location:   String(woRows[i][5] || '').trim(),
    };
  }

  // Helpers
  const fmt = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    if (!v) return '';
    const m = String(v).match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };
  const blankUnits = () => ({ SF: 0, LF: 0, EA: 0 });

  // ── Walk Marking Items
  const miData = miSheet.getDataRange().getValues();

  const dailyMap      = {};   // dateIso → { date, SF, LF, EA }
  const contractorMap = {};   // contractor → { contractor, SF, LF, EA, items }
  const categoryMap   = {};   // <category|unit> → { category, unit, qty, items }
  const woMap         = {};   // woId → { wo_id, contractor, location, SF, LF, EA, items }

  let totalSF = 0;
  let totalLF = 0;
  let totalEA = 0;
  let totalItems = 0;

  for (let i = 1; i < miData.length; i++) {
    const r = miData[i];
    const status = String(r[12] || '').toLowerCase();
    if (status !== 'completed') continue;
    const dateIso = fmt(r[11]);
    if (!dateIso || dateIso < startIso || dateIso > endIso) continue;

    const woId = String(r[1] || '').trim();
    const meta = woById[woId];
    if (!meta) continue;

    const category = String(r[4] || '').trim();
    const unitRaw  = String(r[9] || '').trim().toUpperCase();
    const unit     = (unitRaw === 'SF' || unitRaw === 'LF' || unitRaw === 'EA') ? unitRaw : null;
    if (!unit) continue;   // unrecognized unit — skip rather than mis-bucket

    const qty = Number(r[8]);
    if (isNaN(qty) || qty <= 0) continue;

    totalItems += 1;
    if (unit === 'SF') totalSF += qty;
    else if (unit === 'LF') totalLF += qty;
    else                   totalEA += qty;

    // daily
    let day = dailyMap[dateIso];
    if (!day) {
      day = { date: dateIso, SF: 0, LF: 0, EA: 0 };
      dailyMap[dateIso] = day;
    }
    day[unit] += qty;

    // contractor
    const cKey = meta.contractor || 'Unknown';
    let cBucket = contractorMap[cKey];
    if (!cBucket) {
      cBucket = contractorMap[cKey] = Object.assign({ contractor: cKey, items: 0 }, blankUnits());
    }
    cBucket[unit] += qty;
    cBucket.items += 1;

    // category (keyed by category+unit so "Others" entries with mixed
    // units don't collapse into one inscrutable row)
    const catKey = category + '|' + unit;
    let catBucket = categoryMap[catKey];
    if (!catBucket) {
      catBucket = categoryMap[catKey] = { category, unit, qty: 0, items: 0 };
    }
    catBucket.qty += qty;
    catBucket.items += 1;

    // wo
    let wBucket = woMap[woId];
    if (!wBucket) {
      wBucket = woMap[woId] = Object.assign({
        wo_id: woId, contractor: meta.contractor, location: meta.location, items: 0
      }, blankUnits());
    }
    wBucket[unit] += qty;
    wBucket.items += 1;
  }

  const daily = Object.keys(dailyMap).sort().map(k => dailyMap[k]);

  // ── Shift metrics (count, % of days, longest streak) ──────────
  // A "shift" here = a calendar day with at least one completed
  // Marking Item. Cleaner than reading Daily Sign-In Data because it
  // captures actual production output (not days where the crew was
  // dispatched but didn't produce markings, e.g. waterblasting only).
  const workedDates = new Set(Object.keys(dailyMap));
  const shiftsCount = workedDates.size;

  // Calendar-day count for the range, inclusive. UTC anchors avoid
  // DST drift miscounting.
  const startMs = Date.UTC(
    Number(startIso.slice(0, 4)),
    Number(startIso.slice(5, 7)) - 1,
    Number(startIso.slice(8, 10))
  );
  const endMs = Date.UTC(
    Number(endIso.slice(0, 4)),
    Number(endIso.slice(5, 7)) - 1,
    Number(endIso.slice(8, 10))
  );
  const daysInRange = Math.max(1, Math.round((endMs - startMs) / 86400000) + 1);
  const pctDaysWorked = Math.round((shiftsCount / daysInRange) * 1000) / 10;

  // Longest streak: walk the range day-by-day, counting consecutive
  // days that appear in workedDates. Reset on any gap.
  const pad = (n) => String(n).padStart(2, '0');
  let longestStreak = 0;
  let currentRun    = 0;
  for (let ms = startMs; ms <= endMs; ms += 86400000) {
    const d = new Date(ms);
    const iso = d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    if (workedDates.has(iso)) {
      currentRun += 1;
      if (currentRun > longestStreak) longestStreak = currentRun;
    } else {
      currentRun = 0;
    }
  }

  // Contractors sorted by total quantity contribution. We sort on a
  // single combined number so the top-of-list is the biggest producer
  // even when its work mix is split across units.
  const contractorScore = (b) => b.SF + b.LF + b.EA;
  const byContractor = Object.keys(contractorMap)
    .map(k => contractorMap[k])
    .sort((a, b) => contractorScore(b) - contractorScore(a));

  const byCategory = Object.keys(categoryMap)
    .map(k => categoryMap[k])
    .sort((a, b) => b.qty - a.qty);

  const topWos = Object.keys(woMap)
    .map(k => woMap[k])
    .sort((a, b) => (b.SF + b.LF + b.EA) - (a.SF + a.LF + a.EA))
    .slice(0, 25);

  return {
    range: { start: startIso, end: endIso },
    totals: { SF: totalSF, LF: totalLF, EA: totalEA, items: totalItems },
    shifts: {
      count:           shiftsCount,
      days_in_range:   daysInRange,
      pct_days_worked: pctDaysWorked,
      longest_streak:  longestStreak,
    },
    daily,
    by_contractor: byContractor,
    by_category:   byCategory,
    top_wos:       topWos,
  };
}


// ── action: get_doc_status_calendar ───────────────────────────
//
// Returns the calendar payload for the Doc Status tab. Two calendars
// per request (day + week) for the requested month, plus an
// always-current pending list of every doc instance still missing a
// Done or Sent flag.
//
// Body: { month: 'YYYY-MM' }   (defaults to current month in TZ)
//
// Source data:
//   - Doc Lifecycle Log → authoritative state per doc instance.
//   - Daily Sign-In Data → enumerates days+weeks that had work but
//     don't yet have a Log row (admin can see "you worked Tuesday,
//     no PL filed yet" on the calendar even before backfill).
//
// Color rollup per cell:
//   gray  — work happened, nothing done.
//   amber — at least one of the breakdown's docs is partially done.
//   green — every breakdown is fully sent (or done, for SI which has
//           no separate Sent state).
function handleGetDocStatusCalendar_(body) {
  const d = body.data || {};
  const monthRaw = String(d.month || '').trim();
  const today = new Date();
  const monthEff = /^\d{4}-\d{2}$/.test(monthRaw)
    ? monthRaw
    : Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM');

  const token = _getCacheToken_();
  const key = 'doc_status_v1_' + token + '_' + monthEff;
  const payload = _withScriptCache_(key, 60, () => _buildDocStatusPayload_(monthEff));
  return jsonResponse_(payload);
}

function _buildDocStatusPayload_(monthIso) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Read Doc Lifecycle Log + Work Day Log once. Work Day Log is the
  // canonical "days that worked" source — Daily Sign-In Data only
  // populates once a sign-in sheet has been filled, so it lags.
  const { rows: logRows, byId: logById } = _readDocLifecycle_(ss);
  const wdlSheet = ss.getSheetByName('Work Day Log');
  const wdlData  = wdlSheet ? wdlSheet.getDataRange().getValues() : [];

  // Per-contractor PL eligibility: only contractors in
  // CONFIG.PRODUCTION_LOG_CONTRACTORS need PLs. Other primes' breakdown
  // entries get pl_required=false and the calendar/popover/pending list
  // skip the PL surface for them. Edit the CONFIG list to enable a new
  // prime — no migration needed.
  const PL_ENABLED = new Set(
    (CONFIG.PRODUCTION_LOG_CONTRACTORS || []).map(s => String(s).trim()).filter(Boolean)
  );

  // Date helpers
  const ymd = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  };
  const weekStartIsoFor = (date) => {
    const d = (date instanceof Date) ? date : new Date(String(date) + 'T12:00:00');
    const dow = d.getDay();
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate() - dow);
    return Utilities.formatDate(start, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  };

  // Month bounds
  const [yyyy, mm] = monthIso.split('-').map(Number);
  const monthStart = monthIso + '-01';
  // Use day 0 of next month → last day of this month
  const monthEnd = Utilities.formatDate(
    new Date(yyyy, mm, 0), CONFIG.TIMEZONE, 'yyyy-MM-dd'
  );

  // Walk Work Day Log → group at TWO granularities:
  //   contractorDays[dateIso|contractor|crewChief] = per-crew PL shell
  //     .contracts[cn|bor] = per-(contract, borough) sub-entry (one per SI within this crew)
  //   weekTuples[weekStart|cn|bor] = per-(week, contract, borough) (one per CP)
  //
  // Multi-crew: a contractor running two crews on the same date gets
  // TWO contractorDays entries (one per chief) → two PLs, each with
  // their own contracts list. CP stays single-crew per (week, contract,
  // borough) by design.
  //
  // WDL columns (0-idx): 0=Date, 1=WO, 2=Contractor, 3=Contract#,
  // 4=Borough, 7=Crew Chief (blank for legacy / single-crew rows).
  const contractorDays = {};
  const weekTuples = {};

  for (let i = 1; i < wdlData.length; i++) {
    const r = wdlData[i];
    const dateIso = ymd(r[0]);
    if (!dateIso) continue;
    const woId        = String(r[1] || '').trim();
    const contractor  = String(r[2] || '').trim();
    const contractNum = String(r[3] || '').split('/')[0].trim();
    const borough     = String(r[4] || '').trim();
    const crewChief   = String(r[7] || '').trim();
    if (!woId || !contractor || !contractNum || !borough) continue;

    const cdKey = dateIso + '|' + contractor + '|' + crewChief;
    if (!contractorDays[cdKey]) {
      contractorDays[cdKey] = {
        date: dateIso, contractor, crew_chief: crewChief,
        wo_ids: new Set(),
        contracts: {},
      };
    }
    const cdEntry = contractorDays[cdKey];
    cdEntry.wo_ids.add(woId);

    const cKey = contractNum + '|' + borough;
    if (!cdEntry.contracts[cKey]) {
      cdEntry.contracts[cKey] = {
        contract_num: contractNum, borough,
        wo_ids: new Set(),
      };
    }
    cdEntry.contracts[cKey].wo_ids.add(woId);

    const wkStart = weekStartIsoFor(dateIso);
    const wKey = wkStart + '|' + contractNum + '|' + borough;
    if (!weekTuples[wKey]) {
      weekTuples[wKey] = {
        week_start: wkStart, contractor, contract_num: contractNum, borough,
        wo_ids: new Set(),
      };
    }
    weekTuples[wKey].wo_ids.add(woId);
  }

  // Build the days[] entries — only those whose date falls in the month.
  // Each day cell's breakdown is a list of contractor groups; each group
  // has the contractor's PL state + a list of per-(contract, borough)
  // sub-entries with their SI state.
  const dayCellByDate = {};  // dateIso → { date, status, breakdown[] }
  Object.values(contractorDays).forEach(cd => {
    if (cd.date < monthStart || cd.date > monthEnd) return;
    const plRequired = PL_ENABLED.has(cd.contractor);
    const plId = _plDocId_(cd.date, cd.contractor, cd.crew_chief);
    const pl = logById[plId];
    const contractsArr = Object.values(cd.contracts).map(c => {
      const siId = _docLifecycleId_('Sign-In', cd.date, c.contract_num, c.borough, cd.crew_chief);
      const si = logById[siId];
      return {
        contract_num: c.contract_num,
        borough:      c.borough,
        crew_chief:   cd.crew_chief,
        wo_ids:       Array.from(c.wo_ids).sort(),
        si: si
          ? { doc_id: si.doc_id, done: si.done }
          : { doc_id: siId,      done: false },
      };
    }).sort((a, b) => (a.contract_num + a.borough).localeCompare(b.contract_num + b.borough));

    const bdRow = {
      contractor:  cd.contractor,
      crew_chief:  cd.crew_chief,
      pl_required: plRequired,
      wo_ids:      Array.from(cd.wo_ids).sort(),
      pl: plRequired
        ? (pl
            ? { doc_id: pl.doc_id, done: pl.done, sent: pl.sent }
            : { doc_id: plId,      done: false,    sent: false })
        : null,
      contracts: contractsArr,
    };

    if (!dayCellByDate[cd.date]) {
      dayCellByDate[cd.date] = { date: cd.date, status: 'gray', breakdown: [] };
    }
    dayCellByDate[cd.date].breakdown.push(bdRow);
  });

  // Day cell status rollup: green if every contractor's PL (when required)
  // is Done+Sent AND every contract's SI is Done. Amber if any partial.
  // Gray if all empty. PL is treated as "satisfied" for non-required
  // contractors (they don't need a PL to count as full).
  Object.values(dayCellByDate).forEach(cell => {
    let allFull = cell.breakdown.length > 0, allEmpty = true;
    cell.breakdown.forEach(b => {
      const plFull  = !b.pl_required || (b.pl?.done && b.pl?.sent);
      const plEmpty = !b.pl_required || (!b.pl?.done && !b.pl?.sent);
      const siAllDone  = b.contracts.every(c => c.si.done);
      const siAllBlank = b.contracts.every(c => !c.si.done);
      const full  = plFull  && siAllDone;
      const empty = plEmpty && siAllBlank;
      if (!full)  allFull  = false;
      if (!empty) allEmpty = false;
    });
    cell.status = allFull ? 'green' : (allEmpty ? 'gray' : 'amber');
  });

  // Build the weeks[] entries — include the leading week whose Sunday
  // falls in the previous month (the calendar UI renders that week as
  // its first row, e.g. May view shows the "Apr 26 – May 2" week at top).
  const firstSunIso = weekStartIsoFor(monthStart);
  const weekCellByStart = {};
  Object.values(weekTuples).forEach(w => {
    if (w.week_start < firstSunIso || w.week_start > monthEnd) return;
    const cpId = _docLifecycleId_('Certified Payroll', w.week_start, w.contract_num, w.borough);
    const cp = logById[cpId];
    const wos = Array.from(w.wo_ids).sort();
    const bdRow = {
      contractor:   w.contractor,
      contract_num: w.contract_num,
      borough:      w.borough,
      wo_ids:       wos,
      cp: cp
        ? { doc_id: cp.doc_id, done: cp.done, sent: cp.sent }
        : { doc_id: cpId,      done: false,    sent: false },
    };
    if (!weekCellByStart[w.week_start]) {
      weekCellByStart[w.week_start] = { week_start: w.week_start, status: 'gray', breakdown: [] };
    }
    weekCellByStart[w.week_start].breakdown.push(bdRow);
  });
  Object.values(weekCellByStart).forEach(cell => {
    let anyPartial = false, allFull = cell.breakdown.length > 0, allEmpty = true;
    cell.breakdown.forEach(b => {
      const full  = b.cp.done && b.cp.sent;
      const empty = !b.cp.done && !b.cp.sent;
      if (!full)  allFull  = false;
      if (!empty) allEmpty = false;
      if (!full && !empty) anyPartial = true;
    });
    cell.status = allFull ? 'green' : (allEmpty ? 'gray' : 'amber');
  });

  // Pending list (all-time, FIFO oldest first). We emit one entry per
  // "missing" doc — so a (date, contract, borough) where both PL and
  // SI are unfilled creates two entries. Frontend caps the visible
  // window at 10 with scroll for the rest.
  const pending = [];
  const nowMs = new Date().getTime();
  const pushPending = (g, kind, anchor, doc_id, missing) => {
    const ageDays = Math.max(0, Math.round((nowMs - new Date(anchor + 'T12:00:00').getTime()) / 86400000));
    pending.push({
      kind,
      anchor,
      contractor:   g.contractor,
      crew_chief:   g.crew_chief || '',
      contract_num: g.contract_num,
      borough:      g.borough,
      missing,
      doc_id,
      wo_ids:       Array.from(g.wo_ids).sort(),
      age_days:     ageDays,
    });
  };
  // Day-kind pending. SI is per (contract, borough, crew_chief); PL is
  // per (contractor, crew_chief). Multi-crew shifts produce one SI and
  // one PL pending row per crew so Doc Status surfaces both.
  Object.values(contractorDays).forEach(cd => {
    // SI per (contract, borough) within this crew
    Object.values(cd.contracts).forEach(c => {
      const siId = _docLifecycleId_('Sign-In', cd.date, c.contract_num, c.borough, cd.crew_chief);
      const si = logById[siId];
      if (!si || !si.done) {
        pushPending({
          contractor:   cd.contractor,
          crew_chief:   cd.crew_chief,
          contract_num: c.contract_num,
          borough:      c.borough,
          wo_ids:       c.wo_ids,
        }, 'day', cd.date, siId, ['SI Done']);
      }
    });
    // PL per (contractor, crew) — only for contractors in CONFIG.PRODUCTION_LOG_CONTRACTORS.
    if (PL_ENABLED.has(cd.contractor)) {
      const plId = _plDocId_(cd.date, cd.contractor, cd.crew_chief);
      const pl = logById[plId];
      const plDone = pl && pl.done;
      const plSent = pl && pl.sent;
      const cdProxy = {
        contractor:   cd.contractor,
        crew_chief:   cd.crew_chief,
        contract_num: '',                  // PL spans contracts — none singular
        borough:      '',
        wo_ids:       cd.wo_ids,
      };
      if (!plDone) pushPending(cdProxy, 'day', cd.date, plId, ['PL Done']);
      if (plDone && !plSent) pushPending(cdProxy, 'day', cd.date, plId, ['PL Sent']);
    }
  });
  // Week-kind pending: CP Done, CP Sent
  Object.values(weekTuples).forEach(w => {
    const cpId = _docLifecycleId_('Certified Payroll', w.week_start, w.contract_num, w.borough);
    const cp = logById[cpId];
    const cpDone = cp && cp.done;
    const cpSent = cp && cp.sent;
    if (!cpDone) pushPending(w, 'week', w.week_start, cpId, ['CP Done']);
    if (cpDone && !cpSent) pushPending(w, 'week', w.week_start, cpId, ['CP Sent']);
  });
  // Within same anchor, order: SI Done → PL Done → PL Sent → CP Done → CP Sent.
  // Then break further ties on (contract_num, borough).
  const PENDING_RANK = {
    'SI Done': 0, 'PL Done': 1, 'PL Sent': 2, 'CP Done': 3, 'CP Sent': 4,
  };
  pending.sort((a, b) => {
    if (a.anchor !== b.anchor) return a.anchor < b.anchor ? -1 : 1;
    const ra = PENDING_RANK[a.missing[0]] != null ? PENDING_RANK[a.missing[0]] : 99;
    const rb = PENDING_RANK[b.missing[0]] != null ? PENDING_RANK[b.missing[0]] : 99;
    if (ra !== rb) return ra - rb;
    const ca = String(a.contract_num) + '|' + String(a.borough);
    const cb = String(b.contract_num) + '|' + String(b.borough);
    return ca.localeCompare(cb);
  });

  return {
    month: monthIso,
    days:   Object.values(dayCellByDate).sort((a, b) => a.date < b.date ? -1 : 1),
    weeks:  Object.values(weekCellByStart).sort((a, b) => a.week_start < b.week_start ? -1 : 1),
    pending,
  };
}


// ── action: upload_photo ──────────────────────────────────────

/**
 * Uploads a work-order site photo to Drive.
 * Saves into: Archive / Contractor / ContractNum-Borough / WO#-Location / Photos /
 *
 * body.data:
 *   wo_id     — Work Order # (used to find the archive folder)
 *   filename  — original filename (e.g. "IMG_1234.jpg")
 *   mime_type — MIME type (e.g. "image/jpeg")
 *   data      — base64-encoded file bytes
 */
function handleUploadPhoto_(body) {
  // Express proxy wraps the payload under body.data
  const t0 = Date.now();
  const d = body.data || {};
  const { wo_id, filename, mime_type, data } = d;
  if (!wo_id || !filename || !data) {
    return jsonResponse_({ error: 'Missing required fields: wo_id, filename, data' }, 400);
  }

  const { folder: photosFolder, error: folderErr } = resolveWOSubfolder_(wo_id, 'Photos');
  const tFolder = Date.now();
  if (!photosFolder) {
    _logAutomation_('Photo Upload', 'WO archive folder unreachable', wo_id,
      `filename=${filename} | error=${folderErr}`, 'Error', 'Yes');
    return jsonResponse_({
      error: 'Could not locate or create WO archive folder: ' + folderErr
    }, 500);
  }

  const bytes = Utilities.base64Decode(data);
  const tDecode = Date.now();
  const blob  = Utilities.newBlob(bytes, mime_type || 'image/jpeg', filename);
  const file  = photosFolder.createFile(blob);
  const tCreate = Date.now();

  Logger.log('📸 Photo uploaded for WO ' + wo_id + ': ' + filename +
             ' | folder=' + (tFolder - t0) + 'ms decode=' + (tDecode - tFolder) + 'ms' +
             ' create=' + (tCreate - tDecode) + 'ms total=' + (tCreate - t0) + 'ms' +
             ' size=' + Math.round(bytes.length / 1024) + 'KB');
  return jsonResponse_({
    success: true,
    file_id: file.getId(),
    file_url: file.getUrl(),
    _timing: { folder_ms: tFolder - t0, decode_ms: tDecode - tFolder,
               create_ms: tCreate - tDecode, total_ms: tCreate - t0 },
  });
}


// ── action: list_wo_photos ────────────────────────────────────

/**
 * Returns every photo currently in the WO's Photos/ Drive folder so the
 * Field Report page can render a "previously taken" gallery alongside
 * the live picker. Thumbnail bytes are inlined as base64 so the client
 * doesn't have to chase per-file URLs (Drive thumbnails behind auth are
 * a pain for an unauthenticated webapp tab).
 *
 * body.data: { wo_id }
 * response: { photos: [{ file_id, name, url, thumbnail_b64, mime, created_at }] }
 */
function handleListWOPhotos_(body) {
  const d = body.data || {};
  const woId = String(d.wo_id || '').trim();
  if (!woId) return jsonResponse_({ error: 'Missing wo_id' }, 400);

  const { folder, error } = resolveWOSubfolder_(woId, 'Photos');
  if (!folder) {
    // Soft-fail: a WO that has never had a photo uploaded simply has no
    // folder. Return an empty list so the UI renders a clean empty state.
    Logger.log('list_wo_photos: no Photos folder for ' + woId + ' (' + error + ')');
    return jsonResponse_({ photos: [] });
  }

  // First pass: collect image-file refs + created dates only — cheap, no
  // thumbnail fetch. We sort newest-first and then pull thumbnails for at
  // most MAX_THUMBS files. getThumbnail() is a per-file Drive round-trip,
  // so a folder with dozens of photos would otherwise serialise dozens of
  // them (slow gallery load, and a pathological folder could brush the
  // 6-min execution limit). 120 covers any realistic WO.
  const MAX_THUMBS = 120;
  const refs = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const mime = f.getMimeType() || '';
    if (mime.indexOf('image/') !== 0) continue;
    refs.push({ file: f, mime: mime, created: f.getDateCreated() });
  }
  refs.sort((a, b) => b.created - a.created);   // newest first (Date math → ms)
  if (refs.length > MAX_THUMBS) {
    Logger.log('list_wo_photos: ' + woId + ' has ' + refs.length +
               ' photos; returning newest ' + MAX_THUMBS);
  }

  const out = refs.slice(0, MAX_THUMBS).map(r => {
    const f = r.file;
    let thumb = '';
    try {
      const blob = f.getThumbnail();
      if (blob) thumb = Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      // Drive sometimes refuses thumbnails for very fresh uploads — UI
      // falls back to the file URL, no need to fail the whole listing.
    }
    return {
      file_id:       f.getId(),
      name:          f.getName(),
      url:           f.getUrl(),
      thumbnail_b64: thumb,
      mime:          r.mime,
      created_at:    r.created.toISOString(),
    };
  });
  return jsonResponse_({ photos: out });
}


// ── action: get_wo_photo_content ──────────────────────────────

/**
 * Returns the raw bytes of a Drive photo so the webapp lightbox can
 * show a full-size view. Historic photos only come back from
 * list_wo_photos as 220x220 thumbnails (free Drive thumbnail) — that
 * is unreadable for "what did I take a picture of" review.
 *
 * body.data: { file_id }
 * response: { data: <base64>, mime } or { error }
 */
function handleGetWOPhotoContent_(body) {
  const d = body.data || {};
  const fileId = String(d.file_id || '').trim();
  if (!fileId) return jsonResponse_({ error: 'Missing file_id' }, 400);
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return jsonResponse_({
      data: Utilities.base64Encode(blob.getBytes()),
      mime: file.getMimeType() || 'image/jpeg',
    });
  } catch (err) {
    return jsonResponse_({ error: String(err && err.message || err) }, 500);
  }
}


// ── action: reverse_geocode ───────────────────────────────────

/**
 * Reverse-geocodes a lat/lng into a structured address using Apps
 * Script's built-in Maps service (no API key needed). Feeds the photo
 * watermark on the Field Report capture path.
 *
 * body.data: { lat, lng }
 * response: { address, city, state, zip, country } or { error }
 */
function handleReverseGeocode_(body) {
  const d = body.data || {};
  const lat = Number(d.lat);
  const lng = Number(d.lng);
  if (!isFinite(lat) || !isFinite(lng)) {
    return jsonResponse_({ error: 'Missing or invalid lat/lng' }, 400);
  }
  try {
    const res = Maps.newGeocoder().reverseGeocode(lat, lng);
    const first = res && res.results && res.results[0];
    if (!first) return jsonResponse_({ error: 'No address found' });
    const out = { address: '', city: '', state: '', zip: '', country: '' };
    // Build street address from the parts we have so blank slots don't
    // leave stray spaces (e.g. "  South Ave" when street_number is missing).
    let streetNum = '', route = '';
    (first.address_components || []).forEach(c => {
      const types = c.types || [];
      if (types.indexOf('street_number') !== -1) streetNum = c.long_name;
      else if (types.indexOf('route') !== -1) route = c.long_name;
      else if (types.indexOf('locality') !== -1) out.city = c.long_name;
      else if (!out.city && types.indexOf('sublocality') !== -1) out.city = c.long_name;
      else if (types.indexOf('administrative_area_level_1') !== -1) out.state = c.short_name;
      else if (types.indexOf('postal_code') !== -1) out.zip = c.long_name;
      else if (types.indexOf('country') !== -1) out.country = c.long_name;
    });
    out.address = [streetNum, route].filter(Boolean).join(' ').trim();
    if (!out.address && first.formatted_address) {
      out.address = String(first.formatted_address).split(',')[0].trim();
    }
    return jsonResponse_(out);
  } catch (err) {
    Logger.log('reverse_geocode error: ' + (err && err.message || err));
    return jsonResponse_({ error: String(err && err.message || err) });
  }
}


// ── action: upload_signature ──────────────────────────────────

/**
 * Uploads a crew member's digital signature (PNG) to Drive.
 * Saves into: Archive / ... / WO#-Location / Signatures /
 *
 * body.data:
 *   wo_id        — Work Order #
 *   crew_name    — employee name (for filename)
 *   signature    — "time_in" or "time_out"
 *   work_date    — date string (YYYY-MM-DD)
 *   data         — base64-encoded PNG
 */
function handleUploadSignature_(body) {
  // Express proxy wraps the payload under body.data
  const d = body.data || {};
  const { wo_id, crew_name, signature, work_date, data } = d;
  if (!wo_id || !crew_name || !data) {
    return jsonResponse_({ error: 'Missing required fields: wo_id, crew_name, data' }, 400);
  }

  const { folder: sigsFolder, error: folderErr } = resolveWOSubfolder_(wo_id, 'Signatures');
  if (!sigsFolder) {
    _logAutomation_('Signature Upload', 'WO archive folder unreachable', wo_id,
      `crew=${crew_name} signature=${signature} | error=${folderErr}`, 'Error', 'Yes');
    return jsonResponse_({
      error: 'Could not locate or create WO archive folder: ' + folderErr
    }, 500);
  }

  // Filename: "2026-04-12_John Smith_time_in.png"
  const safeName = String(crew_name).replace(/[^a-zA-Z0-9 _-]/g, '');
  const filename  = (work_date || 'unknown') + '_' + safeName + '_' + (signature || 'sig') + '.png';

  const bytes = Utilities.base64Decode(data);
  const blob  = Utilities.newBlob(bytes, 'image/png', filename);
  const file  = sigsFolder.createFile(blob);

  Logger.log('✍️ Signature saved for ' + crew_name + ' (' + signature + ') on WO ' + wo_id);
  return jsonResponse_({ success: true, file_id: file.getId(), file_url: file.getUrl() });
}


/**
 * Helper: returns (creating if needed) a named subfolder inside the WO's archive folder.
 * Path: Archive / Contractor / ContractNum-Borough / WO#-Location / [subfolderName]
 */
/**
 * Resolve (and create if needed) Archive/…/WO-Location/<subfolderName>.
 * One attempt — callers should use resolveWOSubfolder_ to get a retry
 * wrapper + a readable error message when the whole thing fails.
 */
function getWOSubfolder_(wo_id, subfolderName) {
  const props     = PropertiesService.getScriptProperties();
  const archiveId = props.getProperty('ARCHIVE_ID');
  if (!archiveId) throw new Error('ARCHIVE_ID script property is not set');

  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();
  const woRow   = allRows.find(r => r[0] && String(r[0]) === String(wo_id));

  const contractor = woRow ? String(woRow[1] || 'Unknown') : 'Unknown';
  const contractNum = woRow ? String(woRow[2] || '').split('/')[0] : 'Unknown';
  const borough    = woRow ? String(woRow[3] || '') : '';
  const location   = woRow ? String(woRow[5] || wo_id) : wo_id;

  const archiveRoot      = DriveApp.getFolderById(archiveId);
  const contractorFolder = getOrCreateSubfolder_(archiveRoot, contractor);
  const contractFolder   = getOrCreateSubfolder_(contractorFolder,
                             contractNum + (borough ? ' - ' + getBoroughName_(borough) : ''));
  const woFolder         = getOrCreateSubfolder_(contractFolder, wo_id + ' - ' + location);

  return getOrCreateSubfolder_(woFolder, subfolderName);
}


/**
 * Retry wrapper around getWOSubfolder_ — Drive has transient hiccups
 * (especially under parallel photo uploads for the same WO), and the
 * previous swallow-and-return-null behaviour turned those into opaque
 * "Could not locate or create WO archive folder" errors in the UI with
 * nothing in the logs.
 *
 * Returns { folder, error } — exactly one of the two is set. On final
 * failure the error message is the last exception's message, so the
 * caller can surface it to the user + record it in Automation Log.
 */
function resolveWOSubfolder_(wo_id, subfolderName, maxAttempts) {
  const attempts = maxAttempts || 3;
  let lastErr = null;
  for (let i = 0; i < attempts; i++) {
    try {
      const folder = getWOSubfolder_(wo_id, subfolderName);
      if (folder) return { folder, error: null };
      lastErr = new Error('getWOSubfolder_ returned no folder');
    } catch (err) {
      lastErr = err;
      Logger.log(`⚠️ resolveWOSubfolder_ attempt ${i + 1}/${attempts} failed for ${wo_id}/${subfolderName}: ${err && err.message || err}`);
      // Short backoff — most Drive hiccups clear in well under a second.
      if (i < attempts - 1) Utilities.sleep(500 * (i + 1));
    }
  }
  return { folder: null, error: (lastErr && lastErr.message) ? lastErr.message : String(lastErr) };
}


// ═══════════════════════════════════════════════════════════════
// QUICKBOOKS ONLINE INVOICE INTEGRATION
// ═══════════════════════════════════════════════════════════════
//
// The Apps Script side of the QB invoice flow. Owns:
//   - aggregating per-WO revenue into the 5 pricing-group line items
//     (qty × rate = amount, matching the QB invoice template columns)
//   - persisting the rotating refresh token in Script Properties
//   - caching contractor-name → QB Customer ID lookups
//   - recording the QB invoice number + ID back onto the WO Tracker
//
// Webapp (Express server) owns the actual QB OAuth handshake and
// outbound API calls. See webapp/src/server/qb.js.
//
// WO Tracker invoice columns (1-indexed):
//   27 = Invoice #          (QB DocNumber)
//   28 = Invoice Date
//   29 = Invoice Amount
//   30 = Invoice Sent?      (legacy; kept at "No" — QB now tracks)
//   50 = QB Invoice ID      (Intuit's internal Id; for building view URL)

// Per-group equivalent qty: convert each item's raw qty to the unit
// the pricing group is denominated in. Returns null if the category
// doesn't belong to this group or the multiplier is missing.
function _qbEquivalentQty_(group, category, qty) {
  if (qty == null || isNaN(qty)) return null;
  if (group === 'line4') {
    const m = LINE_WIDTH_MULTIPLIER_[category];
    return m == null ? null : qty * m;
  }
  if (group === 'line12') {
    const m = LINE12_MULTIPLIER_[category];
    return m == null ? null : qty * m;
  }
  if (group === 'preformed') {
    const u = PREFORMED_UNIT_COUNT_[category];
    return u == null ? null : qty * u;
  }
  if (group === 'extruded') {
    const u = EXTRUDED_UNIT_COUNT_[category];
    return u == null ? null : qty * u;
  }
  if (group === 'color_surface') return qty;
  return null;
}

const _QB_GROUP_ORDER = ['line4', 'line12', 'preformed', 'extruded', 'color_surface'];
const _QB_GROUP_LABEL = Object.freeze({
  line4:         '4" Line group',
  line12:        'Crosswalk / Stop Line',
  preformed:     'Preformed L&S',
  extruded:      'Extruded L&S',
  color_surface: 'Color Surface',
});
const _QB_UNIT_LABEL = Object.freeze({
  line4:         'LF (4" equiv)',
  line12:        'LF (12" equiv)',
  preformed:     'Units',
  extruded:      'Units',
  color_surface: 'SF',
});

/**
 * Aggregate one WO's Completed marking items into the QB invoice
 * payload shape — one line per non-empty pricing group with
 * { qty, rate, amount, description }. Math reconciles penny-perfect to
 * priceMarkingItem_'s revenue numbers: per group,
 *   amount  = Σ priceMarkingItem_(item).revenue
 *   qty     = Σ item.qty × multiplier[item.category]
 *   rate    = amount / qty   (= contract rate when uniform, weighted
 *                              blend when items spanned a rate change)
 *
 * Returns:
 *   {
 *     wo_id, contractor, contract_num, borough, location,
 *     work_start, work_end,
 *     totals: { revenue, items },
 *     lines:  [{ group, label, qty, unit_label, rate, amount, description }],
 *     needs_pricing: [{ item_id, category, qty, unit, reason }]
 *   }
 *
 * Empty groups are omitted. If `needs_pricing` is non-empty the caller
 * should refuse to send the invoice until the underlying items are fixed.
 */
function aggregateRevenueByWoForQB_(ss, woId) {
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const miSheet = ss.getSheetByName('Marking Items');
  if (!woSheet || !miSheet) {
    throw new Error('Missing required sheets — Work Order Tracker or Marking Items');
  }

  // WO metadata
  const woRows = woSheet.getDataRange().getValues();
  const woRow  = woRows.slice(1).find(r => String(r[0] || '').trim() === String(woId).trim());
  if (!woRow) throw new Error('WO not found in tracker: ' + woId);

  const fmtDate = (v) => {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(v || '').trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : s;
  };

  const contractor      = String(woRow[1]  || '').trim();
  // Apply billing remap so sub-prime work on a contract they didn't
  // win is priced against — and posted to — the contract they DID win.
  // Pricing reads Contract Pricing keyed by (contractor, contract_num,
  // borough); the remapped tuple is the one that has rates for this
  // contractor.
  const _qbMapped = _billingRemap_(
    String(woRow[2] || '').split('/')[0].trim(),
    String(woRow[3] || '').trim(),
    contractor
  );
  const contractNum  = _qbMapped.contractNum;
  const borough      = _qbMapped.borough;
  const location     = String(woRow[5]  || '').trim();
  const workStartIso = fmtDate(woRow[17]);
  const workEndIso   = fmtDate(woRow[18]) || workStartIso || Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // Marking items walk — collect per-category sums. Each marking item
  // type becomes its own invoice line; multiple Marking Items rows of
  // the same category combine into one line (raw qty + revenue summed).
  const miData = miSheet.getDataRange().getValues();
  const rates  = _loadContractPricing_(ss);
  const woMeta = { contractor, contract_num: contractNum, borough };

  // Per-category accumulators: acc[category] = { group, raw_qty, unit, revenue }
  const acc = {};
  const needsPricing = [];
  let totalItems = 0;

  for (let i = 1; i < miData.length; i++) {
    const r = miData[i];
    if (String(r[1] || '').trim() !== woId) continue;
    if (String(r[12] || '').toLowerCase() !== 'completed') continue;
    const qty = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) continue;

    const item = {
      item_id:        String(r[0] || '').trim(),
      category:       String(r[4] || '').trim(),
      quantity:       qty,
      unit:           String(r[9] || '').trim(),
      date_completed: fmtDate(r[11]),
    };
    const group = PRICING_GROUP_BY_CATEGORY_[item.category] || 'unpriced';
    const priced = priceMarkingItem_(item, woMeta, rates);

    if (priced.reason || group === 'unpriced') {
      needsPricing.push({
        item_id:  item.item_id,
        category: item.category,
        qty:      item.quantity,
        unit:     item.unit,
        reason:   priced.reason || 'unpriced_category',
      });
      continue;
    }

    // _qbEquivalentQty_ as a "is this category priceable for the QB
    // line builder?" gate — returns null for categories without a
    // known multiplier (e.g. Custom Msg in extruded). Still send to
    // needs_pricing if so.
    const equivQty = _qbEquivalentQty_(group, item.category, qty);
    if (equivQty == null) {
      needsPricing.push({
        item_id:  item.item_id,
        category: item.category,
        qty:      item.quantity,
        unit:     item.unit,
        reason:   'no_unit_count',
      });
      continue;
    }

    if (!acc[item.category]) {
      acc[item.category] = { group, raw_qty: 0, unit: item.unit, revenue: 0 };
    }
    acc[item.category].raw_qty += qty;
    acc[item.category].revenue += priced.revenue;
    totalItems += 1;
  }

  // Build line items: one per category, ordered by _QB_GROUP_ORDER
  // (so all line4 rows precede line12, etc.) then alphabetic within
  // each group.
  const round2 = (n) => Math.round(n * 100) / 100;
  const round4 = (n) => Math.round(n * 10000) / 10000;

  const lines = [];
  let totalRevenue = 0;

  // Per-group unit-count tables for the L&S description parenthetical.
  const _LS_UNIT_TABLE = {
    preformed: PREFORMED_UNIT_COUNT_,
    extruded:  EXTRUDED_UNIT_COUNT_,
  };

  _QB_GROUP_ORDER.forEach(group => {
    const cats = Object.keys(acc).filter(c => acc[c].group === group).sort();
    cats.forEach(cat => {
      const bucket = acc[cat];
      if (bucket.raw_qty <= 0) return;
      // rate × qty = amount on every row. Rate = revenue / raw_qty,
      // which equals base × multiplier (line4 / line12 / L&S) or just
      // base (color_surface) under uniform contracts. Blends if the
      // contract rate changed mid-WO — same math priceMarkingItem_
      // already produces, just decomposed at category granularity.
      const rate = round4(bucket.revenue / bucket.raw_qty);
      const unitOut = bucket.unit
        || (group === 'color_surface' ? 'SF'
           : (group === 'line4' || group === 'line12' ? 'LF' : 'EA'));

      // Description: bare category name on lines / color_surface;
      // category + (unit_count Units) on L&S so the multiplier baked
      // into the rate is self-documenting on the invoice.
      let description = cat;
      const lsTable = _LS_UNIT_TABLE[group];
      if (lsTable) {
        const unitCount = lsTable[cat];
        if (unitCount != null) {
          description = `${cat} (${unitCount} Units)`;
        }
      }

      lines.push({
        category:    cat,
        group,
        label:       _QB_GROUP_LABEL[group],
        qty:         round2(bucket.raw_qty),
        unit_label:  unitOut,
        rate,
        amount:      round2(bucket.revenue),
        description,
      });
      totalRevenue += bucket.revenue;
    });
  });

  return {
    wo_id:         woId,
    contractor,
    contract_num:  contractNum,
    borough,
    location,
    work_start:    workStartIso,
    work_end:      workEndIso,
    totals:        { revenue: round2(totalRevenue), items: totalItems },
    lines,
    needs_pricing: needsPricing,
  };
}


/**
 * Record a successfully-posted QB invoice on the WO Tracker.
 *   col 27 = Invoice # (QB DocNumber)
 *   col 28 = Invoice Date (today)
 *   col 29 = Invoice Amount
 *   col 50 = QB Invoice ID (Intuit internal id, for view URL)
 *
 * Refuses to overwrite an existing Invoice # — the webapp's pre-flight
 * "already invoiced?" check should have caught this earlier, but the
 * server-side guard prevents a race from silently double-billing.
 */
function recordQbInvoice_(ss, woId, data) {
  const docNumber = String(data.doc_number || '').trim();
  const qbId      = String(data.qb_invoice_id || '').trim();
  const amount    = Number(data.amount);
  if (!docNumber || !qbId || isNaN(amount)) {
    throw new Error('recordQbInvoice_: doc_number, qb_invoice_id, and amount are required');
  }

  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) throw new Error('Work Order Tracker not found');

  const data2d  = woSheet.getDataRange().getValues();
  const rowIdx0 = data2d.slice(1).findIndex(r => String(r[0] || '').trim() === String(woId).trim());
  if (rowIdx0 === -1) throw new Error('WO not found in tracker: ' + woId);
  const rowNum = rowIdx0 + 2;  // +2 because: slice(1) drops header + sheets are 1-indexed

  const existing = String(data2d[rowIdx0 + 1][26] || '').trim();
  if (existing) {
    throw new Error(`WO ${woId} already has Invoice # ${existing} — refusing to overwrite`);
  }

  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  woSheet.getRange(rowNum, 27).setValue(docNumber);   // Invoice #
  woSheet.getRange(rowNum, 28).setValue(today);       // Invoice Date
  woSheet.getRange(rowNum, 29).setValue(amount);      // Invoice Amount
  woSheet.getRange(rowNum, 45).setValue('Yes');       // Invoice Done? (1-idx col 45)
  woSheet.getRange(rowNum, 50).setValue(qbId);        // QB Invoice ID

  Logger.log(`✅ Recorded QB invoice ${docNumber} (id=${qbId}, $${amount}) on WO ${woId}`);
  return { wo_id: woId, doc_number: docNumber, qb_invoice_id: qbId, amount, invoice_date: today };
}


/**
 * One-shot bootstrap: add the "QB Invoice ID" column at WO Tracker
 * col 50. Idempotent. Run from the Apps Script editor once when first
 * deploying the QB integration. Avoids re-running setupAutomation
 * (which has many side effects).
 */
function setupQBInvoiceCol() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) throw new Error('Work Order Tracker not found');

  const COL = 50;
  const HEADER = 'QB Invoice ID';
  const current = String(woSheet.getRange(1, COL).getValue() || '').trim();
  if (current === HEADER) {
    Logger.log('↻ Col 50 already set to "QB Invoice ID"');
    return;
  }
  if (current) {
    throw new Error(`Col 50 already has header "${current}" — refusing to overwrite. Migrate manually first.`);
  }
  woSheet.getRange(1, COL).setValue(HEADER).setFontWeight('bold');
  Logger.log('✅ Added "QB Invoice ID" header at WO Tracker col 50');
}


/**
 * One-shot bootstrap: pre-populate the QB customer cache so the
 * webapp doesn't have to query QB by DisplayName (which often drifts
 * from our shorter WO Tracker contractor strings — e.g. "Metro Express"
 * in the Sheet vs. "Metro Express Services" in QB).
 *
 * Edit the mapping below with your real WO-Tracker-name → QB-Customer-ID
 * pairs, then run this function once from the Apps Script editor.
 * Re-running overwrites previous entries — safe and idempotent.
 *
 * Keys are case-insensitive and whitespace-normalized at lookup time
 * (so "Metro Express", "metro express", "Metro  Express" all match).
 */
function setupQbCustomerCache() {
  const MAPPING = {
    // WO Tracker contractor string  →  QB Customer ID
    'Metro Express': '10',
    'Denville':      '217',
    'Delan':         '219',
    // Add more contractors here as customers are added in QB:
  };

  const cache = {};
  Object.entries(MAPPING).forEach(([name, id]) => {
    const key = String(name || '').trim().replace(/\s+/g, ' ').toLowerCase();
    cache[key] = String(id);
  });

  PropertiesService.getScriptProperties().setProperty(
    'QB_CUSTOMER_ID_CACHE', JSON.stringify(cache)
  );
  Logger.log('✅ QB customer cache seeded with: ' +
             Object.keys(cache).join(', ') +
             ' (' + Object.keys(cache).length + ' entries)');
}


// ── QB token storage (rotating refresh token) ─────────────────────
// QBO rotates the refresh token every 24–26 hours; the webapp must
// write the rotated value back here immediately on every refresh.
// Failing to do so means the next refresh will fail with invalid_grant.

function _getQbRefreshToken_() {
  return PropertiesService.getScriptProperties().getProperty('QB_REFRESH_TOKEN') || '';
}
function _setQbRefreshToken_(token) {
  if (!token) throw new Error('Empty refresh token — refusing to clobber stored token');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('QB_REFRESH_TOKEN',    String(token));
  props.setProperty('QB_LAST_REFRESH_AT', new Date().toISOString());
}


// ── QB customer-name cache ────────────────────────────────────────
// Map of "Display Name" → QB Customer ID, persisted as a JSON blob in
// Script Properties. Avoids redundant QB customer-query API calls.

function _normalizeCustomerName_(name) {
  return String(name || '').trim().replace(/\s+/g, ' ').toLowerCase();
}
function _getQbCustomerCache_() {
  const raw = PropertiesService.getScriptProperties().getProperty('QB_CUSTOMER_ID_CACHE') || '{}';
  try {
    const parsed = JSON.parse(raw);
    return (parsed && typeof parsed === 'object') ? parsed : {};
  } catch (_) {
    return {};
  }
}
function _setQbCustomerCache_(cache) {
  PropertiesService.getScriptProperties().setProperty('QB_CUSTOMER_ID_CACHE', JSON.stringify(cache || {}));
}


// ── doPost handlers ──────────────────────────────────────────────

function handleGetQbInvoicePayload_(body) {
  const d = body.data || {};
  const woId = String(d.wo_id || '').trim();
  if (!woId) return jsonResponse_({ error: 'wo_id required' }, 400);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Pre-flight: already invoiced?
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const rows    = woSheet.getDataRange().getValues();
  const woRow   = rows.slice(1).find(r => String(r[0] || '').trim() === woId);
  if (!woRow) return jsonResponse_({ error: 'WO not found: ' + woId }, 404);
  const existingDoc = String(woRow[26] || '').trim();
  const existingQbId= String(woRow[49] || '').trim();
  if (existingDoc) {
    return jsonResponse_({
      already_invoiced: true,
      doc_number:       existingDoc,
      qb_invoice_id:    existingQbId,
      amount:           woRow[28] || '',
      invoice_date:     woRow[27] || '',
    });
  }

  const payload = aggregateRevenueByWoForQB_(ss, woId);
  return jsonResponse_(payload);
}

function handleRecordQbInvoice_(body) {
  const d = body.data || {};
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const result = recordQbInvoice_(ss, d.wo_id, d);
  return jsonResponse_(result);
}

/**
 * Clear the recorded QB invoice fields on a WO Tracker row so the
 * next invoice generation creates a fresh one. Called by the webapp
 * after detecting that the QB-side invoice was deleted out from under
 * us (auto-heal path).
 *
 * Wipes col 27 (Invoice #), 28 (Invoice Date), 29 (Invoice Amount),
 * and 50 (QB Invoice ID). Leaves col 30 (Invoice Sent?) alone — it
 * was a legacy flag and isn't part of the QB integration's state.
 */
function handleClearQbInvoice_(body) {
  const d = body.data || {};
  const woId = String(d.wo_id || '').trim();
  if (!woId) return jsonResponse_({ error: 'wo_id required' }, 400);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  if (!woSheet) return jsonResponse_({ error: 'Work Order Tracker not found' }, 404);
  const rows = woSheet.getDataRange().getValues();
  const rowIdx0 = rows.slice(1).findIndex(r => String(r[0] || '').trim() === woId);
  if (rowIdx0 === -1) return jsonResponse_({ error: 'WO not found: ' + woId }, 404);
  const rowNum = rowIdx0 + 2;
  // Clear cols 27, 28, 29 in one range; col 50 separately. Reset col 45
  // (Invoice Done?) to "No" so the INV doc chip on the dashboard reverts
  // to grey/not-done in lockstep.
  woSheet.getRange(rowNum, 27, 1, 3).clearContent();
  woSheet.getRange(rowNum, 45).setValue('No');
  woSheet.getRange(rowNum, 50).clearContent();
  Logger.log(`✅ Cleared QB invoice fields on WO ${woId} (row ${rowNum})`);
  return jsonResponse_({ ok: true, wo_id: woId });
}

function handleGetQbRefreshToken_() {
  return jsonResponse_({ token: _getQbRefreshToken_() });
}

function handleSetQbRefreshToken_(body) {
  const d = body.data || {};
  _setQbRefreshToken_(d.token);
  return jsonResponse_({ ok: true });
}

function handleGetQbCustomerId_(body) {
  const d = body.data || {};
  const name = _normalizeCustomerName_(d.name);
  if (!name) return jsonResponse_({ error: 'name required' }, 400);
  const cache = _getQbCustomerCache_();
  return jsonResponse_({ id: cache[name] || null });
}

function handleSetQbCustomerId_(body) {
  const d = body.data || {};
  const name = _normalizeCustomerName_(d.name);
  const id   = String(d.id || '').trim();
  if (!name || !id) return jsonResponse_({ error: 'name and id required' }, 400);
  const cache = _getQbCustomerCache_();
  cache[name] = id;
  _setQbCustomerCache_(cache);
  return jsonResponse_({ ok: true });
}
