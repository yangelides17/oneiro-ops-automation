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
  TIMEZONE: 'America/New_York'
};


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
  
  // Reports subfolders
  reports.createFolder('Weekly Payroll Reports');
  reports.createFolder('Monthly Workforce Utilization');
  
  // Save folder IDs
  props.setProperty('ROOT_FOLDER_ID', root.getId());
  props.setProperty('SCAN_INBOX_ID', scanInbox.getId());
  props.setProperty('NEEDS_REVIEW_ID', needsReview.getId());
  props.setProperty('APPROVED_SENT_ID', approvedSent.getId());
  props.setProperty('ARCHIVE_ID', archive.getId());
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
    'L/R Arrow', 'Straight Arrow', 'Combination Arrow',
    // Page 2 — miscellaneous
    'Speed Hump Markings', 'Shark Teeth 12x18', 'Shark Teeth 24x36',
    // Page 2 — bike lane
    'Bike Lane Arrow', 'Bike Lane Symbol', 'Bike Lane Green Bar',
    // MMA
    'Bike Lane', 'Pedestrian Space', 'Bus Lane', 'Ped Stop',
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
 * @param {string} dateStr - Date in MM/DD/YYYY format (defaults to today)
 */
function generateDailyDocuments(dateStr) {
  const targetDate = dateStr ? new Date(dateStr) : new Date();
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
    return;
  }
  
  // Group entries by Work Order
  const byWorkOrder = {};
  todaysEntries.forEach(row => {
    const woId = row[1]; // Work Order # column
    if (!byWorkOrder[woId]) byWorkOrder[woId] = [];
    byWorkOrder[woId].push(row);
  });
  
  // Get Work Order details from tracker
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const woData = woSheet.getDataRange().getValues();
  const woHeaders = woData[0];
  
  Logger.log(`📋 Processing ${Object.keys(byWorkOrder).length} work orders for ${targetDate.toDateString()}`);
  
  // Generate Production Log (Metro Express format)
  generateProductionLog_(targetDate, todaysEntries, byWorkOrder, woData, ss);
  
  // Generate Field Reports for completed WOs
  Object.entries(byWorkOrder).forEach(([woId, entries]) => {
    const isComplete = entries.some(e => String(e[14]).toLowerCase() === 'yes');
    if (isComplete) {
      generateFieldReport_(woId, entries, woData, ss);
      generateInvoice_(woId, entries, woData, ss);
    }
  });
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
  'Speed Hump Markings': 'Speed Hump Marking',
  'Shark Teeth 24x36':   'Sharks Teeth 24" 36"',
  'Bike Lane Arrow':     'Bicycle Lane Arrow',
  'Bike Lane Symbol':    'Bicycle Lane Symbol',
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
function aggregateMarkingItemsForPL_(ss, woId) {
  const out = { markings: {}, sqft: '', paint: '' };
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return out;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return out;

  // Col indices (15-col Marking Items schema):
  //   1 WO#, 4 Marking Type, 5 Intersection, 6 Direction,
  //   8 Quantity, 9 Unit, 10 Color/Material, 12 Status
  const woItems = data.slice(1).filter(r =>
    String(r[1] || '').trim() === woId &&
    String(r[12] || '').toLowerCase() === 'completed'
  );
  if (woItems.length === 0) return out;

  let sfSum = 0;
  const colorsSet = {};
  let crosswalkSum = 0;

  woItems.forEach(r => {
    const category = String(r[4] || '').trim();
    const qty      = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;
    const unit     = String(r[9] || '').toUpperCase();

    // MMA SF → Color Surface Treatment 1/2, not the grid
    if (unit === 'SF' &&
        (category === 'Bike Lane' || category === 'Bus Lane' || category === 'Pedestrian Space')) {
      sfSum += qty;
      const color = String(r[10] || '').trim();
      if (color && color.toLowerCase() !== 'n/a') colorsSet[color] = true;
      return;
    }

    // HVX Crosswalk + Stop Line → combined LF cell
    if (category === 'HVX Crosswalk' || category === 'Stop Line') {
      crosswalkSum += qty;
      return;
    }

    // Standard rename; unmapped categories drop silently
    const plLabel = PL_CATEGORY_MAP_[category];
    if (!plLabel) return;
    out.markings[plLabel] = (out.markings[plLabel] || 0) + qty;
  });

  if (crosswalkSum > 0) {
    out.markings['CrossWalks/Stop Lines'] = crosswalkSum;
  }
  if (sfSum > 0) {
    out.sqft  = sfSum;
    out.paint = Object.keys(colorsSet).sort().join(', ');
  }
  return out;
}


/**
 * Generate Metro Thermoplastic Production Daily Log
 *
 * Filter rule: only include WOs where the WO Tracker has
 * `Status = 'Completed'` AND `Work End Date = targetDate`. A multi-day
 * WO therefore appears only on the PL for the day the work was
 * finished, and the quantities shown are the WO's full cumulative
 * Marking Items totals (acknowledged skew for now; ops manager to
 * weigh in).
 */
function generateProductionLog_(targetDate, allEntries, byWorkOrder, woData, ss) {
  const props = PropertiesService.getScriptProperties();
  const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
  const targetDayStr = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // ── Filter WOs to those COMPLETED on targetDate ───────────────
  // WO Tracker col 15 = Status, col 18 = Work End Date.
  const completedByWO = {};
  Object.entries(byWorkOrder).forEach(([woId, entries]) => {
    const woRow = woData.find(r => String(r[0]) === String(woId));
    if (!woRow) return;
    const status = String(woRow[15] || '').trim().toLowerCase();
    if (status !== 'completed') return;

    // Work End Date can come back as a Date object or a string
    const workEndRaw = woRow[18];
    let workEndStr = '';
    if (workEndRaw instanceof Date && !isNaN(workEndRaw.getTime())) {
      workEndStr = Utilities.formatDate(workEndRaw, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    } else if (workEndRaw) {
      const d = new Date(workEndRaw);
      if (!isNaN(d.getTime())) {
        workEndStr = Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      }
    }
    if (workEndStr !== targetDayStr) return;

    completedByWO[woId] = entries;
  });

  if (Object.keys(completedByWO).length === 0) {
    Logger.log('No WOs completed on ' + targetDayStr + ' — skipping Production Log.');
    return null;
  }

  // Only the crew that worked on completed WOs shows up on this PL
  const relevantEntries = [];
  Object.values(completedByWO).forEach(rows => relevantEntries.push(...rows));

  // Get unique employees — track earliest time-in and latest time-out across all completed WOs for the day
  const employees = {};
  relevantEntries.forEach(row => {
    const name = row[6];
    const timeInMins  = parseTimeToMinutes_(row[8]);
    const timeOutMins = parseTimeToMinutes_(row[9]);
    if (!employees[name]) {
      employees[name] = {
        timeIn: formatTime_(row[8]), timeOut: formatTime_(row[9]),
        timeInMins, timeOutMins,
        classification: row[7]
      };
    } else {
      // Correct numeric comparison — not string comparison
      if (timeInMins  < employees[name].timeInMins)  { employees[name].timeIn  = formatTime_(row[8]); employees[name].timeInMins  = timeInMins;  }
      if (timeOutMins > employees[name].timeOutMins) { employees[name].timeOut = formatTime_(row[9]); employees[name].timeOutMins = timeOutMins; }
    }
  });

  const dateFormatted = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'MM/dd/yyyy');

  // ── Build a plain-text summary for admin review ──────────────
  let logContent = `METRO THERMOPLASTIC PRODUCTION DAILY LOG\n`;
  logContent += `Date: ${dateFormatted}\n\n`;
  Object.entries(employees).forEach(([name, info]) => {
    const role = info.classification === 'LP' ? 'Crew Chief' : 'Individual';
    logContent += `${role}: ${name} | In: ${info.timeIn} | Out: ${info.timeOut}\n`;
  });
  logContent += `\n${'─'.repeat(80)}\n\n`;
  logContent += `WORK ORDERS COMPLETED ${dateFormatted}:\n\n`;
  logContent += `${'Borough'.padEnd(8)} ${'WO #'.padEnd(12)} ${'Location'.padEnd(25)} ${'SQFT'.padEnd(10)} ${'Paint/Material'}\n`;
  logContent += `${'─'.repeat(80)}\n`;

  // ── Build the per-WO payload for the filler ──────────────────
  const workOrdersJson = Object.entries(completedByWO).map(([woId, entries]) => {
    const woRow   = woData.find(r => String(r[0]) === String(woId));
    const borough = woRow ? String(woRow[3]).toUpperCase() : '';
    const location = woRow ? String(woRow[5]).toUpperCase() : String(entries[0][5] || '').toUpperCase();
    const agg     = aggregateMarkingItemsForPL_(ss, woId);

    const sqftStr  = agg.sqft  !== '' ? String(agg.sqft)  : '';
    const paintStr = agg.paint !== '' ? String(agg.paint) : '';

    logContent += `${String(borough).padEnd(8)} ${String(woId).padEnd(12)} ${String(location).padEnd(25)} ${sqftStr.padEnd(10)} ${paintStr}\n`;

    return {
      wo_number:    String(woId),
      borough:      borough,
      location:     location,
      sqft:         sqftStr,
      paint:        paintStr,
      complete:     'Y',        // filter guarantees this
      layout_yn:    '',
      layout_hours: '',
      markings:     agg.markings,
    };
  });

  // Production Logs folder — JSON gets written below; the Python worker
  // produces the filled PDF.
  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  const subFolder    = getOrCreateSubfolder_(reviewFolder, 'Production Logs');

  // ── Plain-text skeleton (disabled; kept in case admin wants it back) ──
  // const fileName = `Production_Log_${targetDayStr}.txt`;
  // const file     = subFolder.createFile(fileName, logContent, MimeType.PLAIN_TEXT);

  // ── JSON export for the Python filler ──────────────────────
  const sortedNames = Object.keys(employees);
  const crewChiefName = sortedNames.find(n => employees[n].classification === 'LP') || sortedNames[0];
  const crewMemberNames = sortedNames.filter(n => n !== crewChiefName);

  const logJson = {
    _type:             'production_log',
    date:              dateFormatted,
    crew_number:       '',
    truck_number:      '',
    inspector_present: '',
    gas_tank_refilled: '',
    materials: {
      thermo_white_bags:  '',
      thermo_yellow_bags: '',
      beads_bags:         '',
      paint_cans:         ''
    },
    crew_chief: crewChiefName ? {
      name:     crewChiefName,
      time_in:  employees[crewChiefName].timeIn,
      time_out: employees[crewChiefName].timeOut
    } : { name: '', time_in: '', time_out: '' },
    crew: crewMemberNames.map(n => ({
      name:     n,
      time_in:  employees[n].timeIn,
      time_out: employees[n].timeOut
    })),
    work_orders: workOrdersJson
  };

  const jsonFileName = `Production_Log_${targetDayStr}.json`;
  const jsonFile = subFolder.createFile(jsonFileName, JSON.stringify(logJson, null, 2), MimeType.PLAIN_TEXT);

  // Log it
  const logSheet = ss.getSheetByName('Automation Log');
  logSheet.appendRow([
    new Date(), 'Production Log Generator', 'Daily trigger',
    `${Object.keys(completedByWO).length} completed WO(s) on ${dateFormatted}`,
    jsonFileName, 'Generated',
    '', 'Yes — Review and send to Claudia'
  ]);

  Logger.log('✅ Production log JSON exported: ' + jsonFileName);
  return jsonFile;
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
    const woId = row[1];
    if (!byWorkOrder[woId]) byWorkOrder[woId] = [];
    byWorkOrder[woId].push(row);
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
  const contractNum = woRow[2];
  const borough = woRow[3];
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
 */
function processApprovedDocuments() {
  const props = PropertiesService.getScriptProperties();
  const approvedId = props.getProperty('APPROVED_SENT_ID');
  if (!approvedId) return;
  
  const approvedFolder = DriveApp.getFolderById(approvedId);
  const files = approvedFolder.getFiles();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Skip already-processed files
    if (fileName.startsWith('📨')) continue;
    
    // Determine document type and extract WO info from filename
    let docType = 'Unknown';
    let woId = '';
    
    if (fileName.includes('Production_Log')) {
      docType = 'Production Log';
    } else if (fileName.includes('Field_Report')) {
      docType = 'Field Report';
      const match = fileName.match(/PT[-_]?\d+/i);
      if (match) woId = match[0];
    } else if (fileName.includes('Invoice')) {
      docType = 'Invoice';
      const match = fileName.match(/PT[-_]?\d+/i);
      if (match) woId = match[0];
    } else if (fileName.includes('Certified_Payroll')) {
      docType = 'Certified Payroll';
    }
    
    // Look up who to send this to
    const recipients = getRecipientsForDoc_(docType, woId, ss);
    
    if (recipients.length > 0) {
      // Email the document
      const blob = file.getBlob();
      recipients.forEach(recipient => {
        MailApp.sendEmail({
          to: recipient.email,
          subject: `Oneiro Collection — ${docType}: ${woId || fileName}`,
          htmlBody: `
            <p>Hi ${recipient.name},</p>
            <p>Please find attached the ${docType.toLowerCase()} ${woId ? 'for ' + woId : ''} from Oneiro Collection LLC.</p>
            <p>Best regards,<br>Oneiro Operations</p>
          `,
          attachments: [blob]
        });
      });
      
      Logger.log(`📨 Sent ${docType} to ${recipients.map(r => r.email).join(', ')}`);
    }
    
    // Archive the file
    archiveDocument_(file, docType, woId, ss);

    // Build detailed log summary
    const recipientList = recipients.length > 0
      ? recipients.map(r => `${r.name} <${r.email}>`).join('; ')
      : 'No recipients (archive only)';
    const archiveNote = docType === 'Production Log' || docType === 'Certified Payroll'
      ? `Archived to master folder + duplicated into WO subfolder(s)`
      : woId
        ? `Archived to WO folder: ${woId}`
        : `Archived (doc type: ${docType})`;

    // Log to Automation Log tab before deleting
    const logSheet = ss.getSheetByName('Automation Log');
    logSheet.appendRow([
      new Date(), 'Approve & Send', 'File moved to Approved Docs folder',
      fileName,
      `Emailed to: ${recipientList} | ${archiveNote}`,
      'Completed', '', 'No'
    ]);

    // Delete from Approved Docs — archive is now the single source of truth
    file.setTrashed(true);
    Logger.log(`🗑️ Deleted from Approved Docs: ${fileName}`);
  }
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

/**
 * Archive a document using the correct folder structure:
 *
 *   Archive / [Contractor] / [ContractNum - Borough] /
 *     ├── PT-XXXXX - [Location] /   ← WO folder (Field Reports, Invoices filed directly here)
 *     │     └── Photos/             ← only subfolder inside a WO
 *     ├── Production Logs/          ← master copy; also duplicated into each WO folder
 *     └── Certified Payroll/        ← master copy; also duplicated into each WO folder
 */
function archiveDocument_(file, docType, woId, ss) {
  const props = PropertiesService.getScriptProperties();
  const archiveId = props.getProperty('ARCHIVE_ID');
  if (!archiveId) return;

  const archiveRoot = DriveApp.getFolderById(archiveId);
  const cleanName = file.getName().replace('📨 ', '');

  if (docType === 'Field Report' || docType === 'Invoice') {
    // Single doc — file directly inside the WO subfolder, no type subfolder
    const woFolder = getWOFolder_(archiveRoot, woId, ss);
    if (woFolder) {
      file.makeCopy(cleanName, woFolder);
      Logger.log(`📁 Archived ${docType}: ${cleanName} → WO folder ${woId}`);
    }

  } else if (docType === 'Production Log') {
    // Parse date from filename: Production_Log_YYYY-MM-DD.txt
    const dateMatch = cleanName.match(/(\d{4}-\d{2}-\d{2})/);
    if (!dateMatch) { Logger.log('❌ Could not parse date from Production Log filename'); return; }
    const logDate = new Date(dateMatch[1] + 'T12:00:00');

    // Group WOs worked that day by contractor/contract/borough
    const wosByContract = getWOsForDate_(logDate, ss);

    Object.entries(wosByContract).forEach(([key, wos]) => {
      const [contractor, contractNum, borough] = key.split('|');
      const contractFolder = getOrCreateSubfolder_(
        getOrCreateSubfolder_(archiveRoot, contractor),
        `${contractNum} - ${getBoroughName_(borough)}`
      );
      // Master copy at contract level
      file.makeCopy(cleanName, getOrCreateSubfolder_(contractFolder, 'Production Logs'));
      // Duplicate into each WO folder covered by this log
      wos.forEach(wo => {
        const woFolder = getOrCreateSubfolder_(contractFolder, `${wo.id} - ${wo.location}`);
        file.makeCopy(cleanName, woFolder);
      });
      Logger.log(`📁 Archived Production Log → ${contractor}/${contractNum} + ${wos.length} WO folder(s)`);
    });

  } else if (docType === 'Certified Payroll') {
    // Parse from filename: Certified_Payroll_[contractNum]_[borough]_YYYY-MM-DD.txt
    const match = cleanName.match(/Certified_Payroll_([^_]+)_([^_]+)_(\d{4}-\d{2}-\d{2})/);
    if (!match) { Logger.log('❌ Could not parse contract info from Certified Payroll filename'); return; }
    const [, contractNum, borough, weekStartStr] = match;
    const weekStart = new Date(weekStartStr + 'T12:00:00');

    // Look up contractor from Contract Lookup
    const clData = ss.getSheetByName('Contract Lookup').getDataRange().getValues();
    const clRow = clData.find(r => String(r[0]).includes(contractNum) && String(r[1]) === borough);
    const contractor = clRow ? String(clRow[5]).split(',')[0].trim() : 'General';

    const contractFolder = getOrCreateSubfolder_(
      getOrCreateSubfolder_(archiveRoot, contractor),
      `${contractNum} - ${getBoroughName_(borough)}`
    );
    // Master copy at contract level
    file.makeCopy(cleanName, getOrCreateSubfolder_(contractFolder, 'Certified Payroll'));
    // Duplicate into each WO folder worked during that payroll week for this contract
    const wos = getWOsForPayrollWeek_(contractNum, borough, weekStart, ss);
    wos.forEach(wo => {
      const woFolder = getOrCreateSubfolder_(contractFolder, `${wo.id} - ${wo.location}`);
      file.makeCopy(cleanName, woFolder);
    });
    Logger.log(`📁 Archived Certified Payroll → ${contractor}/${contractNum} + ${wos.length} WO folder(s)`);
  }
}

/** Get or create the WO-level subfolder: Archive/Contractor/Contract-Borough/WO#-Location */
function getWOFolder_(archiveRoot, woId, ss) {
  const woData = ss.getSheetByName('Work Order Tracker').getDataRange().getValues();
  const woRow = woData.find(r => r[0] === woId);
  if (!woRow) return null;
  const contractor  = woRow[1] || 'General';
  const contractNum = String(woRow[2]).split('/')[0];
  const borough     = woRow[3];
  const location    = woRow[5];
  const contractFolder = getOrCreateSubfolder_(
    getOrCreateSubfolder_(archiveRoot, contractor),
    `${contractNum} - ${getBoroughName_(borough)}`
  );
  return getOrCreateSubfolder_(contractFolder, `${woId} - ${location}`);
}

/** Return WOs worked on a given date, grouped by "contractor|contractNum|borough" */
function getWOsForDate_(date, ss) {
  const data = ss.getSheetByName('Daily Sign-In Data').getDataRange().getValues();
  const wosByContract = {};
  const seen = new Set();
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    if (new Date(row[0]).toDateString() !== date.toDateString()) return;
    const woId = row[1];
    if (seen.has(woId)) return;
    seen.add(woId);
    const key = `${row[2]}|${String(row[3]).split('/')[0]}|${row[4]}`;
    if (!wosByContract[key]) wosByContract[key] = [];
    wosByContract[key].push({ id: woId, location: row[5] });
  });
  return wosByContract;
}

/** Return unique WOs for a contract+borough during a payroll week */
function getWOsForPayrollWeek_(contractNum, borough, weekStart, ss) {
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  const data = ss.getSheetByName('Daily Sign-In Data').getDataRange().getValues();
  const wos = [];
  const seen = new Set();
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    const d = new Date(row[0]);
    if (d < weekStart || d > weekEnd) return;
    if (String(row[3]).split('/')[0] !== contractNum || row[4] !== borough) return;
    const woId = row[1];
    if (seen.has(woId)) return;
    seen.add(woId);
    wos.push({ id: woId, location: row[5] });
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
 * @param {string} weekStartStr - Monday date in MM/DD/YYYY format
 */
/**
 * YTD gross pay for an employee across ALL projects, up to and
 * including the payroll week end date. Applies the same day-of-week OT
 * rule the weekly payroll uses (Sat/Sun all OT; weekday >8 OT) so the
 * Total Gross Pay column matches the weekly column's math.
 *
 * signInData: full Daily Sign-In Data getValues() result (incl. header row).
 * empName:    employee name exactly as written in Sign-In Data col 6.
 * stRate, otRate: from Employee Registry for this employee.
 * weekEnd:    Date object marking end of the payroll week (YTD cutoff).
 */
function computeYtdGrossForEmployee_(signInData, empName, stRate, otRate, weekEnd) {
  const normName = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const target   = normName(empName);
  if (!target) return 0;

  const yearStart = new Date(weekEnd.getFullYear(), 0, 1, 0, 0, 0);

  // Aggregate total hours by date for this employee (handles the case
  // where two sign-ins on the same day get combined before the OT rule).
  const byDay = {};   // 'yyyy-MM-dd' → { hours, dow }
  signInData.slice(1).forEach(row => {
    if (!row[0]) return;
    const rowDate = new Date(row[0]);
    if (isNaN(rowDate.getTime())) return;
    if (rowDate < yearStart || rowDate > weekEnd) return;
    if (normName(row[6]) !== target) return;
    const h = Number(row[10]) || 0;
    if (h <= 0) return;
    const key = Utilities.formatDate(rowDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    if (!byDay[key]) byDay[key] = { hours: 0, dow: rowDate.getDay() };
    byDay[key].hours += h;
  });

  let totalST = 0;
  let totalOT = 0;
  Object.values(byDay).forEach(({ hours, dow }) => {
    if (dow === 0 || dow === 6) {
      totalOT += hours;
    } else if (hours <= 8) {
      totalST += hours;
    } else {
      totalST += 8;
      totalOT += hours - 8;
    }
  });
  return totalST * stRate + totalOT * otRate;
}


function generateCertifiedPayroll(weekStartStr) {
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
  
  // Full classification names — add new codes here as needed
  const CLASSIFICATION_NAMES = {
    'LP':   'Line Person',
    'SAT':  'Stripper Assistant',
    'OP':   'Operator',
    'LAB':  'Laborer',
    'FGL':  'Flagger',
    'SUP':  'Supervisor',
  };
  const classificationName = code => CLASSIFICATION_NAMES[String(code).trim().toUpperCase()] || String(code).trim();

  // For each contract group, generate certified payroll entries
  Object.entries(byContract).forEach(([key, entries]) => {
    const [contractNum, borough] = key.split('|');
    
    // Look up Contract ID from lookup table
    const clRow = clData.find(r => 
      String(r[0]).includes(contractNum) && String(r[1]) === borough
    );
    const contractId = clRow ? clRow[3] : '⚠️ MISSING — CHECK LOOKUP TABLE';
    const projectName = clRow ? clRow[4] : '';
    
    // Group by employee
    const byEmployee = {};
    entries.forEach(row => {
      const emp = row[6];
      if (!byEmployee[emp]) byEmployee[emp] = { days: {}, classification: row[7], totalHours: 0 };
      
      const dayOfWeek = new Date(row[0]).getDay(); // 0=Sun, 1=Mon...6=Sat
      const hours = Number(row[10]) || 0;
      byEmployee[emp].days[dayOfWeek] = (byEmployee[emp].days[dayOfWeek] || 0) + hours;
      byEmployee[emp].totalHours += hours;
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

    // Write to Certified Payroll Tracker
    Object.entries(byEmployee).forEach(([empName, info]) => {
      // Look up pay rates + address/SSN4 from Employee Registry by full name.
      const empRow = empByName[normName(empName)] || null;
      if (!empRow) {
        Logger.log(`⚠️ Certified Payroll: no Employee Registry row matches ${JSON.stringify(empName)} — rates/address/SSN will be blank.`);
      }
      const stRate    = empRow ? Number(empRow[6]) : 0;
      const otRate    = empRow ? Number(empRow[7]) : 0;
      const empFringe = empRow ? Number(empRow[9]) : 0;
      const erFringe  = empRow ? Number(empRow[8]) : 0;
      const empAddr   = empRow ? String(empRow[2] || '') : '';
      const empSsn4   = empRow ? String(empRow[3] || '') : '';

      // YTD gross across all projects this year. Used to fill the
      // "Total Gross Pay (All Work)" column on the certified payroll form.
      // Applies the same day-of-week OT rule (Sat/Sun all OT; weekday >8 OT).
      const ytdGross = computeYtdGrossForEmployee_(data, empName, stRate, otRate, weekEnd);

      // OT rules: ALL hours on Saturday (6) and Sunday (0) are OT.
      // On Mon–Fri, hours over 8 in a single day are OT.
      let totalST = 0;
      let totalOT = 0;
      const stHours = ['0','0','0','0','0','0','0'];
      const otHours = ['0','0','0','0','0','0','0'];

      Object.entries(info.days).forEach(([dayOfWeek, dayHours]) => {
        const dow = Number(dayOfWeek);
        const fi  = jsDayToFormDay[dow];
        if (dow === 0 || dow === 6) {
          totalOT += dayHours;
          otHours[fi] = String(dayHours);
        } else {
          if (dayHours <= 8) {
            totalST += dayHours;
            stHours[fi] = String(dayHours);
          } else {
            totalST += 8;
            totalOT += dayHours - 8;
            stHours[fi] = '8';
            otHours[fi] = String(dayHours - 8);
          }
        }
      });

      const grossPay = (totalST * stRate) + (totalOT * otRate);

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
        erFringe, empFringe,
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
        rate_st:         stRate.toFixed(2),
        rate_ot:         otRate.toFixed(2),
        gross_pay:       grossPay.toFixed(2),
        total_gross_pay: ytdGross.toFixed(2),   // YTD across all projects
        net_pay:         '',
        deductions:      '',
        annualized_rate: ''
      });
    });

    // ── JSON export for local PDF filler ─────────────────────────────────────
    const weekEndFormatted = Utilities.formatDate(weekEnd, CONFIG.TIMEZONE, 'MM/dd/yyyy');
    const sigDate  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MM/dd'); // MM/DD only — form pre-prints the year separately
    const sigYear  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yy');    // last 2 digits — form pre-prints "20"

    const cpContractNumber = clRow ? String(clRow[0] || '') : '';

    const cpJson = {
      _type:                 'certified_payroll',
      payroll_number:        '',
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
      signatory: {
        name:  CONFIG.SIGNATORY.name,
        title: CONFIG.SIGNATORY.title,
        date:  sigDate,
        year:  sigYear
      }
    };
    const cpProps = PropertiesService.getScriptProperties();
    const cpFolder = getOrCreateSubfolder_(
      DriveApp.getFolderById(cpProps.getProperty('NEEDS_REVIEW_ID')), 'Certified Payroll'
    );
    const cpJsonName = `Certified_Payroll_${contractNum}_${borough}_${Utilities.formatDate(weekEnd, CONFIG.TIMEZONE, 'yyyy-MM-dd')}.json`;
    cpFolder.createFile(cpJsonName, JSON.stringify(cpJson, null, 2), MimeType.PLAIN_TEXT);
    Logger.log(`✅ Certified payroll JSON exported: ${cpJsonName}`);
    // ─────────────────────────────────────────────────────────────────────────

    Logger.log(`✅ Certified payroll entries created for ${contractNum}/${borough}: ${Object.keys(byEmployee).length} employees`);
  });

  const contractCount = Object.keys(byContract).length;

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
  const response = ui.prompt('Enter week start date (Monday, MM/DD/YYYY):');
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
    } else if (action === 'finalize_field_report_docs') {
      return handleFinalizeFieldReportDocs_(body);
    } else if (action === 'get_dashboard_data') {
      return handleGetDashboardData_();
    } else if (action === 'upload_photo') {
      return handleUploadPhoto_(body);
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
    } else if (action === 'get_drive_file_bytes') {
      return handleGetDriveFileBytes_(body);
    } else if (action === 'approve_doc') {
      return handleApproveDoc_(body);
    } else if (action === 'approve_signin_with_bytes') {
      return handleApproveSignInWithBytes_(body);
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
    // Pull contract + borough + date if we can; otherwise fall back.
    const m = s.match(/Certified_Payroll_(\w+)_(\w+)_(\d{4}-\d{2}-\d{2})/);
    if (m) return `${m[1]}-${m[2]} · ${m[3]}`;
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

  approvals.sort((a, b) => a.created_at < b.created_at ? 1 : -1);
  return jsonResponse_({ approvals });
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

  try {
    const file = DriveApp.getFileById(fileId);
    if (file.isTrashed()) return jsonResponse_({ error: 'File is trashed' }, 404);

    // Safety gate: verify the file lives under NEEDS_REVIEW_ID. Walks up
    // the parents tree (Drive files can have multiple, but Scan Inbox /
    // Needs Review subfolders are single-parent).
    const props         = PropertiesService.getScriptProperties();
    const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');
    if (!_isUnderParent_(file, needsReviewId)) {
      return jsonResponse_({ error: 'File not in Docs Needing Review' }, 403);
    }

    const blob = file.getBlob();
    return jsonResponse_({
      filename:  file.getName(),
      mime_type: blob.getContentType(),
      size:      file.getSize(),
      data:      Utilities.base64Encode(blob.getBytes()),
    });
  } catch (err) {
    return jsonResponse_({ error: String(err) }, 500);
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

  const props          = PropertiesService.getScriptProperties();
  const approvedId     = props.getProperty('APPROVED_SENT_ID');
  const needsReviewId  = props.getProperty('NEEDS_REVIEW_ID');
  if (!approvedId)    return jsonResponse_({ error: 'APPROVED_SENT_ID not set' }, 500);
  if (!needsReviewId) return jsonResponse_({ error: 'NEEDS_REVIEW_ID not set' }, 500);

  try {
    const file = DriveApp.getFileById(fileId);
    if (file.isTrashed()) return jsonResponse_({ error: 'File is trashed' }, 404);
    if (!_isUnderParent_(file, needsReviewId)) {
      return jsonResponse_({ error: 'File not in Docs Needing Review' }, 403);
    }

    const approvedFolder = DriveApp.getFolderById(approvedId);
    file.moveTo(approvedFolder);

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
    return jsonResponse_({ error: String(err) }, 500);
  }
}


// ── action: approve_signin_with_bytes ─────────────────────────
//
// Sign-In docs have a principal sign-off block that's filled DURING
// the approval flow (see webapp PrincipalSignModal + Express
// /api/approvals/:fileId/approve-signin). Express patches the PDF via
// pdf-lib and POSTs the resulting bytes here. This handler does the
// upload + trash + log in one atomic move so the cron sees the signed
// PDF immediately.
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
 *  9  Priority Level        20  SQFT Completed         31  Payment Date
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
    '', '', '',                  // 20-22  SQFT Completed, Paint/Material, Issues
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

  return jsonResponse_({ success: true, work_order_id: d.work_order_id });
}


/**
 * Archive the original WO PDF scan into:
 *   Archive / [Contractor] / [ContractNum - Borough] / [WO# - Location] /
 * Then rename the original in Scan Inbox with ✅ prefix to prevent reprocessing.
 */
function archiveWOFile_(fileId, d) {
  try {
    const props      = PropertiesService.getScriptProperties();
    const archiveId  = props.getProperty('ARCHIVE_ID');
    if (!archiveId || !fileId) return;

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

    Logger.log('📁 WO PDF archived: ' + d.work_order_id + ' → ' + contractor + '/' + contractNum);
  } catch (err) {
    Logger.log('⚠️ Could not archive WO file: ' + err.toString());
    // Non-fatal — WO row was already written to tracker
  }
}


/**
 * Idempotently ensures the WO Tracker has every extra column we've
 * added over time: the 3 CFR cols (Date Entered, School, Prep By) and
 * the 2 scan-tracking cols (Scan File ID, Combined Scan File ID).
 * Safe to call on every scan intake — only writes headers that are
 * actually missing.
 *
 * Column layout (1-indexed): 36 = Date Entered, 37 = School,
 *   38 = Prep By, 39 = Scan File ID, 40 = Combined Scan File ID.
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
 * Trash a Drive file by ID. Used by the Python worker after it fills a PDF
 * from a JSON payload — trashing the source JSON prevents the poll loop
 * from re-processing it if the worker's local `.processed_files.json`
 * state gets lost (e.g. Railway container restart with ephemeral disk).
 *
 * body: { key, file_id }
 */
function handleTrashFile_(body) {
  const fileId = body.file_id;
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
  'Speed Hump Markings': 'EA',
  'Shark Teeth 12x18':   'EA',
  'Shark Teeth 24x36':   'EA',
  'Bike Lane Arrow':     'EA',
  'Bike Lane Symbol':    'EA',
  'Bike Lane Green Bar': 'EA',
  'Ped Stop':            'EA',
  // "Others" is intentionally variable (user picks).
};

function unitForCategory_(category) {
  return CATEGORY_UNITS_[String(category || '').trim()] || '';
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
  const topMarkings = d.top_markings       || [];
  const grid        = d.intersection_grid  || [];
  if (topMarkings.length === 0 && grid.length === 0) return 0;

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
 * MMA:    marking_types = distinct Category joined by ", "
 *         paint_material = distinct Color/Material joined by ", "
 *         sqft_completed = SUM(Quantity WHERE Unit = 'SF')
 *
 * Thermo: marking_types = "N/A"     (too many categories to rollup)
 *         paint_material = "N/A"    (Thermo doesn't record material)
 *         sqft_completed = SUM(Quantity WHERE Unit = 'SF')
 *
 * Only COMPLETED items contribute to marking_types / paint_material.
 * SQFT includes all items with a quantity (regardless of status) so the
 * tracker reflects what's been measured even if status is still Pending.
 */
function computeMarkingRollups_(ss, woId, preloadedData) {
  const blank = { marking_types: '', sqft_completed: '', paint_material: '' };
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

  // Column indices (0-based) under the 15-col schema:
  //   0 Item ID     4 Marking Type   8 Quantity      12 Status
  //   1 WO#         5 Intersection   9 Unit          13 Added By
  //   2 Work Type   6 Direction      10 Material     14 Notes
  //   3 WO Section  7 Description    11 Date Completed
  let sqftSum = 0;
  let hasQty  = false;
  woItems.forEach(r => {
    const unit = String(r[9] || '').toUpperCase();
    if (unit !== 'SF') return;
    const qty = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;
    sqftSum += qty;
    hasQty = true;
  });

  if (anyThermo) {
    return {
      marking_types:  'N/A',
      sqft_completed: hasQty ? sqftSum : '',
      paint_material: 'N/A',
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
    marking_types:  Object.keys(cats).sort().join(', '),
    sqft_completed: hasQty ? sqftSum : '',
    paint_material: Object.keys(mats).sort().join(', '),
  };
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
  if (wasCompleted && anyFieldChanged) {
    sheet.getRange(rowNum, COL.status).setValue('Pending');
    sheet.getRange(rowNum, COL.date_completed).setValue('');
  }

  SpreadsheetApp.flush();
  const item = readMarkingItemById_(sheet, itemId);
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
function finalizeMarkingStatus_(ss, woId, dateOfWork) {
  const sheet = ss.getSheetByName('Marking Items');
  if (!sheet) return { touched: 0, data: null };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { touched: 0, data };

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
        // Keep the in-memory copy in sync so downstream rollup readers
        // don't have to re-fetch the sheet.
        r[12] = 'Completed';
        r[11] = dateOfWork;
        touched += 2;
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

  const ORDER = { 'in progress': 0, 'dispatched': 1, 'received': 2 };

  const wos = allRows.slice(1)
    .filter(r => r[0] && String(r[15]).toLowerCase() !== 'completed')
    .map(r => ({
      id:              String(r[0]),
      contractor:      String(r[1]),
      contract_number: String(r[2]),
      borough:         String(r[3]),
      location:        String(r[5]),
      from_street:     String(r[6]),
      to_street:       String(r[7]),
      due_date:        String(r[8]),
      priority:        String(r[9]),
      work_type:       String(r[10]),
      status:          String(r[15]),
      dispatch_date:   String(r[16] || ''),
      work_start_date: String(r[17] || '')
    }))
    .sort((a, b) => {
      const aO = ORDER[a.status.toLowerCase()] ?? 99;
      const bO = ORDER[b.status.toLowerCase()] ?? 99;
      return aO - bO;
    });

  return jsonResponse_({ wos });
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
 *   marking_types   — string ("Crosswalk: 500 SF, Stop Bar: 10 LF")
 *   sqft_completed  — number
 *   paint_material  — string
 *   issues          — string (appended to existing issues with date prefix)
 *   photos_uploaded — boolean
 *   crew            — [{name, classification, time_in, time_out, hours, overtime}]
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
 *   16  Dispatch Date      20  SQFT Completed
 *   17  Work Start Date    21  Paint/Material Used
 *   18  Work End Date      22  Issues Reported
 */
function handleSubmitFieldReport_(body) {
  const d = body.data || {};

  if (!d.wo_id) return jsonResponse_({ error: 'Missing wo_id' }, 400);
  if (!d.date)  return jsonResponse_({ error: 'Missing date' }, 400);
  if (!Array.isArray(d.crew) || d.crew.length === 0) {
    return jsonResponse_({ error: 'At least one crew member is required' }, 400);
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

  // ── Derive updated WO Tracker values ─────────────────────────

  // Status: Received → Dispatched → In Progress → Complete
  const currentStatus    = String(woRow[15] || 'Received');
  const currentDispatch  = woRow[16] ? String(woRow[16]) : '';
  const currentWorkStart = woRow[17] ? String(woRow[17]) : '';
  const currentWorkEnd   = woRow[18] ? String(woRow[18]) : '';
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
  const { touched: nFinalized, data: markingData } = finalizeMarkingStatus_(ss, d.wo_id, d.date);

  step = 'marking rollup';
  // Reuse the values finalizeMarkingStatus_ already read (with its
  // in-memory updates applied). Saves a full-sheet getDataRange call.
  const rollups = computeMarkingRollups_(ss, d.wo_id, markingData);
  Logger.log(`✅ Marking Items: ${nFinalized} status cells updated; ` +
             `rollup → types=${JSON.stringify(rollups.marking_types)}, ` +
             `sqft=${rollups.sqft_completed}, ` +
             `material=${JSON.stringify(rollups.paint_material)}`);

  // ── Write WO Tracker cols 15–23 (0-indexed) ──────────────────
  // Col 15 (0-indexed) = col 16 (1-indexed); 9 columns → cols 16–24 (1-indexed)
  const woValues = [
    newStatus,                          // col 15: Status
    newDispatch,                        // col 16: Dispatch Date
    newWorkStart,                       // col 17: Work Start Date
    newWorkEnd,                         // col 18: Work End Date
    rollups.marking_types,              // col 19: Marking Types (rollup)
    rollups.sqft_completed,             // col 20: SQFT Completed (rollup)
    rollups.paint_material,             // col 21: Paint/Material Used (rollup)
    newIssues,                          // col 22: Issues Reported
    newPhotos                           // col 23: Photos Uploaded?
  ];
  const woLabels = [
    'Status', 'Dispatch Date', 'Work Start Date', 'Work End Date',
    'Marking Types', 'SQFT Completed', 'Paint/Material',
    'Issues Reported', 'Photos Uploaded?'
  ];
  step = 'WO Tracker write';
  writeRowWithProbing_(woSheet, woRowIdx + 1, 16, woValues, woLabels, 'WO Tracker');

  Logger.log('✅ WO Tracker updated: ' + d.wo_id + ' → ' + newStatus);

  // ── Write Daily Sign-In Data rows (one per crew member) ──────
  // New 14-col schema (cols 12-13 = Admin Reviewed? / Review Notes, filled
  // later by admin — we leave them truly blank). WO-level fields that used
  // to live here (SQFT, Paint/Material, WO Complete?, Issues/Notes) moved
  // to Marking Items rollups + WO Tracker cols 19-22.
  const signInSheet = ss.getSheetByName('Daily Sign-In Data');
  const signInLabels = [
    'Date', 'Work Order #', 'Prime Contractor', 'Contract #', 'Borough',
    'Location', 'Employee Name', 'Classification', 'Time In', 'Time Out',
    'Hours Worked', 'Overtime Hours'
  ];

  // Day of work drives the OT rule below. d.date is ISO "YYYY-MM-DD"
  // (from <input type="date">), so construct in local parts to avoid
  // UTC-shift (new Date("2026-04-20") is interpreted as UTC midnight).
  const dowOfWork = (() => {
    const m = String(d.date || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return -1;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay();
  })();
  const isWeekend = (dowOfWork === 0 || dowOfWork === 6);

  // Build every crew row in memory first, then hand them to the batched
  // appender so the whole crew lands in one setValues round-trip instead
  // of N.
  step = 'Sign-In Data batch build';
  const crewRows = d.crew.map(member => {
    const hours    = parseFloat(member.hours) || 0;
    // Apply the OT rule server-side (authoritative). Client-supplied
    // overtime is ignored — the rule is:
    //   Saturday/Sunday → every hour is OT.
    //   Monday–Friday   → hours over 8 in the day are OT; first 8 are ST.
    const overtime = isWeekend ? hours : Math.max(0, hours - 8);
    return [
      d.date,                              //  0  Date
      d.wo_id,                             //  1  Work Order #
      String(woRow[1]),                    //  2  Prime Contractor
      String(woRow[2]),                    //  3  Contract #
      String(woRow[3]),                    //  4  Borough
      String(woRow[5]),                    //  5  Location
      String(member.name   || '').trim(),  //  6  Employee Name
      String(member.classification || ''), //  7  Classification
      String(member.time_in  || ''),       //  8  Time In
      String(member.time_out || ''),       //  9  Time Out
      hours,                               // 10  Hours Worked
      overtime                             // 11  Overtime Hours
      // cols 12 (Admin Reviewed?) and 13 (Review Notes) omitted so the
      // cells stay truly blank — the dropdown validator rejects an
      // explicit '' but accepts a truly empty cell.
    ];
  });
  step = `Sign-In Data append x${crewRows.length}`;
  appendRowsWithProbing_(signInSheet, crewRows, signInLabels, 'Daily Sign-In Data');

  Logger.log('✅ Sign-In Data: ' + d.crew.length + ' row(s) appended for WO ' + d.wo_id);

  // Sign-In JSON + CFR JSON generation moved to a separate action
  // (`finalize_field_report_docs`) the client fires AFTER it gets this
  // success response. Keeps the user-facing submit path fast (the JSON
  // writes are 2-5s of Drive I/O that don't affect the submitted data
  // integrity — worst case a generated PDF is missing and we see a row
  // in Automation Log).

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
      d.crew.length + ' crew member(s) on ' + d.date,
      newStatus,
      '',
      actionNote
    ],
    ['Timestamp', 'Source', 'Action', 'Related', 'Details', 'Status', 'User', 'Next Steps'],
    'Automation Log'
  );

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
 * Sign-In Log JSON (always) and the Contractor Field Report JSON
 * (only when `wo_complete === true`) to Drive so the Railway worker
 * can pick them up and produce the filled PDFs.
 *
 * The client fires this as a separate POST right after the main
 * `submit_field_report` returns success. Failures go to Automation Log
 * — they never bubble back to the user because by this point the
 * report data is already safely persisted in the spreadsheet.
 *
 * Expects body.data to be the same payload the submit got (includes
 * signatures). Re-reads the WO Tracker row + the newly-written Issues
 * Reported aggregate so we don't depend on the client shipping them.
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

  // Sign-In JSON — always attempted
  try {
    generateSignInJson_(d, woRow, ss);
    result.signin = 'ok';
  } catch (err) {
    Logger.log('⚠️ Sign-In JSON export failed: ' + err);
    result.signin = 'failed';
    try {
      ss.getSheetByName('Automation Log').appendRow([
        new Date(), 'Sign-In JSON Export', 'Failed', d.wo_id,
        String(err), 'Error', '', 'Check logs — sign-in PDF will not be generated'
      ]);
    } catch (logErr) {
      Logger.log('⚠️ Could not write Sign-In failure to Automation Log: ' + logErr);
    }
  }

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


/**
 * Build a SignIn JSON payload from the submitted field report + WO Tracker row,
 * and write it to Drive folder "Needs Review / Sign-In Logs / ". The Railway
 * Python worker (watch_and_fill.py) polls that folder, fills the Sign-In PDF
 * template with embedded signatures, and uploads the result back.
 *
 * Signatures are expected inline as base64 data URLs on each crew member
 * (sig_in_b64, sig_out_b64) and on the crew-leader block
 * (contractor_signature_b64). They are forwarded verbatim — never persisted
 * separately to Drive.
 *
 * Date formatting: the Sign-In PDF shows "M/D/YY". We accept "YYYY-MM-DD"
 * from the web app and reformat.
 */
function generateSignInJson_(d, woRow, ss) {
  const props          = PropertiesService.getScriptProperties();
  const needsReviewId  = props.getProperty('NEEDS_REVIEW_ID');
  if (!needsReviewId) throw new Error('NEEDS_REVIEW_ID not set');

  const reviewFolder   = DriveApp.getFolderById(needsReviewId);
  const signInFolder   = getOrCreateSubfolder_(reviewFolder, 'Sign-In Logs');

  // WO Tracker columns: 1=Prime Contractor, 2=Contract #, 3=Borough, 5=Location
  const primeContractor = String(woRow[1] || '').trim();
  const contractNum     = String(woRow[2] || '').trim();
  const boroughCode     = String(woRow[3] || '').trim();
  const location        = String(woRow[5] || '').trim();
  const boroughName     = boroughCode ? getBoroughName_(boroughCode) : '';

  // "Contract #" field on the sign-in form shows contract + borough
  // (user preference — we don't have a separate Registration # field).
  const contractLabel = boroughName
    ? `${contractNum} - ${boroughName}`
    : contractNum;

  // Look up the prime contractor's address from Contractor Contacts.
  // Sheet layout: col 0 = Contractor, col 5 = Address (billing/yard).
  // Best-effort — blank if sheet missing or contractor not listed.
  let primeAddress = '';
  try {
    const ccSheet = ss.getSheetByName('Contractor Contacts');
    if (ccSheet) {
      const ccData = ccSheet.getDataRange().getValues();
      const ccRow = ccData.find(r => String(r[0] || '').trim() === primeContractor);
      if (ccRow) primeAddress = String(ccRow[5] || '').trim();
    }
  } catch (err) {
    Logger.log(`⚠️ Contractor Contacts lookup failed for ${primeContractor}: ${err}`);
  }

  // Project Name/Location = "<WO#> | <Location>"
  const projectName = location
    ? `${d.wo_id} | ${location}`
    : d.wo_id;

  // Reformat YYYY-MM-DD → M/D/YY to match the paper form's date style
  const dateFmt = (iso) => {
    if (!iso) return '';
    const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return String(iso);
    const [, yyyy, mm, dd] = m;
    return `${parseInt(mm, 10)}/${parseInt(dd, 10)}/${yyyy.slice(-2)}`;
  };
  const dateHeader = dateFmt(d.date);

  const payload = {
    _type:              'signin',
    wo_id:              d.wo_id,
    date:               dateHeader,
    prime_contractor:   primeContractor,
    subcontractor:      CONFIG.EMPLOYER.name,
    contract_number:    contractLabel,
    address:            primeAddress,
    agency:             'NYCDOT',
    project_name:       projectName,
    crew: (d.crew || []).map(m => ({
      name:            String(m.name || '').trim(),
      classification:  String(m.classification || '').trim(),
      time_in:         String(m.time_in  || ''),
      time_out:        String(m.time_out || ''),
      sig_in_b64:      m.sig_in_b64  || '',
      sig_out_b64:     m.sig_out_b64 || '',
    })),
    // Crew-leader block at the bottom of the form. The Field Report web app
    // sends these alongside the per-member sigs.
    // Bottom sign-off block stays blank at generation time. The
    // principal signs + types name/title during the Approvals flow
    // (PrincipalSignModal → pdf-lib in Express). Date is injected at
    // that same step so it matches the actual approval date.
    contractor_name:            '',
    contractor_title:           '',
    date_signed:                '',
    contractor_signature_b64:   '',
  };

  // Filename: SignIn_<WO>_<YYYY-MM-DD>.json — matches the worker's expected
  // naming so deduplication works if the admin re-submits.
  const isoDate = (d.date || '').slice(0, 10) || 'unknown';
  const fileName = `SignIn_${d.wo_id}_${isoDate}.json`;

  // Overwrite any existing file with the same name (same-day re-submit)
  const existing = signInFolder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  signInFolder.createFile(
    fileName,
    JSON.stringify(payload, null, 2),
    MimeType.PLAIN_TEXT
  );

  Logger.log('✅ Sign-In JSON exported: ' + fileName);
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
    'Bike Lane Arrow':    'Bike Lane Arrows',
    'Bike Lane Symbol':   'Bike Lane Symbols',
  };
  // Grid-only categories — never in top table
  const GRID_ONLY = { 'HVX Crosswalk': 1, 'Stop Line': 1, 'Stop Msg': 1 };

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
  // Preserve first-seen order via an index map.
  const rowByIntersection = {};
  const order = [];

  woItems.forEach(r => {
    const intersection = String(r[5] || '').trim();
    if (!intersection) return;
    const category = String(r[4] || '').trim();
    const direction = String(r[6] || '').trim().toUpperCase();
    const qty = parseFloat(r[8]);
    if (isNaN(qty) || qty <= 0) return;

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

    if (category === 'HVX Crosswalk' && ['N','E','S','W'].indexOf(direction) !== -1) {
      const key = direction.toLowerCase();
      gridRow[key] = (parseFloat(gridRow[key]) || 0) + qty;
    } else if (category === 'Stop Msg') {
      gridRow.stop_msg = (parseFloat(gridRow.stop_msg) || 0) + qty;
    } else if (category === 'Stop Line') {
      gridRow.st_line = (parseFloat(gridRow.st_line) || 0) + qty;
    }
    // School Msg 8'/10': no source category → stays blank.
  });

  out.grid = order.slice(0, 10).map(i => rowByIntersection[i]);
  return out;
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
function generateContractorFieldReportJson_(d, woRow, ss, aggregatedIssues) {
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

  // Flatten multi-line issues to a single line for the PDF (the General
  // Remarks field is single-line — a raw \n gets truncated at the first
  // line). Each issue line already carries its own date prefix, so " | "
  // is enough visual separation.
  const flatRemarks = String(aggregatedIssues || '')
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(Boolean)
    .join(' | ');

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
  };

  const isoDate  = (d.date || '').slice(0, 10) || 'unknown';
  const fileName = `CFR_${d.wo_id}_${isoDate}.json`;

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
 */
function handleGetDashboardData_() {
  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();

  const wos = allRows.slice(1)
    .filter(r => r[0])   // skip blank rows
    .map(r => ({
      id:            String(r[0]),
      contractor:    String(r[1]),
      contract_num:  String(r[2]),
      borough:       String(r[3]),
      contract_id:   String(r[4]),
      location:      String(r[5]),
      from_street:   String(r[6]),
      to_street:     String(r[7]),
      due_date:      String(r[8]  || ''),
      priority:      String(r[9]  || ''),
      work_type:     String(r[10] || ''),
      wo_received:   String(r[11] || ''),
      water_blast:   String(r[12] || ''),
      status:        String(r[15] || 'Received'),
      dispatch_date: String(r[16] || ''),
      work_start:    String(r[17] || ''),
      work_end:      String(r[18] || ''),
      marking_types: String(r[19] || ''),
      sqft:          r[20] != null ? String(r[20]) : '',
      paint:         String(r[21] || ''),
      issues:        String(r[22] || ''),
      photos:        String(r[23] || ''),
      prod_log:      String(r[24] || ''),
      field_report:  String(r[25] || ''),
      invoice_sent:  String(r[29] || ''),
      payment_recv:  String(r[30] || '')
    }));

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

  return jsonResponse_({ wos, stats, byContractor, attention });
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
  const d = body.data || {};
  const { wo_id, filename, mime_type, data } = d;
  if (!wo_id || !filename || !data) {
    return jsonResponse_({ error: 'Missing required fields: wo_id, filename, data' }, 400);
  }

  const photosFolder = getWOSubfolder_(wo_id, 'Photos');
  if (!photosFolder) {
    return jsonResponse_({ error: 'Could not locate or create WO archive folder' }, 500);
  }

  const bytes = Utilities.base64Decode(data);
  const blob  = Utilities.newBlob(bytes, mime_type || 'image/jpeg', filename);
  const file  = photosFolder.createFile(blob);

  Logger.log('📸 Photo uploaded for WO ' + wo_id + ': ' + filename);
  return jsonResponse_({ success: true, file_id: file.getId(), file_url: file.getUrl() });
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

  const sigsFolder = getWOSubfolder_(wo_id, 'Signatures');
  if (!sigsFolder) {
    return jsonResponse_({ error: 'Could not locate or create WO archive folder' }, 500);
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
function getWOSubfolder_(wo_id, subfolderName) {
  try {
    const props     = PropertiesService.getScriptProperties();
    const archiveId = props.getProperty('ARCHIVE_ID');
    if (!archiveId) return null;

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
  } catch (err) {
    Logger.log('⚠️ getWOSubfolder_ error: ' + err.toString());
    return null;
  }
}
