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

/**
 * Parse a time string like "9:15 AM" or "1:15 PM" into minutes since midnight.
 * Used for correct earliest-in / latest-out comparisons (string compare fails on AM/PM).
 */
function parseTimeToMinutes_(timeStr) {
  if (!timeStr) return 0;
  const s = String(timeStr).trim().toUpperCase();
  const match = s.match(/(\d+):(\d+)\s*(AM|PM)/);
  if (!match) return 0;
  let h = parseInt(match[1]);
  const m = parseInt(match[2]);
  const ampm = match[3];
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return h * 60 + m;
}

/**
 * Generate Metro Thermoplastic Production Daily Log
 */
function generateProductionLog_(targetDate, allEntries, byWorkOrder, woData, ss) {
  const props = PropertiesService.getScriptProperties();
  const needsReviewId = props.getProperty('NEEDS_REVIEW_ID');

  // Get unique employees — track earliest time-in and latest time-out across all WOs for the day
  const employees = {};
  allEntries.forEach(row => {
    const name = row[6];
    const timeInMins  = parseTimeToMinutes_(row[8]);
    const timeOutMins = parseTimeToMinutes_(row[9]);
    if (!employees[name]) {
      employees[name] = {
        timeIn: row[8], timeOut: row[9],
        timeInMins, timeOutMins,
        classification: row[7]
      };
    } else {
      // Correct numeric comparison — not string comparison
      if (timeInMins  < employees[name].timeInMins)  { employees[name].timeIn  = row[8]; employees[name].timeInMins  = timeInMins;  }
      if (timeOutMins > employees[name].timeOutMins) { employees[name].timeOut = row[9]; employees[name].timeOutMins = timeOutMins; }
    }
  });
  
  // Build the log content
  const dateFormatted = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'MM/dd/yyyy');
  let logContent = `METRO THERMOPLASTIC PRODUCTION DAILY LOG\n`;
  logContent += `Date: ${dateFormatted}\n\n`;
  
  // Crew members
  Object.entries(employees).forEach(([name, info]) => {
    const role = info.classification === 'LP' ? 'Crew Chief' : 'Individual';
    logContent += `${role}: ${name} | In: ${info.timeIn} | Out: ${info.timeOut}\n`;
  });
  
  logContent += `\n${'─'.repeat(80)}\n\n`;
  logContent += `WORK ORDERS:\n\n`;
  logContent += `${'Borough'.padEnd(8)} ${'WO #'.padEnd(12)} ${'Location'.padEnd(25)} ${'SQFT'.padEnd(10)} ${'Paint'.padEnd(15)} ${'Complete?'}\n`;
  logContent += `${'─'.repeat(80)}\n`;
  
  Object.entries(byWorkOrder).forEach(([woId, entries]) => {
    // Look up borough from WO tracker
    const woRow = woData.find(r => r[0] === woId);
    const borough = woRow ? woRow[3] : '??';
    const location = woRow ? woRow[5] : entries[0][5];
    const sqft = entries[0][12];
    const paint = entries[0][13];
    const complete = entries[0][14];
    
    logContent += `${String(borough).padEnd(8)} ${String(woId).padEnd(12)} ${String(location).padEnd(25)} ${String(sqft).padEnd(10)} ${String(paint).padEnd(15)} ${complete}\n`;
  });
  
  // Save to Needs Review folder
  const reviewFolder = DriveApp.getFolderById(needsReviewId);
  const subFolder = getOrCreateSubfolder_(reviewFolder, 'Production Logs');
  const fileName = `Production_Log_${Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'yyyy-MM-dd')}.txt`;
  const file = subFolder.createFile(fileName, logContent, MimeType.PLAIN_TEXT);

  // ── JSON export for local PDF filler ─────────────────────────────────────
  const sortedNames = Object.keys(employees);
  const crewChiefName = sortedNames.find(n => employees[n].classification === 'LP') || sortedNames[0];
  const crewMemberNames = sortedNames.filter(n => n !== crewChiefName);

  const workOrdersJson = Object.entries(byWorkOrder).map(([woId, entries]) => {
    const woRow = woData.find(r => r[0] === woId);
    const borough = woRow ? String(woRow[3]).toUpperCase() : '';
    const location = woRow ? String(woRow[5]).toUpperCase() : String(entries[0][5]).toUpperCase();
    return {
      wo_number:    String(woId),
      borough:      borough,
      location:     location,
      sqft:         String(entries[0][12] || ''),
      paint:        String(entries[0][13] || ''),
      complete:     entries[0][14] ? 'Y' : 'N',
      layout_yn:    '',
      layout_hours: '',
      markings:     {}
    };
  });

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

  const jsonFileName = `Production_Log_${Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'yyyy-MM-dd')}.json`;
  subFolder.createFile(jsonFileName, JSON.stringify(logJson, null, 2), MimeType.PLAIN_TEXT);
  // ─────────────────────────────────────────────────────────────────────────

  // Log it
  const logSheet = ss.getSheetByName('Automation Log');
  logSheet.appendRow([
    new Date(), 'Production Log Generator', 'Manual trigger',
    `${Object.keys(byWorkOrder).length} WOs on ${dateFormatted}`,
    fileName, 'Generated',
    '', 'Yes — Review and send to Claudia'
  ]);

  Logger.log('✅ Production log generated: ' + fileName);
  Logger.log('✅ Production log JSON exported: ' + jsonFileName);
  return file;
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

    // Write to Certified Payroll Tracker
    Object.entries(byEmployee).forEach(([empName, info]) => {
      // Look up pay rates + address/SSN4 from Employee Registry
      const empRow = empData.find(r => String(r[1]).includes(empName.split(' ')[0]));
      const stRate    = empRow ? Number(empRow[6]) : 0;
      const otRate    = empRow ? Number(empRow[7]) : 0;
      const empFringe = empRow ? Number(empRow[9]) : 0;
      const erFringe  = empRow ? Number(empRow[8]) : 0;
      const empAddr   = empRow ? String(empRow[2] || '') : '';
      const empSsn4   = empRow ? String(empRow[3] || '') : '';

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
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
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
  } catch (batchErr) {
    // Probe per-column to identify which one failed
    for (let i = 0; i < values.length; i++) {
      try {
        sheet.getRange(row, startCol + i).setValue(values[i]);
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
  try {
    sheet.appendRow(values);
    return;
  } catch (batchErr) {
    // Probe per-column by writing to the row that appendRow *would have* filled
    const targetRow = sheet.getLastRow() + 1;
    let culprit = null;
    try {
      for (let i = 0; i < values.length; i++) {
        try {
          sheet.getRange(targetRow, i + 1).setValue(values[i]);
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
      if (sheet.getLastRow() >= targetRow) {
        sheet.deleteRow(targetRow);
      }
    }
    if (culprit) throw culprit;

    // Per-cell probe didn't reproduce. That happens when appendRow validates
    // the row atomically but per-cell setValue bypasses validation (e.g.
    // dropdown ranges applied only to specific rows, cross-cell rules).
    // Read the validation rules on row 2 (first data row) and check each
    // value against any list-type rule — that identifies the culprit
    // deterministically.
    try {
      const probeRow   = sheet.getRange(2, 1, 1, values.length);
      const rules      = probeRow.getDataValidations()[0];
      for (let i = 0; i < values.length; i++) {
        const rule = rules[i];
        if (!rule) continue;
        const criteriaType = rule.getCriteriaType();
        const criteriaName = criteriaType ? String(criteriaType) : '';
        // Covers VALUE_IN_LIST and VALUE_IN_RANGE
        if (criteriaName.indexOf('VALUE_IN') === -1) continue;
        const crArgs = rule.getCriteriaValues() || [];
        let allowed = [];
        if (criteriaName.indexOf('LIST') !== -1) {
          allowed = Array.isArray(crArgs[0]) ? crArgs[0].map(String) : crArgs.map(String);
        } else if (criteriaName.indexOf('RANGE') !== -1 && crArgs[0]) {
          try {
            allowed = crArgs[0].getValues().flat().filter(v => v !== '').map(String);
          } catch (rangeErr) {
            allowed = [];
          }
        }
        const raw    = values[i];
        const asStr  = (raw == null) ? '' : String(raw);
        if (asStr === '') continue;  // empty usually permitted
        if (allowed.length && allowed.indexOf(asStr) === -1) {
          throw new Error(
            `${sheetLabel} → "${labels[i] || 'col ' + (i + 1)}" value ` +
            `${JSON.stringify(raw)} not in allowed list [${allowed.join(', ')}]. ` +
            `Criteria: ${criteriaName}`
          );
        }
      }
    } catch (inspectErr) {
      // Re-throw specific diagnosis; otherwise fall through to value dump.
      if (String(inspectErr.message).indexOf(sheetLabel) === 0) throw inspectErr;
    }

    // Fall-through: dump all values with their labels so the caller can
    // eyeball which one violates validation even when automated detection
    // missed it. This is always better than a bare "Invalid Entry" message.
    const summary = values
      .map((v, i) => `${labels[i] || 'col' + (i + 1)}=${JSON.stringify(v)}`)
      .join(' | ');
    throw new Error(
      `${sheetLabel} → batch append failed (per-cell probe could not isolate). ` +
      `Row values: [${summary}]. Original: ${batchErr.message}`
    );
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
    } else if (action === 'get_dashboard_data') {
      return handleGetDashboardData_();
    } else if (action === 'upload_photo') {
      return handleUploadPhoto_(body);
    } else if (action === 'upload_signature') {
      return handleUploadSignature_(body);
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


// ── action: write_wo ──────────────────────────────────────────────────────────

/**
 * Writes a parsed Work Order to the WO Tracker sheet and archives the source PDF.
 *
 * body.file_id  — Drive file ID of the original scanned WO PDF
 * body.data     — normalized dict from parse_work_order.normalize_wo_data()
 *
 * WO Tracker columns (0-indexed, 35 total):
 *  0  Work Order #          11  WO Received Date      22  Issues Reported
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
 */
function handleWriteWO_(body) {
  const fileId = body.file_id;
  const d      = body.data || {};

  if (!d.work_order_id) {
    return jsonResponse_({ error: 'Missing work_order_id in data' }, 400);
  }

  const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const woSheet = ss.getSheetByName('Work Order Tracker');
  const allRows = woSheet.getDataRange().getValues();

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
    ''                           // 34  Notes             ← from crew web app
  ];

  woSheet.appendRow(row);
  Logger.log('✅ WO added to tracker: ' + d.work_order_id
             + (contractIdMissing ? ' (Contract ID not found in lookup)' : ''));

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


/** Wraps an object as a JSON ContentService response. */
function jsonResponse_(obj, statusCode) {
  if (statusCode && statusCode !== 200) obj._status = statusCode;
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
    .filter(r => r[0] && String(r[15]).toLowerCase() !== 'complete')
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
    newStatus = 'Complete';
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

  // Use submitted marking types; fall back to existing if not provided
  const newMarkings = (d.marking_types && d.marking_types.trim())
    ? d.marking_types.trim()
    : String(woRow[19] || '');

  // ── Write WO Tracker cols 15–23 (0-indexed) ──────────────────
  // Col 15 (0-indexed) = col 16 (1-indexed); 9 columns → cols 16–24 (1-indexed)
  const woValues = [
    newStatus,                                                        // col 15: Status
    newDispatch,                                                      // col 16: Dispatch Date
    newWorkStart,                                                     // col 17: Work Start Date
    newWorkEnd,                                                       // col 18: Work End Date
    newMarkings,                                                      // col 19: Marking Types
    (d.sqft_completed != null) ? d.sqft_completed : (woRow[20] || ''),// col 20: SQFT
    (d.paint_material  || String(woRow[21] || '')).trim(),            // col 21: Paint/Material
    newIssues,                                                        // col 22: Issues Reported
    newPhotos                                                         // col 23: Photos Uploaded?
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
  const signInSheet = ss.getSheetByName('Daily Sign-In Data');
  const signInLabels = [
    'Date', 'Work Order #', 'Prime Contractor', 'Contract #', 'Borough',
    'Location', 'Employee Name', 'Classification', 'Time In', 'Time Out',
    'Hours Worked', 'Overtime Hours', 'SQFT Completed', 'Paint/Material',
    'WO Complete?', 'Issues/Notes', 'Admin Reviewed?', 'Review Notes'
  ];

  d.crew.forEach((member, idx) => {
    step = `Sign-In Data row ${idx + 1}/${d.crew.length} (${member.name || '<no name>'})`;
    const hours    = parseFloat(member.hours)    || 0;
    const overtime = parseFloat(member.overtime) || Math.max(0, hours - 8);

    const row = [
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
      overtime,                            // 11  Overtime Hours
      (d.sqft_completed != null) ? d.sqft_completed : '', // 12  SQFT Completed
      (d.paint_material || '').trim(),     // 13  Paint/Material
      d.wo_complete ? 'Yes' : 'No',        // 14  WO Complete?
      (d.issues || '').trim(),             // 15  Issues/Notes
      '',                                  // 16  Admin Reviewed?
      ''                                   // 17  Review Notes
    ];
    appendRowWithProbing_(signInSheet, row, signInLabels, 'Daily Sign-In Data');
  });

  Logger.log('✅ Sign-In Data: ' + d.crew.length + ' row(s) appended for WO ' + d.wo_id);

  // ── Sign-In Log JSON export → Python worker fills the PDF ────
  // Runs on every submit (partial days still need a sign-in sheet).
  // Wrapped so a failure here doesn't reject the whole submit.
  step = 'Sign-In JSON export';
  try {
    generateSignInJson_(d, woRow, ss);
  } catch (err) {
    Logger.log('⚠️ Sign-In JSON export failed: ' + err);
    // The failure-logging write must never throw over the original sign-in
    // error (and must never reject the whole submit). If Automation Log's
    // Status column has dropdown validation that rejects "Error", an
    // un-wrapped appendRow here would surface as the bare "Invalid Entry"
    // message with no indication that it came from the error handler.
    try {
      ss.getSheetByName('Automation Log').appendRow([
        new Date(), 'Sign-In JSON Export', 'Failed', d.wo_id,
        String(err), 'Error', '', 'Check logs — sign-in PDF will not be generated'
      ]);
    } catch (logErr) {
      Logger.log('⚠️ Could not write Sign-In failure to Automation Log: ' + logErr);
    }
  }

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
    address:            '',   // TODO: source from Contractor Contacts yard-address column when populated
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
    contractor_name:            String(d.contractor_name || '').trim(),
    contractor_title:           String(d.contractor_title || 'Crew Leader').trim(),
    date_signed:                dateFmt(d.date_signed || d.date),
    contractor_signature_b64:   d.contractor_signature_b64 || '',
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
    complete:    count('Complete')
  };

  // Contractor breakdown
  const byContractor = {};
  wos.forEach(w => {
    const c = w.contractor || 'Unknown';
    byContractor[c] = (byContractor[c] || 0) + 1;
  });

  // WOs needing attention (issues reported, incomplete docs)
  const attention = wos.filter(w =>
    w.status.toLowerCase() !== 'complete' && (
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
