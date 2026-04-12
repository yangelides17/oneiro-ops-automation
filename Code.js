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
  const approvedSent = root.createFolder('✅ Approved & Sent');
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
 * Processes documents that the admin has moved to the Approved & Sent folder.
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
      new Date(), 'Approve & Send', 'File moved to Approved & Sent folder',
      fileName,
      `Emailed to: ${recipientList} | ${archiveNote}`,
      'Completed', '', 'No'
    ]);

    // Delete from Approved & Sent — archive is now the single source of truth
    file.setTrashed(true);
    Logger.log(`🗑️ Deleted from Approved & Sent: ${fileName}`);
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
    } else {
      return jsonResponse_({ error: 'Unknown action: ' + action }, 400);
    }

  } catch (err) {
    Logger.log('❌ doPost error: ' + err.toString());
    return jsonResponse_({ error: err.toString() }, 500);
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
