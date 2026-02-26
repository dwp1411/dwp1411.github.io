/**
 * Attendance Logger Script
 *
 * This script automates the process of logging hours from the Attendance Tracker to the weekly Timecard sheets.
 *
 * CONFIGURATION:
 * 1. Open your Google Sheet (Attendance Tracker).
 * 2. Go to Extensions > Apps Script.
 * 3. Paste this code into the editor.
 * 4. Save the project.
 *
 * SETUP:
 * 1. Run the `setupTriggers` function once to create a daily trigger (e.g., run every night or every hour) for the "Log All" backup.
 * 2. Run `onOpen` to add the custom menu "Attendance".
 *
 * HOW TO USE:
 * - **Menu Option:** Go to "Attendance" -> "Submit Current Sheet" while on a Supervisor's tab.
 * - **Button:** You can insert a drawing (Insert > Drawing) on each supervisor's sheet, style it as a button (e.g., "Submit Attendance"), and assign the script `submitCurrentSheet` to it.
 *
 * LOGIC:
 * - Finds the "Monday" of the current week.
 * - Searches Drive for the correct Timecard file.
 * - Opens the correct daily tab (e.g., "Monday_").
 * - Reads attendance from the source sheet(s).
 * - Logs 7.5 hours for associates marked 'Y'.
 */

// --- CONFIGURATION ---
const SOURCE_SPREADSHEET_ID = '1YQ0uua9CU04SYBSl1MO_kT9UuO8YNYwJyWMm_3tm8GM'; // The Attendance Tracker Sheet ID
const DESTINATION_FOLDER_ID = '12kzfclHUoALkgkB6c2njL5L8u23II25r'; // The Timecards Folder ID
const SUPERVISOR_TABS = [
  'Abel', 'Andrews', 'Casey', 'Chris', 'Dan', 'Dave',
  'Javier', 'Jesse', 'Mercedes', 'Ramon', 'Roger', 'Tombe'
];
const HOURS_TO_LOG = 7.5;

// --- ENTRY POINTS ---

/**
 * Triggered manually from the menu or a button on the active sheet.
 * Logs attendance ONLY for the currently active supervisor tab.
 */
function submitCurrentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // Validate that we are on a supervisor sheet
  if (!SUPERVISOR_TABS.includes(sheetName)) {
    SpreadsheetApp.getUi().alert(`Current sheet "${sheetName}" is not in the list of Supervisor tabs. Please switch to a valid tab.`);
    return;
  }

  // Run the logic for just this supervisor
  try {
    runAttendanceLog(sheetName);
    SpreadsheetApp.getUi().alert(`Attendance for ${sheetName} has been successfully logged.`);
  } catch (e) {
    console.error(e);
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

/**
 * Main function to log attendance for ALL supervisors.
 * Can be run via a time-driven trigger as a backup.
 */
function logAllAttendance() {
  try {
    runAttendanceLog(null); // null means process all
    console.log('All attendance logged successfully.');
  } catch (e) {
    console.error('Error in logAllAttendance: ' + e.message);
  }
}

// --- CORE LOGIC ---

/**
 * Orchestrates the logging process.
 * @param {string|null} specificSupervisorName - The name of the specific tab to process, or null to process all.
 */
function runAttendanceLog(specificSupervisorName) {
  const today = new Date();
  console.log(`Starting attendance log for: ${today.toDateString()} ` + (specificSupervisorName ? `[${specificSupervisorName}]` : '[ALL]'));

  // 1. Determine the correct Timecard File
  const mondayDate = getMonday(today);
  const file = findTimecardFile(mondayDate);
  if (!file) {
    throw new Error('Could not find a matching Timecard file for the week starting ' + formatDate(mondayDate));
  }
  console.log(`Found Timecard file: ${file.getName()}`);

  // 2. Open the Timecard Spreadsheet and the correct Daily Tab
  const timecardSS = SpreadsheetApp.open(file);
  const dayTabName = getDayTabName(today);
  const targetSheet = timecardSS.getSheetByName(dayTabName);

  if (!targetSheet) {
    throw new Error(`Could not find tab "${dayTabName}" in file "${file.getName()}".`);
  }
  console.log(`Processing for day tab: ${dayTabName}`);

  // 3. Prepare Target Data Maps (Name -> Row, Job -> Col)
  const { nameRowMap, jobColMap } = buildTargetMaps(targetSheet);

  // 4. Process Source Data
  const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);

  if (specificSupervisorName) {
    // Process single supervisor
    const sheet = sourceSS.getSheetByName(specificSupervisorName);
    if (!sheet) throw new Error(`Sheet "${specificSupervisorName}" not found.`);
    processSupervisor(sheet, targetSheet, nameRowMap, jobColMap);
  } else {
    // Process all supervisors
    SUPERVISOR_TABS.forEach(supervisorName => {
      const sheet = sourceSS.getSheetByName(supervisorName);
      if (sheet) {
        processSupervisor(sheet, targetSheet, nameRowMap, jobColMap);
      } else {
        console.warn(`Supervisor sheet "${supervisorName}" not found.`);
      }
    });
  }
}

/**
 * Reads a supervisor's sheet and updates the target timecard sheet.
 */
function processSupervisor(sourceSheet, targetSheet, nameRowMap, jobColMap) {
  console.log(`Processing supervisor: ${sourceSheet.getName()}`);
  const data = sourceSheet.getDataRange().getValues();

  // Iterate rows (assuming Row 1 is header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Safety check for empty rows
    if (!row[0]) continue;

    const name = String(row[0]).trim(); // Col A
    const jobFunction = String(row[2]).trim(); // Col C
    const present = String(row[5]).trim().toUpperCase(); // Col F

    if (present === 'Y') {
      const targetRow = nameRowMap.get(name.toLowerCase());
      const targetCol = jobColMap.get(jobFunction.toLowerCase());

      if (targetRow && targetCol) {
        targetSheet.getRange(targetRow, targetCol).setValue(HOURS_TO_LOG);
      } else {
        if (!targetRow) console.warn(`Name "${name}" from ${sourceSheet.getName()} not found in Timecard.`);
        if (!targetCol) console.warn(`Job "${jobFunction}" from ${sourceSheet.getName()} not found in Timecard headers.`);
      }
    }
  }
}

/**
 * Builds maps for Names (Row index) and Job Functions (Column index) from the target sheet.
 */
function buildTargetMaps(targetSheet) {
  const targetData = targetSheet.getDataRange().getValues();

  // Map Name -> Row Index
  const nameRowMap = new Map();
  for (let r = 0; r < targetData.length; r++) {
    // Column B is Index 1
    if (targetData[r].length > 1) {
      const name = String(targetData[r][1]).trim();
      if (name) {
        nameRowMap.set(name.toLowerCase(), r + 1); // 1-based row index
      }
    }
  }

  // Map Job Function -> Column Index
  const jobColMap = new Map();
  // Job functions are in Row 3 (Index 2)
  if (targetData.length > 2) {
    const headerRow = targetData[2];
    for (let c = 0; c < headerRow.length; c++) {
      const job = String(headerRow[c]).trim();
      if (job) {
        jobColMap.set(job.toLowerCase(), c + 1); // 1-based column index
      }
    }
  }

  return { nameRowMap, jobColMap };
}


// --- HELPER FUNCTIONS ---

function getMonday(d) {
  const date = new Date(d);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is sunday
  return new Date(date.setDate(diff));
}

function getDayTabName(date) {
  const days = ['Sunday', 'Monday_', 'Tuesday_', 'Wednesday_', 'Thursday_', 'Friday_', 'Saturday'];
  return days[date.getDay()];
}

function findTimecardFile(mondayDate) {
  const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const dateString = formatDate(mondayDate);
  const query = `title contains '191 Week' and title contains '${dateString}' and trashed = false`;
  const files = folder.searchFiles(query);
  if (files.hasNext()) return files.next();
  return null;
}

function formatDate(date) {
  const m = date.getMonth() + 1;
  const d = date.getDate();
  const y = date.getFullYear().toString().slice(-2);
  return `${m}/${d}/${y}`;
}

// --- UI & TRIGGERS ---

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
    .addItem('Submit Current Sheet', 'submitCurrentSheet')
    .addSeparator()
    .addItem('Log All (Admin)', 'logAllAttendance')
    .addToUi();
}

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'logAllAttendance' || trigger.getHandlerFunction() === 'logDailyAttendance') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('logAllAttendance')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  console.log('Daily trigger set for 11 PM.');
}
