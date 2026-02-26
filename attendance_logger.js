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
 * 1. Run the `setupTriggers` function once to create a daily trigger (e.g., run every night or every hour).
 *    - Alternatively, you can set up triggers manually in the Triggers dashboard (clock icon on the left).
 * 2. Run `onOpen` to add the custom menu "Attendance" -> "Log Today's Hours" for manual execution.
 *
 * LOGIC:
 * - Finds the "Monday" of the current week based on the current date.
 * - Searches the destination Google Drive folder for a file named like "191 Week ... [Monday Date]".
 * - Opens the correct daily tab in that file (e.g., "Monday_", "Tuesday_", "Wednesday_", "Thursday_", "Friday_", "Saturday", "Sunday").
 * - Reads attendance data from each Supervisor's tab in the source sheet.
 * - If an associate is marked 'Y' in column F, it logs 7.5 hours to their row (Name in Col B) and job function column (Row 3).
 */

// --- CONFIGURATION ---
const SOURCE_SPREADSHEET_ID = '1YQ0uua9CU04SYBSl1MO_kT9UuO8YNYwJyWMm_3tm8GM'; // The Attendance Tracker Sheet ID
const DESTINATION_FOLDER_ID = '12kzfclHUoALkgkB6c2njL5L8u23II25r'; // The Timecards Folder ID
const SUPERVISOR_TABS = [
  'Abel', 'Andrews', 'Casey', 'Chris', 'Dan', 'Dave',
  'Javier', 'Jesse', 'Mercedes', 'Ramon', 'Roger', 'Tombe'
];
const HOURS_TO_LOG = 7.5;

// --- MAIN FUNCTION ---

/**
 * Main function to log attendance for the current day.
 * Can be run manually or via a time-driven trigger.
 */
function logDailyAttendance() {
  const today = new Date();
  console.log(`Starting attendance log for: ${today.toDateString()}`);

  // 1. Determine the correct Timecard File
  const mondayDate = getMonday(today);
  const file = findTimecardFile(mondayDate);
  if (!file) {
    console.error('Could not find a matching Timecard file for this week.');
    SpreadsheetApp.getUi().alert('Error: Could not find a matching Timecard file for the week starting ' + formatDate(mondayDate));
    return;
  }
  console.log(`Found Timecard file: ${file.getName()}`);

  // 2. Open the Timecard Spreadsheet and the correct Daily Tab
  const timecardSS = SpreadsheetApp.open(file);
  const dayTabName = getDayTabName(today);
  const targetSheet = timecardSS.getSheetByName(dayTabName);

  if (!targetSheet) {
    console.error(`Could not find tab "${dayTabName}" in file "${file.getName()}".`);
    return;
  }
  console.log(`Processing for day tab: ${dayTabName}`);

  // 3. Prepare Target Data Maps (Name -> Row, Job -> Col)
  // We read the whole sheet data to build maps for fast lookups.
  // Assuming Names are in Column B (Index 1) and Job Functions in Row 3 (Index 2).
  const targetData = targetSheet.getDataRange().getValues();

  // Map Name -> Row Index
  const nameRowMap = new Map();
  for (let r = 0; r < targetData.length; r++) {
    const name = String(targetData[r][1]).trim(); // Column B
    if (name) {
      nameRowMap.set(name.toLowerCase(), r + 1); // Store 1-based row index
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
        jobColMap.set(job.toLowerCase(), c + 1); // Store 1-based column index
      }
    }
  }

  // 4. Process Source Data
  const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);

  SUPERVISOR_TABS.forEach(supervisorName => {
    const sheet = sourceSS.getSheetByName(supervisorName);
    if (!sheet) {
      console.warn(`Supervisor sheet "${supervisorName}" not found.`);
      return;
    }

    console.log(`Processing supervisor: ${supervisorName}`);
    const data = sheet.getDataRange().getValues();

    // Iterate rows (skip header if necessary, assuming Row 1 is header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = String(row[0]).trim(); // Col A
      const jobFunction = String(row[2]).trim(); // Col C
      const present = String(row[5]).trim().toUpperCase(); // Col F

      if (present === 'Y') {
        // Find target coordinates
        const targetRow = nameRowMap.get(name.toLowerCase());
        const targetCol = jobColMap.get(jobFunction.toLowerCase());

        if (targetRow && targetCol) {
          // Check if value already exists to avoid overwriting with same (optional)
          // But here we just write.
          targetSheet.getRange(targetRow, targetCol).setValue(HOURS_TO_LOG);
        } else {
          if (!targetRow) console.warn(`Name "${name}" from ${supervisorName} not found in Timecard.`);
          if (!targetCol) console.warn(`Job "${jobFunction}" from ${supervisorName} not found in Timecard headers.`);
        }
      }
    }
  });

  console.log('Attendance logging complete.');
}

// --- HELPER FUNCTIONS ---

/**
 * Returns the Monday of the week for the given date.
 * If today is Sunday, it returns the previous Monday (standard business week logic).
 * Adjust if your week starts differently.
 */
function getMonday(d) {
  const date = new Date(d);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is sunday
  return new Date(date.setDate(diff));
}

/**
 * Returns the tab name for the given date.
 * Logic: Monday_, Tuesday_, Wednesday_, Thursday_, Friday_, Saturday, Sunday.
 */
function getDayTabName(date) {
  const days = ['Sunday', 'Monday_', 'Tuesday_', 'Wednesday_', 'Thursday_', 'Friday_', 'Saturday'];
  return days[date.getDay()];
}

/**
 * Searches the destination folder for a file containing "191 Week" and the formatted Monday date.
 */
function findTimecardFile(mondayDate) {
  const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const dateString = formatDate(mondayDate); // e.g., "2/23/26"

  // Search for file with title containing "191 Week" and the date string
  // Note: Date formats in titles can be tricky (e.g. 02/23 vs 2/23). We try exact match first.
  const query = `title contains '191 Week' and title contains '${dateString}' and trashed = false`;
  const files = folder.searchFiles(query);

  if (files.hasNext()) {
    return files.next();
  }
  return null;
}

/**
 * Formats date as M/d/yy (e.g., 2/23/26 or 10/5/25).
 * No leading zeros for single digits to match user example "2/23/26".
 */
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
    .addItem('Log Today\'s Hours', 'logDailyAttendance')
    .addToUi();
}

function setupTriggers() {
  // Deletes existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'logDailyAttendance') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create a new daily trigger (e.g., at 11 PM)
  ScriptApp.newTrigger('logDailyAttendance')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  console.log('Daily trigger set for 11 PM.');
}
