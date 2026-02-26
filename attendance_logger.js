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
 * 1. **Add a Button for Each Supervisor:**
 *    - Go to each supervisor's tab.
 *    - Insert > Drawing > New.
 *    - Create a shape (e.g., a button labeled "Submit Attendance").
 *    - Click "Save and Close".
 *    - Click the three dots on the new drawing -> "Assign script".
 *    - Enter `submitCurrentSheet` and click OK.
 *
 * HOW TO USE:
 * - When a supervisor is done logging their attendance for the day, they simply click the "Submit Attendance" button on their sheet.
 * - The script will log 7.5 hours for everyone marked 'Y' on that specific sheet.
 * - A confirmation message will appear when done.
 *
 * ADVANCED FEATURES:
 * - **Split Shifts / Multiple Jobs:**
 *    - Add a new column "Hours" (Column G) to your Supervisor sheet.
 *    - If an associate works multiple jobs, duplicate their name on a new row.
 *    - Select the first job for Row 1 and enter the hours (e.g., 4) in Col G.
 *    - Select the second job for Row 2 and enter the hours (e.g., 3.5) in Col G.
 *    - Ensure both rows are marked 'Y' for Present.
 *    - The script will log the specific hours for each job. If Col G is blank, it defaults to 7.5.
 *
 * LOGIC:
 * - Finds the "Monday" of the current week.
 * - Searches Drive for the correct Timecard file.
 * - Opens the correct daily tab (e.g., "Monday_").
 * - Reads attendance from the active sheet.
 * - Logs hours (custom from Col G or default 7.5) for associates marked 'Y'.
 * - **Smart Name Matching:** Attempts to match names even if they are formatted differently (e.g., "First Last" vs "Last First").
 */

// --- CONFIGURATION ---
const SOURCE_SPREADSHEET_ID = '1YQ0uua9CU04SYBSl1MO_kT9UuO8YNYwJyWMm_3tm8GM'; // The Attendance Tracker Sheet ID
const DESTINATION_FOLDER_ID = '12kzfclHUoALkgkB6c2njL5L8u23II25r'; // The Timecards Folder ID
const SUPERVISOR_TABS = [
  'Abel', 'Andrews', 'Casey', 'Chris', 'Dan', 'Dave',
  'Javier', 'Jesse', 'Mercedes', 'Ramon', 'Roger', 'Tombe'
];
const DEFAULT_HOURS = 7.5;
const HOURS_COLUMN_INDEX = 6; // Column G (0-based index)

// --- ENTRY POINTS ---

/**
 * Triggered manually via a button on the active sheet.
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

// --- CORE LOGIC ---

/**
 * Orchestrates the logging process.
 * @param {string|null} specificSupervisorName - The name of the specific tab to process.
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
  }
}

/**
 * Reads a supervisor's sheet and updates the target timecard sheet.
 * Uses an update map to aggregate hours before writing, supporting multiple rows for split shifts.
 */
function processSupervisor(sourceSheet, targetSheet, nameRowMap, jobColMap) {
  console.log(`Processing supervisor: ${sourceSheet.getName()}`);
  const data = sourceSheet.getDataRange().getValues();

  // Map to aggregate updates: Key = "row_col", Value = Total Hours
  const updates = new Map();

  // Iterate rows (assuming Row 1 is header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Safety check for empty rows
    if (!row[0]) continue;

    const rawName = String(row[0]).trim(); // Col A
    const jobFunction = String(row[2]).trim(); // Col C
    const present = String(row[5]).trim().toUpperCase(); // Col F

    if (present === 'Y') {
      const targetRow = findTargetRow(rawName, nameRowMap);
      const targetCol = jobColMap.get(jobFunction.toLowerCase());

      if (targetRow && targetCol) {
        // Determine Hours: Read from Col G (Index 6), otherwise use default
        let hours = DEFAULT_HOURS;
        if (row.length > HOURS_COLUMN_INDEX) {
          const customHours = parseFloat(row[HOURS_COLUMN_INDEX]);
          if (!isNaN(customHours) && customHours > 0) {
            hours = customHours;
          }
        }

        // Accumulate hours for this cell (handling potential multiple rows for same job)
        const key = `${targetRow}_${targetCol}`;
        const currentTotal = updates.get(key) || 0;
        updates.set(key, currentTotal + hours);

      } else {
        if (!targetRow) console.warn(`Name "${rawName}" from ${sourceSheet.getName()} not found in Timecard (tried reversing name too).`);
        if (!targetCol) console.warn(`Job "${jobFunction}" from ${sourceSheet.getName()} not found in Timecard headers.`);
      }
    }
  }

  // Apply updates to the sheet
  if (updates.size > 0) {
    console.log(`Writing ${updates.size} updates to Timecard.`);
    updates.forEach((hours, key) => {
      const [r, c] = key.split('_').map(Number);
      targetSheet.getRange(r, c).setValue(hours);
    });
  }
}

/**
 * Tries to find the row index for a given name, checking both:
 * 1. Exact match (case-insensitive)
 * 2. Reversed name match (Last First <-> First Last)
 */
function findTargetRow(rawName, nameRowMap) {
  const name = rawName.toLowerCase();

  // 1. Direct match
  if (nameRowMap.has(name)) {
    return nameRowMap.get(name);
  }

  // 2. Try swapping parts (e.g. "Mercedez Rodriguez" <-> "Rodriguez Mercedez")
  // Split by space
  const parts = name.split(/\s+/);
  if (parts.length >= 2) {
    // Reverse the parts and join them back
    const reversedName = parts.reverse().join(' ');
    if (nameRowMap.has(reversedName)) {
      console.log(`Fuzzy match found: "${rawName}" matched as "${reversedName}"`);
      return nameRowMap.get(reversedName);
    }
  }

  return null;
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
