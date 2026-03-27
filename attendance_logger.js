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
 * - **Silent Mode:** A small message will appear at the bottom right to confirm start/finish. No pop-up box will interrupt you.
 *
 * ADVANCED FEATURES:
 * - **Split Shifts / Multiple Jobs (WITHOUT Duplicating Rows):**
 *    - Since you cannot duplicate rows, simply add new columns to the right of your existing data.
 *    - **Column I:** "Hours 1" (Overrides the default 7.5 hours for the main job).
 *    - **Column L:** "Job 2" (Select the second job function - Column L avoids the button in Column K).
 *    - **Column M:** "Hours 2" (Enter the hours for the second job).
 * - **FortHill Hours:**
 *    - If Column D (Building) is set to "FortHill", the total hours for that associate will also be logged in Column AZ of the Timecard.
 *
 * LOGIC:
 * - Finds the "Monday" of the current week.
 * - Searches Drive for the correct Timecard file.
 * - Opens the correct daily tab (e.g., "Monday_").
 * - Reads attendance from the active sheet.
 * - Clears the supervisor's designated block in the Timecard to remove old roster data.
 * - Writes the names of associates marked 'Y' into Column B starting at the supervisor's assigned row.
 * - Logs hours (custom from Col I/M or default 7.5) into the appropriate Job columns.
 * - Logs FortHill hours to Column AZ if applicable.
 */

// --- CONFIGURATION ---
const SOURCE_SPREADSHEET_ID = '1YQ0uua9CU04SYBSl1MO_kT9UuO8YNYwJyWMm_3tm8GM'; // The Attendance Tracker Sheet ID
const DESTINATION_FOLDER_ID = '1e0n54IsOYFJja4iPgCbwnn-huhvVBMDo'; // The Timecards Folder ID

// Define the exact rows designated for each supervisor on the Timecard sheet.
// Note: Dan is missing from the list provided, so he is excluded here unless added later.
const SUPERVISOR_RANGES = {
  'Andrews':  { start: 17, end: 66 },
  'Casey':    { start: 67, end: 114 },
  'Chris':    { start: 115, end: 177 },
  'Dave':     { start: 178, end: 229 },
  'Javier':   { start: 230, end: 275 },
  'Jesse':    { start: 276, end: 355 },
  'Mercedes': { start: 356, end: 405 },
  'Ramon':    { start: 406, end: 469 },
  'Roger':    { start: 470, end: 510 },
  'Tombe':    { start: 511, end: 517 },
  'Abel':     { start: 518, end: 524 }
};

const DEFAULT_HOURS = 7.5;
// Column Indices (0-based: A=0, B=1, C=2...)
const COL_NAME = 0;      // Column A
const COL_JOB_1 = 2;     // Column C
const COL_BUILDING = 3;  // Column D
const COL_PRESENT = 5;   // Column F (Present for Work Y/N) -- FIXED TO MATCH USER DESCRIPTION "Column F"
const COL_HOURS_1 = 8;   // Column I (Updated/Hours)
const COL_JOB_2 = 11;    // Column L (Avoiding button in K)
const COL_HOURS_2 = 12;  // Column M

// Target Columns
const TARGET_COL_NAME_INDEX = 2; // Column B (1-based for getRange)
const TARGET_COL_FORTHILL_INDEX = 52; // Column AZ (1-based for getRange)

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
  if (!SUPERVISOR_RANGES[sheetName]) {
    SpreadsheetApp.getUi().alert(`Current sheet "${sheetName}" is not a recognized Supervisor tab with assigned rows. Please check configuration.`);
    return;
  }

  // --- POP-UP MENU 1: PROCESSING INDICATOR ---
  // This shows a small toast message at the bottom right to say it started.
  ss.toast('Logging attendance... please wait.', 'Processing', 10);

  // Run the logic for just this supervisor
  try {
    runAttendanceLog(sheetName);

    // --- POP-UP MENU 2: SUCCESS MESSAGE (SILENT MODE) ---
    // Changed from alert to toast as requested. No user click required.
    ss.toast(`Attendance for ${sheetName} has been successfully logged.`, 'Success', 5);

  } catch (e) {
    console.error(e);

    // --- POP-UP MENU 3: ERROR MESSAGE ---
    // We keep the alert for errors because they are critical and need user attention.
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

// --- CORE LOGIC ---

/**
 * Orchestrates the logging process.
 * @param {string} supervisorName - The name of the specific tab to process.
 */
function runAttendanceLog(supervisorName) {
  const today = new Date();
  console.log(`Starting attendance log for: ${today.toDateString()} [${supervisorName}]`);

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

  // 3. Prepare Target Job Function Headers
  const jobColMap = buildJobHeadersMap(targetSheet);

  // 4. Process Source Data
  const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const sheet = sourceSS.getSheetByName(supervisorName);
  if (!sheet) throw new Error(`Sheet "${supervisorName}" not found.`);

  processSupervisor(sheet, targetSheet, jobColMap, supervisorName);
}

/**
 * Reads a supervisor's sheet and writes names and hours to the target timecard sheet block.
 */
function processSupervisor(sourceSheet, targetSheet, jobColMap, supervisorName) {
  console.log(`Processing supervisor: ${supervisorName}`);
  const data = sourceSheet.getDataRange().getValues();
  const rangeInfo = SUPERVISOR_RANGES[supervisorName];

  // 1. PREPARE BLOCK DATA
  const numRowsToClear = rangeInfo.end - rangeInfo.start + 1;
  const maxCols = Math.max(targetSheet.getLastColumn(), TARGET_COL_FORTHILL_INDEX);

  // Create a 2D array filled with empty strings for the entire block (starting from Col B)
  // Number of columns = maxCols - 2 + 1 = maxCols - 1
  const blockData = Array.from({ length: numRowsToClear }, () => Array(maxCols - 1).fill(''));

  let currentTargetRowIndex = 0; // 0-based index for blockData array

  // Iterate source rows (assuming Row 1 is header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Safety check for empty rows
    if (!row[COL_NAME]) continue;

    const rawName = String(row[COL_NAME]).trim();
    const building = String(row[COL_BUILDING]).trim(); // Col D
    const present = String(row[COL_PRESENT]).trim().toUpperCase(); // Col F

    if (present === 'Y') {

      // Check if we have exceeded the allocated space for this supervisor
      if (currentTargetRowIndex >= numRowsToClear) {
        throw new Error(`Out of space for ${supervisorName}. Tried to write more rows than the limit of ${rangeInfo.end}. Please ask an Admin to increase row allocation in the script.`);
      }

      // 1. Write Name to Column B (Index 0 in blockData since it starts at Col B)
      blockData[currentTargetRowIndex][0] = rawName;

      let totalRowHours = 0;

      // 2. Process Job 1
      const job1 = String(row[COL_JOB_1]).trim();
      const targetCol1 = jobColMap.get(job1.toLowerCase());

      if (targetCol1) {
        let hours1 = DEFAULT_HOURS;
        // Check custom hours for Job 1
        if (row.length > COL_HOURS_1) {
          const customHours = parseFloat(row[COL_HOURS_1]);
          if (!isNaN(customHours) && customHours > 0) {
            hours1 = customHours;
          }
        }

        const colIndex = targetCol1 - 2; // targetCol1 is 1-based, blockData starts at Col 2
        blockData[currentTargetRowIndex][colIndex] = hours1;
        totalRowHours += hours1;
      } else {
        console.warn(`Job 1 "${job1}" for "${rawName}" not found in Timecard headers.`);
      }

      // 3. Process Job 2 (Optional)
      if (row.length > COL_JOB_2) {
        const job2 = String(row[COL_JOB_2]).trim();
        if (job2) {
          const targetCol2 = jobColMap.get(job2.toLowerCase());
          if (targetCol2) {
            let hours2 = 0;
            // Check custom hours for Job 2
            if (row.length > COL_HOURS_2) {
              const customHours2 = parseFloat(row[COL_HOURS_2]);
              if (!isNaN(customHours2) && customHours2 > 0) {
                hours2 = customHours2;
              }
            }
            if (hours2 > 0) {
              const colIndex = targetCol2 - 2;
              const currentVal = Number(blockData[currentTargetRowIndex][colIndex]) || 0;
              blockData[currentTargetRowIndex][colIndex] = currentVal + hours2;
              totalRowHours += hours2;
            }
          } else {
             console.warn(`Job 2 "${job2}" for "${rawName}" not found in Timecard headers.`);
          }
        }
      }

      // 4. Process FortHill Hours (Column AZ)
      if (building.toLowerCase().includes('forthill') && totalRowHours > 0) {
        const colIndex = TARGET_COL_FORTHILL_INDEX - 2;
        const currentVal = Number(blockData[currentTargetRowIndex][colIndex]) || 0;
        blockData[currentTargetRowIndex][colIndex] = currentVal + totalRowHours;
      }

      // Increment row pointer for the next present employee
      currentTargetRowIndex++;
    }
  }

  // 2. APPLY UPDATES TO SHEET
  // Write the entire block in one operation to overwrite old data and set new data instantly
  console.log(`Writing data for ${supervisorName} to Timecard block.`);
  targetSheet.getRange(rangeInfo.start, 2, numRowsToClear, maxCols - 1).setValues(blockData);
}

/**
 * Builds a map for Job Functions (Column index) from the target sheet (Row 3).
 * We no longer need the Name map since we are writing names dynamically.
 */
function buildJobHeadersMap(targetSheet) {
  const targetData = targetSheet.getDataRange().getValues();
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
  return jobColMap;
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
