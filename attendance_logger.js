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
 * - Logs hours (custom from Col I/M or default 7.5) for associates marked 'Y'.
 * - Logs FortHill hours to Column AZ if applicable.
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
// Column Indices (0-based: A=0, B=1, C=2...)
const COL_NAME = 0;      // Column A
const COL_JOB_1 = 2;     // Column C
const COL_BUILDING = 3;  // Column D
const COL_PRESENT = 7;   // Column H (Present for Work Y/N)
const COL_HOURS_1 = 8;   // Column I (Updated/Hours)
const COL_JOB_2 = 11;    // Column L (Avoiding button in K)
const COL_HOURS_2 = 12;  // Column M

// Target Column for FortHill Hours (AZ = Index 52 in 1-based getRange)
const TARGET_COL_FORTHILL_INDEX = 52;

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

  // --- POP-UP MENU 1: PROCESSING INDICATOR ---
  // This shows a small toast message at the bottom right to say it started.
  ss.toast('Logging attendance... please wait.', 'Processing', 10);

  // Run the logic for just this supervisor
  try {
    runAttendanceLog(sheetName);

    // --- POP-UP MENU 2: SUCCESS MESSAGE ---
    // This is the main alert box that pops up when done.
    SpreadsheetApp.getUi().alert(`Attendance for ${sheetName} has been successfully logged.`);

  } catch (e) {
    console.error(e);

    // --- POP-UP MENU 3: ERROR MESSAGE ---
    // This alert box pops up if something goes wrong.
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
 * Uses an update map to aggregate hours before writing, supporting multiple jobs per row.
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
    if (!row[COL_NAME]) continue;

    const rawName = String(row[COL_NAME]).trim();
    const building = String(row[COL_BUILDING]).trim(); // Col D
    const present = String(row[COL_PRESENT]).trim().toUpperCase();

    if (present === 'Y') {
      const targetRow = findTargetRow(rawName, nameRowMap);

      if (!targetRow) {
        console.warn(`Name "${rawName}" from ${sourceSheet.getName()} not found in Timecard (tried reversing name too).`);
        continue;
      }

      let totalRowHours = 0;

      // --- Process Job 1 ---
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
        addUpdate(updates, targetRow, targetCol1, hours1);
        totalRowHours += hours1;
      } else {
        console.warn(`Job 1 "${job1}" for "${rawName}" not found in Timecard headers.`);
      }

      // --- Process Job 2 (Optional) ---
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
              addUpdate(updates, targetRow, targetCol2, hours2);
              totalRowHours += hours2;
            }
          } else {
             console.warn(`Job 2 "${job2}" for "${rawName}" not found in Timecard headers.`);
          }
        }
      }

      // --- Process FortHill Hours (Column AZ) ---
      // Requirement: "Making sure that the hours will be logged in the job function... and ALSO add them to column AZ"
      // If Building is "FortHill" (case-insensitive check), add total hours to Col AZ
      if (building.toLowerCase().includes('forthill') && totalRowHours > 0) {
        // This is an ADDITIVE log. If they worked 7.5 hours at FortHill,
        // 7.5 goes to their Job Column AND 7.5 goes to Column AZ.
        addUpdate(updates, targetRow, TARGET_COL_FORTHILL_INDEX, totalRowHours);
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
 * Helper to add hours to the update map, summing if key exists.
 */
function addUpdate(updates, r, c, hours) {
  const key = `${r}_${c}`;
  const currentTotal = updates.get(key) || 0;
  updates.set(key, currentTotal + hours);
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
