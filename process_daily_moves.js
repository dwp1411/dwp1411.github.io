/**
 * This script processes daily "Daily Move Report" files for sites 191 and 199,
 * extracting the total pieces shipped into location "50900101" and logging
 * them into a target tracking spreadsheet under yesterday's date.
 */

function processDailyMoves() {
  const FOLDER_ID = "1hy_IFqn9rHXwp5aYcpXX-tHJ5GZUtwJl";
  const TARGET_SPREADSHEET_ID = "1R0L_etmd77M5xC6hlobr0h5WnQFVkHbSavaPSYCYdRc";
  const TARGET_SHEET_NAME = "Sheet1";
  const TARGET_LOCATION = "50900101";

  // Calculate today's date for filename matching
  const today = new Date();
  const tz = Session.getScriptTimeZone();
  const todayFormatted = Utilities.formatDate(today, tz, "yyyy-MM-dd");

  // Calculate yesterday's date to match the column in the target sheet
  const yesterday = new Date(today.getTime());
  yesterday.setDate(today.getDate() - 1);

  const yestYear = yesterday.getFullYear();
  const yestMonth = yesterday.getMonth();
  const yestDate = yesterday.getDate();

  // Open the target spreadsheet
  const targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const targetSheet = targetSpreadsheet.getSheetByName(TARGET_SHEET_NAME);

  if (!targetSheet) {
    console.error(`Tab "${TARGET_SHEET_NAME}" not found in target spreadsheet.`);
    return;
  }

  // Find the column that contains yesterday's date in Row 1
  const lastCol = targetSheet.getLastColumn();
  if (lastCol < 1) {
    console.error("Target sheet is empty.");
    return;
  }

  const headerValues = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let targetColIndex = -1;

  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = headerValues[i];

    if (cellValue instanceof Date) {
      if (cellValue.getFullYear() === yestYear &&
          cellValue.getMonth() === yestMonth &&
          cellValue.getDate() === yestDate) {
        targetColIndex = i + 1; // 1-based index for Google Sheets
        break;
      }
    } else if (cellValue) {
      // If the date is formatted as text (e.g. "3/1/26")
      const parsedDate = new Date(cellValue);
      if (!isNaN(parsedDate.getTime())) {
        if (parsedDate.getFullYear() === yestYear &&
            parsedDate.getMonth() === yestMonth &&
            parsedDate.getDate() === yestDate) {
          targetColIndex = i + 1;
          break;
        }
      }
    }
  }

  if (targetColIndex === -1) {
    console.error(`Could not find yesterday's date (${Utilities.formatDate(yesterday, tz, "M/d/yy")}) in Row 1 of the target sheet.`);
    return;
  }

  // Access the folder containing the daily reports
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const sites = ["191", "199"];
  const siteSums = { "191": 0, "199": 0 };

  // Process each site
  for (let i = 0; i < sites.length; i++) {
    const site = sites[i];
    const fileName = `whmovrpt${site}_${todayFormatted}`;
    const files = folder.getFilesByName(fileName);

    if (files.hasNext()) {
      const file = files.next();
      let sourceSpreadsheet;

      try {
        sourceSpreadsheet = SpreadsheetApp.open(file);
      } catch (e) {
        console.error(`Error opening file ${fileName}: ${e.message}. Ensuring it's a valid Google Sheet.`);
        continue;
      }

      const sourceSheet = sourceSpreadsheet.getSheets()[0];
      const data = sourceSheet.getDataRange().getValues();
      let sum = 0;

      // Data starts on Row 2 (index 1)
      for (let r = 1; r < data.length; r++) {
        const row = data[r];

        // Ensure row has enough columns: Column X is index 23, Column P is index 15
        if (row.length > 23) {
          const colXValue = row[23];

          // Exact match for the target location
          if (String(colXValue) === TARGET_LOCATION) {
            const pieces = parseFloat(row[15]);
            if (!isNaN(pieces)) {
              sum += pieces;
            }
          }
        }
      }

      siteSums[site] = sum;
      console.log(`Successfully processed ${fileName}. Sum for location ${TARGET_LOCATION}: ${sum}`);
    } else {
      console.warn(`File ${fileName} not found. Defaulting to 0.`);
    }
  }

  // Write the calculated sums to the target sheet
  // Row 2 for 191, Row 3 for 199
  targetSheet.getRange(2, targetColIndex).setValue(siteSums["191"]);
  targetSheet.getRange(3, targetColIndex).setValue(siteSums["199"]);

  console.log(`Successfully updated target sheet for date ${Utilities.formatDate(yesterday, tz, "M/d/yy")}.`);
}

/**
 * Run this function manually ONE TIME to process all historical files
 * currently residing in the Drive folder and backfill the target sheet.
 */
function backfillHistoricalMoves() {
  const FOLDER_ID = "1hy_IFqn9rHXwp5aYcpXX-tHJ5GZUtwJl";
  const TARGET_SPREADSHEET_ID = "1R0L_etmd77M5xC6hlobr0h5WnQFVkHbSavaPSYCYdRc";
  const TARGET_SHEET_NAME = "Sheet1";
  const TARGET_LOCATION = "50900101";

  const targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const targetSheet = targetSpreadsheet.getSheetByName(TARGET_SHEET_NAME);

  if (!targetSheet) {
    console.error(`Tab "${TARGET_SHEET_NAME}" not found in target spreadsheet.`);
    return;
  }

  const lastCol = targetSheet.getLastColumn();
  if (lastCol < 1) {
    console.error("Target sheet is empty.");
    return;
  }

  // Cache the header row mapping dates to column indices
  const headerValues = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap = {};

  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = headerValues[i];
    let d = null;
    if (cellValue instanceof Date) {
      d = cellValue;
    } else if (cellValue) {
      d = new Date(cellValue);
    }

    if (d && !isNaN(d.getTime())) {
      const dateKey = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
      colMap[dateKey] = i + 1; // 1-based index
    }
  }

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    // Check if filename matches expected pattern, e.g. whmovrpt191_2026-05-01
    const match = fileName.match(/whmovrpt(191|199)_(\d{4}-\d{2}-\d{2})/);
    if (!match) continue;

    const site = match[1];
    const fileDateStr = match[2]; // e.g. "2026-05-01"

    // Parse the date from the filename (assumes YYYY-MM-DD local time)
    const fileParts = fileDateStr.split('-');
    const fileDate = new Date(fileParts[0], fileParts[1] - 1, fileParts[2]);

    // The data inside applies to the day before the filename date
    const targetDate = new Date(fileDate.getTime());
    targetDate.setDate(fileDate.getDate() - 1);

    const dateKey = `${targetDate.getFullYear()}-${targetDate.getMonth()}-${targetDate.getDate()}`;
    const targetColIndex = colMap[dateKey];

    if (!targetColIndex) {
      console.warn(`Target column for date ${targetDate.toLocaleDateString()} not found in Row 1. Skipping file ${fileName}.`);
      continue;
    }

    let sourceSpreadsheet;
    try {
      sourceSpreadsheet = SpreadsheetApp.open(file);
    } catch (e) {
      console.error(`Error opening file ${fileName}: ${e.message}. Ensuring it's a valid Google Sheet.`);
      continue;
    }

    const sourceSheet = sourceSpreadsheet.getSheets()[0];
    const data = sourceSheet.getDataRange().getValues();
    let sum = 0;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (row.length > 23) {
        if (String(row[23]) === TARGET_LOCATION) {
          const pieces = parseFloat(row[15]);
          if (!isNaN(pieces)) {
            sum += pieces;
          }
        }
      }
    }

    // Write sum to the correct cell
    const targetRow = (site === "191") ? 2 : 3;
    targetSheet.getRange(targetRow, targetColIndex).setValue(sum);

    console.log(`Backfilled ${fileName}. Written to Row ${targetRow}, Col ${targetColIndex} (Date: ${targetDate.toLocaleDateString()}) - Sum: ${sum}`);
  }

  console.log("Historical backfill completed.");
}
