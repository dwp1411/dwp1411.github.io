/**
 * Processes forwards files and updates Time Card files.
 */
function processForwards() {
  const FORWARDS_FOLDER_ID = '1R5pWBjcsJV_83amg1eCnJlDEtuR7caE4';
  const TIMECARD_FOLDER_ID = '1e0n54IsOYFJja4iPgCbwnn-huhvVBMDo';

  // 1. Find the most recent forwards file
  const forwardsFolder = DriveApp.getFolderById(FORWARDS_FOLDER_ID);
  const files = forwardsFolder.getFiles();
  let mostRecentFile = null;
  let maxDate = 0;

  while (files.hasNext()) {
    const file = files.next();
    const lastUpdated = file.getLastUpdated().getTime();
    if (lastUpdated > maxDate) {
      maxDate = lastUpdated;
      mostRecentFile = file;
    }
  }

  if (!mostRecentFile) {
    throw new Error('No forwards file found in folder: ' + FORWARDS_FOLDER_ID);
  }

  const forwardsFileName = mostRecentFile.getName();
  Logger.log('Most recent forwards file: ' + forwardsFileName);

  // Extract week number: e.g., 'forwards271' -> week 1, 'forwards2710' -> week 10
  // "forwards" is 8 chars, then 2 digits for FY (e.g., "27"). The rest is the week number.
  const weekMatch = forwardsFileName.match(/forwards\d{2}(\d{1,2})/i);
  if (!weekMatch) {
    throw new Error('Could not extract week number from filename: ' + forwardsFileName);
  }
  const weekNum = weekMatch[1];
  Logger.log('Extracted week number: ' + weekNum);

  // 2. Read and aggregate data from the forwards file
  const forwardsSpreadsheet = SpreadsheetApp.openById(mostRecentFile.getId());
  const forwardsSheet = forwardsSpreadsheet.getSheetByName('NAP CRATE');

  if (!forwardsSheet) {
    throw new Error('NAP CRATE tab not found in file: ' + forwardsFileName);
  }

  // Data starts in row 5
  const lastRow = forwardsSheet.getLastRow();
  const aggregatedData = {};

  if (lastRow >= 5) {
    // Column A is index 0, Column D is index 3
    const dataRange = forwardsSheet.getRange(5, 1, lastRow - 4, 4);
    const data = dataRange.getValues();

    for (let i = 0; i < data.length; i++) {
      const dateVal = data[i][0];
      const piecesVal = data[i][3];

      if (dateVal) {
        // Ensure dateVal is a Date object, or convert string to something comparable
        // We'll use the Date's getTime() string representation or format it to be safe
        let dateKey = '';
        if (dateVal instanceof Date) {
           dateVal.setHours(0,0,0,0);
           dateKey = dateVal.getTime().toString();
        } else {
           // Try to parse string date
           const parsedDate = new Date(dateVal);
           if (!isNaN(parsedDate.getTime())) {
             parsedDate.setHours(0,0,0,0);
             dateKey = parsedDate.getTime().toString();
           } else {
             dateKey = String(dateVal).trim();
           }
        }

        // Add only if we successfully parsed the dateKey
        if (dateKey) {
            let pieces = parseFloat(piecesVal);
            if (isNaN(pieces)) {
              pieces = 0;
            }

            if (!aggregatedData[dateKey]) {
              aggregatedData[dateKey] = 0;
            }
            aggregatedData[dateKey] += pieces;
        }
      }
    }
  }
  Logger.log('Aggregated Data: ' + JSON.stringify(aggregatedData));

  // 3. Find the matching Time Card file
  const timecardFolder = DriveApp.getFolderById(TIMECARD_FOLDER_ID);
  const timecardFiles = timecardFolder.getFiles();
  let timecardFile = null;
  const targetPrefix = '191 Week ' + weekNum + ' ';

  while (timecardFiles.hasNext()) {
    const tFile = timecardFiles.next();
    if (tFile.getName().startsWith(targetPrefix)) {
      timecardFile = tFile;
      break;
    }
  }

  if (!timecardFile) {
    throw new Error('Could not find a matching time card file starting with: ' + targetPrefix);
  }

  Logger.log('Found time card file: ' + timecardFile.getName());

  // 4. Write data to the Time Card file
  const timecardSpreadsheet = SpreadsheetApp.openById(timecardFile.getId());
  const timecardSheet = timecardSpreadsheet.getSheetByName('Time Card Data Input File Link');

  if (!timecardSheet) {
    throw new Error('Time Card Data Input File Link tab not found in file: ' + timecardFile.getName());
  }

  // Read row 3 (Columns E to K, indexes 4 to 10) for dates
  // getRange(row, column, numRows, numColumns) -> row 3, col 5 (E), 1 row, 7 cols
  const dateRange = timecardSheet.getRange(3, 5, 1, 7);
  const dateValues = dateRange.getValues()[0];

  const valuesToWrite = [[0, 0, 0, 0, 0, 0, 0]];

  for (let c = 0; c < dateValues.length; c++) {
    const tDateVal = dateValues[c];
    if (tDateVal) {
      let tDateKey = '';
      if (tDateVal instanceof Date) {
        tDateVal.setHours(0,0,0,0);
        tDateKey = tDateVal.getTime().toString();
      } else {
        const parsedTDate = new Date(tDateVal);
        if (!isNaN(parsedTDate.getTime())) {
          parsedTDate.setHours(0,0,0,0);
          tDateKey = parsedTDate.getTime().toString();
        } else {
          tDateKey = String(tDateVal).trim();
        }
      }

      if (tDateKey && aggregatedData.hasOwnProperty(tDateKey)) {
        valuesToWrite[0][c] = aggregatedData[tDateKey];
      }
    }
  }

  // Write the resulting array into row 9, columns E to K
  const writeRange = timecardSheet.getRange(9, 5, 1, 7);
  writeRange.setValues(valuesToWrite);

  Logger.log('Successfully wrote data to: ' + timecardFile.getName());
}

/**
 * INSTRUCTIONS FOR USE:
 *
 * 1. Open your Google Drive, and create a new Google Apps Script project
 *    (or open an existing one where you want this to live).
 * 2. Copy the entire contents of this file and paste it into the script editor.
 * 3. Save the script (Ctrl+S or Cmd+S).
 * 4. To run the script manually:
 *    - Select "processForwards" from the function dropdown at the top.
 *    - Click the "Run" button.
 * 5. If it's your first time running it, Google will ask for authorization
 *    to access your Drive and Spreadsheets. Follow the prompts to allow access.
 * 6. The script will log its progress. You can view the logs by clicking
 *    "Execution log" at the bottom of the editor.
 */
