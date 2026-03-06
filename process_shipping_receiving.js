/**
 * Processes daily Shipping and Receiving Reports.
 * Extracts "Shp" and "Rcv" piece counts from 191 and 199 tabs and logs them to the correct Timecard file.
 */
function processShippingReceivingReport() {
  const SOURCE_FOLDER_ID = "1RYcW4khpFg79Dz_iZTtiemNskSbyjbKU";
  const TIMECARD_FOLDER_ID = "1e0n54IsOYFJja4iPgCbwnn-huhvVBMDo";
  const TARGET_TAB_NAME = "Time Card Data Input File Link";

  // 1. File Selection Logic
  const today = new Date();
  const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
  // Search for the Google Sheet created today containing the date
  // e.g. "WI614_D2_Rpt_2026-03-06_05.46.01"
  const filesIter = sourceFolder.searchFiles(`title contains '${formattedToday}' and mimeType = 'application/vnd.google-apps.spreadsheet'`);

  let targetFile = null;
  while (filesIter.hasNext()) {
    const file = filesIter.next();
    targetFile = file; // Take the first matching file
    break;
  }

  if (!targetFile) {
    throw new Error(`Could not find a Google Sheet in folder ${SOURCE_FOLDER_ID} for today: ${formattedToday}`);
  }

  // 2. Date Calculation
  const fileName = targetFile.getName();

  // Extract the date from the filename (e.g., "WI614_D2_Rpt_2026-03-06_...")
  const dateMatch = fileName.match(/(\d{4}-\d{2}-\d{2})/);
  if (!dateMatch) {
    throw new Error(`Could not extract a date (YYYY-MM-DD) from the filename: ${fileName}`);
  }
  const dateString = dateMatch[1]; // e.g., "2026-03-06"

  // Create a Date object from the filename date string.
  const [year, month, day] = dateString.split('-');
  const generatedDate = new Date(year, parseInt(month, 10) - 1, day);

  const reportDate = new Date(generatedDate.getTime());
  // The data is actually for 2 days ago
  reportDate.setDate(generatedDate.getDate() - 2);

  // Get the Monday of the week for the reportDate.
  const dayOfWeek = reportDate.getDay(); // 0 (Sunday) to 6 (Saturday)

  const diffToMonday = (dayOfWeek === 0 ? -6 : 1 - dayOfWeek);
  const mondayDate = new Date(reportDate.getTime());
  mondayDate.setDate(reportDate.getDate() + diffToMonday);

  // Normalize the times to midnight to ensure accurate matching
  reportDate.setHours(0, 0, 0, 0);
  mondayDate.setHours(0, 0, 0, 0);

  // 3. Data Extraction from Google Sheet
  let rcv199 = 0; // Row 5
  let rcv191 = 0; // Row 6
  let shp191 = 0; // Row 22
  let shp199 = 0; // Row 23

  const sourceSpreadsheet = SpreadsheetApp.openById(targetFile.getId());
  const allSheets = sourceSpreadsheet.getSheets();

  allSheets.forEach(sheet => {
    const tabName = sheet.getName();
    const is191 = tabName.includes("191");
    const is199 = tabName.includes("199");

    if (!is191 && !is199) {
      // Not a target tab, skip
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    // We need Columns A through D.
    // Index 0 is Column A, Index 3 is Column D
    const data = sheet.getRange(1, 1, lastRow, 4).getValues();

    for (let r = 0; r < data.length; r++) {
      const colAValue = data[r][0]; // Index 0 is Column A
      const colDValue = data[r][3]; // Index 3 is Column D

      if (typeof colAValue === 'string') {
        const valUpper = colAValue.trim().toUpperCase();

        // Sum values if Column A matches exactly one of the four types
        // Also ensure Column D contains a numeric value
        const pieces = parseFloat(colDValue);
        if (!isNaN(pieces)) {
          if (valUpper === 'RCV/DOM' || valUpper === 'RCV/IMP') {
            if (is191) rcv191 += pieces;
            if (is199) rcv199 += pieces;
          } else if (valUpper === 'SHP/DOM' || valUpper === 'SHP/IMP') {
            if (is191) shp191 += pieces;
            if (is199) shp199 += pieces;
          }
        }
      }
    }
  });

  // 4. Timecard Update Logic
  // Construct the expected file name part based on the mondayDate
  // Using Utilities.formatDate ensures consistent string formatting matching Google Sheets display
  const targetDateString = `${mondayDate.getMonth() + 1}/${mondayDate.getDate()}/${mondayDate.getFullYear().toString().slice(-2)}`;

  // Search the Timecard folder for files containing the start date string
  const folder = DriveApp.getFolderById(TIMECARD_FOLDER_ID);
  const timecardFilesIter = folder.searchFiles(`title contains '${targetDateString}'`);

  let targetTimecardFile = null;
  while (timecardFilesIter.hasNext()) {
    const file = timecardFilesIter.next();
    // Verify it's exactly the file we want by ensuring it matches the "191 Week" pattern
    if (file.getName().startsWith("191 Week") && file.getName().includes(targetDateString)) {
      targetTimecardFile = file;
      break;
    }
  }

  if (!targetTimecardFile) {
    throw new Error(`Could not find a Timecard file for week starting ${targetDateString} in folder ${TIMECARD_FOLDER_ID}`);
  }

  const timecardSpreadsheet = SpreadsheetApp.openById(targetTimecardFile.getId());
  const inputTab = timecardSpreadsheet.getSheetByName(TARGET_TAB_NAME);

  if (!inputTab) {
    throw new Error(`Could not find the tab '${TARGET_TAB_NAME}' in the Timecard file ${targetTimecardFile.getName()}`);
  }

  // The dates are in Row 3. Get the dates.
  const numColumns = inputTab.getLastColumn();
  if (numColumns === 0) {
    throw new Error(`The '${TARGET_TAB_NAME}' tab is empty.`);
  }

  // Row 3 (Index 3), starting from column 1
  const row3Dates = inputTab.getRange(3, 1, 1, numColumns).getValues()[0];

  let targetColIndex = -1; // Keep track of the 0-based array index

  for (let i = 0; i < row3Dates.length; i++) {
    const cellValue = row3Dates[i];
    if (cellValue instanceof Date) {
      // Create a new Date object to normalize the time
      const cellDate = new Date(cellValue.getTime());
      cellDate.setHours(0, 0, 0, 0);

      if (cellDate.getTime() === reportDate.getTime()) {
        targetColIndex = i;
        break;
      }
    }
  }

  if (targetColIndex === -1) {
    const formattedReportDate = `${reportDate.getMonth() + 1}/${reportDate.getDate()}/${reportDate.getFullYear().toString().slice(-2)}`;
    throw new Error(`Could not find the report date (${formattedReportDate}) in Row 3 of the '${TARGET_TAB_NAME}' tab.`);
  }

  const sheetColIndex = targetColIndex + 1; // 1-based index for Google Sheets range methods

  // Write the totals into the correct rows
  inputTab.getRange(5, sheetColIndex).setValue(rcv199); // 199 Receiving -> Row 5
  inputTab.getRange(6, sheetColIndex).setValue(rcv191); // 191 Receiving -> Row 6
  inputTab.getRange(22, sheetColIndex).setValue(shp191); // 191 Outbound -> Row 22
  inputTab.getRange(23, sheetColIndex).setValue(shp199); // 199 Outbound -> Row 23

  Logger.log(`Successfully processed report for ${formattedToday} (Data Date: ${reportDate.toDateString()}).`);
  Logger.log(`191 Rcv: ${rcv191}, 199 Rcv: ${rcv199}`);
  Logger.log(`191 Shp: ${shp191}, 199 Shp: ${shp199}`);
}
