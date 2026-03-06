/**
 * Processes daily EDI-Parcel Details Reports.
 * Extracts "DSS" and "SS/SL" package counts and logs them to the correct Timecard file.
 */
function processParcelDetailsReport() {
  const SENDER_EMAIL = "mistechops@crateandbarrel.com";
  const EMAIL_SUBJECT = "EDI-Parcel Details Report";

  // 1. Email and Attachment Retrieval Logic
  // Search for emails from today, sorted so we can pick the earliest one
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy/MM/dd");
  const query = `from:${SENDER_EMAIL} subject:"${EMAIL_SUBJECT}" after:${formattedDate}`;

  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    throw new Error(`No emails found matching query: ${query}`);
  }

  // Get all messages from the threads found today
  let allMessages = [];
  threads.forEach(thread => {
    allMessages = allMessages.concat(thread.getMessages());
  });

  // Filter messages specifically received today (after midnight)
  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  const todayMessages = allMessages.filter(msg => msg.getDate().getTime() >= todayStart.getTime());

  if (todayMessages.length === 0) {
    throw new Error(`Found threads, but no messages received today after midnight.`);
  }

  // Sort messages by date (ascending) to get the earliest one
  todayMessages.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

  const earliestMessage = todayMessages[0];
  const attachments = earliestMessage.getAttachments();

  let targetAttachment = null;
  for (let i = 0; i < attachments.length; i++) {
    if (attachments[i].getName().endsWith('.xls')) {
      targetAttachment = attachments[i];
      break;
    }
  }

  if (!targetAttachment) {
    throw new Error("No .xls attachment found in the earliest email.");
  }

  // 2. Date Calculation
  const attachmentName = targetAttachment.getName();

  // Extract the date from the filename (e.g., "ParcelDetailReport_2026-03-06.xls")
  const dateMatch = attachmentName.match(/(\d{4}-\d{2}-\d{2})/);
  if (!dateMatch) {
    throw new Error(`Could not extract a date (YYYY-MM-DD) from the filename: ${attachmentName}`);
  }
  const dateString = dateMatch[1]; // e.g., "2026-03-06"

  // Create a Date object from the filename date string.
  // The filename is the date the report was GENERATED.
  // The "report date" is the date the shipping actually happened (the day BEFORE).
  const [year, month, day] = dateString.split('-');
  const generatedDate = new Date(year, parseInt(month, 10) - 1, day);

  const reportDate = new Date(generatedDate.getTime());
  reportDate.setDate(generatedDate.getDate() - 1); // Subtract one day to get the actual report date

  // Get the Monday of the week for the reportDate.
  // This helps us locate the correct Timecard file since they start on Mondays.
  const dayOfWeek = reportDate.getDay(); // 0 (Sunday) to 6 (Saturday)

  const diffToMonday = (dayOfWeek === 0 ? -6 : 1 - dayOfWeek);
  const mondayDate = new Date(reportDate.getTime());
  mondayDate.setDate(reportDate.getDate() + diffToMonday);

  // Normalize the times to midnight to ensure accurate matching
  reportDate.setHours(0, 0, 0, 0);
  mondayDate.setHours(0, 0, 0, 0);

  // 3. Data Extraction from .xls
  // Temporarily convert the .xls file to a Google Sheet
  const blob = targetAttachment.copyBlob();
  blob.setContentType(MimeType.MICROSOFT_EXCEL);

  const tempFileConfig = {
    title: attachmentName.replace('.xls', ''),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const tempFile = Drive.Files.insert(tempFileConfig, blob);
  const tempSpreadsheetId = tempFile.id;

  let dssCount = 0;
  let ssslCount = 0;

  try {
    const tempSpreadsheet = SpreadsheetApp.openById(tempSpreadsheetId);

    // Process tabs "191" and "199"
    const tabsToProcess = ["191", "199"];

    tabsToProcess.forEach(tabName => {
      const sheet = tempSpreadsheet.getSheetByName(tabName);
      if (!sheet) {
        // Tab not found, safely skip it
        return;
      }

      const lastRow = sheet.getLastRow();
      if (lastRow === 0) return; // Empty sheet

      // We need Column J (Index 10), so we fetch A to J.
      // Column J's array index is 9.
      const data = sheet.getRange(1, 1, lastRow, 10).getValues();

      for (let r = 0; r < data.length; r++) {
        const colJValue = data[r][9]; // Index 9 is Column J

        if (typeof colJValue === 'string') {
          const valUpper = colJValue.trim().toUpperCase();
          if (valUpper === 'DSS') {
            dssCount++;
          } else if (valUpper === 'SS' || valUpper === 'SL') {
            ssslCount++;
          }
        }
      }
    });

  } catch (error) {
    // If an error occurs, we still want to clean up the temporary file, so we re-throw it.
    throw error;
  } finally {
    // Clean up: delete the temporary Google Sheet
    Drive.Files.remove(tempSpreadsheetId);
  }

  // 4. Timecard Update Logic
  const timecardFolderId = "1e0n54IsOYFJja4iPgCbwnn-huhvVBMDo";
  const targetTabName = "Time Card Data Input File Link";

  // Construct the expected file name part based on the mondayDate
  // e.g. "191 Week [WeekNum] M/D/YY" - we just search for the start date portion.
  // Using Utilities.formatDate ensures consistent string formatting matching Google Sheets display
  const targetDateString = `${mondayDate.getMonth() + 1}/${mondayDate.getDate()}/${mondayDate.getFullYear().toString().slice(-2)}`;

  // Search the Timecard folder for files containing the start date string
  const folder = DriveApp.getFolderById(timecardFolderId);
  const filesIter = folder.searchFiles(`title contains '${targetDateString}'`);

  let targetTimecardFile = null;
  while (filesIter.hasNext()) {
    const file = filesIter.next();
    // Verify it's exactly the file we want by ensuring it matches the "191 Week" pattern
    if (file.getName().startsWith("191 Week") && file.getName().includes(targetDateString)) {
      targetTimecardFile = file;
      break;
    }
  }

  if (!targetTimecardFile) {
    throw new Error(`Could not find a Timecard file for week starting ${targetDateString} in folder ${timecardFolderId}`);
  }

  const timecardSpreadsheet = SpreadsheetApp.openById(targetTimecardFile.getId());
  const inputTab = timecardSpreadsheet.getSheetByName(targetTabName);

  if (!inputTab) {
    throw new Error(`Could not find the tab '${targetTabName}' in the Timecard file ${targetTimecardFile.getName()}`);
  }

  // The dates are in Row 3. Get the dates.
  const numColumns = inputTab.getLastColumn();
  if (numColumns === 0) {
    throw new Error(`The '${targetTabName}' tab is empty.`);
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
    throw new Error(`Could not find the report date (${formattedReportDate}) in Row 3 of the '${targetTabName}' tab.`);
  }

  const sheetColIndex = targetColIndex + 1; // 1-based index for Google Sheets range methods

  // Enter DSS count in Rows 19, 24
  // Enter SS/SL count in Rows 20, 25
  inputTab.getRange(19, sheetColIndex).setValue(dssCount);
  inputTab.getRange(24, sheetColIndex).setValue(dssCount);

  inputTab.getRange(20, sheetColIndex).setValue(ssslCount);
  inputTab.getRange(25, sheetColIndex).setValue(ssslCount);

  Logger.log(`Successfully processed report for ${targetDateString}. DSS count: ${dssCount}, SS/SL count: ${ssslCount}`);
}
