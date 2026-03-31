/**
 * Saves the daily shipping/receiving report attachment and converts it to a Google Sheet.
 * Specifically looks for reports sent by mis-warehousing@crateandbarrel.com containing '05.45.' or '05.46.' and ending in '.xls'.
 *
 * IMPORTANT: To use this script, you MUST enable the Drive API v3:
 * 1. Open your Apps Script editor.
 * 2. On the left sidebar, click the plus (+) button next to "Services".
 * 3. Select "Drive API" and click "Add".
 */
function saveShippingReceivingReport() {
  const FOLDER_ID = '1RYcW4khpFg79Dz_iZTtiemNskSbyjbKU';
  const SENDER = 'mis-warehousing@crateandbarrel.com';

  try {
    // Search for emails from the sender within the last day that have attachments
    const query = `from:${SENDER} has:attachment newer_than:1d`;
    const threads = GmailApp.search(query);

    if (threads.length === 0) {
      console.log("No new emails found recently.");
      return;
    }

    let reportFound = false;

    // Iterate through recent threads and messages
    for (let i = 0; i < threads.length; i++) {
      const messages = threads[i].getMessages();

      for (let j = 0; j < messages.length; j++) {
        const message = messages[j];
        const attachments = message.getAttachments();

        for (let k = 0; k < attachments.length; k++) {
          const attachment = attachments[k];
          const fileName = attachment.getName();

          // Look for .xls extension and either '05.45.' or '05.46.' substring in the filename
          if (fileName.toLowerCase().endsWith('.xls') && (fileName.includes('05.45.') || fileName.includes('05.46.'))) {
            console.log(`Found matching report: ${fileName}`);

            // Get base name to check for existing converted Google Sheets
            const baseFileName = fileName.replace(/\.[^/.]+$/, "");
            const folder = DriveApp.getFolderById(FOLDER_ID);

            // Check if file already exists in the destination folder
            const existingConvertedFiles = folder.getFilesByName(baseFileName);

            if (existingConvertedFiles.hasNext()) {
              console.log(`File already exists in folder: ${baseFileName}`);
              reportFound = true; // Mark as found so we don't process it again
              continue; // Skip this one, maybe there are others
            }

            // Get the file content as a blob
            const blob = attachment.copyBlob();

            // Explicitly set the Blob's content type to prevent 'Bad Request' errors
            // Automated systems often send files as 'application/octet-stream'
            blob.setContentType('application/vnd.ms-excel');

            // Metadata for the Google Sheet
            const fileMetadata = {
              name: baseFileName, // Name of the Google Sheet (without .xls extension)
              parents: [FOLDER_ID],
              mimeType: 'application/vnd.google-apps.spreadsheet'
            };

            // Convert to Google Sheet using Drive API
            Drive.Files.create(fileMetadata, blob);

            console.log(`Success! File saved as Google Sheet: ${baseFileName}`);
            reportFound = true;
          }
        }
      }
    }

    if (!reportFound) {
      console.log("No matching '05.45.' or '05.46.' reports found in the recent emails.");
    }

  } catch (e) {
    console.error("Error Detail: " + e.message);
    console.error("Please ensure 'Drive API' is added in the 'Services' tab on the left sidebar.");
  }
}
