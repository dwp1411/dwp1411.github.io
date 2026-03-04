/**
 * Saves the "Forwards" attachment and converts it to a Google Sheet.
 * Fixed "Bad Request" error by specifying the media type.
 */
function saveForwardsAttachment() {
  const FOLDER_ID = '1R5pWBjcsJV_83amg1eCnJlDEtuR7caE4';
  const SENDER = 'forwardsshippedreport@crateandbarrel.com';

  try {
    // Search for the most recent email with an attachment from the sender within the last day
    const query = `from:${SENDER} has:attachment newer_than:2d`;
    const threads = GmailApp.search(query, 0, 1);

    if (threads.length === 0) {
      console.log("No new emails found recently.");
      return;
    }

    const messages = threads[0].getMessages();
    const lastMessage = messages[messages.length - 1];
    const attachments = lastMessage.getAttachments();

    if (attachments.length > 0) {
      const attachment = attachments[0];
      const fileName = attachment.getName().toLowerCase();
      // Keep original attachment name as requested, but we can ensure it doesn't duplicate.
      // E.g., forwards2652.xls
      const newFileName = attachment.getName();

      // Check if folder is accessible
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const existingFiles = folder.getFilesByName(newFileName);

      // If we want to check for a file without the extension (which Google Sheets does)
      const baseFileName = newFileName.replace(/\.[^/.]+$/, "");
      const existingConvertedFiles = folder.getFilesByName(baseFileName);

      if (existingFiles.hasNext() || existingConvertedFiles.hasNext()) {
        console.log("File already exists: " + newFileName);
        return;
      }

      // Get the file content as a blob
      const blob = attachment.copyBlob();

      // FIX: Check the file extension and explicitly set the Blob's content type.
      // Automated systems often send files as 'application/octet-stream' which Drive API cannot convert.
      if (fileName.endsWith('.xlsx')) {
        blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      } else if (fileName.endsWith('.xls')) {
        blob.setContentType('application/vnd.ms-excel');
      } else if (fileName.endsWith('.csv')) {
        blob.setContentType('text/csv');
      }

      // Metadata for the Google Sheet
      const fileMetadata = {
        name: baseFileName, // Name of the Google Sheet (without .xls extension)
        parents: [FOLDER_ID],
        mimeType: 'application/vnd.google-apps.spreadsheet'
      };

      // EXECUTE CONVERSION
      // We pass the configured blob directly to ensure the Drive API recognizes the data
      Drive.Files.create(fileMetadata, blob);

      console.log("Success! File saved as Google Sheet: " + baseFileName);
    }

  } catch (e) {
    console.error("Error Detail: " + e.message);
    // Ensure "Drive API" is added in the 'Services' tab on the left.
  }
}
