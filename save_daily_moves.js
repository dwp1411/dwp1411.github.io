/**
 * Searches Gmail for Daily Move Reports, extracts the CSV attachments,
 * converts them to Google Sheets, and saves them to a specific Drive folder.
 * Labels processed emails to avoid duplicate processing.
 * Only looks back 7 days and checks if the file already exists in the folder.
 *
 * NOTE: Before running this script, you MUST enable the Drive API in the
 * Advanced Google Services (Services -> Add a service -> Drive API).
 */
function processDailyMoveReports() {
  var folderId = "1hy_IFqn9rHXwp5aYcpXX-tHJ5GZUtwJl";
  var labelName = "Processed Moves";

  // Search for emails containing "Daily Move Report for 191" or "199" in the subject
  // Only look at emails from the past 7 days, excluding already labeled ones
  var searchQueries = [
    'subject:"Daily Move Report for 191" newer_than:7d -label:"' + labelName + '"',
    'subject:"Daily Move Report for 199" newer_than:7d -label:"' + labelName + '"'
  ];

  // Get or create the "Processed Moves" label
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  // Iterate over each search query
  for (var q = 0; q < searchQueries.length; q++) {
    var threads = GmailApp.search(searchQueries[q]);

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();

      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var attachments = message.getAttachments();

        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];
          var fileName = attachment.getName();

          // Check if attachment is a CSV file and matches our expected naming
          if (fileName.toLowerCase().indexOf('.csv') !== -1 && fileName.toLowerCase().indexOf('whmovrpt') !== -1) {

            // Format the date from the email to append to the filename
            var emailDate = message.getDate();
            var formattedDate = Utilities.formatDate(emailDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

            // Create the new file name (e.g., whmovrpt199_2026-03-09)
            var baseName = fileName.replace(/\.[cC][sS][vV]$/, "");
            var newFileName = baseName + "_" + formattedDate;

            // Check if a file with this exact name already exists in the destination folder
            // Using Drive API search query to find files by name in the specific folder
            var existingFiles;

            try {
               // Use Drive API v3 to search for the file
               existingFiles = Drive.Files.list({
                 q: "name = '" + newFileName + "' and '" + folderId + "' in parents and trashed = false",
                 fields: "files(id, name)"
               }).files;
            } catch (e) {
               Logger.log("Error checking for existing file: " + e.message);
               // Fallback: If Drive API v3 fails, we'll try to use DriveApp as a backup,
               // though it's less efficient, it's safer than creating duplicates.
               var folder = DriveApp.getFolderById(folderId);
               var filesIterator = folder.getFilesByName(newFileName);
               existingFiles = [];
               while(filesIterator.hasNext()){
                   existingFiles.push(filesIterator.next());
               }
            }

            // If the file exists, log it and skip creating a new one
            if (existingFiles && existingFiles.length > 0) {
              Logger.log("File already exists in folder: " + newFileName + ". Skipping conversion.");
              continue; // Skip the rest of this attachment's processing
            }

            // Prepare blob and explicitly set content type for Drive API v3
            var blob = attachment.copyBlob();
            blob.setContentType(MimeType.CSV);

            var fileMetadata = {
              name: newFileName,
              parents: [folderId],
              mimeType: MimeType.GOOGLE_SHEETS // This forces conversion to Google Sheets format
            };

            try {
              // Create the file using Drive API v3
              var newFile = Drive.Files.create(fileMetadata, blob);
              Logger.log("Successfully converted and saved: " + newFileName);

              // Delay to allow Drive time to propagate the new file (best practice)
              Utilities.sleep(3000);

            } catch (e) {
              Logger.log("Error converting and saving file " + fileName + ": " + e.message);
              // Throw the error so the user can see it if run manually, or log it if triggered.
              throw new Error("Failed to create file via Drive API. Make sure the 'Drive API' service is enabled in the editor. Error details: " + e.message);
            }
          }
        }
      }

      // Apply the label to the thread to prevent it from being processed again
      thread.addLabel(label);
    }
  }
}
