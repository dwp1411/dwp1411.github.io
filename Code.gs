function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Operational Zone Review')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSlideshowData(weekNumber) {
  var spreadsheetId = '1nqIBYvuodJmbO4UZBls2w9q1gB2_QU1Up0FNbTzkkYY';
  var parentFolderId = '1BteCTcpPemrtPa5CuY1vuHQWDXkx4isu';

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheetName = 'Week ' + weekNumber + ' Audit results';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { error: 'Sheet "' + sheetName + '" not found.' };
  }

  var data = sheet.getDataRange().getValues();
  // Assume Row 1 is headers. Zone is Col C (index 2), Score is Col D (index 3), Notes is Col E (index 4)
  var zones = [];

  for (var i = 1; i < data.length; i++) {
    var zoneName = data[i][2];
    var score = data[i][3];
    var notes = data[i][4];

    if (zoneName) {
      zones.push({
        zone: String(zoneName).trim(),
        score: score,
        notes: notes,
        images: []
      });
    }
  }

  // Find the folder for the week in the parent folder
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var weekFolderIter = parentFolder.searchFolders('title contains "Week ' + weekNumber + '"');

  if (weekFolderIter.hasNext()) {
    var weekFolder = weekFolderIter.next();
    var subFoldersIter = weekFolder.getFolders();

    while (subFoldersIter.hasNext()) {
      var subFolder = subFoldersIter.next();
      var subFolderName = subFolder.getName();

      // Try to match subfolder to a zone. For instance, subfolder name might be "1", "Zone 1", etc.
      // Match if zoneName contains the number from the subfolder name, or if they match exactly.
      var matchedZone = null;
      for (var j = 0; j < zones.length; j++) {
        // Simple match: if subfolder name is in the zone string (e.g. "1" is in "Zone 1")
        if (zones[j].zone.toString().indexOf(subFolderName) !== -1 || subFolderName.indexOf(zones[j].zone.toString()) !== -1) {
          matchedZone = zones[j];
          break;
        }
      }

      if (matchedZone) {
        var filesIter = subFolder.getFiles();
        while (filesIter.hasNext()) {
          var file = filesIter.next();
          // Verify it's an image
          var mimeType = file.getMimeType();
          if (mimeType.indexOf('image') !== -1) {
            // Memory: use thumbnail url format
            matchedZone.images.push('https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1920-h1080');
          }
        }
      }
    }
  } else {
    // Return a warning but still send the data if no folder found
    return { error: 'Folder for Week ' + weekNumber + ' not found in Drive.', data: zones };
  }

  return { success: true, data: zones, weekNumber: weekNumber };
}
