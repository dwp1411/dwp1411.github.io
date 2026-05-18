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

  var warningMessage = null;

  // Find the folder for the week in the parent folder
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  // Search using just the number to be more forgiving with "Week12" vs "Week 12"
  var weekFolderIter = parentFolder.searchFolders('title contains "' + weekNumber + '"');

  var foundWeekFolder = null;
  while (weekFolderIter.hasNext()) {
    var folder = weekFolderIter.next();
    // Verify it actually matches the week number via regex to avoid "1" matching "12"
    var match = folder.getName().match(/\d+/);
    if (match && match[0] == weekNumber) {
      foundWeekFolder = folder;
      break;
    }
  }

  if (foundWeekFolder) {
    var subFoldersIter = foundWeekFolder.getFolders();

    while (subFoldersIter.hasNext()) {
      var subFolder = subFoldersIter.next();
      var subFolderName = subFolder.getName();

      // Extract the number from the subfolder name (e.g. "Zone 1" -> "1")
      var subfolderMatch = subFolderName.match(/\d+/);
      var subfolderNum = subfolderMatch ? subfolderMatch[0] : null;

      if (subfolderNum) {
        var matchedZone = null;
        for (var j = 0; j < zones.length; j++) {
          var zoneStr = zones[j].zone.toString();
          var zoneMatch = zoneStr.match(/\d+/);
          var zoneNum = zoneMatch ? zoneMatch[0] : null;

          if (zoneNum === subfolderNum) {
            matchedZone = zones[j];
            break;
          }
        }

        if (matchedZone) {
          var filesIter = subFolder.getFiles();
          while (filesIter.hasNext()) {
            var file = filesIter.next();
            var mimeType = file.getMimeType();
            if (mimeType.indexOf('image') !== -1) {
              matchedZone.images.push('https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1920-h1080');
            }
          }
        }
      }
    }
  } else {
    // Return a warning but still send the data if no folder found
    warningMessage = 'Folder for Week ' + weekNumber + ' not found in Drive. Pictures will not be shown.';
  }

  return { success: true, data: zones, weekNumber: weekNumber, warning: warningMessage };
}
