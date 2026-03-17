function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Payroll')
      .addItem('Submit Hours', 'submitHours')
      .addToUi();
}

function submitHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var crateSheet = ss.getSheetByName('Crate');
  var tempSheet = ss.getSheetByName('Temp');

  if (!crateSheet || !tempSheet) {
    SpreadsheetApp.getUi().alert('Error: "Crate" or "Temp" tab not found in the calculator.');
    return;
  }

  var targetDateValue = crateSheet.getRange('A1').getValue();
  if (!(targetDateValue instanceof Date)) {
    SpreadsheetApp.getUi().alert('Error: A1 in Crate tab must be a valid date.');
    return;
  }

  var targetDate = new Date(targetDateValue);
  targetDate.setHours(0, 0, 0, 0);
  var dayOfWeek = targetDate.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // Calculate Monday of the week
  // If Sunday (0), Monday was 6 days ago. If Monday (1), it's today.
  var diff = targetDate.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  var mondayDate = new Date(targetDate);
  mondayDate.setDate(diff);

  var mondayString = (mondayDate.getMonth() + 1) + '/' + mondayDate.getDate() + '/' + mondayDate.getFullYear().toString().slice(-2);

  var folderId = '1e0n54IsOYFJja4iPgCbwnn-huhvVBMDo';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var timecardFile = null;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(mondayString) !== -1) {
      timecardFile = SpreadsheetApp.open(file);
      break;
    }
  }

  if (!timecardFile) {
    SpreadsheetApp.getUi().alert('Error: Could not find Timecard file for week starting ' + mondayString + ' in the specified folder.');
    return;
  }

  var payrollDropSheet = timecardFile.getSheetByName('Payroll Drop');
  if (!payrollDropSheet) {
    SpreadsheetApp.getUi().alert('Error: "Payroll Drop" tab not found in the timecard file: ' + timecardFile.getName());
    return;
  }

  // Day configuration
  // Monday is 1, Tuesday 2, ..., Saturday 6, Sunday 0
  var dayConfigs = {
    1: { nameCol: 3, regCol: 6, otCol: 7, dateCell: 'F1', tempCol: 6 }, // Mon
    2: { nameCol: 12, regCol: 15, otCol: 16, dateCell: 'O1', tempCol: 7 }, // Tue
    3: { nameCol: 21, regCol: 24, otCol: 25, dateCell: 'Y1', tempCol: 8 }, // Wed
    4: { nameCol: 30, regCol: 33, otCol: 34, dateCell: 'AG1', tempCol: 9 }, // Thu
    5: { nameCol: 39, regCol: 42, otCol: 43, dateCell: 'AP1', tempCol: 10 }, // Fri
    6: { nameCol: 48, regCol: 51, otCol: 52, dateCell: 'AY1', tempCol: 11 }, // Sat
    0: { nameCol: 56, regCol: 59, otCol: 60, dateCell: 'BH1', tempCol: 12 }  // Sun
  };

  var config = dayConfigs[dayOfWeek];

  // Verify date in the timecard file
  var dateInFileValue = payrollDropSheet.getRange(config.dateCell).getValue();
  if (dateInFileValue instanceof Date) {
    dateInFileValue.setHours(0, 0, 0, 0);
    if (dateInFileValue.getTime() !== targetDate.getTime()) {
      SpreadsheetApp.getUi().alert('Error: The date in ' + config.dateCell + ' of the timecard (' + dateInFileValue.toLocaleDateString() + ') does not match the target date (' + targetDate.toLocaleDateString() + ').');
      return;
    }
  } else {
    // If it's not a Date object, maybe it's a string? Try to compare anyway if it looks like a date.
    // For now, if it's not a Date, we'll just log a warning but proceed, or we could be stricter.
    console.warn('Date cell ' + config.dateCell + ' in timecard is not a Date object.');
  }

  // Get data from Crate tab (Name in A, Hours in H)
  var crateLastRow = crateSheet.getLastRow();
  var crateData = [];
  if (crateLastRow > 1) {
    crateData = crateSheet.getRange(2, 1, crateLastRow - 1, 8).getValues();
  }

  // Get data from Temp tab (Name in B, Hours in Mon-Sun columns F-L)
  var tempLastRow = tempSheet.getLastRow();
  var tempData = [];
  if (tempLastRow > 1) {
    tempData = tempSheet.getRange(2, 1, tempLastRow - 1, 12).getValues();
  }

  // Get Payroll Drop names for matching
  var payrollLastRow = payrollDropSheet.getLastRow();
  var payrollNames = payrollDropSheet.getRange(1, config.nameCol, payrollLastRow, 1).getValues();

  /**
   * Updates Regular and OT hours for a given associate.
   */
  function updateHours(name, hours) {
    if (!name || hours === "" || hours === null || hours === undefined) return;

    var hoursNum = parseFloat(hours) || 0;
    var reg = Math.min(hoursNum, 8);
    var ot = Math.max(0, hoursNum - 8);

    var searchName = name.toString().trim().toLowerCase();

    for (var i = 0; i < payrollNames.length; i++) {
      if (payrollNames[i][0].toString().trim().toLowerCase() === searchName) {
        // Update cells (i+1 because row index is 1-based)
        payrollDropSheet.getRange(i + 1, config.regCol).setValue(reg);
        payrollDropSheet.getRange(i + 1, config.otCol).setValue(ot);
        return; // Found and updated, stop searching
      }
    }
  }

  // Process Crate data
  for (var i = 0; i < crateData.length; i++) {
    updateHours(crateData[i][0], crateData[i][7]);
  }

  // Process Temp data
  for (var i = 0; i < tempData.length; i++) {
    updateHours(tempData[i][1], tempData[i][config.tempCol - 1]);
  }

  // No popup needed per user request
}
