function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Payroll')
      .addItem('Submit Hours', 'submitHours')
      .addToUi();
}

/**
 * Main function to synchronize hours from Calculator to Timecard.
 */
function submitHours() {
  var calculatorId = '18O_zZ2TQRRABy_J_GSHxNXVk-WDas1onJ6AzaKlnfG8';
  var ss;

  try {
    // Attempt to get the active spreadsheet first (if running as container-bound)
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || ss.getId() !== calculatorId) {
      // If not active or IDs don't match, open by ID explicitly
      ss = SpreadsheetApp.openById(calculatorId);
    }
  } catch (e) {
    // If openById fails or script is standalone without permissions, notify user
    SpreadsheetApp.getUi().alert('Error: Could not access the Calculator spreadsheet with ID: ' + calculatorId);
    return;
  }

  var crateSheet = ss.getSheetByName('Crate');
  var tempSheet = ss.getSheetByName('Temp');

  if (!crateSheet || !tempSheet) {
    SpreadsheetApp.getUi().alert('Error: "Crate" or "Temp" tab not found in the calculator spreadsheet.');
    return;
  }

  var targetDateValue = crateSheet.getRange('A1').getValue();
  if (!(targetDateValue instanceof Date)) {
    SpreadsheetApp.getUi().alert('Error: Cell A1 in the "Crate" tab must contain a valid date in the format M/D/YYYY.');
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

  // File name date format is M/D/YY (e.g., 3/16/26)
  var mondayString = (mondayDate.getMonth() + 1) + '/' + mondayDate.getDate() + '/' + mondayDate.getFullYear().toString().slice(-2);

  var folderId = '1e0n54IsOYFJja4iPgCbwnn-huhvVBMDo';
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: Could not find the Timecard folder with ID: ' + folderId);
    return;
  }

  var files = folder.getFiles();
  var timecardFile = null;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(mondayString) !== -1) {
      try {
        timecardFile = SpreadsheetApp.open(file);
      } catch (e) {
        console.warn('Found file ' + file.getName() + ' but could not open it.');
      }
      break;
    }
  }

  if (!timecardFile) {
    SpreadsheetApp.getUi().alert('Error: Could not find a Timecard file containing "' + mondayString + '" in its name within the specified folder.');
    return;
  }

  var payrollDropSheet = timecardFile.getSheetByName('Payroll Drop');
  if (!payrollDropSheet) {
    SpreadsheetApp.getUi().alert('Error: "Payroll Drop" tab not found in the timecard file: ' + timecardFile.getName());
    return;
  }

  // Day configuration mapping
  var dayConfigs = {
    1: { nameCol: 3, regCol: 6, otCol: 7, dateCell: 'F1', tempCol: 6 }, // Mon
    2: { nameCol: 12, regCol: 15, otCol: 16, dateCell: 'O1', tempCol: 7 }, // Tue
    3: { nameCol: 21, regCol: 24, otCol: 25, dateCell: 'Y1', tempCol: 8 }, // Wed
    4: { nameCol: 30, regCol: 33, otCol: 34, dateCell: 'AG1', tempCol: 9 }, // Thu
    5: { nameCol: 39, regCol: 42, otCol: 43, dateCell: 'AP1', tempCol: 10 }, // Fri (OT in AQ is 43rd col)
    6: { nameCol: 48, regCol: 51, otCol: 52, dateCell: 'AY1', tempCol: 11 }, // Sat
    0: { nameCol: 56, regCol: 59, otCol: 60, dateCell: 'BH1', tempCol: 12 }  // Sun
  };

  var config = dayConfigs[dayOfWeek];

  // Verify date in the timecard file's Payroll Drop tab
  var dateInFileValue = payrollDropSheet.getRange(config.dateCell).getValue();
  if (dateInFileValue instanceof Date) {
    dateInFileValue.setHours(0, 0, 0, 0);
    if (dateInFileValue.getTime() !== targetDate.getTime()) {
      SpreadsheetApp.getUi().alert('Error: Date mismatch! The target date is ' + targetDate.toLocaleDateString() + ' but the date in cell ' + config.dateCell + ' of the timecard is ' + dateInFileValue.toLocaleDateString());
      return;
    }
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
        payrollDropSheet.getRange(i + 1, config.regCol).setValue(reg);
        payrollDropSheet.getRange(i + 1, config.otCol).setValue(ot);
        return;
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

  // Quiet success per user request
}
