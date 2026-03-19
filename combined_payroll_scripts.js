function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Payroll')
      .addItem('Submit Hours', 'submitHours')
      .addItem('Sync Zoned Hours', 'syncPayrollZonedHours')
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

/**
 * sync_payroll_zoned_hours.js
 *
 * Synchronize Payroll Hours (Column G) with Zoned Hours (Columns I through BB)
 * in the weekly Timecard files.
 */
function syncPayrollZonedHours() {
  const masterSpreadsheetId = '18O_zZ2TQRRABy_J_GSHxNXVk-WDas1onJ6AzaKlnfG8';

  try {
    const masterSs = SpreadsheetApp.openById(masterSpreadsheetId);
    const crateSheet = masterSs.getSheetByName('Crate');

    if (!crateSheet) {
      console.error("Crate tab not found in master spreadsheet.");
      return;
    }

    const targetDateValue = crateSheet.getRange('A1').getValue();
    if (!targetDateValue || !(targetDateValue instanceof Date)) {
      masterSs.toast("Error: Valid date not found in Crate tab A1.");
      return;
    }

    const targetDate = new Date(targetDateValue);
    targetDate.setHours(0, 0, 0, 0);

    // Determine expected tab name based on day of week
    const daysOfWeek = ['Sunday', 'Monday_', 'Tuesday_', 'Wednesday_', 'Thursday_', 'Friday_', 'Saturday'];
    const expectedTabName = daysOfWeek[targetDate.getDay()];

    // Search for 191 Week files in the user's Drive
    const files = DriveApp.searchFiles("title contains '191 Week' and mimeType = 'application/vnd.google-apps.spreadsheet'");

    let targetSheet = null;
    let targetSpreadsheet = null;

    while (files.hasNext()) {
      const file = files.next();
      const ss = SpreadsheetApp.openById(file.getId());
      const sheet = ss.getSheetByName(expectedTabName);

      if (sheet) {
        const dateInB1 = sheet.getRange('B1').getValue();
        if (dateInB1 && dateInB1 instanceof Date) {
          const sheetDate = new Date(dateInB1);
          sheetDate.setHours(0, 0, 0, 0);

          if (sheetDate.getTime() === targetDate.getTime()) {
            targetSheet = sheet;
            targetSpreadsheet = ss;
            break;
          }
        }
      }
    }

    if (!targetSheet) {
      masterSs.toast(`Error: Could not find a timecard file for ${expectedTabName} matching date ${targetDate.toLocaleDateString()}`);
      return;
    }

    // Process the hours
    // Associates start on Row 5
    const startRow = 5;
    const lastRow = targetSheet.getLastRow();

    if (lastRow < startRow) {
      masterSs.toast("Error: No data rows found in the timecard tab.");
      return;
    }

    const numRows = lastRow - startRow + 1;

    // Column G is Payroll Hours (7), Columns I (9) to BB (54) are Zoned Hours.
    // We will get G separately and I to BB separately to keep indices simple.
    const payrollRange = targetSheet.getRange(startRow, 7, numRows, 1);
    const payrollData = payrollRange.getValues();

    const zonedRange = targetSheet.getRange(startRow, 9, numRows, 46); // 54 - 9 + 1 = 46 cols
    const zonedData = zonedRange.getValues();

    let updatesMade = 0;

    for (let r = 0; r < numRows; r++) {
      const payrollHours = parseFloat(payrollData[r][0]);

      // If Payroll Hours is 0 or NaN, skip
      if (isNaN(payrollHours) || payrollHours <= 0) continue;

      let totalZonedHours = 0;
      let zonedCols = [];

      // Find all non-zero zoned hour entries for this associate
      for (let c = 0; c < 46; c++) {
        const val = parseFloat(zonedData[r][c]);
        if (!isNaN(val) && val > 0) {
          totalZonedHours += val;
          zonedCols.push({ index: c, val: val });
        }
      }

      // If total zoned hours is 0 (forgot to zone), skip
      if (totalZonedHours <= 0) continue;

      // Distribute the payroll hours proportionally to the zoned columns
      let remainingPayroll = payrollHours;

      for (let k = 0; k < zonedCols.length; k++) {
        const colInfo = zonedCols[k];

        if (k === zonedCols.length - 1) {
          // Last one gets the remainder to avoid rounding issues (e.g. 3.38 + 5.63 = 9.01 -> 9.00)
          let finalVal = Math.round(remainingPayroll * 100) / 100;
          zonedData[r][colInfo.index] = finalVal;
        } else {
          // Proportional split based on original zoned hours
          const proportion = colInfo.val / totalZonedHours;
          let calculated = Math.round(payrollHours * proportion * 100) / 100;
          zonedData[r][colInfo.index] = calculated;
          remainingPayroll -= calculated;
        }
      }
      updatesMade++;
    }

    // Write back the updated Zoned Hours data
    zonedRange.setValues(zonedData);

    // Notify the user of success via Toast on the master spreadsheet
    masterSs.toast(`Success: Synced hours for ${updatesMade} associates on ${targetDate.toLocaleDateString()}.`, "Sync Complete", 5);

  } catch (error) {
    console.error("Error in syncPayrollZonedHours:", error);
    try {
      SpreadsheetApp.openById(masterSpreadsheetId).toast("An error occurred. Check script execution logs.");
    } catch (e) {
      // Ignore
    }
  }
}
