/**
 * sync_payroll_zoned_hours.js
 *
 * Standalone Google Apps Script to synchronize Payroll Hours (Column G)
 * with Zoned Hours (Columns I through BB) in the weekly Timecard files.
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
