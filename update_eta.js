/**
 * Update the "All Loads" sheet with ETA data from today's "In Transit ETA Now" file.
 *
 * Instructions:
 * 1. Open your "All Loads" spreadsheet.
 * 2. Click on "Extensions" > "Apps Script".
 * 3. Delete any code in the code editor and paste this entire script.
 * 4. Save the project (Ctrl+S or Cmd+S).
 * 5. Refresh your "All Loads" spreadsheet in the browser.
 * 6. You should see a new menu item at the top called "ETA Updates".
 * 7. Click "ETA Updates" > "Update Days Past TT ASN" to run the script.
 * 8. The first time you run it, you'll need to authorize the script.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ETA Updates')
    .addItem('Update Days Past TT ASN', 'updateDaysPastTTASN')
    .addToUi();
}

function updateDaysPastTTASN() {
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. Determine today's file name
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const fileName = `${yyyy}-${mm}-${dd} - In Transit ETA Now`;

    // 2. Find the file in the specific folder
    const folderId = '1D8XcYrF_kj2Lx9mI5ZdeuGRrHLusd19P';
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      ui.alert('File Not Found', `Could not find today's ETA file named:\n"${fileName}"\nin the In Transit ETA folder.`, ui.ButtonSet.OK);
      return;
    }

    const file = files.next();
    const sourceSpreadsheet = SpreadsheetApp.openById(file.getId());
    const sourceSheet = sourceSpreadsheet.getSheets()[0]; // Assumes the data is on the first sheet
    const sourceData = sourceSheet.getDataRange().getValues();

    // 3. Build a map of Trailer-Site to ETA Date
    const etaMap = {};
    // Assuming row 1 is headers, start at i = 1
    for (let i = 1; i < sourceData.length; i++) {
      let trailer = String(sourceData[i][5]).trim(); // Column F is index 5
      let etaValue = sourceData[i][7];               // Column H is index 7
      let site = String(sourceData[i][13]).trim();   // Column N is index 13

      // We only care about site 191 or 199
      if (trailer && etaValue && (site === '191' || site === '199')) {
        let key = trailer + '-' + site;
        // If duplicates exist, this stores the first one found.
        if (!etaMap[key]) {
          let parsedDate = new Date(etaValue);
          if (Object.prototype.toString.call(parsedDate) === "[object Date]" && !isNaN(parsedDate.getTime())) {
            etaMap[key] = parsedDate;
          }
        }
      }
    }

    // 4. Update the "All Loads" sheet
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheet = targetSpreadsheet.getSheetByName('All Loads');

    if (!targetSheet) {
      ui.alert('Sheet Not Found', 'Could not find a tab named "All Loads" in this spreadsheet.', ui.ButtonSet.OK);
      return;
    }

    const targetData = targetSheet.getDataRange().getValues();

    // If there's no data (just headers or empty), nothing to update
    if (targetData.length <= 1) {
      ui.alert('No Data', 'The "All Loads" sheet does not have any rows to update.', ui.ButtonSet.OK);
      return;
    }

    // We need to update Column K (index 10) for rows 2 through targetData.length
    // Let's grab the current values so we don't overwrite rows that shouldn't change
    const targetRange = targetSheet.getRange(2, 11, targetData.length - 1, 1);
    const targetKValues = targetRange.getValues();

    const todayNorm = new Date();
    todayNorm.setHours(0, 0, 0, 0);

    let updateCount = 0;

    for (let i = 1; i < targetData.length; i++) {
      let site = String(targetData[i][2]).trim();     // Column C is index 2
      let trailer = String(targetData[i][3]).trim();  // Column D is index 3

      // If there is no information in column C, ensure column K is blank and skip the rest
      if (site === '') {
        if (targetKValues[i - 1][0] !== '') {
          targetKValues[i - 1][0] = '';
          updateCount++;
        }
        continue;
      }

      // Skip this row entirely if the site in Column C is not 191 or 199
      if (site !== '191' && site !== '199') {
        continue;
      }

      let key = trailer + '-' + site;

      if (etaMap[key]) {
        let etaDate = etaMap[key];
        etaDate.setHours(0, 0, 0, 0);

        let diffDays = Math.round((todayNorm.getTime() - etaDate.getTime()) / (1000 * 60 * 60 * 24));
        let value;

        if (diffDays <= 0) {
          value = diffDays - 1;
        } else {
          value = diffDays;
        }

        targetKValues[i - 1][0] = value;
        updateCount++;
      }
      // If not found in etaMap, targetKValues[i - 1][0] remains as is, preserving the old value.
    }

    // Write back the updated values to Column K
    targetRange.setValues(targetKValues);

    // Show silent success or alert
    targetSpreadsheet.toast(`Successfully updated ${updateCount} loads.`, 'Update Complete', 5);

  } catch (error) {
    ui.alert('Error', 'An error occurred while running the script:\n\n' + error.message, ui.ButtonSet.OK);
  }
}
