function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Payroll')
      .addItem('Submit Hours', 'submitHours')
      .addItem('Sync Zoned Hours', 'syncPayrollZonedHours')
      .addItem('Clear Saved Name Mappings', 'clearNameMappings')
      .addToUi();
}

/**
 * Normalizes a name for comparison.
 */
function normalizeName(nameStr) {
  if (!nameStr) return "";
  return nameStr.toString()
    .trim()
    .toUpperCase()
    .replace(/[.,\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Gets saved name mappings from Document Properties.
 */
function getSavedMappings() {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty('NameMappings');
  return saved ? JSON.parse(saved) : {};
}

/**
 * Saves a new name mapping to Document Properties.
 */
function saveMapping(sourceName, dropName) {
  var props = PropertiesService.getDocumentProperties();
  var mappings = getSavedMappings();
  mappings[normalizeName(sourceName)] = normalizeName(dropName);
  props.setProperty('NameMappings', JSON.stringify(mappings));
}

/**
 * Clears all saved name mappings.
 */
function clearNameMappings() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Clear Mappings', 'Are you sure you want to clear all saved name aliases? This means the script will ask you to match mismatched names again.', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    PropertiesService.getDocumentProperties().deleteProperty('NameMappings');
    ui.alert('All saved name mappings have been cleared.');
  }
}

/**
 * Helper to escape HTML to prevent injection
 */
function escapeHtml(unsafe) {
    return (unsafe || "").toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/**
 * Main function to synchronize hours from Calculator to Timecard.
 */
function submitHours() {
  var calculatorId = '18O_zZ2TQRRABy_J_GSHxNXVk-WDas1onJ6AzaKlnfG8';
  var ss;

  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || ss.getId() !== calculatorId) {
      ss = SpreadsheetApp.openById(calculatorId);
    }
  } catch (e) {
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
  var dayOfWeek = targetDate.getDay();

  var diff = targetDate.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  var mondayDate = new Date(targetDate);
  mondayDate.setDate(diff);

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
  var timecardFileId = null;
  var timecardFile = null;

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(mondayString) !== -1) {
      timecardFileId = file.getId();
      timecardFile = SpreadsheetApp.open(file);
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

  var dayConfigs = {
    1: { nameCol: 3, regCol: 6, otCol: 7, dateCell: 'F1', tempCol: 6 },
    2: { nameCol: 12, regCol: 15, otCol: 16, dateCell: 'O1', tempCol: 7 },
    3: { nameCol: 21, regCol: 24, otCol: 25, dateCell: 'Y1', tempCol: 8 },
    4: { nameCol: 30, regCol: 33, otCol: 34, dateCell: 'AG1', tempCol: 9 },
    5: { nameCol: 39, regCol: 42, otCol: 43, dateCell: 'AP1', tempCol: 10 },
    6: { nameCol: 48, regCol: 51, otCol: 52, dateCell: 'AY1', tempCol: 11 },
    0: { nameCol: 56, regCol: 59, otCol: 60, dateCell: 'BH1', tempCol: 12 }
  };

  var config = dayConfigs[dayOfWeek];

  var dateInFileValue = payrollDropSheet.getRange(config.dateCell).getValue();
  if (dateInFileValue instanceof Date) {
    dateInFileValue.setHours(0, 0, 0, 0);
    if (dateInFileValue.getTime() !== targetDate.getTime()) {
      SpreadsheetApp.getUi().alert('Error: Date mismatch! The target date is ' + targetDate.toLocaleDateString() + ' but the date in cell ' + config.dateCell + ' of the timecard is ' + dateInFileValue.toLocaleDateString());
      return;
    }
  }

  // Extract Source Data
  var crateLastRow = crateSheet.getLastRow();
  var crateData = [];
  if (crateLastRow > 1) {
    crateData = crateSheet.getRange(2, 1, crateLastRow - 1, 8).getValues();
  }

  var tempLastRow = tempSheet.getLastRow();
  var tempLastCol = tempSheet.getLastColumn();
  var tempData = [];

  if (tempLastRow > 0 && tempLastCol > 0) {
    var fullTempData = tempSheet.getRange(1, 1, Math.min(tempLastRow, 50), tempLastCol).getValues();
    var headerRowIndex = -1;
    var nameColIdx = -1;
    var daysCols = { 1: -1, 2: -1, 3: -1, 4: -1, 5: -1, 6: -1, 0: -1 };

    for (var r = 0; r < fullTempData.length; r++) {
      for (var c = 0; c < fullTempData[r].length; c++) {
        var cellVal = String(fullTempData[r][c]).trim().toLowerCase();
        if (cellVal.indexOf('employee name') !== -1) {
          headerRowIndex = r;
          break;
        }
      }
      if (headerRowIndex !== -1) break;
    }

    if (headerRowIndex !== -1) {
      for (var c = 0; c < tempLastCol; c++) {
        var hVal = String(fullTempData[headerRowIndex][c]).trim().toLowerCase();
        if (hVal.indexOf('employee name') !== -1) nameColIdx = c;
        else if (hVal.indexOf('mon') !== -1) daysCols[1] = c;
        else if (hVal.indexOf('tue') !== -1) daysCols[2] = c;
        else if (hVal.indexOf('wed') !== -1) daysCols[3] = c;
        else if (hVal.indexOf('thu') !== -1) daysCols[4] = c;
        else if (hVal.indexOf('fri') !== -1) daysCols[5] = c;
        else if (hVal.indexOf('sat') !== -1) daysCols[6] = c;
        else if (hVal.indexOf('sun') !== -1) daysCols[0] = c;
      }

      var dataStartRow = headerRowIndex + 2;
      if (tempLastRow >= dataStartRow) {
        var rawTempData = tempSheet.getRange(dataStartRow, 1, tempLastRow - dataStartRow + 1, tempLastCol).getValues();
        var targetTempColIdx = daysCols[dayOfWeek];

        if (nameColIdx !== -1 && targetTempColIdx !== -1) {
          for (var r = 0; r < rawTempData.length; r++) {
            tempData.push([
                null,
                rawTempData[r][nameColIdx],
                rawTempData[r][targetTempColIdx]
            ]);
          }
        }
      }
    } else {
      SpreadsheetApp.getUi().alert('Warning: Could not find "Employee Name" header in Temp tab. Temp data skipped.');
    }
  }

  var payrollLastRow = payrollDropSheet.getLastRow();

  // Overwrite the dynamic formula with static text
  // We grab the full block (Supervisor, Checkbox, Name, Function) to freeze all 4 columns generated by LET
  var blockStartCol = Math.max(1, config.nameCol - 2); // Assumes config.nameCol is the name, -2 is Supervisor
  var nameBlockRange = payrollDropSheet.getRange(1, blockStartCol, payrollLastRow, 4);
  var nameBlockValues = nameBlockRange.getValues();
  nameBlockRange.setValues(nameBlockValues); // Paste values back to destroy the formula and lock data

  var payrollNames = payrollDropSheet.getRange(1, config.nameCol, payrollLastRow, 1).getValues();
  var regRange = payrollDropSheet.getRange(1, config.regCol, payrollLastRow, 1);
  var otRange = payrollDropSheet.getRange(1, config.otCol, payrollLastRow, 1);
  var regValues = regRange.getValues();
  var otValues = otRange.getValues();

  var processedRowIndices = new Set();

  function namesMatch(searchStr, targetStr) {
    if (!searchStr || !targetStr) return false;
    return searchStr === targetStr;
  }

  function updateHours(name, hours) {
    if (!name || hours === "" || hours === null || hours === undefined) return;

    var hoursNum = parseFloat(hours) || 0;
    var reg = Math.min(hoursNum, 8);
    var ot = Math.max(0, hoursNum - 8);

    var searchName = normalizeName(name);
    if (searchName === "") return;

    for (var i = 0; i < payrollNames.length; i++) {
      var currentName = normalizeName(payrollNames[i][0]);

      if (namesMatch(searchName, currentName)) {
        regValues[i][0] = reg;
        otValues[i][0] = ot;
        processedRowIndices.add(i);
        return;
      }
    }
  }

  for (var i = 0; i < crateData.length; i++) {
    updateHours(crateData[i][0], crateData[i][7]);
  }

  for (var i = 0; i < tempData.length; i++) {
    updateHours(tempData[i][1], tempData[i][2]);
  }

  for (var i = 0; i < payrollNames.length; i++) {
    var rowName = payrollNames[i][0].toString().trim();
    if (rowName !== "" && rowName.toLowerCase().indexOf('associate') === -1 && !processedRowIndices.has(i)) {
      regValues[i][0] = 0;
      otValues[i][0] = 0;
    }
  }

  regRange.setValues(regValues);
  otRange.setValues(otValues);
}

/**
 * sync_payroll_zoned_hours.js
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
    const dayOfWeek = targetDate.getDay();

    const daysOfWeek = ['Sunday', 'Monday_', 'Tuesday_', 'Wednesday_', 'Thursday_', 'Friday_', 'Saturday'];
    const expectedTabName = daysOfWeek[dayOfWeek];

    const files = DriveApp.searchFiles("title contains '191 Week' and mimeType = 'application/vnd.google-apps.spreadsheet'");

    let targetSheet = null;
    let targetSpreadsheet = null;
    let timecardFileId = null;

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
            timecardFileId = file.getId();
            break;
          }
        }
      }
    }

    if (!targetSheet) {
      masterSs.toast(`Error: Could not find a timecard file for ${expectedTabName} matching date ${targetDate.toLocaleDateString()}`);
      return;
    }

    const startRow = 5;
    const lastRow = targetSheet.getLastRow();

    if (lastRow < startRow) {
      masterSs.toast("Error: No data rows found in the timecard tab.");
      return;
    }

    const numRows = lastRow - startRow + 1;

    const nameRange = targetSheet.getRange(startRow, 2, numRows, 1);
    const nameData = nameRange.getValues();

    const payrollRange = targetSheet.getRange(startRow, 7, numRows, 1);
    const payrollData = payrollRange.getValues();

    const zonedRange = targetSheet.getRange(startRow, 9, numRows, 46);
    const zonedData = zonedRange.getValues();

    let updatesMade = 0;

    for (let r = 0; r < numRows; r++) {
      let payrollHours = parseFloat(payrollData[r][0]);

      if (isNaN(payrollHours) || payrollHours <= 0) continue;

      if (payrollHours > 0.5) {
        payrollHours -= 0.5;
      }

      let totalZonedHours = 0;
      let zonedCols = [];

      for (let c = 0; c < 46; c++) {
        const val = parseFloat(zonedData[r][c]);
        if (!isNaN(val) && val > 0) {
          totalZonedHours += val;
          zonedCols.push({ index: c, val: val });
        }
      }

      if (totalZonedHours <= 0) continue;

      let remainingPayroll = payrollHours;

      for (let k = 0; k < zonedCols.length; k++) {
        const colInfo = zonedCols[k];

        if (k === zonedCols.length - 1) {
          let finalVal = Math.round(remainingPayroll * 100) / 100;
          zonedData[r][colInfo.index] = finalVal;
        } else {
          const proportion = colInfo.val / totalZonedHours;
          let calculated = Math.round(payrollHours * proportion * 100) / 100;
          zonedData[r][colInfo.index] = calculated;
          remainingPayroll -= calculated;
        }
      }
      updatesMade++;
    }

    zonedRange.setValues(zonedData);

    // Now perform the mismatched check
    const payrollDropSheet = targetSpreadsheet.getSheetByName('Payroll Drop');
    let mismatches = [];

    const mappings = getSavedMappings();

    if (payrollDropSheet) {
      const dropLastRow = payrollDropSheet.getLastRow();

      const dayConfigs = {
        1: { nameCol: 3, regCol: 6, otCol: 7 },
        2: { nameCol: 12, regCol: 15, otCol: 16 },
        3: { nameCol: 21, regCol: 24, otCol: 25 },
        4: { nameCol: 30, regCol: 33, otCol: 34 },
        5: { nameCol: 39, regCol: 42, otCol: 43 },
        6: { nameCol: 48, regCol: 51, otCol: 52 },
        0: { nameCol: 56, regCol: 59, otCol: 60 }
      };

      const config = dayConfigs[dayOfWeek];

      if (config && dropLastRow > 0) {
        const dropNamesRaw = payrollDropSheet.getRange(1, config.nameCol, dropLastRow, 1).getValues();

        var dropNames = [];
        for (var i = 0; i < dropNamesRaw.length; i++) {
            var nameStr = dropNamesRaw[i][0].toString().trim();
            if (nameStr !== "" && nameStr.toLowerCase().indexOf('associate') === -1) {
                dropNames.push(nameStr);
            }
        }

        // We check the Daily Tab to see if anyone has Zoned Hours but 0 Payroll Hours
        for (let r = 0; r < numRows; r++) {
            let dailyName = nameData[r][0].toString().trim();
            if (!dailyName) continue;

            let pHours = parseFloat(payrollData[r][0]);
            let hasZoned = false;

            // Check if they zoned anything
            for (let c = 0; c < 46; c++) {
                const val = parseFloat(zonedData[r][c]);
                if (!isNaN(val) && val > 0) {
                    hasZoned = true;
                    break;
                }
            }

            if (hasZoned && (isNaN(pHours) || pHours <= 0)) {
                // Check if we already have a mapping for them
                var normDaily = normalizeName(dailyName);
                if (mappings[normDaily]) {
                     // We know who they are, but their name is wrong on the daily sheet.
                     // The mapped name wasn't written to the daily sheet yet.
                     mismatches.push({ dailyName: dailyName, rowIdx: r + startRow });
                } else {
                     mismatches.push({ dailyName: dailyName, rowIdx: r + startRow });
                }
            }
        }

        if (mismatches.length > 0) {
            // Display UI
            dropNames.sort();

            var html = `
            <html>
                <head>
                <style>
                    body { font-family: Arial, sans-serif; padding: 10px; }
                    .mismatch { margin-bottom: 15px; border: 1px solid #ccc; padding: 10px; border-radius: 5px; }
                    .mismatch strong { color: #d32f2f; }
                    select { width: 100%; padding: 5px; margin-top: 5px; }
                    .btn { background-color: #1a73e8; color: white; border: none; padding: 8px 16px; cursor: pointer; border-radius: 4px; font-weight: bold;}
                    .btn:hover { background-color: #1557b0; }
                </style>
                <script>
                    function submitForm() {
                        var btn = document.getElementById('submitBtn');
                        btn.disabled = true;
                        btn.innerText = "Processing...";

                        var selects = document.querySelectorAll('select');
                        var updates = [];
                        var mappings = {};

                        for (var i = 0; i < selects.length; i++) {
                            var rowIdx = selects[i].getAttribute('data-row');
                            var sourceName = selects[i].getAttribute('data-source');
                            var mappedTo = selects[i].value;

                            if (mappedTo !== "_SKIP_") {
                                mappings[sourceName] = mappedTo;
                                updates.push({ row: parseInt(rowIdx), name: mappedTo });
                            }
                        }

                        google.script.run
                            .withSuccessHandler(function() { google.script.host.close(); })
                            .processSyncMappings(mappings, updates, ${JSON.stringify(timecardFileId)}, "${expectedTabName}");
                    }
                </script>
                </head>
                <body>
                <p>The following people on the Daily Tab have Zoned Hours but 0 Payroll Hours. Please match them to a name from the Payroll Drop tab so we can fix the Daily Tab.</p>
                <form onsubmit="event.preventDefault(); submitForm();">
            `;

            var dropOptionsHtml = '<option value="_SKIP_">-- Skip / Ignore --</option>';
            for (var i=0; i<dropNames.length; i++) {
                dropOptionsHtml += '<option value="' + escapeHtml(dropNames[i]) + '">' + escapeHtml(dropNames[i]) + '</option>';
            }

            for (var i = 0; i < mismatches.length; i++) {
                html += '<div class="mismatch">Daily Tab Name: <strong>' + escapeHtml(mismatches[i].dailyName) + '</strong><br>';
                html += '<select data-source="' + escapeHtml(mismatches[i].dailyName) + '" data-row="' + mismatches[i].rowIdx + '">';

                // Pre-select if we already have an alias
                var normDaily = normalizeName(mismatches[i].dailyName);
                var preselectedHtml = dropOptionsHtml;
                if (mappings[normDaily]) {
                    var alias = mappings[normDaily];
                    // Very simple preselection string replacement
                    preselectedHtml = preselectedHtml.replace('value="' + escapeHtml(alias) + '"', 'value="' + escapeHtml(alias) + '" selected');
                }

                html += preselectedHtml;
                html += '</select></div>';
            }

            html += '<br><button id="submitBtn" class="btn" type="submit">Fix Names & Resume Sync</button></form></body></html>';

            var htmlOutput = HtmlService.createHtmlOutput(html)
                .setWidth(450)
                .setHeight(500)
                .setTitle('Resolve Zoned/Payroll Mismatches');

            SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Resolve Zoned/Payroll Mismatches');
            return;
        }
      }
    }

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

/**
 * Called by the HTML dialog to fix the Daily Tab names, save mappings, and resume the sync.
 */
function processSyncMappings(newMappings, updates, timecardFileId, expectedTabName) {
    for (var source in newMappings) {
       saveMapping(source, newMappings[source]);
    }

    var ss = SpreadsheetApp.openById(timecardFileId);
    var targetSheet = ss.getSheetByName(expectedTabName);

    if (targetSheet && updates.length > 0) {
        for (var i = 0; i < updates.length; i++) {
            // Write the correct name directly to Column B on the Daily Tab
            targetSheet.getRange(updates[i].row, 2).setValue(updates[i].name);
        }
    }

    if (updates.length > 0) {
        // Force spreadsheet recalculation by sleeping momentarily
        SpreadsheetApp.flush();
        Utilities.sleep(1000);

        // Re-trigger the sync function to calculate the now-populated payroll hours
        syncPayrollZonedHours();
    } else {
        // If everything was skipped, just exit gracefully.
        var masterSs = SpreadsheetApp.openById('18O_zZ2TQRRABy_J_GSHxNXVk-WDas1onJ6AzaKlnfG8');
        masterSs.toast("Sync completed (mismatches skipped).", "Sync Complete", 5);
    }
}
