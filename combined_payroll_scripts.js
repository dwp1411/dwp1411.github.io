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
    .replace(/[.,\-]/g, ' ') // replace commas, dots, hyphens with space
    .replace(/\s+/g, ' ')    // collapse multiple spaces into one
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

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(mondayString) !== -1) {
      timecardFileId = file.getId();
      break;
    }
  }

  if (!timecardFileId) {
    SpreadsheetApp.getUi().alert('Error: Could not find a Timecard file containing "' + mondayString + '" in its name within the specified folder.');
    return;
  }

  var timecardFile = SpreadsheetApp.openById(timecardFileId);
  var payrollDropSheet = timecardFile.getSheetByName('Payroll Drop');
  if (!payrollDropSheet) {
    SpreadsheetApp.getUi().alert('Error: "Payroll Drop" tab not found in the timecard file: ' + timecardFile.getName());
    return;
  }

  var dayConfigs = {
    1: { nameCol: 3, regCol: 6, otCol: 7, dateCell: 'F1' },
    2: { nameCol: 12, regCol: 15, otCol: 16, dateCell: 'O1' },
    3: { nameCol: 21, regCol: 24, otCol: 25, dateCell: 'Y1' },
    4: { nameCol: 30, regCol: 33, otCol: 34, dateCell: 'AG1' },
    5: { nameCol: 39, regCol: 42, otCol: 43, dateCell: 'AP1' },
    6: { nameCol: 48, regCol: 51, otCol: 52, dateCell: 'AY1' },
    0: { nameCol: 56, regCol: 59, otCol: 60, dateCell: 'BH1' }
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
  var updates = [];

  var crateLastRow = crateSheet.getLastRow();
  if (crateLastRow > 1) {
    var crateData = crateSheet.getRange(2, 1, crateLastRow - 1, 8).getValues();
    for (var i = 0; i < crateData.length; i++) {
        updates.push({ name: crateData[i][0], hours: crateData[i][7] });
    }
  }

  var tempLastRow = tempSheet.getLastRow();
  var tempLastCol = tempSheet.getLastColumn();

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
            updates.push({ name: rawTempData[r][nameColIdx], hours: rawTempData[r][targetTempColIdx] });
          }
        }
      }
    } else {
      SpreadsheetApp.getUi().alert('Warning: Could not find "Employee Name" header in Temp tab. Temp data skipped.');
    }
  }

  // Get Payroll Drop names
  var payrollLastRow = payrollDropSheet.getLastRow();
  var payrollNamesRaw = payrollDropSheet.getRange(1, config.nameCol, payrollLastRow, 1).getValues();

  var dropNames = [];
  var dropNamesNormalized = [];

  for (var i = 0; i < payrollNamesRaw.length; i++) {
      var nameStr = payrollNamesRaw[i][0].toString().trim();
      if (nameStr !== "" && nameStr.toLowerCase().indexOf('associate') === -1) {
          dropNames.push(nameStr);
          dropNamesNormalized.push(normalizeName(nameStr));
      }
  }

  // Cross-reference names
  var mappings = getSavedMappings();
  var mismatches = [];

  for (var i = 0; i < updates.length; i++) {
      var uName = updates[i].name;
      var hours = updates[i].hours;

      if (!uName || hours === "" || hours === null || hours === undefined) continue;

      var normUName = normalizeName(uName);
      if (normUName === "") continue;

      // Check if it exactly matches a drop name
      var isMatch = dropNamesNormalized.indexOf(normUName) !== -1;

      // Check if it's in our saved aliases
      if (!isMatch && mappings[normUName]) {
          var mappedName = mappings[normUName];
          if (dropNamesNormalized.indexOf(mappedName) !== -1) {
             isMatch = true;
          }
      }

      if (!isMatch) {
          // It's a true mismatch
          mismatches.push(uName);
      }
  }

  // Deduplicate mismatches
  var uniqueMismatches = mismatches.filter(function(item, pos) {
      return mismatches.indexOf(item) == pos;
  });

  if (uniqueMismatches.length > 0) {
      // Build HTML UI
      dropNames.sort(); // Sort alphabetically for the dropdown

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
              var mappings = {};

              for (var i = 0; i < selects.length; i++) {
                 var sourceName = selects[i].getAttribute('data-source');
                 var mappedTo = selects[i].value;
                 if (mappedTo !== "_SKIP_") {
                    mappings[sourceName] = mappedTo;
                 }
              }

              google.script.run
                  .withSuccessHandler(function() { google.script.host.close(); })
                  .processManualMappings(mappings, ${JSON.stringify(timecardFileId)}, ${targetDateValue.getTime()});
            }
          </script>
        </head>
        <body>
          <p>We found some names in Crate/Temp that don't match the <b>Payroll Drop</b> tab. Please map them below so we can learn them for the future, or skip them.</p>
          <form onsubmit="event.preventDefault(); submitForm();">
      `;

      var dropOptionsHtml = '<option value="_SKIP_">-- Skip / Ignore --</option>';
      for (var i=0; i<dropNames.length; i++) {
          dropOptionsHtml += '<option value="' + dropNames[i] + '">' + dropNames[i] + '</option>';
      }

      for (var i = 0; i < uniqueMismatches.length; i++) {
          html += '<div class="mismatch">Found: <strong>' + uniqueMismatches[i] + '</strong><br>';
          html += '<select data-source="' + uniqueMismatches[i] + '">';
          html += dropOptionsHtml;
          html += '</select></div>';
      }

      html += '<br><button id="submitBtn" class="btn" type="submit">Save & Continue Sync</button></form></body></html>';

      var htmlOutput = HtmlService.createHtmlOutput(html)
          .setWidth(450)
          .setHeight(500)
          .setTitle('Resolve Name Mismatches');

      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Resolve Name Mismatches');
      return; // Execution stops here and resumes in processManualMappings
  }

  // If no mismatches, proceed immediately
  executeSubmitHours(timecardFileId, targetDateValue.getTime());
}

/**
 * Called by the HTML dialog to save user mappings and resume execution.
 */
function processManualMappings(newMappings, timecardFileId, targetDateTime) {
   for (var source in newMappings) {
       saveMapping(source, newMappings[source]);
   }
   executeSubmitHours(timecardFileId, targetDateTime);
}

/**
 * The second half of submitHours that actually writes the data.
 */
function executeSubmitHours(timecardFileId, targetDateTime) {
   var calculatorId = '18O_zZ2TQRRABy_J_GSHxNXVk-WDas1onJ6AzaKlnfG8';
   var ss = SpreadsheetApp.openById(calculatorId);
   var crateSheet = ss.getSheetByName('Crate');
   var tempSheet = ss.getSheetByName('Temp');

   var targetDate = new Date(targetDateTime);
   var dayOfWeek = targetDate.getDay();

   var timecardFile = SpreadsheetApp.openById(timecardFileId);
   var payrollDropSheet = timecardFile.getSheetByName('Payroll Drop');

   var dayConfigs = {
    1: { nameCol: 3, regCol: 6, otCol: 7 },
    2: { nameCol: 12, regCol: 15, otCol: 16 },
    3: { nameCol: 21, regCol: 24, otCol: 25 },
    4: { nameCol: 30, regCol: 33, otCol: 34 },
    5: { nameCol: 39, regCol: 42, otCol: 43 },
    6: { nameCol: 48, regCol: 51, otCol: 52 },
    0: { nameCol: 56, regCol: 59, otCol: 60 }
  };
  var config = dayConfigs[dayOfWeek];

  // Re-fetch Crate and Temp data
  var updates = [];

  var crateLastRow = crateSheet.getLastRow();
  if (crateLastRow > 1) {
    var crateData = crateSheet.getRange(2, 1, crateLastRow - 1, 8).getValues();
    for (var i = 0; i < crateData.length; i++) {
        updates.push({ name: crateData[i][0], hours: crateData[i][7] });
    }
  }

  var tempLastRow = tempSheet.getLastRow();
  var tempLastCol = tempSheet.getLastColumn();

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
            updates.push({ name: rawTempData[r][nameColIdx], hours: rawTempData[r][targetTempColIdx] });
          }
        }
      }
    }
  }

  var payrollLastRow = payrollDropSheet.getLastRow();
  var payrollNames = payrollDropSheet.getRange(1, config.nameCol, payrollLastRow, 1).getValues();
  var regRange = payrollDropSheet.getRange(1, config.regCol, payrollLastRow, 1);
  var otRange = payrollDropSheet.getRange(1, config.otCol, payrollLastRow, 1);
  var regValues = regRange.getValues();
  var otValues = otRange.getValues();

  var processedRowIndices = new Set();
  var mappings = getSavedMappings();

  function updateHours(name, hours) {
    if (!name || hours === "" || hours === null || hours === undefined) return;

    var hoursNum = parseFloat(hours) || 0;
    var reg = Math.min(hoursNum, 8);
    var ot = Math.max(0, hoursNum - 8);

    var searchName = normalizeName(name);
    if (searchName === "") return;

    // Apply alias if it exists
    if (mappings[searchName]) {
        searchName = mappings[searchName];
    }

    for (var i = 0; i < payrollNames.length; i++) {
      var currentName = normalizeName(payrollNames[i][0]);

      if (searchName === currentName) {
        regValues[i][0] = reg;
        otValues[i][0] = ot;
        processedRowIndices.add(i);
        return;
      }
    }
  }

  for (var i = 0; i < updates.length; i++) {
    updateHours(updates[i].name, updates[i].hours);
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

    const payrollDropSheet = targetSpreadsheet.getSheetByName('Payroll Drop');
    let mismatches = [];

    // Get saved aliases so we don't alert for known mismatches
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
        const dropNames = payrollDropSheet.getRange(1, config.nameCol, dropLastRow, 1).getValues();
        const dropReg = payrollDropSheet.getRange(1, config.regCol, dropLastRow, 1).getValues();
        const dropOt = payrollDropSheet.getRange(1, config.otCol, dropLastRow, 1).getValues();

        const dailyTabNamesList = [];
        for (let r = 0; r < numRows; r++) {
          let normName = normalizeName(nameData[r][0]);

          // Apply alias to daily tab names as well if we have one
          // This way, if the daily tab says "Mike Smith" but Drop says "Michael Smith",
          // and we mapped "Mike Smith" -> "Michael Smith", this matches perfectly.
          if (mappings[normName]) {
              normName = mappings[normName];
          }

          const pHours = parseFloat(payrollData[r][0]);
          if (normName !== "") {
            dailyTabNamesList.push({ name: normName, hours: isNaN(pHours) ? 0 : pHours });
          }
        }

        for (let i = 0; i < dropLastRow; i++) {
          const name = dropNames[i][0];
          const reg = parseFloat(dropReg[i][0]) || 0;
          const ot = parseFloat(dropOt[i][0]) || 0;
          const totalHours = reg + ot;

          if (totalHours > 0) {
            const searchName = normalizeName(name);
            let foundMatch = false;
            let dailyHours = 0;

            for (let j = 0; j < dailyTabNamesList.length; j++) {
              if (searchName === dailyTabNamesList[j].name) {
                foundMatch = true;
                dailyHours = dailyTabNamesList[j].hours;
                break;
              }
            }

            if (!foundMatch || dailyHours <= 0) {
              mismatches.push(name.toString().trim());
            }
          }
        }
      }
    }

    if (mismatches.length > 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Spelling / Missing Names Alert',
        `The following associates have hours in 'Payroll Drop' for ${expectedTabName} but show 0 hours (or are missing) in the daily tab, likely due to a spelling mismatch:\n\n` +
        mismatches.join('\n') +
        `\n\nPlease check the timecard to fix their names so hours calculate properly!`,
        ui.ButtonSet.OK
      );
    } else {
      masterSs.toast(`Success: Synced hours for ${updatesMade} associates on ${targetDate.toLocaleDateString()}.`, "Sync Complete", 5);
    }

  } catch (error) {
    console.error("Error in syncPayrollZonedHours:", error);
    try {
      SpreadsheetApp.openById(masterSpreadsheetId).toast("An error occurred. Check script execution logs.");
    } catch (e) {
      // Ignore
    }
  }
}
