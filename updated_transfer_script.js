/**
 * @license
 * Copyright 2024 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * A helper function to reliably parse date values from a sheet.
 * @param {*} dateValue The value from the sheet cell.
 * @return {Date|null} A valid Date object or null.
 */
function parseSheetDate(dateValue) {
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    return dateValue;
  }
  // This will attempt to parse strings or numbers into a date
  const d = new Date(dateValue);
  if (d instanceof Date && !isNaN(d)) {
    return d;
  }
  return null; // Return null if parsing fails
}

/**
 * Sums cells in a specified sheet that are older than a certain number of days
 * and match the background color of a reference cell.
 *
 * @param {string} sheetName The name of the sheet to search (e.g., "199 Aging Splits").
 * @param {string} colorRefA1 A single cell on the *current* sheet whose background color will be matched (e.g., "F1").
 * @param {number} daysOld The age threshold in days.
 * @return {number|string} The sum of the matching cells or an error message.
 * @customfunction
 */
function SUMBYCOLOROLDERTHAN(sheetName, colorRefA1, daysOld) {
  // 1. Argument Validation
  if (!sheetName || typeof sheetName !== 'string') {
    return 'ERROR: First argument (sheetName) is missing or invalid.';
  }
  if (!colorRefA1 || typeof colorRefA1 !== 'string') {
    return 'ERROR: Second argument (colorRefA1) is missing or invalid.';
  }
  if (daysOld === null || typeof daysOld !== 'number') { // Allow 0
    return 'ERROR: Third argument (daysOld) must be a number.';
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName);
  const activeSheet = ss.getActiveSheet();

  if (!targetSheet) {
    return `ERROR: Sheet named "${sheetName}" was not found.`;
  }

  // 2. Color Reference Validation
  let colorRef;
  try {
    colorRef = activeSheet.getRange(colorRefA1);
  } catch (e) {
    return `ERROR: Invalid cell reference "${colorRefA1}".`;
  }
  const refColor = colorRef.getBackground();

  // 3. Data Range and Structure Validation
  const dataRange = targetSheet.getDataRange();
  const allValues = dataRange.getValues();
  const allColors = dataRange.getBackgrounds();

  if (allValues.length < 3) {
    return `ERROR: Sheet "${sheetName}" has fewer than 3 rows of data.`;
  }

  const dates = allValues[2]; // Dates are in the 3rd row (index 2)
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize today to midnight
  const thresholdDate = new Date(today.getTime() - (daysOld * 24 * 60 * 60 * 1000));

  let total = 0;
  let invalidDateFound = false;

  // Start from column index 3 (Column D) where data begins
  for (let j = 3; j < dates.length; j++) {
    const cellDate = parseSheetDate(dates[j]);

    // Check if the date is valid and on or before the threshold date
    if (cellDate) {
      cellDate.setHours(0, 0, 0, 0); // Normalize cell date to midnight for accurate comparison
      if (cellDate <= thresholdDate) {
        // This column is old enough, now sum the colored cells in this column
        // Iterate through the relevant rows (Row 4 to 48, which is index 3 to 47)
        for (let i = 3; i < 48; i++) {
           // Ensure the row/column exists in the arrays to prevent errors
          if (allColors[i] && allValues[i] && allColors[i][j] !== undefined && allValues[i][j] !== undefined) {
            if (allColors[i][j] === refColor && typeof allValues[i][j] === 'number') {
              total += allValues[i][j];
            }
          }
        }
      }
    } else if (dates[j] && !invalidDateFound) {
        // Log only the first invalid date to avoid spamming
        Logger.log(`Skipping invalid date in sheet "${sheetName}", cell ${String.fromCharCode(65 + j)}3. Value: ${dates[j]}`);
        invalidDateFound = true;
    }
  }

  return total;
}


/**
 * Sums all numeric cells in a specified sheet that are older than a certain number of days.
 *
 * @param {string} sheetName The name of the sheet to search (e.g., "199 Aging Splits").
 * @param {number} daysOld The age threshold in days.
 * @return {number|string} The total sum of the matching cells or an error message.
 * @customfunction
 */
function SUMTOTALOLDERTHAN(sheetName, daysOld) {
  // 1. Argument Validation
  if (!sheetName || typeof sheetName !== 'string') {
    return 'ERROR: First argument (sheetName) is missing or invalid.';
  }
  if (daysOld === null || typeof daysOld !== 'number') { // Allow 0
    return 'ERROR: Second argument (daysOld) must be a number.';
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName);

  if (!targetSheet) {
    return `ERROR: Sheet named "${sheetName}" was not found.`;
  }

  // 2. Data Range and Structure Validation
  const dataRange = targetSheet.getDataRange();
  const allValues = dataRange.getValues();

  if (allValues.length < 3) {
    return `ERROR: Sheet "${sheetName}" has fewer than 3 rows of data.`;
  }

  const dates = allValues[2]; // Dates are in the 3rd row (index 2)
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize today to midnight
  const thresholdDate = new Date(today.getTime() - (daysOld * 24 * 60 * 60 * 1000));

  let total = 0;
  let invalidDateFound = false;

  // Start from column index 3 (Column D) where data begins
  for (let j = 3; j < dates.length; j++) {
    const cellDate = parseSheetDate(dates[j]);

    // Check if the date is valid and on or before the threshold date
    if (cellDate) {
      cellDate.setHours(0, 0, 0, 0); // Normalize cell date to midnight for accurate comparison
      if (cellDate <= thresholdDate) {
        // This column is old enough, sum all numbers in this column
        // Iterate through the relevant rows (Row 4 to 48, which is index 3 to 47)
        for (let i = 3; i < 48; i++) {
          // Ensure the row/column exists in the array to prevent errors
          if (allValues[i] && allValues[i][j] !== undefined && typeof allValues[i][j] === 'number') {
            total += allValues[i][j];
          }
        }
      }
    } else if (dates[j] && !invalidDateFound) {
        // Log only the first invalid date to avoid spamming
        Logger.log(`Skipping invalid date in sheet "${sheetName}", cell ${String.fromCharCode(65 + j)}3. Value: ${dates[j]}`);
        invalidDateFound = true;
    }
  }
  return total;
}

/**
 * Creates a custom menu with options to run the transfer scripts.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Run It!!')
      .addItem('Transfer ALL Totals', 'runBothTransfers')
      .addItem('Transfer +7 Day Totals to Dashboard', 'transferToDashboard')
      .addItem('Outbound Picks', 'updateOutboundPicks')
      .addSeparator()
      .addItem('Transfer 199 Totals Only', 'transfer199Data')
      .addItem('Transfer 191 Totals Only', 'transfer191Data')
      .addToUi();
}

/**
 * A single function to run all transfer processes sequentially.
 */
function runBothTransfers() {
  transfer199Data();
  transfer191Data();
  transferToDashboard();
}

/**
 * Sets up and runs the transfer for the "199 Aging Splits" data.
 */
function transfer199Data() {
  // Configuration specific to the "199" data source and its destination rows
  const config = {
    sourceSheetName: "199 Aging Splits",
    destinationRanges: ["A185:A226", "A236:A242"] // Both regular and DC store ranges
  };
  executeTransfer(config);
}

/**
 * Sets up and runs the transfer for the "191 Aging Splits" data.
 */
function transfer191Data() {
  // Configuration specific to the "191" data source and its destination rows
  const config = {
    sourceSheetName: "191 Aging Splits",
    destinationRanges: ["A337:A379", "A389:A395"] // Both regular and DC store ranges
  };
  executeTransfer(config);
}

/**
 * Transfers the calculated +7 day totals to the main dashboard spreadsheet.
 * This function now activates the source sheet to ensure formulas calculate correctly.
 */
function transferToDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("+7 days");
  const ui = SpreadsheetApp.getUi();

  if (!sourceSheet) {
    // Use a more robust error message for non-interactive contexts.
    throw new Error("Source sheet '+7 days' not found. Cannot proceed with the transfer.");
  }

  try {
    // --- FIX: Activate the sheet to ensure formulas can calculate ---
    sourceSheet.activate();
    // Allow the sheet to recalculate before grabbing values.
    SpreadsheetApp.flush();

    // Helper function to safely get a value and default to 0 if it's an error
    const getSafeValue = (cellA1) => {
      const value = sourceSheet.getRange(cellA1).getValue();
      // Check for standard Google Sheets error strings
      if (typeof value === 'string' && value.startsWith('#')) {
        Logger.log(`Found error "${value}" in cell ${cellA1} on sheet "+7 days". Defaulting to 0.`);
        return 0;
      }
      return value;
    };

    // Read the values from the source sheet safely
    const valuesToTransfer = {
      "199_picking": getSafeValue("B2"),
      "199_shipping": getSafeValue("B3"),
      "199_leftover": getSafeValue("B4"),
      "191_picking": getSafeValue("D2"),
      "191_shipping": getSafeValue("D3"),
      "191_leftover": getSafeValue("D4")
    };

    // Open the destination spreadsheet by its ID
    const destSpreadsheetId = "1aHJQKkmx2ijrPmOeY_syCtu3d1U0j6iNuAGWG7JuvJk";
    const destSpreadsheet = SpreadsheetApp.openById(destSpreadsheetId);
    const destSheet = destSpreadsheet.getSheetByName("Dashboard");

    if (!destSheet) {
      // Use a more robust error message
      throw new Error("Destination sheet 'Dashboard' not found in the target spreadsheet.");
    }

    // Write the values to the specified destination cells
    destSheet.getRange("D28").setValue(valuesToTransfer["199_picking"]);
    destSheet.getRange("D29").setValue(valuesToTransfer["199_shipping"]);
    destSheet.getRange("D30").setValue(valuesToTransfer["199_leftover"]);
    destSheet.getRange("H28").setValue(valuesToTransfer["191_picking"]);
    destSheet.getRange("H29").setValue(valuesToTransfer["191_shipping"]);
    destSheet.getRange("H30").setValue(valuesToTransfer["191_leftover"]);

  } catch (e) {
    Logger.log(`Error in transferToDashboard: ${e.toString()}`);
    // Use ui.alert only if the script is running in a context that has a UI.
    try {
      ui.alert("An error occurred during transfer: " + e.message);
    } catch (uiError) {
      Logger.log("Could not display UI alert. The script was likely run from a trigger or non-interactive context.");
    }
  }
}

/**
 * A robust and reusable function to perform the data transfer.
 * This version maps store locations before writing data to prevent errors.
 * @param {object} config - The configuration object with sheet names and ranges.
 */
function executeTransfer(config) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = config.sourceSheetName;

  try {
    // --- 1. GET SOURCE DATA ---
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      ui.alert(`Source sheet named '${sourceSheetName}' not found.`);
      return;
    }
    const sourceDataRange = sourceSheet.getDataRange();
    const sourceValues = sourceDataRange.getValues();
    const sourceBackgrounds = sourceDataRange.getBackgrounds();

    const headers = sourceValues[2]; // Headers are in row 3
    const grandTotalColumnIndex = headers.indexOf("Grand Total");
    if (grandTotalColumnIndex === -1) {
      ui.alert(`Could not find 'Grand Total' in Row 3 of '${sourceSheetName}'.`);
      return;
    }

    // Create a map of { storeNumber: { volume, color } } from the source data
    const storeDataMap = {};
    // Data is in rows 4 through 48 (index 3 to 47)
    for (let i = 3; i < 52; i++) {
      const storeNumber = sourceValues[i][2]; // Store numbers are in column C
      if (storeNumber) {
        storeDataMap[storeNumber.toString().trim()] = { // Use trim() for safety
          volume: sourceValues[i][grandTotalColumnIndex],
          color: sourceBackgrounds[i][grandTotalColumnIndex]
        };
      }
    }

    // --- 2. PREPARE DESTINATION ---
    const destinationSpreadsheet = SpreadsheetApp.openById("1aHJQKkmx2ijrPmOeY_syCtu3d1U0j6iNuAGWG7JuvJk");
    const destinationSheet = destinationSpreadsheet.getSheetByName("Data");
    if (!destinationSheet) {
      ui.alert(`Destination sheet named 'Data' not found.`);
      return;
    }

    const timeZone = destinationSpreadsheet.getSpreadsheetTimeZone();
    const todayFormatted = Utilities.formatDate(new Date(), timeZone, "M/d/yyyy");

    const dateRow = destinationSheet.getRange("1:1").getValues()[0];
    const todayColumnIndex = dateRow.findIndex(cell =>
        cell instanceof Date && Utilities.formatDate(cell, timeZone, "M/d/yyyy") === todayFormatted
    );

    if (todayColumnIndex === -1) {
      ui.alert(`Could not find today's date (${todayFormatted}) in Row 1 of the 'Data' sheet.`);
      return;
    }
    const targetColumn = todayColumnIndex + 1; // 1-based index for getRange()

    // --- 3. MAP DESTINATION STORE LOCATIONS ---
    // This is the new, more robust logic.
    // It creates a map of { storeNumber: targetRow }
    const storeLocationMap = {};
    config.destinationRanges.forEach(rangeA1 => {
      const range = destinationSheet.getRange(rangeA1);
      const storeNumbers = range.getValues();
      const startRow = range.getRow();
      storeNumbers.forEach((row, index) => {
        const storeNumber = row[0];
        if (storeNumber) {
          // Map the store number to its absolute row number in the sheet
          storeLocationMap[storeNumber.toString().trim()] = startRow + index;
        }
      });
    });

    // --- 4. UPDATE DATA CELL BY CELL (SAFEST METHOD) ---
    // While slightly slower, this guarantees data goes to the correct row.
    let updatedCount = 0;
    for (const storeNumber in storeDataMap) {
      if (storeLocationMap.hasOwnProperty(storeNumber)) {
        const targetRow = storeLocationMap[storeNumber];
        const targetCell = destinationSheet.getRange(targetRow, targetColumn);

        // Only update if the cell is empty
        if (targetCell.getValue() === "") {
          const dataToPaste = storeDataMap[storeNumber];
          targetCell.setValue(dataToPaste.volume)
                    .setBackground(dataToPaste.color);
          updatedCount++;
        }
      }
    }



  } catch (e) {
    Logger.log(e.toString());
    ui.alert("An error occurred: " + e.toString());
  }
}



// CONSTANTS
var SPREADSHEET_ID = '1aHJQKkmx2ijrPmOeY_syCtu3d1U0j6iNuAGWG7JuvJk';
var BLUE_COLOR = '#b4a7d6'; // Assumed blue color. Please change if this is incorrect.

var AGING_SPLITS_191_SHEET_NAME = '191 Aging Splits';
var AGING_SPLITS_199_SHEET_NAME = '199 Aging Splits';
var OUTBOUND_PICKS_SHEET_NAME = 'Outbound Picks';

// Mappings for '191 Aging Splits'
var MAPPING_191 = {
  'Store': 3,
  'Out': 4,
  'HDC': 5,
  'DC': 6
};

// Mappings for '199 Aging Splits'
var MAPPING_199 = {
  'Store': 8,
  'Out': 9,
  'DC': 10
};


function updateOutboundPicks() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    var outboundPicksSheet = ss.getSheetByName(OUTBOUND_PICKS_SHEET_NAME);
    if (!outboundPicksSheet) {
      throw new Error('Sheet "' + OUTBOUND_PICKS_SHEET_NAME + '" not found.');
    }

    var today = new Date();
    var dateColumn = findDateColumn(outboundPicksSheet, today);

    if (dateColumn === -1) {
      SpreadsheetApp.getUi().alert('Today\'s date column was not found in \'Outbound Picks\'.');
      return;
    }

    // Process '191 Aging Splits'
    processAgingSheet(ss, AGING_SPLITS_191_SHEET_NAME, MAPPING_191, outboundPicksSheet, dateColumn);

    // Process '199 Aging Splits'
    processAgingSheet(ss, AGING_SPLITS_199_SHEET_NAME, MAPPING_199, outboundPicksSheet, dateColumn);



  } catch (e) {
    Logger.log(e.toString());
    SpreadsheetApp.getUi().alert('An error occurred: ' + e.toString());
  }
}

/**
 * A test function that calculates totals and logs them without modifying the sheet.
 */
function testUpdateOutboundPicks() {
  Logger.log('--- Starting Test Run ---');
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    Logger.log('Calculating totals for: ' + AGING_SPLITS_191_SHEET_NAME);
    var totals191 = calculateAgingSheetTotals(ss, AGING_SPLITS_191_SHEET_NAME, MAPPING_191);
    Logger.log('Totals for 191: ' + JSON.stringify(totals191, null, 2));

    Logger.log('Calculating totals for: ' + AGING_SPLITS_199_SHEET_NAME);
    var totals199 = calculateAgingSheetTotals(ss, AGING_SPLITS_199_SHEET_NAME, MAPPING_199);
    Logger.log('Totals for 199: ' + JSON.stringify(totals199, null, 2));

    Logger.log('--- Test Run Complete ---');
    Logger.log('Review the totals logged above. No data has been written to the spreadsheet.');

  } catch (e) {
    Logger.log('An error occurred during test: ' + e.toString());
  }
}

/**
 * Finds the column number for a given date in the first row of a sheet.
 * @param {Sheet} sheet The sheet to search in.
 * @param {Date} date The date to find.
 * @return {number} The column number (1-indexed) or -1 if not found.
 */
function findDateColumn(sheet, date) {
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var values = range.getValues()[0];

  for (var i = 0; i < values.length; i++) {
    if (values[i] instanceof Date) {
      if (values[i].getFullYear() === date.getFullYear() &&
          values[i].getMonth() === date.getMonth() &&
          values[i].getDate() === date.getDate()) {
        return i + 1;
      }
    }
  }
  return -1;
}

/**
 * Processes an aging splits sheet and updates the outbound picks sheet.
 * @param {Spreadsheet} ss The spreadsheet object.
 * @param {string} sheetName The name of the aging sheet to process.
 * @param {Object} mapping The mapping of categories to row numbers.
 * @param {Sheet} outboundSheet The destination sheet.
 * @param {number} dateColumn The column to write the data to.
 */
function processAgingSheet(ss, sheetName, mapping, outboundSheet, dateColumn) {
  var totals = calculateAgingSheetTotals(ss, sheetName, mapping);
  if (totals) {
    for (var category in totals) {
      var row = mapping[category];
      outboundSheet.getRange(row, dateColumn).setValue(totals[category]);
    }
  }
}

/**
 * Calculates the totals from an aging splits sheet based on category and color.
 * @param {Spreadsheet} ss The spreadsheet object.
 * @param {string} sheetName The name of the aging sheet to process.
 * @param {Object} mapping The mapping of categories to row numbers.
 * @return {Object|null} An object with totals per category, or null if sheet not found.
 */
function calculateAgingSheetTotals(ss, sheetName, mapping) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('Sheet "' + sheetName + '" not found. Skipping.');
    return null;
  }

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var backgroundColors = dataRange.getBackgrounds();

  // Find the 'Grand Total' column
  var grandTotalCol = -1;
  var headerRow = values[2]; // Assuming 'Grand Total' is in row 3
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === 'Grand Total') {
      grandTotalCol = i;
      break;
    }
  }

  if (grandTotalCol === -1) {
    Logger.log('Could not find "Grand Total" column in ' + sheetName);
    return null;
  }

  var totals = {};
  for (var category in mapping) {
    totals[category] = 0;
  }

  for (var r = 3; r < values.length; r++) { // Start from row 4 (index 3)
    var category = values[r][0]; // Column A
    var color = backgroundColors[r][0]; // Column A background

    if (color === BLUE_COLOR && mapping[category]) {
      var storeNumber = values[r][2]; // Column C
      if (storeNumber) { // Make sure there's a store number
          var totalValue = values[r][grandTotalCol];
          if (typeof totalValue === 'number') {
            totals[category] += totalValue;
          }
      }
    }
  }
  return totals;
}
function createOpenTrigger() {
  ScriptApp.newTrigger("myOnOpen")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

function myOnOpen(e) {
  const sheet = SpreadsheetApp.openById('1aHJQKkmx2ijrPmOeY_syCtu3d1U0j6iNuAGWG7JuvJk');
  const sheetName = sheet.getSheetByName(AGING_SPLITS_191_SHEET_NAME);
  Logger.log('Sheet accessed: ' + sheetName.getName());
}