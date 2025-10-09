// --- Existing Script for Tracking Individual Loads ---

function trackAndLogLoadsWithRetry() {
  var maxRetries = 3;
  for (var i = 0; i < maxRetries; i++) {
    try {
      trackAndLogLoads();
      return; // If the script runs successfully, we exit the function.
    } catch (e) {
      console.error('Attempt ' + (i + 1) + ' failed: ' + e.message);
      if (i < maxRetries - 1) {
        Utilities.sleep(3000); // Wait for 3 seconds before trying again.
      } else {
        console.error('Script failed after ' + maxRetries + ' attempts.');
        throw e; // If all retries fail, stop the script and show the error.
      }
    }
  }
}

function trackAndLogLoads() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allLoadsSheet = ss.getSheetByName('All Loads');
  var targetSpreadsheetId = "1IwvKjBXrCDYnK0x1OHG_vwk_9nlwh4Qne59aFzbGC5E";

  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  var inboundLogSheet = targetSpreadsheet.getSheetByName('Inbound Loads Log');
  var unloadedLogSheet = targetSpreadsheet.getSheetByName('Unloaded Loads Log');

  var dataRange = allLoadsSheet.getDataRange();
  var allLoadsData = dataRange.getValues();

  var scriptProperties = PropertiesService.getScriptProperties();
  var processedLoadsString = scriptProperties.getProperty('processedLoads');
  var processedLoads = processedLoadsString ? JSON.parse(processedLoadsString) : {};

  var currentLoadIDs = {};
  var timestamp = new Date();
  var dateFormat = "M/d/yy";

  for (var i = 2; i < allLoadsData.length; i++) {
    var row = allLoadsData[i];
    var trailerNumber = row[3];
    if (!trailerNumber) continue;
    currentLoadIDs[trailerNumber] = true;
    if (!processedLoads.hasOwnProperty(trailerNumber)) {
      var arrivalDateValue = row[7];
      var inboundType = row[2];
      var loadType = row[5];
      var arrivalDate = new Date(arrivalDateValue);
      if (!isNaN(arrivalDate.getTime())) {
        var formattedArrivalDate = Utilities.formatDate(arrivalDate, Session.getScriptTimeZone(), dateFormat);
        inboundLogSheet.appendRow([formattedArrivalDate, trailerNumber, inboundType, loadType, timestamp]);
        processedLoads[trailerNumber] = { arrivalDate: formattedArrivalDate, inboundType: inboundType, loadType: loadType };
      }
    }
  }

  for (var trailerNumber in processedLoads) {
    if (!currentLoadIDs.hasOwnProperty(trailerNumber)) {
      var unloadedLoadData = processedLoads[trailerNumber];
      unloadedLogSheet.appendRow([ unloadedLoadData.arrivalDate, trailerNumber, unloadedLoadData.inboundType, unloadedLoadData.loadType, timestamp ]);
      delete processedLoads[trailerNumber];
    }
  }
  scriptProperties.setProperty('processedLoads', JSON.stringify(processedLoads));
}

// --- FINAL CORRECTED Script for Daily Snapshot Count ---

// IDs of the spreadsheets
const SOURCE_SPREADSHEET_ID = '1HcNv1RbNQ-3f-YMmIPBTFjK6KBeaITjqbg1zdHsstRs';
const DESTINATION_SPREADSHEET_ID = '1uaJ3TPBEtC650JQKKWxM3SNPuSTtCbYtrOzU-Vumtvc';
const SOURCE_SHEET_NAME = 'All Loads';
const DESTINATION_SHEET_NAME = 'Data';
const TRAILER_TYPES_RANGE = 'A54:A59';

/**
 * FINAL CORRECTED VERSION: Finds today's date in row 1 of the destination sheet by comparing
 * formatted date strings (to be timezone-proof) and writes the counts into the corresponding column.
 */
function takeDailySnapshot() {
  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const destinationSpreadsheet = SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);

  const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
  const destinationSheet = destinationSpreadsheet.getSheetByName(DESTINATION_SHEET_NAME);

  if (!sourceSheet || !destinationSheet) {
    Logger.log("Source or destination sheet not found.");
    return;
  }

  // --- Counting logic (same as before) ---
  const trailerTypes = destinationSheet.getRange(TRAILER_TYPES_RANGE).getValues().flat().filter(String);
  if (trailerTypes.length === 0) {
    Logger.log(`No trailer types found in range ${TRAILER_TYPES_RANGE}.`);
    return;
  }
  const lastRowSource = sourceSheet.getLastRow();
  const inboundData = lastRowSource > 0 ? sourceSheet.getRange(`F1:F${lastRowSource}`).getValues().flat() : [];
  const counts = new Map();
  trailerTypes.forEach(type => counts.set(type, 0));
  inboundData.forEach(inbound => {
    if (counts.has(inbound)) {
      counts.set(inbound, counts.get(inbound) + 1);
    }
  });
  const columnData = trailerTypes.map(type => [counts.get(type) || 0]);

  // --- FINAL logic to find the correct column using formatted strings ---
  const datesRow = destinationSheet.getRange(1, 1, 1, destinationSheet.getLastColumn()).getValues()[0];
  // Get the timezone of the spreadsheet itself to ensure consistency
  const spreadsheetTimezone = destinationSpreadsheet.getSpreadsheetTimeZone();
  const todayFormatted = Utilities.formatDate(new Date(), spreadsheetTimezone, "yyyy-MM-dd");

  let targetColumn = -1;
  for (let i = 0; i < datesRow.length; i++) {
    if (datesRow[i] instanceof Date) {
      const sheetDateFormatted = Utilities.formatDate(datesRow[i], spreadsheetTimezone, "yyyy-MM-dd");
      if (sheetDateFormatted === todayFormatted) {
        targetColumn = i + 1; // Column index is 1-based
        break;
      }
    }
  }

  if (targetColumn !== -1) {
    destinationSheet.getRange(54, targetColumn, columnData.length, 1).setValues(columnData);
    Logger.log(`Daily snapshot successfully written to column ${targetColumn}.`);
  } else {
    Logger.log(`Could not find today's date (${todayFormatted}) in row 1 of the destination sheet. Please ensure the date exists and is formatted as a date.`);
  }
}

/**
 * A test function for the final corrected daily snapshot.
 */
function test_takeDailySnapshot() {
    try {
        Logger.log("--- SNAPSHOT TEST RUN (Final Logic) ---");
        Logger.log("This test will check if it can find today's date in row 1 by comparing formatted date strings.");
        takeDailySnapshot(); // Running the actual function to see its logs
    } catch (e) {
        Logger.log("Error during snapshot test run: " + e.toString());
    }
}