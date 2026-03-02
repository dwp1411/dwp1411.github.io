/**
 * Generates a PDF from a specified sheet and sends it as an email attachment.
 * Now includes better formatting and accepts a custom filename.
 */
function sendSheetAsPdf(sheetName, emailList, subject, body, pdfName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    const allSheetNames = ss.getSheets().map(s => `'${s.getName()}'`);
    throw new Error(`Sheet named "${sheetName}" not found. Available sheets are: ${allSheetNames.join(", ")}.`);
  }

  if (!emailList) {
    throw new Error('No recipient email addresses provided. Email not sent.');
  }

  try {
    const spreadSheetId = ss.getId();
    const sheetId = sheet.getSheetId();

   // *** MODIFIED URL for better formatting ***
    const url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/export?format=pdf&gid=${sheetId}` +
                '&portrait=true' +  // MUST be lowercase 'false' for landscape
                '&scale=2' +         // 4 = Fit to Page
                '&size=letter' +     // US Letter paper size
                '&gridlines=false';  // Do not show gridlines
    const params = {
      method: "GET",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    };

    // *** MODIFIED to use the custom PDF filename ***
    const pdfBlob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName);

    // Send the email
    MailApp.sendEmail({
      to: emailList,
      subject: subject,
      body: body,
      attachments: [pdfBlob]
    });

  } catch (e) {
    Logger.log(e.toString());
    // Re-throw the error so the calling function can handle it and show a UI alert.
    throw new Error("An error occurred while creating or sending the PDF: " + e.toString());
  }
}

/**
 * A wrapper function to be called by a button in the Google Sheet.
 * This function contains the configuration for the email.
 */
function onSendEmailButtonClick() {
  const ui = SpreadsheetApp.getUi();

  // --- Get today's date and format it ---
  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd-yyyy");

  // --- CONFIGURATION ---
  const SHEET_TO_SEND = "Dashboard"; // The name of the sheet you want to send
  const RECIPIENT_EMAILS = "NapervilleDCManagers@crateandbarrel.com,NapervilleDCSupervisors@crateandbarrel.com,NapervilleDCLeads@crateandbarrel.com,rymartinez@crateandbarrel.com,mbellich@crateandbarrel.com"; // Comma-separated list of emails

  // *** MODIFIED to include the date in the subject and define a filename ***
  const EMAIL_SUBJECT = `Naperville DBR ${formattedDate}`;
  const PDF_FILENAME = `Naperville DBR Report ${formattedDate}.pdf`;

  const EMAIL_BODY = "Please find the attached report with today's info.";
  // ---------------------

  try {


    // *** MODIFIED to pass the new PDF filename to the sending function ***
    sendSheetAsPdf(SHEET_TO_SEND, RECIPIENT_EMAILS, EMAIL_SUBJECT, EMAIL_BODY, PDF_FILENAME);


  } catch (e) {
    // Log the full error for debugging and show a user-friendly message.
    Logger.log(e.toString());
    ui.alert('Failed to send email. Error: ' + e.message);
  }
}

// --- SCRIPT to find dates in Row 1 and populate a vertical forecast for Naperville, IL ---

/**
 * Main function to update a 5-day forecast on the 'Data' sheet.
 * This script now uses the 5 Day / 3 Hour Forecast API from OpenWeatherMap,
 * as the One Call 3.0 API requires a paid subscription.
 */
function updateForecastOnDataSheet() {

  // --- Configuration ---
  const apiKey = 'e1e607932638179e28607670b0abb14d';

  // *** LOCATION UPDATED HERE ***
  const lat = '41.7733'; // Naperville, IL
  const lon = '-88.1446'; // Naperville, IL

  const dataSheetName = "Data";
  const dateRowNumber = 1; // The row where your dates are

  if (apiKey === 'YOUR_OPENWEATHERMAP_KEY_HERE' || apiKey === '') {
    SpreadsheetApp.getUi().alert('Please enter your OpenWeatherMap API key in the script.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(dataSheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet named '${dataSheetName}' not found.`);
    return;
  }

  // --- Step 1: Read all the dates from Row 1 ---
  const dateRange = sheet.getRange(dateRowNumber, 1, 1, sheet.getMaxColumns());
  const datesFromSheet = dateRange.getValues()[0];

  // --- Step 2: Fetch the 5-day forecast from the API ---
  const apiUrl = `https://api.openweathermap.org/data/2.5/forecast?lat=${lat}&lon=${lon}&appid=${apiKey}&units=imperial`;

  try {
    const response = UrlFetchApp.fetch(apiUrl);
    const data = JSON.parse(response.getContentText());

    if (data.cod !== "200") {
      throw new Error(`API Error: ${data.message || 'Unknown error'}`);
    }

    // --- Process 3-hourly data into daily forecasts ---
    const dailyData = {}; // Use an object to store daily aggregated data

    data.list.forEach(item => {
      const date = new Date(item.dt * 1000);
      const formattedDate = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

      if (!dailyData[formattedDate]) {
        // Initialize data for a new day
        dailyData[formattedDate] = {
          temps: [],
          pops: [],
          winds: [],
          humidities: [],
          icons: [],
          descriptions: [],
          dt: item.dt // Store the timestamp for sorting later
        };
      }

      // Collect data points for the day
      dailyData[formattedDate].temps.push(item.main.temp);
      dailyData[formattedDate].pops.push(item.pop);
      dailyData[formattedDate].winds.push(item.wind.speed);
      dailyData[formattedDate].humidities.push(item.main.humidity);
      dailyData[formattedDate].icons.push(item.weather[0].icon);
      dailyData[formattedDate].descriptions.push(item.weather[0].description);
    });

    // --- Create forecastDays array from aggregated data ---
    const forecastDays = Object.keys(dailyData).map(dateStr => {
      const dayData = dailyData[dateStr];

      const temp_min = Math.min(...dayData.temps);
      const temp_max = Math.max(...dayData.temps);
      const pop = Math.max(...dayData.pops);
      const wind_speed = Math.max(...dayData.winds);
      const humidity = dayData.humidities.reduce((a, b) => a + b, 0) / dayData.humidities.length;

      // For icon and description, pick the one for the middle of the day (around noon).
      const middayIndex = Math.floor(dayData.icons.length / 2);
      const icon = dayData.icons[middayIndex] || dayData.icons[0];
      const description = dayData.descriptions[middayIndex] || dayData.descriptions[0];

      // Capitalize description
      const summary = description.charAt(0).toUpperCase() + description.slice(1);

      return {
        dt: dayData.dt,
        temp: { min: temp_min, max: temp_max },
        pop: pop,
        wind_speed: wind_speed,
        humidity: humidity,
        weather: [{ icon: icon }],
        summary: summary
      };
    }).slice(0, 5); // Ensure we only take 5 days

    // --- Step 3: Loop through the forecast, find the matching date column, and write the data ---
    forecastDays.forEach(day => {
      const forecastDate = new Date(day.dt * 1000);
      const formattedApiDate = Utilities.formatDate(forecastDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");

      let columnIndex = -1;
      for (let i = 0; i < datesFromSheet.length; i++) {
        if (datesFromSheet[i] instanceof Date) {
          const sheetDate = new Date(datesFromSheet[i]);
          const formattedSheetDate = Utilities.formatDate(sheetDate, ss.getSpreadsheetTimeZone(), "M/d/yyyy");
          if (formattedSheetDate === formattedApiDate) {
            columnIndex = i;
            break;
          }
        }
      }

      if (columnIndex !== -1) {
        const columnNumber = columnIndex + 1; // Sheet columns are 1-indexed

        const iconUrl = `=IMAGE("https://openweathermap.org/img/wn/${day.weather[0].icon}@2x.png")`;
        const tempRange = `${Math.round(day.temp.max)}° / ${Math.round(day.temp.min)}°`;
        const chanceOfRain = `${Math.round(day.pop * 100)}%`;
        const wind = `${Math.round(day.wind_speed)} mph`;
        const humidityVal = `${Math.round(day.humidity)}%`;

        sheet.getRange(dateRowNumber + 1, columnNumber).setFormula(iconUrl);
        sheet.getRange(dateRowNumber + 2, columnNumber).setValue(day.summary);
        sheet.getRange(dateRowNumber + 3, columnNumber).setValue(tempRange);
        sheet.getRange(dateRowNumber + 4, columnNumber).setValue(chanceOfRain);
        sheet.getRange(dateRowNumber + 5, columnNumber).setValue(wind);
        sheet.getRange(dateRowNumber + 6, columnNumber).setValue(humidityVal);
      }
    });

    SpreadsheetApp.getUi().alert("5-day forecast has been updated on the 'Data' tab for Naperville, IL.");

  } catch (e) {
    Logger.log(e.toString());
    SpreadsheetApp.getUi().alert("An error occurred fetching the weather. Check logs for details: " + e.toString());
  }
}
/**
 * Calculates a sum, count, or average for a range of cells based on a specified background color.
 *
 * @param {range} inputRange The range of cells to evaluate (e.g., A1:A10).
 * @param {range} colorReferenceCell The single cell whose background color you want to match.
 * @param {string} calculationType Optional. The type of calculation: "SUM" (default), "COUNT", or "AVERAGE".
 * @return The calculated result.
 * @customfunction
 */
function CALCULATEBYCOLOR(inputRange, colorReferenceCell, calculationType) {
  // Check for null or empty arguments
  if (!inputRange || !colorReferenceCell) {
    throw new Error("Error: inputRange and colorReferenceCell arguments cannot be null or empty.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetColor = sheet.getRange(colorReferenceCell).getBackground();
  const range = sheet.getRange(inputRange);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  let sum = 0;
  let count = 0;

  for (let i = 0; i < backgrounds.length; i++) {
    for (let j = 0; j < backgrounds[i].length; j++) {
      if (backgrounds[i][j] === targetColor) {
        if (typeof values[i][j] === 'number') {
          sum += values[i][j];
        }
        count++;
      }
    }
  }

  const calcType = (calculationType || "SUM").toUpperCase();
  switch (calcType) {
    case "COUNT":
      return count;
    case "AVERAGE":
      return count === 0 ? 0 : sum / count;
    case "SUM":
    default:
      return sum;
  }
}
/**
 * @fileoverview This script automates the process of logging a daily total
 * cost from a source spreadsheet to a destination spreadsheet. It is designed
 * to be run on a daily trigger.
 */

/**
 * The ID of the spreadsheet where the daily cost is calculated.
 * This is the "calculator" sheet.
 * @const {string}
 */
const SOURCE_SPREADSHEET_ID = '1Pv6CRX-YvPKq4moiDPJ2U7imsj5Oix3vEnCjuncKI0w';

/**
 * The name of the sheet within the source spreadsheet that contains the total.
 * @const {string}
 */
const SOURCE_SHEET_NAME = 'Sheet1';

/**
 * The cell that contains the total daily detention cost.
 * @const {string}
 */
const SOURCE_CELL = 'O3';

/**
 * The name of the sheet in *this* spreadsheet where the data will be logged.
 * @const {string}
 */
const DESTINATION_SHEET_NAME = 'Data';

/**
 * The row number (1-indexed) where the dates are located.
 * @const {number}
 */
const DATE_ROW = 1;

/**
 * The row number (1-indexed) where the daily cost should be written.
 * @const {number}
 */
const TARGET_ROW = 76;

/**
 * The main function to log the daily total cost. This function will:
 * 1. Fetch the total cost from the source spreadsheet.
 * 2. Find the column in the destination sheet that corresponds to today's date.
 * 3. Write the total cost into the correct cell in the destination sheet.
 * This function is intended to be run daily via a trigger.
 */
function logDailyTotalCost() {
  const ui = SpreadsheetApp.getUi();

  try {
    // --- Step 1: Get the total cost from the source sheet ---
    const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      throw new Error(`Sheet "${SOURCE_SHEET_NAME}" not found in the source spreadsheet.`);
    }

    const totalCost = sourceSheet.getRange(SOURCE_CELL).getValue();

    // Validate that the cost is a number.
    if (typeof totalCost !== 'number') {
      throw new Error(`The value in cell ${SOURCE_CELL} ("${totalCost}") is not a valid number.`);
    }

    // --- Step 2: Find the correct column in the destination sheet ---
    const destinationSpreadsheet = SpreadsheetApp.openById('1aHJQKkmx2ijrPmOeY_syCtu3d1U0j6iNuAGWG7JuvJk');
    const destinationSheet = destinationSpreadsheet.getSheetByName(DESTINATION_SHEET_NAME);
    if (!destinationSheet) {
      throw new Error(`Sheet "${DESTINATION_SHEET_NAME}" not found in this spreadsheet.`);
    }

    const dateRowValues = destinationSheet.getRange(DATE_ROW, 1, 1, destinationSheet.getLastColumn()).getValues()[0];

    // Get today's date and normalize it to the beginning of the day for a clean comparison.
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let targetColumn = -1;

    // Loop through the header row to find a date that matches today.
    for (let i = 0; i < dateRowValues.length; i++) {
      const cellValue = dateRowValues[i];
      if (cellValue instanceof Date) {
        // Also normalize the date from the sheet.
        const sheetDate = new Date(cellValue);
        sheetDate.setHours(0, 0, 0, 0);

        if (sheetDate.getTime() === today.getTime()) {
          targetColumn = i + 1; // getRange is 1-indexed, so add 1.
          break;
        }
      }
    }

    if (targetColumn === -1) {
      throw new Error(`Today's date (${today.toLocaleDateString()}) was not found in row ${DATE_ROW}.`);
    }

    // --- Step 3: Write the cost to the correct cell ---
    destinationSheet.getRange(TARGET_ROW, targetColumn).setValue(totalCost);

    ui.alert('Success', `Successfully logged detention cost of ${totalCost} for today.`, ui.ButtonSet.OK);

  } catch (e) {
    // If anything goes wrong, log the error and show an alert to the user.
    console.error(e);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Adds a custom menu to the spreadsheet UI when the file is opened.
 * This allows the user to manually trigger the logging script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Daily Cost Logger')
      .addItem('Log Today\'s Cost', 'logDailyTotalCost')
      .addToUi();
}
