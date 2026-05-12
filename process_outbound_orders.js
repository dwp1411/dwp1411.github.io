function generateRetailFurnitureOutboundReport() {
  const mainDataSheetId = '1GTGuDDiY0PwYymhRaxMHyCLKTRFjYo7fX6BfPHumF9w';
  const storesSheetId = '1UNfyB49dGQHGkJQIEWqtH62HIUOEQSXRWEuplscP4nA';
  const outboundTrackerId = '18MMQBgiZY8MqBVX4qQHLTxYyUbytTeglv-QmJ124_gQ';

  // 1. Fetch Main Data
  const mainSpreadsheet = SpreadsheetApp.openById(mainDataSheetId);
  const mainSheet = mainSpreadsheet.getSheets()[0]; // Assuming first sheet
  const mainData = mainSheet.getDataRange().getValues();
  const mainHeaders = mainData[0];

  const colSplitDate = mainHeaders.indexOf('SPLIT_DATE');
  const colShipFrom = mainHeaders.indexOf('SHIP_FROM');
  const colShipTo = mainHeaders.indexOf('SHIP_TO');
  const colQty = mainHeaders.indexOf('QTY');

  // 2. Fetch Stores Data
  const storesSpreadsheet = SpreadsheetApp.openById(storesSheetId);
  const storesSheet = storesSpreadsheet.getSheets()[0];
  const storesData = storesSheet.getDataRange().getValues();

  const stores191 = [];
  const stores199 = [];

  for (let i = 1; i < storesData.length; i++) {
    const row = storesData[i];
    if (row[0] && row[1]) {
      stores191.push({ type: String(row[0]).trim().toUpperCase(), store: String(row[1]).trim() });
    }
    if (row[3] && row[4]) {
      stores199.push({ type: String(row[3]).trim().toUpperCase(), store: String(row[4]).trim() });
    }
  }

  // 3. Fetch Outbound Tracker Data
  const trackerSpreadsheet = SpreadsheetApp.openById(outboundTrackerId);
  const trackerSheet = trackerSpreadsheet.getSheetByName('DC 191_199 -- From DC');
  const trackerData = trackerSheet.getDataRange().getValues();

  // Find column indices in tracker
  const trackerHeaders = trackerData[1]; // Row 2 has headers
  const colDCPickDate = trackerHeaders.indexOf('DC Pick Date'); // A
  const colOriginId = trackerHeaders.indexOf('Origin ID'); // B
  const colDestinationId = trackerHeaders.indexOf('Destination ID'); // E
  const colCarrierReadyDay = trackerHeaders.indexOf('Carrier Ready Day'); // Z

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayNames = ["SUN", "MON", "TUES", "WED", "THURS", "FRI", "SAT"];
  const todayDayName = dayNames[today.getDay()];

  const pickingToday = { '191': new Set(), '199': new Set() };
  const shippingToday = { '191': new Set(), '199': new Set() };

  let currentPickDate = null;

  for (let i = 2; i < trackerData.length; i++) { // Start from row 3 (data)
    const row = trackerData[i];

    // Fill down the pick date
    let rawPickDate = row[colDCPickDate];
    if (rawPickDate && rawPickDate !== '') {
      if (typeof rawPickDate === 'string' && rawPickDate.includes('Total')) {
        continue; // Skip total rows
      }
      currentPickDate = new Date(rawPickDate);
      currentPickDate.setHours(0, 0, 0, 0);
    }

    if (!currentPickDate) continue;

    const origin = String(row[colOriginId]).trim();
    const destination = String(row[colDestinationId]).trim();
    const readyDay = String(row[colCarrierReadyDay]).trim().toUpperCase();

    if (origin !== '191' && origin !== '199') continue;

    // Check if picking today
    if (currentPickDate.getTime() === today.getTime()) {
      pickingToday[origin].add(destination);
    }

    // Check if shipping today (lookback 14 days max for safety to avoid matching future/past weeks incorrectly)
    const diffTime = today.getTime() - currentPickDate.getTime();
    const diffDaysTracker = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    if (diffDaysTracker >= 0 && diffDaysTracker <= 14) {
      if (readyDay === todayDayName || (todayDayName === 'TUE' && readyDay === 'TUES') || (todayDayName === 'THU' && readyDay === 'THURS')) {
         shippingToday[origin].add(destination);
      }
    }
  }

  // 4. Process Main Data
  const reportData = {
    '191': {},
    '199': {}
  };

  // Initialize with stores from the reference sheet
  stores191.forEach(s => {
    reportData['191'][s.store] = { type: s.type, days: {}, picking: 0, shipping: 0, intercompany: 0, hdc: 0, outlet: 0 };
  });
  stores199.forEach(s => {
    reportData['199'][s.store] = { type: s.type, days: {}, picking: 0, shipping: 0, intercompany: 0, hdc: 0, outlet: 0 };
  });

  let maxDaysDiff = 0;
  let uniqueDaysDiff = new Set();

  for (let i = 1; i < mainData.length; i++) {
    const row = mainData[i];
    const shipFrom = String(row[colShipFrom]).trim();
    const shipTo = String(row[colShipTo]).trim();
    const qty = Number(row[colQty]) || 0;
    const splitDateStr = row[colSplitDate];

    if (shipFrom !== '191' && shipFrom !== '199') continue;
    if (!splitDateStr) continue;

    const splitDate = new Date(splitDateStr);
    splitDate.setHours(0,0,0,0);

    const diffTimeMain = today.getTime() - splitDate.getTime();
    let diffDays = Math.ceil(diffTimeMain / (1000 * 60 * 60 * 24));

    // Filter out future dates, or if they just show as 0 days ago (or Negative) group them to 0 or 1
    if (diffDays < 1) diffDays = 1; // Minimum 1 day as per standard lookback or "Today"

    if (diffDays > maxDaysDiff) maxDaysDiff = diffDays;
    uniqueDaysDiff.add(diffDays);

    // Ensure store exists in our report map (even if not in the master list, although we should only list what's in the list ideally - but let's add it to not lose data)
    if (!reportData[shipFrom][shipTo]) {
      reportData[shipFrom][shipTo] = { type: 'Store', days: {}, picking: 0, shipping: 0, intercompany: 0, hdc: 0, outlet: 0 };
    }

    if (!reportData[shipFrom][shipTo].days[diffDays]) {
      reportData[shipFrom][shipTo].days[diffDays] = 0;
    }
    reportData[shipFrom][shipTo].days[diffDays] += qty;
  }

  // Calculate POS logic based on flags
  for (const origin of ['191', '199']) {
    for (const store in reportData[origin]) {
      const data = reportData[origin][store];

      // Calculate total pieces for this store
      let totalPieces = 0;
      for (const d in data.days) {
        totalPieces += data.days[d];
      }

      if (pickingToday[origin].has(store)) {
        data.picking = totalPieces;
      }
      if (shippingToday[origin].has(store)) {
        data.shipping = totalPieces;
      }

      const type = data.type.toUpperCase();
      if (type === 'DC') {
        data.intercompany = totalPieces;
      } else if (type === 'HDC') {
        data.hdc = totalPieces;
      } else if (type === 'OUT') {
        data.outlet = totalPieces;
      }
    }
  }

  // 5. Format and generate output
  // We'll create a new spreadsheet (or you can set this to write to a specific ID)
  // To avoid creating a new file every time, we typically would search by name or you can provide the ID.
  // For the script, we create a new file "Retail/Furniture Outbound" or overwrite.

  let newSpreadsheet;
  const targetName = "Retail/Furniture Outbound";

  // Find if it exists to reuse, else create
  const files = DriveApp.getFilesByName(targetName);
  if (files.hasNext()) {
    newSpreadsheet = SpreadsheetApp.open(files.next());
  } else {
    newSpreadsheet = SpreadsheetApp.create(targetName);
  }

  // Prepare dynamic day headers (sorted descending so oldest is first)
  const sortedDays = Array.from(uniqueDaysDiff).sort((a, b) => b - a);

  for (const origin of ['191', '199']) {
    let sheet = newSpreadsheet.getSheetByName(origin);
    if (!sheet) {
      sheet = newSpreadsheet.insertSheet(origin);
    }
    sheet.clear();

    // Headers Row 1 (Empty placeholders for row 1, dates in row 2)
    // Actually, in the image:
    // Row 1: [Location Type] [Store #] [18 days] [15 days] ... [Pos Picking Today] [Pos shipping Today] [Pos Intercompany] [Pos HDC] [Pos Outlet]
    // Row 2: Dates corresponding to those days.

    const headers1 = ['Location Type', 'Store #'];
    const headers2 = ['', ''];

    for (const d of sortedDays) {
      headers1.push(`${d} days`);
      // Calculate the date
      const dateD = new Date(today.getTime() - d * 24 * 60 * 60 * 1000);
      headers2.push(Utilities.formatDate(dateD, Session.getScriptTimeZone(), "M/d/yyyy"));
    }

    headers1.push('Grand Total', 'Pos Picking Today', 'Pos shipping Today', 'Pos Intercompany', 'Pos HDC', 'Pos Outlet');
    headers2.push('', '', '', '', '', '');

    const outputData = [headers1, headers2];
    const storeKeys = Object.keys(reportData[origin]).sort((a,b) => {
        // Sort numerically if possible, otherwise string
        const numA = parseInt(a, 10);
        const numB = parseInt(b, 10);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        return a.localeCompare(b);
    });

    for (const store of storeKeys) {
      const data = reportData[origin][store];
      const row = [data.type, store];

      let storeTotal = 0;
      for (const d of sortedDays) {
        const val = data.days[d] || '';
        row.push(val);
        storeTotal += (data.days[d] || 0);
      }

      row.push(storeTotal); // Grand Total
      row.push(data.picking || '');
      row.push(data.shipping || '');
      row.push(data.intercompany || '');
      row.push(data.hdc || '');
      row.push(data.outlet || '');

      outputData.push(row);
    }

    // Summary Rows
    // We need totals at the bottom
    const totalStoresRow = ['Total for Stores', ''];
    const totalIntercompanyRow = ['Total for Intercompany', ''];
    const totalHDCRow = ['Total for HDC', ''];
    const totalOutletRow = ['Total Outlet', ''];

    const totals = {
       stores: new Array(sortedDays.length).fill(0),
       intercompany: new Array(sortedDays.length).fill(0),
       hdc: new Array(sortedDays.length).fill(0),
       outlet: new Array(sortedDays.length).fill(0)
    };

    for (const store of storeKeys) {
       const data = reportData[origin][store];
       const type = data.type.toUpperCase();
       let targetArr = totals.stores;

       if (type === 'DC') targetArr = totals.intercompany;
       else if (type === 'HDC') targetArr = totals.hdc;
       else if (type === 'OUT' || type === 'OUTLET') targetArr = totals.outlet;

       for (let idx = 0; idx < sortedDays.length; idx++) {
          targetArr[idx] += (data.days[sortedDays[idx]] || 0);
       }
    }

    const sumArray = (arr) => arr.reduce((a,b)=>a+b, 0);

    totalStoresRow.push(...totals.stores, sumArray(totals.stores), '', '', '', '', '');
    totalIntercompanyRow.push(...totals.intercompany, sumArray(totals.intercompany), '', '', '', '', '');
    totalHDCRow.push(...totals.hdc, sumArray(totals.hdc), '', '', '', '', '');
    totalOutletRow.push(...totals.outlet, sumArray(totals.outlet), '', '', '', '', '');

    const expectedLength = outputData[0].length; outputData.push(new Array(expectedLength).fill(''));
    const grandTotalRow = ['Grand Total', '']; while(grandTotalRow.length < expectedLength) { grandTotalRow.push(''); } outputData.push(grandTotalRow); // Adjust formula via app script
    outputData.push(totalStoresRow);
    outputData.push(totalIntercompanyRow);
    outputData.push(totalHDCRow);
    outputData.push(totalOutletRow);

    // Percentages
    const pctStoresRow = ['Percent to total for Stores', ''];
    const pctIntercompanyRow = ['Percent to total for Intercompany', ''];
    const pctHDCRow = ['Percent to total for HDC', ''];
    const pctOutletRow = ['Percent to total for Outlet', ''];

    for (let i = 0; i < sortedDays.length; i++) {
        const grandTotalCol = totals.stores[i] + totals.intercompany[i] + totals.hdc[i] + totals.outlet[i];
        pctStoresRow.push(grandTotalCol ? totals.stores[i] / grandTotalCol : 0);
        pctIntercompanyRow.push(grandTotalCol ? totals.intercompany[i] / grandTotalCol : 0);
        pctHDCRow.push(grandTotalCol ? totals.hdc[i] / grandTotalCol : 0);
        pctOutletRow.push(grandTotalCol ? totals.outlet[i] / grandTotalCol : 0);
    }

    // padding the rest
    pctStoresRow.push(...['','','','','','']);
    pctIntercompanyRow.push(...['','','','','','']);
    pctHDCRow.push(...['','','','','','']);
    pctOutletRow.push(...['','','','','','']);

    outputData.push(pctStoresRow);
    outputData.push(pctIntercompanyRow);
    outputData.push(pctHDCRow);
    outputData.push(pctOutletRow);

    sheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // Apply Formatting
    // Set format for percentage rows
    sheet.getRange(outputData.length - 3, 3, 4, sortedDays.length).setNumberFormat("0.00%");

    // Format headers
    sheet.getRange(1, 1, 2, outputData[0].length).setFontWeight("bold");

    // Background colors for the dates logic
    // Green means in SLA <= 7 days, Red is outside > 7 days. Overridden if picking or shipping today.
    const bgColors = [];
    for (let r = 0; r < storeKeys.length; r++) {
      const store = storeKeys[r];
      const data = reportData[origin][store];
      const isPicking = pickingToday[origin].has(store);
      const isShipping = shippingToday[origin].has(store);

      let overrideColor = null;
      if (isPicking) overrideColor = "#4a86e8"; // Light Blue 2 approximation in hex or you can use Google standard like #cfe2f3
      if (isShipping) overrideColor = "#d9d2e9"; // Light Purple 2 approximation
      // If both picking and shipping, we just take shipping or what you prefer. Let's assume they don't overlap or one takes precedence.

      const rowColors = [];

      for (let c = 0; c < sortedDays.length; c++) {
         const daysDiff = sortedDays[c];
         if (overrideColor) {
             rowColors.push(overrideColor);
         } else {
             if (daysDiff <= 7) {
                 rowColors.push("#00FF00"); // Green
             } else {
                 rowColors.push("#FF0000"); // Red
             }
         }
      }
      bgColors.push(rowColors);
    }

    if (bgColors.length > 0 && bgColors[0].length > 0) {
      sheet.getRange(3, 3, bgColors.length, bgColors[0].length).setBackgrounds(bgColors);
    }

    // Set column widths to be more legible
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 60);
  }
}
