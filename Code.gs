// =============================================================================
// Property Data Entry Portal — Backend
// =============================================================================
// Flow:
//   1. User lands on the page, sees a list of households from the master sheet
//   2. User selects a household (or creates a new one)
//   3. User uploads an Excel/CSV data entry file for that household
//   4. Backend parses the file, appends rows tagged with the household + today's date
//   5. Frontend displays a dashboard summary of the uploaded data
// =============================================================================

// ---------------------------------------------------------------------------
// Configuration — replace these with your actual Google Sheet ID
// ---------------------------------------------------------------------------
const MASTER_SHEET_ID = 'YOUR_SHEET_ID_HERE';

// Sheet names within the master spreadsheet
const HOUSEHOLDS_SHEET = 'Households';
const DATA_ENTRIES_SHEET = 'Data Entries';

// ---------------------------------------------------------------------------
// doGet — serves the frontend HTML page
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Property Data Entry Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------------------
// getHouseholds — returns the list of existing households for the dropdown
// ---------------------------------------------------------------------------
// Each household has an ID and display name. Reads from the "Households" sheet.
// If the sheet doesn't exist yet, it creates one with headers.
// ---------------------------------------------------------------------------
function getHouseholds() {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = getOrCreateSheet(ss, HOUSEHOLDS_SHEET, ['House ID', 'House Name', 'Address', 'Created Date']);

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return []; // only headers, no data

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var households = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      households.push({
        id: String(data[i][0]).trim(),
        name: String(data[i][1]).trim(),
        address: String(data[i][2]).trim()
      });
    }
  }
  return households;
}

// ---------------------------------------------------------------------------
// addHousehold — creates a new household entry
// ---------------------------------------------------------------------------
function addHousehold(houseId, houseName, address) {
  if (!houseId || !houseName) {
    return { success: false, message: 'House ID and Name are required.' };
  }

  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = getOrCreateSheet(ss, HOUSEHOLDS_SHEET, ['House ID', 'House Name', 'Address', 'Created Date']);

  // Check for duplicate ID
  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim() === houseId.trim()) {
      return { success: false, message: 'House ID "' + houseId + '" already exists.' };
    }
  }

  sheet.appendRow([houseId.trim(), houseName.trim(), (address || '').trim(), new Date()]);
  return { success: true, message: 'Household "' + houseName + '" added.' };
}

// ---------------------------------------------------------------------------
// removeHousehold — removes a household from the Households sheet by ID
// ---------------------------------------------------------------------------
function removeHousehold(houseId) {
  if (!houseId) return { success: false, message: 'No House ID provided.' };

  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = ss.getSheetByName(HOUSEHOLDS_SHEET);
  if (!sheet) return { success: false, message: 'Households sheet not found.' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === houseId.trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Household removed.' };
    }
  }
  return { success: false, message: 'House ID not found.' };
}

// ---------------------------------------------------------------------------
// processDataEntry — parses the uploaded file and appends to Data Entries
// ---------------------------------------------------------------------------
// Receives:
//   houseId   — the selected household ID
//   dataFile  — { name, mimeType, base64 }
//
// Every row from the uploaded file gets tagged with:
//   - House ID (from selection)
//   - Upload Date (today)
//   - Original filename
//
// Returns: the parsed rows for the frontend dashboard display
// ---------------------------------------------------------------------------
function processDataEntry(houseId, dataFile) {
  try {
    if (!houseId) return { success: false, message: 'No household selected.' };
    if (!dataFile) return { success: false, message: 'No data file uploaded.' };

    var rows = parseDataFile(dataFile);
    if (!rows || rows.length === 0) {
      return { success: false, message: 'No data rows found in the uploaded file.' };
    }

    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Get the column headers from the uploaded file
    var uploadHeaders = Object.keys(rows[0]);

    // Master headers: House ID, Upload Date, then all columns from the file
    var masterHeaders = ['House ID', 'Upload Date'].concat(uploadHeaders);

    var sheet = getOrCreateSheet(ss, DATA_ENTRIES_SHEET, masterHeaders);

    // Read existing headers to ensure column alignment
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Add any new columns from this upload that don't exist yet
    for (var h = 0; h < uploadHeaders.length; h++) {
      if (existingHeaders.indexOf(uploadHeaders[h]) === -1) {
        existingHeaders.push(uploadHeaders[h]);
        sheet.getRange(1, existingHeaders.length).setValue(uploadHeaders[h]);
      }
    }

    // Build rows aligned to the master header order
    var rowsToAppend = [];
    for (var i = 0; i < rows.length; i++) {
      var row = [];
      for (var c = 0; c < existingHeaders.length; c++) {
        var header = existingHeaders[c];
        if (header === 'House ID') {
          row.push(houseId);
        } else if (header === 'Upload Date') {
          row.push(today);
        } else {
          row.push(rows[i][header] || '');
        }
      }
      rowsToAppend.push(row);
    }

    if (rowsToAppend.length > 0) {
      sheet.getRange(
        sheet.getLastRow() + 1, 1,
        rowsToAppend.length, existingHeaders.length
      ).setValues(rowsToAppend);
    }

    return {
      success: true,
      message: rowsToAppend.length + ' row(s) added for household ' + houseId + '.',
      headers: uploadHeaders,
      data: rows,
      uploadDate: today,
      houseId: houseId
    };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// getEntriesForHouse — fetches all data entries for a given household
// ---------------------------------------------------------------------------
// Used by the dashboard to show historical data.
// ---------------------------------------------------------------------------
function getEntriesForHouse(houseId) {
  try {
    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var sheet = ss.getSheetByName(DATA_ENTRIES_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, headers: [], data: [], dates: [] };
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var houseIdCol = headers.indexOf('House ID');
    var dateCol = headers.indexOf('Upload Date');

    // Filter rows for the selected household
    var filtered = [];
    var dates = {};
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][houseIdCol]).trim() === houseId) {
        var obj = {};
        for (var c = 0; c < headers.length; c++) {
          if (headers[c] !== 'House ID' && headers[c] !== 'Upload Date') {
            obj[headers[c]] = allData[i][c];
          }
        }
        filtered.push(obj);
        if (allData[i][dateCol]) {
          dates[String(allData[i][dateCol])] = true;
        }
      }
    }

    // Column headers excluding House ID and Upload Date
    var displayHeaders = headers.filter(function(h) {
      return h !== 'House ID' && h !== 'Upload Date';
    });

    return {
      success: true,
      headers: displayHeaders,
      data: filtered,
      dates: Object.keys(dates).sort().reverse()
    };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// parseDataFile — detects CSV vs Excel and returns an array of row objects
// ---------------------------------------------------------------------------
function parseDataFile(dataFile) {
  var bytes = Utilities.base64Decode(dataFile.base64);
  var blob = Utilities.newBlob(bytes, dataFile.mimeType, dataFile.name);
  var name = dataFile.name.toLowerCase();

  if (name.endsWith('.csv')) {
    return parseCsv(blob);
  } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    return parseExcel(blob);
  } else {
    throw new Error('Unsupported file type. Please upload .csv or .xlsx');
  }
}

// ---------------------------------------------------------------------------
// parseCsv — parses a CSV blob into row objects
// ---------------------------------------------------------------------------
function parseCsv(blob) {
  var text = blob.getDataAsString();
  var parsed = Utilities.parseCsv(text);
  if (parsed.length < 2) return [];

  var headers = parsed[0].map(function(h) { return h.trim(); });
  var rows = [];

  for (var i = 1; i < parsed.length; i++) {
    var obj = {};
    var hasData = false;
    for (var c = 0; c < headers.length; c++) {
      var val = (parsed[i][c] || '').trim();
      obj[headers[c]] = val;
      if (val) hasData = true;
    }
    if (hasData) rows.push(obj);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// parseExcel — converts Excel to Google Sheet temporarily, reads data, deletes temp
// ---------------------------------------------------------------------------
function parseExcel(blob) {
  var resource = { title: 'TempUpload_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS };
  var tempFile = Drive.Files.insert(resource, blob, { convert: true });
  var tempSheet = SpreadsheetApp.openById(tempFile.id);
  var data = tempSheet.getSheets()[0].getDataRange().getValues();
  DriveApp.getFileById(tempFile.id).setTrashed(true);

  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    var hasData = false;
    for (var c = 0; c < headers.length; c++) {
      var val = String(data[i][c] || '').trim();
      obj[headers[c]] = val;
      if (val) hasData = true;
    }
    if (hasData) rows.push(obj);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// getOrCreateSheet — returns an existing sheet or creates one with headers
// ---------------------------------------------------------------------------
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}
