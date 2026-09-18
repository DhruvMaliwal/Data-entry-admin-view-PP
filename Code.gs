// =============================================================================
// Property Data Entry Portal — Backend
// =============================================================================
// Flow:
//   1. User lands on the page, sees a list of households (by name only)
//   2. User creates a new household by name, or selects an existing one
//   3. On the household page, they see previous entries (with date/time)
//     and can upload a new data entry sheet
//   4. Uploading generates a dashboard for that entry; the data is stored
//     permanently and can be viewed again later
//
// Storage in the master spreadsheet:
//   - "Households" sheet: House ID (auto), Name, Created At
//   - "Entries"    sheet: Entry ID (auto), House ID, Uploaded At, File Name, Row Count
//   - "Entry Data" sheet: Entry ID, [dynamic columns from uploaded files]
// =============================================================================

// ---------------------------------------------------------------------------
// Configuration — replace with your actual Google Sheet ID
// ---------------------------------------------------------------------------
const MASTER_SHEET_ID = 'YOUR_SHEET_ID_HERE';

const HOUSEHOLDS_SHEET = 'Households';
const ENTRIES_SHEET = 'Entries';
const ENTRY_DATA_SHEET = 'Entry Data';

// ---------------------------------------------------------------------------
// doGet — serves the frontend
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Property Data Entry Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------------------
// getHouseholds — returns the list of existing households
// ---------------------------------------------------------------------------
function getHouseholds() {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = getOrCreateSheet(ss, HOUSEHOLDS_SHEET, ['House ID', 'Name', 'Created At']);

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var houses = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      houses.push({
        id: String(data[i][0]).trim(),
        name: String(data[i][1]).trim(),
        createdAt: data[i][2] ? formatDateTime(data[i][2]) : ''
      });
    }
  }
  return houses;
}

// ---------------------------------------------------------------------------
// addHousehold — creates a household with an auto-generated ID
// ---------------------------------------------------------------------------
function addHousehold(name) {
  if (!name || !name.trim()) {
    return { success: false, message: 'House name is required.' };
  }
  name = name.trim();

  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = getOrCreateSheet(ss, HOUSEHOLDS_SHEET, ['House ID', 'Name', 'Created At']);

  // Check for duplicate name (case-insensitive)
  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][1]).trim().toLowerCase() === name.toLowerCase()) {
      return { success: false, message: 'A house named "' + name + '" already exists.' };
    }
  }

  var id = 'H' + Date.now().toString(36).toUpperCase();
  sheet.appendRow([id, name, new Date()]);
  return { success: true, message: 'House "' + name + '" created.', house: { id: id, name: name } };
}

// ---------------------------------------------------------------------------
// removeHousehold — removes a household + all its entries and data rows
// ---------------------------------------------------------------------------
function removeHousehold(houseId) {
  try {
    if (!houseId) return { success: false, message: 'No House ID provided.' };
    houseId = String(houseId).trim();

    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);

    // Remove from Households
    var hs = ss.getSheetByName(HOUSEHOLDS_SHEET);
    if (hs && hs.getLastRow() > 1) {
      var data = hs.getRange(2, 1, hs.getLastRow() - 1, 1).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]).trim() === houseId) hs.deleteRow(i + 2);
      }
    }

    // Collect entry IDs belonging to this house, delete rows in Entries
    var entryIds = [];
    var es = ss.getSheetByName(ENTRIES_SHEET);
    if (es && es.getLastRow() > 1) {
      var edata = es.getRange(2, 1, es.getLastRow() - 1, 2).getValues();
      for (var j = edata.length - 1; j >= 0; j--) {
        if (String(edata[j][1]).trim() === houseId) {
          entryIds.push(String(edata[j][0]).trim());
          es.deleteRow(j + 2);
        }
      }
    }

    // Delete entry data rows for those entry IDs
    if (entryIds.length > 0) {
      var ds = ss.getSheetByName(ENTRY_DATA_SHEET);
      if (ds && ds.getLastRow() > 1) {
        var ddata = ds.getRange(2, 1, ds.getLastRow() - 1, 1).getValues();
        for (var k = ddata.length - 1; k >= 0; k--) {
          if (entryIds.indexOf(String(ddata[k][0]).trim()) !== -1) ds.deleteRow(k + 2);
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'House removed.' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// getEntriesForHouse — returns metadata for all entries of a given house
// ---------------------------------------------------------------------------
function getEntriesForHouse(houseId) {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === houseId) {
      entries.push({
        id: String(data[i][0]).trim(),
        houseId: String(data[i][1]).trim(),
        uploadedAt: data[i][2] ? formatDateTime(data[i][2]) : '',
        uploadedAtRaw: data[i][2] ? new Date(data[i][2]).getTime() : 0,
        fileName: String(data[i][3] || ''),
        rowCount: Number(data[i][4]) || 0
      });
    }
  }
  // Newest first
  entries.sort(function(a, b) { return b.uploadedAtRaw - a.uploadedAtRaw; });
  return entries;
}

// ---------------------------------------------------------------------------
// getEntryData — returns headers + rows for one entry, for dashboard display
// ---------------------------------------------------------------------------
function getEntryData(entryId) {
  try {
    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);

    // Find the entry metadata
    var es = ss.getSheetByName(ENTRIES_SHEET);
    if (!es) return { success: false, message: 'No entries found.' };
    var edata = es.getDataRange().getValues();
    var meta = null;
    for (var i = 1; i < edata.length; i++) {
      if (String(edata[i][0]).trim() === entryId) {
        meta = {
          id: String(edata[i][0]).trim(),
          houseId: String(edata[i][1]).trim(),
          uploadedAt: edata[i][2] ? formatDateTime(edata[i][2]) : '',
          fileName: String(edata[i][3] || ''),
          rowCount: Number(edata[i][4]) || 0
        };
        break;
      }
    }
    if (!meta) return { success: false, message: 'Entry not found.' };

    // Load data rows for this entry
    var ds = ss.getSheetByName(ENTRY_DATA_SHEET);
    if (!ds || ds.getLastRow() <= 1) {
      return { success: true, meta: meta, headers: [], data: [] };
    }
    var ddata = ds.getDataRange().getValues();
    var allHeaders = ddata[0]; // first column is Entry ID
    var rowObjs = [];
    for (var j = 1; j < ddata.length; j++) {
      if (String(ddata[j][0]).trim() === entryId) {
        var obj = {};
        for (var c = 1; c < allHeaders.length; c++) {
          if (allHeaders[c]) obj[allHeaders[c]] = ddata[j][c];
        }
        rowObjs.push(obj);
      }
    }

    // Determine which columns have any data for this entry
    var displayHeaders = [];
    for (var c = 1; c < allHeaders.length; c++) {
      if (!allHeaders[c]) continue;
      var hasData = false;
      for (var r = 0; r < rowObjs.length; r++) {
        var v = rowObjs[r][allHeaders[c]];
        if (v !== '' && v !== null && v !== undefined) { hasData = true; break; }
      }
      if (hasData) displayHeaders.push(allHeaders[c]);
    }

    return { success: true, meta: meta, headers: displayHeaders, data: rowObjs };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// processDataEntry — parses uploaded CSV, stores it under a new entry ID,
// and returns the dashboard-ready data
// ---------------------------------------------------------------------------
function processDataEntry(houseId, dataFile) {
  try {
    if (!houseId) return { success: false, message: 'No household selected.' };
    if (!dataFile) return { success: false, message: 'No file uploaded.' };

    var rows = parseCsv(dataFile);
    if (!rows || rows.length === 0) {
      return { success: false, message: 'No data rows found in the uploaded file.' };
    }

    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var uploadHeaders = Object.keys(rows[0]);

    // 1) Append to Entries index
    var es = getOrCreateSheet(ss, ENTRIES_SHEET,
      ['Entry ID', 'House ID', 'Uploaded At', 'File Name', 'Row Count']);
    var entryId = 'E' + Date.now().toString(36).toUpperCase();
    var uploadedAt = new Date();
    es.appendRow([entryId, houseId, uploadedAt, dataFile.name, rows.length]);

    // 2) Append rows to Entry Data (first column = Entry ID, then dynamic columns)
    var ds = getOrCreateSheet(ss, ENTRY_DATA_SHEET, ['Entry ID']);
    var existingHeaders = ds.getRange(1, 1, 1, Math.max(1, ds.getLastColumn())).getValues()[0];

    // Add any new columns that don't exist yet
    for (var h = 0; h < uploadHeaders.length; h++) {
      if (existingHeaders.indexOf(uploadHeaders[h]) === -1) {
        existingHeaders.push(uploadHeaders[h]);
        ds.getRange(1, existingHeaders.length).setValue(uploadHeaders[h]).setFontWeight('bold');
      }
    }

    // Build rows aligned to existingHeaders
    var rowsToAppend = [];
    for (var i = 0; i < rows.length; i++) {
      var row = [];
      for (var c = 0; c < existingHeaders.length; c++) {
        var header = existingHeaders[c];
        if (c === 0) row.push(entryId);
        else row.push(rows[i][header] !== undefined ? rows[i][header] : '');
      }
      rowsToAppend.push(row);
    }
    if (rowsToAppend.length > 0) {
      ds.getRange(ds.getLastRow() + 1, 1, rowsToAppend.length, existingHeaders.length)
        .setValues(rowsToAppend);
    }

    return {
      success: true,
      message: rows.length + ' row(s) uploaded.',
      meta: {
        id: entryId,
        houseId: houseId,
        uploadedAt: formatDateTime(uploadedAt),
        fileName: dataFile.name,
        rowCount: rows.length
      },
      headers: uploadHeaders,
      data: rows
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// parseCsv — parses CSV blob (Excel is converted to CSV client-side)
// ---------------------------------------------------------------------------
function parseCsv(dataFile) {
  var bytes = Utilities.base64Decode(dataFile.base64);
  var blob = Utilities.newBlob(bytes, 'text/csv', dataFile.name);
  var text = blob.getDataAsString();
  var parsed = Utilities.parseCsv(text);
  if (parsed.length < 2) return [];

  var headers = parsed[0].map(function(h) { return String(h).trim(); });
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
// getOrCreateSheet — returns existing sheet or creates one with headers
// ---------------------------------------------------------------------------
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// formatDateTime — formats a Date in the script's timezone
// ---------------------------------------------------------------------------
function formatDateTime(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
}
