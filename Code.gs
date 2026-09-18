// =============================================================================
// Property Data Entry Portal — Backend
// =============================================================================
// This Apps Script web app receives property data files (Excel/CSV) and images
// from a browser form, parses the data, stores images in Google Drive, and
// appends rows to a master Google Sheet. Communication between frontend and
// backend uses google.script.run (HTML Service pattern), not HTTP POST.
// =============================================================================

// ---------------------------------------------------------------------------
// Configuration — replace these with your actual Google Sheet and Drive IDs
// ---------------------------------------------------------------------------
const MASTER_SHEET_ID = 'YOUR_SHEET_ID_HERE';
const IMAGE_FOLDER_ID = 'YOUR_FOLDER_ID_HERE';

// Column headers expected in the master sheet (first 7 are fixed; Images is last)
const EXPECTED_HEADERS = [
  'House ID', 'Address', 'Price', 'Beds', 'Baths', 'Description', 'Images'
];

// ---------------------------------------------------------------------------
// doGet — serves the frontend HTML page when the web app URL is opened
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Property Data Entry Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------------------
// processSubmission — main entry point called from the frontend
// ---------------------------------------------------------------------------
// Receives:
//   dataFile  — { name: string, mimeType: string, base64: string }
//   imageFiles — array of { name: string, mimeType: string, base64: string }
//
// Flow:
//   1. Parse the data file into rows keyed by House ID
//   2. Upload each image to Drive, match it to a House ID via filename prefix
//   3. Append one row per house to the master sheet with comma-separated image URLs
//   4. Return a summary object to the frontend
// ---------------------------------------------------------------------------
function processSubmission(dataFile, imageFiles) {
  try {
    // 1. Parse the uploaded data file into an array of row objects
    var rows = parseDataFile(dataFile);
    if (!rows || rows.length === 0) {
      return { success: false, message: 'No data rows found in the uploaded file.' };
    }

    // Build a lookup: House ID → row object (preserves upload order)
    var houseMap = {};
    var houseIds = [];
    for (var i = 0; i < rows.length; i++) {
      var id = String(rows[i]['House ID'] || '').trim();
      if (!id) continue;
      houseMap[id] = rows[i];
      houseMap[id]._imageUrls = [];
      houseIds.push(id);
    }

    // 2. Upload images to Drive and match to houses by filename prefix
    var imageCount = 0;
    var warnings = [];
    var folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);

    if (imageFiles && imageFiles.length > 0) {
      for (var j = 0; j < imageFiles.length; j++) {
        var img = imageFiles[j];
        var houseId = extractHouseId(img.name);

        if (!houseId || !houseMap[houseId]) {
          warnings.push('Skipped image "' + img.name + '" — no matching House ID.');
          continue;
        }

        // Decode base64 → blob, save to Drive, make shareable
        var blob = Utilities.newBlob(
          Utilities.base64Decode(img.base64),
          img.mimeType,
          img.name
        );
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';

        houseMap[houseId]._imageUrls.push(url);
        imageCount++;
      }
    }

    // 3. Append rows to the master Google Sheet
    var sheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheets()[0];
    ensureHeaders(sheet);

    var rowsToAppend = [];
    for (var k = 0; k < houseIds.length; k++) {
      var house = houseMap[houseIds[k]];
      rowsToAppend.push([
        house['House ID']    || '',
        house['Address']     || '',
        house['Price']       || '',
        house['Beds']        || '',
        house['Baths']       || '',
        house['Description'] || '',
        (house._imageUrls || []).join(', ')
      ]);
    }

    if (rowsToAppend.length > 0) {
      sheet.getRange(
        sheet.getLastRow() + 1, 1,
        rowsToAppend.length, EXPECTED_HEADERS.length
      ).setValues(rowsToAppend);
    }

    // 4. Return summary
    return {
      success: true,
      message: rowsToAppend.length + ' house(s) added, ' + imageCount + ' image(s) uploaded.',
      warnings: warnings
    };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// parseDataFile — detects CSV vs Excel and returns an array of row objects
// ---------------------------------------------------------------------------
// Each object has keys matching the header row of the uploaded file.
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
// parseCsv — parses a CSV blob into row objects using Utilities.parseCsv
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
// parseExcel — converts an Excel file to CSV via a temporary Google Sheet,
// then parses the resulting CSV. This avoids needing an external library.
// ---------------------------------------------------------------------------
// Flow: upload blob as a Google Sheet (Drive converts it), read sheet data,
// delete the temp file.
// ---------------------------------------------------------------------------
function parseExcel(blob) {
  // Drive's insert with convert:true turns xlsx → Google Sheet
  var resource = { title: 'TempUpload_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS };
  var tempFile = Drive.Files.insert(resource, blob, { convert: true });
  var tempSheet = SpreadsheetApp.openById(tempFile.id);
  var data = tempSheet.getSheets()[0].getDataRange().getValues();

  // Clean up the temporary file immediately
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
// extractHouseId — gets the House ID from an image filename
// ---------------------------------------------------------------------------
// Convention: everything before the first underscore is the House ID.
// Example: "H101_front.jpg" → "H101"
// If there's no underscore, the filename without extension is used.
// ---------------------------------------------------------------------------
function extractHouseId(filename) {
  if (!filename) return null;
  var nameOnly = filename.replace(/\.[^.]+$/, ''); // strip extension
  var parts = nameOnly.split('_');
  return parts[0].trim() || null;
}

// ---------------------------------------------------------------------------
// ensureHeaders — writes column headers to row 1 if the sheet is empty
// ---------------------------------------------------------------------------
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setValues([EXPECTED_HEADERS]);
    sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setFontWeight('bold');
  }
}
