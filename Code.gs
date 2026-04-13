// ═══════════════════════════════════════════════════════════════════
//  DC / XD Ops Intelligence Suite — Google Apps Script Backend
//  Deploy as: Web App → Execute as: Me → Access: Anyone
// ═══════════════════════════════════════════════════════════════════

// ── SHEET NAME MAPPING ──────────────────────────────────────────────
var SHEET_MAP = {
  'seal':     'Seal Mismatch base',
  'misroute': 'Misroute Tracker Base'
};

// ── MAIN ENTRY POINT ────────────────────────────────────────────────
function doGet(e) {
  var output;
  try {
    output = handleRequest(e);
  } catch (err) {
    output = { error: String(err), stack: err.stack || '' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CORS-friendly POST (optional, same logic) ────────────────────────
function doPost(e) {
  return doGet(e);
}

// ── REQUEST HANDLER ──────────────────────────────────────────────────
function handleRequest(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  // ?sheet=seal  OR  ?sheet=misroute
  var sheetKey = (params.sheet || '').toLowerCase().trim();

  if (!sheetKey) {
    // No sheet specified → return a manifest of available sheets
    return {
      status: 'ok',
      available_keys: Object.keys(SHEET_MAP),
      usage: 'Add ?sheet=seal or ?sheet=misroute to your URL'
    };
  }

  var sheetName = SHEET_MAP[sheetKey];
  if (!sheetName) {
    return {
      error: 'Unknown sheet key "' + sheetKey + '". Valid keys: ' + Object.keys(SHEET_MAP).join(', ')
    };
  }

  return readSheet(sheetName);
}

// ── SHEET READER ─────────────────────────────────────────────────────
function readSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return {
      error: 'Sheet "' + sheetName + '" not found. Available sheets: ' +
             ss.getSheets().map(function(s) { return s.getName(); }).join(', ')
    };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return { headers: [], rows: [], count: 0, sheet: sheetName };
  }

  // Get all data in one API call (efficient)
  var allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  if (allValues.length === 0) {
    return { headers: [], rows: [], count: 0, sheet: sheetName };
  }

  // First row = headers; clean them up
  var rawHeaders = allValues[0];
  var headers = rawHeaders.map(function(h) {
    return String(h == null ? '' : h).trim();
  });

  // Filter out completely empty trailing columns
  var lastNonEmptyCol = headers.length - 1;
  while (lastNonEmptyCol > 0 && headers[lastNonEmptyCol] === '') {
    lastNonEmptyCol--;
  }
  headers = headers.slice(0, lastNonEmptyCol + 1);

  if (headers.length === 0 || headers.every(function(h) { return h === ''; })) {
    return {
      error: 'Sheet "' + sheetName + '" has no header row. Ensure Row 1 contains column names.',
      sheet: sheetName
    };
  }

  // Build row objects
  var rows = [];
  for (var i = 1; i < allValues.length; i++) {
    var row = allValues[i];

    // Skip completely empty rows
    var hasData = row.some(function(cell) {
      return cell !== null && cell !== undefined && cell !== '';
    });
    if (!hasData) continue;

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      if (!key) continue;  // skip blank-header columns
      var val = row[j];

      // Normalise cell value
      if (val instanceof Date) {
        obj[key] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      } else if (val === null || val === undefined) {
        obj[key] = '';
      } else {
        obj[key] = val;
      }
    }
    rows.push(obj);
  }

  return {
    status: 'ok',
    sheet: sheetName,
    headers: headers,
    rows: rows,
    count: rows.length,
    fetched_at: new Date().toISOString()
  };
}
