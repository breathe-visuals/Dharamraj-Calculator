// ============================================================
// Code.gs  —  Jewellery Calculation App
// ============================================================

var HEADER_TEXTS = ['calc id', 'id', 'calculation id', 'calcid'];

/**
 * Returns true if val looks like a real record ID (has at least one digit,
 * is not empty, is not a header label).
 */
function isValidId_(val) {
  if (val === null || val === undefined) return false;
  var s = String(val).trim();
  return s !== '' && HEADER_TEXTS.indexOf(s.toLowerCase()) === -1 && /\d/.test(s);
}

// Robust fallback to find the Calculations sheet even if named range is broken
function getCalculationsSheet_(ss) {
  var rng = ss.getRangeByName('RANGECALC');
  if (rng) {
    try { return rng.getSheet(); } catch(e) {}
  }
  var s = ss.getSheetByName('Calculations');
  if (s) return s;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var val = String(sheets[i].getRange(1, 1).getValue()).trim().toLowerCase();
    if (HEADER_TEXTS.indexOf(val) !== -1) return sheets[i];
  }
  return null;
}

// ── Serve the web app (kept for backwards compatibility) ──────
function doGet(e) {
  // If called with an action param, route to doPost logic
  if (e && e.parameter && e.parameter.action) {
    var result = routeAction_(e.parameter.action, e.parameter.data ? JSON.parse(e.parameter.data) : null);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Otherwise serve a simple status page
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Jewellery Calculator API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── REST API Entry Point for Vercel Frontend ──────────────────
function doPost(e) {
  var result;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;
    var data   = body.data || null;
    result = routeAction_(action, data);
  } catch (err) {
    result = { ok: false, error: 'doPost error: ' + err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Action Router ─────────────────────────────────────────────
function routeAction_(action, data) {
  if (action === 'getConfig')          return getConfig();
  if (action === 'getCalculations')    return getCalculations();
  if (action === 'saveCalculation')    return saveCalculation(data);
  if (action === 'updateCalculation')  return updateCalculation(data);
  if (action === 'deleteCalculation')  return deleteCalculation(data && data.id);
  if (action === 'deleteCalculations') return deleteCalculations(data && data.ids);
  return { ok: false, error: 'Unknown action: ' + action };
}

// ── getConfig ────────────────────────────────────────────────
function getConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rng = ss.getRangeByName('RANGECONFIG');
    var sheet = rng ? rng.getSheet() : ss.getSheetByName('Config');
    if (!sheet) return { ok: false, error: 'Config sheet not found.' };

    var cfg = { prefix: 'CALC', businessName: 'My Business', lastId: 0, defaultGST: 3 };
    sheet.getDataRange().getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (k === 'Prefix')        cfg.prefix       = String(r[1]).trim();
      if (k === 'Business Name') cfg.businessName = String(r[1]).trim();
      if (k === 'LastID')        cfg.lastId       = Number(r[1]) || 0;
      if (k === 'DefaultGST')    cfg.defaultGST   = Number(r[1]) || 3;
    });
    return { ok: true, data: cfg };
  } catch (e) {
    return { ok: false, error: 'getConfig: ' + e.message };
  }
}

// ── getCalculations ──────────────────────────────────────────
function getCalculations() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getCalculationsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Calculations sheet not found.' };
    var lastRow = getLastDataRow_(sheet);
    if (lastRow < 2) return { ok: true, data: [] }; // No data yet

    var dataRange = sheet.getRange(2, 1, lastRow - 1, 20);
    var tz   = Session.getScriptTimeZone();
    var rows = [];

    dataRange.getValues().forEach(function (r) {
      if (!isValidId_(r[0])) return;

      var dateVal = '';
      if (r[1]) {
        try {
          var d = new Date(r[1]);
          if (!isNaN(d.getTime())) {
            var day = ('0' + d.getDate()).slice(-2);
            var month = ('0' + (d.getMonth() + 1)).slice(-2);
            var year = d.getFullYear();
            dateVal = day + '/' + month + '/' + year;
          }
        } catch (_) {}
      }

      rows.push({
        calcId          : String(r[0] || '').trim(),
        date            : dateVal,
        customerName    : String(r[2] || ''),
        weight          : safeNum_(r[3]),
        ratePerG        : safeNum_(r[4]),
        labourPerG      : safeNum_(r[5]),
        otherCharge     : safeNum_(r[6]),
        gstPercent      : safeNum_(r[7]),
        metalAmount     : safeNum_(r[8]),
        labourAmount    : safeNum_(r[9]),
        otherAmount     : safeNum_(r[10]),
        gstAmount       : safeNum_(r[11]),
        totalAmount     : safeNum_(r[12]),
        metalPercent    : safeNum_(r[13]),
        labourPercent   : safeNum_(r[14]),
        otherPercent    : safeNum_(r[15]),
        metal           : String(r[16] || ''),
        remarks         : String(r[17] || ''),
        receivedAmount  : safeNum_(r[18]),
        balanceAmount   : safeNum_(r[19])
      });
    });

    return { ok: true, data: rows };
  } catch (e) {
    return { ok: false, error: 'getCalculations: ' + e.message };
  }
}

// ── saveCalculation ──────────────────────────────────────────
function saveCalculation(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getCalculationsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Calculations sheet not found.' };
    var id    = generateId_(ss);
    
    var lastRow = getLastDataRow_(sheet);
    var newRow = lastRow + 1;

    var rowData = [
      id,
      new Date(data.date),
      data.customerName || '',
      safeNum_(data.weight),
      safeNum_(data.ratePerG),
      safeNum_(data.labourPerG),
      safeNum_(data.otherCharge),
      safeNum_(data.gstPercent),
      safeNum_(data.metalAmount),
      safeNum_(data.labourAmount),
      safeNum_(data.otherAmount),
      safeNum_(data.gstAmount),
      safeNum_(data.totalAmount),
      safeNum_(data.metalPercent),
      safeNum_(data.labourPercent),
      safeNum_(data.otherPercent),
      data.metal || '',
      data.remarks || '',
      safeNum_(data.receivedAmount),
      safeNum_(data.balanceAmount)
    ];

    sheet.getRange(newRow, 1, 1, 20).setValues([rowData]);
    
    // Attempt to force a flush so subsequent getCalculations fetch the updated data.
    SpreadsheetApp.flush();
    return { ok: true, calcId: id };
  } catch (e) {
    return { ok: false, error: 'saveCalculation: ' + e.message };
  }
}

// ── updateCalculation ────────────────────────────────────────
function updateCalculation(data) {
  try {
    if (!isValidId_(data.calcId)) return { ok: false, error: 'Invalid ID: ' + data.calcId };
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getCalculationsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Calculations sheet not found.' };
    var row   = findRowById_(sheet, data.calcId);
    if (!row) return { ok: false, error: 'Record not found: ' + data.calcId };

    sheet.getRange(row, 1, 1, 20).setValues([[
      data.calcId,
      new Date(data.date),
      data.customerName || '',
      safeNum_(data.weight),
      safeNum_(data.ratePerG),
      safeNum_(data.labourPerG),
      safeNum_(data.otherCharge),
      safeNum_(data.gstPercent),
      safeNum_(data.metalAmount),
      safeNum_(data.labourAmount),
      safeNum_(data.otherAmount),
      safeNum_(data.gstAmount),
      safeNum_(data.totalAmount),
      safeNum_(data.metalPercent),
      safeNum_(data.labourPercent),
      safeNum_(data.otherPercent),
      data.metal || '',
      data.remarks || '',
      safeNum_(data.receivedAmount),
      safeNum_(data.balanceAmount)
    ]]);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch (e) {
    return { ok: false, error: 'updateCalculation: ' + e.message };
  }
}
    
// ── deleteCalculation ────────────────────────────────────────
function deleteCalculation(id) {
  try {
    if (!isValidId_(id)) return { ok: false, error: 'Not a valid record ID: ' + id };
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getCalculationsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Calculations sheet not found.' };
    var row   = findRowById_(sheet, id);
    if (!row) return { ok: false, error: 'Record not found: ' + id };
    sheet.deleteRow(row);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: 'deleteCalculation: ' + e.message };
  }
}

// ── deleteCalculations (Bulk) ────────────────────────────────
function deleteCalculations(ids) {
  try {
    if (!Array.isArray(ids) || ids.length === 0) return { ok: false, error: 'No IDs provided' };
    
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getCalculationsSheet_(ss);
    if (!sheet) return { ok: false, error: 'Calculations sheet not found.' };
    
    var rowsToDelete = [];
    ids.forEach(function(id) {
      if (isValidId_(id)) {
        var row = findRowById_(sheet, id);
        if (row) rowsToDelete.push(row);
      }
    });
    
    // Sort rows descending to safely delete from bottom to top without shifting indices
    rowsToDelete.sort(function(a, b) { return b - a; });
    
    rowsToDelete.forEach(function(row) {
      sheet.deleteRow(row);
    });
    
    SpreadsheetApp.flush();
    return { ok: true, deleted: rowsToDelete.length };
  } catch (e) {
    return { ok: false, error: 'deleteCalculations: ' + e.message };
  }
}

// ── Helpers ───────────────────────────────────────────────────

function getLastDataRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;
  var vals = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== "") {
      return i + 1;
    }
  }
  return 1;
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var hit = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(id).trim())
    .matchEntireCell(true)
    .findNext();
  return hit ? hit.getRow() : null;
}

function generateId_(ss) {
  var rng    = ss.getRangeByName('RANGECONFIG');
  var sheet  = rng ? rng.getSheet() : ss.getSheetByName('Config');
  if (!sheet) return 'CALC-' + String(Date.now()).slice(-4);
  
  var vals   = sheet.getDataRange().getValues();
  var prefix = 'CALC', lastId = 0, idRow = -1;

  vals.forEach(function (r, i) {
    var k = String(r[0]).trim();
    if (k === 'Prefix') prefix = String(r[1]).trim();
    if (k === 'LastID') { lastId = Number(r[1]) || 0; idRow = i; }
  });

  var newId = lastId + 1;
  // Since getDataRange() starts at row 1, index i corresponds to sheet row i+1
  if (idRow >= 0) sheet.getRange(idRow + 1, 2).setValue(newId);
  return prefix + '-' + String(newId).padStart(4, '0');
}

function safeNum_(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  var s = String(val).replace(/[^0-9.-]/g, '');
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}