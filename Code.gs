var SPREADSHEET_ID = '1NrHiIHC5mZX-Fvo_BLfiWgrmr-gjLrYIPLgpAKhrbk';
var SHEET_NAME = 'Sheet1';
var HEADERS = [
  'ID',
  'User ID',
  'Name',
  'Location',
  'Note',
  'Date',
  'Start Time',
  'Finish time',
  'Duration (hrs)',
  'Status',
  'Submitted At'
];

function doGet(e) {
  try {
    var p = e && e.parameter ? e.parameter : {};
    var action = String(p.action || '').toLowerCase();

    if (action === 'get') return jsonOut(getRows(p.userId));
    if (action === 'append') return jsonOut(appendRow(p));
    if (action === 'update') return jsonOut(updateRow(p));

    return jsonOut({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && SPREADSHEET_ID) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  ensureHeaders(sh);
  return sh;
}

function ensureHeaders(sh) {
  var current = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var needsHeaders = current.join('') === '';
  if (!needsHeaders) return;
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
}

function getRows(userId) {
  var sh = getSheet();
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };

  var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  if (userId) {
    rows = rows.filter(function(row) {
      return String(row[1]).toLowerCase() === String(userId).toLowerCase();
    });
  }
  return { ok: true, rows: rows };
}

function appendRow(p) {
  var sh = getSheet();
  var id = String(p.id || new Date().getTime());
  var existingRow = findRowById(sh, id);
  if (existingRow > 0) return { ok: true, duplicate: true, id: id };

  sh.appendRow([
    id,
    p.uid || '',
    p.name || '',
    p.loc || '',
    p.note || '',
    p.date || '',
    p.start || '',
    '',
    '',
    'In Progress',
    p.ts || new Date().toISOString()
  ]);
  return { ok: true, id: id };
}

function updateRow(p) {
  var sh = getSheet();
  var id = String(p.id || '');
  var row = findRowById(sh, id);
  if (row < 1) return { ok: false, error: 'Entry not found: ' + id };

  sh.getRange(row, 8, 1, 3).setValues([[
    p.finish || '',
    p.dur || '',
    'Complete'
  ]]);
  return { ok: true, id: id };
}

function findRowById(sh, id) {
  if (!id) return -1;
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  return -1;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
