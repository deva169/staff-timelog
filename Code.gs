var SPREADSHEET_ID = '1NrHiIHC5mZX-Fvo_BLfiWgrmr-gjLrYIPLgpAKhrbk';
var SHEET_NAME = 'Sheet1';
var STAFF_SHEET_NAME = 'Staff Database';
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
var STAFF_HEADERS = [
  'Username',
  'PIN',
  'First name',
  'Last name',
  'DOB',
  'Mobile',
  'Email',
  'Full name',
  'Initials',
  'Status',
  'Created At'
];

function doGet(e) {
  try {
    var p = e && e.parameter ? e.parameter : {};
    var action = String(p.action || '').toLowerCase();

    if (action === 'get') return jsonOut(getRows(p.userId));
    if (action === 'append') return jsonOut(appendRow(p));
    if (action === 'update') return jsonOut(updateRow(p));
    if (action === 'login') return jsonOut(loginStaff(p));
    if (action === 'register') return jsonOut(registerStaff(p));

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

function getStaffSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && SPREADSHEET_ID) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(STAFF_SHEET_NAME) || ss.insertSheet(STAFF_SHEET_NAME);
  ensureStaffHeaders(sh);
  ensureDefaultStaff(sh);
  return sh;
}

function ensureHeaders(sh) {
  var current = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var needsHeaders = current.join('') === '';
  if (!needsHeaders) return;
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
}

function ensureStaffHeaders(sh) {
  var current = sh.getRange(1, 1, 1, STAFF_HEADERS.length).getValues()[0];
  var needsHeaders = current.join('') === '' || String(current[5]).trim() !== 'Mobile';
  if (!needsHeaders) return;
  sh.getRange(1, 1, 1, STAFF_HEADERS.length).setValues([STAFF_HEADERS]);
  sh.setFrozenRows(1);
}

function ensureDefaultStaff(sh) {
  if (findStaffRow(sh, 'deva') > 0) return;
  sh.appendRow(['deva', '1234', 'Deva', '', '', '', '', 'Deva', 'DV', 'Active', new Date().toISOString()]);
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
    'Pending'
  ]]);
  return { ok: true, id: id };
}

function loginStaff(p) {
  var uid = normalizeUsername(p.uid || p.userId || '');
  var pin = String(p.pin || '');
  if (!uid || !pin) return { ok: false, error: 'Enter username and PIN.' };

  var sh = getStaffSheet();
  var row = findStaffRow(sh, uid);
  if (row < 1) return { ok: false, error: 'Invalid username or PIN.' };

  var data = sh.getRange(row, 1, 1, STAFF_HEADERS.length).getDisplayValues()[0];
  if (String(data[1]) !== pin) return { ok: false, error: 'Invalid username or PIN.' };
  if (String(data[9]).toLowerCase() !== 'active') return { ok: false, error: 'Staff login is not active.' };

  return {
    ok: true,
    user: {
      id: data[0],
      name: data[7] || data[2],
      initials: data[8] || initialsFor(data[2], data[3])
    }
  };
}

function registerStaff(p) {
  var first = cleanName(p.first || '');
  var last = cleanName(p.last || '');
  var dob = String(p.dob || '').trim();
  var mobile = String(p.mobile || '').trim();
  var email = String(p.email || '').trim().toLowerCase();
  var pin = String(p.pin || '').trim();

  if (!first || !last || !dob || !mobile || !email || !pin) return { ok: false, error: 'All registration fields are required.' };
  if (!/^\d{4}$/.test(pin)) return { ok: false, error: 'PIN must be exactly 4 digits.' };
  if (!/^[0-9 +()-]{8,20}$/.test(mobile)) return { ok: false, error: 'Enter a valid mobile number.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Enter a valid email address.' };

  var sh = getStaffSheet();
  var username = nextUsername(sh, first, last);
  var fullName = first + ' ' + last;
  var initials = initialsFor(first, last);

  sh.appendRow([
    username,
    pin,
    first,
    last,
    dob,
    mobile,
    email,
    fullName,
    initials,
    'Active',
    new Date().toISOString()
  ]);

  return { ok: true, user: { id: username, name: fullName, initials: initials }, username: username };
}

function nextUsername(sh, first, last) {
  var base = normalizeUsername(first) + '_' + normalizeUsername(last).charAt(0);
  var taken = {};
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var users = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    users.forEach(function(row) { taken[String(row[0]).toLowerCase()] = true; });
  }
  if (!taken[base]) return base;
  for (var i = 1; i < 100; i++) {
    var candidate = base + String(i).padStart(2, '0');
    if (!taken[candidate]) return candidate;
  }
  return base + String(new Date().getTime()).slice(-4);
}

function findStaffRow(sh, username) {
  username = normalizeUsername(username);
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var rows = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    if (normalizeUsername(rows[i][0]) === username) return i + 2;
  }
  return -1;
}

function normalizeUsername(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function cleanName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function initialsFor(first, last) {
  return ((String(first || '').charAt(0) || '') + (String(last || '').charAt(0) || '')).toUpperCase() || 'ST';
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
