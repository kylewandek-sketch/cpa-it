var HELPDESK_EMAIL = 'kyle.anderson@cpaohio.org';
var ADMIN_TOKEN = 'CHANGE_ME';   // set your own; do NOT commit the real token to a public repo

// Drive folder that ticket photos are saved into. The account running this script
// must have EDIT access to it. Falls back to a folder on the script's own Drive.
var PHOTO_FOLDER_ID = '17cTg5MESNSdgb_ZWt_e3Lm2t3bqDShQq';
var PHOTO_FOLDER_FALLBACK = 'CPA IT Ticket Photos';

// Native Google Sheet holding the cart rosters (HS_Cart_1..6, HS_Spares, ...).
// Roster tabs are auto-detected: serials in column B, with "Serial #" in B2.
var ROSTER_SHEET_ID = '1FDVE6KtAEf06_zRYQyHyaNZ_9gXsv3JRJGbIwckv4Mw';

// Full column layout. A = Chromebook S/N. Status=9, Notes=10 (unchanged); new cols appended.
var HEADERS = [
  'Chromebook S/N', 'Timestamp', 'Teacher Email', 'Teacher Name', 'Room #',
  'Issue Type', 'Urgency', 'Description', 'Status', 'Notes',
  'Ticket #', 'Student at Fault', 'Assigned To', 'Resolved At', 'Photo'
];
var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// ---- Dashboard GET endpoint (JSONP) ----
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  var guarded = ['list', 'update', 'archiveTest', 'stats', 'lookup'];
  if (guarded.indexOf(p.action) >= 0 && p.token !== ADMIN_TOKEN) {
    out = { ok: false, error: 'unauthorized' };
  } else if (p.action === 'list') {
    out = listTickets_();
  } else if (p.action === 'update') {
    out = updateTicket_(p);
  } else if (p.action === 'stats') {
    out = stats_();
  } else if (p.action === 'lookup') {
    out = deviceLookup_(p);
  } else if (p.action === 'archiveTest') {
    out = archiveCopy_(false);
  } else if (p.action === 'openCount') {
    out = openCount_(p);          // public: duplicate-open-ticket check for the submit form
  } else {
    out = { ok: true, msg: 'CPA IT Tickets endpoint is live.' };
  }
  var json = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function firstSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }

function ensureHeaders_(sheet) {
  var cur = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var changed = false;
  for (var i = 0; i < HEADERS.length; i++) {
    if (!cur[i]) { sheet.getRange(1, i + 1).setValue(HEADERS[i]); changed = true; }
  }
  if (changed) { sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold'); sheet.setFrozenRows(1); }
}

function listTickets_() {
  var sheet = firstSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };
  var lastCol = Math.max(HEADERS.length, sheet.getLastColumn());
  var v = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows = v.map(function (r, i) {
    return {
      row: i + 2,
      sn: r[0], timestamp: r[1] ? new Date(r[1]).toISOString() : '',
      teacherEmail: r[2], teacherName: r[3], room: r[4],
      issue: r[5], urgency: r[6], description: r[7],
      status: r[8] || 'New', notes: r[9] || '',
      ticketNo: r[10] || '', studentAtFault: r[11] || '', assignedTo: r[12] || '',
      resolvedAt: r[13] ? new Date(r[13]).toISOString() : '',
      photoUrl: r[14] || ''
    };
  });
  return { ok: true, rows: rows };
}

function updateTicket_(p) {
  var row = parseInt(p.row, 10);
  if (!row || row < 2) return { ok: false, error: 'bad row' };
  var sheet = firstSheet_();
  ensureHeaders_(sheet);
  var oldStatus = sheet.getRange(row, 9).getValue();
  if (p.status != null) {
    sheet.getRange(row, 9).setValue(p.status);
    if (p.status !== oldStatus) {
      if (p.status === 'Resolved') sheet.getRange(row, 14).setValue(new Date());
      if (p.status === 'In Progress' || p.status === 'Resolved') sendStatusEmail_(sheet, row, p.status);
    }
  }
  if (p.notes != null) sheet.getRange(row, 10).setValue(p.notes);
  if (p.studentAtFault != null) sheet.getRange(row, 12).setValue(p.studentAtFault);
  if (p.assignedTo != null) sheet.getRange(row, 13).setValue(p.assignedTo);
  return { ok: true };
}

function sendStatusEmail_(sheet, row, status) {
  var r = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  var email = r[2];
  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return;
  var no = r[10] || '';
  var subject = '[Help Desk] Ticket #' + no + ' — ' + status + ' — CB ' + r[0];
  var body = 'Your Chromebook help desk ticket is now: ' + status + '.\n\n' +
    'Ticket #:       ' + no + '\n' +
    'Chromebook S/N: ' + r[0] + '\n' +
    'Issue:          ' + r[5] + '\n\n' +
    (status === 'Resolved'
      ? 'This ticket has been marked resolved. Reply if the problem is not fixed.\n'
      : 'We are working on it and will follow up.\n') +
    '\n' + HELPDESK_EMAIL;
  MailApp.sendEmail(email, subject, body, { name: 'CPA IT Tickets', replyTo: HELPDESK_EMAIL });
}

// Count of OPEN (not Resolved) tickets for a given S/N in the live sheet. Public (no token).
function openCount_(p) {
  var sn = String(p.sn || '').trim().toLowerCase();
  if (!sn) return { ok: true, count: 0 };
  var sheet = firstSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };
  var v = sheet.getRange(2, 1, lastRow - 1, Math.max(9, sheet.getLastColumn())).getValues();
  var c = 0;
  v.forEach(function (r) {
    if (String(r[0]).trim().toLowerCase() === sn && (r[8] || 'New') !== 'Resolved') c++;
  });
  return { ok: true, count: c };
}

// Lifetime aggregates across the live sheet AND every archive tab.
function stats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var byDevice = {}, byStudent = {}, resSum = 0, resCount = 0;
  ss.getSheets().forEach(function (sh) {
    var lr = sh.getLastRow();
    if (lr < 2) return;
    var lc = Math.max(HEADERS.length, sh.getLastColumn());
    var v = sh.getRange(2, 1, lr - 1, lc).getValues();
    v.forEach(function (r) {
      if (!r[0]) return;
      var snk = String(r[0]).trim();
      if (snk) byDevice[snk] = (byDevice[snk] || 0) + 1;
      var stu = String(r[11] || '').trim();
      if (stu) byStudent[stu] = (byStudent[stu] || 0) + 1;
      if (r[8] === 'Resolved' && r[1] && r[13]) {
        var d = (new Date(r[13]) - new Date(r[1])) / 86400000;
        if (d >= 0) { resSum += d; resCount++; }
      }
    });
  });
  function top(o) {
    return Object.keys(o).map(function (k) { return { label: k, value: o[k] }; })
      .sort(function (a, b) { return b.value - a.value; }).slice(0, 12);
  }
  return {
    ok: true, byDevice: top(byDevice), byStudent: top(byStudent),
    avgResolutionDays: resCount ? (resSum / resCount) : null, resolvedCount: resCount
  };
}

// ---- Device history lookup ----
// Given a serial: where it lives (cart/teacher/room/Chromebook #/student) + every
// past ticket for it (live sheet + all archive tabs). Uses createTextFinder so the
// search happens in one optimized pass per workbook rather than tab-by-tab.
function deviceLookup_(p) {
  var sn = String(p.sn || '').trim();
  if (!sn) return { ok: false, error: 'No serial provided.' };
  var out = { ok: true, sn: sn, assignments: [], tickets: [] };

  // 1) Roster assignment — serials live in column B of tabs whose B2 says "Serial #".
  try {
    var rs = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    rs.createTextFinder(sn).matchEntireCell(true).findAll().forEach(function (rng) {
      if (rng.getColumn() !== 2) return;                 // ignore non-serial columns
      var sh = rng.getSheet();
      var hdr = String(sh.getRange(2, 2).getValue() || '').toLowerCase();
      if (hdr.indexOf('serial') < 0) return;             // not a roster tab
      var row = rng.getRow();
      if (row < 3) return;
      out.assignments.push({
        cart: sh.getName(),
        teacher: String(sh.getRange(1, 1).getValue() || ''),
        room: String(sh.getRange(1, 2).getValue() || ''),
        chromebookNo: String(sh.getRange(row, 1).getValue() || ''),
        student: String(sh.getRange(row, 3).getValue() || '')
      });
    });
  } catch (e) { out.rosterError = String(e); }

  // 2) Ticket history — S/N is column A in the live sheet and every archive tab.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.createTextFinder(sn).matchEntireCell(true).findAll().forEach(function (rng) {
      if (rng.getColumn() !== 1) return;
      var sh = rng.getSheet();
      var row = rng.getRow();
      if (row < 2) return;
      var r = sh.getRange(row, 1, 1, Math.max(HEADERS.length, sh.getLastColumn())).getValues()[0];
      out.tickets.push({
        sheet: sh.getName(),
        ticketNo: r[10] || '',
        timestamp: r[1] ? new Date(r[1]).toISOString() : '',
        issue: r[5] || '', urgency: r[6] || '', status: r[8] || 'New',
        notes: r[9] || '', studentAtFault: r[11] || '',
        description: r[7] || '', photoUrl: r[14] || ''
      });
    });
    out.tickets.sort(function (a, b) { return String(b.timestamp).localeCompare(String(a.timestamp)); });
  } catch (e) { out.ticketError = String(e); }

  return out;
}

// ---- Photos ----
// Saves a base64 data-URL photo into the Drive folder and returns its shareable URL.
function savePhoto_(dataUrl, name) {
  if (!dataUrl || String(dataUrl).indexOf('data:') !== 0) return '';
  var m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return '';
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name || 'photo.jpg');
  var file = photoFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file.getUrl();
}

function photoFolder_() {
  try {
    return DriveApp.getFolderById(PHOTO_FOLDER_ID);   // the folder you provided
  } catch (e) {
    // No access to that folder — fall back so photos are never lost.
    var it = DriveApp.getFoldersByName(PHOTO_FOLDER_FALLBACK);
    return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_FALLBACK);
  }
}

// ---- Monthly archive ----
function setupMonthlyArchive() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'archiveMonthly') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archiveMonthly').timeBased().onMonthDay(1).atHour(1).create();
}
function archiveMonthly() { archiveCopy_(true); }
function archiveCopy_(clear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheets()[0];
  var lastRow = src.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'No tickets to archive.' };
  var lastCol = Math.max(HEADERS.length, src.getLastColumn());
  var prev = new Date();
  prev = new Date(prev.getFullYear(), prev.getMonth() - 1, 1);
  var name = MONTHS[prev.getMonth()] + String(prev.getFullYear()).slice(-2) + '_Tickets';
  if (ss.getSheetByName(name)) name += '_' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMddHHmm');
  var dest = ss.insertSheet(name, ss.getNumSheets());
  var all = src.getRange(1, 1, lastRow, lastCol).getValues();
  dest.getRange(1, 1, all.length, lastCol).setValues(all);
  dest.getRange(1, 1, 1, lastCol).setFontWeight('bold');
  dest.setFrozenRows(1);
  if (clear) src.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  return { ok: true, name: name, rows: lastRow - 1, cleared: !!clear };
}

// ---- Ticket submissions ----
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (err) { data = (e.parameter || {}); }
    } else { data = (e && e.parameter) || {}; }

    var sheet = firstSheet_();
    ensureHeaders_(sheet);

    var props = PropertiesService.getScriptProperties();
    var no = (parseInt(props.getProperty('lastTicketNo'), 10) || 1000) + 1;
    props.setProperty('lastTicketNo', String(no));

    var photoUrl = '';
    try {
      photoUrl = savePhoto_(data.photo, 'CB_' + (data.sn || 'unknown') + '_ticket' + no + '.jpg');
    } catch (e) { photoUrl = ''; }   // never fail a ticket because of a photo

    var now = new Date();
    sheet.appendRow([
      data.sn || '', now, data.email || '', data.name || '', data.room || '',
      data.issue || '', data.urgency || '', data.description || '', 'New', '',
      no, data.studentAtFault || '', '', '', photoUrl
    ]);

    sendEmail_(data, now, no, photoUrl);
    return jsonOut_({ ok: true, ticketNo: no });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function sendEmail_(data, now, no, photoUrl) {
  var subject = '[Help Desk] Ticket #' + no + ' — ' + (data.issue || 'Ticket') +
                ' — CB ' + (data.sn || '?') + ' (' + (data.urgency || 'Medium') + ')';
  var lines = [
    'A new Chromebook help desk ticket was submitted.', '',
    'Ticket #:        ' + no,
    'Chromebook S/N:  ' + (data.sn || ''),
    'Issue type:      ' + (data.issue || ''),
    'Urgency:         ' + (data.urgency || ''),
    'Student at fault:' + (data.studentAtFault ? ' ' + data.studentAtFault : ' (none)'),
    '',
    'Description:', (data.description || ''), '',
    'Submitted by:    ' + (data.name || '(no name)'),
    'Teacher email:   ' + (data.email || '(none)'),
    'Room #:          ' + (data.room || ''),
    'Submitted at:    ' + now
  ];
  if (photoUrl) lines.push('', 'Photo: ' + photoUrl);
  lines.push('', HELPDESK_EMAIL);
  var valid = data.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(data.email);
  var recipient = valid ? data.email : HELPDESK_EMAIL;
  MailApp.sendEmail(recipient, subject, lines.join('\n'),
    { name: 'CPA IT Tickets', cc: HELPDESK_EMAIL, replyTo: HELPDESK_EMAIL });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
