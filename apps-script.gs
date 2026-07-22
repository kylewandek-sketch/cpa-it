var HELPDESK_EMAIL = 'kyle.anderson@cpaohio.org';
var ADMIN_TOKEN = 'CHANGE_ME';   // set your own; do NOT commit the real token to a public repo

// Drive folder that ticket photos are saved into. The account running this script
// must have EDIT access to it. Falls back to a folder on the script's own Drive.
var PHOTO_FOLDER_ID = '1CTMn-eBkvMjUN69ALhYd0UvO71Cc0mUN';
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
  var guarded = ['list', 'update', 'archiveTest', 'stats', 'lookup',
                 'todoList', 'todoAdd', 'todoUpdate', 'todoDelete', 'todoReorder'];
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
  } else if (p.action === 'todoList') {
    out = todoList_();
  } else if (p.action === 'todoAdd') {
    out = todoAdd_(p);
  } else if (p.action === 'todoUpdate') {
    out = todoUpdate_(p);
  } else if (p.action === 'todoDelete') {
    out = todoDelete_(p);
  } else if (p.action === 'todoReorder') {
    out = todoReorder_(p);
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

// ---- To-Do list (dashboard "To-Do" tab) ----
// Items live in a 'Todos' sheet tab: ID | Text | Done | Order | Created.
var TODO_SHEET_NAME = 'Todos';
var TODO_HEADERS = ['ID', 'Text', 'Done', 'Order', 'Created'];

function todoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TODO_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(TODO_SHEET_NAME, ss.getNumSheets());
    sh.getRange(1, 1, 1, TODO_HEADERS.length).setValues([TODO_HEADERS]);
    sh.getRange(1, 1, 1, TODO_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function todoList_() {
  var sh = todoSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, todos: [] };
  var v = sh.getRange(2, 1, lastRow - 1, TODO_HEADERS.length).getValues();
  var todos = [];
  v.forEach(function (r) {
    if (!r[0]) return;
    var done = false;
    if (r[2] === true || r[2] === 'TRUE') done = true;
    todos.push({ id: String(r[0]), text: String(r[1]), done: done, order: Number(r[3]) || 0 });
  });
  todos.sort(function (a, b) { return a.order - b.order; });
  return { ok: true, todos: todos };
}

function todoFindRow_(sh, id) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var v = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function todoAdd_(p) {
  var text = String(p.text || '').trim();
  if (!text) return { ok: false, error: 'empty text' };
  var sh = todoSheet_();
  var id = String(new Date().getTime());
  var order = sh.getLastRow();   // new items go to the bottom
  sh.appendRow([id, text, false, order, new Date()]);
  return { ok: true, id: id };
}

function todoUpdate_(p) {
  var sh = todoSheet_();
  var row = todoFindRow_(sh, p.id);
  if (!row) return { ok: false, error: 'not found' };
  if (p.text != null) sh.getRange(row, 2).setValue(String(p.text));
  if (p.done != null) sh.getRange(row, 3).setValue(String(p.done) === 'true');
  return { ok: true };
}

function todoDelete_(p) {
  var sh = todoSheet_();
  var row = todoFindRow_(sh, p.id);
  if (!row) return { ok: false, error: 'not found' };
  sh.deleteRow(row);
  return { ok: true };
}

// ids arrives as a comma-separated list in the new display order.
function todoReorder_(p) {
  var ids = String(p.ids || '').split(',');
  var sh = todoSheet_();
  for (var i = 0; i < ids.length; i++) {
    var row = todoFindRow_(sh, ids[i]);
    if (row) sh.getRange(row, 4).setValue(i + 1);
  }
  return { ok: true };
}

// RUN THIS ONCE from the editor to load the summer 2026 check/repair cross-reference
// findings into the to-do list. It refuses to run if the Todos tab already has items.
function seedTodos() {
  var sh = todoSheet_();
  if (sh.getLastRow() > 1) {
    Logger.log('Todos tab already has items - not seeding again.');
    return 'Todos tab already has items - not seeding again.';
  }
  var items = [
    'Cart O: fix state-testing browser on 20 HP units (2HA99...) - likely enrollment issue, one unit noted not enrolled to @CPAohio.org',
    'Cart O: fix duplicate serial - #6 and #25 both entered with matching serials (2HA99FEN501098M / 2HA99FEN511183W)',
    'Cart C: open tickets - #7, #13, #22 missing from cart; #12, #23, #26, #27 chargers dead; #18 not working + testing browser. Flagged since spring break, never ticketed',
    'Cart G: state-testing browser not working on #25-#28 (also flagged at spring break, no tickets)',
    'Cart Y: #5, #23, #27 not working with dead chargers - no tickets on file',
    'Cart K: #26 missing/broken; #1 power key missing; #23 spacebar missing; #8 hinge dislocating; #4 and #18 chargers dead',
    'Cart K: Adams Smartpass iPad and Karim Shabana iPad both missing/not working',
    'Cart J: #11 charging port missing, #28 screen scratched + no serial label, #9 hinge super loose - identical notes since spring break, never ticketed',
    'Cart A #18 (NXHBNAA0019160FFC07600): in repair since 5/15 (screen will not turn on) - never returned',
    'Cart A (NXHBNAA0019160FE837600): in repair since 4/17 (stuck on white screen) - never returned',
    'Cart B #13 (NXH8VAA0060400FD467611): screen broken, in repair since 1/14 - never returned; note says headphone jack blocked all year',
    'Chase 5 iPads never returned from repair: Ljungren DMQSJ599HGSD (9/3/25), Seggerson DMQPH5ZZFK10 (9/22/25), Wand DMQPH9V4FK10 (9/22/25), Caudill DMQPHDLPFK10 (11/11/25), Buechner DMPPDS94FK10 (3/4/26)',
    'Ljungren Smartpass iPad (DMQPHCTGFK10): screen broken, in repair since 2/27 - never returned; SOY note says Smart Pass not working',
    'Cart D (NXHBNAA0019252726A7600): 3 wifi repairs with the same complaint - replace wifi card or retire the unit',
    'Cart N (G5LG0H3): 3 repair visits (trackpad x2, then screen) - consider retiring',
    'Cart K (NXHBNAA00191610D527600): keyboard still dead after 2 repair visits - repair did not take',
    'Cart D (NXHBNAA001916101BF7600): came back not working after 4/10 repair, in again 4/24 - verify it is actually fixed',
    'Cart H: missing keys on #8, #12, #14, #15, #27; hinges on #10 and #19 - no tickets',
    'Cart H #16: repair log says returned 1/5 but start-of-year check says missing - reconcile',
    'Cart I #4 and #9: marked returned from enrollment fix 1/5 but start-of-year check says not in cart - locate',
    'Cart Y #13 and #25: marked returned 1/5 but start-of-year check says missing - locate',
    'Cart T #26: will not turn on - open ticket',
    'ESL #6: trackpad broken and taped shut - open ticket',
    'Cart B: #1 charger dead since spring break; #23 not working; #3 and #7 missing keys; MP1M1ZX2 number keys acting up again after repair',
    'Cart A: #19 hinge loose; #3 and #7 chargers dead; Eduanny ESL iPad charger missing (flagged both checks)',
    'Caudill iPad #95 was swapped out - update roster (teacher now has 24 iPads)',
    'Miller iPad #44: needs an extra charger',
    'Cart W #28: on loan to SPED - track it and get it back',
    'Cart V #12: state test app missing - reinstall',
    'Cart X #1: keys missing',
    'Cart D #14: D key cap missing (key still works) - since spring break',
    'Re-check Carts AA, F, Y: checker marked every box TRUE (treated checkmark as OK), so Keys Missing / Hinge Broken columns are unreliable',
    'Finish start-of-year checks - untouched sheets: BB (Moorman), Aeh, Perez, Moorman iPads, Title 1, Hunter iPad cart'
  ];
  var now = new Date();
  var base = now.getTime();
  var rows = [];
  for (var i = 0; i < items.length; i++) {
    rows.push([String(base + i), items[i], false, i + 1, now]);
  }
  sh.getRange(2, 1, rows.length, TODO_HEADERS.length).setValues(rows);
  Logger.log('Seeded ' + rows.length + ' to-dos.');
  return 'Seeded ' + rows.length + ' to-dos.';
}

// ---- Photos ----
// RUN THIS ONCE from the editor: it triggers the Drive authorization prompt and
// verifies the script can actually write to PHOTO_FOLDER_ID. Check the log/result.
function testPhotoSetup() {
  var out = [];
  try {
    var f = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    out.push('Folder found: "' + f.getName() + '"');
    var t = f.createFile(Utilities.newBlob('cpa-it test', 'text/plain', 'cpa-it-test.txt'));
    out.push('Write OK: ' + t.getUrl());
    t.setTrashed(true);
    out.push('Cleanup OK — photos will save here.');
  } catch (e) {
    out.push('FAILED: ' + e);
    out.push('If this is an authorization error, approve the Drive prompt and run again.');
    out.push('If it is "not found"/"access denied", the account running this script cannot');
    out.push('edit folder ' + PHOTO_FOLDER_ID + ' — share it with this account as Editor.');
  }
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}
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
    } catch (e) {
      photoUrl = '';                        // never fail a ticket because of a photo
      Logger.log('photo save failed: ' + e); // shows in Executions log
    }

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
