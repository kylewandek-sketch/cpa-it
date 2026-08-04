var SCRIPT_VERSION = '2026-08-04 serial-parse + todo-rebuild';   // shown by checkSetup()

var HELPDESK_EMAIL = 'kyle.anderson@cpaohio.org';
var ADMIN_TOKEN = 'CHANGE_ME';   // set your own; do NOT commit the real token to a public repo

// Drive folder that ticket photos are saved into. The account running this script
// must have EDIT access to it. Falls back to a folder on the script's own Drive.
var PHOTO_FOLDER_ID = '1CTMn-eBkvMjUN69ALhYd0UvO71Cc0mUN';
var PHOTO_FOLDER_FALLBACK = 'CPA IT Ticket Photos';

// Native Google Sheet holding the cart rosters (HS_Cart_1..6, HS_Spares, ...).
// Roster tabs are auto-detected: serials in column B, with "Serial #" in B2.
var ROSTER_SHEET_ID = '1FDVE6KtAEf06_zRYQyHyaNZ_9gXsv3JRJGbIwckv4Mw';

// ---- Roster note settings ----
// When a ticket is opened or moved to In Progress, the teacher's "Describe the
// problem" text is written into the device's row in the master workbook, so the
// roster itself shows why a device is out. Cleared when the ticket is Resolved.
//
// Notes are written into the MIRROR (ROSTER_SHEET_ID), not the master .xlsx --
// SpreadsheetApp cannot open .xlsx files and this account only has view access
// to the master anyway. The mirror is refreshed from the master after a ticket
// comes in; see "Roster mirror" further down.
var NOTE_BOOK_ID = ROSTER_SHEET_ID;

// Cart for a to-do item when the serial is not on any roster tab.
var TODO_GROUP_FALLBACK = 'Unassigned';

// The school-account master workbook. Read-only source for the mirror; nothing
// is ever written back to it.
// This is the workbook the cart notes are typed into.
//   1z1W54tWIvm4XlaqsEbyN4slO2HNCD4FO = "Start of Year ... Check.xlsx" (in use)
//   1phWbiNeULgv0KVErXiodIoAPp_5FWPnL = "2026-2027 Chromebook Carts/Ipads.xlsx"
//     -- the assignment roster; holds students, HS carts and the iPad lists,
//        but not the damage notes
var MASTER_XLSX_ID = '1z1W54tWIvm4XlaqsEbyN4slO2HNCD4FO';

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
  var guarded = ['list', 'update', 'archiveTest', 'stats', 'lookup', 'notesToClear',
                 'todoList', 'todoAdd', 'todoUpdate', 'todoDelete', 'todoReorder',
                 'todoHideGroup', 'todoShowGroup', 'refreshTodos'];
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
  } else if (p.action === 'notesToClear') {
    out = notesToClear_();
  } else if (p.action === 'refreshTodos') {
    out = refreshTodos_();
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
  } else if (p.action === 'todoHideGroup') {
    out = todoHideGroup_(p);
  } else if (p.action === 'todoShowGroup') {
    out = todoShowGroup_(p);
  } else if (p.action === 'snCheck') {
    out = snCheck_(p);            // public: submit form checks the S/N before filing
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
      // roster note: put the description in the device's row while it is being
      // worked on, take it back out once the ticket is closed
      var r = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      if (p.status === 'In Progress') {
        rosterNoteWrite_(r[0], r[10], r[7]);
        ticketTodoWrite_(r[0], r[10], r[7]);
      }
      if (p.status === 'Resolved') {
        rosterNoteClear_(r[0], r[10], r[7]);     // note comes out of the roster
        ticketTodoRemove_(r[10]);                // and the item off the to-do list
      }
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
    'Description:\n' + (r[7] || '(none given)') + '\n\n' +
    (r[9] ? 'IT notes:\n' + r[9] + '\n\n' : '') +
    (status === 'Resolved'
      ? 'This ticket has been marked resolved. Reply if the problem is not fixed.\n'
      : 'We are working on it and will follow up.\n') +
    '\n' + HELPDESK_EMAIL;
  MailApp.sendEmail(email, subject, body, { name: 'CPA IT Tickets', replyTo: HELPDESK_EMAIL });
}

// Is this serial in the roster at all? The submit form asks before filing a
// ticket, so a mistyped S/N gets questioned rather than silently accepted.
// Public (no token); returns nothing about who the device belongs to.
function snCheck_(p) {
  var sn = serialFromScan_(p.sn);
  if (!sn) return { ok: true, found: false };
  try {
    var hits = rosterFindRows_(sn);
    if (!hits.length) return { ok: true, found: false };
    return { ok: true, found: true, cart: rosterPickBest_(hits).sheet.getName() };
  } catch (e) {
    return { ok: false, error: String(e) };   // the form treats an error as "do not block"
  }
}

// Count of OPEN (not Resolved) tickets for a given S/N in the live sheet. Public (no token).
function openCount_(p) {
  var sn = serialFromScan_(p.sn).toLowerCase();
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
  var sn = serialFromScan_(p.sn);
  if (!sn) return { ok: false, error: 'No serial provided.' };
  var out = { ok: true, sn: sn, assignments: [], tickets: [], todos: [] };

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
      // A device can have 1-4 students assigned, one per homeroom: every column
      // whose row-2 header says "Student Assigned" holds a name, and row 1 above
      // it names the homeroom (e.g. "Hepler's HR").
      var lastCol = Math.min(sh.getLastColumn(), 8);
      var heads = sh.getRange(1, 1, 2, lastCol).getValues();   // rows 1 and 2
      var vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
      var students = [];
      for (var c = 3; c <= lastCol; c++) {
        var hdr2 = String(heads[1][c - 1] || '').toLowerCase();
        if (hdr2.indexOf('student') < 0) continue;
        var name = String(vals[c - 1] || '').trim();
        if (!name) continue;
        var hr = String(heads[0][c - 1] || '').trim();
        if (hr) students.push(name + ' (' + hr + ')');
        else students.push(name);
      }
      out.assignments.push({
        cart: sh.getName(),
        teacher: String(heads[0][0] || ''),
        room: String(heads[0][1] || ''),
        chromebookNo: String(vals[0] || ''),
        students: students
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

  // 3) To-do items whose task text mentions this serial (Todos tab, column B).
  //    Only matches tasks that actually contain the serial - cart-level items
  //    without a serial in the text will not show here.
  try {
    var ts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TODO_SHEET_NAME);
    if (ts) {
      ts.createTextFinder(sn).findAll().forEach(function (rng) {
        if (rng.getColumn() !== 2) return;   // task text lives in column B
        var row = rng.getRow();
        if (row < 2) return;
        var r = ts.getRange(row, 1, 1, TODO_HEADERS.length).getValues()[0];
        var done = false;
        if (r[2] === true || r[2] === 'TRUE') done = true;
        out.todos.push({ id: String(r[0]), text: String(r[1]), done: done, group: String(r[5] || '') });
      });
    }
  } catch (e) { out.todoError = String(e); }

  return out;
}

// ---- Roster notes -----------------------------------------------------------
// Finds a serial in the roster workbook and parks the ticket description in the
// device's row. Everything here is best-effort: a missing serial, a locked
// sheet or a bad id must never stop a ticket from being filed or updated.

// Every row in the roster workbook holding this serial. A column counts as a
// serial column when its row-2 header says "Serial #", so tabs that keep the
// serial somewhere other than column B (ESL Translation, MS Staff) work too.
function rosterFindRows_(sn) {
  var hits = [];
  sn = String(sn || '').trim();
  if (!sn) return hits;
  var rs = SpreadsheetApp.openById(NOTE_BOOK_ID);
  rs.createTextFinder(sn).matchEntireCell(true).findAll().forEach(function (rng) {
    var sh = rng.getSheet();
    var row = rng.getRow();
    var col = rng.getColumn();
    if (row < 3) return;
    var hdr = String(sh.getRange(2, col).getValue() || '').toLowerCase();
    if (hdr.indexOf('serial') < 0) return;        // not a serial column
    hits.push({ sheet: sh, row: row });
  });
  return hits;
}

// A serial can appear on several tabs (an iPad on its teacher's tab and again
// on the master iPad list). Prefer the tab that looks like an inventory sheet:
// named like one, listing devices rather than student assignments, and long.
function rosterTabScore_(sh) {
  var name = sh.getName().toLowerCase();
  var score = 0;
  if (name.indexOf('inventory') >= 0) score += 1000;
  if (name.indexOf('ipads') >= 0 || name.indexOf('spares') >= 0) score += 500;
  var lastCol = Math.min(sh.getLastColumn(), 14);   // headers never run past N
  var heads = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  var assigns = false;
  for (var i = 0; i < heads.length; i++) {
    if (String(heads[i]).toLowerCase().indexOf('student') >= 0) assigns = true;
  }
  if (!assigns) score += 200;                     // assignment tabs are not inventory
  score += sh.getLastRow();                       // the master list is the long one
  return score;
}

function rosterPickBest_(hits) {
  var best = null;
  var bestScore = -1;
  for (var i = 0; i < hits.length; i++) {
    var s = rosterTabScore_(hits[i].sheet);
    if (s > bestScore) {
      bestScore = s;
      best = hits[i];
    }
  }
  return best;
}

// The note column for a tab: the first column whose ROW 2 cell is empty. Row 2
// holds the headers (Chromebook # / Serial # / Student Assigned ...), so the
// first blank header is the first unused column, and every note on the tab
// lines up in it.
function rosterNoteColumn_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 3);
  var heads = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  for (var c = 1; c <= lastCol; c++) {
    if (!String(heads[c - 1] || '').trim()) return c;
  }
  return lastCol + 1;                             // every header filled: add a column
}

// Clear the note this ticket left in the row, but only if the cell still holds
// that note -- never wipe something a person typed in the meantime.
function rosterClearInRow_(sh, row, ticketNo, description) {
  var col = rosterNoteColumn_(sh);
  var cur = String(sh.getRange(row, col).getValue() || '').trim();
  if (!cur) return 0;
  if (!description || cur !== String(description).trim()) return 0;   // somebody else's text
  sh.getRange(row, col).clearContent();
  return 1;
}

// Write (or refresh) the note. Skips quietly when the serial is not on any tab.
function rosterNoteWrite_(sn, ticketNo, description) {
  try {
    var text = String(description || '').trim();
    if (!text) return { ok: true, skipped: 'no description' };
    var hits = rosterFindRows_(sn);
    if (!hits.length) return { ok: true, skipped: 'serial not found' };
    var target = rosterPickBest_(hits);
    var col = rosterNoteColumn_(target.sheet);
    target.sheet.getRange(target.row, col).setValue(text);
    return { ok: true, tab: target.sheet.getName(), row: target.row, col: col,
             cbNo: rosterDeviceNumber_(target.sheet, target.row) };
  } catch (e) {
    Logger.log('roster note write failed: ' + e);   // see Executions in the editor
    return { ok: false, error: String(e) };
  }
}

// Remove the note from every tab the serial appears on, so nothing is left
// behind if the device was moved between tabs while the ticket was open.
function rosterNoteClear_(sn, ticketNo, description) {
  try {
    var hits = rosterFindRows_(sn);
    if (!hits.length) return { ok: true, skipped: 'serial not found' };
    var cleared = 0;
    for (var i = 0; i < hits.length; i++) {
      cleared += rosterClearInRow_(hits[i].sheet, hits[i].row, ticketNo, description);
    }
    return { ok: true, cleared: cleared };
  } catch (e) {
    Logger.log('roster note clear failed: ' + e);
    return { ok: false, error: String(e) };
  }
}

// ---- To-dos driven by tickets ----------------------------------------------
// A ticket puts an item on the to-do list, grouped under the cart the device
// belongs to; closing the ticket takes it back off. The item's ID is derived
// from the ticket number, so it can always be found again.

function ticketTodoId_(ticketNo) { return 'ticket-' + ticketNo; }

// The column holding the device's position in the cart -- "Chromebook #" on
// cart tabs, "iPad" on the tablet ones. Falls back to column A.
function rosterNumberColumn_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var heads = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  for (var c = 1; c <= lastCol; c++) {
    var h = String(heads[c - 1] || '').toLowerCase();
    if (h.indexOf('serial') >= 0) continue;
    if (h.indexOf('chromebook') >= 0 || h.indexOf('ipad') >= 0 ||
        h.indexOf('device') >= 0 || h.indexOf('tablet') >= 0) return c;
  }
  return 1;
}

function rosterDeviceNumber_(sh, row) {
  var c = rosterNumberColumn_(sh);
  if (!c) return '';
  var v = sh.getRange(row, c).getValue();
  if (!looksLikePosition_(v)) return '';
  return String(v).trim();
}

// Cart and position for a serial, or null when it is not on any tab.
function rosterDeviceInfo_(sn) {
  try {
    var hits = rosterFindRows_(sn);
    if (!hits.length) return null;
    var best = rosterPickBest_(hits);
    return {
      cart: best.sheet.getName(),
      cbNo: rosterDeviceNumber_(best.sheet, best.row)
    };
  } catch (e) {
    Logger.log('device lookup failed: ' + e);
    return null;
  }
}

// One place builds the trailing detail, so the reconcile step can rebuild the
// exact same text and tell a real edit from a formatting difference.
function deviceTag_(cbNo, sn) {
  if (cbNo) return '  (CB #' + cbNo + ' \u00b7 S/N ' + sn + ')';
  return '  (S/N ' + sn + ')';
}

// Add the item, or refresh it if this ticket already has one.
function ticketTodoWrite_(sn, ticketNo, description, info) {
  try {
    var text = String(description || '').trim();
    if (!text || !ticketNo) return { ok: true, skipped: 'nothing to add' };
    if (!info || !info.cart) info = rosterDeviceInfo_(sn) || { cart: '', cbNo: '' };
    var cart = info.cart || TODO_GROUP_FALLBACK;
    var tail = '  (Ticket #' + ticketNo;
    if (info.cbNo) tail += ' \u00b7 CB #' + info.cbNo;
    tail += ' \u00b7 S/N ' + sn + ')';
    var label = text + tail;
    var sh = todoSheet_();
    var id = ticketTodoId_(ticketNo);
    var row = todoFindRow_(sh, id);
    if (row) {
      sh.getRange(row, 2).setValue(label);
      sh.getRange(row, 6).setValue(cart);
      return { ok: true, updated: id, group: cart };
    }
    sh.appendRow([id, label, false, sh.getLastRow(), new Date(), cart, TODO_STATUS_DEFAULT]);
    return { ok: true, added: id, group: cart };
  } catch (e) {
    Logger.log('ticket todo write failed: ' + e);
    return { ok: false, error: String(e) };
  }
}

function ticketTodoRemove_(ticketNo) {
  try {
    if (!ticketNo) return { ok: true, skipped: 'no ticket number' };
    var sh = todoSheet_();
    var row = todoFindRow_(sh, ticketTodoId_(ticketNo));
    if (!row) return { ok: true, skipped: 'no item for this ticket' };
    sh.deleteRow(row);
    return { ok: true, removed: ticketTodoId_(ticketNo) };
  } catch (e) {
    Logger.log('ticket todo remove failed: ' + e);
    return { ok: false, error: String(e) };
  }
}

// ---- Notes already written in the carts -------------------------------------
// Devices carry hand-typed notes in the roster long before any ticket existed.
// These get pulled onto the to-do list, one item per note, keyed to the serial.
// Safe to run repeatedly: an item is only added when nothing similar is there.

// Columns that can hold a note: a blank row-2 header (the column the ticket
// notes go in) or a header that reads like one. Student columns never count.
var NOTE_HEADER_RE = /(note|issue|repair|comment|problem|damage|broken|status)/i;

function rosterNoteColumns_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 3);
  var heads = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  var cols = [];
  for (var c = 3; c <= lastCol; c++) {
    var h = String(heads[c - 1] || '').trim();
    if (h.toLowerCase().indexOf('student') >= 0) continue;
    if (!h) { cols.push(c); continue; }
    if (NOTE_HEADER_RE.test(h)) cols.push(c);
  }
  return cols;
}

// Which column on this tab holds the serials.
function rosterSerialColumn_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 2);
  var heads = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  for (var c = 1; c <= lastCol; c++) {
    if (String(heads[c - 1] || '').toLowerCase().indexOf('serial') >= 0) return c;
  }
  return 0;
}

// Plenty of cells in a roster row are not notes: ticked checkboxes come back as
// the boolean true, and columns like "Computer Working" hold yes/no flags,
// counts or dates. A note has to read like words.
var NOTE_JUNK_RE = /^(true|false|yes|no|y|n|x|na|n\/a|ok|okay|good|new|none|null|-|--|—|✓|✔)$/i;

function isRealNote_(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === 'boolean') return false;         // a ticked checkbox
  if (typeof value === 'number') return false;          // a count or a serial
  if (Object.prototype.toString.call(value) === '[object Date]') return false;
  var s = String(value).trim();
  if (s.length < 3) return false;
  if (NOTE_JUNK_RE.test(s)) return false;
  if (!/[a-z]{2}/i.test(s)) return false;               // needs actual words in it
  return true;
}

// Roster tabs can hold more than one table -- a cart's Chromebook list with an
// iPad list below it, each with its own header row. A header row read as data
// gives nonsense like serial "Serial #", so every candidate row is checked.
// Serials here run 6+ characters, letters and digits, no spaces or symbols
// (Dell tags are 7, Acer 22, Apple 12).
function looksLikeRosterSerial_(value) {
  if (typeof value !== 'string') value = String(value == null ? '' : value);
  var v = value.trim();
  if (v.length < 6 || v.length > 30) return false;
  if (!/^[A-Za-z0-9]+$/.test(v)) return false;
  if (!/[0-9]/.test(v)) return false;
  if (!/[A-Za-z]/.test(v)) return false;
  return true;
}

// The position in the cart is a plain number. Anything else in that column --
// a student name from a second table, a label -- is not a position.
function looksLikePosition_(value) {
  var v = String(value == null ? '' : value).trim();
  return /^[0-9]{1,4}$/.test(v);
}

function normalizeNote_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
}

// Is this note already represented on the to-do list for this serial?
function todoAlreadyHas_(todos, sn, note) {
  var want = normalizeNote_(note);
  if (!want) return true;
  var snLower = String(sn).toLowerCase();
  for (var i = 0; i < todos.length; i++) {
    var text = normalizeNote_(todos[i].text);
    if (text.indexOf(snLower) >= 0 && (text.indexOf(want) >= 0 || want.indexOf(text) >= 0)) return true;
    if (text === want) return true;
  }
  return false;
}

// Walk every roster tab in the mirror and put any note found onto the to-do
// list. Returns a count; details go to the log.
function importRosterNotes_() {
  var added = 0;
  var refreshed = 0;
  var seen = 0;
  var suppressed = 0;
  // Notes already dealt with: ticked off, or the item deleted. They stay in the
  // workbook until somebody removes them by hand, so without this every refresh
  // would put the item straight back.
  var handled = {};
  var handledList = handledNotes_();
  for (var q = 0; q < handledList.length; q++) handled[handledList[q]] = true;
  var todos = todoList_().todos;
  var sh = todoSheet_();
  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var tabs = ss.getSheets();
  for (var t = 0; t < tabs.length; t++) {
    var tab = tabs[t];
    if (mirrorIsProtectedTab_(tab.getName())) continue;
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;                       // not a roster tab
    var noteCols = rosterNoteColumns_(tab);
    if (!noteCols.length) continue;
    var numCol = rosterNumberColumn_(tab);
    var lastRow = tab.getLastRow();
    if (lastRow < 3) continue;
    var lastCol = tab.getLastColumn();
    var heads2 = tab.getRange(2, 1, 1, lastCol).getValues()[0];
    var vals = tab.getRange(3, 1, lastRow - 2, lastCol).getValues();
    for (var r = 0; r < vals.length; r++) {
      var sn = String(vals[r][snCol - 1] || '').trim();
      if (!looksLikeRosterSerial_(sn)) continue;     // header row or blank
      var cbNo = '';
      if (looksLikePosition_(vals[r][numCol - 1])) cbNo = String(vals[r][numCol - 1]).trim();
      for (var k = 0; k < noteCols.length; k++) {
        var raw = vals[r][noteCols[k] - 1];
        if (!isRealNote_(raw)) continue;             // checkbox, flag, number, date
        var note = String(raw).trim();
        var colHead = String(heads2[noteCols[k] - 1] || '').trim();
        if (colHead && note.toLowerCase() === colHead.toLowerCase()) continue;   // repeated header
        seen++;
        if (handled[noteKey_(sn, note)]) {          // already dealt with
          suppressed++;
          continue;
        }
        var where = tab.getName();
        if (cbNo) where += ' #' + cbNo;
        Logger.log(where + ' | ' + sn + ' | ' + note);
        var id = 'roster-' + sn + '-' + noteCols[k];
        var label = note + deviceTag_(cbNo, sn);
        var existing = todoFindRow_(sh, id);
        if (existing) {                              // keep older items current
          var curText = String(sh.getRange(existing, 2).getValue() || '');
          if (curText !== label) {
            sh.getRange(existing, 2).setValue(label);
            sh.getRange(existing, 6).setValue(tab.getName());
            refreshed++;
          }
          continue;
        }
        if (todoAlreadyHas_(todos, sn, note)) continue;
        sh.appendRow([id, label, false, sh.getLastRow(), new Date(), tab.getName(),
                      TODO_STATUS_DEFAULT]);
        todos.push({ id: id, text: label, done: false, group: tab.getName(),
                     status: TODO_STATUS_DEFAULT });
        added++;
      }
    }
  }
  Logger.log('roster notes found: ' + seen + ', new items added: ' + added +
             ', existing items updated: ' + refreshed +
             ', already dealt with: ' + suppressed);
  return added;
}

// A note that has been dealt with. The master .xlsx cannot be written to, so
// clearing a note in the mirror alone would not stick -- the next sync copies
// it straight back from the master. Keeping the handled ones on a list stops
// them being imported again. Editing the note in the master changes its key,
// so a genuinely new note still comes through.
var NOTE_DONE_KEY = 'handledRosterNotes';
var NOTE_DONE_MAX = 500;

function noteKey_(sn, text) { return String(sn).toLowerCase() + '|' + normalizeNote_(text); }

function handledNotes_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(NOTE_DONE_KEY) || '[]');
  } catch (e) {
    return [];
  }
}

function markNoteHandled_(sn, text) {
  var list = handledNotes_();
  var key = noteKey_(sn, text);
  if (list.indexOf(key) < 0) list.push(key);
  while (list.length > NOTE_DONE_MAX) list.shift();
  PropertiesService.getScriptProperties().setProperty(NOTE_DONE_KEY, JSON.stringify(list));
}

// Forget every handled note, so the next import brings them all back. Useful if
// the list ever gets out of step with the roster.
function clearHandledNotes() {
  PropertiesService.getScriptProperties().deleteProperty(NOTE_DONE_KEY);
  Logger.log('handled-note list cleared');
}

// ---- Closed log -------------------------------------------------------------
// What was closed and when, kept for a week. Used to show which notes are still
// waiting to be deleted from the master workbook, and how long they have waited.
var CLOSED_LOG_SHEET = 'ClosedLog';
var CLOSED_LOG_HEADERS = ['Closed At', 'Item ID', 'Note', 'S/N', 'Cart', 'How'];
var CLOSED_LOG_DAYS = 7;

function closedLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CLOSED_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CLOSED_LOG_SHEET, ss.getNumSheets());
    sh.getRange(1, 1, 1, CLOSED_LOG_HEADERS.length).setValues([CLOSED_LOG_HEADERS]);
    sh.getRange(1, 1, 1, CLOSED_LOG_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Drop anything older than a week, so the log stays a working list.
function closedLogPrune_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var cutoff = new Date().getTime() - CLOSED_LOG_DAYS * 86400000;
  var when = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = when.length - 1; i >= 0; i--) {
    var t = new Date(when[i][0]).getTime();
    if (!t || t < cutoff) sh.deleteRow(i + 2);
  }
}

function logClosed_(id, note, sn, cart, how) {
  try {
    var sh = closedLogSheet_();
    closedLogPrune_(sh);
    sh.appendRow([new Date(), String(id || ''), String(note || ''),
                  String(sn || ''), String(cart || ''), String(how || '')]);
  } catch (e) {
    Logger.log('closed log write failed: ' + e);
  }
}

// Everything closed in the last week, newest first.
function closedHistory_() {
  var out = [];
  try {
    var sh = closedLogSheet_();
    closedLogPrune_(sh);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return out;
    var v = sh.getRange(2, 1, lastRow - 1, CLOSED_LOG_HEADERS.length).getValues();
    for (var i = 0; i < v.length; i++) {
      if (!v[i][0]) continue;
      out.push({
        closedAt: new Date(v[i][0]).toISOString(),
        id: String(v[i][1]), note: String(v[i][2]),
        sn: String(v[i][3]), cart: String(v[i][4]), how: String(v[i][5])
      });
    }
    out.sort(function (a, b) { return String(b.closedAt).localeCompare(String(a.closedAt)); });
  } catch (e) {
    Logger.log('closed history read failed: ' + e);
  }
  return out;
}

// Where a roster-sourced to-do came from: id is roster-<serial>-<column>, and
// the cart tab is the item's group.
function rosterTodoSource_(todo) {
  var m = String(todo.id || '').match(/^roster-(.+)-(\d+)$/);
  if (!m) return null;
  var sn = m[1];
  var col = parseInt(m[2], 10);
  try {
    var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
    var tab = ss.getSheetByName(todo.group);
    if (!tab) return null;
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) return null;
    var lastRow = tab.getLastRow();
    if (lastRow < 3) return null;
    var col_vals = tab.getRange(3, snCol, lastRow - 2, 1).getValues();
    for (var i = 0; i < col_vals.length; i++) {
      if (String(col_vals[i][0]).trim() === sn) {
        return { sheet: tab, row: i + 3, col: col, sn: sn };
      }
    }
  } catch (e) {
    Logger.log('could not locate source of ' + todo.id + ': ' + e);
  }
  return null;
}

// Square the roster-sourced items with the roster itself:
//   ticked off  -> clear the cell in the mirror, remember it, drop the item
//   cell blank  -> the note is gone, so drop the item
//   text edited -> update the item to match
function reconcileRosterTodos_() {
  var sh = todoSheet_();
  var todos = todoList_().todos;
  var dropped = 0;
  var updated = 0;
  for (var i = 0; i < todos.length; i++) {
    var todo = todos[i];
    if (String(todo.id).indexOf('roster-') !== 0) continue;
    var src = rosterTodoSource_(todo);
    if (!src) continue;
    var cell = src.sheet.getRange(src.row, src.col);
    var cur = String(cell.getValue() || '').trim();

    if (todo.done) {
      // The note itself is left alone -- the master workbook is not ours to
      // edit. It goes on the handled list so it is not imported again, and
      // shows up in the "notes to clear" report for someone to delete by hand.
      if (cur) {
        markNoteHandled_(src.sn, cur);
        logClosed_(todo.id, cur, src.sn, todo.group, 'ticked off');
      }
      var doneRow = todoFindRow_(sh, todo.id);
      if (doneRow) { sh.deleteRow(doneRow); dropped++; }
      continue;
    }
    if (!cur) {                                   // note deleted in the roster
      var goneRow = todoFindRow_(sh, todo.id);
      if (goneRow) { sh.deleteRow(goneRow); dropped++; }
      continue;
    }
    var wanted = cur + deviceTag_(rosterDeviceNumber_(src.sheet, src.row), src.sn);
    if (String(todo.text) !== wanted) {            // note reworded in the roster
      var editRow = todoFindRow_(sh, todo.id);
      if (editRow) { sh.getRange(editRow, 2).setValue(wanted); updated++; }
    }
  }
  Logger.log('roster to-dos reconciled: ' + dropped + ' removed, ' + updated + ' updated');
  return dropped;
}

// Notes that have been dealt with but are still sitting in the workbook. The
// script cannot edit the master .xlsx, so this is the list to work through by
// hand. A note drops off it once it is gone from the master.
function notesToClear_() {
  var handled = handledNotes_();
  if (!handled.length) return { ok: true, notes: [] };
  var wanted = {};
  for (var i = 0; i < handled.length; i++) wanted[handled[i]] = true;

  var out = [];
  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var tabs = ss.getSheets();
  for (var t = 0; t < tabs.length; t++) {
    var tab = tabs[t];
    if (mirrorIsProtectedTab_(tab.getName())) continue;
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;
    var noteCols = rosterNoteColumns_(tab);
    if (!noteCols.length) continue;
    var lastRow = tab.getLastRow();
    if (lastRow < 3) continue;
    var vals = tab.getRange(3, 1, lastRow - 2, tab.getLastColumn()).getValues();
    for (var r = 0; r < vals.length; r++) {
      var sn = String(vals[r][snCol - 1] || '').trim();
      if (!sn) continue;
      for (var k = 0; k < noteCols.length; k++) {
        var raw = vals[r][noteCols[k] - 1];
        if (!isRealNote_(raw)) continue;          // checkbox, flag, number, date
        var note = String(raw).trim();
        if (!wanted[noteKey_(sn, note)]) continue;
        out.push({
          cart: tab.getName(),
          sn: sn,
          note: note,
          cell: tab.getRange(r + 3, noteCols[k]).getA1Notation(),
          key: noteKey_(sn, note)
        });
      }
    }
  }
  // Pair each one with when it was closed, so the report can say how long it
  // has been waiting.
  var history = closedHistory_();
  var when = {};
  for (var h = 0; h < history.length; h++) {
    var k = noteKey_(history[h].sn, history[h].note);
    if (!when[k]) when[k] = history[h].closedAt;      // history is newest first
  }
  for (var n = 0; n < out.length; n++) {
    if (when[out[n].key]) out[n].closedAt = when[out[n].key];
  }
  return { ok: true, notes: out, history: history };
}

// Clears out roster items that were never really notes -- an earlier import
// turned ticked checkboxes into to-dos reading "true". Only touches items whose
// id starts with roster-, so anything typed by hand is left alone.
function purgeBogusTodos() {
  var sh = todoSheet_();
  var todos = todoList_().todos;
  var removed = [];
  for (var i = 0; i < todos.length; i++) {
    var todo = todos[i];
    if (String(todo.id).indexOf('roster-') !== 0) continue;
    var text = String(todo.text || '').replace(/\s*\([^)]*S\/N[^)]*\)\s*$/, '').trim();
    if (isRealNote_(text)) continue;
    var row = todoFindRow_(sh, todo.id);
    if (row) {
      sh.deleteRow(row);
      removed.push(todo.group + ': ' + todo.text);
    }
  }
  Logger.log('removed ' + removed.length + ' bogus item(s)');
  for (var k = 0; k < removed.length; k++) Logger.log('  ' + removed[k]);
  return removed.length;
}

// Everything the scheduled job does, on demand from the dashboard's Refresh
// button, and it hands the finished list straight back so the page does not
// have to ask again.
// The mirror sync is FORCED here: pressing Refresh should fetch, not decide the
// master looks unchanged and skip.
function refreshTodos_() {
  var sync = syncMirror_(true);
  var res = rebuildRosterTodos_();
  rebuildTicketTodos();
  var list = todoList_();
  return { ok: true, added: res.added, dropped: res.removed,
           imported: res.imported, kept: res.kept,
           sync: sync, todos: list.todos, hidden: list.hidden,
           carts: rosterCartTabs_() };
}

// ---- Full rebuild of the roster-sourced to-dos ------------------------------
// Refresh used to MERGE: importRosterNotes_ added what was new and
// reconcileRosterTodos_ was the only thing that ever deleted. Reconcile located
// a note's source by tab name and column index, and KEPT anything it could not
// find -- so renaming a cart tab in the master orphaned every item stamped with
// the old name, and the item could never be cleared again. Same story if a
// device moved carts.
//
// Rebuilding sidesteps all of it: the list is only ever what the master says
// now. The sheet is a mirror of the master, not a place things are kept.
//
// The ONLY rows carried over are ticket- items, which rebuildTicketTodos()
// maintains from the ticket sheet; dropping them here would just have that
// function write them straight back with fresh timestamps. Everything else is
// rebuilt from the master, so a row typed into the sheet by hand does not
// survive a refresh. That is deliberate -- the dashboard's Add box was removed
// when this became a pure mirror on 2026-08-04, so nothing offers to create a
// row that would then be discarded.
//
// Status and Done are carried across on a key of serial + note text, which
// survives the note moving tab or column -- a rebuild should not throw away
// work you did on an item that is still there.

// The trailing "  (CB #4 · S/N ABC123)" that deviceTag_ adds, so the original
// note text can be recovered from a stored label.
var TODO_TAG_RE = /\s*\([^)]*S\/N[^)]*\)\s*$/;

function todoStateKey_(sn, noteText) {
  return noteKey_(sn, String(noteText).replace(TODO_TAG_RE, '').trim());
}

// Serial out of a roster- id. Ids are roster-<serial>-<column>.
function todoIdSerial_(id) {
  var m = String(id || '').match(/^roster-(.+)-(\d+)$/);
  return m ? m[1] : '';
}

function rebuildRosterTodos_() {
  var sh = todoSheet_();
  var width = TODO_HEADERS.length;

  // ---- 1. read the sheet as it stands, raw, so Created survives ----
  var keep = [];        // ticket- rows only; everything else is rebuilt
  var before = {};      // state key -> { done, status }
  var beforeKeys = {};
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var rows = sh.getRange(2, 1, lastRow - 1, width).getValues();
    for (var i = 0; i < rows.length; i++) {
      var id = String(rows[i][0] || '');
      if (!id) continue;
      if (id.indexOf('ticket-') === 0) {
        keep.push(rows[i]);                       // rebuildTicketTodos_ owns these
        continue;
      }
      if (id.indexOf('roster-') !== 0) continue;  // stray row: not from the master, dropped
      var sn = todoIdSerial_(id);
      if (!sn) continue;
      var st = String(rows[i][6] || '').trim();
      if (TODO_STATUSES.indexOf(st) < 0) st = TODO_STATUS_DEFAULT;
      var key = todoStateKey_(sn, rows[i][1]);
      before[key] = { done: (rows[i][2] === true || rows[i][2] === 'TRUE'), status: st };
      beforeKeys[key] = true;
    }
  }

  // ---- 2. read every roster note in the mirror ----
  var fresh = [];
  var afterKeys = {};
  var handled = {};
  var handledList = handledNotes_();
  for (var q = 0; q < handledList.length; q++) handled[handledList[q]] = true;

  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var tabs = ss.getSheets();
  for (var t = 0; t < tabs.length; t++) {
    var tab = tabs[t];
    if (mirrorIsProtectedTab_(tab.getName())) continue;
    if (isHiddenGroup_(tab.getName())) continue;    // cart deleted from the to-do page
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;                          // not a roster tab
    var noteCols = rosterNoteColumns_(tab);
    if (!noteCols.length) continue;
    var numCol = rosterNumberColumn_(tab);
    var tabLastRow = tab.getLastRow();
    if (tabLastRow < 3) continue;
    var tabLastCol = tab.getLastColumn();
    var heads2 = tab.getRange(2, 1, 1, tabLastCol).getValues()[0];
    var vals = tab.getRange(3, 1, tabLastRow - 2, tabLastCol).getValues();

    for (var r = 0; r < vals.length; r++) {
      var sn2 = String(vals[r][snCol - 1] || '').trim();
      if (!looksLikeRosterSerial_(sn2)) continue;  // header row of a second table, or blank
      var cbNo = '';
      if (looksLikePosition_(vals[r][numCol - 1])) cbNo = String(vals[r][numCol - 1]).trim();

      for (var k = 0; k < noteCols.length; k++) {
        var raw = vals[r][noteCols[k] - 1];
        if (!isRealNote_(raw)) continue;           // checkbox, flag, number, date
        var note = String(raw).trim();
        var colHead = String(heads2[noteCols[k] - 1] || '').trim();
        if (colHead && note.toLowerCase() === colHead.toLowerCase()) continue;

        var key2 = todoStateKey_(sn2, note);
        // One item per device per note, even where a serial is listed on more
        // than one tab. Two DIFFERENT notes on a device still give two items.
        if (afterKeys[key2]) continue;
        afterKeys[key2] = true;

        var prev = before[key2] || { done: false, status: TODO_STATUS_DEFAULT };

        // A note that was ticked off but is still typed in the master belongs on
        // the "notes to clear" report. Log it once, not on every refresh.
        if (prev.done && !handled[noteKey_(sn2, note)]) {
          markNoteHandled_(sn2, note);
          logClosed_('roster-' + sn2 + '-' + noteCols[k], note, sn2, tab.getName(), 'ticked off');
          handled[noteKey_(sn2, note)] = true;
        }

        fresh.push(['roster-' + sn2 + '-' + noteCols[k],
                    note + deviceTag_(cbNo, sn2),
                    prev.done,
                    0,                             // order filled in below
                    new Date(),
                    tab.getName(),
                    prev.status]);
      }
    }
  }

  // ---- 3. write the sheet back: kept rows first, then the fresh import ----
  // Kept rows keep the Order they already had, so anything dragged into place on
  // the dashboard stays put. Fresh rows go after them.
  var order = 0;
  var out = [];
  for (var a = 0; a < keep.length; a++) {
    var ord = Number(keep[a][3]) || 0;
    if (ord > order) order = ord;
    out.push(keep[a]);
  }
  for (var b = 0; b < fresh.length; b++) {
    fresh[b][3] = ++order;
    out.push(fresh[b]);
  }

  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, width).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, width).setValues(out);
  SpreadsheetApp.flush();

  // ---- 4. count what actually moved, for the dashboard toast ----
  var added = 0;
  var removed = 0;
  for (var ka in afterKeys) { if (!beforeKeys[ka]) added++; }
  for (var kb in beforeKeys) { if (!afterKeys[kb]) removed++; }

  Logger.log('todo rebuild: ' + fresh.length + ' roster item(s) from the master, ' +
             keep.length + ' ticket item(s) kept, ' +
             added + ' new, ' + removed + ' gone');
  return { imported: fresh.length, kept: keep.length, added: added, removed: removed };
}

// Which carts actually exist, by name, straight off the mirror. The dashboard
// draws its grid from this instead of a hardcoded list, so a tile whose tab has
// been renamed or deleted stops being drawn, and a tab added to the master
// turns up on its own.
//
// Deliberately NOT done by auto-hiding the missing ones: hiddenTodoGroups is
// permanent, so a cart auto-hidden today would stay suppressed if a tab of the
// same name were created later. This list is recomputed every time instead.
//
// Hidden carts are included here -- the dashboard subtracts them itself, and
// needs the names to offer them back.
function rosterCartTabs_() {
  var names = [];
  try {
    var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
    var tabs = ss.getSheets();
    for (var i = 0; i < tabs.length; i++) {
      var tab = tabs[i];
      if (mirrorIsProtectedTab_(tab.getName())) continue;
      if (!rosterSerialColumn_(tab)) continue;      // not a roster tab
      names.push(tab.getName());
    }
  } catch (e) {
    Logger.log('could not list cart tabs: ' + e);
  }
  return names;
}

// Diagnostic. Lists every tab in the mirror in workbook order and says whether
// the to-do code can see it, so a tab that is missing from the dashboard can be
// told apart from one that is there but laid out differently.
//
// Run it from the editor and read View > Execution log. Sync the mirror first
// (syncMirrorNow) if the master has changed since the last refresh.
function listCartTabs() {
  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var tabs = ss.getSheets();
  var seen = 0;
  Logger.log(tabs.length + ' tab(s) in the mirror, in workbook order:');
  Logger.log('');
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i];
    var name = tab.getName();
    if (mirrorIsProtectedTab_(name)) {
      Logger.log(pad_(i + 1) + name + '  -- skipped (ticket/todo tab, not roster)');
      continue;
    }
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) {
      // Show what is actually on rows 1 and 2, so a different layout is obvious.
      var lc = Math.min(tab.getLastColumn(), 8);
      var r1 = lc ? tab.getRange(1, 1, 1, lc).getValues()[0].join(' | ') : '';
      var r2 = lc ? tab.getRange(2, 1, 1, lc).getValues()[0].join(' | ') : '';
      Logger.log(pad_(i + 1) + name + '  -- NOT DETECTED (no "serial" header on row 2)');
      Logger.log('        row1: ' + r1);
      Logger.log('        row2: ' + r2);
      continue;
    }
    seen++;
    var noteCols = rosterNoteColumns_(tab);
    var numCol = rosterNumberColumn_(tab);
    var rows = Math.max(0, tab.getLastRow() - 2);
    Logger.log(pad_(i + 1) + name + '  -- ok: serial=col' + snCol +
               ', number=col' + numCol +
               ', notes=col' + (noteCols.length ? noteCols.join('/') : 'NONE') +
               ', ' + rows + ' data row(s)' +
               (isHiddenGroup_(name) ? '   [HIDDEN on the dashboard]' : ''));
  }
  Logger.log('');
  Logger.log(seen + ' tab(s) will show as carts. Hidden: ' +
             (hiddenGroups_().join(', ') || 'none'));
  return seen;
}

function pad_(n) {
  var s = String(n);
  while (s.length < 3) s = ' ' + s;
  return s + '. ';
}

// Run by hand in the editor to rebuild without going through the dashboard.
function rebuildTodosNow() {
  var sync = syncMirror_(true);
  Logger.log('mirror sync: ' + JSON.stringify(sync));
  Logger.log(JSON.stringify(rebuildRosterTodos_()));
  rebuildTicketTodos();
}

// The one-time cross-check. Idempotent, so re-running it is harmless.
// No longer on the Refresh path -- rebuildRosterTodos_ replaced it -- but kept
// for running an import by hand.
function importRosterNotesNow() {
  Logger.log(importRosterNotes_() + ' item(s) added');
}

// Rebuild the ticket-driven items from the ticket sheet: one per open ticket,
// none for closed ones. Run by hand after a mirror reset, since the cart a
// serial belongs to is read from the mirror.
function rebuildTicketTodos() {
  var sheet = firstSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return Logger.log('no tickets');
  var v = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var open = 0;
  var closed = 0;
  for (var i = 0; i < v.length; i++) {
    var sn = v[i][0];
    var no = v[i][10];
    var status = String(v[i][8] || 'New');
    if (!no) continue;
    if (status === 'Resolved') {
      ticketTodoRemove_(no);
      closed++;
    } else {
      ticketTodoWrite_(sn, no, v[i][7]);
      open++;
    }
  }
  Logger.log('to-dos rebuilt: ' + open + ' open ticket items, ' + closed + ' closed ones cleared');
}

// ---- Roster mirror ----------------------------------------------------------
// The master roster is a .xlsx on the school account, which Apps Script cannot
// open or write. So a ticket triggers a refresh of the native mirror
// (ROSTER_SHEET_ID) from that .xlsx: convert a copy, pour the values in tab by
// tab, put the notes back, and re-lock the tabs.
//
// The mirror is DISPOSABLE. Anything typed into it by hand is wiped on the next
// refresh -- edit the master workbook instead.

var MIRROR_TRIGGER_FN = 'mirrorSyncTrigger';
var MIRROR_STAMP_KEY = 'mirrorSyncedFrom';   // master's modified time at last sync
var MIRROR_TABS_KEY = 'mirrorTabs';          // tabs the last sync wrote, JSON array

// Tabs the sync must never write to or delete, whatever the master contains.
// Ticket data, the to-do list and the monthly archives can live in the same
// workbook as the mirrored roster, and they are not the master's to overwrite.
function mirrorIsProtectedTab_(name) {
  var n = String(name || '');
  if (n === TODO_SHEET_NAME) return true;
  if (n.toLowerCase().indexOf('ticket') >= 0) return true;   // Ticketing_System, Jul26_Tickets
  try {
    if (n === firstSheet_().getName()) return true;          // the live ticket sheet
  } catch (e) {}
  return false;
}

// Copy the .xlsx into a throw-away native Sheet with the advanced Drive
// service (editor: Services + > Drive API). It runs on the drive scope this
// script already holds.
function mirrorConvertCopy_() {
  if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.copy) {
    throw new Error('Drive service missing. In the editor: Services + > Drive API.');
  }
  var name = 'TEMP roster conversion ' + new Date().getTime();
  var made;
  try {
    made = Drive.Files.copy({ name: name, mimeType: MimeType.GOOGLE_SHEETS }, MASTER_XLSX_ID);
  } catch (e) {
    // older advanced service (Drive API v2) calls the field "title"
    made = Drive.Files.copy({ title: name, mimeType: MimeType.GOOGLE_SHEETS }, MASTER_XLSX_ID);
  }
  if (!made || !made.id) throw new Error('Drive.Files.copy returned no id');
  return made.id;
}

// Values-only copy of every tab, source -> mirror. Formatting is not carried
// over; the roster is data, and this keeps the run to a few seconds.
function mirrorCopyTabs_(fromSs, toSs, pruneAll) {
  var props = PropertiesService.getScriptProperties();
  var previous = [];
  try { previous = JSON.parse(props.getProperty(MIRROR_TABS_KEY) || '[]'); } catch (e) { previous = []; }

  var written = [];
  var names = {};
  var from = fromSs.getSheets();
  for (var i = 0; i < from.length; i++) {
    var src = from[i];
    var name = src.getName();
    if (mirrorIsProtectedTab_(name)) {
      Logger.log('mirror: leaving "' + name + '" alone (protected)');
      continue;
    }
    names[name] = true;
    var rows = src.getLastRow();
    var cols = src.getLastColumn();
    var dest = toSs.getSheetByName(name);
    if (!dest) dest = toSs.insertSheet(name);
    dest.clear();
    if (rows > 0 && cols > 0) {
      dest.getRange(1, 1, rows, cols).setValues(src.getRange(1, 1, rows, cols).getValues());
    }
    written.push(name);
  }

  // Normally only tabs a previous sync created are removed -- anything else was
  // put here by somebody, not by the mirror. A reset prunes everything that is
  // not in the master, which is how leftovers from an older master go away.
  var candidates = previous;
  if (pruneAll) {
    candidates = [];
    var have = toSs.getSheets();
    for (var h = 0; h < have.length; h++) candidates.push(have[h].getName());
  }
  for (var k = 0; k < candidates.length; k++) {
    var old = candidates[k];
    if (names[old] || mirrorIsProtectedTab_(old)) continue;
    var sh = toSs.getSheetByName(old);
    if (sh && toSs.getSheets().length > 1) {
      var prot = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var q = 0; q < prot.length; q++) prot[q].remove();
      toSs.deleteSheet(sh);
      Logger.log('mirror: removed "' + old + '"');
    }
  }
  props.setProperty(MIRROR_TABS_KEY, JSON.stringify(written));
  return written.length;
}

// Lock every tab so only the account running this script can change it. The
// script itself still writes notes, because it runs as the owner.
function mirrorProtectTabs_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (mirrorIsProtectedTab_(sheets[i].getName())) continue;   // not ours to lock
    try {
      var old = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var j = 0; j < old.length; j++) old[j].remove();
      var p = sheets[i].protect().setDescription('Mirror of the master workbook - edit the master, not this');
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
    } catch (e) {
      Logger.log('could not protect ' + sheets[i].getName() + ': ' + e);
    }
  }
}

// A refresh wipes the notes, so write them back for every ticket that is not
// closed yet.
function mirrorReapplyNotes_() {
  var sheet = firstSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var v = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var n = 0;
  for (var i = 0; i < v.length; i++) {
    var status = String(v[i][8] || 'New');
    if (status === 'Resolved') continue;
    var r = rosterNoteWrite_(v[i][0], v[i][10], v[i][7]);
    if (r && r.ok && !r.skipped) n++;
  }
  return n;
}

// Refresh the mirror from the master. Skips when the master has not changed
// since the last run; pass true to force it.
function syncMirror_(force, pruneAll) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { ok: true, skipped: 'another sync is running' };
  var tempId = null;
  try {
    var master = DriveApp.getFileById(MASTER_XLSX_ID);
    var stamp = master.getLastUpdated().toISOString();
    var props = PropertiesService.getScriptProperties();
    if (!force && props.getProperty(MIRROR_STAMP_KEY) === stamp) {
      return { ok: true, skipped: 'master unchanged' };
    }
    tempId = mirrorConvertCopy_();
    var temp = SpreadsheetApp.openById(tempId);
    var mirror = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    var copied = mirrorCopyTabs_(temp, mirror, pruneAll);
    SpreadsheetApp.flush();
    var notes = mirrorReapplyNotes_();
    mirrorProtectTabs_(mirror);
    props.setProperty(MIRROR_STAMP_KEY, stamp);
    Logger.log('mirror refreshed from master; tabs copied: ' + copied + ', notes restored: ' + notes);
    return { ok: true, tabs: copied, notes: notes };
  } catch (e) {
    Logger.log('mirror sync failed: ' + e);
    return { ok: false, error: String(e) };
  } finally {
    if (tempId) {
      try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e2) {}
    }
    lock.releaseLock();
  }
}

// doPost calls this: books the sync for a minute from now so the teacher's form
// returns straight away. One pending trigger at a time.
function scheduleMirrorSync_() {
  try {
    var all = ScriptApp.getProjectTriggers();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getHandlerFunction() === MIRROR_TRIGGER_FN) return;   // already booked
    }
    ScriptApp.newTrigger(MIRROR_TRIGGER_FN).timeBased().after(60 * 1000).create();
  } catch (e) {
    Logger.log('could not schedule mirror sync: ' + e);
  }
}

// The scheduled handler: clears itself, then syncs.
function mirrorSyncTrigger() {
  try {
    var all = ScriptApp.getProjectTriggers();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getHandlerFunction() === MIRROR_TRIGGER_FN) ScriptApp.deleteTrigger(all[i]);
    }
  } catch (e) {}
  syncMirror_(false);
}

// Run by hand to refresh now, whether or not the master changed.
function syncMirrorNow() {
  Logger.log(syncMirror_(true));
}

// ---- Scheduled refresh ------------------------------------------------------
// Four times a school day: pull the master across, pick up any notes typed into
// the carts since, and square the to-do list with the open tickets.
var TODO_REFRESH_FN = 'scheduledTodoRefresh';
var TODO_REFRESH_HOURS = [7, 9, 12, 15];        // school time zone, see setupTodoTriggers

// Not forced here: when the master has not changed the mirror is already
// current, and this runs unattended four times a day.
function scheduledTodoRefresh() {
  var sync = syncMirror_(false);
  var res = rebuildRosterTodos_();
  rebuildTicketTodos();
  Logger.log('scheduled refresh done. sync: ' + JSON.stringify(sync) +
             ', roster items: ' + res.imported + ', new: ' + res.added +
             ', gone: ' + res.removed + ', kept: ' + res.kept);
}

// Run once to install the four daily triggers. Running it again just replaces
// them, so it is safe to repeat.
//
// The hours follow the SCRIPT PROJECT's time zone, not your computer's: check
// Project Settings shows America/New_York, or the runs will land at the wrong
// local time. Apps Script fires within about 15 minutes of the hour.
function setupTodoTriggers() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === TODO_REFRESH_FN) ScriptApp.deleteTrigger(all[i]);
  }
  for (var h = 0; h < TODO_REFRESH_HOURS.length; h++) {
    ScriptApp.newTrigger(TODO_REFRESH_FN)
      .timeBased().everyDays(1).atHour(TODO_REFRESH_HOURS[h]).nearMinute(0).create();
  }
  Logger.log('installed ' + TODO_REFRESH_HOURS.length + ' daily triggers at ' +
             TODO_REFRESH_HOURS.join(', ') + ' (project time zone: ' +
             Session.getScriptTimeZone() + ')');
}

// One-off cleanup. The mirror still carries tabs from an older master, and
// Lookup searches those too, so stale serials keep resolving. This empties the
// mirror completely and rebuilds it from the current master.
//
// Only the mirror is touched -- tickets, todos and archives are in a different
// workbook (CPA_Chromebook_Ticketing). If anything goes wrong, the mirror's
// File > Version history has the previous state.
function resetMirror() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(MIRROR_TABS_KEY);
  props.deleteProperty(MIRROR_STAMP_KEY);

  // The copy happens first and the pruning second, inside the same run, so a
  // failure part way through leaves the old mirror in place rather than an
  // empty workbook.
  var res = syncMirror_(true, true);
  Logger.log(res);
  if (!res.ok) {
    Logger.log('reset failed, mirror left as it was. Nothing was deleted.');
  }
  return res;
}

// First thing to run when something is not working. Says which version of this
// file is loaded, who it runs as, and whether the Drive service is available.
function checkSetup() {
  Logger.log('script version:  ' + SCRIPT_VERSION);
  try {
    // owner of the mirror -- the account this script must be running as for the
    // notes and the tab locks to work
    Logger.log('mirror owner:    ' + DriveApp.getFileById(ROSTER_SHEET_ID).getOwner().getEmail());
  } catch (e) {
    Logger.log('mirror owner:    (unknown) ' + e);
  }
  var hasDrive = (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.copy);
  Logger.log('Drive service:   ' + (hasDrive ? 'ADDED - conversion will use it' :
             'MISSING - add it with Services + > Drive API'));
  try {
    var f = DriveApp.getFileById(MASTER_XLSX_ID);
    Logger.log('master file:     ' + f.getName() + ' (' + f.getMimeType() + ')');
  } catch (e) {
    Logger.log('master file:     CANNOT READ - ' + e);
  }
  try {
    var m = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    Logger.log('mirror sheet:    ' + m.getName() + ', ' + m.getSheets().length + ' tabs');
  } catch (e) {
    Logger.log('mirror sheet:    CANNOT OPEN - ' + e);
  }
}

// Run this by hand against a real serial to see where a note would land.
function testRosterNote() {
  var sn = 'YX0JK8EFYXN0B6205009';       // a serial from the roster workbook
  Logger.log(rosterNoteWrite_(sn, 9999, 'TEST - screen flickers'));
  Logger.log(rosterNoteClear_(sn, 9999, 'TEST - screen flickers'));
}

// ---- To-Do list (dashboard "To-Do" tab) ----
// Items live in a 'Todos' sheet tab: ID | Text | Done | Order | Created | Group.
// Group is the cart/section the task belongs to (e.g. "Cart O"); the dashboard
// shows each group as a collapsible section. Blank group shows as "General".
var TODO_SHEET_NAME = 'Todos';
var TODO_HEADERS = ['ID', 'Text', 'Done', 'Order', 'Created', 'Group', 'Status'];

// How far along a task is. Everything starts at the first one.
var TODO_STATUSES = ['Untouched', 'Noticed', 'Working', 'Stalled', 'Progressing', 'Complete'];
var TODO_STATUS_DEFAULT = 'Untouched';

function todoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TODO_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(TODO_SHEET_NAME, ss.getNumSheets());
    sh.getRange(1, 1, 1, TODO_HEADERS.length).setValues([TODO_HEADERS]);
    sh.getRange(1, 1, 1, TODO_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // Older versions of this sheet had no Group column - add the header if missing.
  if (sh.getRange(1, 6).getValue() !== 'Group') {
    sh.getRange(1, 6).setValue('Group').setFontWeight('bold');
  }
  // ...and no Status column.
  if (sh.getRange(1, 7).getValue() !== 'Status') {
    sh.getRange(1, 7).setValue('Status').setFontWeight('bold');
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
    var status = String(r[6] || '').trim();
    if (TODO_STATUSES.indexOf(status) < 0) status = TODO_STATUS_DEFAULT;
    todos.push({
      id: String(r[0]), text: String(r[1]), done: done,
      order: Number(r[3]) || 0, group: String(r[5] || ''), status: status
    });
  });
  todos.sort(function (a, b) { return a.order - b.order; });
  return { ok: true, todos: todos, hidden: hiddenGroups_(), carts: rosterCartTabs_() };
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
  sh.appendRow([id, text, false, order, new Date(), String(p.group || '').trim(),
                TODO_STATUS_DEFAULT]);
  return { ok: true, id: id };
}

function todoUpdate_(p) {
  var sh = todoSheet_();
  var row = todoFindRow_(sh, p.id);
  if (!row) return { ok: false, error: 'not found' };
  if (p.text != null) sh.getRange(row, 2).setValue(String(p.text));
  if (p.done != null) sh.getRange(row, 3).setValue(String(p.done) === 'true');
  if (p.group != null) sh.getRange(row, 6).setValue(String(p.group).trim());
  if (p.status != null) {
    var st = String(p.status).trim();
    if (TODO_STATUSES.indexOf(st) < 0) st = TODO_STATUS_DEFAULT;
    sh.getRange(row, 7).setValue(st);
  }
  return { ok: true };
}

// Deleting an item that came from a roster note counts as dealing with it:
// the note is remembered so the next import does not bring the item straight
// back, and it turns up in the "notes to clear" report instead. Items from a
// ticket are not suppressed -- an open ticket belongs on the list, and its item
// returns until the ticket is marked Resolved.
function todoDelete_(p) {
  var sh = todoSheet_();
  var row = todoFindRow_(sh, p.id);
  if (!row) return { ok: false, error: 'not found' };
  var id = String(p.id);
  var suppressed = false;
  if (id.indexOf('roster-') === 0) {
    try {
      var todo = {
        id: id,
        text: String(sh.getRange(row, 2).getValue() || ''),
        group: String(sh.getRange(row, 6).getValue() || '')
      };
      var src = rosterTodoSource_(todo);
      if (src) {
        var cur = String(src.sheet.getRange(src.row, src.col).getValue() || '').trim();
        if (cur) {
          markNoteHandled_(src.sn, cur);
          logClosed_(id, cur, src.sn, todo.group, 'removed from the list');
          suppressed = true;
        }
      }
    } catch (e) {
      Logger.log('could not suppress note for ' + id + ': ' + e);
    }
  }
  sh.deleteRow(row);
  return { ok: true, suppressed: suppressed };
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

// ---- Hidden carts -----------------------------------------------------------
// Renaming a tab in the master leaves the to-do page showing the same physical
// cart twice, once under each name, and the HS_Cart_* tabs are not wanted on
// the page at all. Hiding a cart takes it off the grid and stops that tab being
// imported.
//
// Nothing is deleted. The name goes on a list, rebuildRosterTodos_ skips that
// tab, and the cart's rows simply stop being re-created on the next refresh.
// Rows typed by hand are never rebuilt from anything, so they sit untouched in
// the sheet and come back if the cart is restored.
//
// The list has to live on the server: hiding a cart in the browser alone would
// not hold, because the scheduled refresh reads the master four times a day and
// would put every item straight back.
//
// Tabs the master gains later are picked up as usual -- only names on this list
// are skipped.
var TODO_HIDDEN_KEY = 'hiddenTodoGroups';

function hiddenGroups_() {
  try {
    var v = JSON.parse(PropertiesService.getScriptProperties().getProperty(TODO_HIDDEN_KEY) || '[]');
    return Object.prototype.toString.call(v) === '[object Array]' ? v : [];
  } catch (e) {
    return [];
  }
}

function isHiddenGroup_(name) {
  var list = hiddenGroups_();
  var n = String(name || '').trim().toLowerCase();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i]).trim().toLowerCase() === n) return true;
  }
  return false;
}

// Remember the name so the tab is not imported again. Writes nothing to the
// Todos sheet -- the cart's rows stop being re-created by the next rebuild, and
// anything added from the dashboard stays put.
function todoHideGroup_(p) {
  var name = String(p.group || '').trim();
  if (!name) return { ok: false, error: 'no group given' };
  var list = hiddenGroups_();
  if (!isHiddenGroup_(name)) list.push(name);
  PropertiesService.getScriptProperties().setProperty(TODO_HIDDEN_KEY, JSON.stringify(list));
  Logger.log('cart hidden: ' + name);
  return { ok: true, group: name, hidden: hiddenGroups_() };
}

// Put a cart back. Its notes return on the next refresh.
function todoShowGroup_(p) {
  var name = String(p.group || '').trim();
  if (!name) return { ok: false, error: 'no group given' };
  var list = hiddenGroups_();
  var out = [];
  for (var i = 0; i < list.length; i++) {
    if (String(list[i]).trim().toLowerCase() !== name.toLowerCase()) out.push(list[i]);
  }
  PropertiesService.getScriptProperties().setProperty(TODO_HIDDEN_KEY, JSON.stringify(out));
  Logger.log('cart restored: ' + name);
  return { ok: true, group: name, hidden: out };
}

// Run from the editor if a cart needs bringing back and the dashboard is not to
// hand.
function clearHiddenGroups() {
  PropertiesService.getScriptProperties().deleteProperty(TODO_HIDDEN_KEY);
  Logger.log('all hidden carts restored');
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

    // A factory QR code decodes to a support URL, not a serial. Whatever the
    // form sent, store the serial. Done once here, so the row that is written,
    // the photo file name, the email subject, the roster note and the to-do
    // item all get the same clean value.
    data.sn = serialFromScan_(data.sn);

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
    // refresh the device's row in the roster mirror with what was typed into
    // "Describe the problem"; skipped silently if the serial is not on a tab
    var note = rosterNoteWrite_(data.sn, no, data.description);
    // put the same text on the to-do list, under the cart the device is in
    var where = null;
    if (note && note.tab) where = { cart: note.tab, cbNo: note.cbNo };
    ticketTodoWrite_(data.sn, no, data.description, where);
    // and pull any roster changes across from the master, a minute from now, so
    // the teacher's form is not left waiting on it
    scheduleMirrorSync_();
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

// ---- Serial normalising -----------------------------------------------------
// Factory QR labels encode a support URL rather than the bare serial. A new
// Lenovo reads as https://support.lenovo.com/qrcode?sn=YX0JK6SE&mtm=83T60009US
// -- the serial is the sn parameter, and mtm is the model, not a serial. Our
// own printed labels encode the bare serial and pass through unchanged.
//
// The scanner pages do this too, so the teacher sees the right thing while
// filling the form. It is repeated here because the server is the one place
// every route passes through: a pasted URL, or a page still running the old
// build out of the Pages cache, would otherwise put a link in the sheet.
// serialFromScan() in the three HTML files is the same logic -- keep them in
// step if either changes.
function serialFromScan_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return '';

  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(s)) {
    // A URL. Take the first query parameter that names a serial, so sn wins
    // over mtm on a Lenovo link.
    var q = s.indexOf('?');
    if (q >= 0) {
      var parts = s.slice(q + 1).split(/[&;]/);
      for (var i = 0; i < parts.length; i++) {
        var eq = parts[i].indexOf('=');
        if (eq < 0) continue;
        var key = parts[i].slice(0, eq).toLowerCase().replace(/[^a-z]/g, '');
        if (key !== 'sn' && key !== 'serial' && key !== 'serialno' &&
            key !== 'serialnumber' && key !== 'servicetag') continue;
        var val = '';
        try {
          val = decodeURIComponent(parts[i].slice(eq + 1).replace(/\+/g, ' ')).trim();
        } catch (e) {
          val = parts[i].slice(eq + 1).trim();
        }
        if (val) return val;
      }
    }
    // No serial parameter -- some makers put it in the last path segment.
    var path = s.split(/[?#]/)[0].replace(/\/+$/, '');
    var seg = path.slice(path.lastIndexOf('/') + 1);
    if (looksLikeRosterSerial_(seg)) return seg;
    return s;             // nothing recognisable; keep what we read rather than
                          // mangle it into something wrong
  }

  // Not a URL. Some labels write "S/N: ABC123" rather than the bare value.
  var m = s.match(/\b(?:s\/n|sn|serial(?:\s*(?:no|number))?|service\s*tag)\s*[:=]\s*([A-Za-z0-9-]+)/i);
  if (m && m[1]) return m[1];

  return s;
}

// ---- One-off cleanup of what is already stored ------------------------------
// Tickets filed before the parsing went in have a URL sitting in column A.
// Walks the live ticket sheet and every archive tab. listBadSerials reports and
// writes nothing; fixStoredSerials applies.

function listBadSerials() {
  var found = badSerialRows_();
  if (!found.length) {
    Logger.log('nothing to fix -- every stored serial is already clean');
    return 0;
  }
  Logger.log(found.length + ' row(s) hold a serial that needs rewriting:');
  for (var i = 0; i < found.length; i++) {
    Logger.log('  ' + found[i].tab + ' row ' + found[i].row + ':  ' +
               found[i].was + '   ->   ' + found[i].now);
  }
  Logger.log('run fixStoredSerials() to apply these');
  return found.length;
}

function fixStoredSerials() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var found = badSerialRows_();
  for (var i = 0; i < found.length; i++) {
    ss.getSheetByName(found[i].tab).getRange(found[i].row, 1).setValue(found[i].now);
    Logger.log('fixed ' + found[i].tab + ' row ' + found[i].row + ': ' +
               found[i].was + ' -> ' + found[i].now);
  }
  Logger.log(found.length + ' serial(s) rewritten');
  return found.length;
}

function badSerialRows_() {
  var out = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var live = firstSheet_().getName();
  var tabs = ss.getSheets();
  for (var t = 0; t < tabs.length; t++) {
    var sh = tabs[t];
    var name = sh.getName();
    // Ticket data only: the live sheet and the monthly archives. Not the Todos
    // tab, not anything mirrored.
    if (name !== live && name.toLowerCase().indexOf('ticket') < 0) continue;
    var lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    var v = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < v.length; r++) {
      var was = String(v[r][0] || '').trim();
      if (!was) continue;
      var now = serialFromScan_(was);
      if (now && now !== was) out.push({ tab: name, row: r + 2, was: was, now: now });
    }
  }
  return out;
}
