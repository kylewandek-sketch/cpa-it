var SCRIPT_VERSION = '2026-08-12c roster-matched serials';   // shown by checkSetup()

var HELPDESK_EMAIL = 'kyle.anderson@cpaohio.org';
var ADMIN_TOKEN = 'CHANGE_ME';   // set your own; do NOT commit the real token to a public repo

// Drive folder that ticket photos are saved into. The account running this script
// must have EDIT access to it. Falls back to a folder on the script's own Drive.
var PHOTO_FOLDER_ID = '1rhGgLWy4W9i1U-EcfcAf9ZBbzuNWotCb';   // "Ticket Images"
var PHOTO_FOLDER_FALLBACK = 'CPA IT Ticket Photos';

// The roster workbook: cart tabs, the iPad list, everything the to-do page is
// built from, and the book teachers type their notes into. This IS the master --
// there is no mirror any more.
//
// It used to be a .xlsx on the school account, which SpreadsheetApp cannot open
// at all, so the script kept a converted copy and synced it. Converting the
// workbook to a native Sheet on 2026-08-05 removed the need for that: notes now
// read and write straight through, and a refresh no longer pays for a Drive
// copy of the whole book.
//
// Superseded ids, kept for the record:
//   1z1W54tWIvm4XlaqsEbyN4slO2HNCD4FO  the original .xlsx (now archived)
//   1FDVE6KtAEf06_zRYQyHyaNZ_9gXsv3JRJGbIwckv4Mw  the old mirror
var ROSTER_SHEET_ID = '1WLrGRmlRoaFeg2OrkwP8eXVilkr-LiQRP_RrQrf4d0o';

// The workbook holding tickets, the Todos tab, the ClosedLog and the monthly
// archives, in the shared IT folder.
//
// Reached by id, never through getActiveSpreadsheet(). That is what lets this
// script be a STANDALONE project rather than one bound inside a spreadsheet --
// a bound script is part of its container file and cannot be moved out of it,
// so the container could never be deleted.
//
// Script properties (lastTicketNo, hiddenTodoGroups, handledRosterNotes,
// cartTabsCache) belong to the SCRIPT PROJECT, not to any workbook. They do not
// follow the data, and they do not survive being copied into a new project --
// see the migration notes in the repo if this project is ever recreated.
var TICKET_BOOK_ID = '12CcT0FHlILSGqABfpsK9sRAs2D2YnWZmHR4RYe8dUZA';

function ticketBook_() { return SpreadsheetApp.openById(TICKET_BOOK_ID); }

// The assignment roster: "2026-2027 Chromebook Carts/Ipads". Same cart tabs as
// the check workbook above, but this is the book that says who a device belongs
// to -- row 1 is "TEACHER: <name>" / "ROOM# <n>", row 2 is the headers, and
// column C is the assigned student.
//
// The dashboard's Lookup tab reads its assignment section from HERE. Notes, the
// to-do list and everything else still come from ROSTER_SHEET_ID; this book has
// no note columns.
var ASSIGNMENT_SHEET_ID = '1JQxgBqWzrwg58okUJT39L1xv_MbmoPLNPdh31IjD1DQ';

// ---- Roster note settings ----
// When a ticket is opened or moved to In Progress, the teacher's "Describe the
// problem" text is written into the device's row, so the roster itself shows why
// a device is out. Cleared when the ticket is Resolved. This now lands in the
// workbook people actually open.
var NOTE_BOOK_ID = ROSTER_SHEET_ID;

// Cart for a to-do item when the serial is not on any roster tab.
var TODO_GROUP_FALLBACK = 'Unassigned';

// Tabs in the master that are not carts and should never reach the to-do page:
// the per-teacher lists and the separate iPad cart. Matched on the whole name,
// ignoring case and surrounding spaces. This is for tabs that are structurally
// not carts -- for a one-off choice, use the remove control on the dashboard,
// which writes to hiddenTodoGroups instead.
// 'Spares' is here so the loaner pool's condition notes do not pour into the
// cart to-do list. Remove it from this list if you decide you want them there.
var TODO_SKIP_TABS = [
  'Crider', 'Palsa', 'Buechner', 'Aeh', 'Jablonski', 'Miller', 'Perez',
  'Caudill', 'Moorman', 'Title 1', 'iPad Cart - Hunter', 'Spares'
];

function isSkippedTab_(name) {
  var n = String(name || '').trim().toLowerCase();
  for (var i = 0; i < TODO_SKIP_TABS.length; i++) {
    if (TODO_SKIP_TABS[i].toLowerCase() === n) return true;
  }
  return false;
}

// Everything the to-do page should leave alone: ticket/todo tabs, the skip list
// above, and carts removed from the dashboard.
// Tabs in the roster workbook that are not carts at all. The ticket data and the
// to-do list live in a different workbook now, but the check is cheap and stops
// anything odd being read as a roster tab.
// The Loaner_log tabs in the assignment and check workbooks. Since 2026-08-11
// they are not data at all: A1 holds
//   =IMPORTRANGE("...12CcT0FH...", "Loaners!A:H")
// so each one is a read-only mirror of the Loaners ledger in the ticket book.
//
// This tab MUST be excluded, and the reason is not cosmetic. The mirror spills
// the ledger's headers onto row 1, where rosterHeadInfo_ is looking:
//
//   "Inop S/N"   matches /serial|s\/n/  -> the tab reads as a roster tab
//   "Note"       matches NOTE_HEADER_RE -> its column reads as a note source
//
// Left unguarded that means every serial in the ledger counted a second time in
// both books and reported as filed in two places, plus one to-do item per
// ledger row, in a group named after this tab. The old layout escaped only
// because its header row sat below row 3; the mirror does not.
//
// Matched loosely on purpose -- "Loaner_log", "Loaner_logs", "Loaner Log" all
// count. An exact name is one typo away from letting all of that back in.
var RELOCATION_LOG_RE = /loaner[\s_-]*log/i;

function isNonRosterTab_(name) {
  var n = String(name || '');
  if (n === TODO_SHEET_NAME) return true;
  if (RELOCATION_LOG_RE.test(n)) return true;
  return n.toLowerCase().indexOf('ticket') >= 0;
}

function todoIgnoreTab_(name) {
  return isNonRosterTab_(name) || isSkippedTab_(name) || isHiddenGroup_(name);
}

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
                 'todoHideGroup', 'todoShowGroup', 'refreshTodos', 'validateTodos',
                 'validateData', 'validateLast', 'validateIgnore', 'validateFixSerials',
                 'loanerList', 'loanerOptions', 'loanerIssue', 'loanerReturn', 'loanerResync',
                 'gridTabs', 'gridRead', 'gridSearch', 'gridWrite', 'gridUndo', 'gridBatches'];
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
  } else if (p.action === 'validateTodos') {
    out = validateTodos_(p);
  } else if (p.action === 'validateData') {
    out = validateData_(p);       // full cross-book walk, ~60-90s, button only
  } else if (p.action === 'validateLast') {
    out = validateLast_();        // the cached report, costs nothing
  } else if (p.action === 'validateIgnore') {
    out = validateIgnore_(p);
  } else if (p.action === 'validateFixSerials') {
    out = validateFixSerials_();
  } else if (p.action === 'loanerList') {
    out = loanerList_();
  } else if (p.action === 'loanerOptions') {
    out = loanerOptions_();
  } else if (p.action === 'loanerIssue') {
    out = loanerIssue_(p);
  } else if (p.action === 'loanerReturn') {
    out = loanerReturn_(p);
  } else if (p.action === 'loanerResync') {
    out = loanerResync_();
  } else if (p.action === 'gridTabs') {
    out = gridTabs_(p);
  } else if (p.action === 'gridRead') {
    out = gridRead_(p);
  } else if (p.action === 'gridSearch') {
    out = gridSearch_(p);
  } else if (p.action === 'gridWrite') {
    out = gridWrite_(p);
  } else if (p.action === 'gridUndo') {
    out = gridUndo_(p);
  } else if (p.action === 'gridBatches') {
    out = gridBatches_();
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
  } else if (p.action === 'snResolve') {
    out = snResolve_(p);          // public: scanner asks the books which scanned code is the device
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

// The ticket sheet, found by its header rather than by position.
//
// This was getSheets()[0], so dragging any tab in front of it would silently
// send tickets to the wrong sheet -- and it has already been renamed once, from
// Ticketing_System to Don't_Touch, which shows the name is not dependable
// either. A1 holding "Chromebook S/N" is the thing that actually identifies it.
//
// Cached per execution: this is called on every ticket read and write.
var TICKET_SHEET_CACHE = null;

function firstSheet_() {
  if (TICKET_SHEET_CACHE) return TICKET_SHEET_CACHE;
  var sheets = ticketBook_().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getRange(1, 1).getValue() || '').trim() === HEADERS[0]) {
      TICKET_SHEET_CACHE = sheets[i];
      return TICKET_SHEET_CACHE;
    }
  }
  TICKET_SHEET_CACHE = sheets[0];        // nothing matched: behave as it always did
  Logger.log('WARNING: no tab has "' + HEADERS[0] + '" in A1; falling back to "' +
             TICKET_SHEET_CACHE.getName() + '" as the ticket sheet.');
  return TICKET_SHEET_CACHE;
}

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
    // Same matcher the scanner uses, so a serial that IS in the books under a
    // slightly longer form stops being reported as missing at submit time.
    var hit = rosterResolveOne_(sn);
    if (!hit || hit.ambiguous) return { ok: true, found: false };
    return { ok: true, found: true, sn: hit.sn, how: hit.how, cart: rosterCartFor_(hit.sn) };
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
  var ss = ticketBook_();
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
// Row 1 of an assignment tab holds "TEACHER: Messer" in A1 and "ROOM# 13" in B1.
// The dashboard already prints its own "Teacher:" and "Room:" labels, so strip
// the one baked into the cell. A tab with a bare name in A1 is left alone.
function stripRosterLabel_(value) {
  var s = String(value || '').trim();
  s = s.replace(/^teacher\s*#?\s*:?\s*/i, '');
  s = s.replace(/^room\s*#?\s*:?\s*/i, '');
  return s.trim();
}

function deviceLookup_(p) {
  var sn = serialFromScan_(p.sn);
  if (!sn) return { ok: false, error: 'No serial provided.' };
  var out = { ok: true, sn: sn, assignments: [], tickets: [], todos: [] };

  // 1) Roster assignment — read from the ASSIGNMENT workbook, not the check
  //    workbook: that is the book that says who a device belongs to. Serials
  //    live in column B of tabs whose B2 says "Serial #".
  try {
    var rs = SpreadsheetApp.openById(ASSIGNMENT_SHEET_ID);
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
        teacher: stripRosterLabel_(heads[0][0]),
        room: stripRosterLabel_(heads[0][1]),
        chromebookNo: String(vals[0] || ''),
        students: students
      });
    });
  } catch (e) { out.rosterError = String(e); }

  // 2) Ticket history — S/N is column A in the live sheet and every archive tab.
  try {
    var ss = ticketBook_();
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
    var ts = ticketBook_().getSheetByName(TODO_SHEET_NAME);
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
  var heads = rosterHeads_(sh);
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
  var heads = rosterHeads_(sh);
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
  var heads = rosterHeads_(sh);
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
  var heads = rosterHeads_(sh);
  var cols = [];
  for (var c = 3; c <= lastCol; c++) {
    var h = String(heads[c - 1] || '').trim();
    if (h.toLowerCase().indexOf('student') >= 0) continue;
    if (!h) { cols.push(c); continue; }
    if (NOTE_HEADER_RE.test(h)) cols.push(c);
  }
  return cols;
}

// Cart tabs put their headers on row 2 and data from row 3, but not every tab
// in the master is laid out that way -- the iPad list is not. Find the row the
// headers are actually on by looking for something that names a serial in the
// first three rows. Falls back to 2, which is what the cart tabs use.
//
// Everything that reads a roster tab goes through this, so a tab with headers
// on row 1 is read correctly rather than being skipped as "not a roster tab".
var ROSTER_SERIAL_RE = /serial|s\/n/i;

// Rows 1-3 of a tab, read ONCE and remembered for the rest of the run.
//
// This matters more than it looks. rosterSerialColumn_, rosterNoteColumns_ and
// rosterNumberColumn_ all want the headers, and each is called per tab; without
// the cache every one of them is a separate round trip to the sheet, times ~50
// tabs, and the dashboard's Refresh times out before the script finishes.
var ROSTER_HEAD_CACHE = {};

function rosterHeadReset_() { ROSTER_HEAD_CACHE = {}; }

function rosterHeadInfo_(sh) {
  var key = sh.getSheetId() + '|' + sh.getName();
  var hit = ROSTER_HEAD_CACHE[key];
  if (hit) return hit;

  var info = { row: 2, heads: [], serialCol: 0 };
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastCol && lastRow) {
    var probe = Math.min(6, lastRow);
    var vals = sh.getRange(1, 1, probe, lastCol).getValues();   // the one read

    // 1) A header row naming a serial, in the first three rows. Cart tabs put
    //    it on row 2; the iPad list uses row 2 as well but with its own columns.
    var top = Math.min(3, probe);
    for (var r = 0; r < top && !info.serialCol; r++) {
      for (var c = 0; c < lastCol; c++) {
        if (ROSTER_SERIAL_RE.test(String(vals[r][c] || ''))) {
          info.row = r + 1;
          info.serialCol = c + 1;
          info.heads = vals[r];
          break;
        }
      }
    }

    // 2) No headers at all -- some tabs (Speech) are a bare list, serial in the
    //    first column and data from row 1. Find the column that actually holds
    //    serials. row 0 means "no header row", so data starts at row 1.
    if (!info.serialCol) {
      for (var c2 = 0; c2 < lastCol && !info.serialCol; c2++) {
        for (var r2 = 0; r2 < probe; r2++) {
          if (looksLikeRosterSerial_(vals[r2][c2])) {
            info.row = 0;
            info.serialCol = c2 + 1;
            info.heads = [];
            break;
          }
        }
      }
    }
  }
  ROSTER_HEAD_CACHE[key] = info;
  return info;
}

function rosterHeaderRow_(sh) { return rosterHeadInfo_(sh).row; }
function rosterHeads_(sh) { return rosterHeadInfo_(sh).heads; }
function rosterDataRow_(sh) { return rosterHeadInfo_(sh).row + 1; }

// Tab order is simply the workbook's order now -- getSheets() returns them as
// they sit in the tab strip. The old version had to reconstruct it from a
// property the sync wrote, because the mirror's own order drifted.
function sortByMasterOrder_(names) { return names; }

// Which column on this tab holds the serials.
function rosterSerialColumn_(sh) { return rosterHeadInfo_(sh).serialCol; }

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

// Walk every roster tab and put any note found onto the to-do
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
    if (todoIgnoreTab_(tab.getName())) continue;
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;                       // not a roster tab
    var noteCols = rosterNoteColumns_(tab);
    if (!noteCols.length) continue;
    var numCol = rosterNumberColumn_(tab);
    var lastRow = tab.getLastRow();
    var dataRow = rosterDataRow_(tab);
    if (lastRow < dataRow) continue;
    var lastCol = tab.getLastColumn();
    var heads2 = rosterHeads_(tab);
    var vals = tab.getRange(dataRow, 1, lastRow - dataRow + 1, lastCol).getValues();
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

// A note that has been dealt with. Kept on a list so it is not imported again
// while it is still typed in the workbook; editing the note changes its key, so
// a genuinely new note still comes through.
//
// This exists because the roster used to be an .xlsx the script could not write.
// It can now clear the cell directly, which would make this list and the
// "notes to clear" report unnecessary -- left in place for now so the switch to
// the native workbook is one change rather than two.
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
  var ss = ticketBook_();
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
  var m = String(todo.id || '').match(ROSTER_ID_RE);
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
    var dataRow = rosterDataRow_(tab);
    if (lastRow < dataRow) return null;
    var col_vals = tab.getRange(dataRow, snCol, lastRow - dataRow + 1, 1).getValues();
    for (var i = 0; i < col_vals.length; i++) {
      if (String(col_vals[i][0]).trim() === sn) {
        return { sheet: tab, row: i + dataRow, col: col, sn: sn };
      }
    }
  } catch (e) {
    Logger.log('could not locate source of ' + todo.id + ': ' + e);
  }
  return null;
}

// Square the roster-sourced items with the roster itself:
//   ticked off  -> clear the cell, remember it, drop the item
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
    if (todoIgnoreTab_(tab.getName())) continue;
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;
    var noteCols = rosterNoteColumns_(tab);
    if (!noteCols.length) continue;
    var lastRow = tab.getLastRow();
    var dataRow = rosterDataRow_(tab);
    if (lastRow < dataRow) continue;
    var vals = tab.getRange(dataRow, 1, lastRow - dataRow + 1, tab.getLastColumn()).getValues();
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
          cell: tab.getRange(r + dataRow, noteCols[k]).getA1Notation(),
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
// Timed per stage, so a slow one is visible rather than guessed at. There is no
// mirror sync to pay for any more -- this reads the roster workbook directly.
function refreshTodos_() {
  var t0 = new Date().getTime();
  rosterHeadReset_();
  var res = rebuildRosterTodos_();
  var tBuild = new Date().getTime();

  rebuildTicketTodos();
  var carts = cartTabsRefresh_();
  // Mark the list current, so the next dashboard load sees nothing to do.
  var stamp = rosterStamp_();
  if (stamp) PropertiesService.getScriptProperties().setProperty(ROSTER_STAMP_KEY, stamp);
  var list = todoList_();
  var tEnd = new Date().getTime();

  var timing = { rebuild: tBuild - t0, tickets_carts_list: tEnd - tBuild, total: tEnd - t0 };
  Logger.log('refresh timing (ms): ' + JSON.stringify(timing));

  return { ok: true, added: res.added, dropped: res.removed,
           imported: res.imported, kept: res.kept,
           todos: list.todos, hidden: list.hidden,
           carts: carts, timing: timing };
}

// ---- Validate on load -------------------------------------------------------
// The dashboard asks for this every time it opens. A full rebuild takes ~50s,
// which is far too slow to do on every visit -- but almost every visit finds
// nothing changed, and Drive can tell us that in ONE call.
//
// getLastUpdated() on the roster workbook is a single cheap Drive read. Compare
// it with the stamp from the last rebuild: same means the notes cannot have
// moved, so the stored list is already correct and we return it as-is in about
// a second. Different means somebody typed a note, or the help desk wrote a
// ticket description into a device row, so the rebuild runs.
//
// A ticket ALWAYS changes the roster workbook, because rosterNoteWrite_ puts the
// description in the device's row. So filing a ticket guarantees the next
// dashboard load rebuilds and picks it up.
var ROSTER_STAMP_KEY = 'rosterCheckedAt';

function rosterStamp_() {
  try {
    return DriveApp.getFileById(ROSTER_SHEET_ID).getLastUpdated().toISOString();
  } catch (e) {
    Logger.log('could not read the roster timestamp: ' + e);
    return '';                      // unknown: fall through and rebuild
  }
}

function validateTodos_(p) {
  var t0 = new Date().getTime();
  var stamp = rosterStamp_();
  var props = PropertiesService.getScriptProperties();
  var seen = props.getProperty(ROSTER_STAMP_KEY);

  if (stamp && stamp === seen && String((p && p.force) || '') !== 'true') {
    var cur = todoList_();
    return { ok: true, fresh: true, checked: stamp, ms: new Date().getTime() - t0,
             todos: cur.todos, hidden: cur.hidden, carts: cur.carts };
  }

  rosterHeadReset_();
  var res = rebuildRosterTodos_();
  rebuildTicketTodos();
  var carts = cartTabsRefresh_();
  if (stamp) props.setProperty(ROSTER_STAMP_KEY, stamp);
  var list = todoList_();
  var ms = new Date().getTime() - t0;
  Logger.log('validate: rebuilt in ' + ms + 'ms (' + res.added + ' new, ' + res.removed + ' gone)');
  return { ok: true, fresh: false, checked: stamp, ms: ms,
           added: res.added, dropped: res.removed,
           imported: res.imported, kept: res.kept,
           todos: list.todos, hidden: list.hidden, carts: carts };
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
// work you did on an item that is still there. That key ignores the tab, so two
// tabs showing the same note share one status: it is one physical problem.

// The trailing "  (CB #4 · S/N ABC123)" that deviceTag_ adds, so the original
// note text can be recovered from a stored label.
var TODO_TAG_RE = /\s*\([^)]*S\/N[^)]*\)\s*$/;

function todoStateKey_(sn, noteText) {
  return noteKey_(sn, String(noteText).replace(TODO_TAG_RE, '').trim());
}

// Ids are roster-<serial>-<column>-<sheetId>. The sheet id is what keeps two
// tabs carrying the same note on the same device apart -- without it both rows
// would claim the same id and only one could exist.
//
// The serial is matched non-greedily and the two numeric parts anchor the end.
// Serials are letters and digits only (see looksLikeRosterSerial_), so they
// cannot swallow the numbers.
var ROSTER_ID_RE = /^roster-(.+?)-(\d+)-(\d+)$/;

function todoIdSerial_(id) {
  var m = String(id || '').match(ROSTER_ID_RE);
  return m ? m[1] : '';
}

function rebuildRosterTodos_() {
  var sh = todoSheet_();
  var width = TODO_HEADERS.length;

  // ---- 1. read the sheet as it stands, raw, so Created survives ----
  var keep = [];        // ticket- rows only; everything else is rebuilt
  var before = {};      // state key -> { done, status }
  var beforeIds = {};   // ids that were on the list, for the changed counts
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
      // Two tabs can carry the same note on the same device, so several rows can
      // share one state key. Merge rather than let the last one read win, or the
      // copy you marked Working would be reset by the copy you had not touched.
      // Furthest-along status wins, and done on either counts as done.
      var key = todoStateKey_(sn, rows[i][1]);
      var cand = { done: (rows[i][2] === true || rows[i][2] === 'TRUE'), status: st };
      var held = before[key];
      if (!held) {
        before[key] = cand;
      } else {
        before[key] = {
          done: held.done || cand.done,
          status: TODO_STATUSES.indexOf(cand.status) > TODO_STATUSES.indexOf(held.status)
                    ? cand.status : held.status
        };
      }
      beforeIds[id] = true;
    }
  }

  // ---- 2. read every roster note ----
  // A ticket writes its description into the device's row in the roster, and it
  // already has a ticket- item on the list. Without this the rebuild would
  // import that same text again as a roster- item and the task would appear
  // twice. Keyed the same way, so a reworded ticket stops matching and the
  // roster copy comes back -- which is the right answer, because the text in the
  // roster is then somebody else's.
  var ticketNotes = {};
  try {
    var tsheet = firstSheet_();
    var tLast = tsheet.getLastRow();
    if (tLast > 1) {
      var tv = tsheet.getRange(2, 1, tLast - 1, HEADERS.length).getValues();
      for (var ti = 0; ti < tv.length; ti++) {
        if (String(tv[ti][8] || 'New') === 'Resolved') continue;   // closed: note is cleared
        var tsn = String(tv[ti][0] || '').trim();
        var tdesc = String(tv[ti][7] || '').trim();
        if (tsn && tdesc) ticketNotes[todoStateKey_(tsn, tdesc)] = true;
      }
    }
  } catch (e) {
    Logger.log('could not read open tickets: ' + e);
  }

  var fresh = [];
  var afterIds = {};
  var handled = {};
  var handledList = handledNotes_();
  for (var q = 0; q < handledList.length; q++) handled[handledList[q]] = true;

  // Walked in workbook order, so the Order column matches the grid.
  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var byName = {};
  ss.getSheets().forEach(function (s) { byName[s.getName()] = s; });
  var tabs = sortByMasterOrder_(Object.keys(byName)).map(function (n) { return byName[n]; });
  for (var t = 0; t < tabs.length; t++) {
    var tab = tabs[t];
    if (todoIgnoreTab_(tab.getName())) continue;   // not a cart, or removed from the page
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;                          // not a roster tab
    var noteCols = rosterNoteColumns_(tab);
    if (!noteCols.length) continue;
    var numCol = rosterNumberColumn_(tab);
    var tabLastRow = tab.getLastRow();
    var dataRow = rosterDataRow_(tab);
    if (tabLastRow < dataRow) continue;
    var tabLastCol = tab.getLastColumn();
    var heads2 = rosterHeads_(tab);
    var vals = tab.getRange(dataRow, 1, tabLastRow - dataRow + 1, tabLastCol).getValues();

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
        // An open ticket already covers this exact note on this exact device.
        if (ticketNotes[key2]) continue;

        // EVERY tab that carries a note gets its own item, even when the same
        // device carries the same note on another tab. A device listed on two
        // carts is a roster matter; the to-do page's job is to show what each
        // tab actually says, so both carts show the work.
        var id2 = 'roster-' + sn2 + '-' + noteCols[k] + '-' + tab.getSheetId();
        if (afterIds[id2]) continue;              // same cell twice: impossible, but cheap
        afterIds[id2] = true;

        // Status is keyed on serial + note text, deliberately WITHOUT the tab:
        // two copies of one note share a status, and renaming a tab does not
        // reset it.
        var prev = before[key2] || { done: false, status: TODO_STATUS_DEFAULT };

        // A note that was ticked off but is still typed in the master belongs on
        // the "notes to clear" report. Log it once, not on every refresh.
        if (prev.done && !handled[noteKey_(sn2, note)]) {
          markNoteHandled_(sn2, note);
          logClosed_(id2, note, sn2, tab.getName(), 'ticked off');
          handled[noteKey_(sn2, note)] = true;
        }

        fresh.push([id2,
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
  for (var ka in afterIds) { if (!beforeIds[ka]) added++; }
  for (var kb in beforeIds) { if (!afterIds[kb]) removed++; }

  Logger.log('todo rebuild: ' + fresh.length + ' roster item(s) from the master, ' +
             keep.length + ' ticket item(s) kept, ' +
             added + ' new, ' + removed + ' gone');
  return { imported: fresh.length, kept: keep.length, added: added, removed: removed };
}

// Which carts actually exist, by name, straight off the roster workbook. The
// dashboard draws its grid from this instead of a hardcoded list, so a tile
// whose tab has been renamed or deleted stops being drawn, and a tab added to
// the workbook turns up on its own.
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
      if (isNonRosterTab_(tab.getName())) continue;
      if (isSkippedTab_(tab.getName())) continue;   // per-teacher list, not a cart
      if (!rosterSerialColumn_(tab)) continue;      // not a roster tab
      names.push(tab.getName());
    }
  } catch (e) {
    Logger.log('could not list cart tabs: ' + e);
  }
  return sortByMasterOrder_(names);   // the grid follows the workbook
}

// rosterCartTabs_ walks every tab in the roster workbook, which is too slow to
// do on every todoList call -- that is what the dashboard asks for each time the
// To-Do tab is opened, and it used to be a quick read of one sheet. The list only
// changes when the workbook does, so it is worked out on a refresh and parked in a
// script property for everything else to read instantly.
var CART_TABS_KEY = 'cartTabsCache';

function cartTabsCached_() {
  try {
    var v = JSON.parse(PropertiesService.getScriptProperties().getProperty(CART_TABS_KEY) || 'null');
    if (Object.prototype.toString.call(v) === '[object Array]') return v;
  } catch (e) {}
  return null;                    // never cached; caller works it out the slow way
}

function cartTabsRefresh_() {
  var names = rosterCartTabs_();
  try {
    PropertiesService.getScriptProperties().setProperty(CART_TABS_KEY, JSON.stringify(names));
  } catch (e) {
    Logger.log('could not cache cart tabs: ' + e);
  }
  return names;
}

// Diagnostic. Lists every tab in the roster workbook in order and says whether
// the to-do code can see it, so a tab that is missing from the dashboard can be
// told apart from one that is there but laid out differently.
//
// Run it from the editor and read View > Execution log.
function listCartTabs() {
  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var byName = {};
  ss.getSheets().forEach(function (sh) { byName[sh.getName()] = sh; });
  var tabs = sortByMasterOrder_(Object.keys(byName)).map(function (n) { return byName[n]; });
  var seen = 0;
  Logger.log(tabs.length + ' tab(s) in the roster workbook, in tab order:');
  Logger.log('');
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i];
    var name = tab.getName();
    if (isNonRosterTab_(name)) {
      Logger.log(pad_(i + 1) + name + '  -- skipped (ticket/todo tab, not roster)');
      continue;
    }
    if (isSkippedTab_(name)) {
      Logger.log(pad_(i + 1) + name + '  -- skipped (on TODO_SKIP_TABS, not a cart)');
      continue;
    }
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) {
      // Show the first three rows, so a different layout is obvious.
      var lc = Math.min(tab.getLastColumn(), 8);
      Logger.log(pad_(i + 1) + name + '  -- NOT DETECTED (nothing naming a serial in rows 1-3)');
      for (var q = 1; q <= Math.min(3, tab.getLastRow()); q++) {
        Logger.log('        row' + q + ': ' + (lc ? tab.getRange(q, 1, 1, lc).getValues()[0].join(' | ') : ''));
      }
      continue;
    }
    seen++;
    var noteCols = rosterNoteColumns_(tab);
    var numCol = rosterNumberColumn_(tab);
    var hdrRow = rosterHeaderRow_(tab);
    var rows = Math.max(0, tab.getLastRow() - hdrRow);
    Logger.log(pad_(i + 1) + name + '  -- ok: headers=' + (hdrRow ? 'row' + hdrRow : 'none, data from row 1') +
               ', serial=col' + snCol +
               ', number=col' + numCol +
               ', notes=col' + (noteCols.length ? noteCols.join('/') : 'NONE') +
               ', ' + rows + ' data row(s)' +
               (isHiddenGroup_(name) ? '   [REMOVED on the dashboard]' : ''));
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

// Run listCartTabs against a DIFFERENT workbook, to see whether it holds carts
// and notes the to-do page could be built from. Reads only; changes nothing.
//
// Edit the id below and run it from the editor.
function listCartTabsIn() {
  var id = ASSIGNMENT_SHEET_ID;   // <-- workbook to examine

  var ss = SpreadsheetApp.openById(id);
  Logger.log('workbook: ' + DriveApp.getFileById(id).getName());
  Logger.log('');
  rosterHeadReset_();
  var tabs = ss.getSheets();
  var carts = 0, withNotes = 0;
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i];
    var name = tab.getName();
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) {
      Logger.log(pad_(i + 1) + name + '  -- no serial column, not a roster tab');
      continue;
    }
    carts++;
    var noteCols = rosterNoteColumns_(tab);
    if (noteCols.length) withNotes++;
    var hdrRow = rosterHeaderRow_(tab);
    Logger.log(pad_(i + 1) + name + '  -- headers=' + (hdrRow ? 'row' + hdrRow : 'none') +
               ', serial=col' + snCol +
               ', notes=col' + (noteCols.length ? noteCols.join('/') : 'NONE') +
               ', ' + Math.max(0, tab.getLastRow() - hdrRow) + ' data row(s)' +
               (isSkippedTab_(name) ? '   [on TODO_SKIP_TABS]' : ''));
  }
  Logger.log('');
  Logger.log(tabs.length + ' tab(s): ' + carts + ' look like rosters, ' +
             withNotes + ' of those have a note column.');
  Logger.log('A workbook with no note columns cannot drive the to-do list.');
}

// ---- Cross-workbook comparison ---------------------------------------------
// Reads all three books and reports where they disagree. Editor-only, reads
// only, and slow (~100 tabs) -- run it from the editor, not the web app.
//
// Cart names are normalised before comparing: a leading "C-" is Kyle's marker
// for "beginning-of-year check done" and means nothing structurally, and the
// iPad tab is spelled differently in each book.
var COMPARE_ASSIGNMENT_ID = ASSIGNMENT_SHEET_ID;

function normCart_(name) {
  var n = String(name || '').trim().toLowerCase();
  n = n.replace(/^c-\s*/, '');            // "C-Cart A" -> "cart a"
  n = n.replace(/[^a-z0-9]/g, '');        // "iPad's" / "Ipads" -> "ipads"
  return n;
}

// serial -> [tab names], for every roster-shaped tab in a workbook
function serialMap_(id, skipNonCarts) {
  var out = { byserial: {}, tabs: [], error: '' };
  var ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { out.error = String(e); return out; }
  rosterHeadReset_();
  var tabs = ss.getSheets();
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i], name = tab.getName();
    if (isNonRosterTab_(name)) continue;
    if (skipNonCarts && isSkippedTab_(name)) continue;
    var snCol = rosterSerialColumn_(tab);
    if (!snCol) continue;
    var dataRow = rosterDataRow_(tab), last = tab.getLastRow();
    if (last < dataRow) continue;
    out.tabs.push(name);
    var vals = tab.getRange(dataRow, snCol, last - dataRow + 1, 1).getValues();
    for (var r = 0; r < vals.length; r++) {
      var sn = String(vals[r][0] || '').trim();
      if (!looksLikeRosterSerial_(sn)) continue;
      var key = sn.toUpperCase();
      if (!out.byserial[key]) out.byserial[key] = [];
      if (out.byserial[key].indexOf(name) < 0) out.byserial[key].push(name);
    }
  }
  return out;
}

function cap_(list, n) {
  if (list.length <= n) return list;
  return list.slice(0, n).concat(['... and ' + (list.length - n) + ' more']);
}

function compareBooks() {
  Logger.log('reading the assignment roster...');
  var A = serialMap_(COMPARE_ASSIGNMENT_ID, false);
  Logger.log('reading the check workbook...');
  var C = serialMap_(ROSTER_SHEET_ID, false);
  if (A.error) { Logger.log('assignment roster: ' + A.error); return; }
  if (C.error) { Logger.log('check workbook: ' + C.error); return; }

  Logger.log('');
  Logger.log('================ WORKBOOK COMPARISON ================');
  Logger.log('assignment roster : ' + A.tabs.length + ' roster tabs, ' +
             Object.keys(A.byserial).length + ' distinct serials');
  Logger.log('check workbook    : ' + C.tabs.length + ' roster tabs, ' +
             Object.keys(C.byserial).length + ' distinct serials');
  Logger.log('');

  // ---- 1. tabs present in one book but not the other ----
  var an = {}, cn = {};
  A.tabs.forEach(function (t) { an[normCart_(t)] = t; });
  C.tabs.forEach(function (t) { cn[normCart_(t)] = t; });
  var onlyA = [], onlyC = [];
  for (var k in an) if (!cn[k]) onlyA.push(an[k]);
  for (var k2 in cn) if (!an[k2]) onlyC.push(cn[k2]);
  Logger.log('--- tabs only in the assignment roster (' + onlyA.length + ') ---');
  cap_(onlyA.sort(), 25).forEach(function (t) { Logger.log('    ' + t); });
  Logger.log('--- tabs only in the check workbook (' + onlyC.length + ') ---');
  cap_(onlyC.sort(), 25).forEach(function (t) { Logger.log('    ' + t); });
  Logger.log('');

  // ---- 2. devices in one book but not the other ----
  var missC = [], missA = [], moved = [];
  for (var sn in A.byserial) {
    if (!C.byserial[sn]) { missC.push(sn + '  (' + A.byserial[sn].join(', ') + ')'); continue; }
    var a0 = normCart_(A.byserial[sn][0]), c0 = normCart_(C.byserial[sn][0]);
    if (a0 !== c0) moved.push(sn + '   roster: ' + A.byserial[sn].join('/') +
                              '   check: ' + C.byserial[sn].join('/'));
  }
  for (var sn2 in C.byserial) if (!A.byserial[sn2]) missA.push(sn2 + '  (' + C.byserial[sn2].join(', ') + ')');

  Logger.log('--- in the assignment roster but NOT the check workbook (' + missC.length + ') ---');
  cap_(missC.sort(), 40).forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');
  Logger.log('--- in the check workbook but NOT the assignment roster (' + missA.length + ') ---');
  cap_(missA.sort(), 40).forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');
  Logger.log('--- on DIFFERENT carts in the two books (' + moved.length + ') ---');
  cap_(moved.sort(), 40).forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');

  // ---- 3. a serial on more than one tab, within a book ----
  function dupes(m, label) {
    var d = [];
    for (var sn in m.byserial) if (m.byserial[sn].length > 1) d.push(sn + '  ' + m.byserial[sn].join(' | '));
    Logger.log('--- listed on several tabs in the ' + label + ' (' + d.length + ') ---');
    cap_(d.sort(), 30).forEach(function (x) { Logger.log('    ' + x); });
    Logger.log('');
  }
  dupes(A, 'assignment roster');
  dupes(C, 'check workbook');

  // ---- 4. the to-do list against the check workbook ----
  var todos = todoList_().todos;
  var badCart = [], badSerial = [];
  var cartNames = {};
  C.tabs.forEach(function (t) { cartNames[t] = true; });
  for (var i = 0; i < todos.length; i++) {
    var t = todos[i];
    if (String(t.id).indexOf('roster-') !== 0) continue;
    if (t.group && !cartNames[t.group]) badCart.push(t.group + ' :: ' + t.text);
    var m = String(t.text).match(/S\/N\s+([A-Za-z0-9]+)\)/);
    if (m && !C.byserial[m[1].toUpperCase()]) badSerial.push(m[1] + ' :: ' + t.text);
  }
  Logger.log('--- to-do items on a cart that no longer exists (' + badCart.length + ') ---');
  cap_(badCart.sort(), 25).forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');
  Logger.log('--- to-do items whose serial is not in the check workbook (' + badSerial.length + ') ---');
  cap_(badSerial.sort(), 25).forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');
  Logger.log('================ END ================');
}

// Why a tab produced the items it did -- or none. listCartTabs says whether a
// tab is SEEN; this says what the importer makes of every row on it, and gives
// the reason each cell was passed over.
//
// Edit the name below and run it from the editor. No redeploy needed; this is
// never called by the web app.
function debugTab() {
  var name = 'Cart E';                 // <-- the tab to examine

  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var tab = ss.getSheetByName(name);
  if (!tab) {
    Logger.log('There is no tab named "' + name + '" in the roster workbook.');
    return;
  }
  rosterHeadReset_();
  var info = rosterHeadInfo_(tab);
  var noteCols = rosterNoteColumns_(tab);
  var numCol = rosterNumberColumn_(tab);
  var dataRow = rosterDataRow_(tab);

  Logger.log('tab "' + name + '"');
  Logger.log('  skipped entirely?  ' + (todoIgnoreTab_(name) ? 'YES' : 'no'));
  Logger.log('  header row         ' + (info.row || 'none, data from row 1'));
  Logger.log('  serial column      ' + (info.serialCol || 'NOT FOUND'));
  Logger.log('  number column      ' + numCol);
  Logger.log('  headers            ' + info.heads.join(' | '));
  Logger.log('  note columns       ' + (noteCols.length ? noteCols.join(', ') : 'NONE'));
  Logger.log('');

  if (!info.serialCol || !noteCols.length) {
    Logger.log('  Nothing can be imported: a tab needs both a serial column and');
    Logger.log('  at least one note column. Note columns are col 3 or later whose');
    Logger.log('  header is blank or contains note/issue/repair/comment/problem/');
    Logger.log('  damage/broken/status, and never one containing "student".');
    return;
  }

  var lastRow = tab.getLastRow();
  var vals = tab.getRange(dataRow, 1, lastRow - dataRow + 1, tab.getLastColumn()).getValues();
  var wouldImport = 0;
  var badSerials = 0;
  for (var r = 0; r < vals.length; r++) {
    var row = r + dataRow;
    var sn = String(vals[r][info.serialCol - 1] || '').trim();
    if (!looksLikeRosterSerial_(sn)) {
      badSerials++;
      if (sn) Logger.log('  row ' + row + ': serial "' + sn + '" rejected - whole row skipped');
      continue;
    }
    for (var k = 0; k < noteCols.length; k++) {
      var raw = vals[r][noteCols[k] - 1];
      if (raw === '' || raw === null || raw === undefined) continue;   // blank, unremarkable
      if (!isRealNote_(raw)) {
        Logger.log('  row ' + row + ' col' + noteCols[k] + ': "' + String(raw) +
                   '" (' + typeof raw + ') not treated as a note');
        continue;
      }
      wouldImport++;
      Logger.log('  row ' + row + ' col' + noteCols[k] + ': IMPORT "' + String(raw).trim() +
                 '"  sn=' + sn);
    }
  }
  Logger.log('');
  Logger.log('  ' + wouldImport + ' item(s) would be imported, ' +
             badSerials + ' row(s) skipped for an unusable serial, ' +
             'out of ' + vals.length + ' row(s) read from row ' + dataRow + ' down.');
}

// Run by hand in the editor to rebuild without going through the dashboard.
function rebuildTodosNow() {
  rosterHeadReset_();
  Logger.log(JSON.stringify(rebuildRosterTodos_()));
  rebuildTicketTodos();
  Logger.log('carts: ' + JSON.stringify(cartTabsRefresh_()));
}

// The one-time cross-check. Idempotent, so re-running it is harmless.
// No longer on the Refresh path -- rebuildRosterTodos_ replaced it -- but kept
// for running an import by hand.
function importRosterNotesNow() {
  Logger.log(importRosterNotes_() + ' item(s) added');
}

// Rebuild the ticket-driven items from the ticket sheet: one per open ticket,
// none for closed ones. The cart a serial belongs to is read from the roster
// workbook.
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

// ---- Scheduled refresh ------------------------------------------------------
// Four times a school day: pull the master across, pick up any notes typed into
// the carts since, and square the to-do list with the open tickets.
var TODO_REFRESH_FN = 'scheduledTodoRefresh';
var TODO_REFRESH_HOURS = [7, 9, 12, 15];        // school time zone, see setupTodoTriggers

function scheduledTodoRefresh() {
  rosterHeadReset_();
  var res = rebuildRosterTodos_();
  rebuildTicketTodos();
  cartTabsRefresh_();
  Logger.log('scheduled refresh done. roster items: ' + res.imported + ', new: ' + res.added +
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

// First thing to run when something is not working, and the thing to run after
// any of the three ids above change. Proves each one can actually be reached and
// written, rather than assuming it.
function checkSetup() {
  Logger.log('script version:   ' + SCRIPT_VERSION);
  try {
    // Needs the userinfo.email scope. Useful to know, but never worth failing
    // the whole diagnostic over.
    Logger.log('running as:       ' + Session.getEffectiveUser().getEmail());
  } catch (e) {
    Logger.log('running as:       (add the userinfo.email scope to see this)');
  }
  Logger.log('');

  // ---- roster workbook: read + write ----
  try {
    var rf = DriveApp.getFileById(ROSTER_SHEET_ID);
    Logger.log('roster workbook:  ' + rf.getName());
    Logger.log('  owner:          ' + rf.getOwner().getEmail());
    var native = String(rf.getMimeType()).indexOf('google-apps.spreadsheet') >= 0;
    Logger.log('  type:           ' + rf.getMimeType() +
               (native ? '  (native)' : '  WRONG - an .xlsx cannot be opened or written'));
    var rs = SpreadsheetApp.openById(ROSTER_SHEET_ID);
    Logger.log('  tabs:           ' + rs.getSheets().length);
    var rp = rs.getSheets()[0];
    rp.getRange(1, rp.getMaxColumns()).setValue('write test').clearContent();
    Logger.log('  write access:   OK  (ticket notes can be written)');
  } catch (e) {
    Logger.log('  roster:         FAILED - ' + e);
  }
  Logger.log('');

  // ---- assignment roster: read only (the Lookup tab's assignment section) ----
  try {
    var af = DriveApp.getFileById(ASSIGNMENT_SHEET_ID);
    Logger.log('assignment book:  ' + af.getName());
    Logger.log('  owner:          ' + af.getOwner().getEmail());
    var assignSs = SpreadsheetApp.openById(ASSIGNMENT_SHEET_ID);
    Logger.log('  tabs:           ' + assignSs.getSheets().length);
    Logger.log('  read access:    OK  (Lookup can find assignments)');
  } catch (e) {
    Logger.log('  assignment:     FAILED - ' + e);
  }
  Logger.log('');

  // ---- ticket workbook: read + write, and the tabs the code expects ----
  try {
    var tf = DriveApp.getFileById(TICKET_BOOK_ID);
    Logger.log('ticket workbook:  ' + tf.getName());
    Logger.log('  owner:          ' + tf.getOwner().getEmail());
    var ts = ticketBook_();
    var names = ts.getSheets().map(function (sh) { return sh.getName(); });
    Logger.log('  tabs:           ' + names.join(', '));
    Logger.log('  ticket sheet:   ' + firstSheet_().getName() +
               '  (' + Math.max(0, firstSheet_().getLastRow() - 1) + ' ticket row(s))');
    Logger.log('  Todos tab:      ' + (names.indexOf(TODO_SHEET_NAME) >= 0
                 ? todoList_().todos.length + ' item(s)' : 'MISSING - will be created'));
    var tp = ts.getSheets()[0];
    tp.getRange(1, tp.getMaxColumns()).setValue('write test').clearContent();
    Logger.log('  write access:   OK');
  } catch (e) {
    Logger.log('  ticket book:    FAILED - ' + e);
  }
  Logger.log('');

  // ---- photo folder: read + write ----
  try {
    var pf = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    Logger.log('photo folder:     ' + pf.getName());
    Logger.log('  owner:          ' + pf.getOwner().getEmail());
    var probe = pf.createFile(Utilities.newBlob('cpa-it test', 'text/plain', 'cpa-it-test.txt'));
    probe.setTrashed(true);
    Logger.log('  write access:   OK  (ticket photos will save here)');
  } catch (e) {
    Logger.log('  photo folder:   FAILED - ' + e);
    Logger.log('  share it as Editor with the account above, then run again.');
  }
  Logger.log('');
  Logger.log('ticket counter:   lastTicketNo = ' +
             (PropertiesService.getScriptProperties().getProperty('lastTicketNo') || '(unset, starts at 1001)'));
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
  var ss = ticketBook_();
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
  // Cached: the slow walk happens on a refresh, not on every tab open.
  return { ok: true, todos: todos, hidden: hiddenGroups_(),
           carts: cartTabsCached_() || cartTabsRefresh_() };
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
  var ss = ticketBook_();
  var src = firstSheet_();               // by header, not by position
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

// ---- Sheet editor -----------------------------------------------------------
// The Editor tab on the dashboard: browse one tab of one book and edit cells, or
// search across books and apply one change to rows from several tabs at once.
//
// Three rules hold this together, and all three are enforced HERE rather than in
// the browser, because the browser is only one of the ways this endpoint can be
// called:
//
//   1. Only these three workbooks. A book is named by a key, never by an id sent
//      from outside, so no request can point this at another spreadsheet.
//   2. Serial columns are never writable. Not a UI convention -- gridWrite_
//      refuses them. A wrong serial is a Sheets job, deliberately.
//   3. Every write carries the value the browser last saw. If the cell no longer
//      holds it, the write is refused rather than clobbering whoever changed it
//      in Sheets in the meantime.
var GRID_BOOKS = {
  assignment: { id: ASSIGNMENT_SHEET_ID, label: '2026-2027 Chromebook Carts / iPads' },
  check:      { id: NOTE_BOOK_ID,        label: 'Start-of-Year Check' },
  tickets:    { id: TICKET_BOOK_ID,      label: 'Tickets' }
};
var GRID_ROW_CAP = 400;         // rows returned per tab
var GRID_HIT_CAP = 200;         // search hits returned

var EDIT_LOG_NAME = 'EditLog';
var EDIT_LOG_HEADERS = ['When', 'Book', 'Tab', 'Row', 'Col', 'Header', 'Was', 'Now', 'Batch'];

function gridBook_(key) {
  var b = GRID_BOOKS[String(key || '')];
  if (!b) return null;
  return b;
}

function gridSheet_(key, tab) {
  var b = gridBook_(key);
  if (!b) return null;
  var sh = SpreadsheetApp.openById(b.id).getSheetByName(String(tab || ''));
  if (!sh) return null;
  return sh;
}

// Where the headers are and which column holds the serial. The ticket book is
// a plain table with headers on row 1; the two roster books go through
// rosterHeadInfo_, which already copes with headers on rows 1-3 and with the
// headerless tabs.
function gridHeadInfo_(key, sh) {
  if (key === 'tickets') {
    var lastCol = Math.max(sh.getLastColumn(), 1);
    return { row: 1, heads: sh.getRange(1, 1, 1, lastCol).getValues()[0], serialCol: 1 };
  }
  return rosterHeadInfo_(sh);
}

function gridTabs_(p) {
  var b = gridBook_(p.book);
  if (!b) return { ok: false, error: 'Unknown workbook.' };
  var tabs = [];
  SpreadsheetApp.openById(b.id).getSheets().forEach(function (sh) {
    tabs.push(sh.getName());
  });
  return { ok: true, book: p.book, label: b.label, tabs: tabs };
}

function gridRead_(p) {
  var sh = gridSheet_(p.book, p.tab);
  if (!sh) return { ok: false, error: 'That tab was not found.' };
  var info = gridHeadInfo_(p.book, sh);
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var dataRow = info.row + 1;

  var out = {
    ok: true, book: p.book, tab: sh.getName(),
    headerRow: info.row, headers: [], serialCol: info.serialCol,
    title: '', room: '', rows: [], total: 0, truncated: false,
    // Tells the grid to render every cell locked. gridWrite_ refuses this tab
    // regardless; this just stops you typing into it first.
    readOnly: RELOCATION_LOG_RE.test(sh.getName()),
    readOnlyWhy: 'This tab mirrors the Loaners ledger in the ticket book. ' +
                 'Edit the Loaners tab and this updates itself.'
  };

  // Cart tabs in the assignment book carry the teacher in A1 and the room in B1,
  // both with a baked-in label. The page shows them stripped and writes them
  // back with the label re-applied -- see gridApplyLabel_.
  if (info.row > 1) {
    out.title = stripRosterLabel_(sh.getRange(1, 1).getValue());
    if (lastCol > 1) out.room = stripRosterLabel_(sh.getRange(1, 2).getValue());
  }

  for (var c = 0; c < lastCol; c++) {
    var head = '';
    if (info.heads && info.heads[c] != null) head = String(info.heads[c]);
    out.headers.push(head);
  }

  if (lastRow >= dataRow) {
    out.total = lastRow - dataRow + 1;
    var take = out.total;
    if (take > GRID_ROW_CAP) { take = GRID_ROW_CAP; out.truncated = true; }
    var vals = sh.getRange(dataRow, 1, take, lastCol).getValues();
    for (var r = 0; r < vals.length; r++) {
      var cells = [];
      for (var c2 = 0; c2 < lastCol; c2++) {
        var v = vals[r][c2];
        if (Object.prototype.toString.call(v) === '[object Date]') {
          cells.push(Utilities.formatDate(v, Session.getScriptTimeZone(), 'M/d/yyyy'));
        } else if (v === true || v === false) {
          cells.push(String(v));
        } else {
          cells.push(String(v == null ? '' : v));
        }
      }
      out.rows.push({ row: dataRow + r, cells: cells });
    }
  }
  return out;
}

// Cross-book search. createTextFinder is the same trick deviceLookup_ uses: it
// searches a whole workbook without this script reading 49 tabs itself, which is
// the only reason a search across all three books is affordable at all.
function gridSearch_(p) {
  var q = String(p.q || '').trim();
  if (q.length < 2) return { ok: false, error: 'Type at least two characters.' };
  var keys = String(p.books || 'assignment,check').split(',');
  var out = { ok: true, hits: [], truncated: false };

  // Header info per tab, worked out once. The ticket book reads row 1 from the
  // sheet every time it is asked, and two hundred hits on one tab would other-
  // wise be two hundred reads of the same row.
  var headMemo = {};
  function headFor(key, sh) {
    var memoKey = key + '|' + sh.getSheetId();
    if (!headMemo[memoKey]) headMemo[memoKey] = gridHeadInfo_(key, sh);
    return headMemo[memoKey];
  }

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i].trim();
    var b = gridBook_(key);
    if (!b) continue;
    var ss;
    try { ss = SpreadsheetApp.openById(b.id); }
    catch (e) { continue; }

    var found = ss.createTextFinder(q).findAll();
    for (var f = 0; f < found.length; f++) {
      if (out.hits.length >= GRID_HIT_CAP) { out.truncated = true; break; }
      var rng = found[f];
      var sh = rng.getSheet();
      var row = rng.getRow();
      var info = headFor(key, sh);
      if (row <= info.row) continue;                 // a header, not data
      var head = '';
      if (info.heads && info.heads[rng.getColumn() - 1] != null) {
        head = String(info.heads[rng.getColumn() - 1]);
      }
      var sn = '';
      if (info.serialCol) sn = String(sh.getRange(row, info.serialCol).getValue() || '');
      out.hits.push({
        book: key, bookLabel: b.label, tab: sh.getName(), row: row,
        col: rng.getColumn(), header: head, sn: sn,
        value: String(rng.getValue() == null ? '' : rng.getValue())
      });
    }
    if (out.truncated) break;
  }
  return out;
}

// A1 on a cart tab holds "TEACHER: Miller (D)" and B1 holds "ROOM# 210". The
// editor shows those stripped, so writing back what the page sends would drop
// the label and quietly break everything that reads teacher and room. Whatever
// prefix the cell already carries goes back on.
function gridApplyLabel_(current, next) {
  var s = String(current == null ? '' : current);
  var m = s.match(/^(teacher\s*#?\s*:?\s*|room\s*#?\s*:?\s*)/i);
  if (!m) return next;
  return m[1] + next;
}

function editLogSheet_() {
  var ss = ticketBook_();
  var sh = ss.getSheetByName(EDIT_LOG_NAME);
  if (sh) return sh;
  sh = ss.insertSheet(EDIT_LOG_NAME);
  sh.getRange(1, 1, 1, EDIT_LOG_HEADERS.length).setValues([EDIT_LOG_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// Every edit, one at a time, each guarded by the value the browser last saw.
//
// The browser sends these in chunks because this endpoint is reached over JSONP,
// which is a GET -- a hundred edits do not fit in a URL. Every chunk of one save
// carries the same batch id, so gridUndo_ can put the whole save back however
// many requests it arrived in.
function gridWrite_(p) {
  var edits;
  try { edits = JSON.parse(p.edits || '[]'); }
  catch (e) { return { ok: false, error: 'Could not read the edits.' }; }
  if (!edits.length) return { ok: true, applied: 0, results: [] };

  var batch = String(p.batch || '').trim();
  if (!batch) batch = 'b' + new Date().getTime();
  var log = editLogSheet_();
  var logRows = [];
  var results = [];
  var applied = 0;
  var now = new Date();

  for (var i = 0; i < edits.length; i++) {
    var e = edits[i];
    var res = { i: i, ok: false, why: '' };
    var sh = gridSheet_(e.b, e.t);
    if (!sh) { res.why = 'tab not found'; results.push(res); continue; }

    // The Loaner_log tabs are an IMPORTRANGE mirror of the Loaners ledger.
    // Writing anywhere inside a spilled array turns it into #REF! and takes the
    // mirror with it, so the whole tab is refused rather than any one cell.
    if (RELOCATION_LOG_RE.test(sh.getName())) {
      res.why = 'that tab mirrors the Loaners ledger — edit the ledger instead';
      results.push(res);
      continue;
    }

    var info = gridHeadInfo_(e.b, sh);
    var col = Number(e.c), row = Number(e.r);
    if (!col || !row) { res.why = 'bad cell'; results.push(res); continue; }

    // Rule 2. The serial identifies the device to every other part of this
    // script; nothing here is allowed to change it.
    if (info.serialCol && col === info.serialCol && row > info.row) {
      res.why = 'serial columns are locked';
      results.push(res);
      continue;
    }

    var cell = sh.getRange(row, col);
    var cur = cell.getValue();
    var curText = String(cur == null ? '' : cur);
    if (Object.prototype.toString.call(cur) === '[object Date]') {
      curText = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'M/d/yyyy');
    }
    // Rule 3. Somebody edited this in Sheets while the grid was open.
    //
    // The teacher and room cells are shown to the page with their label peeled
    // off ("Miller (D)", not "TEACHER: Miller (D)"), so the value the page sends
    // back as "what I last saw" is the stripped one. Both spellings count as a
    // match; anything else really has changed underneath.
    var want = String(e.w == null ? '' : e.w).trim();
    var stripped = stripRosterLabel_(curText);
    if (curText.trim() !== want && stripped !== want) {
      res.why = 'changed in the sheet since you loaded it (now "' + curText + '")';
      results.push(res);
      continue;
    }

    var next = String(e.n == null ? '' : e.n);
    var written = gridApplyLabel_(cur, next);
    if (written === '') cell.clearContent();
    else cell.setValue(written);

    var head = '';
    if (info.heads && info.heads[col - 1] != null) head = String(info.heads[col - 1]);
    logRows.push([now, e.b, sh.getName(), row, col, head, curText, written, batch]);
    res.ok = true;
    applied++;
    results.push(res);
  }

  if (logRows.length) {
    log.getRange(log.getLastRow() + 1, 1, logRows.length, EDIT_LOG_HEADERS.length).setValues(logRows);
  }
  return { ok: true, applied: applied, batch: batch, results: results };
}

// Put a whole save back. Guarded the other way round: a cell is only reverted
// while it still holds what this batch wrote, so an edit made after the batch is
// never silently undone.
function gridUndo_(p) {
  var batch = String(p.batch || '').trim();
  if (!batch) return { ok: false, error: 'No batch given.' };
  var log = editLogSheet_();
  var last = log.getLastRow();
  if (last < 2) return { ok: false, error: 'Nothing logged yet.' };

  var vals = log.getRange(2, 1, last - 1, EDIT_LOG_HEADERS.length).getValues();
  var reverted = 0, skipped = 0;
  for (var i = vals.length - 1; i >= 0; i--) {         // newest first
    if (String(vals[i][8]) !== batch) continue;
    var sh = gridSheet_(vals[i][1], vals[i][2]);
    if (!sh) { skipped++; continue; }
    var cell = sh.getRange(Number(vals[i][3]), Number(vals[i][4]));
    var cur = String(cell.getValue() == null ? '' : cell.getValue());
    if (cur.trim() !== String(vals[i][7] == null ? '' : vals[i][7]).trim()) { skipped++; continue; }
    var back = String(vals[i][6] == null ? '' : vals[i][6]);
    if (back === '') cell.clearContent();
    else cell.setValue(back);
    reverted++;
  }
  return { ok: true, reverted: reverted, skipped: skipped };
}

// The most recent saves, so the Editor tab can offer to undo one.
function gridBatches_() {
  var log = editLogSheet_();
  var last = log.getLastRow();
  var out = { ok: true, batches: [] };
  if (last < 2) return out;
  var take = 400;
  var from = Math.max(2, last - take + 1);
  var vals = log.getRange(from, 1, last - from + 1, EDIT_LOG_HEADERS.length).getValues();
  var seen = {};
  for (var i = vals.length - 1; i >= 0; i--) {
    var id = String(vals[i][8]);
    if (!id) continue;
    if (!seen[id]) {
      seen[id] = { batch: id, when: new Date(vals[i][0]).toISOString(), count: 0, books: {} };
      out.batches.push(seen[id]);
    }
    seen[id].count++;
    seen[id].books[String(vals[i][1])] = true;
  }
  out.batches = out.batches.slice(0, 10);
  out.batches.forEach(function (b) { b.bookList = Object.keys(b.books).join(', '); delete b.books; });
  return out;
}

// ---- Spares / loaner pool ---------------------------------------------------
// The spares tab is laid out exactly like a cart tab -- a title on row 1,
// headers on row 2, devices from row 3 -- so every reader in this file finds it
// without a single change: rosterHeadInfo_ sees "Serial #" on row 2, and
// rosterFindRows_ (which requires a serial header on row 2 and data below row
// 2) matches it. rosterTabScore_ already gives a tab named "spares" a bonus, so
// a device sitting on both a cart and the spares tab notes on the spares tab.
//
// Run once from the editor. Safe to run again; it never touches an existing tab.
var SPARES_HEADERS = ['#', 'Serial #', 'Model', 'Condition', 'Loaner'];

function setupSparesTab() {
  var ss = SpreadsheetApp.openById(NOTE_BOOK_ID);
  var sh = ss.getSheetByName(SPARES_TAB_NAME);
  if (sh) {
    Logger.log('the "' + SPARES_TAB_NAME + '" tab already exists — nothing changed');
    return sh.getName();
  }
  sh = ss.insertSheet(SPARES_TAB_NAME);
  sh.getRange(1, 1).setValue('SPARES / LOANER POOL').setFontWeight('bold');
  sh.getRange(2, 1, 1, SPARES_HEADERS.length).setValues([SPARES_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(2);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(5, 260);
  rosterHeadReset_();
  cartTabsRefresh_();
  Logger.log('created "' + SPARES_TAB_NAME + '" — put a serial in column B from row 3 down');
  return sh.getName();
}

// ---- Loaner ledger ----------------------------------------------------------
// The Loaners tab IS the record. The notes this writes onto the two devices are
// a mirror of it, the same way the Todos tab mirrors the roster: nothing ever
// reads a note to decide whether a device is out, and loanerResync_() can
// rebuild every note from the ledger if one gets mangled in Sheets.
//
// There is no id column and no status column, on purpose. A loan is identified
// by inop serial + loaner serial + an empty "Date in", and only one loan for a
// pair can be open at a time, so a return never has to guess which row it
// means. Status is just "is Date in empty".
var LOANER_HEADERS = ['Date out', 'Inop S/N', 'Inop cart', 'Loaner S/N',
                      'Loaner source', 'Student', 'Note', 'Date in'];

// The header on the roster column that carries loaner notes.
//
// Deliberately NOT rosterNoteColumn_(): that returns the first BLANK-header
// column, which is where ticket notes land, so a device with both a ticket and
// a loan would have one silently overwrite the other.
//
// Giving the column a header also keeps loaner notes out of the to-do list for
// free. rosterNoteColumns_() only treats a column as a note source when its
// header is blank or matches NOTE_HEADER_RE (note|issue|repair|comment|problem|
// damage|broken|status), and "Loaner" is neither. Rename this and that stops
// being true -- "Loaner status" would start generating to-dos.
var LOANER_NOTE_HEADER = 'Loaner';

function loanerSheet_() {
  var ss = ticketBook_();
  var sh = ss.getSheetByName(LOANER_SHEET_NAME);
  if (sh) return sh;
  sh = ss.insertSheet(LOANER_SHEET_NAME);
  sh.getRange(1, 1, 1, LOANER_HEADERS.length).setValues([LOANER_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(7, 240);
  return sh;
}

// create=false when clearing: never add a column to a tab just to blank a cell
// in it.
function loanerNoteColumn_(sh, create) {
  var info = rosterHeadInfo_(sh);
  if (!info.row) return 0;                  // headerless tab (Speech): nowhere to put a header
  var heads = info.heads;
  for (var c = 0; c < heads.length; c++) {
    if (String(heads[c] || '').trim().toLowerCase() === LOANER_NOTE_HEADER.toLowerCase()) {
      return c + 1;
    }
  }
  if (!create) return 0;
  var col = sh.getLastColumn() + 1;
  sh.getRange(info.row, col).setValue(LOANER_NOTE_HEADER).setFontWeight('bold');
  rosterHeadReset_();                       // the cached header row is now stale
  return col;
}

function loanerShortDate_(d) {
  var dt = new Date(d);
  if (isNaN(dt.getTime())) return '';
  return (dt.getMonth() + 1) + '/' + dt.getDate();
}

// The two notes a loan puts on the two devices. Both the writer and the
// validator build them here, so a wording change can never make the validator
// report drift that does not exist.
function loanerInopNote_(loanerSn, dateOut, source) {
  var text = 'INOP — loaner ' + loanerSn + ' out ' + loanerShortDate_(dateOut);
  if (String(source || '').trim()) text += ' (' + String(source).trim() + ')';
  return text;
}

function loanerOutNote_(student, inopSn, dateOut) {
  var who = String(student || '').trim();
  var text = 'LOANER — out';
  if (who) text += ' to ' + who;
  return text + ', replaces ' + inopSn + ', ' + loanerShortDate_(dateOut);
}

// A relocation carries no broken device, so the loaner wording is wrong for it:
// it is not out to anyone in place of anything, it simply lives somewhere else
// now. Saying so plainly also stops the note reading "replaces , 8/11".
function loanerMovedNote_(source, dateOut) {
  var from = String(source || '').trim();
  var text = 'MOVED —';
  if (from) text += ' from ' + from + ',';
  return text + ' ' + loanerShortDate_(dateOut);
}

// The note that belongs on the device that went out, whichever kind of row this
// is. Every caller goes through here -- the writer, the validator and the return
// path -- so the three can never disagree about what a row's note should say.
function loanerDeviceNote_(L) {
  if (L.inopSn) return loanerOutNote_(L.student, L.inopSn, L.dateOut);
  return loanerMovedNote_(L.source, L.dateOut);
}

// Only ever writes into a cell that is empty or already holds a loaner note.
//
// Anything else in that column was typed by a person, and a re-sync across a
// ledger full of open loans would otherwise overwrite it with no undo behind it.
// Refusing instead means re-sync cannot destroy anything: the worst case is a
// note that does not get written, which the validator then reports.
var LOANER_NOTE_RE = /^(INOP|LOANER|MOVED)\s*[—-]/i;

function loanerNoteWrite_(sn, text) {
  try {
    var hits = rosterFindRows_(sn);
    if (!hits.length) return { ok: true, skipped: 'not on any tab' };
    var target = rosterPickBest_(hits);
    var col = loanerNoteColumn_(target.sheet, true);
    if (!col) return { ok: true, skipped: 'tab has no header row' };
    var cell = target.sheet.getRange(target.row, col);
    var cur = String(cell.getValue() || '').trim();
    if (cur && !LOANER_NOTE_RE.test(cur)) {
      return { ok: true, skipped: 'cell holds something typed by hand: "' + cur + '"' };
    }
    cell.setValue(text);
    return { ok: true, tab: target.sheet.getName(), row: target.row };
  } catch (e) {
    Logger.log('loaner note write failed for ' + sn + ': ' + e);
    return { ok: false, error: String(e) };
  }
}

// Clears on every tab the serial appears on, so nothing is left behind if the
// device moved tabs while it was out -- and only when the cell still holds our
// text, so a note somebody typed by hand is never wiped.
function loanerNoteClear_(sn, text) {
  try {
    var hits = rosterFindRows_(sn);
    var cleared = 0;
    for (var i = 0; i < hits.length; i++) {
      var col = loanerNoteColumn_(hits[i].sheet, false);
      if (!col) continue;
      var cell = hits[i].sheet.getRange(hits[i].row, col);
      var cur = String(cell.getValue() || '').trim();
      if (!cur) continue;
      if (cur !== String(text).trim()) continue;
      cell.clearContent();
      cleared++;
    }
    return { ok: true, cleared: cleared };
  } catch (e) {
    Logger.log('loaner note clear failed for ' + sn + ': ' + e);
    return { ok: false, error: String(e) };
  }
}

// Every row, newest first, with the two things the page would otherwise have to
// work out for itself: whether it is open, and how long it has been out.
function loanerList_() {
  var sh = loanerSheet_();
  var out = { ok: true, loans: [], openCount: 0 };
  var last = sh.getLastRow();
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, LOANER_HEADERS.length).getValues();
  var today = new Date();
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    var inopSn = String(v[1] || '').trim();
    var loanerSn = String(v[3] || '').trim();
    if (!inopSn && !loanerSn) continue;

    var dOut = null, dIn = null;
    if (v[0]) dOut = new Date(v[0]);
    if (v[7]) dIn = new Date(v[7]);
    var open = !dIn;

    var days = '';
    if (dOut) {
      var end = today;
      if (dIn) end = dIn;
      days = Math.round((end.getTime() - dOut.getTime()) / 86400000);
      if (days < 0) days = 0;
    }

    var rec = {
      row: i + 2,
      dateOut: '', inopSn: inopSn, inopCart: String(v[2] || ''),
      loanerSn: loanerSn, source: String(v[4] || ''),
      student: String(v[5] || ''), note: String(v[6] || ''),
      dateIn: '', open: open, days: days
    };
    if (dOut) rec.dateOut = dOut.toISOString();
    if (dIn) rec.dateIn = dIn.toISOString();
    out.loans.push(rec);
    if (open) out.openCount++;
  }
  out.loans.sort(function (a, b) { return String(b.dateOut).localeCompare(String(a.dateOut)); });
  return out;
}

// What the page puts behind the type-ahead fields: every value already used in
// that column, most recently used first, so the vocabulary teaches itself and
// never needs maintaining. Typing something new is always allowed.
function loanerOptions_() {
  var sh = loanerSheet_();
  var out = { ok: true, carts: [], sources: [], students: [] };
  var last = sh.getLastRow();
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, LOANER_HEADERS.length).getValues();
  var seen = { carts: {}, sources: {}, students: {} };

  function take(list, mark, value) {
    var s = String(value || '').trim();
    if (!s) return;
    var key = s.toLowerCase();
    if (mark[key]) return;
    mark[key] = true;
    list.push(s);
  }

  for (var i = vals.length - 1; i >= 0; i--) {      // newest row first
    take(out.carts, seen.carts, vals[i][2]);
    take(out.sources, seen.sources, vals[i][4]);
    take(out.students, seen.students, vals[i][5]);
  }
  return out;
}

// The open loan for a serial, on either side of the pair.
function loanerOpenFor_(loans, sn) {
  var key = String(sn || '').trim().toUpperCase();
  for (var i = 0; i < loans.length; i++) {
    if (!loans[i].open) continue;
    if (loans[i].inopSn.toUpperCase() === key) return loans[i];
    if (loans[i].loanerSn.toUpperCase() === key) return loans[i];
  }
  return null;
}

function loanerIssue_(p) {
  var inopSn = serialFromScan_(p.inopSn);
  var loanerSn = serialFromScan_(p.loanerSn);
  if (!inopSn) return { ok: false, error: 'The broken device needs a serial.' };
  if (!loanerSn) return { ok: false, error: 'The loaner needs a serial.' };
  if (inopSn.toUpperCase() === loanerSn.toUpperCase()) {
    return { ok: false, error: 'The loaner and the broken device are the same serial.' };
  }

  var existing = loanerList_().loans;
  var busy = loanerOpenFor_(existing, loanerSn);
  if (busy) {
    return { ok: false, error: 'That loaner is already out (row ' + busy.row + ', to ' +
             (busy.student || 'nobody named') + ' since ' + loanerShortDate_(busy.dateOut) + ').' };
  }
  var already = loanerOpenFor_(existing, inopSn);
  if (already && String(p.replace) !== '1') {
    return { ok: false, needsReplace: true, open: already,
             error: 'That device already has loaner ' + already.loanerSn + ' out (row ' +
                    already.row + ').' };
  }

  // Replacing an open loan: close the old one first, so the pair key stays
  // unique and the old loaner's note comes off.
  if (already && String(p.replace) === '1') {
    loanerReturn_({ row: already.row });
  }

  var now = new Date();
  var sh = loanerSheet_();
  var row = [now, inopSn, String(p.inopCart || '').trim(), loanerSn,
             String(p.source || '').trim(), String(p.student || '').trim(),
             String(p.note || '').trim(), ''];
  sh.appendRow(row);
  var rowNo = sh.getLastRow();

  var inopNote = loanerInopNote_(loanerSn, now, p.source);
  var outNote = loanerOutNote_(p.student, inopSn, now);
  var w1 = loanerNoteWrite_(inopSn, inopNote);
  var w2 = loanerNoteWrite_(loanerSn, outNote);

  return { ok: true, row: rowNo, inopSn: inopSn, loanerSn: loanerSn,
           inopNote: w1, loanerNote: w2 };
}

function loanerReturn_(p) {
  var sh = loanerSheet_();
  var loans = loanerList_().loans;
  var target = null;

  if (p.row) {
    for (var i = 0; i < loans.length; i++) {
      if (String(loans[i].row) === String(p.row)) { target = loans[i]; break; }
    }
  } else {
    // No row given: the open loan for whichever serial was handed in.
    var sn = serialFromScan_(p.sn || p.loanerSn || p.inopSn);
    target = loanerOpenFor_(loans, sn);
  }
  if (!target) return { ok: false, error: 'No open loan found for that.' };
  if (!target.open) return { ok: false, error: 'That loan was already returned.' };

  var c1 = loanerNoteClear_(target.inopSn, loanerInopNote_(target.loanerSn, target.dateOut, target.source));
  var c2 = loanerNoteClear_(target.loanerSn, loanerDeviceNote_(target));
  sh.getRange(target.row, 8).setValue(new Date());

  return { ok: true, row: target.row, cleared: (c1.cleared || 0) + (c2.cleared || 0) };
}

// Rebuild the device notes from the ledger. This is the answer to any note that
// got mangled or deleted in Sheets -- the ledger is the record, so the notes can
// always be made to match it again.
//
// Open rows are always checked. Closed rows are only checked when they closed in
// the last 30 days: an old return whose note is already gone costs a text-finder
// pass to prove nothing, and there is no upper bound on how many of those pile
// up over a year.
function loanerResync_() {
  var loans = loanerList_().loans;
  var written = 0, cleared = 0, checked = 0;
  var skipped = [];
  var cutoff = new Date().getTime() - (30 * 86400000);

  for (var i = 0; i < loans.length; i++) {
    var L = loans[i];

    var inopNote = loanerInopNote_(L.loanerSn, L.dateOut, L.source);
    var outNote = loanerDeviceNote_(L);      // LOANER - ... or MOVED - ...

    if (L.open) {
      checked++;
      // The device that went out always gets a note. The INOP note only exists
      // when there is a broken device to put it on, so relocations skip it.
      var w2 = loanerNoteWrite_(L.loanerSn, outNote);
      if (w2.ok && !w2.skipped) written++;
      if (w2.skipped) skipped.push(L.loanerSn + ' — ' + w2.skipped);
      if (L.inopSn) {
        var w1 = loanerNoteWrite_(L.inopSn, inopNote);
        if (w1.ok && !w1.skipped) written++;
        if (w1.skipped) skipped.push(L.inopSn + ' — ' + w1.skipped);
      }
      continue;
    }
    if (!L.dateIn) continue;
    if (new Date(L.dateIn).getTime() < cutoff) continue;
    checked++;
    cleared += (loanerNoteClear_(L.inopSn, inopNote).cleared || 0);
    cleared += (loanerNoteClear_(L.loanerSn, outNote).cleared || 0);
  }
  return { ok: true, checked: checked, written: written, cleared: cleared,
           skipped: skipped.length, skippedList: cap_(skipped, 20) };
}

// Read-only. Finds where loaner records are actually being kept, when they are
// not in the Loaners tab this script reads.
//
// Two passes, cheapest first: tab NAMES across all three books, which costs one
// getSheets() per book; then, only if that finds nothing, the first two rows of
// every tab looking for a header that mentions a loan. The second pass is ~110
// reads, so it is deliberately the fallback rather than the default.
function findLoanerLog() {
  var books = [
    { label: 'ticket book',     id: TICKET_BOOK_ID },
    { label: 'check workbook',  id: NOTE_BOOK_ID },
    { label: 'assignment book', id: ASSIGNMENT_SHEET_ID }
  ];
  var nameRe = /loan/i;
  var headRe = /loaner|loan\s*out|out\s*to|replaces/i;
  var hits = [];
  var all = [];

  for (var b = 0; b < books.length; b++) {
    var ss;
    try { ss = SpreadsheetApp.openById(books[b].id); }
    catch (e) { Logger.log(books[b].label + ': CANNOT OPEN — ' + e); continue; }
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var nm = sheets[s].getName();
      all.push({ book: books[b].label, sheet: sheets[s], name: nm });
      if (nameRe.test(nm)) {
        hits.push(books[b].label + ' · "' + nm + '"  (' +
                  Math.max(sheets[s].getLastRow() - 1, 0) + ' data row(s))');
      }
    }
  }

  Logger.log('=== tabs whose NAME mentions a loan (' + hits.length + ') ===');
  for (var h = 0; h < hits.length; h++) Logger.log('   ' + hits[h]);

  if (hits.length) {
    Logger.log('');
    Logger.log('If one of those is your log, tell me the book and tab name.');
    return hits;
  }

  Logger.log('');
  Logger.log('None. Scanning the first two rows of all ' + all.length + ' tabs...');
  var found = [];
  for (var i = 0; i < all.length; i++) {
    var sh = all[i].sheet;
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (!lastRow || !lastCol) continue;
    var vals;
    try { vals = sh.getRange(1, 1, Math.min(2, lastRow), lastCol).getValues(); }
    catch (e2) { continue; }
    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        if (typeof vals[r][c] !== 'string') continue;
        if (!headRe.test(vals[r][c])) continue;
        found.push(all[i].book + ' · "' + all[i].name + '"  row ' + (r + 1) +
                   ' col ' + (c + 1) + ': "' + vals[r][c] + '"  (' +
                   Math.max(lastRow - 1, 0) + ' data row(s))');
        r = vals.length;                 // one hit per tab is enough
        break;
      }
    }
  }
  Logger.log('=== tabs with a loaner-ish header (' + found.length + ') ===');
  for (var f = 0; f < found.length; f++) Logger.log('   ' + found[f]);
  if (!found.length) Logger.log('   nothing found — the log is in another spreadsheet entirely');
  return found;
}

// Proves the ledger and the spares tab are reachable and writable before you
// rely on them. Run from the editor.
function checkLoanerSetup() {
  var sh = loanerSheet_();
  Logger.log('ledger tab: "' + sh.getName() + '" in the ticket book, ' +
             Math.max(sh.getLastRow() - 1, 0) + ' row(s)');
  var spares = SpreadsheetApp.openById(NOTE_BOOK_ID).getSheetByName(SPARES_TAB_NAME);
  if (spares) {
    Logger.log('spares tab: "' + spares.getName() + '", ' +
               Math.max(spares.getLastRow() - 2, 0) + ' device(s)');
  } else {
    Logger.log('spares tab: MISSING — run setupSparesTab()');
  }
  var open = loanerList_().openCount;
  Logger.log(open + ' loan(s) currently out');
}

// ---- Data validation --------------------------------------------------------
// One button on the Reports tab answering "is the data between these books
// clean?" -- and if not, exactly which rows are wrong.
//
// compareBooks() has done most of this since the migration, but it only ever
// wrote to Logger, so the dashboard could not read it. This is the same
// comparison returned as JSON the page can group, act on and export, plus the
// checks that only start to matter once the app itself is writing notes.
//
// It reads BOTH workbooks end to end -- roughly 90 tab reads. That makes it a
// manual button with a spinner, never something on the page-load path. One
// getValues() per tab does all the work for that tab; adding a second read per
// tab is what has caused every timeout in this project so far.

var VALIDATE_CACHE_KEY = 'validationReport';
var VALIDATE_IGNORE_KEY = 'validationIgnored';
var VALIDATE_GROUP_CAP = 50;          // findings shown per group before "and N more"

// The spares / loaner pool tab. It lives in the check workbook only, by design,
// so its devices must never be reported as missing from the assignment roster.
var SPARES_TAB_NAME = 'Spares';

// The loaner ledger. Not created yet -- every check below that touches it is
// skipped when the tab is absent, so this file is safe to paste before the tab
// exists.
var LOANER_SHEET_NAME = 'Loaners';

function isSparesTab_(name) {
  return String(name || '').trim().toLowerCase() === SPARES_TAB_NAME.toLowerCase();
}

// Findings the user has waved off. Keyed by group + text, so the same oddity
// stays quiet across runs but a NEW instance of the same kind still reports.
function validateIgnored_() {
  var map = {};
  var raw = PropertiesService.getScriptProperties().getProperty(VALIDATE_IGNORE_KEY);
  if (!raw) return map;
  try {
    var list = JSON.parse(raw);
    for (var i = 0; i < list.length; i++) map[list[i]] = true;
  } catch (e) {
    // A corrupt property must not stop a validation run.
  }
  return map;
}

function validateIgnore_(p) {
  var key = String(p.key || '').trim();
  if (!key) return { ok: false, error: 'no key given' };
  var props = PropertiesService.getScriptProperties();
  var list = [];
  var raw = props.getProperty(VALIDATE_IGNORE_KEY);
  if (raw) {
    try { list = JSON.parse(raw); } catch (e) { list = []; }
  }
  if (String(p.undo) === '1') {
    var keep = [];
    for (var i = 0; i < list.length; i++) {
      if (list[i] !== key) keep.push(list[i]);
    }
    list = keep;
  } else if (list.indexOf(key) < 0) {
    list.push(key);
  }
  // A script property tops out near 9KB. Drop the oldest rather than fail.
  while (JSON.stringify(list).length > 8000) list.shift();
  props.setProperty(VALIDATE_IGNORE_KEY, JSON.stringify(list));
  return { ok: true, ignored: list.length };
}

// Run from the editor when the ignore list has collected things that should be
// reported again.
function clearValidationIgnores() {
  PropertiesService.getScriptProperties().deleteProperty(VALIDATE_IGNORE_KEY);
  Logger.log('validation ignore list cleared');
}

// Something that was probably meant to be a serial but is not one: it holds
// both letters and digits, so it is not a label like "Serial #" or a student
// name, yet it fails the serial test. A space in the middle, a stray period, a
// pasted URL. Anything that fails BOTH tests is almost certainly not a serial
// cell at all, and reporting it would bury the real problems.
function looksAttemptedSerial_(value) {
  var s = String(value == null ? '' : value).trim();
  if (s.length < 5) return false;
  if (!/[0-9]/.test(s)) return false;
  if (!/[A-Za-z]/.test(s)) return false;
  if (looksLikeRosterSerial_(s)) return false;
  return true;
}

// One pass over a workbook. Every serial, every spelling of those serials, the
// cells that look like a failed serial, and -- when asked -- the note cells that
// will never reach the to-do list. All of it from a single getValues() per tab.
function validateScan_(id, label, scanNotes) {
  var out = { label: label, tabs: [], byserial: {}, variants: {},
              bad: [], notes: [], loanerNotes: [], error: '' };
  var ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { out.error = String(e); return out; }

  rosterHeadReset_();
  var tabs = ss.getSheets();
  for (var i = 0; i < tabs.length; i++) {
    var sh = tabs[i];
    var name = sh.getName();
    if (isNonRosterTab_(name)) continue;
    if (name === LOANER_SHEET_NAME) continue;

    var info = rosterHeadInfo_(sh);
    if (!info.serialCol) continue;
    var dataRow = info.row + 1;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < dataRow || !lastCol) continue;

    out.tabs.push(name);
    var vals = sh.getRange(dataRow, 1, lastRow - dataRow + 1, lastCol).getValues();
    var noteCols = [];
    var loanerCol = 0;
    if (scanNotes) {
      noteCols = rosterNoteColumns_(sh);
      loanerCol = loanerNoteColumn_(sh, false);   // false: never create one while reading
    }

    for (var r = 0; r < vals.length; r++) {
      var row = vals[r];
      var rawSn = row[info.serialCol - 1];

      // A date sitting in the serial column stringifies into something with
      // letters and digits in it. It is wrong, but it is not a mistyped serial,
      // and flagging it fills the report with noise from second tables whose
      // columns do not line up with the first.
      var isDate = (Object.prototype.toString.call(rawSn) === '[object Date]');
      var sn = String(rawSn == null ? '' : rawSn).trim();

      if (sn && !isDate) {
        if (looksLikeRosterSerial_(sn)) {
          var key = sn.toUpperCase();
          if (!out.byserial[key]) out.byserial[key] = [];
          if (out.byserial[key].indexOf(name) < 0) out.byserial[key].push(name);

          // Every spelling seen, so "5cd1234abc " and "5CD1234ABC" can be
          // reported as one device typed two ways. A trailing space is
          // invisible in Sheets and breaks every match in this script.
          var spelling = String(rawSn);
          if (!out.variants[key]) out.variants[key] = [];
          if (out.variants[key].indexOf(spelling) < 0) out.variants[key].push(spelling);
        } else if (looksAttemptedSerial_(sn)) {
          out.bad.push({ tab: name, row: dataRow + r, value: sn });
        }
      }

      // Note cells that will never become a to-do. Only text counts: ticked
      // checkboxes, counts and dates are supposed to be ignored, and the junk
      // words ("ok", "na", "x") are deliberate shorthand rather than mistakes.
      for (var n = 0; n < noteCols.length; n++) {
        var cell = row[noteCols[n] - 1];
        if (typeof cell !== 'string') continue;
        var text = cell.trim();
        if (!text) continue;
        if (NOTE_JUNK_RE.test(text)) continue;
        if (isRealNote_(text)) continue;
        out.notes.push({ tab: name, row: dataRow + r, col: noteCols[n], value: text });
      }

      // Loaner notes are collected whole, not judged. Section 8 checks them
      // against the ledger, which is the only thing that knows what is out.
      if (loanerCol) {
        var lnv = String(row[loanerCol - 1] || '').trim();
        if (lnv) out.loanerNotes.push({ tab: name, row: dataRow + r, value: lnv });
      }
    }
  }
  return out;
}

// Why a note cell was passed over, in the words the report shows.
// The set of carts a device sits on, normalised and order-independent, so two
// books that file it the same way compare equal however their tabs are ordered
// or prefixed. normCart_ already folds "C-Cart F" and "iPad's" onto "cartf" and
// "ipads"; this is what makes the comparison about filing rather than ordering.
function cartSetKey_(tabs) {
  var norm = [];
  for (var i = 0; i < tabs.length; i++) {
    var n = normCart_(tabs[i]);
    if (norm.indexOf(n) < 0) norm.push(n);
  }
  norm.sort();
  return norm.join('+');
}

function noteRejectReason_(text) {
  var s = String(text || '').trim();
  if (s.length < 3) return 'under 3 characters';
  if (!/[a-z]{2}/i.test(s)) return 'no two letters together, so it does not read as words';
  return 'does not read as a note';
}

function validateData_(p) {
  var began = new Date().getTime();
  var ignored = validateIgnored_();
  var findings = [];

  // Findings are capped per group for display, but the counts on the tiles must
  // be the real numbers -- the first run reported "103 blockers" while the list
  // itself said "and 140 more", which is worse than either number alone.
  var tally = { blocker: 0, warn: 0, info: 0 };

  function add(sev, group, text, extra, silent) {
    var key = group + '|' + text;
    if (ignored[key]) return;
    if (!silent) tally[sev]++;
    var f = { sev: sev, group: group, text: text, key: key };
    if (extra) {
      if (extra.book) f.book = extra.book;
      if (extra.tab) f.tab = extra.tab;
      if (extra.row) f.row = extra.row;
      if (extra.fix) f.fix = extra.fix;
    }
    findings.push(f);
  }

  // Adds a whole list under one group, capped so a single bad import cannot
  // return ten thousand rows to the browser.
  function addList(sev, group, list, extra) {
    var shown = list.length;
    if (shown > VALIDATE_GROUP_CAP) shown = VALIDATE_GROUP_CAP;
    for (var i = 0; i < shown; i++) add(sev, group, list[i].text, list[i].at || extra);
    if (list.length > shown) {
      // The marker line is not itself a finding, so it is added silently and the
      // rows it stands for are tallied instead.
      add(sev, group, '… and ' + (list.length - shown) + ' more', extra, true);
      tally[sev] += (list.length - shown);
    }
  }

  var A = validateScan_(ASSIGNMENT_SHEET_ID, 'assignment roster', false);
  if (A.error) return { ok: false, error: 'assignment roster: ' + A.error };
  var C = validateScan_(ROSTER_SHEET_ID, 'check workbook', true);
  if (C.error) return { ok: false, error: 'check workbook: ' + C.error };

  // Serials that only ever appear on the spares tab belong to one book on
  // purpose, and are exempt from the cross-book comparison below.
  var sparesOnly = {};
  for (var sp in C.byserial) {
    var onlySpares = true;
    for (var t = 0; t < C.byserial[sp].length; t++) {
      if (!isSparesTab_(C.byserial[sp][t])) { onlySpares = false; break; }
    }
    if (onlySpares) sparesOnly[sp] = true;
  }

  // ---- 1. cells that look like a broken serial ----
  var badList = [];
  A.bad.forEach(function (b) {
    badList.push({ text: 'assignment · ' + b.tab + ' row ' + b.row + '   "' + b.value + '"',
                   at: { book: 'assignment', tab: b.tab, row: b.row } });
  });
  C.bad.forEach(function (b) {
    badList.push({ text: 'check · ' + b.tab + ' row ' + b.row + '   "' + b.value + '"',
                   at: { book: 'check', tab: b.tab, row: b.row } });
  });
  addList('blocker', 'Cell in the serial column is not a valid serial', badList);

  // ---- 2. one serial typed two different ways ----
  var variantList = [];
  function collectVariants(scan, bookLabel) {
    for (var k in scan.variants) {
      if (scan.variants[k].length < 2) continue;
      var quoted = [];
      for (var v = 0; v < scan.variants[k].length; v++) quoted.push('"' + scan.variants[k][v] + '"');
      variantList.push({ text: bookLabel + ' · ' + k + '   typed as ' + quoted.join(' and '),
                         at: { book: bookLabel } });
    }
  }
  collectVariants(A, 'assignment');
  collectVariants(C, 'check');
  addList('blocker', 'Same serial spelled more than one way', variantList);

  // ---- 3. a device listed on a lot of tabs inside one book ----
  //
  // Being on TWO tabs is normal in this data and always has been -- the pooled
  // carts (E, W) and the per-teacher iPad lists deliberately repeat devices that
  // also sit on a home cart, which is exactly why the to-do page shows the notes
  // each tab carries rather than deduping across tabs. Flagging that produced
  // 191 "blockers" describing the filing system.
  //
  // Three or more tabs is a different story, so that is what this reports, and
  // as a notice. Change 'info' to 'blocker' and the floor from 2 to 1 to go back
  // to reporting every repeat.
  var manyTabs = [];
  function collectManyTabs(scan, bookLabel) {
    for (var k in scan.byserial) {
      if (scan.byserial[k].length < 3) continue;
      manyTabs.push({ text: bookLabel + ' · ' + k + '   ' + scan.byserial[k].join(' | '),
                      at: { book: bookLabel, tab: scan.byserial[k][0] } });
    }
  }
  collectManyTabs(A, 'assignment');
  collectManyTabs(C, 'check');
  addList('info', 'Device listed on three or more tabs in one book', manyTabs);

  // ---- 4. the same device filed differently in the two books ----
  var movedList = [], missingFromC = [], missingFromA = [];
  for (var snA in A.byserial) {
    if (!C.byserial[snA]) {
      missingFromC.push({ text: snA + '   assignment: ' + A.byserial[snA].join(', '),
                          at: { book: 'assignment', tab: A.byserial[snA][0] } });
      continue;
    }
    // Compare the whole SET of tabs, not the first one in each book.
    //
    // The books list their tabs in their own order, so "Cart W / Cart F" and
    // "C-Cart F / C-Cart W" are the same filing read two different ways --
    // comparing element [0] of each called every one of those a mismatch, and
    // that alone accounted for most of the 86 findings in this group.
    if (cartSetKey_(A.byserial[snA]) !== cartSetKey_(C.byserial[snA])) {
      movedList.push({ text: snA + '   assignment: ' + A.byserial[snA].join('/') +
                             '   check: ' + C.byserial[snA].join('/'),
                       at: { book: 'assignment', tab: A.byserial[snA][0] } });
    }
  }
  for (var snC in C.byserial) {
    if (A.byserial[snC]) continue;
    if (sparesOnly[snC]) continue;
    missingFromA.push({ text: snC + '   check: ' + C.byserial[snC].join(', '),
                        at: { book: 'check', tab: C.byserial[snC][0] } });
  }
  addList('blocker', 'On different carts in the two books', movedList);
  addList('warn', 'In the assignment roster but not the check workbook', missingFromC);
  addList('warn', 'In the check workbook but not the assignment roster', missingFromA);

  // ---- 5. tabs that exist in one book only ----
  // Both sides filtered the same way, and only for Spares -- which genuinely
  // lives in one book.
  //
  // This used to drop TODO_SKIP_TABS from the check side only, on the wrong
  // assumption that the per-teacher lists were check-book-only. They are in both
  // books, so filtering one side reported Crider, Palsa, Aeh and the rest as
  // "assignment only" when they match perfectly. Filtering neither side also
  // means a genuine naming split -- "B Miller" against "Buechner" -- reports for
  // the right reason instead of being hidden by the skip list.
  var an = {}, cn = {};
  function collectTabKeys(tabs, into) {
    for (var i = 0; i < tabs.length; i++) {
      if (isSparesTab_(tabs[i])) continue;
      into[normCart_(tabs[i])] = tabs[i];
    }
  }
  collectTabKeys(A.tabs, an);
  collectTabKeys(C.tabs, cn);
  var tabOnly = [];
  for (var ka in an) {
    if (!cn[ka]) tabOnly.push({ text: 'assignment only · ' + an[ka], at: { book: 'assignment', tab: an[ka] } });
  }
  for (var kc in cn) {
    if (!an[kc]) tabOnly.push({ text: 'check only · ' + cn[kc], at: { book: 'check', tab: cn[kc] } });
  }
  addList('warn', 'Cart tab exists in one book only', tabOnly);

  // ---- 6. serials stored on tickets ----
  var urlRows = [];
  try {
    var bad = badSerialRows_();
    bad.forEach(function (b) {
      urlRows.push({ text: b.tab + ' row ' + b.row + '   "' + b.was + '"  →  ' + b.now,
                     at: { book: 'tickets', tab: b.tab, row: b.row, fix: 'serials' } });
    });
  } catch (e) {
    add('warn', 'Could not read the ticket book', String(e));
  }
  addList('warn', 'Ticket serial stored as a URL or label', urlRows, { fix: 'serials' });

  var unknownTicket = [];
  try {
    var live = firstSheet_();
    var lastT = live.getLastRow();
    if (lastT > 1) {
      var col = live.getRange(2, 1, lastT - 1, 1).getValues();
      var seenT = {};
      for (var tr = 0; tr < col.length; tr++) {
        var ts = String(col[tr][0] || '').trim().toUpperCase();
        if (!ts) continue;
        // Not every ticket is about a device. A "not a Chromebook or iPad"
        // ticket carries a teacher and room here, which is not a serial and must
        // never be reported as one -- otherwise every projector ticket becomes a
        // permanent warning. Anything not shaped like a serial is skipped; a
        // mistyped serial still IS serial-shaped, so real typos still report.
        if (!looksLikeRosterSerial_(ts)) continue;
        if (seenT[ts]) continue;
        seenT[ts] = true;
        if (A.byserial[ts] || C.byserial[ts]) continue;
        unknownTicket.push({ text: ts + '   (ticket sheet row ' + (tr + 2) + ')',
                             at: { book: 'tickets', tab: live.getName(), row: tr + 2 } });
      }
    }
  } catch (e2) {
    add('warn', 'Could not read ticket serials', String(e2));
  }
  addList('warn', 'Open ticket for a serial that is in neither book', unknownTicket);

  // ---- 7. the to-do list against the check workbook ----
  try {
    var todos = todoList_().todos;
    var cartNames = {};
    C.tabs.forEach(function (n) { cartNames[n] = true; });
    var badCart = [], badSn = [];
    for (var i2 = 0; i2 < todos.length; i2++) {
      var td = todos[i2];
      if (String(td.id).indexOf('roster-') !== 0) continue;
      if (td.group && !cartNames[td.group]) {
        badCart.push({ text: td.group + '   ::   ' + td.text });
      }
      var m = String(td.text).match(/S\/N\s+([A-Za-z0-9]+)\)/);
      if (m && !C.byserial[m[1].toUpperCase()]) {
        badSn.push({ text: m[1] + '   ::   ' + td.text });
      }
    }
    addList('warn', 'To-do sitting on a cart that no longer exists', badCart);
    addList('warn', 'To-do whose serial is not in the check workbook', badSn);
  } catch (e3) {
    add('warn', 'Could not read the to-do list', String(e3));
  }

  // ---- 8. the loaner ledger against the notes it is supposed to have written ----
  // The ledger is the record; the notes on the devices are its mirror. So every
  // check here reads one way -- ledger first, notes second -- and each finding
  // is fixable by re-syncing the notes rather than by editing them.
  try {
    var ledger = ticketBook_().getSheetByName(LOANER_SHEET_NAME);
    if (ledger && ledger.getLastRow() > 1) {
      var loans = loanerList_().loans;
      var openBy = {}, dupOpen = [], badDates = [], unknownLoan = [], missingNote = [];
      var expected = {};        // every note text the ledger says should exist

      // Every loaner note actually on a device, gathered once by validateScan_.
      var found = {};
      for (var fn = 0; fn < C.loanerNotes.length; fn++) found[C.loanerNotes[fn].value] = true;

      for (var li = 0; li < loans.length; li++) {
        var L = loans[li];
        var inopKey = L.inopSn.toUpperCase();
        var loanKey = L.loanerSn.toUpperCase();

        if (L.dateOut && L.dateIn && new Date(L.dateIn) < new Date(L.dateOut)) {
          badDates.push({ text: 'row ' + L.row + '   out ' + loanerShortDate_(L.dateOut) +
                                ', back ' + loanerShortDate_(L.dateIn),
                          at: { book: 'loaners', row: L.row } });
        }

        if (!L.open) continue;

        if (loanKey) {
          if (openBy[loanKey]) {
            dupOpen.push({ text: L.loanerSn + '   open on rows ' + openBy[loanKey] + ' and ' + L.row,
                           at: { book: 'loaners', row: L.row } });
          } else {
            openBy[loanKey] = L.row;
          }
        }
        if (inopKey && !A.byserial[inopKey] && !C.byserial[inopKey]) {
          unknownLoan.push({ text: L.inopSn + '   broken device, Loaners row ' + L.row,
                             at: { book: 'loaners', row: L.row } });
        }
        if (loanKey && !A.byserial[loanKey] && !C.byserial[loanKey]) {
          unknownLoan.push({ text: L.loanerSn + '   loaner, Loaners row ' + L.row,
                             at: { book: 'loaners', row: L.row } });
        }

        // The note this row implies on the device that went out. Registered as
        // expected for EVERY open row, loan or relocation, so a relocation note
        // sitting on a device is recognised rather than reported as an orphan.
        // Kyle keeps those deliberately -- seeing where a device came from is
        // useful on a cart tab.
        var wantOut = loanerDeviceNote_(L);
        expected[wantOut] = true;

        // A row with no Inop S/N is a relocation, not a loan: the device simply
        // lives in another cart now, nothing is broken, and there is no pair to
        // annotate. Its note is legitimate but not required, so nothing further
        // is checked -- reporting a missing one would be 60 findings about
        // devices that are exactly where they should be.
        //
        // Inop S/N is the field that separates the two, and it is the same test
        // loanerResync_ uses before writing anything.
        if (!L.inopSn) continue;

        // A real loan: both notes are expected to be there, and their absence is
        // drift worth reporting.
        var wantInop = loanerInopNote_(L.loanerSn, L.dateOut, L.source);
        expected[wantInop] = true;

        if (!found[wantInop]) {
          missingNote.push({ text: L.inopSn + '   row ' + L.row + '   expected: "' + wantInop + '"',
                             at: { book: 'loaners', row: L.row, fix: 'resync' } });
        }
        if (!found[wantOut]) {
          missingNote.push({ text: L.loanerSn + '   row ' + L.row + '   expected: "' + wantOut + '"',
                             at: { book: 'loaners', row: L.row, fix: 'resync' } });
        }
      }

      // The other direction: a loaner note on a device the ledger says nothing
      // about. Usually a return that got half done by hand.
      var orphanNote = [];
      for (var on = 0; on < C.loanerNotes.length; on++) {
        var note = C.loanerNotes[on];
        if (expected[note.value]) continue;
        orphanNote.push({ text: note.tab + ' row ' + note.row + '   "' + note.value + '"',
                          at: { book: 'check', tab: note.tab, row: note.row, fix: 'resync' } });
      }

      addList('blocker', 'Same loaner out on two open rows', dupOpen);
      addList('blocker', 'Loaner returned before it went out', badDates);
      addList('warn', 'Loan is open but its note is missing from the device', missingNote, { fix: 'resync' });
      addList('warn', 'Loaner note on a device with no open loan behind it', orphanNote, { fix: 'resync' });
      addList('warn', 'Loaner serial not found in either book', unknownLoan);
    }
  } catch (e4) {
    add('warn', 'Could not read the loaner ledger', String(e4));
  }

  // ---- 9. notes that will never become a to-do ----
  var noteList = [];
  C.notes.forEach(function (n) {
    noteList.push({ text: n.tab + ' row ' + n.row + '   "' + n.value + '"   — ' + noteRejectReason_(n.value),
                    at: { book: 'check', tab: n.tab, row: n.row } });
  });
  addList('info', 'Note that will never become a to-do', noteList);

  var out = {
    ok: true,
    runAt: new Date().toISOString(),
    seconds: Math.round((new Date().getTime() - began) / 1000),
    counts: tally,
    findings: findings,
    scanned: {
      assignmentTabs: A.tabs.length, assignmentSerials: Object.keys(A.byserial).length,
      checkTabs: C.tabs.length, checkSerials: Object.keys(C.byserial).length
    }
  };

  // Cached so re-opening the Reports tab shows the last run rather than paying
  // 90 tab reads again. Cache values top out near 100KB; an oversized report
  // simply is not cached.
  try {
    var json = JSON.stringify(out);
    if (json.length < 90000) CacheService.getScriptCache().put(VALIDATE_CACHE_KEY, json, 21600);
  } catch (e5) {
    // Caching is a convenience; never fail the run over it.
  }
  return out;
}

// The last run, if it is still cached. Lets the page show a report on open
// without ever triggering the walk itself.
function validateLast_() {
  var json = null;
  try { json = CacheService.getScriptCache().get(VALIDATE_CACHE_KEY); } catch (e) { json = null; }
  if (!json) return { ok: true, empty: true };
  try { return JSON.parse(json); } catch (e2) { return { ok: true, empty: true }; }
}

// The one fix the validator can apply itself: rewriting ticket serials that were
// stored as a URL or a "S/N: ..." label. fixStoredSerials() already did this
// from the editor; this is the same call with a return value for the page.
function validateFixSerials_() {
  try {
    var n = fixStoredSerials();
    return { ok: true, fixed: n };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// A ticket is about a Chromebook or iPad when its S/N column actually holds a
// serial. With the help desk's "not a device" box ticked that column holds a
// teacher name and room instead -- which looksLikeRosterSerial_ rejects, since
// it has spaces in it -- so the shape of the value is a reliable test on its
// own. The flag the form sends is honoured first, so stated intent wins over
// anything inferred.
function isDeviceTicket_(sn, nonDeviceFlag) {
  if (String(nonDeviceFlag || '') === '1') return false;
  return looksLikeRosterSerial_(sn);
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
      // The S/N column can now hold "Messer · Rm 13" rather than a serial, so the
      // filename is scrubbed of anything that does not belong in one.
      var tag = String(data.sn || 'unknown').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
      photoUrl = savePhoto_(data.photo, 'CB_' + (tag || 'unknown') + '_ticket' + no + '.jpg');
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

    // Everything below annotates a DEVICE. A ticket filed with the help desk's
    // "not a Chromebook or iPad" box ticked has a teacher and room in the S/N
    // column instead of a serial, so there is no device row to write a note on
    // and no cart to group a to-do under. The To-Do page is the history of the
    // Chromebooks and iPads; a projector does not belong in it.
    if (isDeviceTicket_(data.sn, data.nonDevice)) {
      // put what was typed into "Describe the problem" on the device's row in the
      // roster workbook; skipped silently if the serial is not on a tab
      var note = rosterNoteWrite_(data.sn, no, data.description);
      // put the same text on the to-do list, under the cart the device is in
      var where = null;
      if (note && note.tab) where = { cart: note.tab, cbNo: note.cbNo };
      ticketTodoWrite_(data.sn, no, data.description, where);
    }
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
// Every code a scan could plausibly be, longest first. A label often carries
// two -- Acer prints a short SNID beside the serial, Lenovo an MTM -- and length
// alone cannot always tell which is the device. The books can: whoever calls
// this may check the candidates against the S/N column and use the one that is
// really there, falling back to the first (longest) when the books cannot say.
function serialCandidates_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return [];

  // Longest first, duplicates dropped. Sort is stable, so codes of equal length
  // stay in the order they were read.
  function byLongest(list) {
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var v = String(list[i] == null ? '' : list[i]).trim();
      if (v && out.indexOf(v) < 0) out.push(v);
    }
    out.sort(function (a, b) { return b.length - a.length; });
    return out;
  }

  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(s)) {
    // A URL. Collect every query parameter that NAMES a serial. Selecting by
    // name first is what keeps mtm out: it is the model, and on a Lenovo link it
    // is longer than the serial beside it, so comparing lengths across the whole
    // query string would pick the wrong one.
    var found = [];
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
        if (val) found.push(val);
      }
    }
    if (found.length) return byLongest(found);

    // No serial parameter -- some makers put it in the last path segment.
    var path = s.split(/[?#]/)[0].replace(/\/+$/, '');
    var seg = path.slice(path.lastIndexOf('/') + 1);
    if (looksLikeRosterSerial_(seg)) return [seg];
    return [s];           // nothing recognisable; keep what we read rather than
                          // mangle it into something wrong
  }

  // Not a URL. Some labels write "S/N: ABC123" rather than the bare value, and
  // some print two labelled codes.
  //
  // Two patterns, because the colon is only safe to make optional after a label
  // that cannot be the start of another word. "SN" cannot: "SNID 12345678" would
  // match SN and capture "ID". So bare SN still requires its colon.
  var labelled = [];
  var labelRes = [
    /\b(?:s\/n|serial(?:\s*(?:no|number))?|service\s*tag)\b\s*[:=]?\s*([A-Za-z0-9-]+)/gi,
    /\bsn\s*[:=]\s*([A-Za-z0-9-]+)/gi
  ];
  for (var r = 0; r < labelRes.length; r++) {
    var m = labelRes[r].exec(s);
    while (m) {
      if (m[1]) labelled.push(m[1]);
      m = labelRes[r].exec(s);
    }
  }
  if (labelled.length) return byLongest(labelled);

  // Nothing carries a label. If the scan is plainly two or more codes -- on
  // separate lines, or spaces between them -- offer them longest first. Model
  // and part numbers are dropped, for the same reason as above: an MTM like
  // 83T60009US outruns the serial it sits next to.
  function bareTokens(text) {
    var tokens = text.split(/[\s,;|=:]+/);
    var out = [];
    for (var j = 0; j < tokens.length; j++) {
      if (looksLikeRosterSerial_(tokens[j])) out.push(tokens[j]);
    }
    return out;
  }

  var all = bareTokens(s);
  if (all.length > 1) {
    var kept = bareTokens(s.replace(
      /\b(?:mtm|model|type|p\/?n|part(?:\s*(?:no|number))?|product|sku)\s*[:=]?\s*[A-Za-z0-9-]+/gi, ' '));
    if (kept.length) return byLongest(kept);
    return byLongest(all);
  }

  // One code, or none we recognise.
  return [s];
}

// The single best guess, for every caller that just wants a value. Same answer
// serialFromScan has always given, now expressed as "the first candidate".
function serialFromScan_(raw) {
  var list = serialCandidates_(raw);
  if (!list.length) return '';
  return list[0];
}

// ---- Lining a scanned code up with the S/N column ---------------------------
// Exact whole-cell first, which is all the books ever needed. Only when that
// misses do we line the two up by their TAILS: a factory QR often carries a
// prefix the sheet never had, so scan and cell agree from some point onwards.
// Matching on the tail rather than the head matters -- Chromebooks from one
// purchase share a long prefix and differ at the end, so the head is exactly
// the part that does not tell them apart.
//
// Eight characters is the floor: the shortest real serial in these books is
// Lenovo's 8, and a shorter overlap starts colliding across 1,200-odd devices.
// A tail that lands on two different devices is REFUSED, not guessed at --
// hanging a ticket on the wrong Chromebook is worse than asking the teacher to
// read the label again.
var SERIAL_TAIL_MIN = 8;
var SERIAL_INDEX_CACHE_KEY = 'serialTailIndex_v1';
var serialIndexMemo_ = null;

function serialTail_(sn) {
  var s = String(sn || '').toUpperCase();
  if (s.length < SERIAL_TAIL_MIN) return '';
  return s.slice(-SERIAL_TAIL_MIN);
}

// { exact: {SERIAL: true}, tails: {TAIL: [SERIAL, ...]} } over the check
// workbook's serial columns. Building it walks every roster tab, so it is
// memoised per execution and cached for six hours. The exact path below never
// builds it -- only a scan that misses outright pays for this.
function serialIndex_() {
  if (serialIndexMemo_) return serialIndexMemo_;
  var cached = null;
  try { cached = CacheService.getScriptCache().get(SERIAL_INDEX_CACHE_KEY); } catch (e) { cached = null; }
  if (cached) {
    try {
      serialIndexMemo_ = JSON.parse(cached);
      return serialIndexMemo_;
    } catch (e) {}
  }
  var map = serialMap_(NOTE_BOOK_ID, false);
  var idx = { exact: {}, tails: {} };
  for (var sn in map.byserial) {
    idx.exact[sn] = true;
    var t = serialTail_(sn);
    if (!t) continue;
    if (!idx.tails[t]) idx.tails[t] = [];
    if (idx.tails[t].indexOf(sn) < 0) idx.tails[t].push(sn);
  }
  try {
    var json = JSON.stringify(idx);
    if (json.length < 90000) CacheService.getScriptCache().put(SERIAL_INDEX_CACHE_KEY, json, 21600);
  } catch (e) {}
  serialIndexMemo_ = idx;
  return idx;
}

// Throw the cached index away. Worth calling after anything that rewrites
// serials in the books, so a scan is not matched against a stale list.
function serialIndexReset_() {
  serialIndexMemo_ = null;
  try { CacheService.getScriptCache().remove(SERIAL_INDEX_CACHE_KEY); } catch (e) {}
}

// One scanned code -> the serial the books actually hold.
// { sn: <serial>, how: 'exact'|'tail' }, or { ambiguous: true }, or null.
function rosterResolveOne_(scan) {
  var up = String(scan || '').trim().toUpperCase();
  if (!up) return null;

  // One optimised whole-cell search, no index built.
  try {
    if (rosterFindRows_(up).length) return { sn: up, how: 'exact' };
  } catch (e) {
    return null;                          // the caller treats an error as "do not block"
  }

  var tail = serialTail_(up);
  if (!tail) return null;

  var idx = serialIndex_();
  if (idx.exact[up]) return { sn: up, how: 'exact' };

  var list = idx.tails[tail] || [];
  var hits = [];
  for (var i = 0; i < list.length; i++) {
    var cand = list[i];
    // "Lining up" means one is the tail of the other, in either direction:
    // the scan may carry a prefix the sheet lacks, or the reverse.
    if (up.length >= cand.length) {
      if (up.slice(up.length - cand.length) === cand) hits.push(cand);
    } else {
      if (cand.slice(cand.length - up.length) === up) hits.push(cand);
    }
  }
  if (hits.length > 1) return { ambiguous: true };
  if (hits.length === 1) {
    // The index can be up to six hours old, so confirm the winner is still in
    // the books before handing it back. New devices never need this -- they are
    // found by the live search above -- but a serial that has since been
    // removed or corrected would otherwise keep being matched from cache.
    try {
      if (!rosterFindRows_(hits[0]).length) {
        serialIndexReset_();
        return null;
      }
    } catch (e) {
      return null;
    }
    return { sn: hits[0], how: 'tail' };
  }
  return null;
}

function rosterCartFor_(sn) {
  try {
    var hits = rosterFindRows_(sn);
    if (!hits.length) return '';
    return rosterPickBest_(hits).sheet.getName();
  } catch (e) { return ''; }
}

// Public (no token). Takes the RAW scan, not a serial: the server derives the
// candidates itself so the page and the script cannot drift apart on which
// codes a label offers.
//
// The books decide. Length only breaks a tie they cannot: when two of the
// scanned codes are both real devices, or when none of them are.
function snResolve_(p) {
  var raw = String(p.sn == null ? '' : p.sn);
  var list = serialCandidates_(raw);
  if (!list.length) return { ok: true, found: false, sn: '' };

  var ambiguous = false;
  var found = [];
  for (var i = 0; i < list.length; i++) {
    var hit = rosterResolveOne_(list[i]);
    if (!hit) continue;
    if (hit.ambiguous) { ambiguous = true; continue; }
    found.push(hit);
  }

  if (found.length) {
    // `list` came in longest first and `found` kept that order, so found[0] is
    // the longest of the codes that are genuinely in the books.
    var how = found[0].how;
    if (found.length > 1) how = 'longest';
    return { ok: true, found: true, sn: found[0].sn, how: how,
             cart: rosterCartFor_(found[0].sn), candidates: list.length };
  }

  // Nothing matched. Hand back the same guess as before so the form still has a
  // value to work with, and say why it is only a guess.
  return { ok: true, found: false, sn: list[0], ambiguous: ambiguous, candidates: list.length };
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
  var ss = ticketBook_();
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
  var ss = ticketBook_();
  var live = firstSheet_().getName();
  var tabs = ss.getSheets();
  for (var t = 0; t < tabs.length; t++) {
    var sh = tabs[t];
    var name = sh.getName();
    // Ticket data only: the live sheet and the monthly archives. Not the Todos
    // tab, and not the roster workbook.
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
