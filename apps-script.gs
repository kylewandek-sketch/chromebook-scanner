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

// ---- To-Do list (dashboard "To-Do" tab) ----
// Items live in a 'Todos' sheet tab: ID | Text | Done | Order | Created | Group.
// Group is the cart/section the task belongs to (e.g. "Cart O"); the dashboard
// shows each group as a collapsible section. Blank group shows as "General".
var TODO_SHEET_NAME = 'Todos';
var TODO_HEADERS = ['ID', 'Text', 'Done', 'Order', 'Created', 'Group'];

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
    todos.push({
      id: String(r[0]), text: String(r[1]), done: done,
      order: Number(r[3]) || 0, group: String(r[5] || '')
    });
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
  sh.appendRow([id, text, false, order, new Date(), String(p.group || '').trim()]);
  return { ok: true, id: id };
}

function todoUpdate_(p) {
  var sh = todoSheet_();
  var row = todoFindRow_(sh, p.id);
  if (!row) return { ok: false, error: 'not found' };
  if (p.text != null) sh.getRange(row, 2).setValue(String(p.text));
  if (p.done != null) sh.getRange(row, 3).setValue(String(p.done) === 'true');
  if (p.group != null) sh.getRange(row, 6).setValue(String(p.group).trim());
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

// RUN THIS from the editor if your Todos tab still has the old ungrouped items:
// it DELETES the whole Todos tab (all items and checkmarks, including any you
// added yourself) and reloads the cross-reference findings grouped by cart.
function reseedTodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TODO_SHEET_NAME);
  if (sh) ss.deleteSheet(sh);
  return seedTodos();
}

// RUN THIS ONCE from the editor to load the summer 2026 check/repair cross-reference
// findings into the to-do list, grouped by cart. It refuses to run if the Todos tab
// already has items - use reseedTodos() to wipe and reload instead.
function seedTodos() {
  var sh = todoSheet_();
  if (sh.getLastRow() > 1) {
    Logger.log('Todos tab already has items - not seeding again.');
    return 'Todos tab already has items - not seeding again.';
  }
  // [group, task]
  var items = [
    ['Cart A', '#18 (NXHBNAA0019160FFC07600): in repair since 5/15, screen will not turn on - never returned, chase it'],
    ['Cart A', '(NXHBNAA0019160FE837600): in repair since 4/17, stuck on white screen - never returned, chase it'],
    ['Cart A', '#19: hinge loose / falls'],
    ['Cart A', '#3: charger dead + keys work intermittently'],
    ['Cart A', '#7: charger dead'],
    ['Cart A', 'Eduanny ESL iPad: replace missing charger (flagged both checks)'],
    ['Cart B', '#13 (NXH8VAA0060400FD467611): screen broken, in repair since 1/14 - never returned; headphone jack blocked all year'],
    ['Cart B', '#1: charger dead since spring break'],
    ['Cart B', '#23 (MP1M1ZXR): not working'],
    ['Cart B', '#3 (MP1M1ZYY): keys missing'],
    ['Cart B', '#7 (MP1HDZ13): keys missing'],
    ['Cart B', '#6 (MP1M1ZX2): number keys acting up again after 4/17 repair - verify'],
    ['Cart C', '#7: missing from cart - locate'],
    ['Cart C', '#13: missing from cart - locate'],
    ['Cart C', '#22: missing from cart - locate'],
    ['Cart C', '#12: charger dead'],
    ['Cart C', '#23: charger dead'],
    ['Cart C', '#26: charger dead + state-testing browser not working'],
    ['Cart C', '#27: charger dead'],
    ['Cart C', '#18: not working + state-testing browser not working'],
    ['Cart D', '(NXHBNAA0019252726A7600): 3 wifi repairs, same complaint each time - replace wifi card or retire'],
    ['Cart D', '(NXHBNAA001916101BF7600): came back not working after 4/10 repair, in again 4/24 - verify it is actually fixed'],
    ['Cart D', '#14: D key cap missing (key still works) - since spring break'],
    ['Cart G', '#25: state-testing browser not working'],
    ['Cart G', '#26: state-testing browser not working'],
    ['Cart G', '#27: state-testing browser not working'],
    ['Cart G', '#28: state-testing browser not working'],
    ['Cart H', '#16: repair log says returned 1/5 but start-of-year check says missing - reconcile'],
    ['Cart H', '#8: keys missing'],
    ['Cart H', '#12: keys missing'],
    ['Cart H', '#14: keys missing'],
    ['Cart H', '#15: keys missing'],
    ['Cart H', '#27: keys missing'],
    ['Cart H', '#10: hinge broken'],
    ['Cart H', '#19: hinge broken'],
    ['Cart I', '#4: marked returned from enrollment fix 1/5 but not in cart - locate'],
    ['Cart I', '#9: marked returned from enrollment fix 1/5 but not in cart - locate'],
    ['Cart J', '#11: charging port missing, charger must be plugged in backwards'],
    ['Cart J', '#28: screen scratched + no serial label'],
    ['Cart J', '#9: hinge super loose'],
    ['Cart K', '#26: missing / broken screen'],
    ['Cart K', '#1: power key missing'],
    ['Cart K', '#23: spacebar missing'],
    ['Cart K', '#8: hinge starting to dislocate'],
    ['Cart K', '#4: charger dead'],
    ['Cart K', '#18: charger dead'],
    ['Cart K', '(NXHBNAA00191610D527600): keyboard still dead after 2 repair visits - repair did not take'],
    ['Cart K', 'Adams Smartpass iPad: missing / not working'],
    ['Cart K', 'Karim Shabana iPad: missing / not working'],
    ['Cart N', '(G5LG0H3): 3 repair visits (trackpad x2, then screen) - consider retiring'],
    ['Cart O', 'Fix state-testing browser on 20 HP units (2HA99...) - likely enrollment issue, one unit noted not enrolled to @CPAohio.org'],
    ['Cart O', 'Fix duplicate serial - #6 and #25 both entered with matching serials (2HA99FEN501098M / 2HA99FEN511183W)'],
    ['Cart T', '#26: will not turn on - open ticket'],
    ['Cart V', '#12: state test app missing - reinstall'],
    ['Cart W', '#28: on loan to SPED - track it and get it back'],
    ['Cart X', '#1: keys missing'],
    ['Cart Y', '#5: not working + charger dead'],
    ['Cart Y', '#23: not working + charger dead'],
    ['Cart Y', '#27: not working + charger dead'],
    ['Cart Y', '#13: marked returned 1/5 but start-of-year check says missing - locate'],
    ['Cart Y', '#25: marked returned 1/5 but start-of-year check says missing - locate'],
    ['ESL', '#6: trackpad broken and taped shut - open ticket'],
    ['iPads', 'Ljungren (DMQSJ599HGSD): in repair since 9/3/25 - never returned, chase it'],
    ['iPads', 'Seggerson (DMQPH5ZZFK10): in repair since 9/22/25, dead backlight - never returned, chase it'],
    ['iPads', 'Wand (DMQPH9V4FK10): in repair since 9/22/25, battery EOL - never returned, chase it'],
    ['iPads', 'Caudill (DMQPHDLPFK10): in repair since 11/11/25, dead battery - never returned, chase it'],
    ['iPads', 'Buechner (DMPPDS94FK10): screen broken, in repair since 3/4/26 - never returned, chase it'],
    ['iPads', 'Ljungren Smartpass (DMQPHCTGFK10): screen broken, in repair since 2/27 - never returned; Smart Pass not working'],
    ['iPads', 'Caudill #95: swapped out - update roster (teacher now has 24 iPads)'],
    ['iPads', 'Miller #44: needs an extra charger'],
    ['Checks to finish', 'Re-check Cart AA: every box marked TRUE (checkmark treated as OK) - Keys Missing / Hinge Broken columns unreliable'],
    ['Checks to finish', 'Re-check Cart F: every box marked TRUE (checkmark treated as OK)'],
    ['Checks to finish', 'Re-check Cart Y: every box marked TRUE (checkmark treated as OK)'],
    ['Checks to finish', 'Start-of-year check untouched: BB (Moorman)'],
    ['Checks to finish', 'Start-of-year check untouched: Aeh iPads'],
    ['Checks to finish', 'Start-of-year check untouched: Perez iPads'],
    ['Checks to finish', 'Start-of-year check untouched: Moorman iPads'],
    ['Checks to finish', 'Start-of-year check untouched: Title 1 iPads'],
    ['Checks to finish', 'Start-of-year check untouched: Hunter iPad cart']
  ];
  var now = new Date();
  var base = now.getTime();
  var rows = [];
  for (var i = 0; i < items.length; i++) {
    rows.push([String(base + i), items[i][1], false, i + 1, now, items[i][0]]);
  }
  sh.getRange(2, 1, rows.length, TODO_HEADERS.length).setValues(rows);
  Logger.log('Seeded ' + rows.length + ' to-dos in cart groups.');
  return 'Seeded ' + rows.length + ' to-dos in cart groups.';
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
