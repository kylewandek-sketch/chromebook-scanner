/**
 * Chromebook Help Desk — Google Apps Script Web App
 *
 * Receives ticket submissions from the scanner page, appends a row to the bound
 * Google Sheet (row 1 = headers, column A = Chromebook S/N), and emails the help desk.
 *
 * SETUP
 *  1. Create a new Google Sheet (this will hold the tickets).
 *  2. Extensions ▸ Apps Script. Delete any sample code, paste this whole file, Save.
 *  3. Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as: Me
 *       - Who has access: "Anyone" (or "Anyone within CPA Ohio" if teachers are signed in to the domain)
 *  4. Authorize when prompted. Copy the Web app URL (ends with /exec).
 *  5. Send that URL back — it gets pasted into ticket.html as TICKET_ENDPOINT.
 */

var HELPDESK_EMAIL = 'kyle.anderson@cpaohio.org';

// Shared secret the dashboard must send to read/modify tickets. Change this to your own
// random string if you like — just update it in the dashboard the first time you sign in.
var ADMIN_TOKEN = 'de5c23a248ee2d5a66e65ec8';

// Column order. Column A is the Chromebook S/N, as required. (Notes is added as column 10.)
var HEADERS = [
  'Chromebook S/N', 'Timestamp', 'Teacher Email', 'Teacher Name',
  'Room #', 'Issue Type', 'Urgency', 'Description', 'Status'
];

// GET endpoint. Serves the management dashboard's data + edits (JSONP to avoid CORS).
//   ?action=list|update|delete  &token=...  &callback=...
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  if (['list', 'update'].indexOf(p.action) >= 0 && p.token !== ADMIN_TOKEN) {
    out = { ok: false, error: 'unauthorized' };
  } else if (p.action === 'list') {
    out = listTickets_();
  } else if (p.action === 'update') {
    out = updateTicket_(p);
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

function ensureNotesHeader_(sheet) {
  if (!sheet.getRange(1, 10).getValue()) sheet.getRange(1, 10).setValue('Notes');
}

function listTickets_() {
  var sheet = firstSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };
  var lastCol = Math.max(10, sheet.getLastColumn());
  var v = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows = v.map(function (r, i) {
    return {
      row: i + 2,
      sn: r[0], timestamp: r[1] ? new Date(r[1]).toISOString() : '',
      teacherEmail: r[2], teacherName: r[3], room: r[4],
      issue: r[5], urgency: r[6], description: r[7],
      status: r[8] || 'New', notes: r[9] || ''
    };
  });
  return { ok: true, rows: rows };
}

function updateTicket_(p) {
  var row = parseInt(p.row, 10);
  if (!row || row < 2) return { ok: false, error: 'bad row' };
  var sheet = firstSheet_();
  ensureNotesHeader_(sheet);
  if (p.status != null) sheet.getRange(row, 9).setValue(p.status);
  if (p.notes != null) sheet.getRange(row, 10).setValue(p.notes);
  return { ok: true };
}

// ---- Monthly archive ----
// Run setupMonthlyArchive() ONCE to install a trigger that fires on the 1st of each month.
// It moves all rows out of the live tickets sheet into a new tab named "[Mon][YY]_Tickets"
// for the PREVIOUS month (e.g. run Aug 1 -> "Jul26_Tickets"), leaving the live sheet empty
// (headers kept) for the new month.
var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function setupMonthlyArchive() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'archiveMonthly') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archiveMonthly').timeBased().onMonthDay(1).atHour(1).create();
}

function archiveMonthly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheets()[0];
  var lastRow = src.getLastRow();
  if (lastRow < 2) return;                      // nothing to archive
  var lastCol = Math.max(HEADERS.length, src.getLastColumn());

  var prev = new Date();
  prev = new Date(prev.getFullYear(), prev.getMonth() - 1, 1);   // first day of previous month
  var name = MONTHS[prev.getMonth()] + String(prev.getFullYear()).slice(-2) + '_Tickets';
  if (ss.getSheetByName(name)) {
    name += '_' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMddHHmm');
  }

  var dest = ss.insertSheet(name, ss.getNumSheets());
  var all = src.getRange(1, 1, lastRow, lastCol).getValues();    // include header row
  dest.getRange(1, 1, all.length, lastCol).setValues(all);
  dest.getRange(1, 1, 1, lastCol).setFontWeight('bold');
  dest.setFrozenRows(1);

  src.getRange(2, 1, lastRow - 1, lastCol).clearContent();       // empty the live sheet, keep headers
}

function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); }
      catch (err) { data = (e.parameter || {}); }   // fallback for form-encoded
    } else {
      data = (e && e.parameter) || {};
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // Ensure header row exists.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var now = new Date();
    var row = [
      data.sn || '',
      now,
      data.email || '',
      data.name || '',
      data.room || '',
      data.issue || '',
      data.urgency || '',
      data.description || '',
      'New'
    ];
    sheet.appendRow(row);

    sendEmail_(data, now);

    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function sendEmail_(data, now) {
  var subject = '[Help Desk] ' + (data.issue || 'Ticket') +
                ' — CB ' + (data.sn || '?') + ' (' + (data.urgency || 'Medium') + ')';

  var lines = [
    'A new Chromebook help desk ticket was submitted.',
    '',
    'Chromebook S/N: ' + (data.sn || ''),
    'Issue type:     ' + (data.issue || ''),
    'Urgency:        ' + (data.urgency || ''),
    '',
    'Description:',
    (data.description || ''),
    '',
    'Submitted by:   ' + (data.name || '(no name)'),
    'Teacher email:  ' + (data.email || '(none)'),
    'Room #:         ' + (data.room || ''),
    'Submitted at:   ' + now,
    '',
    HELPDESK_EMAIL,
  ];
  var body = lines.join('\n');

  // Send FROM the deploying account (you) TO the address in the "Your email" field,
  // and CC the help desk. If no valid email was entered, fall back to the help desk address.
  var valid = data.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(data.email);
  var recipient = valid ? data.email : HELPDESK_EMAIL;
  MailApp.sendEmail(recipient, subject, body,
    { name: 'CPA IT Tickets', cc: HELPDESK_EMAIL, replyTo: HELPDESK_EMAIL });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
