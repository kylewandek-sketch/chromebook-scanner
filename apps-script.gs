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

// Column order. Column A is the Chromebook S/N, as required.
var HEADERS = [
  'Chromebook S/N', 'Timestamp', 'Teacher Email', 'Teacher Name',
  'Room #', 'Issue Type', 'Urgency', 'Description', 'Status'
];

function doGet() {
  return ContentService
    .createTextOutput('Chromebook Help Desk endpoint is live. Submit tickets via POST.')
    .setMimeType(ContentService.MimeType.TEXT);
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
  ];
  var body = lines.join('\n');

  // Send FROM the deploying account (you) TO the address in the "Your email" field.
  // If no valid email was entered, fall back to the help desk address.
  var valid = data.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(data.email);
  var recipient = valid ? data.email : HELPDESK_EMAIL;
  MailApp.sendEmail(recipient, subject, body, { name: 'Chromebook Help Desk' });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
