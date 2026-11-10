/**
 * Wedding website -> Google Sheet
 * Saves every RSVP and every "special note for the newlyweds" from the website.
 * Tabs are created automatically the first time something arrives:
 *   "RSVPs"  : Date | Name | Answer | Guests | Device ID
 *   "Notes"  : Date | Name | Note | Device ID
 * If the same guest submits the RSVP again, their row is updated (not duplicated).
 * Free: runs on your own Google account.
 */
var ANSWERS = { yes: 'Joyfully accepts', maybe: 'Maybe', no: 'Regretfully declines' };

function tab(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f4e7cd');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Text typed by guests is stored as plain text (never as a formula).
function clean(v, max) {
  var t = String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var id = String(p.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 40);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (p.action === 'rsvp_open') {
      var answer = ANSWERS[p.answer];
      var name = clean(p.name, 80);
      if (!answer || !name) return reply({ ok: false, error: 'Missing name or answer' });
      var guests = Math.max(0, Math.min(20, parseInt(p.guests, 10) || 0));
      var sh = tab('RSVPs', ['Date', 'Name', 'Answer', 'Guests', 'Device ID']);
      var row = [new Date(), name, answer, guests, id];
      var last = sh.getLastRow();
      if (id && last > 1) {
        var ids = sh.getRange(2, 5, last - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === id) { sh.getRange(i + 2, 1, 1, 5).setValues([row]); return reply({ ok: true, updated: true }); }
        }
      }
      sh.appendRow(row);
      return reply({ ok: true });
    }
    if (p.action === 'note_open') {
      var note = clean(p.text, 500);
      if (!note) return reply({ ok: false, error: 'Empty note' });
      tab('Notes', ['Date', 'Name', 'Note', 'Device ID']).appendRow([new Date(), clean(p.name, 80), note, id]);
      return reply({ ok: true });
    }
    return reply({ ok: true, message: 'Wedding sheet is connected.' });
  } finally {
    lock.releaseLock();
  }
}
