/**
 * Wedding guest list + website connection (free, runs on your Google account).
 *
 * Tab "Guests" (made automatically):
 *   Name | WhatsApp | Ladies | Gents | Children | Code | Invite link | Send invite |
 *   Answer | Answered on | Note
 * You fill: Name, WhatsApp (with country code, e.g. 254712345678), Ladies, Gents, Children.
 * Menu  Wedding > Make invite links  fills Code, Invite link and a "Send" link that opens
 * WhatsApp with a ready message to that guest.
 * When a guest answers on the website, Answer / Answered on (and Note) fill in by themselves.
 */
var SITE = 'https://anew-story-begins.github.io/Anew-story-begins/';
var HEAD = ['Name', 'WhatsApp', 'Ladies', 'Gents', 'Children', 'Code', 'Invite link', 'Send invite', 'Answer', 'Answered on', 'Note', 'Coming: Ladies', 'Coming: Gents', 'Coming: Children', 'Coming: Total', 'Send reminder'];
var C = { name: 1, phone: 2, ladies: 3, gents: 4, children: 5, code: 6, link: 7, send: 8, answer: 9, at: 10, note: 11, cl: 12, cg: 13, cc: 14, ct: 15, rem: 16 };
// Reminder wording (change when the real date / venue is known)
var WHEN = 'Wednesday, 11 November 2026 at 4:15 PM (Kenya time)';
var VENUE = '';
var ANSWERS = { yes: 'Joyfully accepts', maybe: 'Maybe', no: 'Regretfully declines' };

// Run this once from the editor (choose "setup", press Run). It creates the Guests tab.
function setup() {
  guests();
  SpreadsheetApp.getActive().toast('Guests tab is ready.', 'Wedding');
}

// Adds the "Wedding" menu every time the sheet is opened.
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('Wedding').addItem('Make invite links', 'makeLinks').addItem('Make reminder links (guests who said yes)', 'makeReminders').addToUi(); } catch (err) {}
}

function guests() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('Guests');
  if (!sh) {
    sh = ss.insertSheet('Guests', 0);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#f4e7cd');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.setColumnWidth(1, 200); sh.setColumnWidth(7, 260); sh.setColumnWidth(8, 120); sh.setColumnWidth(11, 260);
  }
  // add any new headings to an existing sheet
  if (sh.getLastColumn() < HEAD.length) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#f4e7cd');
  return sh;
}

function newCode(used) {
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789', c;
  do { c = ''; for (var i = 0; i < 8; i++) c += chars.charAt(Math.floor(Math.random() * chars.length)); } while (used[c]);
  used[c] = true; return c;
}

function counts(l, g, k) {
  var p = [];
  if (l > 0) p.push(l + (l == 1 ? ' Lady' : ' Ladies'));
  if (g > 0) p.push(g + (g == 1 ? ' Gent' : ' Gents'));
  if (k > 0) p.push(k + (k == 1 ? ' Child' : ' Children'));
  return p.join(', ');
}

function makeLinks() {
  var sh = guests(), last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getActive().toast('Add guests first (from row 2).', 'Wedding'); return; }
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues(), used = {}, made = 0;
  rows.forEach(function (r) { if (r[C.code - 1]) used[r[C.code - 1]] = true; });
  rows.forEach(function (r, i) {
    var name = String(r[C.name - 1]).trim(); if (!name) return;
    var code = r[C.code - 1] || newCode(used), link = SITE + '?g=' + code;
    var phone = String(r[C.phone - 1]).replace(/\D/g, '');
    var msg = 'Assalam o Alaikum ' + name + ',\nYou are warmly invited to our wedding' +
      (counts(+r[2], +r[3], +r[4]) ? ' (' + counts(+r[2], +r[3], +r[4]) + ')' : '') + '.\nPlease open your invitation: ' + link;
    var row = i + 2;
    sh.getRange(row, C.code).setValue(code);
    sh.getRange(row, C.link).setValue(link);
    if (phone.length >= 8) sh.getRange(row, C.send).setFormula('=HYPERLINK("https://wa.me/' + phone + '?text=' + encodeURIComponent(msg) + '","Send ➜")');
    else sh.getRange(row, C.send).setValue('Add WhatsApp number');
    made++;
  });
  SpreadsheetApp.getActive().toast(made + ' invite link(s) ready. Click "Send ➜" for each guest.', 'Wedding');
}

// For every guest who answered "Joyfully accepts", puts a "Send reminder ➜" link in the
// Send reminder column. Click it: WhatsApp opens with the reminder text; attach the reminder picture and press send.
function makeReminders() {
  var sh = guests(), last = sh.getLastRow(), made = 0;
  if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  rows.forEach(function (r, i) {
    var cell = sh.getRange(i + 2, C.rem), name = String(r[C.name - 1]).trim();
    if (!name || r[C.answer - 1] !== ANSWERS.yes) { cell.setValue(''); return; }
    var phone = String(r[C.phone - 1]).replace(/\D/g, '');
    if (phone.length < 8) { cell.setValue('Add WhatsApp number'); return; }
    var who = counts(+r[C.cl - 1], +r[C.cg - 1], +r[C.cc - 1]);
    var msg = 'Assalam o Alaikum ' + name + ',\nA gentle reminder: our wedding is on ' + WHEN +
      (VENUE ? ' at ' + VENUE : '') + '.\nWe are so happy you are joining us' + (who ? ' (' + who + ')' : '') +
      ' and look forward to celebrating with you.\nYour invitation: ' + (r[C.link - 1] || SITE);
    cell.setFormula('=HYPERLINK("https://wa.me/' + phone + '?text=' + encodeURIComponent(msg) + '","Send reminder ➜")');
    made++;
  });
  SpreadsheetApp.getActive().toast(made + ' reminder link(s) ready. Click "Send reminder ➜", attach the picture, send.', 'Wedding');
}

function clean(v, max) {
  var t = String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
function reply(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function findRow(sh, code) {
  code = String(code || '').trim(); if (!/^[a-z0-9]{6,12}$/.test(code)) return 0;
  var last = sh.getLastRow(); if (last < 2) return 0;
  var codes = sh.getRange(2, C.code, last - 1, 1).getValues();
  for (var i = 0; i < codes.length; i++) if (String(codes[i][0]) === code) return i + 2;
  return 0;
}
function log(name, headers, row) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.getRange(1, 1, 1, headers.length).setFontWeight('bold'); sh.setFrozenRows(1); }
  sh.appendRow(row);
}

function doGet(e) {
  var p = (e && e.parameter) || {}, sh = guests(), row = findRow(sh, p.g || p.code);
  if (p.action === 'guest') {
    if (!row) return reply({ ok: false });
    var r = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
    return reply({ ok: true, name: String(r[0]), ladies: +r[2] || 0, gents: +r[3] || 0, children: +r[4] || 0,
      answer: String(r[C.answer - 1] || '') });
  }
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (p.action === 'rsvp_open') {
      var a = ANSWERS[p.answer], nm = clean(p.name, 80);
      if (!a) return reply({ ok: false });
      var n = function (v) { return Math.max(0, Math.min(20, parseInt(v, 10) || 0)); };
      var cl = n(p.ladies), cg = n(p.gents), cc = n(p.children);
      if (row) {
        var inv = sh.getRange(row, C.ladies, 1, 3).getValues()[0];
        cl = Math.min(cl, +inv[0] || 0); cg = Math.min(cg, +inv[1] || 0); cc = Math.min(cc, +inv[2] || 0);  // never more than invited
        if (p.answer === 'no') { cl = cg = cc = 0; }
        sh.getRange(row, C.answer, 1, 2).setValues([[a, new Date()]]);
        sh.getRange(row, C.cl, 1, 4).setValues([[cl, cg, cc, cl + cg + cc]]);
      }
      else log('Other RSVPs', ['Date', 'Name', 'Answer', 'Ladies', 'Gents', 'Children', 'Total'], [new Date(), nm, a, cl, cg, cc, cl + cg + cc]);
      return reply({ ok: true });
    }
    if (p.action === 'note_open') {
      var t = clean(p.text, 500); if (!t) return reply({ ok: false });
      if (row) sh.getRange(row, C.note).setValue(t);
      else log('Other notes', ['Date', 'Name', 'Note'], [new Date(), clean(p.name, 80), t]);
      return reply({ ok: true });
    }
    return reply({ ok: true, message: 'Wedding sheet is connected.' });
  } finally { lock.releaseLock(); }
}
