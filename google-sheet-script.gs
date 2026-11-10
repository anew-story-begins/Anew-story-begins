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
var HEAD = ['Name', 'WhatsApp', 'Ladies', 'Gents', 'Children', 'Code', 'Invite link', 'Send invite', 'Answer', 'Answered on', 'Note', 'Coming: Ladies', 'Coming: Gents', 'Coming: Children', 'Coming: Total', 'Send reminder', 'Invite sent?', 'Reminder sent?', 'Wants updates', 'Wants reminder', 'Send update'];
var C = { name: 1, phone: 2, ladies: 3, gents: 4, children: 5, code: 6, link: 7, send: 8, answer: 9, at: 10, note: 11, cl: 12, cg: 13, cc: 14, ct: 15, rem: 16, isent: 17, rsent: 18, wup: 19, wrem: 20, upd: 21 };
// Reminder wording (change when the real date / venue is known)
var WHEN = 'Tuesday, 10 November 2026 at 4:45 PM (Kenya time)';
var VENUE = '';
var ANSWERS = { yes: 'Joyfully accepts', maybe: 'Maybe', no: 'Regretfully declines' };

// Run this once from the editor (choose "setup", press Run). It creates the Guests tab.
function setup() {
  guests(); summary();
  SpreadsheetApp.getActive().toast('Guests tab is ready.', 'Wedding');
}

// Adds the "Wedding" menu every time the sheet is opened.
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('Wedding').addItem('Make invite links', 'makeLinks').addItem('Make reminder links (yes and maybe)', 'makeReminders').addItem('Send an event update…', 'makeUpdates').addToUi(); } catch (err) {}
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
    sh.getRange(row, C.isent, 1, 2).insertCheckboxes();   // tick these after you send
    sh.getRange(row, C.code).setValue(code);
    sh.getRange(row, C.link).setValue(link);
    if (phone.length >= 8) sh.getRange(row, C.send).setFormula('=HYPERLINK("https://wa.me/' + phone + '?text=' + encodeURIComponent(msg) + '","Send ➜")');
    else sh.getRange(row, C.send).setValue('Add WhatsApp number');
    made++;
  });
  summary();
  SpreadsheetApp.getActive().toast(made + ' invite link(s) ready. Click "Send ➜" for each guest.', 'Wedding');
}

// For every guest who answered "Joyfully accepts", puts a "Send reminder ➜" link in the
// Send reminder column. Click it: WhatsApp opens with the reminder text; attach the reminder picture and press send.
function makeReminders() {
  var sh = guests(), last = sh.getLastRow(), made = 0;
  if (last < 2) return;
  for (var row = 2; row <= last; row++) if (reminderCell(sh, row)) made++;
  summary();
  SpreadsheetApp.getActive().toast(made + ' reminder link(s) ready. Click "Send reminder ➜", attach the picture, send.', 'Wedding');
}

// Fills (or clears) the "Send reminder" link of one guest. Called by the menu and live after every RSVP.
function reminderCell(sh, row) {
  var r = sh.getRange(row, 1, 1, HEAD.length).getValues()[0], cell = sh.getRange(row, C.rem), name = String(r[C.name - 1]).trim();
  var ans = r[C.answer - 1];
  if (!name || (ans !== ANSWERS.yes && ans !== ANSWERS.maybe) || r[C.wrem - 1] === 'No') { cell.setValue(''); return false; }   // every yes/maybe guest, unless they switched reminders off
  var phone = String(r[C.phone - 1]).replace(/\D/g, '');
  if (phone.length < 8) { cell.setValue('Add WhatsApp number'); return false; }
  var who = counts(+r[C.cl - 1], +r[C.cg - 1], +r[C.cc - 1]);
  var msg = 'Assalam o Alaikum ' + name + ',\nA gentle reminder: our wedding is on ' + WHEN +
    (VENUE ? ' at ' + VENUE : '') + '.\n' + (ans === ANSWERS.yes ? 'We are so happy you are joining us' + (who ? ' (' + who + ')' : '') + ' and look forward to celebrating with you.'
    : 'We would love to see you there.');   // text only: attach the reminder picture yourself
  cell.setFormula('=HYPERLINK("https://wa.me/' + phone + '?text=' + encodeURIComponent(msg) + '","Send reminder ➜")');
  return true;
}

// Asks you to type an update (e.g. "The venue is now ..."), then puts a "Send update ➜" WhatsApp link
// on every guest who has not declined (unless they switched updates off).
function makeUpdates() {
  var ui = SpreadsheetApp.getUi(), res = ui.prompt('Event update', 'Type the update to send to guests:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var text = String(res.getResponseText() || '').trim(); if (!text) return;
  var sh = guests(), last = sh.getLastRow(), made = 0;
  if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  rows.forEach(function (r, i) {
    var cell = sh.getRange(i + 2, C.upd), name = String(r[C.name - 1]).trim(), phone = String(r[C.phone - 1]).replace(/\D/g, '');
    if (!name || r[C.answer - 1] === ANSWERS.no || r[C.wup - 1] === 'No') { cell.setValue(''); return; }   // everyone not declined, unless they switched updates off
    if (phone.length < 8) { cell.setValue('Add WhatsApp number'); return; }
    var msg = 'Assalam o Alaikum ' + name + ',\nWedding update: ' + text + '\nYour invitation: ' + (r[C.link - 1] || SITE);
    cell.setFormula('=HYPERLINK("https://wa.me/' + phone + '?text=' + encodeURIComponent(msg) + '","Send update ➜")');
    made++;
  });
  SpreadsheetApp.getActive().toast(made + ' update link(s) ready in the "Send update" column.', 'Wedding');
}

// "Summary" tab: every number is a live formula, so it updates by itself the moment a guest answers.
function summary() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('Summary');
  if (sh) return lists(sh);
  sh = ss.insertSheet('Summary', 1);
  var rows = [
    ['WEDDING SUMMARY (updates live)', ''],
    ['Invitations (families)', '=COUNTA(Guests!A2:A)'],
    ['Invited ladies', '=SUM(Guests!C2:C)'],
    ['Invited gents', '=SUM(Guests!D2:D)'],
    ['Invited children', '=SUM(Guests!E2:E)'],
    ['Invited people in total', '=B3+B4+B5'],
    ['', ''],
    ['Joyfully accepts', '=COUNTIF(Guests!I2:I,"Joyfully accepts")'],
    ['Maybe', '=COUNTIF(Guests!I2:I,"Maybe")'],
    ['Regretfully declines', '=COUNTIF(Guests!I2:I,"Regretfully declines")'],
    ['Not answered yet', '=COUNTA(Guests!A2:A)-COUNTA(Guests!I2:I)'],
    ['', ''],
    ['Coming: gents', '=SUMIF(Guests!I2:I,"Joyfully accepts",Guests!M2:M)'],
    ['Coming: ladies', '=SUMIF(Guests!I2:I,"Joyfully accepts",Guests!L2:L)'],
    ['Coming: children', '=SUMIF(Guests!I2:I,"Joyfully accepts",Guests!N2:N)'],
    ['COMING IN TOTAL (said yes)', '=B13+B14+B15'],
    ['Maybe coming (people)', '=SUMIF(Guests!I2:I,"Maybe",Guests!O2:O)'],
    ['', ''],
    ['Answers without a personal link: yes (people)', "=IFERROR(SUMIF('Other RSVPs'!C2:C,\"Joyfully accepts\",'Other RSVPs'!G2:G),0)"],
    ['Last updated', '=MAX(Guests!J2:J)']
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1:B1').setFontWeight('bold').setBackground('#f4e7cd');
  sh.getRange('A16:B16').setFontWeight('bold').setBackground('#f4e7cd');
  sh.getRange('B20').setNumberFormat('d mmm yyyy, h:mm am/pm');
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 160);
  return lists(sh);
}

// Live name lists on the Summary tab: who has not answered, whose invite / reminder is not ticked as sent.
function lists(sh) {
  if (sh.getRange('J1').getValue() === 'YES / MAYBE, REMINDER NOT SENT YET') return sh;   // already up to date
  var L = [
    ['D', 'NOT ANSWERED YET', '=IFERROR(FILTER(Guests!A2:B,Guests!A2:A<>"",Guests!I2:I=""),"Everyone has answered")'],
    ['G', 'INVITE NOT TICKED AS SENT', '=IFERROR(FILTER(Guests!A2:B,Guests!A2:A<>"",Guests!Q2:Q<>TRUE),"All invites sent")'],
    ['J', 'YES / MAYBE, REMINDER NOT SENT YET', '=IFERROR(FILTER(Guests!A2:B,((Guests!I2:I="Joyfully accepts")+(Guests!I2:I="Maybe"))*(Guests!T2:T<>"No"),Guests!R2:R<>TRUE),"All reminders sent")']
  ];
  L.forEach(function (l) {
    sh.getRange(l[0] + '1').setValue(l[1]);
    sh.getRange(l[0] + '1').setFontWeight('bold').setBackground('#f4e7cd');
    sh.getRange(l[0] + '2').setFormula(l[2]);
  });
  [4, 7, 10].forEach(function (c) { sh.setColumnWidth(c, 200); sh.setColumnWidth(c + 1, 140); });
  return sh;
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
    summary();
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
        reminderCell(sh, row);   // reminder link appears / disappears live
      }
      else log('Other RSVPs', ['Date', 'Name', 'Answer', 'Ladies', 'Gents', 'Children', 'Total'], [new Date(), nm, a, cl, cg, cc, cl + cg + cc]);
      return reply({ ok: true });
    }
    if (p.action === 'prefs') {           // the two switches on the website
      if (row) {
        sh.getRange(row, C.wup, 1, 2).setValues([[p.updates === '1' ? 'Yes' : 'No', p.daybefore === '1' ? 'Yes' : 'No']]);
        reminderCell(sh, row);
      }
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
