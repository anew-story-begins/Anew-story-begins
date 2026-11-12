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
// Wedding WhatsApp group (guests who turn on Updates & Reminders are invited to join it)
var GROUP = 'https://chat.whatsapp.com/Bxaf5dKSHY02HZL1HQibKn';
var HEAD = ['Name', 'WhatsApp', 'Ladies', 'Gents', 'Children', 'Code', 'Invite link', 'Send invite', 'Answer', 'Answered on', 'Note', 'Coming: Ladies', 'Coming: Gents', 'Coming: Children', 'Coming: Total', 'Send reminder', 'Invite sent?', 'Reminder sent?', 'Wants updates', 'Wants reminder', 'Send update'];
var C = { name: 1, phone: 2, ladies: 3, gents: 4, children: 5, code: 6, link: 7, send: 8, answer: 9, at: 10, note: 11, cl: 12, cg: 13, cc: 14, ct: 15, rem: 16, isent: 17, rsent: 18, wup: 19, wrem: 20, upd: 21 };
// Reminder wording (change when the real date / venue is known)
var WHEN = 'Tuesday, 10 November 2026 at 4:45 PM (Kenya time)';
var VENUE = '';
var ANSWERS = { yes: 'Joyfully accepts', maybe: 'Maybe', no: 'Regretfully declines' };

// Run this once from the editor (choose "setup", press Run). It creates the Guests tab.
function setup() {
  guests(); summary(); try { showSendLink(); } catch (e) {}
  SpreadsheetApp.getActive().toast('Guests tab is ready.', 'Wedding');
}

// Adds the "Wedding" menu every time the sheet is opened.
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('Wedding').addItem('Make invite links', 'makeLinks').addItem('Send invites one by one ➜', 'sendSidebar').addItem('Get my phone sending page link', 'showSendLink').addItem('Make reminder links (yes and maybe)', 'makeReminders').addItem('Send an event update…', 'makeUpdates').addToUi(); } catch (err) {}
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
    var msg = inviteMsg(r, link);
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
// Send reminder column. Click it: the guest's WhatsApp chat opens; attach the reminder picture and press send.
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
  // picture only: the link just opens this guest's WhatsApp chat (no text). Attach the reminder picture and send.
  cell.setFormula('=HYPERLINK("https://wa.me/' + phone + '","Send reminder ➜")');
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
    var msg = 'Hi\nWedding update: ' + text + '\nYour invitation: ' + (r[C.link - 1] || SITE);
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
  var T3 = 'REMINDER NOT SENT YET (said yes / maybe)';
  if (sh.getRange('L1').getValue() === T3) return sh;   // already up to date
  sh.getRange('D1:Z200').clear();
  var ok = '((Guests!I2:I="Joyfully accepts")+(Guests!I2:I="Maybe"))*(Guests!T2:T<>"No")*(Guests!R2:R<>TRUE)';
  var L = [
    ['D', 'NOT ANSWERED YET', ['Name', 'WhatsApp'], '=IFERROR(FILTER(Guests!A2:B,Guests!A2:A<>"",Guests!I2:I=""),"Everyone has answered")'],
    ['H', 'INVITE NOT SENT YET', ['Name', 'WhatsApp'], '=IFERROR(FILTER(Guests!A2:B,Guests!A2:A<>"",Guests!Q2:Q<>TRUE),"All invites sent")'],
    ['L', T3, ['Name', 'WhatsApp', 'Answer'], '=IFERROR(FILTER({Guests!A2:B,Guests!I2:I},' + ok + '),"All reminders sent")']
  ];
  L.forEach(function (l) {
    var col = l[0].charCodeAt(0) - 64;
    sh.getRange(1, col).setValue(l[1]);
    sh.getRange(1, col, 1, l[2].length).setFontWeight('bold').setBackground('#f4e7cd');
    sh.getRange(2, col, 1, l[2].length).setValues([l[2]]).setFontWeight('bold');
    sh.getRange(3, col).setFormula(l[3]);
    sh.setColumnWidth(col, 200); sh.setColumnWidth(col + 1, 130); if (l[2].length > 2) sh.setColumnWidth(col + 2, 150);
  });
  sh.getRange('E:E').setNumberFormat('0'); sh.getRange('I:I').setNumberFormat('0'); sh.getRange('M:M').setNumberFormat('0');
  return sh;
}

// ---------- "Send invites one by one" panel ----------
// Opens a panel on the right. Each click on "Send to next guest" opens WhatsApp for the next guest
// whose "Invite sent?" box is not ticked (message and personal link already typed) and ticks the box.
function inviteMsg(r, link) {
  return 'A new story begins\n' + link;
}
function sendSidebar() {
  var html = '<div style="font:14px Arial;padding:6px">' +
    '<div id="info" style="margin-bottom:12px">Loading…</div>' +
    '<button id="go" style="width:100%;padding:12px;font-size:15px;background:#25D366;color:#fff;border:0;border-radius:6px;cursor:pointer" disabled>Send to next guest ➜</button>' +
    '<button id="skip" style="width:100%;margin-top:8px;padding:8px;border:1px solid #ccc;background:#fff;border-radius:6px;cursor:pointer" disabled>Skip this guest</button>' +
    '<p style="color:#666;font-size:12px;margin-top:14px">WhatsApp opens with the message ready. Press send there, then come back and click again.</p></div>' +
    '<script>var cur=null,after=1;' +
    'function show(g){cur=g;var go=document.getElementById("go"),sk=document.getElementById("skip");' +
    'if(!g){document.getElementById("info").innerHTML="<b>All invites are ticked as sent.</b>";go.disabled=sk.disabled=true;return}' +
    'document.getElementById("info").innerHTML="Next: <b>"+g.name.replace(/</g,"&lt;")+"</b><br>"+g.phone+"<br><span style=color:#666>"+g.left+" still to send</span>";go.disabled=sk.disabled=false}' +
    'function load(){google.script.run.withSuccessHandler(show).nextGuest(after)}' +
    'document.getElementById("go").onclick=function(){if(!cur)return;window.open(cur.url,"_blank");after=1;google.script.run.withSuccessHandler(load).markInviteSent(cur.row)};' +
    'document.getElementById("skip").onclick=function(){if(!cur)return;after=cur.row+1;load()};load();</script>';
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('Send invites'));
}
function nextGuest(fromRow) {
  var sh = guests(), last = sh.getLastRow(); if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues(), left = 0, found = null;
  if (rows.some(function (r) { return String(r[C.name - 1]).trim() && !r[C.code - 1]; })) { makeLinks(); rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues(); }
  rows.forEach(function (r, i) {
    var row = i + 2, name = String(r[C.name - 1]).trim(), phone = String(r[C.phone - 1]).replace(/\D/g, '');
    if (!name || phone.length < 8 || r[C.isent - 1] === true) return;
    left++;
    if (!found && row >= (fromRow || 1)) found = { row: row, name: name, phone: '+' + phone,
      url: 'https://wa.me/' + phone + '?text=' + encodeURIComponent(inviteMsg(r, r[C.link - 1] || SITE + '?g=' + r[C.code - 1])) };
  });
  if (found) found.left = left;
  return found;
}
function markInviteSent(row) { guests().getRange(row, C.isent).insertCheckboxes().setValue(true); }

// ---------- Private phone sending page ----------
// A page only you can open (secret key). It lists guests with big buttons that open the WhatsApp APP directly
// (not the Sheets in-app browser), and ticks "Invite sent?" / "Reminder sent?" for you.
function sendKey() {
  var pr = PropertiesService.getScriptProperties(), k = pr.getProperty('SEND_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); pr.setProperty('SEND_KEY', k); }
  return k;
}
function sendPageUrl() { var u = ScriptApp.getService().getUrl(); return u ? u + '?action=send&k=' + sendKey() : ''; }
function showSendLink() {
  var u = sendPageUrl(), sh = summary();
  sh.getRange('A23').setValue('Phone sending page (private, do not share):').setFontWeight('bold');
  sh.getRange('A24').setValue(u || 'Deploy the web app first (Deploy > Manage deployments).');
  sh.getRange('A26').setValue('Wedding WhatsApp group (send reminders and updates here):').setFontWeight('bold');
  sh.getRange('A27').setFormula('=HYPERLINK("' + GROUP + '","Open the wedding group ➜")');
  SpreadsheetApp.getActive().toast('Your private sending page link is on the Summary tab (cell A24).', 'Wedding', 8);
}
function esc(t) { return String(t).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function sendPage(k) {
  var sh = guests(), last = sh.getLastRow(), inv = [], rem = [];
  if (last >= 2) sh.getRange(2, 1, last - 1, HEAD.length).getValues().forEach(function (r, i) {
    var row = i + 2, name = String(r[C.name - 1]).trim(), phone = String(r[C.phone - 1]).replace(/\D/g, '');
    if (!name || phone.length < 8) return;
    var link = r[C.link - 1] || (r[C.code - 1] ? SITE + '?g=' + r[C.code - 1] : '');
    if (link) inv.push({ row: row, name: name, phone: phone, done: r[C.isent - 1] === true, text: inviteMsg(r, link) });
    var a = r[C.answer - 1];
    if ((a === ANSWERS.yes || a === ANSWERS.maybe) && r[C.wrem - 1] !== 'No') rem.push({ row: row, name: name, phone: phone, done: r[C.rsent - 1] === true, text: '' });
  });
  function list(arr, kind) {
    arr.sort(function (x, y) { return x.done - y.done; });
    if (!arr.length) return '<p class="m">Nobody here yet.</p>';
    return arr.map(function (g) {
      var q = 'phone=' + g.phone + (g.text ? '&text=' + encodeURIComponent(g.text) : '');
      return '<div class="g' + (g.done ? ' d' : '') + '" id="' + kind + g.row + '"><div><b>' + esc(g.name) + '</b><br><small>+' + g.phone + '</small></div>' +
        (kind === 'r'
          ? '<span class="bs"><a class="b2" target="_blank" rel="noopener" href="https://api.whatsapp.com/send?' + q + '">Chat</a>' +
            '<a class="b" target="_blank" rel="noopener" href="' + SITE + 'share.html?n=' + encodeURIComponent(g.name) + '" onclick="mk(\'r\',' + g.row + ')">' + (g.done ? 'Sent ✓ again' : 'Share picture ➜') + '</a></span>'
          : '<a class="b" target="_blank" rel="noopener" href="https://api.whatsapp.com/send?' + q + '" onclick="mk(\'' + kind + '\',' + g.row + ')">' + (g.done ? 'Sent ✓ again' : 'Send ➜') + '</a>') + '</div>';
    }).join('');
  }
  var html = '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>' +
    'body{font:16px Arial;margin:0;background:#f7f1e6;color:#3a160f}h2{margin:0;padding:14px 16px;background:#5a0f18;color:#f3e6cf;font-size:18px}' +
    '.t{display:flex}.t button{flex:1;padding:12px;border:0;background:#eadfcb;font-size:15px}.t button.on{background:#fff;font-weight:bold}' +
    '.g{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid #e3d6bf;background:#fff}.g.d{opacity:.5}' +
    '.b{background:#25D366;color:#fff;text-decoration:none;padding:12px 16px;border-radius:8px;font-weight:bold;white-space:nowrap}.bs{display:flex;gap:8px;align-items:center}.b2{color:#5a0f18;text-decoration:none;padding:12px 10px;border:1px solid #c9b48f;border-radius:8px;white-space:nowrap}.m{padding:16px}.grp{margin:12px 16px;padding:14px;border:1px solid #c9b48f;border-radius:10px;background:#fff;display:flex;flex-direction:column;gap:8px}.grp span{font-size:13px;color:#6b4a2a}' +
    '.n{font-size:13px;padding:10px 16px;color:#6b4a2a}</style></head><body><h2>Send from your phone</h2>' +
    '<div class="t"><button id="ti" class="on" onclick="tab(1)">Invites</button><button id="tr" onclick="tab(0)">Reminders (picture)</button></div>' +
    '<div id="pi"><p class="n">Tap Send ➜, WhatsApp opens with the message ready, press send, come back for the next.</p>' + list(inv, 'i') + '</div>' +
    '<div id="pr" style="display:none">' +
    '<div class="grp"><b>Send to everyone in the wedding group</b><span>Share the picture, choose WhatsApp, then pick the group.</span>' +
    '<div class="bs"><a class="b2" target="_blank" rel="noopener" href="' + GROUP + '">Open group</a><a class="b" target="_blank" rel="noopener" href="' + SITE + 'share.html?n=' + encodeURIComponent('the wedding group') + '">Share picture ➜</a></div></div>' +
    '<p class="n">Or send to one guest at a time:</p>' + list(rem, 'r') + '</div>' +
    '<script>var K=' + JSON.stringify(k) + ';function tab(i){document.getElementById("pi").style.display=i?"":"none";document.getElementById("pr").style.display=i?"none":"";' +
    'document.getElementById("ti").className=i?"on":"";document.getElementById("tr").className=i?"":"on"}' +
    'function mk(kind,row){var e=document.getElementById(kind+row);e.className="g d";var bb=e.querySelector("a.b");if(bb)bb.textContent="Sent ✓ again";google.script.run.webMark(K,kind,row)}</script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Send invites').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function webMark(k, kind, row) {
  if (k !== sendKey()) return;
  row = parseInt(row, 10); if (!(row >= 2)) return;
  guests().getRange(row, kind === 'r' ? C.rsent : C.isent).insertCheckboxes().setValue(true);
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
  if (p.action === 'send') return p.k && p.k === sendKey() ? sendPage(p.k) : ContentService.createTextOutput('Not allowed');
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
