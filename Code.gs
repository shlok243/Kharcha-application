// ====== Kharcha: expense tracker backend (Google Apps Script) ======
// Writes into a Google Sheet that keeps one block per month (title row, header row, entries, totals row).
//
// SETUP: replace the placeholder below with your own spreadsheet ID.
// It is the long text between /d/ and /edit in your sheet's URL:
//   https://docs.google.com/spreadsheets/d/<THIS_PART>/edit
// Never commit your real ID to a public repository.
const SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const TZ = 'Asia/Kolkata'; // change to your own timezone, e.g. 'America/New_York'
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
                'August', 'September', 'October', 'November', 'December'];
const YELLOW = '#ffff00', GREEN = '#00ff00', RED = '#ff0000';

// Columns: A Date | B Received | C Spent | D Specification | E Person/mode | F Time (added by the app)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kharcha')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- helpers ----------
function sheet_() {
  if (SHEET_ID.indexOf('YOUR_') === 0) {
    throw new Error('Open Code.gs and set SHEET_ID to your spreadsheet ID first.');
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('Sheet1') || ss.getSheets()[0];
}
function isNum_(v) { return typeof v === 'number' && !isNaN(v); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v); }
function pad_(n) { return ('0' + n).slice(-2); }

// A subtotal / total row: numbers in B-D, nothing in E (e.g. ",415,7217,6820")
function isTotal_(r) {
  return isNum_(r[3]) && String(r[4]).trim() === '' && (isNum_(r[1]) || isNum_(r[2]));
}
// A real entry: has an amount in Received or Spent and is not a total row
function isEntry_(r) {
  return !isTotal_(r) && (isNum_(r[1]) || isNum_(r[2]));
}
function dateKey_(v, tz) {
  if (isDate_(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + '-' + pad_(m[1]) + '-' + pad_(m[2]) : '';
}
function modeOf_(s) {
  s = String(s).toLowerCase().replace(/\s/g, '');
  if (!s) return 'Other';
  if (s.indexOf('fam') >= 0) return 'FamPay';
  if (s.indexOf('cash') >= 0) return 'Cash';
  if (/^gp|^goog/.test(s)) return 'GPay';
  return 'Other';
}
function nowParts_() {
  const p = Utilities.formatDate(new Date(), TZ, 'yyyy,M,d,H,m').split(',').map(Number);
  const o = { y: p[0], mo: p[1], d: p[2], h: p[3], mi: p[4] };
  o.monthKey = o.y + '-' + pad_(o.mo);
  o.dayKey = o.monthKey + '-' + pad_(o.d);
  return o;
}

// Reads the whole sheet once and finds the last month block (the one after the last "Date" header row)
function read_() {
  const sh = sheet_();
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const lastRow = sh.getLastRow();
  let vals = [], disp = [], hdr = 0;
  if (lastRow > 0) {
    const rng = sh.getRange(1, 1, lastRow, 6);
    vals = rng.getValues();
    disp = rng.getDisplayValues();
    for (let i = lastRow - 1; i >= 0; i--) {
      if (String(vals[i][0]).trim().toLowerCase() === 'date') { hdr = i + 1; break; }
    }
  }
  const b = { sh: sh, tz: tz, lastRow: lastRow, vals: vals, disp: disp, hdr: hdr, lastEntryKey: '' };
  b.lastIsTotal = lastRow > 0 && isTotal_(vals[lastRow - 1]);
  if (hdr > 0) {
    for (let i = lastRow - 1; i >= hdr; i--) {
      if (isEntry_(vals[i])) {
        const k = dateKey_(vals[i][0], tz);
        if (k) { b.lastEntryKey = k; break; }
      }
    }
  }
  return b;
}

// Totals for the last month block
function sums_(b) {
  const s = { spent: 0, received: 0, n: 0, modes: { Cash: 0, GPay: 0, FamPay: 0, Other: 0 } };
  if (b.hdr === 0) return s;
  for (let i = b.hdr; i < b.lastRow; i++) {
    const r = b.vals[i];
    if (!isEntry_(r)) continue;
    s.n++;
    if (isNum_(r[1])) s.received += r[1];
    if (isNum_(r[2])) { s.spent += r[2]; s.modes[modeOf_(r[4])] += r[2]; }
  }
  return s;
}

function ensureRows_(sh, n) {
  if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
}

// ---------- drawing the sheet ----------
function writeTitle_(sh, row, text) {
  const r = sh.getRange(row, 1, 1, 5);
  r.merge();
  sh.getRange(row, 1).setValue(text);
  r.setBackground(YELLOW).setFontWeight('bold').setHorizontalAlignment('center');
}
function writeHeader_(sh, row) {
  sh.getRange(row, 1, 1, 6)
    .setValues([['Date', 'Received', 'Spent', 'Specification', 'Person/mode', 'Time']])
    .setBackground(YELLOW).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 2).setBackground(GREEN);
  sh.getRange(row, 3).setBackground(RED);
}
function writeTotals_(sh, row, s) {
  sh.getRange(row, 2, 1, 3)
    .setValues([[s.received, s.spent, s.spent - s.received]])
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(row, 2).setBackground(GREEN);
  sh.getRange(row, 3).setBackground(RED);
}
function writeEntry_(sh, row, e, p) {
  const rng = sh.getRange(row, 1, 1, 6);
  const amtFmt = e.amount % 1 === 0 ? '0' : '0.00';
  const dateSerial = Date.UTC(p.y, p.mo - 1, p.d) / 86400000 + 25569; // timezone-proof date
  const timeSerial = (p.h * 60 + p.mi) / 1440;
  rng.setNumberFormats([['m/d/yyyy', amtFmt, amtFmt, '@', '@', 'h:mm AM/PM']]);
  rng.setValues([[dateSerial,
                  e.type === 'Received' ? e.amount : '',
                  e.type === 'Spent' ? e.amount : '',
                  e.desc, e.mode, timeSerial]]);
  rng.setHorizontalAlignment('center').setFontWeight('normal');
}

// ---------- called from the app ----------
function addEntry(entry) {
  const amount = Number(entry.amount);
  if (!amount || amount <= 0 || amount > 10000000) throw new Error('Enter a valid amount');
  const e = {
    amount: amount,
    type: entry.type === 'Received' ? 'Received' : 'Spent',
    desc: String(entry.desc || '').trim().slice(0, 100),
    mode: ['Cash', 'GPay', 'FamPay'].indexOf(entry.mode) >= 0 ? entry.mode : 'Cash'
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const b = read_();
    const sh = b.sh;
    const p = nowParts_();
    const needNew = b.hdr === 0 || (b.lastEntryKey !== '' && b.lastEntryKey.slice(0, 7) !== p.monthKey);
    let start, entryRow;

    if (!needNew) {
      // same month: write on the next empty line of the current block
      entryRow = b.lastRow + 1;
      start = entryRow;
      ensureRows_(sh, entryRow);
      if (String(b.vals[b.hdr - 1][5]).trim() === '') {
        sh.getRange(b.hdr, 6).setValue('Time').setBackground(YELLOW).setFontWeight('bold').setHorizontalAlignment('center');
      }
    } else {
      // new month: close the old block with totals, then start a new block
      let r;
      if (b.lastRow > 0) {
        start = b.lastRow + 1;
        if (b.hdr > 0 && !b.lastIsTotal) {
          const s = sums_(b);
          const totalsRow = b.lastRow + 2;
          ensureRows_(sh, totalsRow);
          writeTotals_(sh, totalsRow, s);
          PropertiesService.getScriptProperties().setProperty(
            'closed', JSON.stringify({ key: b.lastEntryKey.slice(0, 7), s: s }));
          r = totalsRow + 2;
        } else {
          r = b.lastRow + 2;
        }
      } else {
        start = 1;
        r = 1;
      }
      ensureRows_(sh, r + 2);
      writeTitle_(sh, r, MONTHS[p.mo - 1] + ' ' + p.y);
      writeHeader_(sh, r + 1);
      entryRow = r + 2;
    }

    writeEntry_(sh, entryRow, e, p);
    PropertiesService.getScriptProperties().setProperty(
      'lastOp', JSON.stringify({ start: start, count: entryRow - start + 1 }));
  } finally {
    lock.releaseLock();
  }
  return getStats();
}

// Undo only the last thing the app wrote (the entry, plus the month heading/totals if that entry created them)
function deleteLast() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const op = JSON.parse(props.getProperty('lastOp') || 'null');
    const sh = sheet_();
    if (!op || sh.getLastRow() !== op.start + op.count - 1) throw new Error('Nothing to undo');
    sh.deleteRows(op.start, op.count);
    props.deleteProperty('lastOp');
  } finally {
    lock.releaseLock();
  }
  return getStats();
}

// Numbers shown inside the app
function getStats() {
  const b = read_();
  const p = nowParts_();
  let monthSpent = 0, monthReceived = 0, todaySpent = 0;
  const current = b.hdr > 0 && (b.lastEntryKey === '' || b.lastEntryKey.slice(0, 7) === p.monthKey);
  if (current) {
    for (let i = b.hdr; i < b.lastRow; i++) {
      const r = b.vals[i];
      if (!isEntry_(r)) continue;
      if (isNum_(r[1])) monthReceived += r[1];
      if (isNum_(r[2])) {
        monthSpent += r[2];
        if (dateKey_(r[0], b.tz) === p.dayKey) todaySpent += r[2];
      }
    }
  }
  const recent = [];
  for (let i = b.lastRow - 1; i >= 0 && recent.length < 6; i--) {
    const r = b.vals[i];
    if (!isEntry_(r)) continue;
    const k = dateKey_(r[0], b.tz);
    let when = k ? k.slice(8) + ' ' + MONTHS[Number(k.slice(5, 7)) - 1].slice(0, 3) : '';
    const t = String(b.disp[i][5]).trim();
    if (/^\d{1,2}:\d{2}\s?[AP]M$/i.test(t)) when += (when ? ', ' : '') + t;
    const spent = isNum_(r[2]);
    recent.push({
      amount: spent ? r[2] : r[1], type: spent ? 'Spent' : 'Received',
      desc: String(r[3]), mode: String(r[4]), when: when
    });
  }
  const op = JSON.parse(PropertiesService.getScriptProperties().getProperty('lastOp') || 'null');
  return {
    monthLabel: MONTHS[p.mo - 1] + ' ' + p.y,
    monthSpent: monthSpent, monthReceived: monthReceived, todaySpent: todaySpent,
    recent: recent,
    canUndo: !!op && b.lastRow === op.start + op.count - 1
  };
}

// ---------- monthly auto-calculation ----------
// Run setupMonthlyReport() ONCE from the editor.
function setupMonthlyReport() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'monthlyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monthlyReport').timeBased().onMonthDay(1).atHour(8).inTimezone(TZ).create();
}

// Runs by itself on the 1st of every month: writes the totals row under last month and emails you the summary
function monthlyReport() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const b = read_();
    const p = nowParts_();

    if (b.hdr > 0 && b.lastEntryKey !== '' && b.lastEntryKey.slice(0, 7) !== p.monthKey) {
      const s = sums_(b);
      if (!b.lastIsTotal) {
        ensureRows_(b.sh, b.lastRow + 2);
        writeTotals_(b.sh, b.lastRow + 2, s);
      }
      props.setProperty('closed', JSON.stringify({ key: b.lastEntryKey.slice(0, 7), s: s }));
    }

    let py = p.y, pm = p.mo - 1;
    if (pm === 0) { pm = 12; py--; }
    const prevKey = py + '-' + pad_(pm);
    if (props.getProperty('emailed') === prevKey) return;

    const label = MONTHS[pm - 1] + ' ' + py;
    const closed = JSON.parse(props.getProperty('closed') || 'null');
    let body;
    if (closed && closed.key === prevKey) {
      const s = closed.s;
      body = label + '\n\nSpent: ₹' + s.spent + '\nReceived: ₹' + s.received +
        '\nNet spent: ₹' + (s.spent - s.received) +
        '\n\nCash: ₹' + s.modes.Cash + '\nGPay: ₹' + s.modes.GPay + '\nFamPay: ₹' + s.modes.FamPay +
        '\nOther / not noted: ₹' + s.modes.Other + '\nEntries: ' + s.n;
    } else {
      body = 'No entries were logged in ' + label + '.';
    }
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Expense report - ' + label, body);
    props.setProperty('emailed', prevKey);
  } finally {
    lock.releaseLock();
  }
}
