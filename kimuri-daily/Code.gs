/***********************************************************************
 * KIMURI DAILY – backend (Google Apps Script)
 * Sheet: "KIMURI OCT 26" (one tab per date, named d-m-yyyy)
 *
 *  - Serves a small JSON API for the Vercel app (doPost)
 *  - Creates a new tab per day; yesterday's CLOSING stock becomes today's BALANCE
 *  - Builds the summaries and sends them to Telegram
 *  - Daily auto-run + "Kimuri Bot" menu inside the sheet
 *
 * SECRETS ARE NOT IN THIS FILE. Add them in:
 *   Project Settings (gear icon) > Script properties
 *     TG_TOKEN    = 8851012923:AAGf_IoKMPWkETwwskZ1qWoTYOrJmzSgcXo
 *     TG_CHAT_ID  = 474423713
 *     API_KEY     = https://script.google.com/macros/s/AKfycbzyG2B6OOWvSnrWZJHBHcvvATAPZNSmrj2XlLEbfwYL7K267qeGe0mpWghoe4Gnqusprg/exec (same value goes in Vercel as GAS_KEY)
 ***********************************************************************/

/* ============================== CONFIG ============================== */
const CONFIG = {
  BUSINESS: 'KIMURI',
  SEND_HOUR: 21,        // daily auto-run time (24h, script timezone)
  PERIOD_DAYS: 15,      // length of each periodic report
  EMAIL_COPY: false     // true = also email each summary to yourself
};

const SHEET_ID = '1PFHMaJqEm95_p09p4_Lwsc1aFqzl8lQRCcb1zpu9yMo';
const MAX_ROWS = 80;
const META = ['TOTAL SALES', 'EXPENSES', 'EXPECTED CASH', 'DEBTS', 'MPESA', 'CASH', 'CASH AT HAND', 'DEFICIT'];
const KG_ITEMS = ['BEEF', 'PORK', 'BONES'];

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function prop_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Missing script property: ' + key);
  return v;
}

/* ============================ WEB API =============================== */
// The Vercel app calls this (through its own /api/gas proxy).
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.key !== prop_('API_KEY')) throw new Error('Unauthorized');
    out = { ok: true, data: route_(req.action, req.data || {}) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('Kimuri Daily API is running.');
}

function route_(action, d) {
  switch (action) {
    case 'getDay':      return getDay(d.iso);
    case 'saveDay':     return saveDay(d);
    case 'getSummary':  return getSummary(d.iso);
    case 'sendSummary': return sendSummary(d.iso);
    case 'getReport':   return getReport(d);
    case 'listDays':    return listDays();
    case 'sendUnsent':  return { sent: runDailyJob() };
    default: throw new Error('Unknown action: ' + action);
  }
}

/* ========================= SCHEDULED JOB ============================ */
// Sends a summary for every completed day not yet sent, then any due periodic report.
// Returns how many daily summaries were sent.
function runDailyJob() {
  const props = PropertiesService.getScriptProperties();
  const sent = JSON.parse(props.getProperty('SENT') || '[]');
  const days = loadDays_();
  let count = 0;

  days.forEach(d => {
    if (sent.indexOf(d.name) === -1) {
      sendMessage_(dailyMessage_(d));
      sent.push(d.name);
      props.setProperty('SENT', JSON.stringify(sent));
      count++;
      Utilities.sleep(1500);
    }
  });

  periodReports_(days, props);
  return count;
}

/* =========================== PARSING ================================ */
function loadDays_() {
  const days = [];
  ss_().getSheets().forEach(sh => {
    const dn = dayNumFromName_(sh.getName());
    if (dn === null) return;                 // skips tabs like "Sheet11"
    const d = parseSheet_(sh, dn);
    if (d) days.push(d);
  });
  return days.sort((a, b) => a.dayNum - b.dayNum);
}

function parseSheet_(sh, dayNum) {
  const rows = Math.min(sh.getLastRow(), MAX_ROWS);
  if (rows < 2) return null;
  const data = sh.getRange(1, 1, rows, 10).getValues();
  const d = { name: sh.getName(), dayNum: dayNum, items: [], meta: {}, beefVar: 0, kitchenVar: 0 };

  data.forEach(r => {
    const label = String(r[1]).trim().toUpperCase();
    const hLabel = String(r[7]).trim().toUpperCase();
    if (label && META.indexOf(label) !== -1) {
      d.meta[label] = num_(r[9]);
    } else if (label && typeof r[5] === 'number') {          // item row (has a price)
      const qty = num_(r[8]), amt = num_(r[9]);
      if (amt !== 0) d.items.push({ item: label, qty: qty, amt: amt });
    }
    if (hLabel === 'BEEF') d.beefVar = num_(r[9]);
    if (hLabel.indexOf('KITCHEN') === 0) d.kitchenVar = num_(r[9]);
  });

  const m = d.meta;
  if (!(m['TOTAL SALES'] > 0) || !(m['CASH AT HAND'] > 0)) return null;  // day not complete yet
  d.sales = m['TOTAL SALES'];
  d.expenses = m['EXPENSES'] || 0;
  d.expected = m['EXPECTED CASH'] || (d.sales - d.expenses);
  d.cashAtHand = m['CASH AT HAND'];
  d.variance = d.cashAtHand - d.expected;    // negative = shortage
  return d;
}

/* ============================ MESSAGES ============================== */
function dailyMessage_(d) {
  const m = d.meta;
  const status = d.variance < 0 ? 'SHORT ' + fmt_(-d.variance)
               : d.variance > 0 ? 'OVER ' + fmt_(d.variance) : 'BALANCED';
  let t = '*' + CONFIG.BUSINESS + ' – ' + dateLabel_(d.dayNum) + '*\n';
  t += 'Total sales: ' + fmt_(d.sales) + '\n';
  t += 'Expenses: ' + fmt_(d.expenses) + '\n';
  t += 'Expected cash: ' + fmt_(d.expected) + '\n';
  t += 'Cash at hand: ' + fmt_(d.cashAtHand) +
       ' (M-Pesa ' + fmt_(m['MPESA'] || 0) + ', Cash ' + fmt_(m['CASH'] || 0) +
       ', Debts ' + fmt_(m['DEBTS'] || 0) + ')\n';
  t += '*Result: ' + status + '*\n';
  t += 'Beef ' + signed_(d.beefVar) + ' | Kitchen & Pork ' + signed_(d.kitchenVar) + '\n\n';
  t += '*Sales per item*\n';
  d.items.slice().sort((a, b) => b.amt - a.amt).forEach(i => {
    t += i.item + ' (' + qtyStr_(i) + '): ' + fmt_(i.amt) + '\n';
  });
  return t;
}

function periodMessage_(days, start, end) {
  let sales = 0, exp = 0, beef = 0, kit = 0, short = 0, over = 0, daysShort = 0, daysOver = 0;
  const per = {};
  let worst = days[0];
  days.forEach(d => {
    sales += d.sales; exp += d.expenses; beef += d.beefVar; kit += d.kitchenVar;
    if (d.variance < 0) { short += -d.variance; daysShort++; }
    if (d.variance > 0) { over += d.variance; daysOver++; }
    if (d.variance < worst.variance) worst = d;
    d.items.forEach(i => {
      per[i.item] = per[i.item] || { qty: 0, amt: 0 };
      per[i.item].qty += i.qty; per[i.item].amt += i.amt;
    });
  });
  const netVar = over - short;

  let t = '*' + CONFIG.BUSINESS + ' – ' + (end - start + 1) + '-DAY REPORT*\n';
  t += dateLabel_(start) + ' to ' + dateLabel_(end) + ' (' + days.length + ' days)\n\n';
  t += 'Total sales: ' + fmt_(sales) + '\n';
  t += 'Total expenses: ' + fmt_(exp) + '\n';
  t += 'Sales less expenses: ' + fmt_(sales - exp) + '\n';
  t += 'Average daily sales: ' + fmt_(sales / days.length) + '\n\n';
  t += '*Losses & gains*\n';
  t += 'Total shortages: -' + fmt_(short) + ' (' + daysShort + ' days)\n';
  t += 'Total overages: +' + fmt_(over) + ' (' + daysOver + ' days)\n';
  t += 'Net cash variance: ' + signed_(netVar) + '\n';
  t += 'Beef variance: ' + signed_(beef) + '\n';
  t += 'Kitchen & Pork variance: ' + signed_(kit) + '\n';
  t += 'Worst day: ' + dateLabel_(worst.dayNum) + ' (' + signed_(worst.variance) + ')\n\n';
  t += '*Sales per item*\n';
  Object.keys(per).sort((a, b) => per[b].amt - per[a].amt).forEach(k => {
    const u = KG_ITEMS.indexOf(k) !== -1 ? ' kg' : '';
    t += k + ' (' + Math.round(per[k].qty * 10) / 10 + u + '): ' + fmt_(per[k].amt) + '\n';
  });
  return t;
}

// Sends a report each time a full block of PERIOD_DAYS days has been completed.
function periodReports_(days, props) {
  if (!days.length) return;
  const latest = days[days.length - 1].dayNum;
  const last = props.getProperty('LAST_REPORT_END');
  let start = last ? Number(last) + 1 : days[0].dayNum;

  while (latest >= start + CONFIG.PERIOD_DAYS - 1) {
    const end = start + CONFIG.PERIOD_DAYS - 1;
    const inRange = days.filter(d => d.dayNum >= start && d.dayNum <= end);
    if (inRange.length) {
      sendMessage_(periodMessage_(inRange, start, end));
      Utilities.sleep(1500);
    }
    props.setProperty('LAST_REPORT_END', String(end));
    start = end + 1;
  }
}

/* ============================ TELEGRAM ============================== */
function sendMessage_(text) {
  const token = prop_('TG_TOKEN'), chatId = prop_('TG_CHAT_ID');
  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  for (let i = 0; i < text.length; i += 3500) {
    const chunk = text.substring(i, i + 3500);
    let res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chatId, text: chunk, parse_mode: 'Markdown' })
    });
    if (res.getResponseCode() !== 200) {     // retry as plain text if formatting is rejected
      res = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ chat_id: chatId, text: chunk })
      });
    }
    Logger.log(res.getResponseCode() + ' ' + res.getContentText().substring(0, 120));
    if (res.getResponseCode() !== 200) throw new Error('Telegram error: ' + res.getContentText().substring(0, 150));
  }
  if (CONFIG.EMAIL_COPY) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), CONFIG.BUSINESS + ' summary', text);
  }
}

function testTelegram() { sendMessage_('✅ ' + CONFIG.BUSINESS + ' bot is connected.'); }

/* ============================ HELPERS =============================== */
function dayNumFromName_(name) {
  const m = String(name).trim().match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  return m ? Math.round(Date.UTC(+m[3], +m[2] - 1, +m[1]) / 86400000) : null;
}
function dateLabel_(dayNum) {
  const d = new Date(dayNum * 86400000);
  return d.getUTCDate() + '/' + (d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
}
function isoFromDayNum_(dayNum) { return new Date(dayNum * 86400000).toISOString().slice(0, 10); }
function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
function fmt_(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function signed_(n) { return (n < 0 ? '-' : '+') + fmt_(Math.abs(n)); }
function qtyStr_(i) {
  return Math.round(i.qty * 10) / 10 + (KG_ITEMS.indexOf(i.item) !== -1 ? 'kg' : '');
}
function nameFromIso_(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso))) throw new Error('Invalid date: ' + iso);
  const p = iso.split('-');
  return (+p[2]) + '-' + (+p[1]) + '-' + p[0];
}
function markSent_(name) {
  const props = PropertiesService.getScriptProperties();
  const sent = JSON.parse(props.getProperty('SENT') || '[]');
  if (sent.indexOf(name) === -1) { sent.push(name); props.setProperty('SENT', JSON.stringify(sent)); }
}

function latestDateSheet_(ss) {
  let best = null, bd = -1;
  ss.getSheets().forEach(sh => {
    const dn = dayNumFromName_(sh.getName());
    if (dn !== null && dn > bd) { bd = dn; best = sh; }
  });
  return best;
}

// Latest date tab dated BEFORE dayNum (null if none).
function prevDateSheet_(ss, dayNum) {
  let best = null, bd = -1;
  ss.getSheets().forEach(sh => {
    const dn = dayNumFromName_(sh.getName());
    if (dn !== null && dn < dayNum && dn > bd) { bd = dn; best = sh; }
  });
  return best;
}

// Finds item rows and label rows by content, so layout changes don't break anything.
function scan_(sh) {
  const rows = Math.min(sh.getLastRow(), MAX_ROWS);
  const data = sh.getRange(1, 1, rows, 10).getValues();
  const out = { items: [], rows: {}, beefRow: 0, kitchenRow: 0, beefVar: '', kitchenVar: '' };
  data.forEach((r, i) => {
    const label = String(r[1]).trim().toUpperCase();
    const h = String(r[7]).trim().toUpperCase();
    const row = i + 1;
    if (label && META.indexOf(label) !== -1) out.rows[label] = row;
    else if (label && typeof r[5] === 'number')
      out.items.push({ item: label, row: row, price: r[5], open: r[2], bal: r[3], util: r[6], close: r[7] });
    if (h === 'BEEF') { out.beefRow = row; out.beefVar = r[9]; }
    if (h.indexOf('KITCHEN') === 0) { out.kitchenRow = row; out.kitchenVar = r[9]; }
  });
  return out;
}

// Map of ITEM -> closing stock on a given tab.
function closingMap_(sh) {
  const map = {};
  scan_(sh).items.forEach(i => { map[i.item] = i.close; });
  return map;
}

function summaryText_(sh) {
  const d = parseSheet_(sh, dayNumFromName_(sh.getName()));
  return d ? dailyMessage_(d)
           : 'Saved. The summary appears once Total Sales and Cash at Hand are filled in.';
}

/* ====================== API: DAY ENTRY (the form) =================== */
function getDay(iso) {
  const ss = ss_(), name = nameFromIso_(iso);
  let sh = ss.getSheetByName(name);
  const exists = !!sh;
  const prev = prevDateSheet_(ss, dayNumFromName_(name));
  if (!exists) sh = prev || latestDateSheet_(ss);   // template for the item list
  if (!sh) throw new Error('No date tab found to use as a template.');

  const s = scan_(sh), r = s.rows;
  const clean = v => (v === '' || v == null) ? '' : v;
  const val = row => (exists && row) ? clean(sh.getRange(row, 10).getValue()) : '';
  const carry = (!exists && prev) ? closingMap_(prev) : {};

  return {
    exists: exists,
    prevName: (!exists && prev) ? prev.getName() : '',
    items: s.items.map(i => ({
      item: i.item, price: i.price,
      open:  exists ? clean(i.open)  : '',
      bal:   exists ? clean(i.bal)   : clean(carry[i.item]),   // yesterday's closing
      util:  exists ? clean(i.util)  : '',
      close: exists ? clean(i.close) : ''
    })),
    expenses: val(r['EXPENSES']), debts: val(r['DEBTS']),
    mpesa: val(r['MPESA']), cash: val(r['CASH']),
    beefVar: exists ? clean(s.beefVar) : '', kitchenVar: exists ? clean(s.kitchenVar) : ''
  };
}

function saveDay(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_(), name = nameFromIso_(p.iso);
    let sh = ss.getSheetByName(name);
    const isNew = !sh;
    let carry = {};

    if (isNew) {                               // new day: copy the previous tab as a template
      const prev = prevDateSheet_(ss, dayNumFromName_(name));
      const tpl = prev || latestDateSheet_(ss);
      if (!tpl) throw new Error('No date tab found to use as a template.');
      if (prev) carry = closingMap_(prev);
      sh = tpl.copyTo(ss).setName(name);

      // wipe the copied entries (formulas are left alone)
      const t = scan_(sh), tr = t.rows;
      t.items.forEach(it => {
        sh.getRange(it.row, 3, 1, 2).clearContent();   // opening, balance
        sh.getRange(it.row, 7, 1, 2).clearContent();   // utilities, closing
      });
      ['EXPENSES', 'DEBTS', 'MPESA', 'CASH'].forEach(k => {
        if (tr[k]) sh.getRange(tr[k], 10).clearContent();
      });
      if (t.beefRow) sh.getRange(t.beefRow, 10).clearContent();
      if (t.kitchenRow) sh.getRange(t.kitchenRow, 10).clearContent();
    }

    const s = scan_(sh), r = s.rows;
    const cell = v => (v === '' || v == null || isNaN(parseFloat(v))) ? '' : parseFloat(v);

    const byName = {};
    p.items.forEach(i => byName[i.item] = i);
    s.items.forEach(it => {
      const v = byName[it.item];
      if (!v) return;
      // brand-new day with an empty balance: fall back to yesterday's closing
      const bal = (isNew && cell(v.bal) === '') ? cell(carry[it.item]) : cell(v.bal);
      sh.getRange(it.row, 3, 1, 2).setValues([[cell(v.open), bal]]);            // opening, balance
      sh.getRange(it.row, 7, 1, 2).setValues([[cell(v.util), cell(v.close)]]);  // utilities, closing
    });

    const put = (row, v) => { if (row) sh.getRange(row, 10).setValue(cell(v)); };
    put(r['EXPENSES'], p.expenses); put(r['DEBTS'], p.debts);
    put(r['MPESA'], p.mpesa);       put(r['CASH'], p.cash);
    put(s.beefRow, p.beefVar);      put(s.kitchenRow, p.kitchenVar);

    // keep the derived cells as live formulas
    if (r['EXPECTED CASH'] && r['TOTAL SALES'] && r['EXPENSES'])
      sh.getRange(r['EXPECTED CASH'], 10).setFormula('=J' + r['TOTAL SALES'] + '-J' + r['EXPENSES']);
    if (r['CASH AT HAND'] && r['DEBTS'] && r['CASH'])
      sh.getRange(r['CASH AT HAND'], 10).setFormula('=SUM(J' + r['DEBTS'] + ':J' + r['CASH'] + ')');
    if (r['DEFICIT'] && r['CASH AT HAND'] && r['EXPECTED CASH'])
      sh.getRange(r['DEFICIT'], 10).setFormula('=J' + r['CASH AT HAND'] + '-J' + r['EXPECTED CASH']);

    SpreadsheetApp.flush();

    const d = parseSheet_(sh, dayNumFromName_(name));
    return {
      ok: true,
      created: isNew,
      summary: summaryText_(sh),
      result: d ? { complete: true, sales: d.sales, expenses: d.expenses, expected: d.expected,
                    cashAtHand: d.cashAtHand, variance: d.variance }
                : { complete: false }
    };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== API: SUMMARIES & REPORTS ===================== */
function getSummary(iso) {
  const sh = ss_().getSheetByName(nameFromIso_(iso));
  return sh ? summaryText_(sh) : 'No data saved for this date yet.';
}

function sendSummary(iso) {
  const sh = ss_().getSheetByName(nameFromIso_(iso));
  if (!sh) throw new Error('No data saved for this date yet.');
  const d = parseSheet_(sh, dayNumFromName_(sh.getName()));
  if (!d) throw new Error('Total Sales and Cash at Hand must be filled in first.');
  const text = dailyMessage_(d);
  sendMessage_(text);
  markSent_(d.name);
  periodReports_(loadDays_(), PropertiesService.getScriptProperties());   // sends the periodic report when due
  return text;
}

// o = { start:'yyyy-mm-dd', end:'yyyy-mm-dd', send:true|false }; no dates = last PERIOD_DAYS days.
function getReport(o) {
  o = o || {};
  const days = loadDays_();
  if (!days.length) return 'No completed days yet.';
  let start, end;
  if (o.start && o.end) {
    start = dayNumFromName_(nameFromIso_(o.start));
    end = dayNumFromName_(nameFromIso_(o.end));
    if (start > end) { const t = start; start = end; end = t; }
  } else {
    end = days[days.length - 1].dayNum;
    start = end - CONFIG.PERIOD_DAYS + 1;
  }
  const inRange = days.filter(d => d.dayNum >= start && d.dayNum <= end);
  if (!inRange.length) return 'No completed days in that range.';
  const text = periodMessage_(inRange, start, end);
  if (o.send) sendMessage_(text);
  return text;
}

function getPeriod(send) { return getReport({ send: !!send }); }

// Recent completed days for the summary page (newest first).
function listDays() {
  const sent = JSON.parse(PropertiesService.getScriptProperties().getProperty('SENT') || '[]');
  return loadDays_().slice().reverse().slice(0, 60).map(d => ({
    iso: isoFromDayNum_(d.dayNum), name: d.name, sales: d.sales,
    variance: d.variance, sent: sent.indexOf(d.name) !== -1
  }));
}

/* ============================== SETUP =============================== */
// Run once: schedules runDailyJob every day at SEND_HOUR.
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runDailyJob').timeBased().everyDays(1).atHour(CONFIG.SEND_HOUR).create();
}

// Optional, run once BEFORE runDailyJob if you don't want existing days re-sent.
function markExistingAsSent() {
  const names = ss_().getSheets().map(s => s.getName()).filter(n => dayNumFromName_(n) !== null);
  PropertiesService.getScriptProperties().setProperty('SENT', JSON.stringify(names));
}

/* ======================= SHEET MENU (date triggers) ======================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Kimuri Bot')
    .addItem('Send summary: this tab', 'menuSendThisTab')
    .addItem('Send summary: pick a date…', 'menuSendDate')
    .addItem('Send daily summaries: date range…', 'menuSendRange')
    .addSeparator()
    .addItem('Send report: date range…', 'menuReportRange')
    .addItem('Send report: last 15 days', 'menuReportLast')
    .addSeparator()
    .addItem('Send all unsent days', 'runDailyJob')
    .addToUi();
}

function toast_(msg) { SpreadsheetApp.getActive().toast(msg, 'Kimuri Bot', 5); }

function sendDay_(d) { sendMessage_(dailyMessage_(d)); markSent_(d.name); }

function parseDateInput_(s) {
  const m = String(s).trim().match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  return m ? Math.round(Date.UTC(+m[3], +m[2] - 1, +m[1]) / 86400000) : null;
}

function askRange_(title) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(title, 'Enter two dates, e.g.  3-10-2026 to 17-10-2026', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const parts = r.getResponseText().trim().split(/\s+to\s+|\s*,\s*|\s*;\s*/i);
  if (parts.length !== 2) { ui.alert('Please enter two dates separated by "to".'); return null; }
  const a = parseDateInput_(parts[0]), b = parseDateInput_(parts[1]);
  if (a === null || b === null) { ui.alert('Could not read those dates. Use d-m-yyyy.'); return null; }
  return { start: Math.min(a, b), end: Math.max(a, b) };
}

function menuSendThisTab() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const dn = dayNumFromName_(sh.getName());
  if (dn === null) return ui.alert('This tab is not a date tab (expected d-m-yyyy, e.g. 3-10-2026).');
  const d = parseSheet_(sh, dn);
  if (!d) return ui.alert('Total Sales and Cash at Hand must be filled in first.');
  sendDay_(d);
  toast_('Sent summary for ' + sh.getName());
}

function menuSendDate() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Send summary for a date', 'Date (d-m-yyyy), e.g. 5-10-2026', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const dn = parseDateInput_(r.getResponseText());
  if (dn === null) return ui.alert('Could not read that date. Use d-m-yyyy.');
  const d = loadDays_().filter(x => x.dayNum === dn)[0];
  if (!d) return ui.alert('No completed tab found for ' + dateLabel_(dn) + '.');
  sendDay_(d);
  toast_('Sent summary for ' + dateLabel_(dn));
}

function menuSendRange() {
  const ui = SpreadsheetApp.getUi();
  const rg = askRange_('Send daily summaries');
  if (!rg) return;
  const days = loadDays_().filter(d => d.dayNum >= rg.start && d.dayNum <= rg.end);
  if (!days.length) return ui.alert('No completed days in that range.');
  const ok = ui.alert('Send ' + days.length + ' daily summaries (' +
                      dateLabel_(rg.start) + ' to ' + dateLabel_(rg.end) + ')?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  days.forEach(d => { sendDay_(d); Utilities.sleep(1500); });
  toast_('Sent ' + days.length + ' summaries');
}

function menuReportRange() {
  const ui = SpreadsheetApp.getUi();
  const rg = askRange_('Send report for a date range');
  if (!rg) return;
  const days = loadDays_().filter(d => d.dayNum >= rg.start && d.dayNum <= rg.end);
  if (!days.length) return ui.alert('No completed days in that range.');
  sendMessage_(periodMessage_(days, rg.start, rg.end));
  toast_('Report sent for ' + days.length + ' days');
}

function menuReportLast() {
  const text = getPeriod(true);
  toast_(text.indexOf('No completed') === 0 ? text : 'Report sent');
}
