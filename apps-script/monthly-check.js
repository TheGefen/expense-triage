// Monthly email: which cards need new files.
// Install: open "Expenses (App)" > Extensions > Apps Script, paste this whole file, save,
// then run setupMonthlyEmail once. It emails only the Google account that installs it.

const APP_URL = 'https://thegefen.github.io/expense-triage/';
const TIMEZONE = 'Asia/Jerusalem';
const SOURCE_NAMES = { cal: 'CAL', isracard: 'Isracard', bank: 'Bank', max: 'MAX' };
// Rule issuers like isracard1234 name a card with no export yet.
const sourceName = s => SOURCE_NAMES[s] || String(s).replace(/^([a-z]+)(\d{4})$/, (m, a, b) => `${SOURCE_NAMES[a] || a} ·${b}`);

function setupMonthlyEmail() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monthlyCheck')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyCheck').timeBased().onMonthDay(11).atHour(8).inTimezone(TIMEZONE).create();
}

function monthlyCheck() {
  const ss = SpreadsheetApp.getActive();
  const config = {};
  readTab(ss, 'App_Config').forEach(c => { config[c.key] = c.value; });
  const lastRead = config.lastRead ? new Date(JSON.parse(config.lastRead).at) : null;
  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  const report = buildReport({
    tx: readTab(ss, 'App_Transactions'),
    rules: readTab(ss, 'App_Rules'),
    cards: readTab(ss, 'App_Cards'),
    today: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'),
    unreadFiles: lastRead ? filesAddedSince(folder, lastRead) : [],
  });
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), report.subject, report.body);
}

function readTab(ss, name) {
  const values = ss.getSheetByName(name).getDataRange().getValues();
  const head = values.shift();
  return values.map(r => {
    const o = {};
    head.forEach((h, i) => { o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i], TIMEZONE, 'yyyy-MM-dd') : r[i]; });
    return o;
  });
}

// Raw files are only listed here, never opened or changed.
function filesAddedSince(folder, since) {
  const names = [];
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const files = subfolders.next().getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() > since) names.push(f.getName());
    }
  }
  return names;
}

function buildReport({ tx, rules, cards, today, unreadFiles }) {
  const f = freshness(tx, rules, today);
  const nick = c => {
    if (c.card === 'bank') return 'Bank';
    const k = cards.find(x => String(x.card) === c.card);
    return k ? `${k.nickname} ·${c.card}` : `${sourceName(c.source)} ·${c.card}`;
  };
  const day = d => `${Number(d.slice(8, 10))}/${Number(d.slice(5, 7))}`;
  const groups = {};
  tx.forEach(t => { if (t.status === 'new' && t.by !== 'rule') groups[t.group] = true; });
  const left = Object.keys(groups).length;

  const lines = [];
  if (unreadFiles.length) lines.push(`${unreadFiles.length} new file(s) in Drive not read yet. Open the app and press Re-read Drive first:`, ...unreadFiles.map(n => `  • ${n}`), '');
  if (f.stale.length) lines.push(`Needs new files (latest data before ${day(f.cutoff)}):`, ...f.stale.map(c => `  • ${nick(c)}: latest ${day(c.date)}`), '');
  if (f.fresh.length) lines.push(`Up to date: ${f.fresh.map(nick).join(', ')}`);
  if (f.missing.length) lines.push(`No files yet: ${f.missing.map(sourceName).join(', ')}`);
  lines.push('', `Still to review: ${left} purchases`, `Open: ${APP_URL}`);

  const n = f.stale.length;
  const subject = unreadFiles.length ? 'Expenses: new files waiting to be read'
    : n ? `Expenses: ${n} ${n === 1 ? 'card needs' : 'cards need'} new files` : 'Expenses: all cards up to date';
  return { subject, body: lines.join('\n'), stale: f.stale };
}

// Same as Model.freshness in model.js; keep the two the same.
function freshness(tx, rules, today) {
  const first = new Date(Date.UTC(Number(today.slice(0, 4)), Number(today.slice(5, 7)) - 1, 1));
  const cutoff = new Date(first.getTime() - 7 * 864e5).toISOString().slice(0, 10);
  const latest = new Map();
  tx.forEach(t => {
    const d = [String(t.date || ''), String(t.chargeDate || '')].sort().pop();
    const k = String(t.card);
    if (!latest.has(k) || d > latest.get(k).date) latest.set(k, { card: k, source: String(t.source), date: d });
  });
  const all = [...latest.values()];
  const sources = new Set(all.map(x => x.source));
  const missing = [...new Set(rules.filter(r => r.kind === 'cardCharge' && !sources.has(r.issuer)).map(r => r.issuer))];
  return { cutoff, stale: all.filter(x => x.date < cutoff), fresh: all.filter(x => x.date >= cutoff), missing };
}
