// Turns raw export rows (arrays of cells, as read from Excel) into transactions.
// Columns are found by header name, so a moved column still parses and a renamed
// one raises a ParseError instead of importing wrong numbers.
(function (root) {
  class ParseError extends Error {
    constructor(file, message, details) {
      super(`${file}: ${message}`);
      this.name = 'ParseError';
      this.file = file;
      this.details = details || {};
    }
  }

  // Real exports have line breaks inside header cells and invisible direction marks.
  const clean = v => (v === null || v === undefined ? '' : String(v).replace(/[‎‏‪-‮]/g, '').replace(/\s+/g, ' ').trim());
  const round2 = n => Math.round(n * 100) / 100;

  function parseAmount(v) {
    if (typeof v === 'number') return v;
    const s = clean(v);
    if (!s) return null;
    const m = s.match(/\d[\d,]*(\.\d+)?/);
    if (!m) return null;
    const n = parseFloat(m[0].replace(/,/g, ''));
    return s.indexOf('-') !== -1 && s.indexOf('-') < s.indexOf(m[0]) ? -n : n;
  }

  // Returns yyyy-mm-dd. Accepts dd/mm/yy, dd/mm/yyyy, dd.mm.yy and Date objects.
  function parseDate(v) {
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    const m = clean(v).match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{2}|\d{4})$/);
    if (!m) return null;
    const y = m[3].length === 2 ? `20${m[3]}` : m[3];
    return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }

  function findHeader(rows, required, file) {
    for (let i = 0; i < rows.length; i++) {
      const cells = rows[i].map(clean);
      if (required.every(name => cells.includes(name))) {
        const col = {};
        cells.forEach((c, j) => { if (c) col[c] = j; });
        return { index: i, col };
      }
    }
    const best = rows.map(r => r.map(clean).filter(Boolean)).sort(
      (a, b) => required.filter(n => b.includes(n)).length - required.filter(n => a.includes(n)).length
    )[0] || [];
    throw new ParseError(file, `missing columns: ${required.filter(n => !best.includes(n)).join(', ')}`, { found: best, expected: required });
  }

  function parseBank(rows, file) {
    const h = findHeader(rows, ['תאריך', 'תיאור', 'חובה', 'זכות', 'אסמכתא'], file);
    const out = [];
    for (let i = h.index + 1; i < rows.length; i++) {
      const r = rows[i];
      const date = parseDate(r[h.col['תאריך']]);
      if (!date) {
        if (r.some(c => clean(c))) throw new ParseError(file, `row ${i + 1}: unreadable date "${clean(r[h.col['תאריך']])}"`);
        continue;
      }
      const debit = parseAmount(r[h.col['חובה']]);
      const credit = parseAmount(r[h.col['זכות']]);
      if (debit === null && credit === null) throw new ParseError(file, `row ${i + 1}: no amount`);
      out.push({
        source: 'bank', card: 'bank', date, merchant: clean(r[h.col['תיאור']]),
        amount: round2(debit || 0), income: round2(credit || 0), ref: clean(r[h.col['אסמכתא']]), file,
      });
    }
    return out;
  }

  function parseCal(rows, file) {
    const amountCol = 'סכום בש"ח';
    const h = findHeader(rows, ['תאריך עסקה', 'שם בית עסק', amountCol, 'כרטיס', 'מועד חיוב', 'סוג עסקה'], file);
    const out = [];
    for (let i = h.index + 1; i < rows.length; i++) {
      const r = rows[i];
      const date = parseDate(r[h.col['תאריך עסקה']]);
      if (!date) continue; // totals and footer text have no date
      const amount = parseAmount(r[h.col[amountCol]]);
      if (amount === null) throw new ParseError(file, `row ${i + 1}: no amount`);
      const card = (clean(r[h.col['כרטיס']]).match(/\d{4}/) || [])[0];
      if (!card) throw new ParseError(file, `row ${i + 1}: no card number`);
      const type = clean(r[h.col['סוג עסקה']]);
      const note = h.col['הערות'] === undefined ? '' : clean(r[h.col['הערות']]);
      const payments = (note.match(/ב-(\d+) תשלומים/) || [])[1];
      out.push({
        source: 'cal', card, date, merchant: clean(r[h.col['שם בית עסק']]), amount: round2(amount),
        chargeDate: parseDate(r[h.col['מועד חיוב']]), type, note,
        payments: payments ? Number(payments) : null, pending: note.includes('בקליטה'), file,
      });
    }
    return out;
  }

  function parseIsracard(rows, file) {
    const h = findHeader(rows, ['תאריך רכישה', 'שם בית עסק', 'סכום עסקה', 'סכום חיוב', "מס' שובר"], file);
    const before = rows.slice(0, h.index).map(r => r.map(clean).join(' ')).join(' ');
    const card = (before.match(/-\s*(\d{4})/) || [])[1];
    if (!card) throw new ParseError(file, 'card number not found above the table');
    const charge = before.match(/לחיוב ב-(\d{1,2})\.(\d{1,2})/);
    const year = (before.match(/(20\d{2})/) || [])[1];
    if (!charge || !year) throw new ParseError(file, 'statement charge date not found above the table');
    const chargeDate = `${year}-${charge[2].padStart(2, '0')}-${charge[1].padStart(2, '0')}`;
    const out = [];
    let statedTotal = null;
    for (let i = h.index + 1; i < rows.length; i++) {
      const r = rows[i];
      const line = r.map(clean).join(' ');
      if (line.includes('סה"כ לחיוב')) { statedTotal = parseAmount(r[h.col['סכום חיוב']]); continue; }
      const date = parseDate(r[h.col['תאריך רכישה']]);
      if (!date) continue;
      const amount = parseAmount(r[h.col['סכום חיוב']]);
      if (amount === null) throw new ParseError(file, `row ${i + 1}: no charged amount`);
      out.push({
        source: 'isracard', card, date, merchant: clean(r[h.col['שם בית עסק']]), amount: round2(amount),
        chargeDate, voucher: clean(r[h.col["מס' שובר"]]), note: h.col['פירוט נוסף'] === undefined ? '' : clean(r[h.col['פירוט נוסף']]), file,
      });
    }
    if (statedTotal === null) throw new ParseError(file, 'statement total row not found');
    const sum = round2(out.reduce((s, t) => s + t.amount, 0));
    if (Math.abs(sum - statedTotal) > 0.01) throw new ParseError(file, `rows add up to ${sum} but the statement says ${statedTotal}`);
    // A waived fee is charged 0, so it is not an expense.
    return out.filter(t => t.amount !== 0);
  }

  // Merges rows from every file. Rows that look identical inside one file are real
  // (installments, two coffees), so the count kept is the highest count in any one file.
  function dedupe(all) {
    const key = t => t.voucher ? `v|${t.card}|${t.voucher}` : `${t.source}|${t.card}|${t.date}|${t.merchant}|${t.amount}|${t.income || 0}`;
    const groups = new Map();
    all.forEach(t => {
      const k = key(t);
      if (!groups.has(k)) groups.set(k, new Map());
      const byFile = groups.get(k);
      if (!byFile.has(t.file)) byFile.set(t.file, []);
      byFile.get(t.file).push(t);
    });
    const kept = [];
    const duplicates = [];
    groups.forEach(byFile => {
      const lists = [...byFile.values()].sort((a, b) => b.length - a.length);
      // Prefer the copy with a known charge date: pending rows get one later.
      const pool = lists.flat().sort((a, b) => (b.chargeDate ? 1 : 0) - (a.chargeDate ? 1 : 0));
      const n = lists[0].length;
      kept.push(...pool.slice(0, n));
      duplicates.push(...pool.slice(n));
    });
    return { kept, duplicates };
  }

  // Rules come from the private Sheet, never from this public code. First match wins.
  const classifyBank = (t, rules) => rules.find(r => t.merchant.includes(r.match)) || { kind: 'triage' };

  // Compares bank card charges with the card rows, month by month and per issuer.
  // Rows lacking a charge date (CAL installments, pending rows) can't be placed.
  function reconcile(bankRows, cardRows, rules) {
    const month = d => d.slice(0, 7);
    const lines = new Map();
    const get = k => { if (!lines.has(k)) lines.set(k, { bank: 0, items: 0 }); return lines.get(k); };
    bankRows.forEach(t => {
      const r = classifyBank(t, rules);
      if (r.kind === 'cardCharge') get(`${r.issuer}|${month(t.date)}`).bank += t.amount - t.income;
    });
    cardRows.forEach(t => { if (t.chargeDate) get(`${t.source}|${month(t.chargeDate)}`).items += t.amount; });
    return [...lines.entries()].map(([k, v]) => {
      const [issuer, m] = k.split('|');
      return { issuer, month: m, bank: round2(v.bank), items: round2(v.items), diff: round2(v.bank - v.items) };
    }).sort((a, b) => (a.month + a.issuer).localeCompare(b.month + b.issuer));
  }

  root.Parsers = { ParseError, clean, parseAmount, parseDate, parseBank, parseCal, parseIsracard, dedupe, classifyBank, reconcile };
})(typeof window !== 'undefined' ? window : globalThis);
