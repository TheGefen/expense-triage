// Pure logic between raw rows and the app: no network and no page access, so it can be tested.
(function (root) {
  const P = root.Parsers;
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const FIRST_DATE = '2025-01-01';
  const PARSERS = { bank: P.parseBank, cal: P.parseCal, isracard: P.parseIsracard };
  const FOLDERS = { beinleumi: 'bank', cal: 'cal', isracard: 'isracard' };

  // Finds the category rows of a year tab. Rows named SUM..., and everything from Summary on, are skipped.
  function parseYearTab(values) {
    const header = values.findIndex(r => r[2] === 'January');
    if (header < 0) throw new Error('Year tab has no month header');
    const rows = [];
    let group = '';
    let income = false;
    for (let i = header + 1; i < values.length; i++) {
      const a = String(values[i][0] || '').trim();
      const b = String(values[i][1] || '').trim();
      if (a === 'Summary') break;
      if (a === 'Income') { income = true; group = 'Income'; continue; }
      if (a) group = a;
      if (!b || /^SUM/i.test(b)) continue;
      rows.push({ row: i, group, label: b, category: income ? `Income: ${b}` : b, income });
    }
    return { header, rows };
  }

  function sumFormula(cat, month) {
    const col = cat.income ? 'I' : 'H';
    const label = cat.label.replace(/"/g, '""');
    const key = cat.income ? `Income: ${label}` : label;
    // The m prefix stops Sheets from reading "2026-01" as a date inside the condition.
    return `SUMIFS(App_Transactions!$${col}:$${col},App_Transactions!$P:$P,"${key}",App_Transactions!$F:$F,"m${month}",App_Transactions!$O:$O,"done")`;
  }

  // Keeps what the cell already had and adds the app's items on top.
  function mergedFormula(existing, cat, month) {
    const add = sumFormula(cat, month);
    if (existing === '' || existing === null || existing === undefined) return `=${add}`;
    if (typeof existing === 'string' && existing.startsWith('=')) return `=(${existing.slice(1)})+${add}`;
    const n = P.parseAmount(existing);
    if (n === null) throw new Error(`Cell for ${cat.label} ${month} has unexpected text: ${existing}`);
    return `=${n}+${add}`;
  }

  const applyAliases = (rows, aliases) => {
    if (!aliases.length) return rows;
    const map = new Map(aliases.map(a => [a.found, a.expected]));
    return rows.map(r => r.map(c => (map.has(P.clean(c)) ? map.get(P.clean(c)) : c)));
  };

  function parseFiles(files, aliases, skipped) {
    const rows = [];
    const errors = [];
    files.forEach(f => {
      if (skipped.includes(f.name)) return;
      const kind = FOLDERS[String(f.folder).toLowerCase()];
      try {
        if (!kind) throw new P.ParseError(f.name, `unknown folder "${f.folder}"`);
        if (f.error) throw new P.ParseError(f.name, f.error);
        rows.push(...PARSERS[kind](applyAliases(f.rows, aliases), f.name));
      } catch (e) {
        if (e.name !== 'ParseError') throw e;
        errors.push({ file: f.name, message: e.message, found: e.details.found || null, expected: e.details.expected || null });
      }
    });
    return { rows, errors };
  }

  const isInstallment = t => t.source === 'cal' && t.payments > 1;
  const groupOf = (t, id) => (isInstallment(t) ? `inst|${t.card}|${t.date}|${t.merchant}|${t.payments}` : id);

  // Ids must survive re-reads, so they never include the file name or a charge date.
  function withIds(kept) {
    const seen = new Map();
    return kept.map(t => {
      if (t.voucher) return { ...t, id: `v|${t.card}|${t.voucher}` };
      const base = `${t.source}|${t.card}|${t.date}|${t.merchant}|${t.amount}|${t.income || 0}`;
      const n = (seen.get(base) || 0) + 1;
      seen.set(base, n);
      return { ...t, id: `${base}#${n}` };
    });
  }

  // Rows not yet in the Sheet, with their starting status.
  function newRows(kept, existing, rules, recorded = []) {
    const ids = new Set(existing.map(t => t.id));
    const decidedGroups = new Map(existing.filter(t => t.status !== 'new').map(t => [t.group, t]));
    return withIds(kept).filter(t => t.date >= FIRST_DATE && !ids.has(t.id)).map(t => {
      const group = groupOf(t, t.id);
      const row = { ...t, group, month: `m${t.date.slice(0, 7)}`, income: t.income || 0, status: 'new', category: '', userNote: '', by: '', at: '' };
      const prior = decidedGroups.get(group);
      if (prior) return { ...row, status: prior.status, category: prior.category, by: 'installment' };
      // Already typed into the year tab by hand: counting it again would double it.
      if (recorded.some(r => r.source === t.source && t.chargeDate && t.chargeDate < r.chargedBefore)) {
        return { ...row, status: 'ignored', userNote: 'Already in the year tab', by: 'rule' };
      }
      if (t.source !== 'bank') return row;
      const r = P.classifyBank(t, rules);
      if (r.kind === 'cardCharge') return { ...row, status: 'ignored', userNote: `Card charge: ${r.issuer}`, by: 'rule' };
      if (r.kind === 'investment') return { ...row, status: 'ignored', userNote: 'Investment', by: 'rule' };
      if (r.kind === 'category') return { ...row, status: 'done', category: r.category, by: 'rule' };
      return row;
    });
  }

  // Pending rows can change merchant name when final, so exact matching misses them.
  function possibleDuplicates(rows) {
    const out = [];
    rows.filter(t => t.pending).forEach(p => {
      rows.forEach(o => {
        if (o !== p && !o.pending && o.card === p.card && o.date === p.date && o.amount === p.amount && o.merchant !== p.merchant) out.push({ pending: p, other: o });
      });
    });
    return out;
  }

  const normMerchant = m => String(m).toLowerCase().replace(/[\d*]+/g, '').replace(/\s+/g, ' ').trim();

  function queue(tx, card) {
    const groups = new Map();
    tx.filter(t => t.status === 'new' && (!card || t.card === card)).forEach(t => {
      if (!groups.has(t.group)) groups.set(t.group, []);
      groups.get(t.group).push(t);
    });
    return [...groups.values()].sort((a, b) => a[0].date.localeCompare(b[0].date));
  }

  // Own history first, then seed rules. Top three become the Suggested buttons.
  function suggest(t, tx, merchantRules) {
    const name = normMerchant(t.merchant);
    const counts = new Map();
    let total = 0;
    tx.forEach(o => {
      if (o.status === 'done' && o.category && normMerchant(o.merchant) === name) {
        counts.set(o.category, (counts.get(o.category) || 0) + 1);
        total++;
      }
    });
    const list = [...counts.entries()].sort((a, b) => b[1] - a[1]).map(e => e[0]);
    let why = total ? `${list[0]} ${counts.get(list[0])} of the last ${total} times` : '';
    const seed = merchantRules.find(r => t.merchant.toLowerCase().includes(String(r.match).toLowerCase()));
    if (seed && !list.includes(seed.category)) {
      list.push(seed.category);
      if (!why) why = `Usually ${seed.category}`;
    }
    return { list: list.slice(0, 3), why: why || 'No pattern yet' };
  }

  // Most-used categories, counted from decisions, with seed rule counts until history builds up.
  function popular(tx, merchantRules, categories, exclude, n) {
    const valid = new Set(categories);
    const counts = new Map();
    tx.forEach(t => { if (t.status === 'done' && t.by !== 'rule' && valid.has(t.category)) counts.set(t.category, (counts.get(t.category) || 0) + 1); });
    tx.forEach(t => {
      if (t.status !== 'new') return;
      const r = merchantRules.find(m => t.merchant.toLowerCase().includes(String(m.match).toLowerCase()));
      if (r && valid.has(r.category)) counts.set(r.category, (counts.get(r.category) || 0) + 0.01);
    });
    return [...counts.entries()].sort((a, b) => b[1] - a[1]).map(e => e[0]).filter(c => !exclude.includes(c)).slice(0, n);
  }

  root.Model = { MONTHS, parseYearTab, sumFormula, mergedFormula, parseFiles, withIds, newRows, possibleDuplicates, normMerchant, queue, suggest, popular };
})(typeof window !== 'undefined' ? window : globalThis);
