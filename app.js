// Pages and interaction. Data comes from a backend: Google, or made-up data in demo mode.
(function () {
  const DEMO = new URLSearchParams(location.search).has('demo');
  const api = DEMO ? window.DemoBackend : window.GoogleBackend;
  const { Model } = window;
  const app = document.getElementById('app');
  const nav = document.getElementById('nav');
  const PAGES = [['triage', 'Triage'], ['all', 'All'], ['summary', 'Summary'], ['data', 'Data'], ['settings', 'Settings']];

  const ui = { page: 'triage', card: '', error: '', busy: '', history: [], skipped: new Set(), allFilter: 'all', search: '', editing: null, year: '', quarter: 0, showAll: false, note: '' };
  let db = null;

  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = a => `${a < 0 ? '−' : ''}₪${Math.abs(a).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const whole = a => `${a < 0 ? '−' : ''}₪${Math.round(Math.abs(a)).toLocaleString('en-US')}`;
  const shortDate = d => { const [y, m, day] = String(d).split('-'); return `${day}/${m}/${y}`; };
  const nickname = t => {
    if (t.card === 'bank') return 'Bank';
    const c = db.cards.find(x => String(x.card) === String(t.card));
    return c ? `${c.nickname} ·${t.card}` : `${t.source} ·${t.card}`;
  };
  // Sends only the merchant name to Google, never the amount or card.
  const googleLink = (m, label) => `<a class="btn link" href="https://www.google.com/search?q=${encodeURIComponent(m)}" target="_blank" rel="noopener noreferrer">${label}</a>`;
  const lastRead = () => (db.config.lastRead ? JSON.parse(db.config.lastRead.value) : null);

  function categories() {
    const seen = new Map();
    Object.keys(db.years).sort().reverse().forEach(y => {
      Model.parseYearTab(db.years[y]).rows.forEach(c => { if (!seen.has(c.category)) seen.set(c.category, c); });
    });
    return [...seen.values()];
  }

  async function guarded(label, fn) {
    ui.busy = label;
    ui.error = '';
    render();
    try {
      await fn();
    } catch (e) {
      ui.error = e.message;
      if (e.signIn) db = null;
      api.log('error', '', e.message);
      console.error(e);
    }
    ui.busy = '';
    render();
  }

  // ---------- Sign-in and first-time setup ----------
  function pageSignIn() {
    return `<div class="card done"><h1>Family Expenses</h1><p class="sub">Sign in with the Google account that has access to the expenses folder.</p>
      <button class="btn primary" data-act="signin">Sign in with Google</button></div>`;
  }
  function pageSetup() {
    return `<div class="card"><h1>First-time setup</h1>
      <p>This creates a new Google Sheet called "Expenses (App)" in your folder, copied from Expenses.xlsx. The original and the bank and card folders are not changed.</p>
      <label class="sub" for="seed">Choose the starting rules file from your PC (seed.local.json)</label>
      <input type="file" id="seed" accept=".json">
      <div class="actions"><button class="btn primary" data-act="create">Create my copy</button></div>
      <p class="sub">Only needs doing once, by one person. If you're not the owner, ask them to share the folder with you.</p></div>`;
  }

  async function signIn() {
    await guarded('Signing in…', async () => {
      await api.signIn();
      const found = await api.findCopy();
      if (!found.exists) { db = { needsSetup: true }; return; }
      db = await api.load();
    });
  }
  async function createCopy() {
    const file = document.getElementById('seed').files[0];
    if (!file) { ui.error = 'Choose the rules file first.'; render(); return; }
    await guarded('Creating your copy. This can take a minute…', async () => {
      const seed = JSON.parse(await file.text());
      await api.createCopy(seed);
      db = await api.load();
    });
    if (db && !db.needsSetup) rereadDrive();
  }

  // ---------- Re-read ----------
  function rereadDrive() {
    return guarded('Reading your files…', async () => {
      db = await api.load();
      const skipped = db.config.skippedFiles ? JSON.parse(db.config.skippedFiles.value) : [];
      const files = await api.readRawFiles();
      const parsed = Model.parseFiles(files, db.aliases, skipped);
      const { kept, duplicates } = window.Parsers.dedupe(parsed.rows);
      const recorded = db.config.alreadyRecorded ? JSON.parse(db.config.alreadyRecorded.value) : [];
      const fresh = Model.newRows(kept, db.tx, db.rules, recorded);
      if (fresh.length) await api.appendTx(fresh);
      const bank = kept.filter(t => t.source === 'bank');
      const cards = kept.filter(t => t.source !== 'bank');
      const coverage = {};
      kept.forEach(t => {
        const k = t.source === 'bank' ? 'Bank' : `${t.source} ·${t.card}`;
        const c = coverage[k] = coverage[k] || { from: t.date, to: t.date };
        if (t.date < c.from) c.from = t.date;
        if (t.date > c.to) c.to = t.date;
      });
      const summary = {
        at: new Date().toISOString(),
        files: files.length,
        total: kept.length,
        added: fresh.length,
        errors: parsed.errors,
        duplicates: duplicates.slice(0, 50).map(t => ({ merchant: t.merchant, amount: t.amount, date: t.date, file: t.file })),
        duplicateCount: duplicates.length,
        possible: Model.possibleDuplicates(cards).slice(0, 20).map(p => ({ a: p.pending.merchant, b: p.other.merchant, amount: p.other.amount, date: p.other.date })),
        reconcile: window.Parsers.reconcile(bank, cards, db.rules).filter(r => r.bank !== 0),
        coverage,
      };
      await api.setConfig('lastRead', JSON.stringify(summary), db.config.lastRead);
      parsed.errors.forEach(e => api.log('error', e.file, e.message));
      db = await api.load();
      ui.history = []; // the loaded rows replace the ones Back was holding
    });
  }

  // ---------- Decisions ----------
  async function decide(rows, status, category) {
    const at = new Date().toISOString();
    const by = DEMO ? 'demo' : '';
    const prev = rows.map(t => ({ t, status: t.status, category: t.category, userNote: t.userNote, by: t.by, at: t.at }));
    rows.forEach(t => { Object.assign(t, { status, category: category || '', userNote: ui.note || t.userNote, by: by || 'app', at }); });
    ui.history.push(prev);
    ui.note = '';
    ui.showAll = false;
    render();
    try {
      await api.updateTx(rows);
    } catch (e) {
      prev.forEach(p => Object.assign(p.t, { status: p.status, category: p.category, userNote: p.userNote, by: p.by, at: p.at }));
      ui.history.pop();
      ui.error = `Not saved: ${e.message}`;
      render();
    }
  }
  async function undo() {
    const prev = ui.history.pop();
    if (!prev) return;
    prev.forEach(p => Object.assign(p.t, { status: p.status, category: p.category, userNote: p.userNote, by: p.by, at: p.at }));
    ui.skipped.delete(prev[0].t.group);
    render();
    try { await api.updateTx(prev.map(p => p.t)); } catch (e) { ui.error = `Not saved: ${e.message}`; render(); }
  }

  // ---------- Triage ----------
  function pageTriage() {
    const cats = categories();
    const names = cats.map(c => c.category);
    const all = Model.queue(db.tx, ui.card);
    const q = all.filter(g => !ui.skipped.has(g[0].group)).concat(all.filter(g => ui.skipped.has(g[0].group)));
    const cardsSeen = [...new Set(db.tx.map(t => String(t.card)))];
    const chips = `<div class="chips"><button class="chip ${ui.card ? '' : 'on'}" data-act="card" data-v="">All</button>${cardsSeen.map(c => {
      const t = db.tx.find(x => String(x.card) === c);
      return `<button class="chip ${ui.card === c ? 'on' : ''}" data-act="card" data-v="${esc(c)}">${esc(nickname(t))}</button>`;
    }).join('')}</div>`;
    const totalDecidable = db.tx.filter(t => t.by !== 'rule').length || 1;
    const pct = Math.round(100 * (1 - db.tx.filter(t => t.status === 'new').length / totalDecidable));
    const head = `<div class="row"><h1>Triage</h1><span class="sub">${all.length} left to review</span></div>${chips}<div class="progress"><i style="width:${pct}%"></i></div>`;
    if (!db.tx.length) return `${head}<div class="card done"><h2>No transactions yet</h2><button class="btn primary" data-act="reread">Read my files</button></div>`;
    if (!q.length) return `${head}<div class="card done"><h2>All done</h2><p class="sub">Nothing left to review here.</p><button class="btn" data-act="undo" ${ui.history.length ? '' : 'disabled'}>← Back</button></div>`;

    const g = q[0];
    const t = g[0];
    const total = g.reduce((s, x) => s + (x.amount || 0) - (x.income || 0), 0);
    const isIncome = t.income > 0;
    const pool = names.filter(n => n.startsWith('Income: ') === isIncome);
    const sug = Model.suggest(t, db.tx, db.merchantRules);
    const suggested = sug.list.filter(c => pool.includes(c));
    const pop = Model.popular(db.tx, db.merchantRules, pool, suggested, 6);
    const btn = (c, cls = '') => `<button class="cat ${cls}" data-act="pick" data-v="${esc(c)}">${esc(c.replace('Income: ', ''))}</button>`;
    const groups = [...new Set(cats.map(c => c.group))].map(gr => {
      const inGroup = cats.filter(c => c.group === gr && pool.includes(c.category));
      return inGroup.length ? `<div class="tagline">${esc(gr)}</div><div class="grid">${inGroup.map(c => btn(c.category)).join('')}</div>` : '';
    }).join('');
    const tags = [];
    if (g.length > 1 || t.payments > 1) tags.push(`Installments: ${g.length} payment${g.length > 1 ? 's' : ''} seen of ${t.payments}. All count in the month bought.`);
    if (total < 0 && !isIncome) tags.push('Refund: reduces the category total');
    if (/\$/.test(t.note)) tags.push(t.note);
    if (t.pending) tags.push('Still processing at the card company');
    const needsNote = /^(BIT|PAYBOX|העברה)/i.test(t.merchant) || t.source === 'bank';

    return `${head}<div class="card"><div class="sub">${esc(nickname(t))}</div><div class="merchant" dir="auto">${esc(t.merchant)}</div>
      <div class="amount">${money(isIncome ? -total : total)}${isIncome ? ' in' : ''}</div>
      <div class="meta"><span>${shortDate(t.date)}</span><span dir="auto">${esc(t.type)}</span></div>
      <div class="actions" style="margin-top:8px">${googleLink(t.merchant, 'Google it ↗')}</div>
      ${tags.map(x => `<span class="tag" dir="auto">${esc(x)}</span>`).join('')}
      ${needsNote ? `<div style="margin-top:10px"><input id="note" placeholder="Optional note, e.g. who this was for" aria-label="Note" value="${esc(ui.note)}"></div>` : ''}
      ${suggested.length ? `<div class="tagline">Suggested</div><div class="grid">${suggested.map(c => btn(c, 'top')).join('')}</div>` : ''}
      <div class="sub" style="margin-top:6px">${esc(sug.why)}</div>
      ${pop.length ? `<div class="tagline">Popular</div><div class="grid">${pop.map(c => btn(c)).join('')}</div>` : ''}
      <div class="actions"><button class="btn" data-act="toggleAll">${ui.showAll ? 'Hide categories' : 'All categories'}</button></div>
      ${ui.showAll ? groups : ''}
      <div class="actions"><button class="btn" data-act="undo" ${ui.history.length ? '' : 'disabled'}>← Back</button><button class="btn" data-act="waive">Waive</button><button class="btn" data-act="skip">Skip for now</button></div></div>`;
  }

  // ---------- All transactions ----------
  function pageAll() {
    const filters = [['all', 'Everything'], ['done', 'Categorized'], ['new', 'To review'], ['waived', 'Waived'], ['ignored', 'Auto-ignored']];
    const s = ui.search.toLowerCase();
    const rows = db.tx.filter(t => (ui.allFilter === 'all' || t.status === ui.allFilter)
      && (!s || `${t.merchant} ${t.amount} ${t.income} ${t.userNote} ${t.category}`.toLowerCase().includes(s)))
      .sort((a, b) => String(b.date).localeCompare(String(a.date)));
    const names = categories().map(c => c.category);
    const status = t => {
      if (t.status === 'done') return esc(t.category.replace('Income: ', 'Income: ')) + (t.by === 'rule' ? ' <span class="sub">(rule)</span>' : '');
      if (t.status === 'waived') return '<span class="status mute">Waived</span>';
      if (t.status === 'ignored') return '<span class="status mute">Auto-ignored</span>';
      return '<span class="status warnS">To review</span>';
    };
    const shown = rows.slice(0, 200);
    return `<h1>All transactions</h1><p class="sub">Search anything, fix a past decision, or restore something you waived.</p>
    <div class="card"><input id="search" placeholder="Search merchant, amount, note…" aria-label="Search" value="${esc(ui.search)}">
    <div class="chips">${filters.map(([k, l]) => `<button class="chip ${ui.allFilter === k ? 'on' : ''}" data-act="allFilter" data-v="${k}">${l}</button>`).join('')}</div>
    <div class="scroll"><table><tr><th>Date</th><th>Merchant</th><th>Source</th><th>Category</th><th class="n">₪</th><th></th></tr>
    ${shown.map(t => `<tr><td>${esc(String(t.date).slice(8))}/${esc(String(t.date).slice(5, 7))}/${esc(String(t.date).slice(2, 4))}</td><td dir="auto"><a href="https://www.google.com/search?q=${encodeURIComponent(t.merchant)}" target="_blank" rel="noopener noreferrer" title="Google it">${esc(t.merchant)}</a>${t.userNote ? `<div class="sub" dir="auto">${esc(t.userNote)}</div>` : ''}</td><td>${esc(nickname(t))}</td><td>${status(t)}</td><td class="n">${t.income > 0 ? `+${money(t.income)}` : money(t.amount)}</td>
      <td>${ui.editing === t.id ? `<select data-act="setCat" data-id="${esc(t.id)}" aria-label="Category"><option value="">Choose…</option><option value="__new">Back to review</option><option value="__waive">Waive</option>${names.map(n => `<option>${esc(n)}</option>`).join('')}</select>` : `<button class="btn" data-act="edit" data-id="${esc(t.id)}">Edit</button>`}</td></tr>`).join('')}
    </table></div>${rows.length > shown.length ? `<p class="sub">Showing the newest ${shown.length} of ${rows.length}. Search to narrow it down.</p>` : ''}</div>`;
  }

  // ---------- Summary ----------
  function yearData(year) {
    const values = db.years[year];
    if (!values) return null;
    const months = ui.quarter ? [(ui.quarter - 1) * 3, (ui.quarter - 1) * 3 + 1, (ui.quarter - 1) * 3 + 2] : Model.MONTHS.map((_, i) => i);
    const num = v => (typeof v === 'number' ? v : 0);
    return Model.parseYearTab(values).rows.map(c => {
      const byMonth = Model.MONTHS.map((_, m) => num(values[c.row] && values[c.row][m + 2]));
      return { ...c, byMonth, total: months.reduce((s, m) => s + byMonth[m], 0), months };
    });
  }
  function pageSummary() {
    const years = Object.keys(db.years).sort();
    const year = ui.year || years[years.length - 1];
    const data = yearData(year);
    if (!data) return '<h1>Summary</h1><div class="card">No data for this year.</div>';
    const prev = yearData(String(Number(year) - 1));
    const out = data.filter(c => !c.income);
    const spent = out.reduce((s, c) => s + c.total, 0);
    const income = data.filter(c => c.income).reduce((s, c) => s + c.total, 0);
    const active = Model.MONTHS.map((_, m) => m).filter(m => data[0].months.includes(m) && out.some(c => c.byMonth[m] !== 0));
    const qs = ['Whole year', 'Jan–Mar', 'Apr–Jun', 'Jul–Sep', 'Oct–Dec'];
    const ranked = out.filter(c => c.total > 0).sort((a, b) => b.total - a.total);
    const compare = ranked.slice(0, 8).map(c => {
      const p = prev && prev.find(x => x.category === c.category);
      const pm = p ? p.byMonth.filter(v => v !== 0) : [];
      const pa = pm.length ? pm.reduce((s, v) => s + v, 0) / pm.length : 0;
      const a = c.total / (active.length || 1);
      return `<tr><td>${esc(c.label)}</td><td class="n">${pa ? whole(pa) : '–'}</td><td class="n">${whole(a)}</td><td class="n">${pa ? `${a >= pa ? '+' : '−'}${Math.abs(Math.round((a - pa) / pa * 100))}%` : '–'}</td></tr>`;
    }).join('');
    return `<div class="row"><h1>Summary</h1><select id="year" style="width:auto" aria-label="Year">${years.map(y => `<option ${y === year ? 'selected' : ''}>${y}</option>`).join('')}</select></div>
    <div class="chips">${qs.map((q, i) => `<button class="chip ${ui.quarter === i ? 'on' : ''}" data-act="quarter" data-v="${i}">${q}</button>`).join('')}</div>
    <div class="tiles"><div class="card tile"><span class="sub">Income</span><b>${whole(income)}</b></div><div class="card tile"><span class="sub">Spent</span><b>${whole(spent)}</b></div>
    <div class="card tile"><span class="sub">Saved</span><b>${income ? `${Math.round((income - spent) / income * 100)}%` : '–'}</b></div><div class="card tile"><span class="sub">Monthly average spend</span><b>${whole(spent / (active.length || 1))}</b></div></div>
    <div class="two"><div class="card"><h2>Where the money went</h2><svg id="donut" class="donut" viewBox="0 0 220 220" role="img" aria-label="Spending share by category" style="display:block;max-width:220px;margin:0 auto 10px"></svg><ul class="legend" id="legend"></ul></div>
    <div class="card"><h2>Biggest categories</h2>${ranked.slice(0, 6).map(c => `<div style="margin-bottom:10px"><div class="row" style="font-size:14px"><span>${esc(c.label)}</span><span class="sub">${whole(c.total)}</span></div><div class="bar"><i style="width:${100 * c.total / ranked[0].total}%"></i></div></div>`).join('') || '<p class="sub">Nothing yet.</p>'}</div></div>
    <div class="card"><h2>Spending by month</h2><svg id="bars" class="bars" viewBox="0 0 640 230" role="img" aria-label="Monthly spending, stacked by type" style="width:100%"></svg><ul class="legend wide" id="barLegend"></ul></div>
    <div class="card"><h2>Compared with last year</h2><div class="scroll"><table><tr><th>Category</th><th class="n">${Number(year) - 1} avg/mo</th><th class="n">${year} avg/mo</th><th class="n">Change</th></tr>${compare}</table></div></div>`;
  }

  const NS = 'http://www.w3.org/2000/svg';
  const mk = (n, a, p) => { const e = document.createElementNS(NS, n); Object.entries(a).forEach(([k, v]) => e.setAttribute(k, v)); p.appendChild(e); return e; };
  function hover(el, text) {
    const tip = document.getElementById('tip');
    el.addEventListener('mousemove', e => { tip.style.display = 'block'; tip.textContent = text; tip.style.left = `${Math.min(e.clientX + 12, innerWidth - 190)}px`; tip.style.top = `${e.clientY - 34}px`; });
    el.addEventListener('mouseleave', () => { tip.style.display = 'none'; });
  }
  function drawCharts() {
    const don = document.getElementById('donut');
    if (!don) return;
    const years = Object.keys(db.years).sort();
    const data = yearData(ui.year || years[years.length - 1]).filter(c => !c.income);
    const ranked = data.filter(c => c.total > 0).sort((a, b) => b.total - a.total);
    let slices = ranked.slice(0, 7).map(c => [c.label, c.total]);
    const rest = ranked.slice(7).reduce((s, c) => s + c.total, 0);
    if (rest > 0) slices = slices.concat([['Other', rest]]);
    const total = slices.reduce((s, x) => s + x[1], 0);
    let ang = -Math.PI / 2;
    const legend = document.getElementById('legend');
    slices.forEach(([n, v], k) => {
      const pct = Math.round(v / total * 100);
      if (slices.length > 1) {
        const sw = v / total * 2 * Math.PI;
        const a0 = ang + 0.01; const a1 = ang + sw - 0.01; const R = 105; const r = 66; const lg = sw > Math.PI ? 1 : 0;
        const pt = (rad, a) => `${110 + rad * Math.cos(a)} ${110 + rad * Math.sin(a)}`;
        hover(mk('path', { d: `M${pt(R, a0)}A${R} ${R} 0 ${lg} 1 ${pt(R, a1)}L${pt(r, a1)}A${r} ${r} 0 ${lg} 0 ${pt(r, a0)}Z`, fill: `var(--s${k + 1})` }, don), `${n}: ${whole(v)} (${pct}%)`);
        ang += sw;
      }
      legend.insertAdjacentHTML('beforeend', `<li><i style="background:var(--s${k + 1})"></i><span>${esc(n)}</span><span>${whole(v)} · ${pct}%</span></li>`);
    });
    mk('text', { x: 110, y: 107, 'text-anchor': 'middle', 'font-size': 12, fill: 'var(--ink2)' }, don).textContent = 'Total';
    mk('text', { x: 110, y: 127, 'text-anchor': 'middle', 'font-size': 19, 'font-weight': 700, fill: 'var(--ink)' }, don).textContent = whole(total);

    const sv = document.getElementById('bars');
    const groups = [...new Set(data.map(c => c.group))].slice(0, 8);
    const series = groups.map(g => Model.MONTHS.map((_, m) => data.filter(c => c.group === g).reduce((s, c) => s + Math.max(c.byMonth[m], 0), 0)));
    const totals = Model.MONTHS.map((_, m) => series.reduce((s, x) => s + x[m], 0));
    const top = Math.max(1000, ...totals);
    const step = Math.ceil(top / 4000) * 1000;
    const h = 190; const sc = (h - 6) / (step * 4);
    for (let i = 0; i <= 4; i++) {
      const y = h + 6 - i * step * sc;
      mk('line', { x1: 34, x2: 636, y1: y, y2: y, stroke: 'var(--line)' }, sv);
      mk('text', { x: 28, y: y + 4, 'text-anchor': 'end', 'font-size': 11, fill: 'var(--ink2)' }, sv).textContent = i ? `${i * step / 1000}k` : '0';
    }
    Model.MONTHS.forEach((name, j) => {
      const x = 40 + j * 50;
      let y0 = h + 6;
      series.forEach((d, k) => {
        if (!d[j]) return;
        const hh = d[j] * sc;
        hover(mk('rect', { x, y: y0 - hh + 1, width: 34, height: Math.max(hh - 2, 1), rx: 2, fill: `var(--s${k + 1})` }, sv), `${name} · ${groups[k]}: ${whole(d[j])}`);
        y0 -= hh;
      });
      mk('text', { x: x + 17, y: h + 24, 'text-anchor': 'middle', 'font-size': 11, fill: 'var(--ink2)' }, sv).textContent = name.slice(0, 3);
    });
    document.getElementById('barLegend').innerHTML = groups.map((g, k) => `<li><i style="background:var(--s${k + 1})"></i><span>${esc(g)}</span></li>`).join('');
  }

  // ---------- Data ----------
  function pageData() {
    const r = lastRead();
    const head = `<h1>Data &amp; refresh</h1><p class="sub">Your bank and card folders are only ever read. All updates go to the app's copy.</p>
      <div class="card row"><div><b>${r ? `Last read: ${new Date(r.at).toLocaleString('en-GB')}` : 'Not read yet'}</b>${r ? `<div class="sub">${r.files} files · ${r.total} transactions · ${r.added} new that time</div>` : ''}</div><button class="btn primary" data-act="reread">Re-read Drive</button></div>`;
    if (!r) return head;
    const errors = r.errors.map((e, i) => `<div class="card alert"><div class="row"><h2 style="margin:0">⚠ Could not read this file</h2><span class="status badS">Action needed</span></div>
      <p style="margin:8px 0" dir="auto">${esc(e.message)}</p><p class="sub">Nothing from this file was imported. Your other data is safe.</p>
      ${e.found ? `<pre dir="auto">Found: ${esc(e.found.join(' | '))}\nExpected: ${esc(e.expected.join(' | '))}</pre>
        ${e.expected.filter(x => !e.found.includes(x)).map(x => `<label class="sub">Which column is "${esc(x)}"?</label><select data-map="${i}" data-expected="${esc(x)}">${e.found.filter(f => !e.expected.includes(f)).map(f => `<option>${esc(f)}</option>`).join('')}</select>`).join('')}
        <div class="actions"><button class="btn primary" data-act="applyMap" data-i="${i}">Use this and re-read</button></div>` : ''}
      <div class="actions"><button class="btn" data-act="skipFile" data-v="${esc(e.file)}">Skip this file</button><button class="btn" data-act="copyErr" data-i="${i}">Copy error</button></div></div>`).join('');
    const status = x => {
      if (x.items === 0 && (x.issuer === 'max' || /\d/.test(x.issuer))) return '<span class="status warnS">Export later</span>';
      if (x.items === 0) return '<span class="status mute">No card file for this month</span>';
      if (Math.abs(x.diff) < 0.01) return '<span class="status ok">✓ Matches</span>';
      return `<span class="status badS">Differs ${esc(x.diff.toFixed(2))}</span>`;
    };
    const recon = r.reconcile.slice().reverse().map(x => `<tr><td>${esc(x.month)}</td><td>${esc(x.issuer)}</td><td class="n">${x.bank.toFixed(2)}</td><td class="n">${x.items.toFixed(2)}</td><td>${status(x)}</td></tr>`).join('');
    return `${head}${errors}
      <div class="card"><h2>Does each charge add up?</h2><p class="sub" style="margin-top:0">Bank charges compared with the card items behind them. CAL installments have no charge date in the export, so a difference there can be normal.</p>
      <div class="scroll"><table><tr><th>Month</th><th>Card</th><th class="n">Bank</th><th class="n">Items</th><th>Status</th></tr>${recon}</table></div></div>
      <div class="card"><h2>Duplicates</h2><p class="sub" style="margin-top:0">${r.duplicateCount} rows appeared in more than one file and were kept once.</p>
      ${r.possible.map(p => `<p dir="auto"><span class="status warnS">Check</span> ${esc(p.a)} / ${esc(p.b)}, ${money(p.amount)}, ${shortDate(p.date)}: same card, date and amount but different names. Waive one if they're the same purchase.</p>`).join('')}
      ${r.duplicates.length ? `<details><summary class="sub">Show them</summary><div class="scroll"><table>${r.duplicates.map(d => `<tr><td>${shortDate(d.date)}</td><td dir="auto">${esc(d.merchant)}</td><td class="n">${money(d.amount)}</td><td class="sub">${esc(d.file)}</td></tr>`).join('')}</table></div></details>` : ''}</div>
      <div class="card"><h2>Coverage</h2><div class="scroll"><table><tr><th>Source</th><th>From</th><th>To</th></tr>${Object.entries(r.coverage).map(([k, c]) => `<tr><td>${esc(k)}</td><td>${shortDate(c.from)}</td><td>${shortDate(c.to)}</td></tr>`).join('')}</table></div></div>`;
  }

  // ---------- Settings ----------
  function pageSettings() {
    const kinds = { cardCharge: 'Ignore: card charge', investment: 'Ignore: investment', category: 'Category' };
    const cats = categories();
    const groups = [...new Set(cats.map(c => c.group))];
    return `<h1>Settings</h1>
    <div class="card"><h2>Card nicknames</h2><div class="scroll"><table><tr><th>Card</th><th>Nickname</th></tr>
      ${db.cards.map((c, i) => `<tr><td>${esc(c.issuer)} ·${esc(c.card)}</td><td><input data-nick="${i}" value="${esc(c.nickname)}" aria-label="Nickname for ${esc(c.card)}"></td></tr>`).join('')}</table></div>
      <div class="actions"><button class="btn primary" data-act="saveCards">Save nicknames</button></div></div>
    <div class="card"><h2>Bank line rules</h2><p class="sub" style="margin-top:0">Applied to new bank lines the next time you re-read.</p><div class="scroll"><table><tr><th>Description contains</th><th>Treat as</th><th></th></tr>
      ${db.rules.map((r, i) => `<tr><td dir="auto">${esc(r.match)}</td><td>${esc(kinds[r.kind] || r.kind)}${r.category ? `: ${esc(r.category)}` : ''}</td><td><button class="btn" data-act="delRule" data-i="${i}">Remove</button></td></tr>`).join('')}</table></div></div>
    <div class="card"><h2>Categories</h2><p class="sub" style="margin-top:0">These come from the rows of the year tabs in your copy.</p>${groups.map(g => `<p class="sub" style="margin:6px 0"><b>${esc(g)}:</b> ${cats.filter(c => c.group === g).map(c => esc(c.label)).join(', ')}</p>`).join('')}</div>
    <div class="card"><h2>Your sheet</h2><a href="${esc(api.copyUrl())}" target="_blank" rel="noopener">Open Expenses (App) in Google Sheets</a></div>`;
  }

  // ---------- Render and events ----------
  function render() {
    const needsSignIn = !db;
    const setup = db && db.needsSetup;
    nav.classList.toggle('hide', needsSignIn || setup);
    const banner = `${DEMO ? '<div class="error">Demo mode: made-up data, not connected to Google.</div>' : ''}${ui.busy ? `<div class="card sub">${esc(ui.busy)}</div>` : ''}${ui.error ? `<div class="error" role="alert">${esc(ui.error)}</div>` : ''}`;
    if (needsSignIn) { app.innerHTML = banner + pageSignIn(); return; }
    if (setup) { app.innerHTML = banner + pageSetup(); return; }
    const errs = (lastRead() || { errors: [] }).errors.length;
    nav.innerHTML = PAGES.map(([k, l]) => `<button data-act="page" data-v="${k}" class="${ui.page === k ? 'on' : ''}">${l}${k === 'data' && errs ? ` <span class="badge">${errs}</span>` : ''}</button>`).join('');
    const pages = { triage: pageTriage, all: pageAll, summary: pageSummary, data: pageData, settings: pageSettings };
    // A failing page must say so, and the tabs must keep working, rather than leaving the old page frozen.
    try {
      app.innerHTML = banner + pages[ui.page]();
      drawCharts();
    } catch (e) {
      console.error(e);
      api.log('error', ui.page, e.stack || e.message);
      app.innerHTML = `${banner}<div class="card alert"><h2>This page hit a problem</h2><p class="sub">The other tabs still work and your data is safe. Please send this message to get it fixed:</p><pre>${esc(`${ui.page}: ${e.stack || e.message}`)}</pre></div>`;
    }
  }

  function showPage(page) {
    ui.page = page;
    ui.editing = null;
    window.scrollTo(0, 0);
    if (page !== 'summary') { render(); return; }
    // Totals are calculated by the Sheet, so fetch them after any new decisions.
    guarded('Updating totals…', async () => { db.years = (await api.load()).years; });
  }
  window.addEventListener('popstate', e => {
    if (db && !db.needsSetup) showPage((e.state && e.state.page) || 'triage');
  });

  const currentGroup = () => {
    const all = Model.queue(db.tx, ui.card);
    return (all.filter(g => !ui.skipped.has(g[0].group))[0]) || all[0];
  };
  const actions = {
    signin: signIn,
    create: createCopy,
    reread: rereadDrive,
    page: el => {
      // A history entry per tab, so the phone or browser Back button returns to the last tab.
      if (el.dataset.v !== ui.page) history.pushState({ page: el.dataset.v }, '', `#${el.dataset.v}`);
      showPage(el.dataset.v);
    },
    card: el => { ui.card = el.dataset.v; render(); },
    toggleAll: () => { ui.showAll = !ui.showAll; render(); },
    pick: el => decide(currentGroup(), 'done', el.dataset.v),
    waive: () => decide(currentGroup(), 'waived', ''),
    skip: () => { ui.skipped.add(currentGroup()[0].group); ui.note = ''; ui.showAll = false; render(); },
    undo,
    allFilter: el => { ui.allFilter = el.dataset.v; render(); },
    edit: el => { ui.editing = el.dataset.id; render(); },
    quarter: el => { ui.quarter = Number(el.dataset.v); render(); },
    saveCards: () => guarded('Saving…', async () => {
      document.querySelectorAll('[data-nick]').forEach(i => { db.cards[Number(i.dataset.nick)].nickname = i.value; });
      await api.saveCards(db.cards);
    }),
    delRule: el => guarded('Saving…', async () => {
      db.rules.splice(Number(el.dataset.i), 1);
      await api.saveRules(db.rules);
    }),
    skipFile: el => guarded('Saving…', async () => {
      const list = db.config.skippedFiles ? JSON.parse(db.config.skippedFiles.value) : [];
      list.push(el.dataset.v);
      await api.setConfig('skippedFiles', JSON.stringify(list), db.config.skippedFiles);
      const r = lastRead();
      r.errors = r.errors.filter(e => e.file !== el.dataset.v);
      await api.setConfig('lastRead', JSON.stringify(r), db.config.lastRead);
      db = await api.load();
    }),
    copyErr: el => { navigator.clipboard.writeText(lastRead().errors[Number(el.dataset.i)].message); },
    applyMap: async el => {
      const picks = [...document.querySelectorAll(`[data-map="${el.dataset.i}"]`)];
      await guarded('Saving…', async () => { for (const s of picks) await api.addAlias({ expected: s.dataset.expected, found: s.value }); });
      rereadDrive();
    },
  };

  document.addEventListener('click', e => {
    const el = e.target.closest('[data-act]');
    if (el && el.tagName !== 'SELECT' && actions[el.dataset.act]) actions[el.dataset.act](el);
  });
  document.addEventListener('input', e => {
    if (e.target.id === 'note') ui.note = e.target.value;
    if (e.target.id === 'search') {
      ui.search = e.target.value;
      const pos = e.target.selectionStart;
      render();
      const s = document.getElementById('search');
      s.focus();
      s.setSelectionRange(pos, pos);
    }
  });
  document.addEventListener('change', e => {
    if (e.target.id === 'year') { ui.year = e.target.value; render(); }
    if (e.target.dataset.act === 'setCat') {
      const t = db.tx.find(x => x.id === e.target.dataset.id);
      const v = e.target.value;
      if (!v) return;
      const rows = db.tx.filter(x => x.group === t.group);
      ui.editing = null;
      if (v === '__new') decide(rows, 'new', '');
      else if (v === '__waive') decide(rows, 'waived', '');
      else decide(rows, 'done', v);
    }
  });

  if (DEMO) { db = null; signIn(); } else { render(); }
})();
