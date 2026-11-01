// Talks to Google. Raw folders are only read; every write goes to the "Expenses (App)" Sheet.
(function (root) {
  const CLIENT_ID = '408903367049-v98i2hqo8p5580k3ai6pabl8vq7di5j2.apps.googleusercontent.com';
  const FOLDER_ID = '137eGnljhO4uYI_e43EjwHPOwiZChjCOK';
  const SOURCE_NAME = 'Expenses.xlsx';
  const COPY_NAME = 'Expenses (App)';
  const SCOPES = 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets';
  const SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets';
  const DRIVE = 'https://www.googleapis.com/drive/v3/files';

  const TX_COLS = ['id', 'group', 'source', 'card', 'date', 'month', 'merchant', 'amount', 'income', 'chargeDate', 'type', 'note', 'payments', 'file', 'status', 'category', 'userNote', 'by', 'at'];
  const TABS = {
    App_Transactions: TX_COLS,
    App_Rules: ['match', 'kind', 'issuer', 'category'],
    App_Merchants: ['match', 'category'],
    App_Cards: ['card', 'issuer', 'nickname'],
    App_Aliases: ['expected', 'found'],
    App_Config: ['key', 'value'],
    App_Log: ['at', 'by', 'level', 'file', 'message'],
  };
  // Proposed rows, inserted inside the summed ranges so existing SUM formulas stretch over them.
  const NEW_ROWS = [['Others', ['Life insurance', 'Pharmacy', 'Cash', 'Bank fees']], ['TV', ['Subscriptions']]];

  let token = null;
  let tokenClient = null;
  let sheetId = null;
  let user = '';

  function signIn() {
    return new Promise((resolve, reject) => {
      tokenClient = tokenClient || google.accounts.oauth2.initTokenClient({ client_id: CLIENT_ID, scope: SCOPES, callback: '' });
      tokenClient.callback = r => {
        if (r.error) return reject(new Error(r.error_description || r.error));
        token = r.access_token;
        resolve();
      };
      tokenClient.error_callback = e => reject(new Error(e.message || 'Sign-in was closed'));
      tokenClient.requestAccessToken({ prompt: '' });
    });
  }

  async function call(url, options = {}) {
    const res = await fetch(url, { ...options, headers: { Authorization: `Bearer ${token}`, ...(options.body ? { 'Content-Type': 'application/json' } : {}) } });
    if (res.status === 401) {
      const e = new Error('Your Google sign-in expired. Please sign in again.');
      e.signIn = true;
      throw e;
    }
    if (!res.ok) throw new Error(`Google error ${res.status}: ${(await res.text()).slice(0, 300)}`);
    return res;
  }
  const json = async (url, options) => (await call(url, options)).json();
  const post = (url, body) => json(url, { method: 'POST', body: JSON.stringify(body) });

  async function listFolder(id) {
    let files = [];
    let pageToken = '';
    do {
      const q = encodeURIComponent(`'${id}' in parents and trashed=false`);
      const page = await json(`${DRIVE}?q=${q}&fields=nextPageToken,files(id,name,mimeType)&pageSize=200${pageToken ? `&pageToken=${pageToken}` : ''}`);
      files = files.concat(page.files);
      pageToken = page.nextPageToken || '';
    } while (pageToken);
    return files;
  }

  const getValues = async (range, render = 'UNFORMATTED_VALUE') => (await json(`${SHEETS}/${sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=${render}`)).values || [];
  const putValues = (data, input = 'RAW') => post(`${SHEETS}/${sheetId}/values:batchUpdate`, { valueInputOption: input, data });
  const appendValues = (tab, values) => post(`${SHEETS}/${sheetId}/values/${tab}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, { values });
  const batch = requests => post(`${SHEETS}/${sheetId}:batchUpdate`, { requests });

  const objects = (rows, cols) => rows.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c] = r[j] === undefined ? '' : r[j]; });
    return o;
  });
  const clearAndWrite = async (tab, cols, items) => {
    await post(`${SHEETS}/${sheetId}/values/${tab}:clear`, {});
    await putValues([{ range: `${tab}!A1`, values: [cols, ...items.map(o => cols.map(c => o[c] ?? ''))] }]);
  };

  async function findCopy() {
    const files = await listFolder(FOLDER_ID);
    const copy = files.find(f => f.name === COPY_NAME && f.mimeType === 'application/vnd.google-apps.spreadsheet');
    sheetId = copy ? copy.id : null;
    return { exists: !!copy, source: files.find(f => f.name === SOURCE_NAME) || null };
  }

  // One-time: copy Expenses, add the app tabs and a 2026 tab, and connect the month cells.
  async function createCopy(seed) {
    const { source, exists } = await findCopy();
    if (exists) throw new Error(`"${COPY_NAME}" already exists in the folder.`);
    if (!source) throw new Error(`${SOURCE_NAME} was not found in the folder.`);
    const made = await post(`${DRIVE}/${source.id}/copy?fields=id`, { name: COPY_NAME, parents: [FOLDER_ID], mimeType: 'application/vnd.google-apps.spreadsheet' });
    sheetId = made.id;

    await batch(Object.keys(TABS).map(title => ({ addSheet: { properties: { title } } })));
    await putValues(Object.entries(TABS).map(([tab, cols]) => ({ range: `${tab}!A1`, values: [cols] })));

    const meta = await json(`${SHEETS}/${sheetId}?fields=sheets.properties`);
    const tab2025 = meta.sheets.map(s => s.properties).find(p => p.title === '2025');
    if (!tab2025) throw new Error('The 2025 tab was not found.');

    // Bottom row first, so earlier inserts don't move the later positions.
    let values = await getValues('2025!A:N', 'FORMULA');
    const places = NEW_ROWS.map(([after, labels]) => ({ at: values.findIndex(r => r[1] === after), labels })).sort((a, b) => b.at - a.at);
    for (const p of places) {
      if (p.at < 0) throw new Error('Could not find where to add the new categories in the 2025 tab.');
      await batch([{ insertDimension: { range: { sheetId: tab2025.sheetId, dimension: 'ROWS', startIndex: p.at, endIndex: p.at + p.labels.length }, inheritFromBefore: true } }]);
      await putValues(p.labels.map((l, i) => ({ range: `2025!A${p.at + i + 1}:N${p.at + i + 1}`, values: [['', l, ...Array(12).fill('')]] })));
    }

    await batch([{ duplicateSheet: { sourceSheetId: tab2025.sheetId, newSheetName: '2026' } }]);
    values = await getValues('2025!A:N', 'FORMULA');
    const { rows } = Model.parseYearTab(values);
    const data = [];
    rows.forEach(cat => {
      const cells2025 = Model.MONTHS.map((_, m) => Model.mergedFormula(values[cat.row][m + 2], cat, `2025-${String(m + 1).padStart(2, '0')}`));
      const cells2026 = Model.MONTHS.map((_, m) => `=${Model.sumFormula(cat, `2026-${String(m + 1).padStart(2, '0')}`)}`);
      data.push({ range: `2025!C${cat.row + 1}:N${cat.row + 1}`, values: [cells2025] });
      data.push({ range: `2026!C${cat.row + 1}:N${cat.row + 1}`, values: [cells2026] });
    });
    await putValues(data, 'USER_ENTERED');

    await clearAndWrite('App_Rules', TABS.App_Rules, seed.bankRules);
    await clearAndWrite('App_Merchants', TABS.App_Merchants, seed.merchantRules);
    await clearAndWrite('App_Cards', TABS.App_Cards, seed.cards);
    await clearAndWrite('App_Config', TABS.App_Config, [{ key: 'alreadyRecorded', value: JSON.stringify(seed.alreadyRecorded || []) }, { key: 'setupComplete', value: 'yes' }]);
  }

  async function load() {
    const ranges = Object.keys(TABS).filter(t => t !== 'App_Log').map(t => `ranges=${t}`).join('&');
    const res = await json(`${SHEETS}/${sheetId}/values:batchGet?${ranges}&valueRenderOption=UNFORMATTED_VALUE`);
    const get = t => objects(res.valueRanges.find(v => v.range.startsWith(t)).values || [], TABS[t]);
    const config = {};
    get('App_Config').forEach(c => { config[c.key] = { value: c.value, _row: c._row }; });
    const years = {};
    const meta = await json(`${SHEETS}/${sheetId}?fields=sheets.properties.title`);
    for (const title of meta.sheets.map(s => s.properties.title).filter(t => /^20\d\d$/.test(t) && t >= '2025')) years[title] = await getValues(`${title}!A:N`);
    return { tx: get('App_Transactions'), rules: get('App_Rules'), merchantRules: get('App_Merchants'), cards: get('App_Cards'), aliases: get('App_Aliases'), config, years };
  }

  // Reads each bank and card file as a table of rows. Files are only downloaded, never changed.
  async function readRawFiles() {
    const top = await listFolder(FOLDER_ID);
    const folders = top.filter(f => f.mimeType === 'application/vnd.google-apps.folder');
    const list = [];
    for (const folder of folders) (await listFolder(folder.id)).forEach(f => list.push({ ...f, folder: folder.name }));
    const out = [];
    for (let i = 0; i < list.length; i += 4) {
      out.push(...await Promise.all(list.slice(i, i + 4).map(async f => {
        if (!/\.xlsx?$/i.test(f.name)) return { name: f.name, folder: f.folder, rows: [], error: 'not an Excel file' };
        try {
          const buf = await (await call(`${DRIVE}/${f.id}?alt=media`)).arrayBuffer();
          const wb = XLSX.read(buf, { type: 'array', cellDates: true });
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
          return { name: f.name, folder: f.folder, rows };
        } catch (e) {
          if (e.signIn) throw e;
          return { name: f.name, folder: f.folder, rows: [], error: `could not open the file: ${e.message}` };
        }
      })));
    }
    return out;
  }

  const txData = rows => rows.map(t => ({ range: `App_Transactions!O${t._row}:S${t._row}`, values: [[t.status, t.category, t.userNote, t.by, t.at]] }));

  root.GoogleBackend = {
    signIn,
    setUser: u => { user = u; },
    findCopy,
    createCopy,
    load,
    readRawFiles,
    copyUrl: () => `https://docs.google.com/spreadsheets/d/${sheetId}/edit`,
    updateTx: rows => putValues(txData(rows)),
    appendTx: rows => appendValues('App_Transactions', rows.map(t => TX_COLS.map(c => t[c] ?? ''))),
    saveCards: items => clearAndWrite('App_Cards', TABS.App_Cards, items),
    saveRules: items => clearAndWrite('App_Rules', TABS.App_Rules, items),
    addAlias: a => appendValues('App_Aliases', [[a.expected, a.found]]),
    setConfig: (key, value, existing) => (existing
      ? putValues([{ range: `App_Config!B${existing._row}`, values: [[value]] }])
      : appendValues('App_Config', [[key, value]])),
    log: (level, file, message) => appendValues('App_Log', [[new Date().toISOString(), user, level, file, message]]).catch(() => {}),
  };
})(window);
