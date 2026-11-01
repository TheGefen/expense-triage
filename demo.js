// Demo mode (?demo=1): in-memory, made-up data, no Google connection.
(function (root) {
  
  const yearTab = () => [
    ['Outcome'], ['Type', 'Details', ...root.Model.MONTHS, 'AVG'],
    ['Essential expenses', 'Groceries', 2000, 2100, 1900, 2200, 2000, 2050, 2100, 1950, 2000, 2100, 0, 0],
    ['', 'Others', 0, 100, 0, 0, 50, 0, 0, 0, 0, 0, 0, 0],
    ['Car', 'Car Gas', 500, 400, 450, 500, 420, 480, 500, 400, 450, 500, 0, 0],
    ['Kids Essential', 'Kids Others', 300, 250, 400, 300, 250, 300, 350, 300, 250, 300, 0, 0],
    ['Luxry', 'Food Out', 600, 700, 500, 800, 600, 700, 650, 600, 700, 800, 0, 0],
    ['', 'Clothing', 100, 0, 200, 0, 100, 0, 150, 0, 100, 0, 0, 0],
    ['', 'Electronics', 0, 300, 0, 0, 0, 500, 0, 0, 200, 0, 0, 0],
    ['Income', '', ...root.Model.MONTHS, 'AVG'],
    ['', 'Wage #1', 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 0, 0],
    ['', 'Others', 0, 0, 200, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['Summary', '', ...root.Model.MONTHS, 'AVG'],
  ];
  const state = {
    tx: [],
    rules: [{ match: 'כרטיס דמו', kind: 'cardCharge', issuer: 'cal' }, { match: 'מעסיק דמו', kind: 'category', category: 'Income: Wage #1' }],
    merchantRules: [{ match: 'סופר דמו', category: 'Groceries' }, { match: 'קפה דמו', category: 'Food Out' }],
    cards: [{ card: '1111', issuer: 'cal', nickname: 'Demo Visa' }, { card: '2222', issuer: 'cal', nickname: 'Demo Online' }, { card: '9999', issuer: 'isracard', nickname: 'Demo Isracard' }],
    aliases: [],
    config: {},
    years: { 2025: yearTab(), 2026: yearTab() },
  };
  const calHead = ['תאריך עסקה', 'שם בית עסק', 'סכום בש"ח', 'כרטיס', 'מועד חיוב', 'סוג עסקה', 'הערות'];
  const calFile = [calHead,
    ['14/10/26', 'קפה דמו', '₪ 38.00', 'ויזה 1111', '10/11/26', 'רגילה', ''],
    ['18/10/26', 'סופר דמו', '₪ 412.30', 'ויזה 1111', '10/11/26', 'רגילה', ''],
    ['20/10/26', 'BIT', '₪ 150.00', 'ויזה 1111', '10/11/26', 'רגילה', ''],
    ['22/1/26', 'רהיטים דמו', '₪ 265.00', 'ויזה 1111', '', 'תשלומים', 'עסקה ב-12 תשלומים'],
    ['22/1/26', 'רהיטים דמו', '₪ 265.00', 'ויזה 1111', '', 'תשלומים', 'עסקה ב-12 תשלומים'],
    ['25/6/26', 'החזר דמו', '-₪ 58.00', 'ויזה 2222', '10/7/26', 'זיכוי', ''],
    ['10/10/26', 'רכב דמו', '₪ 120.00', 'ויזה 2222', '10/11/26', 'רגילה', ''],
  ];
  const brokenFile = [calHead.map(c => (c === 'סכום בש"ח' ? 'סכום חיוב' : c)), ['14/10/26', 'קפה דמו', '₪ 5.00', 'ויזה 1111', '10/11/26', 'רגילה', '']];
  const bankFile = [['', 'יתרה', 'תאריך ערך', 'זכות', 'חובה', 'תיאור', 'אסמכתא', 'סוג פעולה', 'תאריך'],
    ['', '0', '10/10/2026', '', '1,000.00', 'כרטיס דמו', '1', '162', '10/10/2026'],
    ['', '0', '01/10/2026', '9,000.00', '', 'מעסיק דמו', '2', '222', '01/10/2026'],
    ['', '0', '28/04/2026', '', '800.00', 'שיק דמו', '3', '162', '28/04/2026']];

  const renumber = () => state.tx.forEach((t, i) => { t._row = i + 2; });
  root.DemoBackend = {
    signIn: async () => {},
    setUser: () => {},
    findCopy: async () => ({ exists: true }),
    createCopy: async () => {},
    load: async () => JSON.parse(JSON.stringify(state)),
    readRawFiles: async () => [
      { name: 'demo-bank.xlsx', folder: 'Beinleumi', rows: bankFile },
      { name: 'demo-cal.xlsx', folder: 'CAL', rows: calFile },
      { name: 'demo-cal-overlap.xlsx', folder: 'CAL', rows: calFile },
      { name: 'demo-cal-new-format.xlsx', folder: 'CAL', rows: brokenFile },
    ],
    copyUrl: () => '#',
    updateTx: async rows => { rows.forEach(r => { Object.assign(state.tx[r._row - 2], r); }); },
    appendTx: async rows => { state.tx.push(...JSON.parse(JSON.stringify(rows))); renumber(); },
    saveCards: async items => { state.cards = items; },
    saveRules: async items => { state.rules = items; },
    addAlias: async a => { state.aliases.push(a); },
    setConfig: async (key, value) => { state.config[key] = { value }; },
    log: async (level, file, message) => { console.warn('log', level, file, message); },
  };

})(window);
