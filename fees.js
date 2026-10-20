/* fees.js — fee schedule estimates (v0.11). A TEST feature, Off by default (Config "Fee estimates").
   Pure: no Office.js, no DOM, no network. Loaded after core.js by index.html and dashboard.html, and by tests/.
   1. parseFeeSchedule: an uploaded fee schedule workbook (xlsx.js readWorkbook) -> normalized rows (FEE_COLS).
   2. Matching: next visit location -> fee tab (telehealth: the provider's home clinic), payer -> fee payer column.
   3. estimateMath: allowed amounts -> patient cost (deductible, coinsurance, copay, OOP cap), MD and NP.
   No rate is ever bundled here except the made-up demo schedule (DEMO_WORKBOOK), whose numbers are invented. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.SoharFees = factory(root.SoharCore);
})(typeof self !== 'undefined' ? self : this, function (C) {
  'use strict';

  const str = (v) => (v == null ? '' : String(v).trim());
  const cents = (n) => Math.round((n + (n >= 0 ? 1e-9 : -1e-9)) * 100) / 100;
  const num = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/[$,\s]/g, '');
    if (!/^-?(\d+(\.\d*)?|\.\d+)$/.test(s)) return null;
    return Number(s);
  };
  const keyOf = (s) => str(s).toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').trim();
  const colLetter = (c) => { let s = '', n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }; // 0-based
  const usd = (n) => (n == null ? '' : `$${cents(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);

  // ---------- config ----------
  const FEE_CONFIG_ROWS = [
    { setting: 'Fee estimates', value: 'Off', what: 'Test feature. On: upload a fee schedule (Admin step 4 or the full-screen Benefits tab) and see patient cost estimates for the next visit. Off: the add-in works as before and reads no fee sheets.' },
    { setting: 'Estimate deductible', value: 'Applies', what: 'Fee estimates, until staff change it for a fee payer: Applies or Waived.' },
    { setting: 'Estimate coinsurance', value: 'After deductible', what: 'Fee estimates, until staff change it for a fee payer: After deductible, Before deductible or None.' },
    { setting: 'Estimate copay', value: 'Added', what: 'Fee estimates, until staff change it for a fee payer: Added (only when Sohar returned a copay above $0) or None.' },
  ];
  const TOGGLES = { deductible: ['Applies', 'Waived'], coinsurance: ['After deductible', 'Before deductible', 'None'], copay: ['Added', 'None'] };
  const TOGGLE_DEFAULTS = { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' };
  function pickToggle(kind, v, d) {
    const s = str(v).toLowerCase();
    if (!s) return d;
    return TOGGLES[kind].find((o) => o.toLowerCase() === s) || TOGGLES[kind].find((o) => o.toLowerCase().split(' ')[0] === s.split(' ')[0]) || d;
  }
  const STALE_DAYS = 90;
  /** Config {Setting: Value} -> {on, defaults}. Off unless the row says On / Yes / True. */
  function feeConfig(cfg) {
    const c = cfg || {};
    return {
      on: /^(on|yes|y|true|1)$/i.test(str(c['Fee estimates'])),
      defaults: {
        deductible: pickToggle('deductible', c['Estimate deductible'], TOGGLE_DEFAULTS.deductible),
        coinsurance: pickToggle('coinsurance', c['Estimate coinsurance'], TOGGLE_DEFAULTS.coinsurance),
        copay: pickToggle('copay', c['Estimate copay'], TOGGLE_DEFAULTS.copay),
      },
    };
  }

  // ---------- parsing ----------
  const FEE_COLS = ['Tab', 'Locality', 'Locations', 'Code', 'Description', 'Payer', 'Level', 'Rate', 'Note'];
  const FEE_META_COLS = ['Key', 'Value'];
  const CODE_RE = /^(?:\d{5}|[A-Z]\d{4}|\d{4}[A-Z])$/;
  const ERROR_RE = /^#(ref!|value!|n\/a|div\/0!|name\?|null!|num!|spill!|calc!)/i;

  // A local code (e.g. an in-house IV ketamine code) is one upper-case token of 4-8 letters / digits, never words.
  const LOCAL_CODE_RE = /^[A-Z0-9]{4,8}$/;
  const isCode = (p) => CODE_RE.test(p) || LOCAL_CODE_RE.test(p);
  /**
   * Code cell -> list of codes ('99417/G2212' -> both; '99417G2212', two codes typed together, -> both), [] for a row
   * that is not a code row (a note such as "Mail paper claims, not portal"), null when blank.
   */
  function parseCodes(v) {
    const s = typeof v === 'number' ? (Number.isInteger(v) ? String(v) : '') : str(v);
    if (!s) return typeof v === 'number' ? [] : null;
    const parts = s.split(/\s*[/,&;]\s*|\s+or\s+/i).map((x) => x.trim()).filter(Boolean)
      .map((p) => (CODE_RE.test(p.toUpperCase()) ? p.toUpperCase() : p))
      .flatMap((p) => (p.length === 10 && CODE_RE.test(p.slice(0, 5)) && CODE_RE.test(p.slice(5)) ? [p.slice(0, 5), p.slice(5)] : [p]));
    return parts.length && parts.every(isCode) ? [...new Set(parts)] : [];
  }

  /** A rate cell -> {rate} or {rate: null, note, kind}. 0, NA, errors, percentages and text are not rates. */
  function parseRate(v) {
    if (v == null || (typeof v === 'string' && !v.trim())) return { rate: null, note: '', kind: 'blank' };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { rate: null, note: String(v), kind: 'text' };
      return v > 0 ? { rate: Math.round(v * 10000) / 10000, note: '' } : { rate: null, note: String(v), kind: v === 0 ? 'zero' : 'text' };
    }
    if (typeof v === 'boolean') return { rate: null, note: String(v), kind: 'text' };
    const s = str(v);
    if (ERROR_RE.test(s)) return { rate: null, note: s, kind: 'error' };
    if (/^n\s*\/?\s*a\.?$/i.test(s)) return { rate: null, note: s, kind: 'na' };
    if (s.includes('%')) return { rate: null, note: s, kind: 'percent' };
    const n = num(s);
    if (n == null) return { rate: null, note: s, kind: 'text' };
    return n > 0 ? { rate: Math.round(n * 10000) / 10000, note: '' } : { rate: null, note: s, kind: 'zero' };
  }

  /** "Locations: North Bend, East Fork and Old Mill" / "Location: Lakeside. Hill Top" -> names; null when not one. */
  function parseLocations(text) {
    const m = /^\s*locations?\s*:\s*([\s\S]*)$/i.exec(str(text));
    if (!m) return null;
    return m[1].split(/\s*(?:,|;|\.(?=\s|$)|\band\b|&)\s*/i).map((s) => s.trim()).filter((s) => s.length >= 3 && !/^(the|and|st)$/i.test(s));
  }

  function levelOf(v) {
    const s = str(v).toLowerCase();
    if (!s || ERROR_RE.test(s)) return null;
    // One column for both levels ("MD/NP", "MD & Nurse", "NP/MD"): the rate applies to MD and NP.
    const md = /(^|[^a-z])m\.?\s?d\b/.test(s), np = /(^|[^a-z])(np\b|n\.p\b|nurse|pmhnp|arnp|aprn)/.test(s);
    if (md && np) return 'Both';
    if (/^m\.?\s?d\b/.test(s)) return 'MD';
    if (/^(nurse|np\b|n\.p\.?|pmhnp|arnp|aprn|nurse practitioner)/.test(s)) return 'NP';
    return null;
  }
  const headerText = (v) => { const s = typeof v === 'number' ? '' : str(v); return ERROR_RE.test(s) ? '' : s.replace(/\s+/g, ' '); };

  /**
   * One worksheet -> {tab, rows, skipped}. tab: {name, title, locality, locations, payers, codes, notes, unparsed}.
   * Layout (header offsets differ per tab): a row with "CPT Code"; description in the next column; "Gross Charge" and
   * "Cash" columns; then payer names, each over one or two columns, with the next row saying MD and Nurse / NP. A
   * payer name carries across its merged cells (merges) or onto the NP column right after its MD column. Code rows
   * follow the MD / NP row and stop at the first row whose code cell is not a code; that text becomes a tab note.
   */
  function parseFeeSheet(sheet) {
    const name = str(sheet.name), rows = sheet.rows || [], merges = sheet.merges || [];
    if (sheet.hidden) return { skipped: { name, reason: 'hidden tab' } };
    let hr = -1, codeCol = -1;
    for (let r = 0; r < Math.min(rows.length, 40) && hr < 0; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) if (/^cpt\s*codes?$/i.test(str(row[c]))) { hr = r; codeCol = c; break; }
    }
    if (hr < 0) return { skipped: { name, reason: 'no "CPT Code" header row' } };
    let title = '', locations = null;
    for (let r = 0; r < hr; r++) {
      for (const v of rows[r] || []) {
        if (typeof v !== 'string') continue;
        if (!locations) { const l = parseLocations(v); if (l) { locations = l; continue; } }
        if (!title && /fee schedule/i.test(v)) title = str(v).replace(/\s+/g, ' ');
      }
    }
    const lm = /\blocality\s*(\d+)/i.exec(title);
    const locality = lm ? `Locality ${lm[1]}` : '';
    const hdr = rows[hr] || [], lvl = rows[hr + 1] || [];
    const width = Math.max(hdr.length, lvl.length);
    const descCol = codeCol + 1;
    let grossCol = -1, cashCol = -1;
    for (let c = descCol + 1; c < width; c++) {
      const h = headerText(hdr[c]);
      if (grossCol < 0 && /^gross/i.test(h)) grossCol = c;
      else if (cashCol < 0 && /^(cash|self[\s-]*pay)/i.test(h)) cashCol = c;
      if (grossCol >= 0 && cashCol >= 0) break;
    }
    const fixedEnd = Math.max(descCol, grossCol, cashCol);
    // Merged header cells: column -> the first column of its merged range.
    const mergeStart = new Map();
    for (const m of merges) if (m.r1 <= hr && hr <= m.r2 && m.c2 > m.c1) for (let c = m.c1; c <= m.c2; c++) mergeStart.set(c, m.c1);
    const notes = [];
    const payers = []; // {name, cols: [{c, level}], start, level0}
    const byCol = new Map();
    let cur = null, unlabeled = 0;
    const unlabeledCols = []; // {col: 'AP', level: 'NP'}: rate columns skipped because no payer name could be found
    for (let c = fixedEnd + 1; c < width; c++) {
      const h = headerText(hdr[c]), L = levelOf(lvl[c]), ms = mergeStart.get(c);
      let p = null;
      if (ms != null && ms !== c && byCol.has(ms)) p = byCol.get(ms);
      else if (h) {
        const adjacentNp = L === 'NP' && cur && cur.start === c - 1 && cur.level0 === 'MD' && !mergeStart.has(c - 1);
        if (adjacentNp && keyOf(h) === keyOf(cur.name)) p = cur; // same name repeated over its NP column
        else if (adjacentNp && cur.cols.length === 1) { cur.name = `${cur.name} ${h}`; p = cur; } // "Imperial" + "Med Adv"
        else { cur = { name: h, cols: [], start: c, level0: L }; payers.push(cur); p = cur; }
      } else if (L === 'NP' && cur && cur.start === c - 1 && cur.level0 === 'MD' && cur.cols.length === 1) p = cur;
      if (p) byCol.set(c, p);
      if (!L) continue;
      const levels = L === 'Both' ? ['MD', 'NP'] : [L];
      const skip = (why) => { unlabeled++; unlabeledCols.push({ col: colLetter(c), level: L === 'Both' ? 'MD/NP' : L, why }); };
      if (!p) { skip('no payer name'); continue; }
      if (p.cols.some((x) => levels.includes(x.level))) { skip(`second ${levels.join('/')} column under ${p.name}`); continue; }
      levels.forEach((level, i) => p.cols.push({ c, level, shared: i > 0 }));
    }
    const rateCols = [];
    if (grossCol >= 0) rateCols.push({ c: grossCol, payer: 'Gross Charge', level: 'Gross' });
    if (cashCol >= 0) rateCols.push({ c: cashCol, payer: 'Cash', level: 'Cash' });
    const used = new Map();
    const named = [];
    for (const p of payers) {
      if (!p.cols.length) { if (p.name) notes.push(`Column "${p.name}" has no MD or NP label; skipped.`); continue; }
      let nm = p.name, k = 2;
      while (used.has(keyOf(nm))) nm = `${p.name} (${k++})`;
      used.set(keyOf(nm), true);
      named.push(nm);
      for (const col of p.cols) rateCols.push({ c: col.c, payer: nm, level: col.level, shared: col.shared });
    }
    if (!named.length) return { skipped: { name, reason: 'no payer columns labeled MD or NP' } };
    const out = [];
    const unparsed = { zero: 0, na: 0, error: 0, percent: 0, text: 0 };
    const codes = new Set();
    let stopped = false, afterStop = 0;
    const rowText = (row) => row.filter((v) => typeof v === 'string' && str(v)).map(str).join(' ').slice(0, 300);
    const locText = (locations || []).join('; ');
    for (let r = hr + 2; r < rows.length; r++) {
      const row = rows[r] || [];
      const list = parseCodes(row[codeCol]);
      if (list === null) { if (stopped && rowText(row)) notes.push(rowText(row)); continue; }
      if (stopped) { if (list.length) afterStop++; else if (rowText(row)) notes.push(rowText(row)); continue; }
      if (!list.length) { stopped = true; if (rowText(row)) notes.push(rowText(row)); continue; }
      const desc = str(row[descCol]).replace(/\s+/g, ' ');
      for (const code of list) codes.add(code);
      for (const rc of rateCols) {
        const pr = parseRate(row[rc.c]);
        if (pr.kind === 'blank') continue;
        if (pr.rate == null && !rc.shared) unparsed[pr.kind] = (unparsed[pr.kind] || 0) + 1; // an MD/NP cell counts once
        for (const code of list) {
          out.push({ Tab: name, Locality: locality, Locations: locText, Code: code, Description: desc, Payer: rc.payer, Level: rc.level, Rate: pr.rate == null ? '' : pr.rate, Note: pr.note });
        }
      }
    }
    return {
      rows: out,
      tab: { name, title, locality, locations: locations || [], payers: named, codes: codes.size, notes: notes.slice(0, 20), unlabeled, unlabeledCols, unparsed, afterStop, hasGross: grossCol >= 0, hasCash: cashCol >= 0 },
    };
  }

  /** Workbook ({sheets}) -> {rows, tabs, skipped, counts}. Visible tabs only; tabs without MD / NP payer columns are skipped. */
  function parseFeeSchedule(wb) {
    const rows = [], tabs = [], skipped = [];
    for (const s of (wb && wb.sheets) || []) {
      const res = parseFeeSheet(s);
      if (res.skipped) { skipped.push(res.skipped); continue; }
      tabs.push(res.tab);
      for (const r of res.rows) rows.push(r);
    }
    const codes = new Set(rows.map((r) => r.Code));
    const unparsed = tabs.reduce((n, t) => n + Object.values(t.unparsed).reduce((a, b) => a + b, 0), 0);
    return { rows, tabs, skipped, counts: { tabs: tabs.length, codes: codes.size, rows: rows.length, unparsed } };
  }

  /**
   * Meta rows (Key | Value) saved with the schedule: file, upload time, counts, the schedule's row count (to catch a
   * schedule and info saved from different uploads), tab summary with skipped columns and cells that were not rates.
   */
  function feeMetaRows(fileName, uploadedIso, parsed) {
    const tabs = (parsed.tabs || []).map((t) => ({ name: t.name, title: t.title, locality: t.locality, locations: t.locations, payers: t.payers, codes: t.codes, notes: t.notes,
      unlabeledCols: t.unlabeledCols || [], unparsed: t.unparsed || {} }));
    return [
      { Key: 'File', Value: str(fileName) }, { Key: 'Uploaded', Value: str(uploadedIso) },
      { Key: 'Tabs', Value: String(parsed.counts.tabs) }, { Key: 'Codes', Value: String(parsed.counts.codes) },
      { Key: 'Rows', Value: String((parsed.rows || []).length) },
      { Key: 'Tab Summary', Value: JSON.stringify(tabs).slice(0, 32000) },
      { Key: 'Skipped', Value: JSON.stringify(parsed.skipped || []).slice(0, 32000) },
    ];
  }
  function readMeta(metaRows) {
    const m = {};
    for (const r of metaRows || []) m[str(r.Key)] = r.Value;
    let tabs = [], skipped = [];
    try { tabs = JSON.parse(m['Tab Summary'] || '[]'); } catch (e) { tabs = []; }
    try { skipped = JSON.parse(m.Skipped || '[]'); } catch (e) { skipped = []; }
    const rows = str(m.Rows) === '' ? null : parseInt(m.Rows, 10);
    return { file: str(m.File), uploaded: str(m.Uploaded), tabs: parseInt(m.Tabs, 10) || tabs.length, codes: parseInt(m.Codes, 10) || 0, rows: isNaN(rows) ? null : rows, tabSummary: tabs, skipped };
  }
  /** False when the info was saved for a different schedule than the rows on the sheet (a half-failed upload). */
  function metaMatches(metaRows, scheduleRowCount) {
    const m = readMeta(metaRows);
    return m.rows == null || m.rows === scheduleRowCount;
  }
  /** Plain-English lines for what a parse skipped: rate columns with no payer name, cells that were not rates. */
  function parseWarnings(tabs) {
    const out = [];
    for (const t of tabs || []) {
      const cols = t.unlabeledCols || [];
      if (cols.length) out.push(`${t.name}: ${cols.length} rate column${cols.length === 1 ? '' : 's'} skipped (${cols.map((c) => `${c.col} ${c.level}, ${c.why}`).join('; ')})`);
    }
    const n = (tabs || []).reduce((a, t) => a + Object.values(t.unparsed || {}).reduce((x, y) => x + y, 0), 0);
    if (n) out.push(`${n} rate cell${n === 1 ? '' : 's'} not read as rates (0, NA, errors, percentages or text)`);
    return out;
  }
  /** "Fee schedule: <file>, uploaded <date>, N tabs, N codes", a warning when older than 90 days or mismatched, and skips. */
  function feeStatus(metaRows, scheduleRowCount, now) {
    const m = readMeta(metaRows);
    if (!m.file || !scheduleRowCount) return { loaded: false, text: 'No fee schedule uploaded yet.', warn: '', issues: [] };
    const up = new Date(m.uploaded);
    const days = isNaN(up) ? null : Math.floor(((now || new Date()) - up) / 86400000);
    const when = isNaN(up) ? '?' : `${up.getMonth() + 1}/${up.getDate()}/${up.getFullYear()}`;
    const match = metaMatches(metaRows, scheduleRowCount);
    const warn = [
      match ? '' : 'Fee Schedule Info does not match the saved fee schedule rows (an upload may have half-failed). Upload the fee schedule again.',
      days != null && days > STALE_DAYS ? `The fee schedule is ${days} days old. Upload the current one.` : '',
    ].filter(Boolean).join(' ');
    return {
      loaded: true, days, stale: days != null && days > STALE_DAYS, mismatch: !match,
      text: `Fee schedule: ${m.file}, uploaded ${when}, ${m.tabs} tab${m.tabs === 1 ? '' : 's'}, ${m.codes} code${m.codes === 1 ? '' : 's'}`,
      warn, issues: match ? parseWarnings(m.tabSummary) : [],
    };
  }

  /** Normalized rows -> Map tab -> {name, locality, locations, payers: Map name -> {MD: Map, NP: Map}, gross, cash, desc}. */
  function feeIndex(rows, tabSummary) {
    const tabs = new Map();
    const tabOf = (n) => {
      if (!tabs.has(n)) tabs.set(n, { name: n, locality: '', locations: [], payers: new Map(), gross: new Map(), cash: new Map(), desc: new Map() });
      return tabs.get(n);
    };
    for (const r of rows || []) {
      const t = tabOf(str(r.Tab));
      if (!t.locality) t.locality = str(r.Locality);
      if (!t.locations.length && str(r.Locations)) t.locations = str(r.Locations).split(/\s*;\s*/).filter(Boolean);
      const code = str(r.Code).toUpperCase(), level = str(r.Level), rate = num(r.Rate);
      if (!code) continue;
      if (str(r.Description) && !t.desc.has(code)) t.desc.set(code, str(r.Description));
      if (level === 'Gross' || level === 'Cash') { const m = level === 'Gross' ? t.gross : t.cash; if (rate != null && rate > 0 && !m.has(code)) m.set(code, rate); continue; }
      if (level !== 'MD' && level !== 'NP') continue;
      const name = str(r.Payer);
      if (!t.payers.has(name)) t.payers.set(name, { MD: new Map(), NP: new Map() });
      const p = t.payers.get(name);
      if (rate != null && rate > 0 && !p[level].has(code)) p[level].set(code, rate);
    }
    for (const s of tabSummary || []) {
      const t = tabs.get(str(s.name));
      if (t && Array.isArray(s.locations) && s.locations.length) t.locations = s.locations.map(str);
    }
    return tabs;
  }

  // ---------- location -> tab ----------
  const compact = (s) => s.replace(/ /g, '');
  // Whole words of ln that, run together, spell kn without spaces ("El Dorado Hills" ~ "ElDorado"). Never part of a
  // word, so "San Marino" is not "Marin" and "Orangevale" is not "Orange".
  function joinedWords(ln, kn) {
    const target = compact(kn), words = ln.split(' ');
    for (let i = 0; i < words.length; i++) {
      let acc = '';
      for (let j = i; j < words.length && acc.length < target.length; j++) { acc += words[j]; if (acc === target) return true; }
    }
    return false;
  }
  function isTele(text, teleKeywords) {
    const t = str(text).toLowerCase();
    if (!t) return false;
    const kws = (teleKeywords && teleKeywords.length ? teleKeywords : ['tele']).map((k) => str(k).toLowerCase()).filter(Boolean);
    return kws.some((k) => t.includes(k));
  }
  /**
   * A NextGen location -> {tab, why}. Matches each tab's Locations list and tab name, ignoring case and punctuation,
   * as whole words either way ("San Rafael Clinic" ~ "San Rafael"). Longest match wins; a tie between tabs is no
   * match (never a guess). Telehealth words in a Locations list are ignored.
   */
  function matchLocationTab(loc, tabs, teleKeywords) {
    const ln = keyOf(loc);
    if (!ln) return { tab: '', why: 'No location' };
    const hits = new Map();
    for (const t of tabs || []) {
      for (const k of [t.name, ...(t.locations || [])]) {
        if (isTele(k, teleKeywords)) continue;
        const kn = keyOf(k);
        if (kn.length < 2) continue;
        const ok = ` ${ln} `.includes(` ${kn} `) || (compact(kn).length >= 5 && (joinedWords(ln, kn) || joinedWords(kn, ln))) || (ln.length >= 4 && ` ${kn} `.includes(` ${ln} `));
        if (ok) hits.set(t.name, Math.max(hits.get(t.name) || 0, compact(kn).length));
      }
    }
    if (!hits.size) return { tab: '', why: `No fee tab for location ${str(loc)}` };
    const best = Math.max(...hits.values());
    const top = [...hits].filter(([, n]) => n === best).map(([t]) => t);
    if (top.length > 1) return { tab: '', why: `Location ${str(loc)} matches more than one fee tab (${top.join(', ')}); pick one on Fee Locations` };
    return { tab: top[0], why: '' };
  }

  /**
   * Merge an editable mapping table: rows a person changed (Source "manual", or the value differs from Auto Match)
   * are kept as they are and marked manual; every other key gets the automatic value. keys: [{key, label}].
   * auto(label) -> {value, note}. Returns rows sorted by label.
   */
  function mergeEditable(prev, keys, cols, auto) {
    const [kc, vc] = cols;
    const out = new Map();
    for (const r of prev || []) {
      const k = keyOf(r[kc]);
      if (!k || out.has(k)) continue;
      const v = str(r[vc]), a = str(r['Auto Match']);
      if (/^manual/i.test(str(r.Source)) || v !== a) out.set(k, { [kc]: str(r[kc]), [vc]: v, Source: 'manual', 'Auto Match': a, Note: str(r.Note) });
      else out.set(k, { [kc]: str(r[kc]), [vc]: '', Source: 'auto', 'Auto Match': '', Note: '', stale: true });
    }
    for (const { label } of keys) { const k = keyOf(label); if (k && !out.has(k)) out.set(k, { [kc]: str(label), [vc]: '', Source: 'auto', 'Auto Match': '', Note: '', stale: true }); }
    for (const row of out.values()) {
      if (!row.stale) continue;
      delete row.stale;
      const a = auto(row[kc]);
      row[vc] = a.value; row['Auto Match'] = a.value; row.Note = a.note || '';
    }
    return [...out.values()].sort((a, b) => a[kc].localeCompare(b[kc]));
  }

  const FEE_LOCATION_COLS = ['NextGen Location', 'Fee Tab', 'Source', 'Auto Match', 'Note'];
  /** Fee Locations rows for these NextGen locations (and every earlier row). Manual edits win and persist. */
  function feeLocationRows(prev, locations, tabs, teleKeywords) {
    const keys = [...new Set((locations || []).map(str).filter(Boolean))].map((label) => ({ label }));
    return mergeEditable(prev, keys, ['NextGen Location', 'Fee Tab'], (loc) => {
      if (isTele(loc, teleKeywords)) return { value: '', note: 'Telehealth: the provider home clinic picks the fee tab' };
      const m = matchLocationTab(loc, tabs, teleKeywords);
      return { value: m.tab, note: m.why };
    });
  }

  const PROVIDER_HOME_COLS = ['Provider', 'Home Location', 'Source', 'Auto Match', 'Note'];
  /** Provider -> most frequent non-telehealth location in the report ({loc, n} or {loc: '', tie}). */
  function providerHomes(appts, teleKeywords) {
    const counts = new Map();
    for (const a of appts || []) {
      const p = str(a.rendering), loc = str(a.loc);
      if (!p || !loc || isTele(loc, teleKeywords) || isTele(a.event, teleKeywords)) continue;
      if (!counts.has(p)) counts.set(p, new Map());
      const m = counts.get(p); m.set(loc, (m.get(loc) || 0) + 1);
    }
    const out = new Map();
    for (const [p, m] of counts) {
      const best = Math.max(...m.values());
      const top = [...m].filter(([, n]) => n === best).map(([l]) => l).sort();
      out.set(p, top.length === 1 ? { loc: top[0], n: best } : { loc: '', tie: top, n: best });
    }
    return out;
  }
  function providerHomeRows(prev, appts, teleKeywords) {
    const homes = providerHomes(appts, teleKeywords);
    const providers = new Set((appts || []).map((a) => str(a.rendering)).filter(Boolean));
    return mergeEditable(prev, [...providers].map((label) => ({ label })), ['Provider', 'Home Location'], (p) => {
      const h = homes.get(p);
      if (!h) return { value: '', note: 'No in-person visits in the report' };
      if (!h.loc) return { value: '', note: `Tie: ${h.tie.join(' / ')} (${h.n} visits each); pick one` };
      return { value: h.loc, note: `${h.n} in-person visit${h.n === 1 ? '' : 's'} in the report` };
    });
  }

  // ---------- payer -> fee payer column ----------
  // Fee names abbreviate: "Med Adv", "MCR", "BS 65", "TriWest/VA". State words and plan-type words carry no payer identity.
  const FEE_STOP = new Set(['only', 'loc', 'commercial', 'ppo', 'hmo', 'epo', 'pos', 'plus', 'california', 'washington', 'texas', 'georgia',
    'tx', 'wa', 'ga', 'nw', 'northwest', 'socal', 'norcal', 'region', 'rate', 'rates']);
  const FEE_GENERIC = new Set(['health', 'healthcare', 'care', 'northern', 'southern', 'west', 'east', 'behavioral', 'services', 'network', 'group',
    'medical', 'alliance', 'partners', 'options', 'systems', 'national', 'company', 'solutions', 'foundation', 'benefits', 'administrators', 'direct']);
  const MEDICARE_TOK = new Set(['medicare', 'advantage', 'senior', '65', 'mapd', 'dsnp', 'snp']);
  function feeTokens(s) {
    const t = ` ${str(s).toLowerCase().replace(/\(([^)]*)\)/g, ' $1 ')} `
      .replace(/\bmed\.?\s*adv\.?\b/g, ' medicare advantage ').replace(/\b(mcr|mcare)\b/g, ' medicare ').replace(/\bmedi[\s-]?care\b/g, ' medicare ');
    return (C ? C.autoTokens(t) : t.split(/[^a-z0-9]+/)).filter((x) => x && !FEE_STOP.has(x));
  }
  const w = (t) => (FEE_GENERIC.has(t) ? 0.5 : 1);
  const medicareish = (toks) => toks.some((t) => MEDICARE_TOK.has(t));
  /**
   * Patient payer name -> fee payer column on one tab: {feePayer, confidence: High | Medium | Low, note}.
   * Every word of the fee name (or one side of a "/" name) must be in the payer name. The most specific match wins;
   * a tie is Low. A Medicare / Advantage payer only matches a Medicare / Advantage column. Only-generic words: Medium.
   */
  function matchFeePayer(payer, feePayers) {
    const P = new Set(feeTokens(payer));
    if (!P.size) return { feePayer: '', confidence: 'Low', note: 'No payer name' };
    const cands = [];
    for (const name of feePayers || []) {
      // "Beacon/Carelon (2)" is the parser's name for a second column headed "Beacon/Carelon": match it as that name,
      // so the two columns tie and are flagged instead of the first one winning.
      const base = String(name).replace(/\s*\(\d+\)$/, '');
      const variants = [base, ...(base.includes('/') ? base.split('/') : [])].map(feeTokens).filter((v) => v.length);
      let best = null;
      for (const V of variants) {
        const matched = V.filter((t) => P.has(t));
        if (matched.length !== V.length) continue;
        const spec = matched.reduce((a, t) => a + w(t), 0), distinctive = matched.some((t) => !FEE_GENERIC.has(t));
        if (!best || spec > best.spec) best = { spec, distinctive, toks: V };
      }
      if (best) cands.push({ name, ...best });
    }
    if (!cands.length) return { feePayer: '', confidence: 'Low', note: `No fee column matches "${str(payer)}"` };
    const top = Math.max(...cands.map((c) => c.spec));
    const tops = cands.filter((c) => c.spec === top);
    if (tops.length > 1) return { feePayer: '', confidence: 'Low', note: `"${str(payer)}" matches ${tops.map((c) => c.name).join(' and ')}; pick one` };
    const c = tops[0];
    if (medicareish([...P]) && !medicareish(c.toks)) return { feePayer: '', confidence: 'Low', note: `Medicare plan; "${c.name}" is not a Medicare column. Map it on Fee Payer Map` };
    return c.distinctive ? { feePayer: c.name, confidence: 'High', note: '' } : { feePayer: c.name, confidence: 'Medium', note: 'Matched on generic words only; check it' };
  }

  const FEE_PAYER_MAP_COLS = ['Payer', 'Fee Tab', 'Fee Payer', 'Confidence', 'Source', 'Auto Match', 'Note'];
  const pairKey = (payer, tab) => `${keyOf(payer)}|${keyOf(tab)}`;
  /**
   * Fee Payer Map rows: one per payer and fee tab seen. A row a person changed (Source manual, or Fee Payer differs from
   * Auto Match) is kept; a manual row with a blank Fee Tab applies to every tab. Others are auto-matched again.
   */
  function feePayerMapRows(prev, pairs, index) {
    const out = new Map();
    const isManual = (r) => /^manual/i.test(str(r.Source)) || str(r['Fee Payer']) !== str(r['Auto Match']);
    // Payers with a manual every-tab row (blank Fee Tab): that row is the answer on every tab, so no auto rows for them.
    const everyTab = new Set((prev || []).filter((r) => keyOf(r.Payer) && !keyOf(r['Fee Tab']) && isManual(r)).map((r) => keyOf(r.Payer)));
    for (const r of prev || []) {
      const k = pairKey(r.Payer, r['Fee Tab']);
      if (!keyOf(r.Payer) || out.has(k)) continue;
      const v = str(r['Fee Payer']), a = str(r['Auto Match']);
      if (isManual(r)) out.set(k, { Payer: str(r.Payer), 'Fee Tab': str(r['Fee Tab']), 'Fee Payer': v, Confidence: 'Manual', Source: 'manual', 'Auto Match': a, Note: str(r.Note) });
      else if (keyOf(r['Fee Tab']) && !everyTab.has(keyOf(r.Payer))) out.set(k, { Payer: str(r.Payer), 'Fee Tab': str(r['Fee Tab']), stale: true });
    }
    for (const { payer, tab } of pairs || []) {
      const k = pairKey(payer, tab);
      if (keyOf(payer) && keyOf(tab) && !out.has(k) && !everyTab.has(keyOf(payer))) out.set(k, { Payer: str(payer), 'Fee Tab': str(tab), stale: true });
    }
    for (const row of out.values()) {
      if (!row.stale) continue;
      delete row.stale;
      const t = index.get(row['Fee Tab']);
      const m = t ? matchFeePayer(row.Payer, [...t.payers.keys()]) : { feePayer: '', confidence: 'Low', note: `Fee tab ${row['Fee Tab']} is not in the uploaded fee schedule` };
      Object.assign(row, { 'Fee Payer': m.feePayer, Confidence: m.confidence, Source: 'auto', 'Auto Match': m.feePayer, Note: m.note });
    }
    return [...out.values()].sort((a, b) => a.Payer.localeCompare(b.Payer) || a['Fee Tab'].localeCompare(b['Fee Tab']));
  }
  /**
   * The Fee Payer Map row for a payer on a tab. A person's choice beats an automatic one: a manual row for this payer
   * and tab, else a manual row for this payer with a blank tab (every tab), else the automatic row for this payer and tab.
   */
  function lookupFeePayer(rows, payer, tab) {
    const rs = rows || [];
    const manual = (r) => /^manual/i.test(str(r.Source));
    const exact = rs.filter((r) => pairKey(r.Payer, r['Fee Tab']) === pairKey(payer, tab));
    const every = rs.find((r) => keyOf(r.Payer) === keyOf(payer) && !keyOf(r['Fee Tab']));
    return exact.find(manual) || (every && manual(every) ? every : null) || exact[0] || every || null;
  }
  /** A fee payer name as typed -> the tab's column name (case and punctuation ignored), or ''. */
  function resolvePayerColumn(t, name) {
    if (!t || !keyOf(name)) return '';
    return [...t.payers.keys()].find((p) => keyOf(p) === keyOf(name)) || '';
  }

  const FEE_PAYER_LOG_COLS = ['When', 'Payer', 'Fee Tab', 'Old Fee Payer', 'Old Confidence', 'New Fee Payer', 'New Confidence', 'By', 'Note'];
  /** Log rows for Fee Payer Map rows that differ from their last logged state (like the Payer Map Log). */
  function feePayerLog(rows, prevLog, when) {
    const last = new Map();
    for (const l of prevLog || []) last.set(pairKey(l.Payer, l['Fee Tab']), l);
    const out = [];
    for (const r of rows || []) {
      const l = last.get(pairKey(r.Payer, r['Fee Tab']));
      if (l && str(l['New Fee Payer']) === str(r['Fee Payer']) && str(l['New Confidence']) === str(r.Confidence)) continue;
      out.push({ When: str(when), Payer: r.Payer, 'Fee Tab': r['Fee Tab'], 'Old Fee Payer': l ? str(l['New Fee Payer']) : '', 'Old Confidence': l ? str(l['New Confidence']) : '',
        'New Fee Payer': str(r['Fee Payer']), 'New Confidence': str(r.Confidence), By: r.Source === 'manual' ? 'Sheet edit' : 'Auto-match', Note: str(r.Note) });
    }
    return out;
  }

  // ---------- estimate codes ----------
  const ESTIMATE_CODE_COLS = ['Event keyword', 'Codes', 'Source', 'Note'];
  const SEED_NOTE = 'default, review';
  const ESK_NOTE = 'Confirm the dose code: G2082 (up to 56 mg) or G2083 (over 56 mg, e.g. 84 mg)';
  // The longest keyword found in the event wins; on a tie the earlier row wins, so "TMS F/U" is a follow-up visit.
  const ESTIMATE_CODE_SEEDS = [
    ['TMS mapping', '90867 x1', 'TMS initial session with mapping'], ['TMS initial', '90867 x1', 'TMS initial session with mapping'],
    ['motor threshold', '90869 x1', 'Motor threshold re-determination'],
    ['esketamine', 'G2083 x1', ESK_NOTE], ['ESK', 'G2083 x1', ESK_NOTE], ['Spravato', 'G2083 x1', ESK_NOTE],
    ['intake', '90792 x1', 'Psychiatric evaluation with medical services'], ['consult', '90792 x1', 'Psychiatric evaluation with medical services'],
    ['new', '90792 x1', 'New patient evaluation'],
    ['follow up', '99214 x1', 'Established patient visit'], ['F/U', '99214 x1', 'Established patient visit'], ['med management', '99214 x1', 'Established patient visit'],
    ['therapy 60', '90837 x1', 'Psychotherapy, 60 minutes'],
    ['TMS', '90868 x1', 'TMS treatment session'],
  ].map(([k, c, n]) => ({ 'Event keyword': k, Codes: c, Source: SEED_NOTE, Note: n }));

  /** "90868 x1; 90869 x2, G2083" -> {codes: [{code, units}], bad: [...]}. Units default to 1. */
  function parseCodeList(text) {
    const codes = [], bad = [];
    for (const part of str(text).split(/\s*[;,+\n]\s*/).filter(Boolean)) {
      const m = /^([A-Z0-9]{4,8}?)(?:\s*[x×*]\s*(\d+(?:\.\d+)?))?$/i.exec(part.trim());
      if (m && isCode(m[1].toUpperCase())) codes.push({ code: m[1].toUpperCase(), units: m[2] ? Number(m[2]) : 1 });
      else bad.push(part);
    }
    return { codes, bad };
  }
  /** Event -> {codes, keyword, note}: the Estimate Codes row with the longest keyword found in the event as whole words. */
  function codesForEvent(event, codeRows) {
    const ev = ` ${keyOf(event)} `;
    let best = null;
    for (const r of codeRows || []) {
      const k = keyOf(r['Event keyword']);
      if (!k || !ev.includes(` ${k} `)) continue;
      if (!best || k.length > best.k.length) best = { k, r };
    }
    if (!best) return { codes: [], keyword: '', note: '' };
    const p = parseCodeList(best.r.Codes);
    return { codes: p.codes, keyword: str(best.r['Event keyword']), note: [str(best.r.Note), /review/i.test(str(best.r.Source)) ? 'default codes, review' : ''].filter(Boolean).join('; '), bad: p.bad };
  }

  // ---------- defaults remembered per fee payer ----------
  const ESTIMATE_DEFAULT_COLS = ['Fee Payer', 'Deductible', 'Coinsurance', 'Copay', 'Updated'];
  /**
   * Toggle defaults for a fee payer: each toggle staff chose for it (a filled cell on Estimate Defaults), else Config.
   * copayRemembered says whether staff chose the copay; if not, the copay is only added when Sohar returned one > $0.
   */
  function defaultsFor(rows, feePayer, cfgDefaults) {
    const d = { ...TOGGLE_DEFAULTS, ...(cfgDefaults || {}) };
    const r = keyOf(feePayer) ? (rows || []).find((x) => keyOf(x['Fee Payer']) === keyOf(feePayer)) : null;
    if (!r) return { ...d, remembered: false, copayRemembered: false };
    return { deductible: pickToggle('deductible', r.Deductible, d.deductible), coinsurance: pickToggle('coinsurance', r.Coinsurance, d.coinsurance), copay: pickToggle('copay', r.Copay, d.copay),
      remembered: true, copayRemembered: !!str(r.Copay) };
  }
  /**
   * Save the toggles staff changed for a fee payer. changes holds only what staff picked ({deductible: 'Waived'});
   * other toggles keep what the row already had, and stay blank (Config default) when nothing was chosen. So a copay
   * of None shown for one $0-copay patient is never saved for the whole fee payer unless staff picked it.
   */
  function rememberDefaults(rows, feePayer, changes, when) {
    const old = (rows || []).find((x) => keyOf(x['Fee Payer']) === keyOf(feePayer)) || {};
    const ch = changes || {};
    const val = (kind, col) => (str(ch[kind]) ? pickToggle(kind, ch[kind], TOGGLE_DEFAULTS[kind]) : str(old[col]) ? pickToggle(kind, old[col], TOGGLE_DEFAULTS[kind]) : '');
    const row = { 'Fee Payer': str(feePayer), Deductible: val('deductible', 'Deductible'), Coinsurance: val('coinsurance', 'Coinsurance'), Copay: val('copay', 'Copay'), Updated: str(when) };
    const rest = (rows || []).filter((x) => keyOf(x['Fee Payer']) && keyOf(x['Fee Payer']) !== keyOf(feePayer));
    return rest.concat([row]).sort((a, b) => str(a['Fee Payer']).localeCompare(str(b['Fee Payer'])));
  }

  // ---------- estimate math ----------
  /** 0.2, 20 or "20%" -> 0.2; null when unknown. */
  function coinsRate(v) {
    const s = str(v);
    if (!s) return null;
    const n = parseFloat(s.replace(/[%\s]/g, ''));
    if (isNaN(n) || n < 0) return null;
    return s.includes('%') || n > 1 ? n / 100 : n;
  }
  /**
   * lines: [{code, units, allowed}] (allowed per unit, null when there is no rate). ben: {deductibleRemaining,
   * coinsurance, copay, oopRemaining} (null / '' = unknown). toggles: {deductible, coinsurance, copay}.
   * In order: deductible (Applies: min(allowed, remaining)); coinsurance (After deductible: rate x (allowed - deductible
   * part); Before deductible: rate x allowed, no deductible step, and it lowers the remaining deductible); copay added;
   * then capped at the allowed amount and the OOP max remaining; never below 0; cents. Unknown benefits count as $0
   * and are listed in missing.
   */
  function estimateMath(lines, ben, toggles) {
    const t = { ...TOGGLE_DEFAULTS, ...(toggles || {}) };
    const b = ben || {};
    const missing = [];
    const ls = (lines || []).map((l) => ({ code: l.code, units: Number(l.units) > 0 ? Number(l.units) : 1, allowed: l.allowed == null ? null : Number(l.allowed), desc: l.desc || '' }))
      .map((l) => ({ ...l, total: l.allowed == null ? null : cents(l.allowed * l.units) }));
    for (const l of ls) if (l.allowed == null) missing.push(`No rate for ${l.code}`);
    const known = ls.filter((l) => l.total != null);
    const allowed = cents(known.reduce((a, l) => a + l.total, 0));
    const ded = num(b.deductibleRemaining), coins = coinsRate(b.coinsurance), copay = num(b.copay), oop = num(b.oopRemaining);
    const before = t.coinsurance === 'Before deductible';
    let dedPart = 0, coinsPart = 0, copayPart = 0;
    if (t.deductible === 'Applies' && !before) {
      if (ded == null) missing.push('Deductible remaining');
      else dedPart = cents(Math.min(allowed, Math.max(0, ded)));
    }
    if (t.coinsurance !== 'None') {
      if (coins == null) missing.push('Coinsurance');
      else coinsPart = cents(coins * (before ? allowed : allowed - dedPart));
    }
    if (t.copay === 'Added') {
      if (copay == null) missing.push('Copay');
      else copayPart = cents(Math.max(0, copay));
    }
    const subtotal = cents(dedPart + coinsPart + copayPart);
    let patient = subtotal, cap = '';
    if (patient > allowed) { patient = allowed; cap = 'allowed'; }
    if (oop == null) missing.push('OOP max remaining (no cap applied)');
    else if (patient > Math.max(0, oop)) { patient = cents(Math.max(0, oop)); cap = 'oop'; }
    const oopMet = oop != null && oop <= 0;
    if (oopMet) { patient = 0; cap = 'oop'; } // OOP max already met: $0 whatever the other benefits are
    patient = Math.max(0, cents(patient));
    const toward = t.deductible === 'Applies' ? Math.min(before ? coinsPart : dedPart, patient) : 0;
    // No deductible, coinsurance or copay from Sohar at all (and the OOP max not met): a $0 "estimate" would mislead,
    // so only the allowed amount.
    const noBenefits = ded == null && coins == null && copay == null && !oopMet;
    // Some codes have a rate and some do not: the total leaves those codes out, so it is partial.
    const noRate = ls.filter((l) => l.allowed == null).map((l) => l.code);
    const partial = noRate.length > 0 && known.length > 0;
    return {
      lines: ls, allowed, deductible: dedPart, coinsurance: coinsPart, copay: copayPart, subtotal, cap, patient: noBenefits ? null : patient,
      plan: noBenefits ? null : cents(Math.max(0, allowed - patient)), coinsRate: coins, deductibleRemaining: ded, oopRemaining: oop,
      deductibleAfter: ded == null ? null : cents(Math.max(0, ded - toward)), toggles: t, missing, noBenefits, partial, noRate,
    };
  }
  /** "partial: no rate for 90869" when an estimate total leaves out codes with no rate, else ''. */
  function partialNote(e) {
    return e && e.partial && e.patient != null ? `partial: no rate for ${e.noRate.join(', ')}` : '';
  }
  /** True when a benefit (not a rate) was missing from Sohar. */
  const benefitsMissing = (e) => !!(e && e.missing && e.missing.some((m) => !/^No rate for /.test(m)));

  /** Plain-English lines for one estimate, for the pane and the full-screen panel. */
  function estimateText(e) {
    if (!e) return [];
    const out = e.lines.map((l) => (l.allowed == null ? `${l.code} x${l.units}: no rate on the fee schedule` : `${l.code} x${l.units} at ${usd(l.allowed)} = ${usd(l.total)}${l.desc ? ` (${l.desc.slice(0, 60)})` : ''}`));
    out.push(`Allowed: ${usd(e.allowed)}`);
    if (e.noBenefits) {
      out.push('No deductible, coinsurance or copay from Sohar for this plan, so no patient amount. Re-verify, or quote from the allowed amount.');
      return out;
    }
    const t = e.toggles;
    if (t.coinsurance === 'Before deductible') out.push('Deductible: skipped (coinsurance before deductible)');
    else if (t.deductible === 'Waived') out.push('Deductible: waived');
    else out.push(`Deductible: ${usd(e.deductible)}${e.deductibleRemaining == null ? ' (remaining unknown)' : ` of ${usd(e.deductibleRemaining)} remaining`}`);
    const pct = e.coinsRate == null ? '?' : `${Math.round(e.coinsRate * 1000) / 10}%`;
    if (t.coinsurance === 'None') out.push('Coinsurance: none');
    else if (e.coinsRate == null) out.push('Coinsurance: unknown, counted as $0');
    else out.push(`Coinsurance ${pct} ${t.coinsurance === 'Before deductible' ? `x ${usd(e.allowed)}` : `x (${usd(e.allowed)} - ${usd(e.deductible)})`} = ${usd(e.coinsurance)}`);
    out.push(t.copay === 'None' ? 'Copay: none' : `Copay: ${usd(e.copay)}`);
    if (e.cap === 'oop') out.push(`Capped at the OOP max remaining: ${usd(e.oopRemaining)}`);
    if (e.cap === 'allowed') out.push(`Capped at the allowed amount: ${usd(e.allowed)}`);
    out.push(`Patient pays: ${usd(e.patient)} · Plan pays: ${usd(e.plan)}`);
    if (e.partial) out.push(`Partial: no rate for ${e.noRate.join(', ')}, so the total leaves ${e.noRate.length === 1 ? 'it' : 'them'} out.`);
    if (e.deductibleAfter != null && t.deductible === 'Applies') out.push(`Deductible remaining after this visit: ${usd(e.deductibleAfter)}`);
    if (e.missing.length) out.push(`Missing: ${e.missing.join('; ')}. Estimated with what is known.`);
    return out;
  }

  // ---------- per patient ----------
  /**
   * Everything the estimate needs, from the pane's state. Pure. Returns the mapping tables to write back (manual edits
   * kept), new Fee Payer Map Log rows, per-patient contexts and the rate tables they use.
   * inp: {results, scheduleRows, tabSummary, locations, homes, payerMap, payerLog, codes, defaults, appts, teleKeywords,
   * cfgDefaults, when, pers?}
   */
  function prepareEstimates(inp) {
    const tele = inp.teleKeywords;
    const index = feeIndex(inp.scheduleRows || [], inp.tabSummary);
    const tabs = [...index.values()];
    const byPer = new Map();
    for (const r of inp.results || []) if (r['Sent As'] === 'Verification' && !/^Superseded/.test(str(r.Action)) && str(r['Per Nbr'])) byPer.set(str(r['Per Nbr']), r);
    const pers = (inp.pers || [...byPer.keys()]).map(str).filter((p) => byPer.has(p));
    const appts = inp.appts || [];
    const homeRows = providerHomeRows(inp.homes, appts, tele);
    // Home clinics count as locations too: one typed on Provider Home Clinic (often for a telehealth-only provider)
    // may not be any visit's location in this report, and it still needs a fee tab.
    const locationRows = feeLocationRows(inp.locations, [...appts.map((a) => a.loc), ...pers.map((p) => byPer.get(p)['Next Location']), ...homeRows.map((r) => r['Home Location'])], tabs, tele);
    const seeded = !(inp.codes || []).some((r) => str(r['Event keyword']));
    const codeRows = seeded ? ESTIMATE_CODE_SEEDS.map((r) => ({ ...r })) : inp.codes;
    const locRow = new Map(locationRows.map((r) => [keyOf(r['NextGen Location']), r]));
    const homeRow = new Map(homeRows.map((r) => [keyOf(r.Provider), r]));
    const tabFor = (loc) => {
      const r = locRow.get(keyOf(loc));
      const v = r ? str(r['Fee Tab']) : '';
      if (v && index.has(v)) return { tab: v };
      const t = v ? tabs.find((x) => keyOf(x.name) === keyOf(v)) : null;
      if (t) return { tab: t.name };
      if (v) return { tab: '', flag: `Fee tab "${v}" (Fee Locations, ${str(loc)}) is not in the uploaded fee schedule` };
      return { tab: '', flag: (r && r.Note && !/^telehealth/i.test(r.Note)) ? r.Note : `No fee tab for location ${str(loc)}` };
    };
    const ctxs = {};
    const pairs = [];
    for (const per of pers) {
      const r = byPer.get(per);
      const ctx = { per, last: str(r['Last Name']), first: str(r['First Name']), event: str(r['Next Event']), location: str(r['Next Location']),
        provider: str(r['Next Provider']), appt: str(r['Next Appt']), telehealth: false, home: '', tab: '', locality: '', flags: [] };
      if (!ctx.appt) ctx.flags.push('No upcoming visit');
      ctx.telehealth = isTele(ctx.location, tele) || isTele(ctx.event, tele);
      let site = ctx.location;
      if (ctx.telehealth) {
        const h = homeRow.get(keyOf(ctx.provider));
        ctx.home = h ? str(h['Home Location']) : '';
        site = ctx.home;
        if (!ctx.home) ctx.flags.push(`Telehealth visit: no home clinic for ${ctx.provider || 'the provider'} (Provider Home Clinic)`);
      }
      if (site) {
        const t = tabFor(site);
        ctx.tab = t.tab;
        if (t.flag) ctx.flags.push(t.flag);
      } else if (!ctx.telehealth && ctx.appt) ctx.flags.push('No location on the next visit');
      if (ctx.tab) ctx.locality = index.get(ctx.tab).locality;
      const eligible = /^complete\.eligible/i.test(str(r['Sohar Status']));
      ctx.payer = eligible && str(r['Sohar Payer']) ? str(r['Sohar Payer']) : str(r['NextGen Payer']) || str(r['Sohar Payer']);
      ctx.payerSource = eligible && str(r['Sohar Payer']) ? 'Sohar' : 'NextGen';
      if (!ctx.payer) ctx.flags.push('No payer name');
      if (ctx.tab && ctx.payer) pairs.push({ payer: ctx.payer, tab: ctx.tab });
      const oon = r.oon || {};
      ctx.oon = !!oon.outOfNetwork || /\b(oon|out[\s-]*of[\s-]*network)\b/i.test(str(r.Network));
      ctx.benefits = {
        inn: { deductibleRemaining: str(r['Deductible Remaining']), coinsurance: str(r.Coinsurance), copay: str(r.Copay), oopRemaining: str(r['OOP Remaining']) },
        oon: { deductibleRemaining: str(oon['Deductible Remaining']), coinsurance: str(oon.Coinsurance), copay: str(oon.Copay), oopRemaining: str(oon['OOP Remaining']) },
      };
      const ce = codesForEvent(ctx.event, codeRows);
      ctx.codes = ce.codes; ctx.codeKeyword = ce.keyword; ctx.codeNote = ce.note;
      if (!ce.codes.length && ctx.event) ctx.flags.push(`No default codes for "${ctx.event}": add a row on Estimate Codes or enter codes`);
      ctxs[per] = ctx;
    }
    const payerRows = feePayerMapRows(inp.payerMap, pairs, index);
    const logRows = feePayerLog(payerRows, inp.payerLog, inp.when);
    const tables = { rates: {}, gross: {}, desc: {} };
    const addTab = (tab) => {
      if (tables.gross[tab]) return;
      const t = index.get(tab);
      tables.gross[tab] = Object.fromEntries(t.gross);
      tables.desc[tab] = Object.fromEntries(t.desc);
    };
    let flagged = 0, ready = 0;
    for (const ctx of Object.values(ctxs)) {
      ctx.feePayer = ''; ctx.confidence = ''; ctx.rateKey = '';
      if (ctx.tab && ctx.payer) {
        const row = lookupFeePayer(payerRows, ctx.payer, ctx.tab);
        const t = index.get(ctx.tab);
        const col = row ? resolvePayerColumn(t, row['Fee Payer']) : '';
        ctx.confidence = row ? str(row.Confidence) : '';
        // Out of network uses the Gross Charge, so no payer column is needed.
        if (ctx.oon && (!col || ctx.confidence === 'Low')) { /* no flag */ }
        else if (row && str(row['Fee Payer']) && !col) ctx.flags.push(`Fee payer "${row['Fee Payer']}" is not on the ${ctx.tab} tab (Fee Payer Map)`);
        else if (!col || ctx.confidence === 'Low') ctx.flags.push(`Map payer to a fee column: ${ctx.payer} on ${ctx.tab}${row && row.Note ? ` (${row.Note})` : ''}`);
        else {
          ctx.feePayer = col;
          if (ctx.confidence === 'Medium') ctx.flags.push(`Fee payer ${col} matched with medium confidence; check Fee Payer Map`);
          ctx.rateKey = `${ctx.tab}|${col}`;
          if (!tables.rates[ctx.rateKey]) {
            const p = t.payers.get(col), codes = new Set([...p.MD.keys(), ...p.NP.keys()]);
            tables.rates[ctx.rateKey] = Object.fromEntries([...codes].map((c) => [c, { MD: p.MD.get(c) ?? null, NP: p.NP.get(c) ?? null }]));
          }
        }
      }
      if (ctx.tab) addTab(ctx.tab);
      const d = defaultsFor(inp.defaults, ctx.feePayer, inp.cfgDefaults);
      const ben = ctx.oon ? ctx.benefits.oon : ctx.benefits.inn;
      // Copay is Added by default only when Sohar returned one above $0, unless staff chose Added for this fee payer.
      if (!d.copayRemembered && d.copay === 'Added' && !((num(ben.copay) || 0) > 0)) d.copay = 'None';
      ctx.defaults = { deductible: d.deductible, coinsurance: d.coinsurance, copay: d.copay };
      ctx.ok = !!(ctx.tab && (ctx.oon || ctx.feePayer));
      if (ctx.ok && ctx.codes.length) ready++; else flagged++;
    }
    return { locationRows, homeRows, payerRows, logRows, codeRows, seeded, contexts: ctxs, tables, counts: { patients: pers.length, ready, flagged, tabs: tabs.length } };
  }

  /**
   * One patient's estimate: {md, np, flags, basis}. over: {codes, toggles} from the editors. Out of network: allowed is
   * the Gross Charge and out-of-network benefits are used (MD and NP are the same then).
   */
  function estimatePatient(ctx, tables, over) {
    const o = over || {};
    const flags = (ctx.flags || []).slice();
    const codes = o.codes || ctx.codes || [];
    const toggles = { ...(ctx.defaults || TOGGLE_DEFAULTS), ...(o.toggles || {}) };
    if (!ctx.ok) return { md: null, np: null, flags, basis: '' };
    if (!codes.length) return { md: null, np: null, flags: flags.length ? flags : ['Enter codes to estimate'], basis: '' };
    const rates = tables.rates[ctx.rateKey] || {}, gross = tables.gross[ctx.tab] || {}, desc = tables.desc[ctx.tab] || {};
    const lines = (level) => codes.map(({ code, units }) => ({ code, units, desc: desc[code] || '',
      allowed: ctx.oon ? (gross[code] ?? null) : (rates[code] ? rates[code][level] : null) }));
    const ben = ctx.oon ? ctx.benefits.oon : ctx.benefits.inn;
    const md = estimateMath(lines('MD'), ben, toggles), np = estimateMath(lines('NP'), ben, toggles);
    if (!ctx.oon && !np.lines.some((l) => l.allowed != null)) { np.none = `No NP rate for ${ctx.feePayer} on ${ctx.tab}`; }
    if (!ctx.oon && !md.lines.some((l) => l.allowed != null)) { md.none = `No MD rate for ${ctx.feePayer} on ${ctx.tab}`; }
    const basis = ctx.oon ? `Out of network: Gross Charge on ${ctx.tab}, out-of-network benefits` : `${ctx.tab}${ctx.locality ? ` (${ctx.locality})` : ''} · ${ctx.feePayer}`;
    return { md: md.none ? { ...md, patient: null } : md, np: np.none ? { ...np, patient: null } : np, flags, basis };
  }

  // ---------- demo (made-up numbers; the layout mimics a real fee schedule) ----------
  function demoTab(name, opts) {
    const rows = [['Mindful Health Solutions', '', '', '', '', '', opts.locations ? `Locations: ${opts.locations}` : '']];
    rows.push([`${opts.title} Fee Schedule 2026${opts.locality ? ` Locality ${opts.locality}` : ''}`]);
    rows.push([]);
    const hdr = ['', 'CPT Code', 'Description', 'Gross Charge', 'Cash'], lvl = ['', '', '', '', 'Rate'];
    const payers = opts.payers;
    for (const [p, cols] of payers) { hdr.push(p); lvl.push('MD'); if (cols > 1) { hdr.push(''); lvl.push('Nurse'); } }
    rows.push(hdr, lvl);
    // Every made-up amount is whole dollars ending in 51, 71 or 91, plus 11 cents (51.11, 271.11, 1191.11): an obviously fake pattern.
    const fake = (x) => {
      const h = Math.floor(x / 100) * 100;
      let best = h + 51;
      for (const v of [h - 9, h + 51, h + 71, h + 91]) if (v > 0 && Math.abs(v - x) < Math.abs(best - x)) best = v;
      return Math.round((best + 0.11) * 100) / 100;
    };
    const codes = [['90792', 'Psych diagnostic evaluation with medical services', 591.11], ['90837', 'Psychotherapy, 60 minutes', 251.11], ['90867', 'TMS initial with mapping', 451.11],
      ['90868', 'TMS subsequent delivery', 271.11], ['90869', 'TMS motor threshold re-determination', 291.11], ['99214', 'Established patient visit, moderate', 191.11],
      ['G2082', 'Esketamine visit, up to 56 mg', 691.11], ['G2083', 'Esketamine visit, over 56 mg', 951.11], ['99417/G2212', 'Prolonged service, each 15 minutes', 91.11]];
    codes.forEach(([code, desc, gross], i) => {
      const row = ['', code, desc, gross, fake(gross * 0.613)];
      payers.forEach(([, cols], j) => {
        const md = fake(gross * (0.42 + ((j * 7 + i * 3) % 11) / 50)); // made-up numbers
        row.push(i === 4 && j === 1 ? 'NA' : i === 2 && j === 2 ? '#REF!' : md);
        if (cols > 1) row.push(i === 5 && j === 0 ? '0.00' : fake(md * 0.853));
      });
      rows.push(row);
    });
    rows.push(['', 'Mail paper claims to the plan (made-up note)']);
    return { name, hidden: false, rows, merges: [] };
  }
  const DEMO_WORKBOOK = {
    sheets: [
      demoTab('Marin', { title: 'Marin', locality: 52, locations: 'San Rafael', payers: [['Aetna', 2], ['Blue Shield', 2], ['Cigna', 2], ['Anthem', 2], ['Kaiser', 1], ['UMR', 2], ['Imperial Med Adv', 2], ['Carelon/Beacon', 2]] }),
      demoTab('SF', { title: 'San Francisco', locality: 5, locations: 'Castro Street, Telehealth', payers: [['Aetna', 2], ['Blue Shield', 2], ['BS 65', 2], ['Cigna', 2], ['Noridian', 2], ['San Francisco Health Plan', 2]] }),
      demoTab('Orange', { title: 'Orange County', locality: 26, locations: 'Newport Beach, Mission Viejo and Costa Mesa', payers: [['Aetna', 2], ['Blue Shield', 2], ['Cigna', 2], ['Anthem', 2], ['UMR', 2], ['Carelon/Beacon', 2], ['TriWest/VA', 2]] }),
      demoTab('East Bay', { title: 'Oakland', locality: 7, locations: 'Oakland, Walnut Creek', payers: [['Aetna', 2], ['Blue Shield', 2], ['Kaiser', 1], ['Medicare', 2]] }),
      demoTab('Seattle', { title: 'Washington', payers: [['Aetna', 2], ['Premera BC', 2], ['Regence Med Adv', 2]] }),
      { name: 'Summary', hidden: true, rows: [['Made-up summary tab'], ['', 'CPT Code', 'Aetna']], merges: [] },
    ],
  };
  const demoSchedule = () => parseFeeSchedule(DEMO_WORKBOOK);

  return {
    FEE_CONFIG_ROWS, TOGGLES, TOGGLE_DEFAULTS, feeConfig, pickToggle, STALE_DAYS,
    FEE_COLS, FEE_META_COLS, parseCodes, parseRate, parseLocations, parseFeeSheet, parseFeeSchedule, feeMetaRows, readMeta, feeStatus, feeIndex,
    isTele, matchLocationTab, FEE_LOCATION_COLS, feeLocationRows, PROVIDER_HOME_COLS, providerHomes, providerHomeRows,
    feeTokens, matchFeePayer, FEE_PAYER_MAP_COLS, feePayerMapRows, lookupFeePayer, FEE_PAYER_LOG_COLS, feePayerLog,
    ESTIMATE_CODE_COLS, ESTIMATE_CODE_SEEDS, parseCodeList, codesForEvent,
    ESTIMATE_DEFAULT_COLS, defaultsFor, rememberDefaults,
    coinsRate, estimateMath, partialNote, benefitsMissing, estimateText, prepareEstimates, estimatePatient, usd,
    metaMatches, parseWarnings,
    DEMO_WORKBOOK, demoSchedule,
  };
});
