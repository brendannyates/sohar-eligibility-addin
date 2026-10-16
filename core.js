/* core.js — pure functions, no Office.js, no DOM. Loaded by index.html and by tests/. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SoharCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const VERIFICATION_COLS = ['firstName', 'lastName', 'dateOfBirth', 'state', 'memberId', 'payerId',
    'placeOfServiceCode', 'specialtyCode', 'id', 'subscriberFirstName', 'subscriberLastName', 'subscriberDateOfBirth'];
  const DISCOVERY_COLS = ['firstName', 'lastName', 'dateOfBirth', 'state', 'zipCode', 'payerId', 'id'];

  const PATIENT_COLS = ['Per Nbr', 'MRN', 'Last Name', 'First Name', 'DOB', 'State', 'Zip', 'Pat Sts', 'Route', 'Route Reason',
    'In Scope', 'Next Appt', 'Next Event', 'POS', 'Primary Payer', 'Sohar payerId', 'Confidence', 'Member ID', 'Member ID (Lst Enc)',
    'Member ID Mismatch', 'Secondary Payer', 'Secondary Sohar payerId', 'Appt Count', 'First Appt', 'Last Appt',
    'Subscriber First', 'Subscriber Last', 'Subscriber DOB', 'Flags'];
  const REVIEW_COLS = ['Severity', 'Issue', 'Per Nbr', 'Detail', 'Fix'];
  // Sohar Results sheet. The add-in creates this table itself, so it is not part of the template contract.
  const RESULTS_COLS = ['Next Appt', 'Next Time', 'Next Event', 'Next Location', 'Next Provider', 'Upcoming Appts', 'Upcoming Detail',
    'Per Nbr', 'MRN', 'Last Name', 'First Name', 'DOB', 'Sent As', 'Sohar Status', 'Work Queue', 'Action', 'Choose in NextGen', 'Benefit Flag',
    'Plan', 'Network', 'Coverage Level', 'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance',
    'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status',
    'NextGen Payer', 'NextGen Member ID', 'Sohar Payer', 'Sohar payerId', 'Sohar Member ID', 'Coverage Differs', 'Match',
    'Sohar Ref', 'Sohar Link', 'Sohar Date', 'Discovery Follow-up'];
  // Saved Sohar result records (hidden sheet), so a later upload adds to earlier ones instead of replacing them.
  const UPLOAD_COLS = ['Kind', 'Key', 'File', 'Uploaded', 'Record'];
  // NextGen Tasks sheet (created by the add-in). Done survives rebuilds via Task Key.
  const TASK_COLS = ['Done', 'Due By', 'Next Appt', 'Next Time', 'Next Location', 'Per Nbr', 'MRN', 'Last Name', 'First Name', 'Result Type',
    'Task', 'Details', 'Choose in NextGen', 'Sohar Link', 'Sohar Ref', 'Task Key'];
  // Benefits CSV for the TMS Tracker. Treat as a contract: add columns at the end only.
  const BENEFITS_EXPORT_COLS = ['Per Nbr', 'MRN', 'Last Name', 'First Name', 'DOB', 'Sohar Payer', 'Plan', 'Network', 'Coverage Level',
    'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status',
    'OOP Max', 'OOP Remaining', 'OOP Status', 'Benefit Flag', 'Sohar Status', 'Sohar Date', 'Sohar Ref'];

  // NextGen header -> canonical field. Matching ignores case, spaces and punctuation.
  const HEADER_SYNONYMS = {
    patName: ['Pat Name', 'Patient Name', 'Patient'],
    dob: ['Birth Dt', 'DOB', 'Birth Date', 'Date of Birth', 'Pat DOB'],
    perNbr: ['Per Nbr', 'Person Nbr', 'Person Number', 'Person #', 'Person No'],
    mrn: ['Md Rc', 'MRN', 'Medical Record', 'Med Rec Nbr', 'Chart Nbr'],
    apptDt: ['Appt Dt', 'Appointment Date', 'Appt Date', 'Date'],
    begTm: ['Beg Tm', 'Start Time', 'Begin Time', 'Time'],
    event: ['Event', 'Appt Type', 'Appointment Type', 'Visit Type'],
    loc: ['Loc Name', 'Location', 'Location Name', 'Practice'],
    sts: ['Sts', 'Status', 'Appt Status', 'Appointment Status'],
    patSts: ['Pat Sts', 'Patient Status'],
    payer: ['Payer Name (Cht Guar)', 'Payer Name', 'Primary Payer', 'Primary Insurance', 'Payer'],
    secPayer: ['Sec Payer (Cht Guar)', 'Secondary Payer', 'Sec Payer', 'Secondary Insurance'],
    polNbr: ['Pol Nbr (Cht Guar)', 'Policy Nbr', 'Policy Number', 'Member ID', 'Subscriber ID', 'Pol Nbr'],
    polNbrLstEnc: ['Pol Nbr (Lst Enc)', 'Policy Nbr (Last Enc)'],
    zip: ['Zip', 'Pat Zip', 'Zip Code', 'Postal Code', 'Patient Zip', 'Home Zip', 'Guar Zip'],
    state: ['Pat State', 'Patient State', 'Home State', 'State'],
    guarName: ['Guar Name', 'Guarantor Name', 'Guarantor'],
    rendering: ['Rendering', 'Rendering Provider', 'Provider'],
    resName: ['Res Name', 'Resource Name', 'Resource'],
    finClass: ['Fin Class (Cht Guar)', 'Fin Class', 'Financial Class'],
    finClassLstEnc: ['Fin Class (Lst Enc)'],
    secPayerLstEnc: ['Sec Payer (Lst Enc)'],
    authCd: ['Auth Cd (Cht Guar)', 'Auth Cd', 'Auth Code'],
    authCdLstEnc: ['Auth Cd (Lst Enc)'],
  };
  // Matching is exact after normHeader, so "Nxt Appt Dt", "Nxt Beg Tm" and "Nxt Event" never land on apptDt/begTm/event.
  const REQUIRED = ['perNbr', 'apptDt', 'payer'];
  const FOOTER_RE = /^total appointments/i;

  function normHeader(s) {
    return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  }
  function isBlank(v) {
    return v == null || String(v).trim() === '';
  }
  function str(v) {
    return isBlank(v) ? '' : String(v).trim();
  }

  /**
   * Find the header row in a pasted report. Returns {headerRow, colMap, headers, missing, optionalMissing,
   * footerRow, dataEnd, leftoverRows, warnings} or null.
   * Data stops at the first row after the header whose first non-blank cell starts with "Total Appointments"
   * (footerRow; -1 when there is none). dataEnd is the exclusive end index for data rows.
   * leftoverRows counts rows with content below the footer (left over from pasting over a longer report).
   */
  function detectHeader(rows, maxScan) {
    const limit = Math.min(rows.length, maxScan || 40);
    const synIndex = {};
    for (const [field, names] of Object.entries(HEADER_SYNONYMS)) {
      for (const n of names) {
        const k = normHeader(n);
        if (!synIndex[k]) synIndex[k] = [];
        synIndex[k].push(field);
      }
    }
    let best = null;
    for (let r = 0; r < limit; r++) {
      const row = rows[r] || [];
      const colMap = {};
      let hits = 0;
      for (let c = 0; c < row.length; c++) {
        const k = normHeader(row[c]);
        if (!k || !synIndex[k]) continue;
        for (const f of synIndex[k]) {
          if (colMap[f] == null) { colMap[f] = c; hits++; }
        }
      }
      if (colMap.perNbr != null && colMap.apptDt != null && (!best || hits > best.hits)) {
        best = { headerRow: r, colMap, hits, headers: row.map(str) };
      }
    }
    if (!best) return null;
    best.missing = REQUIRED.filter((f) => best.colMap[f] == null);
    best.optionalMissing = ['patName', 'dob', 'zip', 'polNbr', 'event', 'loc', 'sts'].filter((f) => best.colMap[f] == null);
    best.warnings = [];
    best.footerRow = -1;
    for (let r = best.headerRow + 1; r < rows.length; r++) {
      const first = (rows[r] || []).find((v) => !isBlank(v));
      if (first != null && FOOTER_RE.test(str(first))) { best.footerRow = r; break; }
    }
    best.dataEnd = best.footerRow >= 0 ? best.footerRow : rows.length;
    best.leftoverRows = 0;
    if (best.footerRow >= 0) {
      for (let r = best.footerRow + 1; r < rows.length; r++) if ((rows[r] || []).some((v) => !isBlank(v))) best.leftoverRows++;
    }
    if (best.leftoverRows) {
      const n = best.leftoverRows;
      best.warnings.push(`${n} leftover row${n > 1 ? 's' : ''} below the report total ${n > 1 ? 'were' : 'was'} ignored — upload the report instead of pasting, which clears the sheet.`);
    }
    // Sanity check: provider names in the payer column mean the columns are shifted.
    const cm = best.colMap;
    if (cm.payer != null && (cm.rendering != null || cm.resName != null)) {
      const staff = new Set(), payers = new Set();
      for (let r = best.headerRow + 1; r < best.dataEnd; r++) {
        const row = rows[r] || [];
        for (const f of ['rendering', 'resName']) if (cm[f] != null && !isBlank(row[cm[f]])) staff.add(str(row[cm[f]]).toLowerCase());
        if (!isBlank(row[cm.payer])) payers.add(str(row[cm.payer]));
      }
      const overlap = [...payers].filter((p) => staff.has(p.toLowerCase()));
      best.shiftedColumns = overlap;
      if (overlap.length) {
        best.warnings.push(`The payer column holds provider names (${overlap.slice(0, 3).join('; ')}${overlap.length > 3 ? '; …' : ''}); the columns look shifted. Re-run the report and upload it again.`);
      }
    } else best.shiftedColumns = [];
    return best;
  }
  const dataEndOf = (det, rows) => (det.dataEnd != null ? det.dataEnd : rows.length);

  // ---------- dates ----------
  const MS_DAY = 86400000;
  function excelSerialToDate(n) {
    // Excel 1900 system, serial 1 = 1900-01-01; 25569 = 1970-01-01 (UTC)
    return new Date(Math.round((n - 25569) * MS_DAY));
  }
  function parseDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : dayOf(v);
    if (typeof v === 'number') return v > 59 ? excelSerialToDate(v) : null;
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) {
      let y = +m[3]; if (y < 100) y += y < 50 ? 2000 : 1900;
      return new Date(Date.UTC(y, +m[1] - 1, +m[2]));
    }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    const d = new Date(s);
    return isNaN(d) ? null : utcDay(d);
  }
  function utcDay(d) {
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }
  /** A Date already at UTC midnight is a day value (from parseDate or an as-of picker): keep it. Else take its local day. */
  function dayOf(d) {
    return d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && d.getUTCMilliseconds() === 0
      ? new Date(d.getTime()) : utcDay(d);
  }
  function fmtDate(d) {
    if (!d) return '';
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return `${mm}/${dd}/${d.getUTCFullYear()}`;
  }

  // ---------- names ----------
  const SUFFIX = /^(jr|sr|ii|iii|iv|v)\.?$/i;
  /** "Last, First Middle" -> {firstName, lastName}. "First Last" handled; multi-token without comma is flagged. */
  function parseName(raw) {
    const s = str(raw).replace(/\s+/g, ' ');
    if (!s) return { firstName: '', lastName: '', flag: 'Missing name' };
    if (s.includes(',')) {
      const [last, rest] = s.split(',', 2).map((x) => x.trim());
      const parts = (rest || '').split(' ').filter(Boolean);
      let lastName = last;
      // "Smith, John Jr" -> suffix belongs to last name
      if (parts.length > 1 && SUFFIX.test(parts[parts.length - 1])) lastName = `${last} ${parts.pop()}`;
      return { firstName: parts[0] || '', lastName, flag: parts[0] ? '' : 'First name missing' };
    }
    const parts = s.split(' ').filter(Boolean);
    if (parts.length === 1) return { firstName: '', lastName: parts[0], flag: 'Single-token name' };
    if (parts.length === 2) return { firstName: parts[0], lastName: parts[1], flag: '' };
    return { firstName: parts[0], lastName: parts.slice(-1)[0], flag: 'Name has no comma; check first/last split' };
  }

  // ---------- payer map ----------
  const VERIFIABLE = new Set(['exact', 'probable', 'confirmed', 'auto-high', 'auto-medium']);
  // Confident enough to stand in for a Fin Class when deriving the last-encounter payer (Auto-Medium is not).
  const CONFIDENT = new Set(['exact', 'probable', 'confirmed', 'auto-high']);
  const isVip = (s) => /\bvip\b/i.test(str(s));
  const isSelfPay = (s) => /^self[\s-]*pay\b/i.test(str(s));
  function payerMapIndex(mapRows) {
    if (mapRows instanceof Map) return mapRows;
    const idx = new Map();
    for (const r of mapRows || []) {
      const name = str(r.nextgenPayerName);
      if (!name) continue;
      idx.set(name.toLowerCase(), {
        payerId: str(r.payerId), payerName: str(r.payerName),
        confidence: str(r.confidence), notes: str(r.notes),
      });
    }
    return idx;
  }
  function lookupPayer(name, idx) {
    if (isBlank(name)) return { status: 'none', payerId: '', confidence: 'No payer on file' };
    const hit = idx.get(str(name).toLowerCase());
    if (!hit) return { status: 'unmapped', payerId: '', confidence: 'Unmapped' };
    const ok = VERIFIABLE.has(hit.confidence.toLowerCase()) && hit.payerId;
    return { status: ok ? 'mapped' : 'review', ...hit };
  }

  /** Score Sohar payers against a NextGen name; returns top n {payerId, payerName, score}. */
  function suggestPayers(name, soharPayers, n) {
    const STOP = new Set(['of', 'the', 'and', 'health', 'plan', 'plans', 'inc', 'ca', 'northern', 'southern', 'california', 'medicare', 'advantage', 'generic']);
    const toks = (s) => str(s).toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter((t) => t && !STOP.has(t));
    const q = toks(name);
    if (!q.length) return [];
    const out = [];
    for (const p of soharPayers || []) {
      const hay = new Set(toks(`${p.payerName} ${p.aliasPayerNames || ''}`));
      let score = 0;
      for (const t of q) if (hay.has(t)) score += 1;
      // bonus for first token match and for shorter names (more specific)
      if (score && hay.has(q[0])) score += 0.5;
      if (score) score += Math.max(0, 1 - toks(p.payerName).length / 8);
      if (score) out.push({ payerId: p.payerId, payerName: p.payerName, score });
    }
    out.sort((a, b) => b.score - a.score || a.payerName.localeCompare(b.payerName));
    return out.slice(0, n || 5);
  }

  // ---------- automatic payer matching ----------
  // Normalization (used for both exact and fuzzy): lowercase, "&" -> "and", punctuation -> space,
  // a few joined phrases (united healthcare -> unitedhealthcare, medi cal -> medcal, tri west -> triwest),
  // abbreviations expanded (ca/calif -> california, bs -> blue shield, bc -> blue cross, uhc -> unitedhealthcare,
  // va -> veterans affairs, nor -> northern, adv -> advantage), filler words dropped (AUTO_STOP), duplicates removed.
  const AUTO_STOP = new Set(['the', 'of', 'and', 'inc', 'co', 'corp', 'llc', 'plan', 'plans', 'insurance', 'ins', 'generic']);
  const AUTO_ABBR = { ca: ['california'], calif: ['california'], bs: ['blue', 'shield'], bc: ['blue', 'cross'], bcbs: ['blue', 'cross', 'shield'],
    uhc: ['unitedhealthcare'], va: ['veterans', 'affairs'], veteran: ['veterans'], nor: ['northern'], adv: ['advantage'] };
  const AUTO_PHRASES = [[/\bunited health ?care\b/g, 'unitedhealthcare'], [/\bmedi cal\b/g, 'medcal'], [/\btri west\b/g, 'triwest']];
  // Generic words count half: they rarely tell two payers apart.
  const AUTO_GENERIC = new Set(['health', 'healthcare', 'care', 'california', 'northern', 'southern', 'west', 'east', 'behavioral', 'services',
    'solutions', 'network', 'foundation', 'group', 'medical', 'benefits', 'administrators', 'options', 'systems', 'national', 'alliance', 'partners', 'company']);
  const AUTO_MEDIUM = { minScore: 0.75, minCoverage: 0.66, minGap: 0.15 };
  function autoTokens(s) {
    let t = ` ${str(s).toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ')} `;
    for (const [re, to] of AUTO_PHRASES) t = t.replace(re, to);
    const out = [];
    for (const w of t.split(/\s+/)) {
      if (!w || AUTO_STOP.has(w)) continue;
      for (const x of AUTO_ABBR[w] || [w]) if (!out.includes(x)) out.push(x);
    }
    return out;
  }
  const autoWeight = (t) => (AUTO_GENERIC.has(t) ? 0.5 : 1);
  const autoPrepCache = new WeakMap();
  function autoPrep(soharPayers) {
    const list = soharPayers || [];
    if (autoPrepCache.has(list)) return autoPrepCache.get(list);
    const prepared = [];
    for (const p of list) {
      if (!str(p.payerId) && !str(p.payerName)) continue;
      const variants = [{ text: str(p.payerName), kind: 'name' }];
      for (const a of str(p.aliasPayerNames).split(/\s*[;|\n,]\s*/)) if (a) variants.push({ text: a, kind: 'alias' });
      prepared.push({ payerId: str(p.payerId), payerName: str(p.payerName), variants: variants.map((v) => ({ ...v, toks: autoTokens(v.text) })).filter((v) => v.toks.length) });
    }
    if (soharPayers) autoPrepCache.set(list, prepared);
    return prepared;
  }

  /**
   * Best Sohar payer for a NextGen payer name, with a confidence:
   *  Auto-High   exact match after normalization to a Sohar payerName or one of its aliasPayerNames (one payerId only).
   *  Auto-Medium fuzzy: score >= 0.75, the query's words matched >= 66% (by weight), and the best payer beats the
   *              best payer with a different payerId by >= 0.15, and no unmatched non-generic query word appears in
   *              another payer's name or alias (a second brand, e.g. "Western Health Advantage Optum").
   *  Auto-Low    anything weaker; payerId is still the best guess ('' when no Sohar payer shares a distinctive word).
   * Fuzzy score per name/alias = 0.7 * coverage (share of the query's word weight found) + 0.3 * precision
   * (share of the Sohar name's word weight found); generic words weigh 0.5, others 1. A candidate scores 0 unless
   * at least one non-generic word matches (when the query has one). Returns {payerId, payerName, confidence, score, evidence}.
   */
  function autoMatchPayer(name, soharPayers) {
    const q = autoTokens(name);
    const none = (evidence) => ({ payerId: '', payerName: '', confidence: 'Auto-Low', score: 0, evidence });
    if (!q.length) return none('No words to match');
    const prepared = autoPrep(soharPayers);
    if (!prepared.length) return none('Sohar payer list is empty');
    const key = q.join(' ');
    const exact = [];
    for (const p of prepared) {
      const v = p.variants.find((x) => x.toks.join(' ') === key);
      if (v && !exact.some((e) => e.p.payerId === p.payerId)) exact.push({ p, v });
    }
    if (exact.length === 1) {
      const { p, v } = exact[0];
      return { payerId: p.payerId, payerName: p.payerName, confidence: 'Auto-High', score: 1, evidence: `Exact ${v.kind} match: "${v.text}"` };
    }
    if (exact.length > 1) {
      const { p } = exact[0];
      return { payerId: p.payerId, payerName: p.payerName, confidence: 'Auto-Low', score: 1,
        evidence: `Same name on ${exact.length} Sohar payers: ${exact.map((e) => `${e.p.payerName} (${e.p.payerId})`).join(', ')}` };
    }
    const qW = q.reduce((s, t) => s + autoWeight(t), 0);
    const qHasStrong = q.some((t) => autoWeight(t) === 1);
    const scored = [];
    for (const p of prepared) {
      let best = null;
      for (const v of p.variants) {
        const hit = q.filter((t) => v.toks.includes(t));
        if (!hit.length || (qHasStrong && !hit.some((t) => autoWeight(t) === 1))) continue;
        const hW = hit.reduce((s, t) => s + autoWeight(t), 0);
        const cW = v.toks.reduce((s, t) => s + autoWeight(t), 0);
        const coverage = hW / qW, precision = hW / cW;
        const score = 0.7 * coverage + 0.3 * precision;
        if (!best || score > best.score) best = { score, coverage, hit, v };
      }
      if (best) scored.push({ p, ...best });
    }
    if (!scored.length) return none('No Sohar payer shares a distinctive word');
    scored.sort((a, b) => b.score - a.score || a.p.payerName.localeCompare(b.p.payerName));
    const top = scored[0];
    const runner = scored.find((s) => s.p.payerId !== top.p.payerId);
    const gap = top.score - (runner ? runner.score : 0);
    const r2 = (n) => Math.round(n * 100) / 100;
    // A distinctive word the best payer did not match but another payer's name has (e.g. "Optum" in
    // "Western Health Advantage Optum") means the name may be a second brand: never Auto-Medium.
    const otherBrand = q.filter((t) => autoWeight(t) === 1 && !top.hit.includes(t))
      .find((t) => prepared.some((p) => p.payerId !== top.p.payerId && p.variants.some((v) => v.toks.includes(t))));
    const medium = !otherBrand && top.score >= AUTO_MEDIUM.minScore && top.coverage >= AUTO_MEDIUM.minCoverage && gap >= AUTO_MEDIUM.minGap;
    const evidence = `Matched ${top.hit.length}/${q.length} words (${top.hit.join(', ')}) in ${top.v.kind} "${top.v.text}"; score ${r2(top.score)}`
      + (runner ? `; next best ${runner.p.payerName} ${r2(runner.score)}` : '; no other candidate')
      + (otherBrand ? `; unmatched "${otherBrand}" is in another payer's name` : '');
    return { payerId: top.p.payerId, payerName: top.p.payerName, confidence: medium ? 'Auto-Medium' : 'Auto-Low', score: r2(top.score), evidence };
  }

  /**
   * Fin Class -> Sohar payer, learned from this report: for each Fin Class (Cht Guar), the NextGen payer most
   * patients with that Fin Class have, among payers confidently mapped (Exact, Probable, Confirmed, Auto-High).
   * entries: patients (uses x.finClass and 'Primary Payer') or {finClass, payer} objects, one per patient.
   * payerMap: Payer Map rows or a payerMapIndex Map. opts.overrides: {finClass: payerId} or
   * [{finClass, payerId, payerName}] (Config); overrides win.
   * Returns Map(normalized fin class -> {finClass, nextgenPayer, payerId, payerName, count, total, source}).
   */
  function finClassPayerIndex(entries, payerMap, opts) {
    const idx = payerMapIndex(payerMap);
    const fk = (s) => str(s).toLowerCase().replace(/\s+/g, ' ');
    const tally = new Map();
    for (const e of entries || []) {
      const fc = str(e.x ? e.x.finClass : e.finClass);
      const payer = str(e.x ? e['Primary Payer'] : e.payer);
      if (!fc || !payer) continue;
      if (!tally.has(fk(fc))) tally.set(fk(fc), { finClass: fc, total: 0, payers: new Map() });
      const t = tally.get(fk(fc));
      t.total++;
      t.payers.set(payer, (t.payers.get(payer) || 0) + 1);
    }
    const out = new Map();
    for (const [k, t] of tally) {
      const ranked = [...t.payers].filter(([n]) => { const h = idx.get(n.toLowerCase()); return h && h.payerId && CONFIDENT.has(h.confidence.toLowerCase()); })
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
      if (!ranked.length) continue;
      const [nextgenPayer, count] = ranked[0];
      const h = idx.get(nextgenPayer.toLowerCase());
      out.set(k, { finClass: t.finClass, nextgenPayer, payerId: h.payerId, payerName: h.payerName, count, total: t.total, source: 'report' });
    }
    const ov = (opts && opts.overrides) || null;
    const ovList = Array.isArray(ov) ? ov : ov ? Object.entries(ov).map(([finClass, payerId]) => ({ finClass, payerId })) : [];
    for (const o of ovList) {
      if (!str(o.finClass) || !str(o.payerId)) continue;
      out.set(fk(o.finClass), { finClass: str(o.finClass), nextgenPayer: str(o.nextgenPayer), payerId: str(o.payerId), payerName: str(o.payerName), count: 0, total: 0, source: 'config' });
    }
    return out;
  }
  const finClassLookup = (fidx, fc) => (str(fc) ? fidx.get(str(fc).toLowerCase().replace(/\s+/g, ' ')) || null : null);

  // ---------- build ----------
  const NOT_KEPT_FUTURE = new Set(['cancelled', 'canceled', 'no show', 'noshow', 'deleted']);

  /**
   * rows: 2D array from the Import sheet (including title rows). cfg:
   *  { asOf: Date, scope: 'upcoming'|'all', locationState: {locName: 'CA'}, defaultState, teleKeywords: ['tele'],
   *    posTele: '10', posOffice: '11', specialty: 'PSY', payerMap: [ {nextgenPayerName,payerId,payerName,confidence,notes} ],
   *    useGuarantorAsSubscriber: bool, soharPayers: [{payerId, payerName, aliasPayerNames}] (optional; turns on auto-matching),
   *    finClassPayers: {finClass: payerId} | [{finClass, payerId, payerName}] (optional Config overrides) }
   * Each patient has the PATIENT_COLS keys plus a non-column property x (see README / API contract).
   * autoMapped: Payer Map rows to add for NextGen payer names that were not in the map.
   */
  function buildPatients(rows, cfg) {
    const det = detectHeader(rows);
    if (!det) return { error: 'Could not find the header row. The Import sheet needs the NextGen columns, including "Per Nbr" and "Appt Dt".' };
    if (det.missing.length) return { error: `Header row found but missing: ${det.missing.join(', ')}.`, detection: det };
    const cm = det.colMap;
    const asOf = cfg.asOf ? dayOf(cfg.asOf) : utcDay(new Date());
    const tele = (cfg.teleKeywords && cfg.teleKeywords.length ? cfg.teleKeywords : ['tele']).map((k) => k.toLowerCase());
    const locState = {};
    for (const [k, v] of Object.entries(cfg.locationState || {})) locState[str(k).toLowerCase()] = str(v).toUpperCase();

    const get = (row, f) => (cm[f] == null ? '' : row[cm[f]]);
    const byPatient = new Map();
    let dataRows = 0;
    const end = dataEndOf(det, rows);
    for (let r = det.headerRow + 1; r < end; r++) {
      const row = rows[r] || [];
      const per = str(get(row, 'perNbr'));
      const dt = parseDate(get(row, 'apptDt'));
      if (!per || !dt) continue; // blank rows
      dataRows++;
      const rec = {
        dt, tmRaw: get(row, 'begTm'), status: str(get(row, 'sts')), event: str(get(row, 'event')), loc: str(get(row, 'loc')),
        rendering: str(get(row, 'rendering')),
        payer: str(get(row, 'payer')), secPayer: str(get(row, 'secPayer')), secPayerLst: str(get(row, 'secPayerLstEnc')),
        pol: str(get(row, 'polNbr')), polLst: str(get(row, 'polNbrLstEnc')),
        finClass: str(get(row, 'finClass')), finClassLst: str(get(row, 'finClassLstEnc')),
        authCd: str(get(row, 'authCd')), authCdLst: str(get(row, 'authCdLstEnc')),
        name: str(get(row, 'patName')), dob: parseDate(get(row, 'dob')), zip: str(get(row, 'zip')),
        state: str(get(row, 'state')), patSts: str(get(row, 'patSts')), mrn: str(get(row, 'mrn')), guar: str(get(row, 'guarName')),
      };
      if (!byPatient.has(per)) byPatient.set(per, []);
      byPatient.get(per).push(rec);
    }
    if (!dataRows) return { error: 'Header row found but no appointment rows under it.', detection: det };

    // Per patient: latest row first; chart-level fields come from the latest row that has them.
    const pts = [];
    for (const [per, recs] of byPatient) {
      recs.sort((a, b) => b.dt - a.dt);
      const latest = recs[0];
      const pick = (f) => { const h = recs.find((x) => x[f]); return h ? h[f] : ''; };
      const of = (f) => latest[f] || pick(f);
      pts.push({ per, recs, latest, pick, payerName: of('payer'), secName: of('secPayer'), finClass: of('finClass'), finClassLst: of('finClassLst') });
    }

    // R2: auto-match NextGen payer names that are not in the Payer Map. Rows already in the map win.
    const autoMapped = [];
    const mapRows = (cfg.payerMap || []).slice();
    const known = payerMapIndex(mapRows);
    if (cfg.soharPayers && cfg.soharPayers.length) {
      const names = new Map();
      for (const t of pts) for (const n of [t.payerName, t.secName]) if (n) names.set(n.toLowerCase(), { name: n, n: (names.get(n.toLowerCase()) || { n: 0 }).n + 1 });
      for (const { name, n } of names.values()) {
        if (known.has(name.toLowerCase()) || isSelfPay(name) || isVip(name)) continue;
        const m = autoMatchPayer(name, cfg.soharPayers);
        const entry = { nextgenPayerName: name, payerId: m.payerId, payerName: m.payerName, confidence: m.confidence,
          notes: `Auto-matched by the add-in: ${m.evidence}`, score: m.score, patients: n };
        autoMapped.push(entry);
        mapRows.push(entry);
      }
      autoMapped.sort((a, b) => b.patients - a.patients || a.nextgenPayerName.localeCompare(b.nextgenPayerName));
    }
    const idx = payerMapIndex(mapRows);
    const finIdx = finClassPayerIndex(pts.map((t) => ({ finClass: t.finClass, payer: t.payerName })), idx, { overrides: cfg.finClassPayers });

    const patients = [];
    const review = [];
    const addReview = (sev, issue, per, detail, fix) => review.push({ Severity: sev, Issue: issue, 'Per Nbr': per, Detail: detail, Fix: fix });

    for (const { per, recs, latest, pick, payerName, secName, finClass, finClassLst } of pts) {
      const future = recs.filter((x) => x.dt >= asOf && !NOT_KEPT_FUTURE.has(x.status.toLowerCase()))
        .sort((a, b) => a.dt - b.dt || timeMins(a.tmRaw) - timeMins(b.tmRaw));
      const next = future[0] || null;
      const posEvent = (next || latest).event;
      const pos = tele.some((k) => posEvent.toLowerCase().includes(k)) ? (cfg.posTele || '10') : (cfg.posOffice || '11');

      const nm = parseName(pick('name'));
      const dob = recs.map((x) => x.dob).find(Boolean) || null;
      const zip = str(pick('zip')).replace(/[^0-9-]/g, '').slice(0, 10);
      const state = (str(pick('state')) || locState[latest.loc.toLowerCase()] || cfg.defaultState || '').toUpperCase().slice(0, 2);

      const memberId = latest.pol || pick('pol');
      const memberLst = latest.polLst || pick('polLst');
      const lk = lookupPayer(payerName, idx);
      const secLk = lookupPayer(secName, idx);
      const conf = str(lk.confidence).toLowerCase();
      const autoLow = conf === 'auto-low';
      const vip = isVip(payerName) || isVip(finClass);

      const distinctPayers = new Set(recs.map((x) => x.payer).filter(Boolean));
      const flags = [];
      const verifyRows = [];
      let route, reason, soharId = lk.payerId || '';

      // Last-encounter insurance, derived from Fin Class (Lst Enc) via this report's Fin Class -> payer index.
      const lastEncRow = () => {
        if (!finClassLst || isVip(finClassLst) || isSelfPay(finClassLst) || !memberLst) return null;
        const f = finClassLookup(finIdx, finClassLst);
        if (!f) return null;
        return { payerId: f.payerId, payerName: f.payerName, memberId: memberLst, source: 'last encounter', nextgenPayer: f.nextgenPayer,
          confidence: f.source === 'config' ? 'Fin Class (config)' : 'Fin Class (report)' };
      };

      if (vip) {
        // R4: VIP / partial VIP = MHS-sponsored. Verify other insurance on file, else discovery only.
        // A real chart payer (Fin Class VIP / Partial VIP) is that other insurance.
        flags.push('VIP');
        const chartOk = !isVip(payerName) && !isSelfPay(payerName) && lk.status === 'mapped' && memberId;
        const le = chartOk ? null : lastEncRow();
        const secOk = secName && !isVip(secName) && !isSelfPay(secName) && secLk.status === 'mapped';
        if (chartOk) verifyRows.push({ payerId: lk.payerId, payerName: lk.payerName || '', memberId, source: 'guarantor', nextgenPayer: payerName, confidence: lk.confidence || '' });
        else if (le) verifyRows.push(le);
        if (verifyRows.length) {
          route = 'verification'; reason = `VIP (MHS-sponsored); verifying other insurance: ${verifyRows[0].source} ${verifyRows[0].nextgenPayer || verifyRows[0].payerName}`;
        } else if (secOk) {
          // The report has no secondary policy number (Pol Nbr (Cht Guar) is the VIP entry's), so discovery with the secondary's payer as the hint.
          route = 'discovery'; reason = `VIP (MHS-sponsored); secondary ${secName} has no policy number in the report: discovery with its payer as the hint`;
          soharId = secLk.payerId;
        } else {
          route = 'discovery'; reason = 'VIP (MHS-sponsored), no other insurance on file: discovery only';
          soharId = '';
        }
      } else if (lk.status === 'mapped' && memberId) {
        route = 'verification'; reason = /^auto/.test(conf) ? `auto-mapped payer (${lk.confidence}) + member ID` : 'mapped payer + member ID';
        verifyRows.push({ payerId: lk.payerId, payerName: lk.payerName || '', memberId, source: 'guarantor', nextgenPayer: payerName, confidence: lk.confidence || '' });
      } else if (lk.status === 'none') { route = 'discovery'; reason = 'no payer on file'; }
      else if (lk.status === 'review' && conf === 'not a payer') {
        route = 'discovery'; reason = `payer needs review (${lk.confidence}): ${payerName}`;
        if (!isSelfPay(payerName)) { route = 'exclude'; reason = `not a payer: ${payerName}`; }
      } else if (lk.status === 'unmapped' || lk.status === 'review') {
        // R3: low confidence + a different last encounter -> send both insurances; the results decide.
        const differs = (finClassLst && finClassLst.toLowerCase() !== finClass.toLowerCase()) || (memberLst && idKey(memberLst) !== idKey(memberId));
        const le = differs ? lastEncRow() : null;
        if (autoLow) soharId = '';
        if (le) {
          if (lk.payerId && memberId) verifyRows.push({ payerId: lk.payerId, payerName: lk.payerName || '', memberId, source: 'guarantor', nextgenPayer: payerName, confidence: lk.confidence || '' });
          verifyRows.push(le);
          route = 'verification';
          reason = verifyRows.length > 1 ? `two insurances: chart ${payerName} (${lk.confidence}) and last encounter ${finClassLst}` : `chart payer ${lk.confidence}: ${payerName}; verifying last-encounter insurance (${finClassLst})`;
          if (verifyRows.length > 1) flags.push('Two insurances sent');
          addReview('Low', verifyRows.length > 1 ? 'Two insurances sent to verification' : 'Last-encounter insurance sent to verification', per,
            `Chart: ${payerName}${lk.payerName ? ` (best guess ${lk.payerName}, ${lk.confidence})` : ` (${lk.confidence})`}; last encounter: ${finClassLst} -> ${le.payerName}`,
            'The Sohar results decide; Choose in NextGen shows which insurance to keep.');
          if (autoLow) addReview('High', 'Payer needs manual classification', per, `${payerName}${lk.payerName ? ` -> best guess ${lk.payerName}` : ''}`, 'Classify it in the pane (Admin, step 1) or on the Payer Map sheet.');
        } else if (lk.status === 'unmapped') {
          route = 'discovery'; reason = `payer not in Payer Map: ${payerName}`;
          addReview('High', 'Unmapped payer', per, payerName, 'Map it in step 2 or add a row on the Payer Map sheet.');
        } else {
          route = 'discovery'; reason = `payer needs review (${lk.confidence}): ${payerName}`;
          if (autoLow) addReview('High', 'Payer needs manual classification', per, `${payerName}${lk.payerName ? ` -> best guess ${lk.payerName}` : ''}`, 'Sent to discovery. Classify it in the pane (Admin, step 1) or on the Payer Map sheet so it verifies next time.');
        }
      } else { route = 'discovery'; reason = 'mapped payer but no member ID'; flags.push('No member ID'); }
      if (route === 'verification') soharId = verifyRows[0].payerId;

      if (!nm.firstName || !nm.lastName) { flags.push('Name incomplete'); addReview('High', 'Name missing or incomplete', per, pick('name') || '(blank)', 'Both Sohar imports need firstName and lastName. Re-run the report with Pat Name, or fill in on the Patients sheet.'); }
      else if (nm.flag) { flags.push(nm.flag); addReview('Low', 'Name split uncertain', per, pick('name'), 'Check first/last on the Patients sheet.'); }
      if (!dob) { flags.push('DOB missing'); addReview('High', 'DOB missing', per, '', 'Both Sohar imports need dateOfBirth (MM/DD/YYYY).'); }
      if (route === 'discovery' && !zip) { flags.push('Zip missing'); addReview('Medium', 'Zip missing for discovery', per, '', 'Discovery works best with home ZIP. Add a ZIP column to the report or fill in on Patients.'); }
      if (!state) { flags.push('State missing'); addReview('Medium', 'State missing', per, latest.loc, 'Add the location to Config > Location to State.'); }
      if (memberId && memberLst && memberId !== memberLst) { flags.push('Member ID differs from last encounter'); addReview('Low', 'Member ID mismatch', per, `Cht Guar vs Lst Enc differ`, 'Confirm which policy is current before upload.'); }
      if (distinctPayers.size > 1) { flags.push('Payer changed across rows'); addReview('Low', 'Payer varies across appointments', per, [...distinctPayers].join(' | '), 'Latest appointment payer was used.'); }
      if (secName && secLk.status === 'unmapped') addReview('Low', 'Secondary payer unmapped', per, secName, 'Optional: map it so the Secondary Sohar payerId fills in.');

      let subF = '', subL = '', subD = '';
      if (cfg.useGuarantorAsSubscriber && latest.guar && str(latest.guar).toLowerCase() !== str(pick('name')).toLowerCase()) {
        const g = parseName(latest.guar); subF = g.firstName; subL = g.lastName;
      }

      const inScope = cfg.scope === 'all' || !!next;
      patients.push({
        'Per Nbr': per, 'MRN': pick('mrn'), 'Last Name': nm.lastName, 'First Name': nm.firstName, 'DOB': fmtDate(dob),
        'State': state, 'Zip': zip, 'Pat Sts': latest.patSts, 'Route': route, 'Route Reason': reason, 'In Scope': inScope ? 'Yes' : 'No',
        'Next Appt': fmtDate(next && next.dt), 'Next Event': next ? next.event : '', 'POS': pos,
        'Primary Payer': payerName, 'Sohar payerId': soharId, 'Confidence': lk.confidence || '',
        'Member ID': memberId, 'Member ID (Lst Enc)': memberLst, 'Member ID Mismatch': (memberId && memberLst && memberId !== memberLst) ? 'Yes' : 'No',
        'Secondary Payer': secName, 'Secondary Sohar payerId': str(secLk.confidence).toLowerCase() === 'auto-low' ? '' : secLk.payerId || '',
        'Appt Count': recs.length, 'First Appt': fmtDate(recs[recs.length - 1].dt), 'Last Appt': fmtDate(latest.dt),
        'Subscriber First': subF, 'Subscriber Last': subL, 'Subscriber DOB': subD, 'Flags': flags.join('; '),
        x: {
          finClass, finClassLstEnc: finClassLst, secPayer: secName, secPayerLstEnc: latest.secPayerLst || pick('secPayerLst'),
          authCd: latest.authCd || pick('authCd'), authCdLstEnc: latest.authCdLst || pick('authCdLst'),
          guarZip: zip, location: next ? next.loc : latest.loc, rendering: next ? next.rendering : '',
          vip, payerConfidence: lk.confidence || '', payerGuess: lk.payerId ? { payerId: lk.payerId, payerName: lk.payerName || '' } : null,
          appointments: future.map((a) => ({ date: fmtDate(a.dt), time: fmtTime(a.tmRaw), event: a.event, loc: a.loc, rendering: a.rendering, status: a.status, authCd: a.authCd, authCdLstEnc: a.authCdLst })),
          verifyRows,
        },
      });
    }
    patients.sort((a, b) => (a.Route > b.Route ? -1 : a.Route < b.Route ? 1 : 0) || ((b['In Scope'] === 'Yes') - (a['In Scope'] === 'Yes')) || a['Per Nbr'].localeCompare(b['Per Nbr']));

    const unmapped = {};
    for (const p of patients) {
      if (p.Route === 'discovery' && p['Route Reason'].startsWith('payer not in Payer Map')) unmapped[p['Primary Payer']] = (unmapped[p['Primary Payer']] || 0) + 1;
    }
    return { detection: det, dataRows, patients, review, unmappedPayers: unmapped, autoMapped, finClassIndex: finIdx, asOf };
  }

  /** Sohar import rows. One verification row per verifyRows entry (same id = Per Nbr); discovery one per patient. */
  function buildExports(patients, cfg) {
    const specialty = (cfg && cfg.specialty) || 'PSY';
    const inScope = patients.filter((p) => p['In Scope'] === 'Yes');
    const verification = [];
    for (const p of inScope.filter((x) => x.Route === 'verification')) {
      const vrs = p.x && p.x.verifyRows && p.x.verifyRows.length ? p.x.verifyRows : [{ payerId: p['Sohar payerId'], memberId: p['Member ID'] }];
      for (const v of vrs) {
        verification.push({
          firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, memberId: v.memberId,
          payerId: v.payerId, placeOfServiceCode: p.POS, specialtyCode: specialty, id: p['Per Nbr'],
          subscriberFirstName: p['Subscriber First'], subscriberLastName: p['Subscriber Last'], subscriberDateOfBirth: p['Subscriber DOB'],
        });
      }
    }
    const discovery = inScope.filter((p) => p.Route === 'discovery').map((p) => ({
      firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, zipCode: p.Zip,
      payerId: p['Sohar payerId'], id: p['Per Nbr'],
    }));
    return { verification, discovery };
  }

  function csvCell(v) {
    const s = v == null ? '' : String(v);
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }
  function toCSV(cols, rows) {
    const lines = [cols.join(',')];
    for (const r of rows) lines.push(cols.map((c) => csvCell(r[c])).join(','));
    return lines.join('\r\n') + '\r\n';
  }
  /** RFC 4180 CSV -> 2D array of strings. Handles quotes, embedded newlines, CRLF and a UTF-8 BOM. */
  function parseCSV(text) {
    const s = String(text || '').replace(/^﻿/, '');
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (q) {
        if (ch === '"') { if (s[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && s[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  /** Excel time fraction, Date, or text -> "h:mm AM". */
  /** Minutes after midnight for sorting: Excel fractions, Dates, "9:00 AM", "13:30". Unreadable sorts last. */
  function timeMins(v) {
    if (v == null || v === '') return 9999;
    if (typeof v === 'number') return Math.round((v % 1) * 1440);
    if (v instanceof Date) return v.getUTCHours() * 60 + v.getUTCMinutes();
    const s = str(v);
    if (/^\d*\.\d+$/.test(s)) return Math.round((+s % 1) * 1440);
    const m = s.match(/^(\d{1,2}):(\d{2})\s*([AP]M)?/i);
    if (!m) return 9999;
    let h = +m[1] % (m[3] ? 12 : 24);
    if (m[3] && /pm/i.test(m[3])) h += 12;
    return h * 60 + +m[2];
  }

  function fmtTime(v) {
    if (v == null || v === '') return '';
    let mins;
    if (typeof v === 'number') mins = Math.round((v % 1) * 1440);
    else if (v instanceof Date) mins = v.getUTCHours() * 60 + v.getUTCMinutes();
    else return str(v);
    const h = Math.floor(mins / 60) % 24, m = mins % 60;
    return `${h % 12 || 12}:${String(m).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`;
  }

  // ---------- Sohar results ----------
  /** Identify a Sohar export by its header row: 'verification', 'discovery', or null. */
  function resultKind(headers) {
    const h = new Set((headers || []).map(normHeader));
    if (h.has('requestid') && (h.has('responsepayerid') || h.has('responsepayername'))) return 'discovery';
    if (h.has('status') && (h.has('verificationid') || h.has('requestid'))) return 'verification';
    return null;
  }

  /** 2D rows from a Sohar export -> {kind, records: [{header: value}]}. Header keys are normalized. */
  function readResultRows(rows) {
    const start = (rows || []).findIndex((r) => (r || []).some((v) => !isBlank(v)));
    if (start < 0) return { kind: null, records: [] };
    const headers = rows[start].map(normHeader);
    const records = [];
    for (let i = start + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!r.some((v) => !isBlank(v))) continue;
      const o = {};
      headers.forEach((k, j) => { if (k) o[k] = r[j]; });
      records.push(o);
    }
    return { kind: resultKind(rows[start]), records };
  }

  // Member IDs are compared without punctuation, case or leading zeros (Excel drops zeros from numeric IDs).
  const idKey = (v) => str(v).toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+/, '');
  const nameKey = (first, last) => `${str(last).toLowerCase()}|${str(first).toLowerCase()}`;

  // ---------- benefits ----------
  const money = (v) => { if (isBlank(v)) return null; const n = parseFloat(String(v).replace(/[$,%\s]/g, '')); return isNaN(n) ? null : n; };
  const fmtMoney = (n) => (n == null ? '' : String(Math.round(n * 100) / 100));
  const fmtPct = (v) => { const n = money(v); if (n == null) return ''; return `${Math.round((n <= 1 ? n * 100 : n) * 10) / 10}%`; };

  /**
   * Accumulator status: 'Met', 'Approaching' (remaining at or under pct of total), 'Not met', 'None' (plan has $0), or ''.
   */
  function accumulatorStatus(total, remaining, metFlag, pct) {
    if (/^(true|yes|y)$/i.test(str(metFlag))) return 'Met';
    const t = money(total), r = money(remaining);
    if (t === 0) return 'None';
    if (r != null) {
      if (r <= 0) return 'Met';
      if (t != null && t > 0 && r / t <= pct) return 'Approaching';
      return 'Not met';
    }
    if (/^(false|no|n)$/i.test(str(metFlag))) return 'Not met';
    return '';
  }

  /**
   * Benefits from a Sohar record (normalized keys). Uses Sohar's in-network tier-one columns
   * (innTotalDeductible, innRemainingDeductible, innOutOfPocket, innOutOfPocketRemaining, innCopay, innCoinsurance).
   * Other exports fall back to a loose match that skips family, out-of-network and tier 2/3 columns.
   * opts.approachingPct: fraction, default 0.2.
   */
  function benefitFields(rec, opts) {
    const pct = opts && opts.approachingPct != null ? opts.approachingPct : 0.2;
    const keys = Object.keys(rec).filter((k) => !/family|outofnetwork|^oon|oopoon|tiertwo|tierthree|coveragelevel/.test(k));
    const exact = (k) => (k in rec && !isBlank(rec[k]) ? rec[k] : '');
    const loose = (test) => { const k = keys.find((x) => test(x) && !isBlank(rec[x])); return k ? rec[k] : ''; };
    const pick = (k, test) => exact(k) || loose(test);
    const oop = (k) => k.includes('outofpocket') || k.startsWith('oop');
    const rem = (k) => k.includes('remain');
    const met = (k) => k.endsWith('met') || k.includes('metamount');
    const ded = {
      total: pick('inntotaldeductible', (k) => k.includes('deductible') && !rem(k) && !met(k)),
      remaining: pick('innremainingdeductible', (k) => k.includes('deductible') && rem(k)),
      met: loose((k) => k.includes('deductible') && met(k)),
    };
    const oo = {
      total: pick('innoutofpocket', (k) => oop(k) && !rem(k) && !met(k)),
      remaining: pick('innoutofpocketremaining', (k) => oop(k) && rem(k)),
      met: loose((k) => oop(k) && met(k)),
    };
    const mt = (b) => (money(b.met) != null && money(b.total) != null && money(b.remaining) == null ? fmtMoney(money(b.total) - money(b.met)) : b.remaining);
    const dedRem = mt(ded), oopRem = mt(oo);
    return {
      'Copay': str(pick('inncopay', (k) => k.includes('copay'))),
      'Coinsurance': fmtPct(pick('inncoinsurance', (k) => k.includes('coinsurance'))),
      'Deductible': str(ded.total), 'Deductible Remaining': str(dedRem),
      'Deductible Status': accumulatorStatus(ded.total, dedRem, money(ded.met) == null ? ded.met : '', pct),
      'OOP Max': str(oo.total), 'OOP Remaining': str(oopRem),
      'OOP Status': accumulatorStatus(oo.total, oopRem, money(oo.met) == null ? oo.met : '', pct),
      'Coverage Level': str(exact('inncoveragelevelcode') || exact('oopinncoveragelevelcode')),
      'Plan': str(exact('planname')), 'Network': str(exact('networktype')),
      'Coverage Start': str(exact('coveragestartdate')), 'Coverage End': str(exact('coverageenddate')),
    };
  }

  /** One-line benefit note for prep: "OOP max met", "Deductible approaching", "No deductible". */
  function benefitFlag(b) {
    const parts = [];
    if (b['OOP Status'] === 'Met') parts.push('OOP max met');
    else if (b['Deductible Status'] === 'Met') parts.push('Deductible met');
    if (b['OOP Status'] === 'Approaching') parts.push('OOP max approaching');
    if (b['Deductible Status'] === 'Approaching' && b['OOP Status'] !== 'Met') parts.push('Deductible approaching');
    if (!parts.length && b['Deductible Status'] === 'None') parts.push('No deductible');
    return parts.join('; ');
  }

  /** Bucket a result row for the admin worklist. */
  /** The verifyRows entry a Sohar verification record answers: same member ID, else same Sohar payer name. */
  function matchVerifyRow(p, rec) {
    const vrs = (p && p.x && p.x.verifyRows) || [];
    if (!vrs.length) return null;
    const mid = idKey(rec.memberid);
    const byId = mid ? vrs.filter((v) => idKey(v.memberId) === mid) : [];
    if (byId.length === 1) return byId[0];
    const pn = str(rec.payername).toLowerCase();
    const pool = byId.length ? byId : vrs;
    const byName = pn ? pool.filter((v) => str(v.payerName).toLowerCase() === pn) : [];
    if (byName.length) return byName[0];
    return byId[0] || (vrs.length === 1 ? vrs[0] : null);
  }

  function workQueue(row) {
    if (row.Match === 'No match') return 'Unmatched';
    if (/^Pending/.test(row.Action)) return 'Pending';
    if (/^Eligible\. No action|^Coverage found matches|^Superseded/.test(row.Action)) return 'Clear';
    return 'Action needed';
  }

  /** LASFIR: first three letters of last name + first three of first name. */
  function lasfir(last, first) {
    const l = (s) => str(s).toUpperCase().replace(/[^A-Z]/g, '').slice(0, 3);
    return l(last) + l(first);
  }

  /**
   * Plain-text note for admins, grouped by location then appointment. LASFIR + Per Nbr only (minimum necessary).
   * opts.mode: 'action' (default: rows needing action) or 'benefits' (deductible/OOP met or approaching).
   */
  function adminMessage(rows, opts) {
    const o = opts || {};
    const benefits = o.mode === 'benefits';
    const todo = benefits
      ? rows.filter((r) => /met|approaching/i.test(r['Benefit Flag'] || '') && r['Next Appt'])
      : rows.filter((r) => r['Work Queue'] === 'Action needed' || r['Work Queue'] === 'Unmatched');
    if (!todo.length) return benefits ? 'Sohar benefits: no upcoming patients have a deductible or OOP max met or approaching.\n' : 'Sohar eligibility: nothing needs action in this list.\n';
    const byLoc = new Map();
    for (const r of todo) {
      const loc = r['Next Location'] || 'No upcoming appointment';
      if (!byLoc.has(loc)) byLoc.set(loc, []);
      byLoc.get(loc).push(r);
    }
    const n = todo.length, s = n > 1 ? 's' : '';
    const from = o.asOf ? ` (appointments from ${fmtDate(o.asOf)})` : '';
    const head = benefits
      ? `Sohar benefits${from}: ${n} upcoming patient${s} with a deductible or out-of-pocket max met or approaching. Check patient responsibility before the visit.`
      : `Sohar eligibility follow-up${from}: ${n} patient${s} need${n > 1 ? '' : 's'} action before their visit.`;
    const lines = [head, ''];
    for (const [loc, list] of [...byLoc].sort((a, b) => a[0].localeCompare(b[0]))) {
      lines.push(`${loc} (${list.length})`);
      for (const r of list) {
        const when = r['Next Appt'] ? `${String(r['Next Appt']).slice(0, 5)}${r['Next Time'] ? ' ' + r['Next Time'] : ''}` : 'no appt';
        const tag = lasfir(r['Last Name'], r['First Name']);
        const id = r['Per Nbr'] ? `${tag} · Per ${r['Per Nbr']}` : `${tag ? tag + ' · ' : ''}Sohar ref ${r['Sohar Ref'] || '?'}`;
        if (benefits) {
          const amt = [r['Deductible Status'] && r['Deductible Status'] !== 'None' ? `ded remaining ${r['Deductible Remaining'] || '?'}` : '', r['OOP Remaining'] !== '' && r['OOP Remaining'] != null ? `OOP remaining ${r['OOP Remaining']}` : ''].filter(Boolean).join(', ');
          lines.push(`- ${when} · ${id} · ${r['Benefit Flag']}${amt ? ` (${amt})` : ''}`);
        } else {
          lines.push(`- ${when} · ${id} · ${r.Action}${r['Benefit Flag'] ? ` (${r['Benefit Flag'].toLowerCase()})` : ''}`);
        }
      }
      lines.push('');
    }
    return lines.join('\n').trim() + '\n';
  }

  /** Sohar resolved/created timestamp in ms (full time when present), NaN when blank. */
  function soharTime(v) {
    // Excel serial (from an .xlsx output), time of day kept: Date.parse would read 46309.04 as a year.
    if (typeof v === 'number' || /^\d+(\.\d+)?$/.test(str(v))) { const n = +v; return n > 59 ? Math.round((n - 25569) * 86400000) : NaN; }
    const t = Date.parse(str(v));
    if (!isNaN(t)) return t;
    const d = parseDate(v);
    return d ? d.getTime() : NaN;
  }

  function verificationAction(status) {
    const s = str(status).toLowerCase();
    if (s === 'complete.eligible') return 'Eligible. No action.';
    if (s === 'complete.ineligible') return 'Ineligible. Confirm coverage with the patient before the visit; consider discovery.';
    if (s === 'error.member.id') return 'Member ID rejected. Check the card, fix in NextGen, re-verify.';
    if (s === 'error.member.name') return 'Name rejected. Match the name to the card, re-verify.';
    if (s.startsWith('error')) return 'Sohar error. Review the request and re-verify.';
    if (s === 'pending' || !s) return 'Pending at Sohar. Recheck later.';
    return 'Review.';
  }

  /**
   * Join Sohar result files to the patients' upcoming appointments.
   * importRows: Import sheet 2D array. files: [{name, rows}] (2D, header first).
   * cfg: as for buildPatients, plus approachingPct (fraction, default 0.2).
   * Returns {results, unknownFiles, counts, asOf} or {error}.
   */
  function mergeSoharResults(importRows, files, cfg) {
    const built = buildPatients(importRows, { ...cfg, scope: 'all' });
    if (built.error) return { error: `The Import sheet has to hold the appointment report to combine results. ${built.error}` };
    const asOf = built.asOf;
    const det = built.detection, cm = det.colMap;
    const get = (row, f) => (cm[f] == null ? '' : row[cm[f]]);
    const bopts = { approachingPct: cfg.approachingPct != null ? cfg.approachingPct : 0.2 };
    // Sohar payerId -> NextGen payer names (from the Payer Map), so discovery can say which NextGen insurance to pick.
    const nextgenFor = new Map();
    for (const m of cfg.payerMap || []) {
      const id = str(m.payerId), name = str(m.nextgenPayerName);
      if (!id || !name || !VERIFIABLE.has(str(m.confidence).toLowerCase())) continue;
      if (!nextgenFor.has(id)) nextgenFor.set(id, []);
      if (!nextgenFor.get(id).includes(name)) nextgenFor.get(id).push(name);
    }

    // Upcoming, still-expected appointments per patient (same rule as scope "upcoming").
    const upcoming = new Map();
    for (let r = det.headerRow + 1; r < dataEndOf(det, importRows); r++) {
      const row = importRows[r] || [];
      const per = str(get(row, 'perNbr')); const dt = parseDate(get(row, 'apptDt'));
      if (!per || !dt || dt < asOf || NOT_KEPT_FUTURE.has(str(get(row, 'sts')).toLowerCase())) continue;
      if (!upcoming.has(per)) upcoming.set(per, []);
      upcoming.get(per).push({ dt, tmRaw: get(row, 'begTm'), event: str(get(row, 'event')), loc: str(get(row, 'loc')), rendering: str(get(row, 'rendering')) });
    }
    for (const list of upcoming.values()) list.sort((a, b) => a.dt - b.dt || timeMins(a.tmRaw) - timeMins(b.tmRaw));

    const byPer = new Map(), byMember = new Map(), byName = new Map(), byNameDob = new Map();
    const push = (m, k, p) => { if (!k) return; if (!m.has(k)) m.set(k, []); if (!m.get(k).includes(p)) m.get(k).push(p); };
    for (const p of built.patients) {
      byPer.set(p['Per Nbr'], p);
      push(byMember, idKey(p['Member ID']), p);
      push(byMember, idKey(p['Member ID (Lst Enc)']), p);
      if (p['Last Name']) push(byName, nameKey(p['First Name'], p['Last Name']), p);
      if (p['Last Name'] && p.DOB) push(byNameDob, `${nameKey(p['First Name'], p['Last Name'])}|${p.DOB}`, p);
    }

    const results = [], unknownFiles = [], kinds = new Set(), seen = { verification: new Set(), discovery: new Set() };
    const counts = { matched: 0, unmatched: 0, notReturned: 0, withBenefits: 0, byAction: {}, byQueue: {}, deductible: {}, oop: {} };

    const emit = (p, kind, f) => {
      const ups = p ? upcoming.get(p['Per Nbr']) || [] : [];
      const n = ups[0];
      const row = {
        'Next Appt': n ? fmtDate(n.dt) : '', 'Next Time': n ? fmtTime(n.tmRaw) : '', 'Next Event': n ? n.event : '',
        'Next Location': n ? n.loc : '', 'Next Provider': n ? n.rendering : '', 'Upcoming Appts': p ? ups.length : '',
        'Upcoming Detail': ups.slice(0, 6).map((u) => `${fmtDate(u.dt).slice(0, 5)} ${u.event}`.trim()).join('; ') + (ups.length > 6 ? `; +${ups.length - 6} more` : ''),
        'Per Nbr': p ? p['Per Nbr'] : '', 'MRN': p ? p.MRN : '', 'Last Name': p ? p['Last Name'] : str(f.lastName), 'First Name': p ? p['First Name'] : str(f.firstName),
        'DOB': p ? p.DOB : f.dob || '', 'Sent As': kind === 'verification' ? 'Verification' : 'Discovery',
        'NextGen Payer': p ? p['Primary Payer'] : '', 'NextGen Member ID': p ? p['Member ID'] : '',
        ...(f.benefits || {}), ...f.cols,
      };
      for (const c of RESULTS_COLS) if (row[c] == null) row[c] = '';
      row.x = p ? p.x || null : null; // non-column: patient extras (appointments, fin class, verifyRows) for the rules
      row['Benefit Flag'] = benefitFlag(row);
      results.push(row);
      if (p) seen[kind].add(p['Per Nbr']);
    };

    const pickOne = (list, last) => {
      if (!list || !list.length) return null;
      if (list.length === 1) return list[0];
      const l = str(last).toLowerCase();
      const hits = list.filter((p) => p['Last Name'].toLowerCase() === l);
      return hits.length === 1 ? hits[0] : null;
    };

    for (const file of files || []) {
      // A file is either raw rows from an upload or saved records ({kind, records}) from earlier uploads.
      const { kind, records } = file.records ? file : readResultRows(file.rows);
      if (!kind) { unknownFiles.push(file.name); continue; }
      kinds.add(kind);
      for (const r of records) {
        let p = null, match = '';
        const first = str(r.firstname || r.subscriberfirstname), last = str(r.lastname || r.subscriberlastname);
        const dob = parseDate(r.dateofbirth);
        const cols = {};
        let benefits = {};
        if (kind === 'verification') {
          // Newer exports echo our id (Per Nbr); older ones only carry member ID and name.
          p = byPer.get(str(r.id)) || null;
          if (p) match = 'Per Nbr';
          if (!p) { p = pickOne(byMember.get(idKey(r.memberid)), last); if (p) match = 'Member ID'; }
          if (!p && last) { p = pickOne(byName.get(nameKey(first, last)), last); if (p) match = 'Name only. Confirm.'; }
          benefits = benefitFields(r, bopts);
          let action = verificationAction(r.status);
          const end = parseDate(r.coverageenddate);
          const next = p && (upcoming.get(p['Per Nbr']) || [])[0];
          if (str(r.status).toLowerCase() === 'complete.eligible' && end && next && end < next.dt) {
            action = `Coverage ends ${fmtDate(end)}, before the next visit. Confirm new coverage with the patient.`;
          }
          Object.assign(cols, {
            'Sohar Status': str(r.status) + (str(r.statusreason) ? ` (${str(r.statusreason)})` : ''), 'Action': action,
            'Sohar Payer': str(r.payername), 'Sohar payerId': '', 'Sohar Member ID': str(r.memberid), 'Coverage Differs': '',
            'Sohar Ref': str(r.verificationid || r.requestid), 'Sohar Link': str(r.url), 'Sohar Date': str(r.resolved || r.created),
          });
          // Which insurance we sent: match on member ID, else Sohar payer name. Non-chart rows show what was sent.
          const vr = p ? matchVerifyRow(p, r) : null;
          if (vr) {
            cols['Sohar payerId'] = vr.payerId;
            if (vr.source !== 'guarantor') { cols['NextGen Payer'] = vr.nextgenPayer || ''; cols['NextGen Member ID'] = vr.memberId; }
          }
        } else {
          p = byPer.get(str(r.id)) || null;
          if (p) match = 'Per Nbr';
          else if (last) {
            p = pickOne(byNameDob.get(`${nameKey(first, last)}|${fmtDate(dob)}`), last);
            if (p) match = 'Name + DOB';
          }
          const found = !isBlank(r.responsepayerid) || !isBlank(r.responsepayername);
          const resolved = !isBlank(r.resolved);
          const status = found ? 'Coverage found' : resolved ? 'No coverage found' : 'Pending';
          const differs = found && p && (str(r.responsepayerid) !== p['Sohar payerId'] || (!isBlank(r.memberid) && idKey(r.memberid) !== idKey(p['Member ID'])));
          const getId = isBlank(r.memberid) ? ' Get the member ID from the Sohar link.' : '';
          const options = found ? nextgenFor.get(str(r.responsepayerid)) || [] : [];
          const current = p && options.some((o) => o.toLowerCase() === p['Primary Payer'].toLowerCase()) ? p['Primary Payer'] : '';
          let choose = '', pickText = '';
          if (found) {
            if (current) { choose = current; pickText = ` Keep "${current}" in NextGen.`; }
            else if (options.length === 1) { choose = options[0]; pickText = ` Choose "${options[0]}" in NextGen.`; }
            else if (options.length > 1) {
              choose = `One of: ${options.join(' | ')}`;
              pickText = ` Choose one of ${options.map((o) => `"${o}"`).join(', ')} in NextGen; match the plan type on the card.`;
            } else {
              choose = `Not in Payer Map: ${str(r.responsepayername)} (${str(r.responsepayerid)})`;
              pickText = ' No NextGen insurance is mapped to this Sohar payer yet; add it to the Payer Map.';
            }
          }
          cols['Choose in NextGen'] = choose;
          let action;
          if (!found) action = resolved ? 'No coverage found. Confirm self-pay or ask the patient for a card.' : 'Pending at Sohar. Recheck later.';
          else if (p && !p['Primary Payer']) action = `Coverage found. Add it to NextGen, then verify.${pickText}${getId}`;
          else if (differs) action = `Different coverage found. Confirm with the patient, update NextGen, then verify.${pickText}${getId}`;
          else action = 'Coverage found matches NextGen. Verify.';
          Object.assign(cols, {
            'Sohar Status': status, 'Action': action, 'Sohar Payer': str(r.responsepayername), 'Sohar payerId': str(r.responsepayerid),
            'Sohar Member ID': str(r.memberid), 'Coverage Differs': found ? (differs ? 'Yes' : 'No') : '',
            'Sohar Ref': str(r.requestid), 'Sohar Link': str(r.url), 'Sohar Date': str(r.resolved || r.created),
          });
        }
        if (!p) { match = 'No match'; cols.Action = `Not matched to a patient in the Import report. ${cols.Action}`; counts.unmatched++; }
        else counts.matched++;
        cols.Match = match;
        if (benefits['Deductible'] || benefits['OOP Max'] || benefits['Copay']) counts.withBenefits++;
        emit(p, kind, { firstName: first, lastName: last, dob: fmtDate(dob), cols, benefits });
      }
    }

    // Patients who would have been sent for an uploaded file type but came back with nothing.
    for (const p of built.patients) {
      const kind = p.Route;
      if (!kinds.has(kind) || seen[kind].has(p['Per Nbr']) || !upcoming.has(p['Per Nbr'])) continue;
      counts.notReturned++;
      emit(p, kind, { cols: { 'Sohar Status': 'Not in Sohar file', 'Action': 'Not in the results. Check it was uploaded to Sohar.', 'Match': '' } });
    }

    // Discovery only finds the policy (payer, member ID). Copay, deductible and OOP come from a verification
    // submission in Sohar, so every policy discovery finds goes back to Sohar as a verification.
    // A verification result is the follow-up when it is for the same patient and the same policy:
    // same member ID when both files have one, else the same Sohar payer.
    const discoveryVerification = [];
    const specialty = cfg.specialty || 'PSY';
    const idRejected = (per, payerId) => results.some((r) => r['Sent As'] === 'Verification' && r['Per Nbr'] === per
      && r['Sohar payerId'] === payerId && /^error\.member\.id/i.test(str(r['Sohar Status'])));
    for (const d of results) {
      if (d['Sent As'] !== 'Discovery' || d['Sohar Status'] !== 'Coverage found' || !d['Per Nbr'] || !d['Sohar payerId']) continue;
      const p = byPer.get(d['Per Nbr']);
      const payer = d['Sohar Payer'].toLowerCase();
      const member = idKey(d['Sohar Member ID']);
      const samePolicy = (r) => (member && idKey(r['Sohar Member ID']) ? idKey(r['Sohar Member ID']) === member : !!payer && r['Sohar Payer'].toLowerCase() === payer);
      // Patient first sent for verification: this discovery followed a failed verification, so only a newer one counts.
      const dT = soharTime(d['Sohar Date']);
      const afterDiscovery = (r) => p.Route !== 'verification' || isNaN(dT) || soharTime(r['Sohar Date']) > dT;
      const v = results
        .filter((r) => r['Sent As'] === 'Verification' && r['Per Nbr'] === d['Per Nbr'] && r['Sohar Status'] !== 'Not in Sohar file' && samePolicy(r) && afterDiscovery(r))
        .sort((a, b) => String(b['Sohar Date']).localeCompare(String(a['Sohar Date'])))[0];
      if (!v) {
        d['Discovery Follow-up'] = 'Ready to verify';
        d.Action = `${d.Action} Copay, deductible and OOP need a Sohar verification: this patient is on the discovery verification import.`;
        discoveryVerification.push({
          firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State,
          // Member ID from discovery; NextGen's only when NextGen already has this payer and Sohar did not reject it.
          memberId: d['Sohar Member ID'] || (p['Sohar payerId'] === d['Sohar payerId'] && !idRejected(d['Per Nbr'], d['Sohar payerId']) ? p['Member ID'] : ''),
          payerId: d['Sohar payerId'], placeOfServiceCode: p.POS, specialtyCode: specialty, id: p['Per Nbr'],
          subscriberFirstName: p['Subscriber First'], subscriberLastName: p['Subscriber Last'], subscriberDateOfBirth: p['Subscriber DOB'],
        });
        continue;
      }
      d['Discovery Follow-up'] = `Verified: ${v['Sohar Status']}`;
      v['Discovery Follow-up'] = 'From discovery';
      const inNextGen = /^Coverage found matches/.test(d.Action);
      if (/^complete\.eligible/i.test(v['Sohar Status'])) {
        if (inNextGen) d.Action = 'Coverage found matches NextGen. Verified eligible.';
        else {
          d.Action = `${d.Action.replace(', then verify.', '.')} Verified eligible; benefits are on the verification row.`;
          if (!/^Coverage ends/.test(v.Action)) v.Action = `Eligible on ${v['Sohar Payer'] || d['Sohar Payer']}, the plan discovery found. Add it to NextGen; benefits are on this row.`;
        }
      } else {
        d.Action = `${d.Action} Verification came back ${v['Sohar Status']}; see the verification row.`;
      }
    }

    // Two insurances verified for one patient (R3 / R4): keep every row, mark the better one.
    // Rank: complete.eligible > pending > ineligible > error; ties go to the latest Sohar Date.
    // The other policy's row becomes "Superseded: ..." (work queue Clear, no NextGen task).
    counts.twoInsurances = 0;
    const vRank = (r) => { const s = str(r['Sohar Status']).toLowerCase(); return /^complete\.eligible/.test(s) ? 0 : /^pending/.test(s) || !s ? 1 : /^complete\.ineligible/.test(s) ? 2 : 3; };
    const vByPer = new Map();
    for (const r of results) {
      if (r['Sent As'] !== 'Verification' || !r['Per Nbr'] || r.Match === 'No match' || r['Sohar Status'] === 'Not in Sohar file') continue;
      const policy = `${idKey(r['Sohar Member ID'])}|${str(r['Sohar Payer']).toLowerCase()}`;
      if (!vByPer.has(r['Per Nbr'])) vByPer.set(r['Per Nbr'], new Map());
      const m = vByPer.get(r['Per Nbr']);
      const prev = m.get(policy);
      if (!prev || String(r['Sohar Date']) > String(prev['Sohar Date'])) m.set(policy, r);
    }
    // Only policies checked together count: saved uploads build up across runs, so a policy whose latest result is
    // more than TWO_INS_DAYS older than the patient's newest result is history, not a second insurance.
    const TWO_INS_DAYS = 3;
    const dayMs = (r) => { const d = parseDate(r['Sohar Date']); return d ? d.getTime() : NaN; };
    for (const [per, all] of vByPer) {
      const newest = Math.max(...[...all.values()].map(dayMs).filter((n) => !isNaN(n)));
      const m = [...all.values()].filter((r) => { const t = dayMs(r); return isNaN(t) || !isFinite(newest) || newest - t <= TWO_INS_DAYS * 86400000; });
      if (m.length < 2) continue;
      counts.twoInsurances++;
      const list = m.sort((a, b) => vRank(a) - vRank(b) || String(b['Sohar Date']).localeCompare(String(a['Sohar Date'])));
      const best = list[0];
      if (vRank(best) !== 0) {
        for (const r of list) r.Action = `${r.Action} Two insurances were checked; neither came back eligible.`;
        continue;
      }
      const bestPayer = best['Sohar Payer'] || 'the other insurance';
      for (const r of list.slice(1)) r.Action = `Superseded: ${r['Sohar Status'] || 'no status'} on ${r['Sohar Payer'] || 'this insurance'}. ${bestPayer} came back eligible for this patient; use that row.`;
      const p = byPer.get(per);
      const vr = p ? matchVerifyRow(p, { memberid: best['Sohar Member ID'], payername: best['Sohar Payer'] }) : null;
      if (!vr || vr.source === 'guarantor') {
        if (/^Eligible\. No action/.test(best.Action)) best.Action = `${best.Action} Two insurances were checked; the chart insurance is the active one.`;
        else best.Action = `${best.Action} Two insurances were checked; this one is active.`;
        continue;
      }
      const options = nextgenFor.get(vr.payerId) || [];
      const choose = vr.nextgenPayer && (!options.length || options.some((o) => o.toLowerCase() === vr.nextgenPayer.toLowerCase())) ? vr.nextgenPayer
        : options.length === 1 ? options[0] : options.length ? `One of: ${options.join(' | ')}` : vr.nextgenPayer || '';
      best['Choose in NextGen'] = choose;
      const text = `Eligible on ${bestPayer}, the ${vr.source} insurance. Two insurances were checked; make this the chart insurance in NextGen${choose ? ` (choose ${/^One of/.test(choose) ? choose.replace(/^One of: /, 'one of: ') : `"${choose}"`})` : ''}.`;
      best.Action = /^Coverage ends/.test(best.Action) ? `${best.Action} ${text}` : text;
    }

    // Verification failed (ineligible or any error) -> discovery, so Sohar can look for the patient's real coverage.
    // Only patients with an upcoming visit, and only each policy's latest result for policies still on the chart
    // (exact member ID, else the Sohar payer; old history for a policy since replaced does not count).
    // Skipped when any current policy is eligible or pending, when the failure was itself the follow-up of a
    // discovery (no looping), or when a discovery already ran after the failure.
    const verificationDiscovery = [];
    const failed = (r) => /^(complete\.ineligible|error)/i.test(str(r['Sohar Status']));
    const onChart = (p, r) => {
      const vrs = (p.x && p.x.verifyRows) || [];
      const mid = idKey(r['Sohar Member ID']);
      if (mid) return [p['Member ID'], p['Member ID (Lst Enc)'], ...vrs.map((v) => v.memberId)].some((m) => idKey(m) === mid);
      const pn = str(r['Sohar Payer']).toLowerCase();
      return !!pn && vrs.some((v) => str(v.payerName).toLowerCase() === pn);
    };
    const discTimes = new Map();
    for (const r of results) {
      if (r['Sent As'] !== 'Discovery' || !r['Per Nbr'] || r['Sohar Status'] === 'Not in Sohar file') continue;
      if (!discTimes.has(r['Per Nbr'])) discTimes.set(r['Per Nbr'], []);
      discTimes.get(r['Per Nbr']).push(soharTime(r['Sohar Date']));
    }
    for (const [per, all] of vByPer) {
      const p = byPer.get(per);
      if (!p || !upcoming.has(per)) continue;
      const latest = [...all.values()].filter((r) => onChart(p, r));
      if (!latest.length || !latest.every(failed) || latest.some((r) => r['Discovery Follow-up'] === 'From discovery')) continue;
      const failT = Math.max(...latest.map((r) => soharTime(r['Sohar Date'])));
      if ((discTimes.get(per) || []).some((t) => isNaN(t) || isNaN(failT) || t >= failT)) continue;
      // Payer as a hint only for errors (payer right, ID or name wrong); ineligible may mean the patient changed plans.
      const hint = latest.find((r) => /^error/i.test(str(r['Sohar Status'])) && r['Sohar payerId']);
      verificationDiscovery.push({
        firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, zipCode: p.Zip,
        payerId: hint ? hint['Sohar payerId'] : '', id: p['Per Nbr'],
      });
      for (const r of latest) {
        r['Discovery Follow-up'] = 'Sent to discovery';
        r.Action = `${r.Action} This patient is on the discovery import (from verification).`;
      }
    }

    for (const row of results) {
      row['Work Queue'] = workQueue(row);
      counts.byAction[row.Action] = (counts.byAction[row.Action] || 0) + 1;
      counts.byQueue[row['Work Queue']] = (counts.byQueue[row['Work Queue']] || 0) + 1;
      if (row['Deductible Status']) counts.deductible[row['Deductible Status']] = (counts.deductible[row['Deductible Status']] || 0) + 1;
      if (row['OOP Status']) counts.oop[row['OOP Status']] = (counts.oop[row['OOP Status']] || 0) + 1;
    }

    // Soonest appointment first; rows with no upcoming appointment last.
    const t = (r) => (r['Next Appt'] ? parseDate(r['Next Appt']).getTime() : Infinity);
    results.sort((a, b) => t(a) - t(b) || a['Last Name'].localeCompare(b['Last Name']));
    return { results, unknownFiles, counts, asOf, discoveryVerification, verificationDiscovery };
  }

  /** Stable key for one Sohar result, so a re-upload replaces it and a new file adds to the saved set. */
  function resultKey(kind, r) {
    if (kind === 'discovery') return `d|${str(r.id) || str(r.requestid) || `${nameKey(r.firstname, r.lastname)}|${str(r.dateofbirth)}`}`;
    const who = str(r.id) || str(r.verificationid || r.requestid) || `${idKey(r.memberid)}|${nameKey(r.firstname, r.lastname)}`;
    // Member ID too: two policies with the same payer (R3) are different results.
    return `v|${who}|${str(r.payername).toLowerCase()}|${idKey(r.memberid)}`;
  }

  /**
   * Add newly uploaded Sohar files to the saved result records. Same key: the new upload wins.
   * saved: rows of UPLOAD_COLS. Returns {saved, files: [{name, kind, records}], added, replaced, unknownFiles}.
   */
  function accumulateResults(saved, uploads, uploadedAt) {
    const by = new Map();
    for (const s of saved || []) {
      let rec; try { rec = JSON.parse(s.Record); } catch (e) { continue; }
      const kind = str(s.Kind), key = kind ? resultKey(kind, rec) : String(s.Key); // re-keyed so older saved keys migrate
      by.set(key, { Kind: kind, Key: key, File: str(s.File), Uploaded: str(s.Uploaded), rec });
    }
    let added = 0, replaced = 0;
    const unknownFiles = [];
    for (const f of uploads || []) {
      const { kind, records } = readResultRows(f.rows);
      if (!kind) { unknownFiles.push(f.name); continue; }
      for (const rec of records) {
        const key = resultKey(kind, rec);
        if (by.has(key)) replaced++; else added++;
        by.set(key, { Kind: kind, Key: key, File: f.name, Uploaded: uploadedAt || '', rec });
      }
    }
    const files = ['verification', 'discovery'].map((kind) => ({ name: `saved ${kind} results`, kind, records: [...by.values()].filter((x) => x.Kind === kind).map((x) => x.rec) }))
      .filter((f) => f.records.length);
    const out = [...by.values()].map((x) => ({ Kind: x.Kind, Key: x.Key, File: x.File, Uploaded: x.Uploaded, Record: JSON.stringify(x.rec) }));
    return { saved: out, files, added, replaced, unknownFiles };
  }

  /** Keep member IDs an admin typed on the Discovery Verify sheet, matched on id (Per Nbr) + payerId. */
  function keepMemberIds(rows, previous) {
    const typed = new Map();
    for (const p of previous || []) if (str(p.memberId)) typed.set(`${str(p.id)}|${str(p.payerId)}`, str(p.memberId));
    return rows.map((r) => (r.memberId ? r : { ...r, memberId: typed.get(`${r.id}|${r.payerId}`) || '' }));
  }

  // ---------- NextGen task list ----------
  const TASK_TYPES = ['Add insurance in NextGen', 'Update insurance in NextGen', 'Fix member ID in NextGen', 'Fix name in NextGen',
    'Update benefits in NextGen', 'Add payer to Payer Map', 'Confirm self-pay or get card', 'Find patient for Sohar result', 'Resend to Sohar', 'Send for verification'];

  /**
   * One task per thing an admin has to do in NextGen (or the Payer Map) before the visit.
   * Every eligible verification with benefits gets an "Update benefits" task with the values to enter,
   * because the appointment report does not show what NextGen currently holds.
   */
  function buildTasks(results) {
    const tasks = [];
    const add = (r, task, details) => {
      const key = `${r['Per Nbr'] || 'ref:' + r['Sohar Ref']}|${task}`;
      tasks.push({
        'Done': '', 'Due By': r['Next Appt'], 'Next Appt': r['Next Appt'], 'Next Time': r['Next Time'], 'Next Location': r['Next Location'],
        'Per Nbr': r['Per Nbr'], 'MRN': r.MRN, 'Last Name': r['Last Name'], 'First Name': r['First Name'], 'Result Type': r['Sent As'],
        'Task': task, 'Details': details, 'Choose in NextGen': r['Choose in NextGen'] || '', 'Sohar Link': r['Sohar Link'], 'Sohar Ref': r['Sohar Ref'], 'Task Key': key,
      });
    };
    const benefitText = (r) => [
      r.Copay !== '' ? `Copay ${r.Copay}` : '', r.Coinsurance ? `Coinsurance ${r.Coinsurance}` : '',
      r.Deductible !== '' ? `Deductible ${r.Deductible} (${r['Deductible Remaining'] || 0} left, ${r['Deductible Status'] || '?'})` : '',
      r['OOP Max'] !== '' ? `OOP max ${r['OOP Max']} (${r['OOP Remaining'] || 0} left, ${r['OOP Status'] || '?'})` : '',
      r['Coverage Level'] ? `Level ${r['Coverage Level']}` : '', r.Plan ? `Plan ${r.Plan}` : '',
      r['Coverage Start'] || r['Coverage End'] ? `Coverage ${r['Coverage Start'] || '?'} to ${r['Coverage End'] || 'open'}` : '',
    ].filter(Boolean).join('; ');
    for (const r of results || []) {
      const status = String(r['Sohar Status'] || '').toLowerCase();
      const memberHint = r['Sohar Member ID'] ? `Member ID ${r['Sohar Member ID']}.` : 'Member ID: open the Sohar link.';
      if (/^Superseded/.test(r.Action || '')) continue; // the other insurance came back eligible
      if (r.Match === 'No match') { add(r, 'Find patient for Sohar result', `Sohar ${r['Sent As'].toLowerCase()} result did not match the Import report. ${r.Action}`); continue; }
      if (status === 'not in sohar file') { add(r, 'Resend to Sohar', 'Upcoming patient missing from the Sohar results.'); continue; }
      if (r['Sent As'] === 'Discovery') {
        if (status !== 'coverage found') { if (status === 'no coverage found') add(r, 'Confirm self-pay or get card', 'Sohar found no coverage.'); continue; }
        const choose = r['Choose in NextGen'] || '';
        if (/^Not in Payer Map/.test(choose)) add(r, 'Add payer to Payer Map', `${r['Sohar Payer']} (${r['Sohar payerId']}) has no NextGen insurance in the Payer Map.`);
        const pick = choose && !/^Not in Payer Map/.test(choose) ? `Choose ${choose.replace(/^One of: /, 'one of: ')}.` : `Sohar payer: ${r['Sohar Payer']}.`;
        const verified = /^Verified/.test(r['Discovery Follow-up'] || '');
        const next = verified ? `Sohar ${r['Discovery Follow-up'].toLowerCase()}; benefits are on the verification row.` : 'Then send for verification to get benefits (step 5: discovery verification import).';
        if (!r['NextGen Payer'] || /^self pay$/i.test(r['NextGen Payer'])) add(r, 'Add insurance in NextGen', `${pick} ${memberHint} ${next}`);
        else if (r['Coverage Differs'] === 'Yes') add(r, 'Update insurance in NextGen', `NextGen has ${r['NextGen Payer']}; Sohar found ${r['Sohar Payer']}. Confirm with the patient. ${pick} ${memberHint}${verified ? ` ${next}` : ''}`);
        else if (!verified) add(r, 'Send for verification', `Discovery matches NextGen (${r['NextGen Payer']}). It is on the discovery verification import.`);
        continue;
      }
      // Verification
      if (status.startsWith('complete.ineligible')) add(r, 'Update insurance in NextGen', `Sohar: ${r['Sohar Status']} for ${r['Sohar Payer'] || r['NextGen Payer']}. Confirm current coverage with the patient; ${r['Discovery Follow-up'] === 'Sent to discovery' ? 'they are on the discovery import (from verification).' : 'consider discovery.'}`);
      else if (status === 'error.member.id') add(r, 'Fix member ID in NextGen', `Sohar rejected member ID ${r['NextGen Member ID'] || '(blank)'} for ${r['NextGen Payer']}. Check the card, then re-verify.`);
      else if (status === 'error.member.name') add(r, 'Fix name in NextGen', 'Sohar rejected the name. Match it to the insurance card, then re-verify.');
      else if (status === 'complete.eligible') {
        if (/^Coverage ends|make this the chart insurance/.test(r.Action)) add(r, 'Update insurance in NextGen', r.Action);
        const ben = benefitText(r);
        if (ben) add(r, 'Update benefits in NextGen', ben);
      }
    }
    const t = (x) => (x['Due By'] ? parseDate(x['Due By']).getTime() : Infinity);
    const mins = (x) => { const m = String(x['Next Time'] || '').match(/(\d+):(\d+)\s*(AM|PM)/i); return m ? ((+m[1] % 12) + (/pm/i.test(m[3]) ? 12 : 0)) * 60 + +m[2] : 9999; };
    tasks.sort((a, b) => t(a) - t(b) || mins(a) - mins(b) || String(a['Next Location']).localeCompare(String(b['Next Location'])) || String(a['Last Name']).localeCompare(String(b['Last Name'])) || TASK_TYPES.indexOf(a.Task) - TASK_TYPES.indexOf(b.Task));
    return tasks;
  }

  /** Open tasks as a note for admins, by location then visit. LASFIR + Per Nbr only. */
  function taskMessage(tasks, opts) {
    const open = (tasks || []).filter((t) => !String(t.Done || '').trim());
    if (!open.length) return 'NextGen tasks: nothing open.\n';
    const byLoc = new Map();
    for (const t of open) {
      const loc = t['Next Location'] || 'No upcoming appointment';
      if (!byLoc.has(loc)) byLoc.set(loc, []);
      byLoc.get(loc).push(t);
    }
    const o = opts || {};
    const lines = [`NextGen updates before the visit${o.asOf ? ` (appointments from ${fmtDate(o.asOf)})` : ''}: ${open.length} open task${open.length > 1 ? 's' : ''}.`, ''];
    for (const [loc, list] of [...byLoc].sort((a, b) => a[0].localeCompare(b[0]))) {
      lines.push(`${loc} (${list.length})`);
      for (const t of list) {
        const when = t['Next Appt'] ? `${String(t['Next Appt']).slice(0, 5)}${t['Next Time'] ? ' ' + t['Next Time'] : ''}` : 'no appt';
        const tag = lasfir(t['Last Name'], t['First Name']);
        const who = t['Per Nbr'] ? `${tag} · Per ${t['Per Nbr']}` : `Sohar ref ${t['Sohar Ref'] || '?'}`;
        lines.push(`- ${when} · ${who} · ${t.Task}: ${t.Details}`);
      }
      lines.push('');
    }
    return lines.join('\n').trim() + '\n';
  }

  /** Carry Done marks over from the previous task table, matched on Task Key. */
  function keepDone(tasks, previous) {
    const done = new Map();
    for (const p of previous || []) if (String(p['Done'] || '').trim()) done.set(String(p['Task Key']), String(p['Done']).trim());
    return tasks.map((t) => (done.has(t['Task Key']) ? { ...t, Done: done.get(t['Task Key']) } : t));
  }

  // ---------- pane <-> dialog messages ----------
  // Office caps the size of one dialog message, so payloads travel as numbered pieces.
  function chunkMessage(obj, size) {
    const s = JSON.stringify(obj), n = Math.max(1, Math.ceil(s.length / (size || 20000)));
    const id = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
    const out = [];
    for (let i = 0; i < n; i++) out.push(JSON.stringify({ type: 'chunk', id, i, n, data: s.slice(i * (size || 20000), (i + 1) * (size || 20000)) }));
    return out;
  }
  /** Returns a function that takes one raw message and returns the whole object once every piece has arrived, else null. */
  function messageAssembler() {
    const parts = new Map();
    return (raw) => {
      let m; try { m = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch (e) { return null; }
      if (!m || m.type !== 'chunk') return m || null;
      if (!parts.has(m.id)) parts.set(m.id, []);
      const list = parts.get(m.id); list[m.i] = m.data;
      if (list.filter((x) => x != null).length < m.n) return null;
      parts.delete(m.id);
      return JSON.parse(list.join(''));
    };
  }

  // ---------- IR Portal cross-reference ----------
  // Column names follow the IR Portal Tracker add-in (Filtered tab; headers in row 1, data may start in column C).
  const IR_HEADERS = {
    id: ['incident id', 'incident'], last: ['last name', 'last'], first: ['first name', 'first'], dob: ['dob', 'date of birth'],
    desc: ['description', 'description option', 'description type'], fac: ['facility', 'location'], mrn: ['mrn'],
    dos: ['dos', 'date of service'], prov: ['provider', 'rendering provider'], due: ['due date', 'due'], act: ['last activity'],
  };
  const IR_COLS = ['Category', 'Description', 'Incident ID', 'DOS', 'Due Date', 'Facility', 'Provider', 'MRN', 'Last Name', 'First Name',
    'Per Nbr', 'Next Appt', 'Next Location', 'Sohar Result', 'Sohar Payer', 'Plan', 'Coverage Start', 'Coverage End', 'Benefit Flag',
    'Choose in NextGen', 'Open NextGen Tasks', 'Suggested Action', 'Match'];

  /** Insurance-relevant IR Portal descriptions only; everything else (signatures, charts, Dx) returns ''. */
  function irCategory(desc) {
    const d = str(desc).toLowerCase();
    if (/missing insurance/.test(d)) return 'Missing insurance';
    if (/verify insurance|insurance on file inactive|member not found/.test(d)) return 'Verify insurance';
    if (/missing demo/.test(d)) return 'Missing demographics';
    if (/out of network/.test(d)) return 'Out of network';
    if (/auth/.test(d)) return 'Authorization';
    return '';
  }

  /** IR Portal rows (2D) -> {records} or {error}. Finds the header row by name, wherever the columns start. */
  function parseIrPortal(rows) {
    const want = (h) => Object.keys(IR_HEADERS).find((k) => IR_HEADERS[k].includes(str(h).toLowerCase()));
    const hi = (rows || []).slice(0, 30).findIndex((r) => { const ks = (r || []).map(want); return ks.includes('mrn') && ks.includes('desc'); });
    if (hi < 0) return { error: 'This does not look like the IR Portal Filtered tab: it needs MRN and Description columns.' };
    const idx = {};
    rows[hi].forEach((h, i) => { const k = want(h); if (k && idx[k] == null) idx[k] = i; });
    const records = [];
    for (const r of rows.slice(hi + 1)) {
      if (!(r || []).some((v) => !isBlank(v))) continue;
      const o = {};
      for (const [k, i] of Object.entries(idx)) o[k] = r[i];
      if (isBlank(o.mrn) && isBlank(o.last)) continue;
      records.push(o);
    }
    return { records };
  }

  function irAction(cat, ver, disc, dos) {
    const payer = (r) => r['Sohar Payer'] || r['NextGen Payer'] || 'the payer';
    const verOk = ver && /^complete\.eligible/i.test(ver['Sohar Status']);
    const verBad = ver && /^complete\.ineligible/i.test(ver['Sohar Status']);
    const found = disc && disc['Sohar Status'] === 'Coverage found';
    const none = disc && disc['Sohar Status'] === 'No coverage found';
    const start = ver ? parseDate(ver['Coverage Start']) : null;
    const lateStart = start && dos && start > dos ? ` Coverage started ${fmtDate(start)}, after the ${fmtDate(dos)} DOS.` : '';
    const pick = found && disc['Choose in NextGen'] ? ` (${disc['Choose in NextGen']})` : '';
    if (cat === 'Missing insurance' || cat === 'Missing demographics') {
      if (found) return `Sohar found ${payer(disc)}. Add it in NextGen${pick}, then resolve the IR item.`;
      if (verOk) return `Sohar verified ${payer(ver)} active${ver.Plan ? ` (${ver.Plan})` : ''}. Confirm it is on the patient in NextGen and on the claim, then resolve the IR item.${lateStart}`;
      if (none) return 'Sohar found no coverage. Confirm self-pay or get the card from the patient.';
      if (verBad) return `Sohar shows ${payer(ver)} inactive. Get current insurance from the patient.`;
    }
    if (cat === 'Verify insurance') {
      if (verOk) return `Sohar shows ${payer(ver)} active now. Check it covered the DOS before rebilling.${lateStart}`;
      if (verBad) return `Sohar agrees ${payer(ver)} is inactive. Run discovery or get new insurance from the patient.`;
      if (found) return `Sohar found ${payer(disc)}. Update NextGen${pick} and rebill.`;
      if (none) return 'Sohar found no coverage. Confirm self-pay or get the card.';
    }
    if (cat === 'Out of network') return `Sohar shows ${payer(ver || disc)}${ver && ver.Network ? `, network type ${ver.Network}` : ''}. Sohar does not check provider participation; confirm the network with the payer.`;
    if (cat === 'Authorization') return `Insurance per Sohar: ${payer(ver || disc)}${ver && ver.Plan ? ` (${ver.Plan})` : ''}. Sohar does not check authorizations; request or update the auth.`;
    return (ver || disc).Action || 'Review.';
  }

  /**
   * IR Portal items that touch insurance, joined to Sohar results and open NextGen tasks for the same patient.
   * Match: MRN (ignoring punctuation and leading zeros), else last + first name + DOB.
   */
  function irCrossReference(irRecords, results, tasks) {
    const mk = (v) => str(v).toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+/, '');
    const byMrn = new Map(), byName = new Map();
    for (const r of results || []) {
      if (r.Match === 'No match') continue;
      const m = mk(r.MRN);
      if (m) { if (!byMrn.has(m)) byMrn.set(m, []); byMrn.get(m).push(r); }
      if (r['Last Name'] && r.DOB) { const k = `${nameKey(r['First Name'], r['Last Name'])}|${r.DOB}`; if (!byName.has(k)) byName.set(k, []); byName.get(k).push(r); }
    }
    const openTasks = (per) => (tasks || []).filter((t) => t['Per Nbr'] === per && !str(t.Done)).map((t) => t.Task);
    const out = [];
    for (const ir of irRecords || []) {
      const cat = irCategory(ir.desc);
      if (!cat) continue;
      const dos = parseDate(ir.dos), due = parseDate(ir.due), dob = parseDate(ir.dob);
      let rs = byMrn.get(mk(ir.mrn)) || [], match = rs.length ? 'MRN' : '';
      if (!rs.length && dob) { rs = byName.get(`${nameKey(ir.first, ir.last)}|${fmtDate(dob)}`) || []; if (rs.length) match = 'Name + DOB'; }
      const latest = (kind) => rs.filter((r) => r['Sent As'] === kind && r['Sohar Status'] !== 'Not in Sohar file').sort((a, b) => String(b['Sohar Date']).localeCompare(String(a['Sohar Date'])))[0];
      const ver = latest('Verification'), disc = latest('Discovery'), any = ver || disc || {};
      if (!ver && !disc) continue; // only items Sohar has something to say about

      out.push({
        'Category': cat, 'Description': str(ir.desc), 'Incident ID': str(ir.id), 'DOS': fmtDate(dos), 'Due Date': fmtDate(due),
        'Facility': str(ir.fac), 'Provider': str(ir.prov), 'MRN': str(ir.mrn), 'Last Name': str(ir.last), 'First Name': str(ir.first),
        'Per Nbr': any['Per Nbr'] || '', 'Next Appt': any['Next Appt'] || '', 'Next Location': any['Next Location'] || '',
        'Sohar Result': [ver ? `Verification: ${ver['Sohar Status']}` : '', disc ? `Discovery: ${disc['Sohar Status']}` : ''].filter(Boolean).join('; '),
        'Sohar Payer': (ver && ver['Sohar Payer']) || (disc && disc['Sohar Payer']) || '', 'Plan': (ver && ver.Plan) || '',
        'Coverage Start': (ver && ver['Coverage Start']) || '', 'Coverage End': (ver && ver['Coverage End']) || '',
        'Benefit Flag': (ver && ver['Benefit Flag']) || '', 'Choose in NextGen': (disc && disc['Choose in NextGen']) || '',
        'Open NextGen Tasks': any['Per Nbr'] ? openTasks(any['Per Nbr']).join('; ') : '',
        'Suggested Action': irAction(cat, ver, disc, dos), 'Match': match,
      });
    }
    // Most urgent first: due date, then DOS.
    const t = (s) => (s ? parseDate(s).getTime() : Infinity);
    out.sort((a, b) => t(a['Due Date']) - t(b['Due Date']) || t(a.DOS) - t(b.DOS) || a.Facility.localeCompare(b.Facility));
    return out;
  }

  /** Cross-reference as a note for admins, by facility. LASFIR + MRN last 4 only. */
  function irMessage(rows) {
    if (!rows.length) return 'IR Portal cross-check: no insurance-related IR items.\n';
    const byFac = new Map();
    for (const r of rows) { const f = r.Facility || 'No facility'; if (!byFac.has(f)) byFac.set(f, []); byFac.get(f).push(r); }
    const lines = [`IR Portal cross-check with Sohar: ${rows.length} insurance-related item${rows.length > 1 ? 's' : ''}.`, ''];
    for (const [f, list] of [...byFac].sort((a, b) => a[0].localeCompare(b[0]))) {
      lines.push(`${f} (${list.length})`);
      for (const r of list) lines.push(`- ${lasfir(r['Last Name'], r['First Name'])} · MRN ...${String(r.MRN).slice(-4)} · DOS ${r.DOS || '?'} · ${r.Category}: ${r['Suggested Action']}`);
      lines.push('');
    }
    return lines.join('\n').trim() + '\n';
  }

  /** Verified patients' benefits, one row per Per Nbr (latest Sohar result wins). Contract for the TMS Tracker import. */
  function benefitsExport(results) {
    const by = new Map();
    for (const r of results || []) {
      if (r['Sent As'] !== 'Verification' || !r['Per Nbr'] || !r['Sohar Ref'] || /^Superseded/.test(r.Action || '')) continue;
      const prev = by.get(r['Per Nbr']);
      if (!prev || String(r['Sohar Date']) > String(prev['Sohar Date'])) by.set(r['Per Nbr'], r);
    }
    return [...by.values()].map((r) => Object.fromEntries(BENEFITS_EXPORT_COLS.map((c) => [c, r[c] == null ? '' : r[c]])));
  }

  // Final file for the TMS Tracker and ESK Tracker imports: the benefits columns, then these. Add columns at the end only.
  // Secondary Payer and Cost-share Flags (Secondary may cover / QMB / VIP on the day of export) let a later copay-log
  // report join on Per Nbr + Sohar Date and spot copays a secondary or QMB should have covered.
  const TRACKER_EXPORT_COLS = [...BENEFITS_EXPORT_COLS, 'Service', 'Source', 'Exported', 'Secondary Payer', 'Cost-share Flags'];
  const COST_SHARE_PILLS = ['Secondary may cover', 'QMB', 'VIP'];

  /** "TMS", "ESK", "TMS; ESK" or '' from a patient's upcoming NextGen events. */
  function serviceLine(text) {
    const s = String(text || '');
    return [/\btms\b|magv|\bmt\b/i.test(s) ? 'TMS' : '', /esketamine|spravato|\besk\b/i.test(s) ? 'ESK' : ''].filter(Boolean).join('; ');
  }

  /**
   * One row per patient with a Sohar verification, for the tracker imports. Patients whose discovered coverage
   * is still waiting on a verification are left out and returned as pending, so their benefits are not stale.
   */
  function trackerExport(results, exportedAt, secondaryByPer) {
    const sec = (per) => (secondaryByPer instanceof Map ? secondaryByPer.get(per) : secondaryByPer && secondaryByPer[per]) || '';
    const pendingPer = new Set((results || []).filter((r) => r['Discovery Follow-up'] === 'Ready to verify' && r['Per Nbr']).map((r) => r['Per Nbr']));
    const events = new Map();
    for (const r of results || []) if (r['Per Nbr']) events.set(r['Per Nbr'], `${events.get(r['Per Nbr']) || ''} ${r['Next Event']} ${r['Upcoming Detail']}`);
    const by = new Map();
    for (const r of results || []) {
      if (r['Sent As'] !== 'Verification' || !r['Per Nbr'] || !r['Sohar Ref'] || pendingPer.has(r['Per Nbr']) || /^Superseded/.test(r.Action || '')) continue;
      const prev = by.get(r['Per Nbr']);
      if (!prev || String(r['Sohar Date']) > String(prev['Sohar Date'])) by.set(r['Per Nbr'], r);
    }
    const rows = [...by.values()].map((r) => ({
      ...Object.fromEntries(BENEFITS_EXPORT_COLS.map((c) => [c, r[c] == null ? '' : r[c]])),
      'Service': serviceLine(events.get(r['Per Nbr'])),
      'Source': r['Discovery Follow-up'] === 'From discovery' ? 'Discovery, then verification' : 'Verification',
      'Exported': exportedAt || '',
      'Secondary Payer': str(sec(r['Per Nbr']) || r['Secondary Payer']),
      'Cost-share Flags': String(r.Pills || '').split(';').map((x) => x.trim()).filter((x) => COST_SHARE_PILLS.includes(x)).join('; '),
    }));
    const pending = (results || []).filter((r) => r['Discovery Follow-up'] === 'Ready to verify' && r['Per Nbr']);
    return { rows, pending };
  }

  // ---------- saved file names ----------
  /** Local calendar date as YYYY-MM-DD (not UTC, so an evening save is still today's date). */
  function localDate(d) {
    const x = d || new Date();
    return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`;
  }
  const FILE_LABELS = {
    verification: 'Sohar Ready Import - Verification',
    discovery: 'Sohar Ready Import - Discovery',
    discoveryVerify: 'Sohar Ready Import - Verification (from Discovery)',
    verifyFailed: 'Sohar Ready Import - Discovery (from Verification)',
    tracker: 'Sohar Tracker Import (TMS + ESK)',
    benefits: 'Sohar Benefits',
    worklist: 'Sohar Results Worklist',
    tasks: 'NextGen Tasks',
    ir: 'IR Portal Cross-Reference',
  };
  /** "Sohar Ready Import - Verification 2026-10-15.csv" */
  function fileName(kind, d) { return `${FILE_LABELS[kind] || kind} ${localDate(d)}.csv`; }

  // ---------- benefits sheet ----------
  // Benefits sheet (created by the add-in): one row per verified patient, OOP max closest to met first.
  const BENEFIT_SHEET_COLS = ['Flag', 'Next Appt', 'Next Time', 'Next Location', 'Service', 'Per Nbr', 'MRN', 'Last Name', 'First Name',
    'Sohar Payer', 'Plan', 'Copay', 'Coinsurance',
    'OOP Max', 'OOP Met', 'OOP Remaining', 'OOP % Met', 'OOP Status',
    'Deductible', 'Deductible Met', 'Deductible Remaining', 'Deductible % Met', 'Deductible Status',
    'Coverage Level', 'Coverage End', 'Source', 'Sohar Date'];
  const BENEFIT_PCT_COLS = ['OOP % Met', 'Deductible % Met'];

  /** Spent so far and share of the total, from total and remaining. Blank when either is unknown or the total is $0. */
  function spent(total, remaining) {
    const t = money(total), r = money(remaining);
    if (t == null || r == null || t <= 0) return { met: '', pct: '' };
    const met = Math.min(t, Math.max(0, t - r));
    return { met: fmtMoney(met), pct: Math.round((met / t) * 1000) / 1000 };
  }

  const FLAG_RANK = { 'OOP max met': 0, 'OOP max approaching': 1, 'Deductible met': 2, 'Deductible approaching': 3, '': 4 };
  function benefitSheetFlag(r) {
    if (r['OOP Status'] === 'Met') return 'OOP max met';
    if (r['OOP Status'] === 'Approaching') return 'OOP max approaching';
    if (r['Deductible Status'] === 'Met') return 'Deductible met';
    if (r['Deductible Status'] === 'Approaching') return 'Deductible approaching';
    return '';
  }

  /**
   * Benefits sheet rows from the merged Sohar results: latest verification per patient, skipping ones whose
   * discovered coverage still waits on a verification. Sorted: OOP met, OOP approaching, then highest OOP % met.
   * Percent columns are fractions (0.85), so Excel can show them as % with data bars.
   */
  function benefitsView(results) {
    const tr = trackerExport(results, '');
    const byPer = new Map((results || []).filter((r) => r['Sent As'] === 'Verification' && !/^Superseded/.test(r.Action || '')).map((r) => [r['Per Nbr'], r]));
    const rows = tr.rows.map((b) => {
      const r = byPer.get(b['Per Nbr']) || {};
      const oop = spent(b['OOP Max'], b['OOP Remaining']);
      const ded = spent(b.Deductible, b['Deductible Remaining']);
      return {
        'Flag': benefitSheetFlag(b), 'Next Appt': str(r['Next Appt']), 'Next Time': str(r['Next Time']), 'Next Location': str(r['Next Location']),
        'Service': b.Service, 'Per Nbr': b['Per Nbr'], 'MRN': b.MRN, 'Last Name': b['Last Name'], 'First Name': b['First Name'],
        'Sohar Payer': b['Sohar Payer'], 'Plan': b.Plan, 'Copay': b.Copay, 'Coinsurance': b.Coinsurance,
        'OOP Max': b['OOP Max'], 'OOP Met': oop.met, 'OOP Remaining': b['OOP Remaining'], 'OOP % Met': oop.pct, 'OOP Status': b['OOP Status'],
        'Deductible': b.Deductible, 'Deductible Met': ded.met, 'Deductible Remaining': b['Deductible Remaining'], 'Deductible % Met': ded.pct, 'Deductible Status': b['Deductible Status'],
        'Coverage Level': b['Coverage Level'], 'Coverage End': b['Coverage End'], 'Source': b.Source, 'Sohar Date': b['Sohar Date'],
      };
    });
    const pctOf = (v) => (v === '' || v == null ? -1 : v);
    const day = (s) => { const d = parseDate(s); return d ? d.getTime() : 8.64e15; };
    rows.sort((a, b) => FLAG_RANK[a.Flag] - FLAG_RANK[b.Flag] || pctOf(b['OOP % Met']) - pctOf(a['OOP % Met']) || day(a['Next Appt']) - day(b['Next Appt']));
    return { rows, pending: tr.pending };
  }

  /** Counts for the pane: verified with benefits, OOP met / approaching, deductible met / approaching, and the OOP watch list. */
  function benefitsSummary(view) {
    const rows = view.rows || [];
    const n = (col, v) => rows.filter((r) => r[col] === v).length;
    return {
      verified: rows.length,
      withAmounts: rows.filter((r) => r['OOP Max'] !== '' || r.Deductible !== '').length,
      oopMet: n('OOP Status', 'Met'), oopApproaching: n('OOP Status', 'Approaching'),
      dedMet: n('Deductible Status', 'Met'), dedApproaching: n('Deductible Status', 'Approaching'),
      watch: rows.filter((r) => r['OOP Status'] === 'Met' || r['OOP Status'] === 'Approaching'),
      pending: (view.pending || []).length,
    };
  }

  function summarize(patients) {
    const s = { patients: patients.length, inScope: 0, verification: 0, discovery: 0, exclude: 0, highIssues: 0 };
    for (const p of patients) {
      if (p['In Scope'] !== 'Yes') continue;
      s.inScope++;
      s[p.Route] = (s[p.Route] || 0) + 1;
    }
    return s;
  }

  return {
    VERIFICATION_COLS, DISCOVERY_COLS, PATIENT_COLS, REVIEW_COLS, RESULTS_COLS, UPLOAD_COLS, BENEFITS_EXPORT_COLS,
    localDate, fileName, FILE_LABELS, BENEFIT_SHEET_COLS, BENEFIT_PCT_COLS, benefitsView, benefitsSummary,
    resultKey, accumulateResults, keepMemberIds, TRACKER_EXPORT_COLS, serviceLine, trackerExport, TASK_COLS, TASK_TYPES, IR_COLS, HEADER_SYNONYMS,
    irCategory, parseIrPortal, irCrossReference, irMessage,
    normHeader, detectHeader, parseDate, fmtDate, fmtTime, timeMins, soharTime, excelSerialToDate, parseName,
    payerMapIndex, lookupPayer, suggestPayers, buildPatients, buildExports, toCSV, parseCSV, summarize,
    VERIFIABLE, autoMatchPayer, autoTokens, AUTO_MEDIUM, finClassPayerIndex, matchVerifyRow, isVip,
    resultKind, readResultRows, mergeSoharResults, benefitFields, accumulatorStatus, benefitFlag, benefitsExport, buildTasks, taskMessage, keepDone, chunkMessage, messageAssembler, workQueue, lasfir, adminMessage,
  };
});
