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
  // v0.12: the patient column is the MRN (the add-in renames an older template's "Per Nbr" header in place).
  const REVIEW_COLS = ['Severity', 'Issue', 'MRN', 'Detail', 'Fix'];
  // Sohar Results sheet. The add-in creates this table itself, so it is not part of the template contract.
  const RESULTS_COLS = ['Next Appt', 'Next Time', 'Next Event', 'Next Location', 'Next Provider', 'Upcoming Appts', 'Upcoming Detail',
    'Per Nbr', 'MRN', 'Last Name', 'First Name', 'DOB', 'Sent As', 'Sohar Status', 'Work Queue', 'Action', 'Choose in NextGen', 'Benefit Flag',
    'Plan', 'Network', 'Coverage Level', 'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance',
    'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status',
    'NextGen Payer', 'NextGen Member ID', 'Sohar Payer', 'Sohar payerId', 'Sohar Member ID', 'Coverage Differs', 'Match',
    'Sohar Ref', 'Sohar Link', 'Sohar Date', 'Discovery Follow-up', 'Searched As'];
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
    guarDob: ['Guar Birth Dt', 'Guar DOB', 'Guar Dob', 'Guarantor DOB', 'Guarantor Birth Date', 'Guarantor Birth Dt', 'Guar Birth Date'],
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

  // ---------- patient key (v0.12) ----------
  // The patient key is the MRN (Md Rc). It is the Sohar id and what every saved table keys on. A patient with no MRN,
  // or one whose MRN another Per Nbr also has, uses "PER-<Per Nbr>" (flagged on Review). Rows carry the key as a
  // non-column `key` (or x.key); else the MRN column; else the Per Nbr column as a PER- key (rows saved before v0.12).
  const PER_PREFIX = 'PER-';
  const perKey = (per) => (str(per) ? `${PER_PREFIX}${str(per)}` : '');
  function keyOf(r) {
    if (!r) return '';
    const k = str(r.key) || str(r.x && r.x.key);
    if (k) return k;
    return str(r.MRN) || perKey(r['Per Nbr']);
  }
  /** A saved table's patient column: MRN (v0.12) or, on rows saved earlier, Per Nbr read as a PER- key. */
  const savedKey = (r) => (r ? str(r.MRN) || perKey(r['Per Nbr']) : '');
  /**
   * Convert a saved PER- key to the patient's key today (perToKey: Map(Per Nbr -> key) from buildPatients). Anything
   * else comes back as it is, so converting twice changes nothing.
   */
  function migrateKey(k, perToKey) {
    const s = str(k);
    if (!s.startsWith(PER_PREFIX) || !perToKey) return s;
    return str(perToKey.get(s.slice(PER_PREFIX.length))) || s;
  }
  /** Per Nbr -> key for every patient buildPatients returned. */
  function perKeyMap(patients) {
    const m = new Map();
    for (const p of patients || []) if (str(p['Per Nbr'])) m.set(str(p['Per Nbr']), keyOf(p));
    return m;
  }
  // Sohar ids with a tag: "<key>-G" guarantor discovery, "<key>-M" payer mapping discovery (member ID sent),
  // "<key>-T1".."-T9" payer test verifications. The tag says what the row was; the key ties it to the patient.
  const ID_TAGS = { G: 'Guarantor', M: 'Payer mapping', T: 'Payer test' };
  function splitId(id) {
    const s = str(id);
    const m = /^(.+)-(G|M|T\d)$/.exec(s);
    if (!m) return { base: s, tag: '', test: 0 };
    return { base: m[1], tag: m[2][0], test: m[2][0] === 'T' ? +m[2].slice(1) : 0 };
  }
  const searchedAs = (tag) => ID_TAGS[tag] || 'Patient';
  /** "MRN 900" (or "Per 100" for a patient with no usable MRN) for messages and the pane. */
  function idLabel(r) { const k = keyOf(r); return !k ? '' : k.startsWith(PER_PREFIX) ? `Per ${k.slice(PER_PREFIX.length)}` : `MRN ${k}`; }

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
  // 'Learned' = an automatic match Sohar agreed with on enough patients (learnPayers). It verifies like Auto-High.
  const VERIFIABLE = new Set(['exact', 'probable', 'confirmed', 'auto-high', 'auto-medium', 'learned']);
  // Confident enough to stand in for a Fin Class when deriving the last-encounter payer (Auto-Medium is not).
  const CONFIDENT = new Set(['exact', 'probable', 'confirmed', 'auto-high', 'learned']);
  // Labels the add-in writes itself. Every other value (Exact, Probable, Confirmed, Not a payer, Review, anything typed)
  // is a human label and is never changed automatically.
  const AUTO_LABELS = new Set(['auto-high', 'auto-medium', 'auto-low', 'learned']);
  const KNOWN_LABELS = new Set([...VERIFIABLE, 'auto-low', 'not a payer', 'review']);
  // Human labels that also teach the matcher (their names become aliases of their payerId).
  const ALIAS_LABELS = new Set(['exact', 'probable', 'confirmed', 'learned']);
  const isAutoLabel = (c) => AUTO_LABELS.has(str(c).toLowerCase());
  /** True when a Payer Map row needs a person: Auto-Low, Review, an unrecognized label (blank included), or a verifiable label with no payerId. */
  function needsClassification(confidence, payerId) {
    const c = str(confidence).toLowerCase();
    if (c === 'not a payer') return false;
    if (c === 'auto-low' || c === 'review' || !KNOWN_LABELS.has(c)) return true;
    return VERIFIABLE.has(c) && !str(payerId);
  }
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
   * Payer Map names a person (or Sohar learning) settled on: Exact, Probable, Confirmed, Learned rows with a payerId.
   * Returns [{name, payerId, payerName}] for autoMatchPayer's opts.confirmed.
   */
  function confirmedNames(mapRows) {
    const out = [];
    for (const r of mapRows instanceof Map ? [...mapRows].map(([, v]) => v) : mapRows || []) {
      const name = str(r.nextgenPayerName || r.name), id = str(r.payerId);
      if (name && id && ALIAS_LABELS.has(str(r.confidence).toLowerCase())) out.push({ name, payerId: id, payerName: str(r.payerName) });
    }
    return out;
  }
  // Sohar payers plus confirmed Payer Map names as extra 'confirmed' variants of their payerId. Cached per (list, confirmed).
  const withConfirmedCache = new WeakMap();
  function autoPrepWith(soharPayers, confirmed) {
    const base = autoPrep(soharPayers);
    if (!confirmed || !confirmed.length) return base;
    const hit = withConfirmedCache.get(confirmed);
    if (hit && hit.base === base) return hit.prepared;
    const byId = new Map();
    const prepared = base.map((p) => { const c = { ...p, variants: p.variants.slice() }; if (c.payerId) byId.set(c.payerId, c); return c; });
    for (const n of confirmed) {
      const toks = autoTokens(n.name);
      if (!toks.length || !str(n.payerId)) continue;
      let p = byId.get(str(n.payerId));
      if (!p) { p = { payerId: str(n.payerId), payerName: str(n.payerName) || str(n.name), variants: [] }; byId.set(p.payerId, p); prepared.push(p); }
      if (!p.variants.some((v) => v.toks.join(' ') === toks.join(' '))) p.variants.push({ text: str(n.name), kind: 'confirmed', toks });
    }
    withConfirmedCache.set(confirmed, { base, prepared });
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
   * opts.confirmed ([{name, payerId, payerName}], see confirmedNames): Payer Map names labeled Exact, Probable, Confirmed
   * or Learned count as aliases of their payerId, for the exact match ('Same as confirmed name "X"') and as fuzzy candidates.
   */
  function autoMatchPayer(name, soharPayers, opts) {
    const q = autoTokens(name);
    const none = (evidence) => ({ payerId: '', payerName: '', confidence: 'Auto-Low', score: 0, evidence });
    if (!q.length) return none('No words to match');
    const prepared = autoPrepWith(soharPayers, opts && opts.confirmed);
    if (!prepared.length) return none('Sohar payer list is empty');
    const key = q.join(' ');
    const exact = [];
    for (const p of prepared) {
      const v = p.variants.find((x) => x.toks.join(' ') === key);
      if (v && !exact.some((e) => e.p.payerId === p.payerId)) exact.push({ p, v });
    }
    if (exact.length === 1) {
      const { p, v } = exact[0];
      return { payerId: p.payerId, payerName: p.payerName, confidence: 'Auto-High', score: 1,
        evidence: v.kind === 'confirmed' ? `Same as confirmed name "${v.text}"` : `Exact ${v.kind} match: "${v.text}"` };
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
    // Learned from Sohar outcomes (learnFinClasses): beats this report's guess; Config overrides still win.
    for (const l of (opts && opts.learned) || []) {
      if (!str(l.finClass) || !str(l.payerId)) continue;
      out.set(fk(l.finClass), { finClass: str(l.finClass), nextgenPayer: str(l.nextgenPayer), payerId: str(l.payerId), payerName: str(l.payerName), count: +l.patients || 0, total: 0, source: 'learned' });
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
   *    finClassPayers: {finClass: payerId} | [{finClass, payerId, payerName}] (optional Config overrides),
   *    learnedFinClasses: [{finClass, payerId, payerName, nextgenPayer}] (optional, from learnFinClasses; Config overrides win) }
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
        guarDob: parseDate(get(row, 'guarDob')),
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
      const confirmed = confirmedNames(mapRows);
      const names = new Map();
      for (const t of pts) for (const n of [t.payerName, t.secName]) if (n) names.set(n.toLowerCase(), { name: n, n: (names.get(n.toLowerCase()) || { n: 0 }).n + 1 });
      for (const { name, n } of names.values()) {
        if (known.has(name.toLowerCase()) || isSelfPay(name) || isVip(name)) continue;
        const m = autoMatchPayer(name, cfg.soharPayers, { confirmed });
        const entry = { nextgenPayerName: name, payerId: m.payerId, payerName: m.payerName, confidence: m.confidence,
          notes: `Auto-matched by the add-in: ${m.evidence}`, score: m.score, patients: n };
        autoMapped.push(entry);
        mapRows.push(entry);
      }
      autoMapped.sort((a, b) => b.patients - a.patients || a.nextgenPayerName.localeCompare(b.nextgenPayerName));
    }
    const idx = payerMapIndex(mapRows);
    const finIdx = finClassPayerIndex(pts.map((t) => ({ finClass: t.finClass, payer: t.payerName })), idx, { overrides: cfg.finClassPayers, learned: cfg.learnedFinClasses });

    const patients = [];
    const review = [];
    // Review rows name the patient by key (MRN), filled in once every patient's key is known (below).
    const addReview = (sev, issue, per, detail, fix) => review.push({ Severity: sev, Issue: issue, MRN: '', _per: per, Detail: detail, Fix: fix });

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
      // Auto-Low, Review, or a label the add-in does not know: a person has to classify the name.
      const needsClass = lk.status === 'review' && needsClassification(lk.confidence, lk.payerId);
      const classifyReview = (fix) => addReview('High', 'Payer needs manual classification', per, `${payerName}${lk.payerName ? ` -> best guess ${lk.payerName}` : ''}${autoLow ? '' : ` (confidence: ${lk.confidence || 'blank'})`}`, fix);
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
          confidence: f.source === 'config' ? 'Fin Class (config)' : f.source === 'learned' ? 'Fin Class (learned)' : 'Fin Class (report)' };
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
          // One verification per insurance: the last encounter naming the same payer and member ID is the same policy.
          if (!verifyRows.some((v) => v.payerId === le.payerId && idKey(v.memberId) === idKey(le.memberId))) verifyRows.push(le);
          route = 'verification';
          reason = verifyRows.length > 1 ? `two insurances: chart ${payerName} (${lk.confidence}) and last encounter ${finClassLst}` : `chart payer ${lk.confidence}: ${payerName}; verifying last-encounter insurance (${finClassLst})`;
          if (verifyRows.length > 1) flags.push('Two insurances sent');
          addReview('Low', verifyRows.length > 1 ? 'Two insurances sent to verification' : 'Last-encounter insurance sent to verification', per,
            `Chart: ${payerName}${lk.payerName ? ` (best guess ${lk.payerName}, ${lk.confidence})` : ` (${lk.confidence})`}; last encounter: ${finClassLst} -> ${le.payerName}`,
            'The Sohar results decide; Choose in NextGen shows which insurance to keep.');
          if (needsClass) classifyReview('Classify it in the pane (Admin, step 1) or on the Payer Map sheet.');
        } else if (lk.status === 'unmapped') {
          route = 'discovery'; reason = `payer not in Payer Map: ${payerName}`;
          addReview('High', 'Unmapped payer', per, payerName, 'Map it in step 1 or add a row on the Payer Map sheet.');
        } else {
          route = 'discovery'; reason = `payer needs review (${lk.confidence}): ${payerName}`;
          if (needsClass) classifyReview('Sent to discovery. Classify it in the pane (Admin, step 1) or on the Payer Map sheet so it verifies next time.');
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

      // Guarantor: a Guar Name on the chart that is not the patient, with Guar Birth Dt when the report has that column.
      const guarRaw = latest.guar || pick('guar');
      const gName = parseName(guarRaw);
      const gDob = recs.map((x) => x.guarDob).find(Boolean) || null;
      const guarantor = guarRaw && gName.lastName && nameKey(gName.firstName, gName.lastName) !== nameKey(nm.firstName, nm.lastName)
        ? { name: str(guarRaw), firstName: gName.firstName, lastName: gName.lastName, dob: fmtDate(gDob) } : null;
      let subF = '', subL = '', subD = '';
      if (cfg.useGuarantorAsSubscriber && guarantor) { subF = guarantor.firstName; subL = guarantor.lastName; subD = guarantor.dob; }

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
          verifyRows, guarantor, apptCount: future.length,
        },
      });
    }
    // Patient key: the MRN; PER-<Per Nbr> when the MRN is blank or another Per Nbr has the same MRN (both flagged).
    const mrnPers = new Map();
    for (const p of patients) { const m = idKey(p.MRN); if (m) mrnPers.set(m, (mrnPers.get(m) || 0) + 1); }
    const perToKey = new Map();
    for (const p of patients) {
      const m = str(p.MRN), shared = m && mrnPers.get(idKey(m)) > 1;
      p.x.key = m && !shared ? m : perKey(p['Per Nbr']);
      perToKey.set(p['Per Nbr'], p.x.key);
      if (shared) addReview('High', 'MRN shared by two patients', p['Per Nbr'], `MRN ${m} is on more than one Per Nbr in the report`, `Fix the MRN in NextGen. Until then Sohar gets ${p.x.key} for this patient.`);
      else if (!m && p['In Scope'] === 'Yes') addReview('Medium', 'MRN missing', p['Per Nbr'], `No MRN (Md Rc) in the report for Per Nbr ${p['Per Nbr']}`, `Sohar gets ${p.x.key} as the id. Add Md Rc to the NextGen report.`);
    }
    for (const r of review) { r.MRN = str(r._per) ? perToKey.get(r._per) || perKey(r._per) : ''; delete r._per; }
    patients.sort((a, b) => (a.Route > b.Route ? -1 : a.Route < b.Route ? 1 : 0) || ((b['In Scope'] === 'Yes') - (a['In Scope'] === 'Yes')) || a['Per Nbr'].localeCompare(b['Per Nbr']));

    const unmapped = {};
    for (const p of patients) {
      if (p.Route === 'discovery' && p['Route Reason'].startsWith('payer not in Payer Map')) unmapped[p['Primary Payer']] = (unmapped[p['Primary Payer']] || 0) + 1;
    }
    return { detection: det, dataRows, patients, review, unmappedPayers: unmapped, autoMapped, finClassIndex: finIdx, asOf, perToKey };
  }

  // Payer mapping discovery (v0.12): the Sohar discovery template plus the member ID at the end, so Sohar can find the
  // payer of the policy NextGen has. Not a Sohar template: Sohar accepting the extra column is unconfirmed.
  const MAPPING_DISCOVERY_COLS = [...DISCOVERY_COLS, 'memberId'];

  /**
   * A discovery patient whose NextGen payer has no confident Sohar payer (unmapped, Auto-Low, Review or an unknown
   * label) but whose chart has a member ID: the payer mapping discovery sends that member ID instead of a plain discovery.
   */
  function isMappingCandidate(p) {
    if (!p || p.Route !== 'discovery' || p['In Scope'] !== 'Yes') return false;
    const payer = str(p['Primary Payer']);
    if (!payer || !str(p['Member ID']) || isSelfPay(payer) || isVip(payer) || (p.x && p.x.vip)) return false;
    return needsClassification(p.Confidence, p['Sohar payerId']);
  }

  /**
   * Sohar import rows (in-scope patients). id is the patient key (MRN).
   *  verification: one row per insurance (verifyRows entry); rows marked reverifySkip (verified eligible recently) left out.
   *  discovery: one row per discovery patient, plus one guarantor row (id <key>-G: the guarantor's name and DOB, the
   *    patient's state and ZIP) for each patient in cfg.guarantorRequests (Set/Map of keys) whose guarantor DOB is known.
   *  mapping: discovery patients isMappingCandidate picks, with their member ID (id <key>-M), when cfg.mappingDiscovery
   *    is not false. They are on this file instead of the discovery file: still one discovery each.
   *  guarantorNoDob: requested guarantor rows left out because the report has no guarantor DOB.
   */
  function buildExports(patients, cfg) {
    const c = cfg || {};
    const specialty = c.specialty || 'PSY';
    const inScope = patients.filter((p) => p['In Scope'] === 'Yes');
    const verification = [];
    for (const p of inScope.filter((x) => x.Route === 'verification')) {
      const vrs = p.x && p.x.verifyRows && p.x.verifyRows.length ? p.x.verifyRows : [{ payerId: p['Sohar payerId'], memberId: p['Member ID'] }];
      for (const v of vrs) {
        if (v.reverifySkip) continue;
        verification.push({
          firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, memberId: v.memberId,
          payerId: v.payerId, placeOfServiceCode: p.POS, specialtyCode: specialty, id: keyOf(p),
          subscriberFirstName: p['Subscriber First'], subscriberLastName: p['Subscriber Last'], subscriberDateOfBirth: p['Subscriber DOB'],
        });
      }
    }
    const disc = (p, id) => ({ firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, zipCode: p.Zip, payerId: p['Sohar payerId'], id });
    const asked = c.guarantorRequests instanceof Map || c.guarantorRequests instanceof Set ? c.guarantorRequests : new Set(c.guarantorRequests || []);
    const reasonOf = (k) => (asked instanceof Map ? str((asked.get(k) || {}).Reason || asked.get(k)) : '');
    const discovery = [], mapping = [], guarantorNoDob = [];
    const gRows = [];
    for (const p of inScope) {
      const g = p.x && p.x.guarantor;
      if (!g || !asked.has(keyOf(p))) continue;
      if (!g.dob) { guarantorNoDob.push(p); continue; }
      gRows.push({ p, row: { firstName: g.firstName, lastName: g.lastName, dateOfBirth: g.dob, state: p.State, zipCode: p.Zip, payerId: p['Sohar payerId'], id: `${keyOf(p)}-G` } });
    }
    // A guarantor resubmission after the patient's own discovery found nothing replaces the patient row (no repeat).
    const replaced = new Set(gRows.filter(({ p }) => reasonOf(keyOf(p)) === GUARANTOR_REASONS.none).map(({ p }) => keyOf(p)));
    for (const p of inScope.filter((x) => x.Route === 'discovery')) {
      if (replaced.has(keyOf(p))) continue;
      if (c.mappingDiscovery !== false && isMappingCandidate(p)) mapping.push({ ...disc(p, `${keyOf(p)}-M`), memberId: p['Member ID'] });
      else discovery.push(disc(p, keyOf(p)));
    }
    for (const { row } of gRows) discovery.push(row);
    return { verification, discovery, mapping, guarantorNoDob };
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

  /**
   * Out-of-network benefits from a Sohar verification record (normalized keys oonCopay, oonCoinsurance, oonTotalDeductible,
   * oonRemainingDeductible, oonOutOfPocket, oonOutOfPocketRemaining) and whether Sohar says the plan is out of network
   * (networkType OON / out of network, networkStatus out of network, or inNetwork false). Used only by fee estimates (v0.11).
   */
  function oonBenefitFields(rec) {
    const g = (k) => (rec && k in rec && !isBlank(rec[k]) ? str(rec[k]) : '');
    const net = `${g('networktype')} ${g('networkstatus')}`;
    return {
      'Copay': g('ooncopay'), 'Coinsurance': fmtPct(g('ooncoinsurance')),
      'Deductible': g('oontotaldeductible'), 'Deductible Remaining': g('oonremainingdeductible'),
      'OOP Max': g('oonoutofpocket'), 'OOP Remaining': g('oonoutofpocketremaining'),
      outOfNetwork: /\b(oon|out[\s-]*of[\s-]*network|non[\s-]*par(ticipating)?)\b/i.test(net) || /^(false|no|n)$/i.test(g('innetwork') || g('isinnetwork')),
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
    if (row.Match === 'No match' || row.Match === 'No identifiers') return 'Unmatched';
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
   * Plain-text note for admins, grouped by location then appointment. LASFIR + MRN only (minimum necessary).
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
        const id = keyOf(r) ? `${tag} · ${idLabel(r)}` : `${tag ? tag + ' · ' : ''}Sohar ref ${r['Sohar Ref'] || '?'}`;
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

  /** Verification subscriber columns: the guarantor's when the policy was found through the guarantor, else the patient row's. */
  function subscriberFor(p, viaGuarantor) {
    const g = p && p.x && p.x.guarantor;
    if (viaGuarantor && g) return { subscriberFirstName: g.firstName, subscriberLastName: g.lastName, subscriberDateOfBirth: g.dob };
    return { subscriberFirstName: p['Subscriber First'], subscriberLastName: p['Subscriber Last'], subscriberDateOfBirth: p['Subscriber DOB'] };
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
    // guessFor: the same from Auto-Low / Review rows (and unrecognized labels, which are treated like Review), used only
    // when no verifiable row names the payer.
    const nextgenFor = new Map(), guessFor = new Map();
    for (const m of cfg.payerMap || []) {
      const id = str(m.payerId), name = str(m.nextgenPayerName), c = str(m.confidence).toLowerCase();
      if (!id || !name) continue;
      const into = VERIFIABLE.has(c) ? nextgenFor : needsClassification(c, id) ? guessFor : null;
      if (!into) continue;
      if (!into.has(id)) into.set(id, []);
      if (!into.get(id).includes(name)) into.get(id).push(name);
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
    // byKey: patient key (MRN, else PER-<Per Nbr>), and PER-<Per Nbr> for everyone (saved rows from before v0.12).
    const byKey = new Map();
    for (const p of built.patients) {
      byPer.set(p['Per Nbr'], p);
      byKey.set(idKey(keyOf(p)), p);
      if (!byKey.has(idKey(perKey(p['Per Nbr'])))) byKey.set(idKey(perKey(p['Per Nbr'])), p);
      push(byMember, idKey(p['Member ID']), p);
      push(byMember, idKey(p['Member ID (Lst Enc)']), p);
      if (p['Last Name']) push(byName, nameKey(p['First Name'], p['Last Name']), p);
      if (p['Last Name'] && p.DOB) push(byNameDob, `${nameKey(p['First Name'], p['Last Name'])}|${p.DOB}`, p);
    }

    const results = [], unknownFiles = [], kinds = new Set(), seen = { verification: new Set(), discovery: new Set() };
    const counts = { matched: 0, unmatched: 0, noIdentifiers: 0, notReturned: 0, withBenefits: 0, byAction: {}, byQueue: {}, deductible: {}, oop: {} };
    // Sohar ref memory (cfg.refs: Sohar Refs rows). Refs learned here come back as refs (new rows for the Sohar Refs sheet).
    const refIdx = refIndex(cfg.refs || []);
    const learned = new Map();
    // Refs of blank-id records no remembered ref matched: if this pass learns one of them, refsPending says to run again.
    const unresolved = [];
    const freshKeys = cfg.freshKeys instanceof Set ? cfg.freshKeys : new Set(cfg.freshKeys || []);
    const stamps = new Map();
    // Sohar ids: the MRN since Config "MRN ids since" (cfg.mrnSince), the Per Nbr before. A check Sohar created before
    // that time is read as a Per Nbr; after it, as a patient key; with no time (or no setting), the key first.
    const since = cfg.mrnSince != null && cfg.mrnSince !== '' ? soharTime(cfg.mrnSince) : NaN;
    const idMatch = (r) => {
      const { base, tag, test } = splitId(r.id);
      if (!base) return { p: null, match: '', tag: '', test: 0 };
      const t = soharTime(r.created || r.resolved);
      const legacy = isNaN(since) || isNaN(t) ? null : t < since;
      let p = null, match = '';
      if (legacy !== true || tag) { p = byKey.get(idKey(base)) || null; if (p) match = 'MRN'; }
      if (!p && legacy !== false && !tag) { p = byPer.get(base) || null; if (p) match = 'Per Nbr (old id)'; }
      return { p, match, tag, test };
    };
    const byIdOrRef = (m) => m === 'MRN' || m === 'Per Nbr (old id)' || m === 'Sohar ref';
    // Payer test verifications (ids <key>-T<n>): kept apart from the patient's results (cfg.payerTestRows: Payer Tests rows;
    // cfg.payerTests is the Config settings).
    const testIdx = new Map((Array.isArray(cfg.payerTestRows) ? cfg.payerTestRows : []).map((t) => [str(t['Test Id']), t]));
    const payerTests = [];

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
        // Who Sohar searched: the patient, the guarantor (id <key>-G) or the payer mapping discovery (<key>-M).
        'Searched As': f.tag ? searchedAs(f.tag) : 'Patient',
        ...(f.benefits || {}), ...f.cols,
      };
      for (const c of RESULTS_COLS) if (row[c] == null) row[c] = '';
      row.key = p ? keyOf(p) : ''; // non-column: patient key (MRN), what tasks, flags and saved tables key on
      row.x = p ? p.x || null : null; // non-column: patient extras (appointments, fin class, verifyRows) for the rules
      if (f.key) row.resultKey = f.key; // non-column: resultKey of the Sohar record, for the payer outcomes ledger
      if (f.chooseGuess) row.chooseGuess = true; // non-column: Choose in NextGen came from an Auto-Low / Review guess
      if (f.sentPayerId) row.sentPayerId = f.sentPayerId; // non-column: Sohar payerId is the one sent, not today's chart
      if (f.oon) row.oon = f.oon; // non-column: out-of-network benefits (oonBenefitFields), for fee estimates only
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
        // A blank id with a remembered ref: the Per Nbr (and resultKey) of the record that ref was first matched on.
        const hit = !str(r.id) ? refLookup(refIdx, kind, r) : null;
        if (!str(r.id) && !hit) unresolved.push(refsOf(kind, r).map((ref) => `${kind}|${ref}`));
        const idm = idMatch(r);
        // A payer test answers "does this policy verify under this Sohar payer?": it goes to payerTests, not the
        // patient's results (no task, no benefits, no re-verify history).
        if (idm.tag === 'T') {
          const t = testIdx.get(str(r.id));
          payerTests.push({ 'Test Id': str(r.id), key: idm.p ? keyOf(idm.p) : splitId(r.id).base, test: t || null, matched: !!idm.p,
            'Sohar Status': str(r.status) + (str(r.statusreason) ? ` (${str(r.statusreason)})` : ''), 'Sohar Payer': str(r.payername),
            'Sohar Ref': str(r.verificationid || r.requestid), 'Sohar Link': str(r.url), 'Sohar Date': str(r.resolved || r.created), resultKey: resultKey(kind, r, refIdx) });
          continue;
        }
        const f0 = { key: resultKey(kind, r, refIdx), tag: idm.tag }; // non-column extras for this row (emit)
        let benefits = {};
        if (kind === 'verification') {
          // Exports echo our id (the MRN; the Per Nbr before v0.12); older ones only carry member ID and name.
          p = idm.p; match = idm.match;
          if (!p && hit) { p = byKey.get(idKey(hit.per)) || null; if (p) match = 'Sohar ref'; }
          if (!p) { p = pickOne(byMember.get(idKey(r.memberid)), last); if (p) match = 'Member ID'; }
          if (!p && last) { p = pickOne(byName.get(nameKey(first, last)), last); if (p) match = 'Name only. Confirm.'; }
          benefits = benefitFields(r, bopts);
          f0.oon = oonBenefitFields(r); // non-column: out-of-network benefits, read only by fee estimates (v0.11)
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
          // Sohar exports carry no payerId, so vr is today's chart. The payerId sent is recorded on the saved record
          // (_sentpayerid) when it is first uploaded; a record from this upload (cfg.freshKeys) gets one now (stamps).
          const sent = str(r.payerid || r._sentpayerid);
          if (sent) { cols['Sohar payerId'] = sent; f0.sentPayerId = sent; }
          else if (vr && str(vr.payerId) && freshKeys.has(f0.key) && (byIdOrRef(match) || match === 'Member ID')) {
            f0.sentPayerId = str(vr.payerId);
            stamps.set(f0.key, f0.sentPayerId);
          }
        } else {
          p = idm.p; match = idm.match;
          if (!p && hit) { p = byKey.get(idKey(hit.per)) || null; if (p) match = 'Sohar ref'; }
          if (!p && last) {
            p = pickOne(byNameDob.get(`${nameKey(first, last)}|${fmtDate(dob)}`), last);
            if (p) match = 'Name + DOB';
          }
          const found = !isBlank(r.responsepayerid) || !isBlank(r.responsepayername);
          const resolved = !isBlank(r.resolved);
          const status = found ? 'Coverage found' : resolved ? 'No coverage found' : 'Pending';
          // The patient's Sohar payer, else the Payer Map's best guess for their NextGen payer (Auto-Low sends none).
          const ownId = p ? p['Sohar payerId'] || (p.x && p.x.payerGuess ? str(p.x.payerGuess.payerId) : '') : '';
          const differs = found && p && (str(r.responsepayerid) !== ownId || (!isBlank(r.memberid) && idKey(r.memberid) !== idKey(p['Member ID'])));
          const getId = isBlank(r.memberid) ? ' Get the member ID from the Sohar link.' : '';
          const sure = found ? nextgenFor.get(str(r.responsepayerid)) || [] : [];
          // Only guesses (Auto-Low / Review) name this payer: name them, and ask for the mapping to be confirmed.
          const guessed = found && !sure.length ? guessFor.get(str(r.responsepayerid)) || [] : [];
          const options = sure.length ? sure : guessed;
          const confirmText = guessed.length ? ' The Payer Map only guesses this NextGen insurance; confirm the payer mapping.' : '';
          const current = p && options.some((o) => o.toLowerCase() === p['Primary Payer'].toLowerCase()) ? p['Primary Payer'] : '';
          let choose = '', pickText = '';
          if (found) {
            if (current) { choose = current; pickText = ` Keep "${current}" in NextGen.${confirmText}`; }
            else if (options.length === 1) { choose = options[0]; pickText = ` Choose "${options[0]}" in NextGen.${confirmText}`; }
            else if (options.length > 1) {
              choose = `One of: ${options.join(' | ')}`;
              pickText = ` Choose one of ${options.map((o) => `"${o}"`).join(', ')} in NextGen; match the plan type on the card.${confirmText}`;
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
          else action = `Coverage found matches NextGen. Verify.${confirmText}`;
          // Guarantor search (<key>-G): the policy is the guarantor's; the patient is a dependent on it.
          if (idm.tag === 'G') {
            const g = p && p.x && p.x.guarantor;
            action = found ? `Found through the guarantor${g ? ` (${g.lastName}, ${g.firstName})` : ''}: the patient is likely a dependent on this policy. ${action}`
              : resolved ? 'Guarantor search found no coverage either. Confirm self-pay or ask the patient for a card.' : 'Guarantor search pending at Sohar. Recheck later.';
          }
          if (guessed.length) f0.chooseGuess = true;
          Object.assign(cols, {
            'Sohar Status': status, 'Action': action, 'Sohar Payer': str(r.responsepayername), 'Sohar payerId': str(r.responsepayerid),
            'Sohar Member ID': str(r.memberid), 'Coverage Differs': found ? (differs ? 'Yes' : 'No') : '',
            'Sohar Ref': str(r.requestid), 'Sohar Link': str(r.url), 'Sohar Date': str(r.resolved || r.created),
          });
        }
        // No id, no remembered ref, and nothing else to match on: an export anomaly or a check run outside the batch.
        const noIds = !p && !str(r.id) && !hit && !first && !last && (kind === 'verification' ? !str(r.memberid) : !dob);
        if (noIds) {
          match = 'No identifiers'; counts.noIdentifiers++;
          cols.Action = `No patient identifiers in this Sohar row (no id, ${kind === 'verification' ? 'member ID or name' : 'name or DOB'}). Likely an export anomaly or a check run outside the batch: open the Sohar link to see who it is. ${cols.Action}`;
        } else if (!p) {
          match = 'No match'; counts.unmatched++;
          cols.Action = `Not matched to a patient in the Import report.${hit ? ` Sohar ref remembered for ${/^PER-/.test(hit.per) ? `Per Nbr ${hit.per.slice(4)}` : `MRN ${hit.per}`}, who is not in this report.` : ''} ${cols.Action}`;
        } else counts.matched++;
        cols.Match = match;
        // Learn this record's refs only from a match on the id (MRN, or an old Per Nbr id) or Member ID (never a name match).
        if (p && (match === 'MRN' || match === 'Per Nbr (old id)' || match === 'Member ID')) {
          for (const ref of refsOf(kind, r)) {
            const k = `${kind}|${ref}`;
            if (!refIdx.has(k) && !learned.has(k)) learned.set(k, { 'Sohar Ref': ref, 'MRN': keyOf(p), 'Kind': kind, 'Result Key': f0.key });
          }
        }
        if (benefits['Deductible'] || benefits['OOP Max'] || benefits['Copay']) counts.withBenefits++;
        emit(p, kind, { firstName: first, lastName: last, dob: fmtDate(dob), cols, benefits, ...f0 });
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
    const discoveryVerification = [], dvSeen = new Set();
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
        // Patient and guarantor searches finding the same payer: one verification (the first found wins).
        const dvKey = `${keyOf(p)}|${d['Sohar payerId']}`;
        if (dvSeen.has(dvKey)) continue;
        dvSeen.add(dvKey);
        discoveryVerification.push({
          firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State,
          // Member ID from discovery; NextGen's only when NextGen already has this payer and Sohar did not reject it.
          memberId: d['Sohar Member ID'] || (p['Sohar payerId'] === d['Sohar payerId'] && !idRejected(d['Per Nbr'], d['Sohar payerId']) ? p['Member ID'] : ''),
          payerId: d['Sohar payerId'], placeOfServiceCode: p.POS, specialtyCode: specialty, id: keyOf(p),
          ...subscriberFor(p, d['Searched As'] === 'Guarantor'),
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
        payerId: hint ? hint['Sohar payerId'] : '', id: keyOf(p),
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
    // refsPending: blank-id records whose ref was only learned in this pass (same upload, or an original saved before the
    // Sohar Refs sheet existed). Run accumulateResults + mergeSoharResults again with refs + these refs before saving.
    const refsPending = unresolved.filter((ks) => ks.some((k) => learned.has(k))).length;
    return { results, unknownFiles, counts, asOf, discoveryVerification, verificationDiscovery, patients: built.patients, perToKey: built.perToKey, payerTests, refs: [...learned.values()], refsPending,
      sentPayerIds: [...stamps].map(([key, payerId]) => ({ key, payerId })) };
  }

  // ---------- Sohar ref memory ----------
  // Sohar exports can arrive with id, memberId, names and DOB blank. Each Sohar check has refs (requestId, the id in its
  // app.soharhealth.com/verifications/<id> link, verificationId on old exports). A ref is remembered with the patient key (MRN) and
  // the resultKey of the record it came on, once that record matched by its id or Member ID (hidden Sohar Refs sheet,
  // append-only; "Clear saved Sohar results" keeps it). A later record with a blank id is then the same patient and the
  // same result.
  const SOHAR_REF_COLS = ['Sohar Ref', 'MRN', 'Kind', 'Result Key', 'First Matched'];
  const SOHAR_URL_ID = /\/verifications\/([^/?#\s]+)/i;
  /** The refs of a Sohar record (normalized keys): requestid, the verification id in its url, verificationid. */
  function refsOf(kind, r) {
    const m = SOHAR_URL_ID.exec(str(r && r.url));
    const out = [];
    for (const v of [r && r.requestid, m ? m[1] : '', r && r.verificationid]) { const s = str(v); if (s && !out.includes(s)) out.push(s); }
    return out;
  }
  const refIndexCache = new WeakMap();
  /** Sohar Refs rows (or a Map from here) -> Map("kind|ref" -> {per, key}). The first row for a ref wins. */
  function refIndex(rows) {
    if (rows instanceof Map) return rows;
    const list = rows || [];
    if (refIndexCache.has(list)) return refIndexCache.get(list);
    const idx = new Map();
    for (const r of list) {
      const ref = str(r['Sohar Ref']), kind = str(r.Kind).toLowerCase(), per = savedKey(r);
      const k = `${kind}|${ref}`;
      if (ref && kind && per && !idx.has(k)) idx.set(k, { per, key: str(r['Result Key']) });
    }
    if (rows) refIndexCache.set(list, idx);
    return idx;
  }
  /** The remembered {per, key} for a record's refs, else null. */
  function refLookup(refs, kind, r) {
    if (!refs) return null;
    const idx = refIndex(refs);
    for (const ref of refsOf(kind, r)) { const hit = idx.get(`${kind}|${ref}`); if (hit) return hit; }
    return null;
  }
  /** Append-only: previous rows stay; fresh refs are added unless Kind + Sohar Ref is already there. */
  // perToKey (optional, perKeyMap): saved PER- keys become the patient's MRN, so a row saved before v0.12 converts on the next save.
  function appendRefs(previous, fresh, when, perToKey) {
    const seen = new Set(), rows = [];
    const norm = (r) => Object.fromEntries(SOHAR_REF_COLS.map((c) => [c, r[c] == null ? '' : str(r[c])]));
    const k = (r) => `${str(r.Kind).toLowerCase()}|${str(r['Sohar Ref'])}`;
    for (const r of previous || []) { if (!str(r['Sohar Ref']) || seen.has(k(r))) continue; seen.add(k(r)); rows.push(norm({ ...r, MRN: migrateKey(savedKey(r), perToKey) })); }
    let added = 0;
    for (const r of fresh || []) {
      if (!str(r['Sohar Ref']) || !savedKey(r) || seen.has(k(r))) continue;
      seen.add(k(r)); rows.push(norm({ ...r, 'First Matched': str(r['First Matched']) || str(when) })); added++;
    }
    return { rows, added };
  }

  /**
   * Stable key for one Sohar result, so a re-upload replaces it and a new file adds to the saved set.
   * refs (Sohar Refs rows or refIndex): a record with a blank id whose ref is remembered gets the key the original
   * id-bearing record had, so a re-export with blank identifiers never duplicates a saved result or a ledger row.
   */
  function resultKey(kind, r, refs) {
    if (refs && !str(r.id)) { const hit = refLookup(refs, kind, r); if (hit && hit.key) return hit.key; }
    if (kind === 'discovery') return `d|${str(r.id) || str(r.requestid) || `${nameKey(r.firstname, r.lastname)}|${str(r.dateofbirth)}`}`;
    const who = str(r.id) || str(r.verificationid || r.requestid) || `${idKey(r.memberid)}|${nameKey(r.firstname, r.lastname)}`;
    // Member ID too: two policies with the same payer (R3) are different results.
    return `v|${who}|${str(r.payername).toLowerCase()}|${idKey(r.memberid)}`;
  }

  /**
   * Add newly uploaded Sohar files to the saved result records. Same key: the new upload wins.
   * saved: rows of UPLOAD_COLS. refs: Sohar Refs rows (optional, see resultKey).
   * When a record replaces one from the same Sohar check (a shared ref), identifiers the new one has blank (id, member ID,
   * names, DOB) are kept from the earlier one. Returns {saved, files: [{name, kind, records}], added, replaced, unknownFiles,
   * fresh (Set of the keys this upload added or replaced)}.
   */
  // _sentpayerid (the payerId sent, stamped by stampSentPayerIds) belongs to the Sohar check, so it is kept the same way.
  const IDENTITY_FIELDS = ['id', 'memberid', 'firstname', 'lastname', 'dateofbirth', 'subscriberfirstname', 'subscriberlastname', 'subscriberdateofbirth', '_sentpayerid'];
  function keepIdentity(kind, rec, prev) {
    if (!prev) return rec;
    const mine = refsOf(kind, rec);
    if (!mine.length || !refsOf(kind, prev).some((x) => mine.includes(x))) return rec;
    const out = { ...rec };
    for (const f of IDENTITY_FIELDS) if (isBlank(out[f]) && !isBlank(prev[f])) out[f] = prev[f];
    return out;
  }
  function accumulateResults(saved, uploads, uploadedAt, refs) {
    const idx = refs ? refIndex(refs) : null;
    const by = new Map();
    for (const s of saved || []) {
      let rec; try { rec = JSON.parse(s.Record); } catch (e) { continue; }
      const kind = str(s.Kind), key = kind ? resultKey(kind, rec, idx) : String(s.Key); // re-keyed so older saved keys migrate
      const prev = by.get(key);
      by.set(key, { Kind: kind, Key: key, File: str(s.File), Uploaded: str(s.Uploaded), rec: prev ? keepIdentity(kind, rec, prev.rec) : rec });
    }
    let added = 0, replaced = 0;
    const unknownFiles = [], fresh = new Set();
    for (const f of uploads || []) {
      const { kind, records } = readResultRows(f.rows);
      if (!kind) { unknownFiles.push(f.name); continue; }
      for (const rec of records) {
        const key = resultKey(kind, rec, idx);
        const prev = by.get(key);
        if (prev) replaced++; else added++;
        fresh.add(key);
        by.set(key, { Kind: kind, Key: key, File: f.name, Uploaded: uploadedAt || '', rec: prev ? keepIdentity(kind, rec, prev.rec) : rec });
      }
    }
    const files = ['verification', 'discovery'].map((kind) => ({ name: `saved ${kind} results`, kind, records: [...by.values()].filter((x) => x.Kind === kind).map((x) => x.rec) }))
      .filter((f) => f.records.length);
    const out = [...by.values()].map((x) => ({ Kind: x.Kind, Key: x.Key, File: x.File, Uploaded: x.Uploaded, Record: JSON.stringify(x.rec) }));
    return { saved: out, files, added, replaced, unknownFiles, fresh };
  }

  /**
   * Record the payerId sent on saved verification records (mergeSoharResults sentPayerIds: [{key, payerId}], for records
   * from this upload). Sohar exports have no payerId column, so without this a later merge could only use today's chart.
   * A record that already has one keeps it. Returns {saved, stamped}.
   */
  function stampSentPayerIds(saved, stamps) {
    const by = new Map((stamps || []).map((s) => [str(s.key), str(s.payerId)]));
    let stamped = 0;
    const out = (saved || []).map((s) => {
      const id = by.get(str(s.Key));
      if (!id || str(s.Kind) !== 'verification') return s;
      let rec; try { rec = JSON.parse(s.Record); } catch (e) { return s; }
      if (str(rec._sentpayerid)) return s;
      stamped++;
      return { ...s, Record: JSON.stringify({ ...rec, _sentpayerid: id }) };
    });
    return { saved: out, stamped };
  }

  /**
   * One results pass (the pane's processResults): add the uploads to the saved records, match them (mergeSoharResults),
   * and when a blank-id record's ref was only learned in this pass (res.refsPending: same upload, or an original saved
   * before the Sohar Refs sheet existed) add and match again with those refs, so it is neither a duplicate nor
   * "No identifiers". Records from this upload get the payerId sent recorded (stampSentPayerIds).
   * cfg: as for mergeSoharResults (refs and freshKeys are set here). Returns {acc, res (null when nothing is saved; may
   * hold res.error), refs (learned this pass, for appendRefs), saved (UPLOAD_COLS rows to write)}.
   */
  function combineResults(importRows, saved, uploads, uploadedAt, refs, cfg) {
    let acc = accumulateResults(saved, uploads, uploadedAt, refs || []);
    if (!acc.files.length) return { acc, res: null, refs: [], saved: acc.saved };
    let res = mergeSoharResults(importRows, acc.files, { ...cfg, refs: refs || [], freshKeys: acc.fresh });
    if (res.error) return { acc, res, refs: [], saved: acc.saved };
    let learned = res.refs;
    if (res.refsPending) {
      const more = appendRefs(refs || [], learned, '').rows;
      acc = accumulateResults(saved, uploads, uploadedAt, more);
      res = mergeSoharResults(importRows, acc.files, { ...cfg, refs: more, freshKeys: acc.fresh });
      if (res.error) return { acc, res, refs: [], saved: acc.saved };
      learned = learned.concat(res.refs);
    }
    return { acc, res, refs: learned, saved: stampSentPayerIds(acc.saved, res.sentPayerIds).saved };
  }

  /**
   * Keep member IDs an admin typed on the Discovery Verify sheet, matched on id (patient key) + payerId. perToKey
   * (optional): a row saved before v0.12 has the Per Nbr as its id; it is matched as that patient's key too.
   */
  function keepMemberIds(rows, previous, perToKey) {
    const typed = new Map();
    for (const p of previous || []) {
      if (!str(p.memberId)) continue;
      typed.set(`${str(p.id)}|${str(p.payerId)}`, str(p.memberId));
      const old = perToKey && perToKey.get(str(p.id));
      if (old && !typed.has(`${old}|${str(p.payerId)}`)) typed.set(`${old}|${str(p.payerId)}`, str(p.memberId));
    }
    return rows.map((r) => (r.memberId ? r : { ...r, memberId: typed.get(`${r.id}|${r.payerId}`) || '' }));
  }

  // ---------- NextGen task list ----------
  const TASK_TYPES = ['Add insurance in NextGen', 'Update insurance in NextGen', 'Fix member ID in NextGen', 'Fix name in NextGen',
    'Update benefits in NextGen', 'Add payer to Payer Map', 'Confirm payer mapping', 'Confirm self-pay or get card', 'Find patient for Sohar result', 'Resend to Sohar', 'Send for verification'];

  /**
   * One task per thing an admin has to do in NextGen (or the Payer Map) before the visit.
   * Every eligible verification with benefits gets an "Update benefits" task with the values to enter,
   * because the appointment report does not show what NextGen currently holds.
   */
  const taskKey = (key, ref, task, searched) => `${str(key) || 'ref:' + str(ref)}|${str(task)}${searched === 'Guarantor' ? '|G' : ''}`;
  function buildTasks(results) {
    const tasks = [];
    const add = (r, task, details) => {
      // Task Key: patient key (MRN) + task; a guarantor search's task is its own (|G).
      const key = taskKey(keyOf(r), r['Sohar Ref'], task, r['Searched As']);
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
      if (r.Match === 'No identifiers') continue; // one Review row for all of them (noIdentifierReviewRows), not a task each
      if (r.Match === 'No match') { add(r, 'Find patient for Sohar result', `Sohar ${r['Sent As'].toLowerCase()} result did not match the Import report. ${r.Action}`); continue; }
      if (status === 'not in sohar file') { add(r, 'Resend to Sohar', 'Upcoming patient missing from the Sohar results.'); continue; }
      if (r['Sent As'] === 'Discovery') {
        if (status !== 'coverage found') { if (status === 'no coverage found') add(r, 'Confirm self-pay or get card', 'Sohar found no coverage.'); continue; }
        const choose = r['Choose in NextGen'] || '';
        if (/^Not in Payer Map/.test(choose)) add(r, 'Add payer to Payer Map', `${r['Sohar Payer']} (${r['Sohar payerId']}) has no NextGen insurance in the Payer Map.`);
        else if (r.chooseGuess || /Payer Map only guesses/.test(r.Action || '')) add(r, 'Confirm payer mapping', `${r['Sohar Payer']} (${r['Sohar payerId']}) is only a guess on the Payer Map for ${choose.replace(/^One of: /, 'one of: ')}. Confirm it in the pane (Admin, step 1) or on the Payer Map sheet.`);
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

  /** Open tasks as a note for admins, by location then visit. LASFIR + MRN only. */
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
        const who = keyOf(t) ? `${tag} · ${idLabel(t)}` : `Sohar ref ${t['Sohar Ref'] || '?'}`;
        lines.push(`- ${when} · ${who} · ${t.Task}: ${t.Details}`);
      }
      lines.push('');
    }
    return lines.join('\n').trim() + '\n';
  }

  /** Carry Done marks over from the previous task table, matched on Task Key. */
  // A task row saved before v0.12 has a Per Nbr Task Key; its MRN column gives the key it has now.
  function keepDone(tasks, previous) {
    const done = new Map();
    for (const p of previous || []) {
      const d = String(p['Done'] || '').trim();
      if (!d) continue;
      done.set(String(p['Task Key']), d);
      const now = taskKey(keyOf(p), p['Sohar Ref'], p.Task, /\|G$/.test(String(p['Task Key'])) ? 'Guarantor' : '');
      if (!done.has(now)) done.set(now, d);
    }
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
      if (r.Match === 'No match' || r.Match === 'No identifiers') continue;
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
    mapping: 'Sohar Ready Import - Discovery (Payer Mapping)',
    payerTests: 'Sohar Ready Import - Verification (Payer Tests)',
    skipped: 'Sohar Skipped (verified recently)',
    clinicList: 'Sohar Clinic List',
    clinicFlags: 'Sohar Clinic Flags',
    metrics: 'Sohar Payer Metrics',
    themes: 'Sohar Guarantor Themes',
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

  // ---------- payer learning (v0.10) ----------
  // Every Sohar outcome is recorded on the hidden Payer Outcomes ledger (payerOutcomes + appendOutcomes). learnPayers rolls it
  // up per NextGen payer name; planLearning turns the rollup into Payer Map changes (promote, demote, re-score) for automatic
  // labels only. Every Payer Map change is logged (mapChangeLog, sheetEditLog). Rules approved by Brendan on 2026-10-20.
  const LEARN_DEFAULTS = { on: true, promotePatients: 3, demotePayerErrors: 2, demoteDiscoveryShare: 0.5, demoteMinPatients: 3, suggestPatients: 3, suggestShare: 0.8 };
  // Optional Config rows, added on build when missing so they can be seen and edited.
  const LEARNING_CONFIG_ROWS = [
    { setting: 'Learning', value: 'On', what: 'On or Off. On: Sohar outcomes promote, demote and re-score the add-in\'s own payer matches (Auto-High, Auto-Medium, Auto-Low, Learned) and teach Fin Class payers. Off: outcomes are still recorded. Human labels are never changed.' },
    { setting: 'Learn: promote patients', value: 3, what: 'Patients Sohar must agree on, with no conflicts, before an Auto-Low or Auto-Medium payer becomes Learned (and before a last-encounter Fin Class is learned).' },
    { setting: 'Learn: demote payer errors', value: 2, what: 'Patients with a Sohar payer error that drop an Auto-High, Auto-Medium or Learned payer to Auto-Low.' },
    { setting: 'Learn: demote discovery share', value: '50%', what: 'Share of discovery patients finding a different payer for the same member ID that drops an Auto-High, Auto-Medium or Learned payer to Auto-Low. Other coverage found (a different member ID) never counts.' },
    { setting: 'Learn: demote min patients', value: 3, what: 'Discovery patients needed before the discovery share can demote a payer.' },
    { setting: 'Learn: suggest patients', value: 3, what: 'Discovery patients returning the same other payer for the same member ID before the pane suggests it for an unmapped, Auto-Low or Review name.' },
    { setting: 'Learn: suggest share', value: '80%', what: 'Share of discovery patients that must return that payer for a suggestion.' },
  ];
  /** Config {Setting: Value} -> learning settings. Counts are whole numbers >= 1; shares accept 50%, 0.5 or 50. */
  function learningConfig(cfg) {
    const c = cfg || {};
    const raw = (k) => (c[k] == null ? '' : String(c[k]).trim());
    const count = (k, d) => { const n = parseInt(raw(k), 10); return isNaN(n) || n < 1 ? d : n; };
    const share = (k, d) => {
      const s = raw(k), n = parseFloat(s.replace('%', ''));
      if (isNaN(n) || n < 0) return d;
      return s.includes('%') || n > 1 ? Math.min(1, n / 100) : n;
    };
    return {
      on: !/^(off|no|n|false|0)$/i.test(raw('Learning')),
      promotePatients: count('Learn: promote patients', LEARN_DEFAULTS.promotePatients),
      demotePayerErrors: count('Learn: demote payer errors', LEARN_DEFAULTS.demotePayerErrors),
      demoteDiscoveryShare: share('Learn: demote discovery share', LEARN_DEFAULTS.demoteDiscoveryShare),
      demoteMinPatients: count('Learn: demote min patients', LEARN_DEFAULTS.demoteMinPatients),
      suggestPatients: count('Learn: suggest patients', LEARN_DEFAULTS.suggestPatients),
      suggestShare: share('Learn: suggest share', LEARN_DEFAULTS.suggestShare),
    };
  }

  /** New note first, earlier notes after " | ", at most max characters (default 500). Repeating the latest note adds nothing. */
  function mergeNotes(note, old, max) {
    const n = str(note), o = str(old), m = max || 500;
    const out = !n ? o : !o || o === n || o.startsWith(`${n} | `) ? (o || n) : `${n} | ${o}`;
    return out.length > m ? out.slice(0, m) : out;
  }

  // Hidden Payer Outcomes sheet: one row per Sohar outcome, append-only, keyed by the result + Sohar Date.
  // v0.12: MRN (the patient key; rows saved earlier have Per Nbr) and Member Match (discovery: Sohar returned the member ID
  // NextGen has, so the payer it found is that policy's payer; only then is a different payer evidence against the mapping).
  const PAYER_OUTCOME_COLS = ['Key', 'MRN', 'Kind', 'NextGen Payer', 'Fin Class', 'Source', 'Sent payerId', 'Sent Confidence', 'Sohar Status',
    'Returned payerId', 'Returned Payer', 'Signal', 'Sohar Date', 'Recorded', 'Member Match'];

  // An outage or slow payer says nothing about whether the mapping is right: such errors are neutral.
  const TRANSIENT_ERROR = /unavailable|time[\s-]?out|timed out|\bdown\b|retry|try again|not responding|no response|temporar|busy|maintenance|connection|throttl|rate limit/;
  // A non-member error counts against the payer only when its text says the payer was wrong or not found.
  const PAYER_WRONG = /not found|invalid|unknown|incorrect|wrong|mismatch|not supported|unsupported|no match/;
  /**
   * What a verification status (status + " (statusReason)") says about the payer that was sent: 'agree'
   * (complete.eligible), 'payer error' (error.payer*, or another non-member error saying the payer was wrong or not
   * found), 'member id error', 'ineligible', else 'neutral'. Transient errors (unavailable, timeout, not responding,
   * retry...) and error.payer.npi (the provider's NPI setup at the payer) are neutral. Only 'agree' and 'payer error' move learning. The real Sohar error.payer* strings are not
   * yet confirmed; check them against a Sohar output before relying on payer-error demotes.
   */
  function verificationSignal(status) {
    const s = str(status).toLowerCase();
    if (/^complete\.eligible/.test(s)) return 'agree';
    if (/^complete\.ineligible/.test(s)) return 'ineligible';
    if (/^error\.member\.id/.test(s)) return 'member id error';
    // error.payer.npi is the provider's NPI setup at the payer, not a wrong payer: neutral.
    if (/^error/.test(s) && !TRANSIENT_ERROR.test(s) && !/^error\.payer\.npi/.test(s)) {
      if (/^error\.payer/.test(s)) return 'payer error';
      if (!/^error\.member/.test(s) && /payer/.test(s) && PAYER_WRONG.test(s)) return 'payer error';
    }
    return 'neutral';
  }

  /**
   * Ledger rows from merged Sohar results (mergeSoharResults: results + patients). Unmatched and "Not in Sohar file" rows
   * are skipped. Verification: NextGen Payer is the insurance sent (the proxy name for a last-encounter row), Sent payerId
   * the payerId sent; a verification of a payer discovery found is Source 'discovery' and neutral. Discovery: the found
   * payer is compared with the patient's Sohar payerId (else the Payer Map guess), which is recorded as Sent payerId:
   * 'agree' (same payer), 'other payer' (another payer for the same member ID: the mapping looks wrong), 'found' (nothing
   * to compare with, same member ID), 'other coverage' (another payer and no member ID match: a different policy, not
   * evidence; e.g. Cigna found for a Blue Card patient), 'neutral' (nothing found). Member Match is Yes when Sohar returned
   * the member ID NextGen has, or, for the payer mapping discovery (member ID sent), returned none or the same one.
   * A guarantor search is about the guarantor's policy: Source 'guarantor', neutral.
   */
  function payerOutcomes(results, patients, recordedAt) {
    const byPer = new Map((patients || []).map((p) => [str(p['Per Nbr']), p]));
    const out = [];
    for (const r of results || []) {
      const per = str(r['Per Nbr']);
      // Name matches (and no-identifier rows) never stand for the patient: another person's result must not teach.
      if (!per || !r.resultKey || untrustedMatch(r) || r['Sohar Status'] === 'Not in Sohar file') continue;
      const p = byPer.get(per);
      if (!p) continue;
      const x = p.x || {};
      const date = str(r['Sohar Date']);
      const o = { 'Key': `${r.resultKey}|${date}`, 'MRN': keyOf(p), 'Sohar Status': str(r['Sohar Status']), 'Sohar Date': date, 'Recorded': str(recordedAt) };
      if (r['Sent As'] === 'Verification') {
        const fromDisc = r['Discovery Follow-up'] === 'From discovery';
        const vr = matchVerifyRow(p, { memberid: r['Sohar Member ID'], payername: r['Sohar Payer'] });
        // The payerId sent (recorded at upload) wins over today's chart; the chart row describes it only when they agree.
        const sentId = str(r.sentPayerId) || (vr ? str(vr.payerId) : str(r['Sohar payerId']));
        const vrSent = vr && str(vr.payerId) === sentId ? vr : null;
        const lastEnc = !!vrSent && vrSent.source === 'last encounter';
        Object.assign(o, {
          'Kind': 'Verification', 'NextGen Payer': str(r['NextGen Payer']),
          'Fin Class': str(lastEnc ? x.finClassLstEnc : x.finClass), 'Source': fromDisc ? 'discovery' : lastEnc ? 'last encounter' : 'chart',
          'Sent payerId': sentId, 'Sent Confidence': vrSent ? str(vrSent.confidence) : '',
          'Returned payerId': '', 'Returned Payer': str(r['Sohar Payer']),
          'Signal': fromDisc ? 'neutral' : verificationSignal(r['Sohar Status']),
        });
      } else {
        // VIP with only a secondary on file: discovery used the secondary's payer, so the outcome is about the secondary.
        const sec = !!x.vip && !!str(p['Secondary Payer']) && !!str(p['Sohar payerId']) && p['Sohar payerId'] === p['Secondary Sohar payerId'];
        const chartVr = (x.verifyRows || []).find((v) => v && v.source === 'guarantor');
        const guess = x.payerGuess ? str(x.payerGuess.payerId) : '';
        const compare = sec ? str(p['Sohar payerId']) : (chartVr ? str(chartVr.payerId) : p.Route === 'verification' ? '' : str(p['Sohar payerId'])) || guess;
        const ret = str(r['Sohar payerId']);
        const found = r['Sohar Status'] === 'Coverage found' && !!ret;
        const viaGuar = r['Searched As'] === 'Guarantor', mapping = r['Searched As'] === 'Payer mapping';
        const back = idKey(r['Sohar Member ID']), mine = idKey(p['Member ID']);
        const memberMatch = !sec && !viaGuar && found && !!mine && (back ? back === mine : mapping);
        Object.assign(o, {
          'Kind': 'Discovery', 'NextGen Payer': str(sec ? p['Secondary Payer'] : p['Primary Payer']), 'Fin Class': sec ? '' : str(x.finClass),
          'Source': viaGuar ? 'guarantor' : sec ? 'secondary' : mapping ? 'payer mapping' : 'discovery', 'Sent payerId': compare, 'Sent Confidence': sec ? '' : str(p.Confidence),
          'Returned payerId': ret, 'Returned Payer': str(r['Sohar Payer']),
          'Signal': !found || viaGuar ? 'neutral' : compare && ret === compare ? 'agree' : !memberMatch ? 'other coverage' : !compare ? 'found' : 'other payer',
          'Member Match': found && !viaGuar ? (memberMatch ? 'Yes' : 'No') : '',
        });
      }
      out.push(Object.fromEntries(PAYER_OUTCOME_COLS.map((c) => [c, o[c] == null ? '' : o[c]])));
    }
    return out;
  }

  /**
   * Append-only: previous rows stay; fresh rows are added unless their Key is already there. A row saved before v0.12
   * keeps its patient as MRN (its Per Nbr as a PER- key, converted with perToKey when given).
   */
  function appendOutcomes(previous, fresh, perToKey) {
    const seen = new Set(), rows = [];
    const norm = (r) => Object.fromEntries(PAYER_OUTCOME_COLS.map((c) => [c, c === 'MRN' ? migrateKey(savedKey(r), perToKey) : r[c] == null ? '' : r[c]]));
    for (const r of previous || []) { const k = str(r.Key); if (!k || seen.has(k)) continue; seen.add(k); rows.push(norm(r)); }
    let added = 0;
    for (const r of fresh || []) { const k = str(r.Key); if (!k || seen.has(k)) continue; seen.add(k); rows.push(norm(r)); added++; }
    return { rows, added };
  }

  const outcomeTime = (o) => { const t = soharTime(o['Sohar Date']); return isNaN(t) ? -Infinity : t; };
  const DEMOTE_MARK = 'Demoted by Sohar learning';
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;

  // "Looks right" on a human label Sohar disagrees with writes this note; the flag stays quiet until the conflicts grow
  // past the count recorded then. The newest note comes first (mergeNotes), so the first match is the latest review.
  const FLAG_ACK = 'Sohar disagreement reviewed';
  const flagAckNote = (conflicts, date) => `${FLAG_ACK} ${str(date)}, conflicts=${Number(conflicts) || 0}`;
  function flagAckConflicts(notes) {
    const m = /Sohar disagreement reviewed[^|]*?conflicts=(\d+)/.exec(str(notes));
    return m ? Number(m[1]) : -1;
  }

  const emptyGroup = (name) => ({ name, pers: new Set(), ver: new Map(), disc: new Map(), unrel: new Set(), last: null });
  /**
   * Ledger outcomes per NextGen payer name (lower case), keeping each patient's latest deciding outcome: per
   * (Sent payerId, patient) for verification (agree / payer error), per patient for discovery (a payer found).
   * Last-encounter verifications are left out: they test the Fin Class guess (learnFinClasses), not the proxy name.
   */
  function groupOutcomes(outcomes) {
    const g = new Map();
    for (const o of outcomes || []) {
      const name = str(o['NextGen Payer']), per = savedKey(o);
      if (!name || !per || isVip(name) || isSelfPay(name)) continue;
      const kind = str(o.Kind).toLowerCase(), sig = str(o.Signal), src = str(o.Source).toLowerCase();
      if (kind === 'verification' && src === 'last encounter') continue;
      // A guarantor search is about the guarantor's policy, and a payer test is judged on its own (payerTestSummary).
      if (src === 'guarantor' || src === 'payer test') continue;
      const k = name.toLowerCase();
      if (!g.has(k)) g.set(k, emptyGroup(name));
      const e = g.get(k);
      e.pers.add(per);
      if (!e.last || outcomeTime(o) >= outcomeTime(e.last)) e.last = o;
      // Only outcomes that say something about the payer compete for "latest": a later member ID error or empty
      // discovery does not erase an earlier agreement or conflict.
      if (kind === 'verification') {
        const pid = str(o['Sent payerId']);
        if (!pid || (sig !== 'agree' && sig !== 'payer error')) continue;
        if (!e.ver.has(pid)) e.ver.set(pid, new Map());
        const m = e.ver.get(pid), prev = m.get(per);
        if (!prev || outcomeTime(o) >= outcomeTime(prev)) m.set(per, o);
      } else if (kind === 'discovery' && str(o['Returned payerId'])) {
        // Evidence only when discovery found the mapped payer, or Sohar returned the member ID NextGen has (Member Match;
        // rows saved before v0.12 have none). Other coverage (another policy) is counted apart and never moves the mapping.
        if (sig !== 'agree' && str(o['Member Match']) !== 'Yes') { e.unrel.add(per); continue; }
        const prev = e.disc.get(per);
        if (!prev || outcomeTime(o) >= outcomeTime(prev)) e.disc.set(per, o);
      }
    }
    return g;
  }

  /** A name's grouped outcomes judged against payerId P: patient counts, top other payer, score and the demote rule. */
  function judgeGroup(e, P, L) {
    const agree = new Set(), perr = new Set(), disc = new Set(), other = new Set(), tally = new Map();
    if (e && P) for (const [per, o] of e.ver.get(P) || []) { const s = str(o.Signal); if (s === 'agree') agree.add(per); else if (s === 'payer error') perr.add(per); }
    for (const [per, o] of e ? e.disc : []) {
      const ret = str(o['Returned payerId']);
      if (!ret) continue;
      disc.add(per);
      if (P && ret === P) { agree.add(per); continue; }
      if (P) other.add(per);
      const t = tally.get(ret) || { payerId: ret, payerName: '', n: 0 };
      t.n++; t.payerName = str(o['Returned Payer']) || t.payerName;
      tally.set(ret, t);
    }
    const ranked = [...tally.values()].sort((a, b) => b.n - a.n || a.payerName.localeCompare(b.payerName) || a.payerId.localeCompare(b.payerId));
    const top = ranked[0] ? { ...ranked[0], share: ranked[0].n / disc.size } : null;
    const conflicts = perr.size + other.size;
    // Patients whose discovery found other coverage (not the same member ID) and no evidence: shown, never counted.
    const unrelated = e ? [...e.unrel].filter((per) => !e.disc.has(per)).length : 0;
    return {
      agree: agree.size, payerErrors: perr.size, discovery: disc.size, otherPayer: other.size, conflicts, top, unrelated,
      score: Math.round((100 * (agree.size + 1)) / (agree.size + conflicts + 2)),
      demoteRule: !!P && (perr.size >= L.demotePayerErrors || (disc.size >= L.demoteMinPatients && other.size / disc.size >= L.demoteDiscoveryShare)),
    };
  }

  /**
   * Roll the ledger up per NextGen payer name (case-insensitive), counting distinct patients from each patient's latest
   * deciding outcome: per (name, Sent payerId) for verification (agree / payer error), per name for discovery (a payer
   * found; re-judged against today's payerId, so a corrected mapping gets credit). Against the Payer Map's current payerId P:
   *  agree = latest verification of P eligible, or latest discovery found P; payer errors = latest verification of P a
   *  payer error; discovery = latest discovery found a payer; other payer = found one that is not P.
   * Last-encounter verifications do not count for the proxy name (see groupOutcomes).
   * score = round(100 * (agree + 1) / (agree + conflicts + 2)), conflicts = payer errors + other payer.
   * decision: 'promote' (Auto-Low / Auto-Medium, agree >= promote patients, no conflicts), 'demote' (Auto-High / Auto-Medium /
   * Learned meeting the demote rule), 'flag' (a human label meeting the demote rule; never changed; quiet after a person
   * reviewed it until the conflicts grow), else ''.
   * suggestion: for unmapped, Auto-Low or Review names, the payer discovery returned for >= suggest patients and >= suggest share.
   * A name on the Payer Map more than once is marked duplicate and never learned or suggested (learningReviewRows flags it).
   * One row per name in the ledger or the Payer Map, names with outcomes first. groups: groupOutcomes(outcomes), optional.
   */
  function learnPayers(outcomes, payerMap, cfg, groups) {
    const L = { ...LEARN_DEFAULTS, ...(cfg || {}) };
    const idx = payerMapIndex(payerMap);
    const g = groups || groupOutcomes(outcomes);
    const names = new Map(), times = new Map();
    for (const r of payerMap instanceof Map ? [] : payerMap || []) {
      const n = str(r.nextgenPayerName), k = n.toLowerCase();
      if (!n) continue;
      if (!names.has(k)) names.set(k, n);
      times.set(k, (times.get(k) || 0) + 1);
    }
    for (const [k, e] of g) if (!names.has(k)) names.set(k, e.name);
    for (const k of idx.keys()) if (!names.has(k)) names.set(k, k);
    const rows = [];
    for (const [k, name] of names) {
      const e = g.get(k) || emptyGroup(name);
      const h = idx.get(k) || null;
      const P = h ? h.payerId : '', conf = h ? h.confidence : '', c = conf.toLowerCase();
      const duplicate = (times.get(k) || 0) > 1;
      const j = judgeGroup(e, P, L), top = j.top;
      const why = `${plural(j.payerErrors, 'patient')} with a Sohar payer error; ${j.otherPayer} of ${j.discovery} discovery patients found another payer${top && j.otherPayer ? ` (most often ${top.payerName || top.payerId}, ${top.payerId})` : ''}`;
      let decision = '', note = '';
      if (duplicate) {
        // Routing reads the last row, the pane and learning would write the first: learn nothing until one row remains.
      } else if (h && P && (c === 'auto-low' || c === 'auto-medium') && j.agree >= L.promotePatients && j.conflicts === 0) {
        decision = 'promote'; note = `Learned from Sohar: ${j.agree} patients agreed`;
      } else if (h && (c === 'auto-high' || c === 'auto-medium' || c === 'learned') && j.demoteRule) {
        decision = 'demote'; note = `${DEMOTE_MARK} (was ${conf}): ${why}`;
      } else if (h && !isAutoLabel(c) && j.demoteRule && j.conflicts > flagAckConflicts(h.notes)) {
        decision = 'flag'; note = `Sohar disagrees with Payer Map: ${why}`;
      }
      const suggestion = !duplicate && (!h || c === 'auto-low' || c === 'review') && top && top.n >= L.suggestPatients && top.share >= L.suggestShare ? top : null;
      const lt = e.last ? outcomeTime(e.last) : -Infinity;
      rows.push({
        key: k, nextgenPayer: name, mapped: !!h, payerId: P, payerName: h ? h.payerName : '', confidence: conf, notes: h ? h.notes : '',
        agree: j.agree, payerErrors: j.payerErrors, discovery: j.discovery, otherPayer: j.otherPayer, unrelated: j.unrelated, conflicts: j.conflicts, score: j.score, n: e.pers.size, top,
        decision, note, suggestion, duplicate, duplicates: times.get(k) || 0,
        lastOutcome: e.last ? `${isFinite(lt) ? `${fmtDate(new Date(lt))} ` : ''}${str(e.last.Kind)}: ${str(e.last.Signal)}` : '',
      });
    }
    rows.sort((a, b) => b.n - a.n || a.nextgenPayer.toLowerCase().localeCompare(b.nextgenPayer.toLowerCase()));
    return rows;
  }

  /**
   * Fin Class -> payerId learned from last-encounter verifications: a Fin Class whose verified payerId has >= promote
   * patients agreeing and none with a payer error, and no other payerId with any agreeing patient (an ambiguous Fin Class
   * such as "Commercial" is never learned). Returns [{finClass, payerId, payerName, nextgenPayer, patients}].
   */
  function learnFinClasses(outcomes, cfg) {
    const L = { ...LEARN_DEFAULTS, ...(cfg || {}) };
    const fk = (s) => str(s).toLowerCase().replace(/\s+/g, ' ');
    const g = new Map();
    for (const o of outcomes || []) {
      if (str(o.Source).toLowerCase() !== 'last encounter' || str(o.Kind).toLowerCase() !== 'verification') continue;
      const fc = str(o['Fin Class']), pid = str(o['Sent payerId']), per = savedKey(o);
      if (!fc || !pid || !per || isVip(fc) || isSelfPay(fc) || (o.Signal !== 'agree' && o.Signal !== 'payer error')) continue; // latest deciding outcome
      if (!g.has(fk(fc))) g.set(fk(fc), { finClass: fc, ids: new Map() });
      const ids = g.get(fk(fc)).ids;
      if (!ids.has(pid)) ids.set(pid, new Map());
      const prev = ids.get(pid).get(per);
      if (!prev || outcomeTime(o) >= outcomeTime(prev)) ids.get(pid).set(per, o);
    }
    const out = [];
    for (const { finClass, ids } of g.values()) {
      const stats = [...ids].map(([payerId, m]) => {
        const list = [...m.values()];
        const agreeing = list.filter((o) => o.Signal === 'agree');
        return { payerId, agree: agreeing.length, errors: list.filter((o) => o.Signal === 'payer error').length, latest: agreeing.sort((a, b) => outcomeTime(b) - outcomeTime(a))[0] };
      });
      const ok = stats.filter((s) => s.agree >= L.promotePatients && !s.errors);
      if (ok.length !== 1 || stats.some((s) => s !== ok[0] && s.agree > 0)) continue;
      const s = ok[0];
      out.push({ finClass, payerId: s.payerId, payerName: str(s.latest['Returned Payer']), nextgenPayer: str(s.latest['NextGen Payer']), patients: s.agree });
    }
    return out.sort((a, b) => a.finClass.localeCompare(b.finClass));
  }

  /**
   * Payer Map changes from the ledger, applied to a copy of the map in order: learning (promote, demote) for automatic
   * labels, then re-score of Auto-Low rows the matcher now places at Auto-Medium or Auto-High (upgrade only; never a row
   * learning demoted or one with Sohar conflicts). With cfg.on false nothing changes; the rollup, flags and suggestions remain.
   * Returns {rollup (after the changes), changes: [{entry, by, old, score, evidence}], payerMap, finClasses, counts}.
   */
  function planLearning(outcomes, payerMap, soharPayers, cfg) {
    const L = { ...LEARN_DEFAULTS, ...(cfg || {}) };
    const groups = groupOutcomes(outcomes);
    const first = learnPayers(outcomes, payerMap, L, groups);
    const map = (payerMap || []).map((r) => ({ ...r }));
    const rowOf = (name) => map.find((r) => str(r.nextgenPayerName).toLowerCase() === str(name).toLowerCase());
    const changes = [];
    const apply = (row, entry, by, score, evidence) => {
      changes.push({ entry, by, score, evidence, old: { payerId: str(row.payerId), confidence: str(row.confidence) } });
      Object.assign(row, { payerId: entry.payerId, payerName: entry.payerName, confidence: entry.confidence, notes: mergeNotes(entry.notes, row.notes) });
    };
    if (L.on) {
      for (const r of first) {
        if (r.decision !== 'promote' && r.decision !== 'demote') continue;
        const row = rowOf(r.nextgenPayer);
        if (!row) continue;
        apply(row, { nextgenPayerName: str(row.nextgenPayerName), payerId: r.payerId, payerName: r.payerName,
          confidence: r.decision === 'promote' ? 'Learned' : 'Auto-Low', notes: r.note }, 'Learning', r.score, r.note);
      }
      if (soharPayers && soharPayers.length) {
        const firstBy = new Map(first.map((r) => [r.key, r]));
        const confirmed = confirmedNames(map);
        for (const row of map) {
          const name = str(row.nextgenPayerName), k = name.toLowerCase();
          if (!name || str(row.confidence).toLowerCase() !== 'auto-low' || isVip(name) || isSelfPay(name)) continue;
          const fr = firstBy.get(k);
          if ((fr && (fr.conflicts > 0 || fr.duplicate)) || str(row.notes).includes(DEMOTE_MARK)) continue;
          const m = autoMatchPayer(name, soharPayers, { confirmed });
          if (m.confidence !== 'Auto-Medium' && m.confidence !== 'Auto-High') continue;
          // Judge the evidence against the payerId the matcher picked, not only the current one (often blank): no
          // upgrade into a mapping Sohar already contradicts, which learning would demote on the next build.
          if (fr && fr.suggestion && fr.suggestion.payerId !== m.payerId) continue;
          const j = judgeGroup(groups.get(k), m.payerId, L);
          if (j.demoteRule || j.payerErrors || j.otherPayer) continue;
          apply(row, { nextgenPayerName: name, payerId: m.payerId, payerName: m.payerName, confidence: m.confidence, notes: `Re-scored by the add-in: ${m.evidence}` }, 'Re-score', m.score, m.evidence);
        }
      }
    }
    const rollup = changes.length ? learnPayers(outcomes, map, L) : first;
    const by = (b, conf) => changes.filter((c) => c.by === b && (!conf || c.entry.confidence === conf)).length;
    return {
      rollup, changes, payerMap: map, finClasses: L.on ? learnFinClasses(outcomes, L) : [],
      counts: { promoted: by('Learning', 'Learned'), demoted: by('Learning', 'Auto-Low'), rescored: by('Re-score'),
        suggestions: rollup.filter((r) => r.suggestion).length, flagged: rollup.filter((r) => r.decision === 'flag').length },
    };
  }

  /** "Sohar: 9 agreed, 1 other payer, score 83" for a rollup row with outcomes, else ''. */
  function learningEvidence(r) {
    if (!r || !r.n) return '';
    if (!r.agree && !r.conflicts && !r.discovery) return `Sohar: ${plural(r.n, 'patient')} checked, nothing that confirms or contradicts the payer yet${r.unrelated ? ` (${r.unrelated} found other coverage, not evidence)` : ''}`;
    const parts = [`${r.agree} agreed`];
    if (r.payerErrors) parts.push(plural(r.payerErrors, 'payer error'));
    if (r.otherPayer) parts.push(`${r.otherPayer} other payer${r.otherPayer === 1 ? '' : 's'} for the same member ID`);
    parts.push(`score ${r.score}`);
    if (r.unrelated) parts.push(`${r.unrelated} found other coverage (not evidence)`);
    return `Sohar: ${parts.join(', ')}`;
  }
  /** "3 of 4 discovery patients" or "2 of 2 payer tests", for a suggestion. */
  const suggestionWhy = (s, r) => (s.source === 'test' ? `${s.n} of ${s.tried} payer test patient${s.tried === 1 ? '' : 's'} eligible` : `${s.n} of ${r.discovery} discovery patients (same member ID)`);

  // Payer Learning sheet: one row per NextGen payer in the ledger or the Payer Map. No patient identifiers.
  const PAYER_LEARNING_COLS = ['NextGen Payer', 'Sohar payerId', 'Sohar payerName', 'Confidence', 'Evidence score', 'Patients agreed', 'Payer errors',
    'Discovery checked', 'Discovery other payer', 'Most common other payer', 'Suggestion', 'Last outcome', 'Discovery other coverage (not evidence)', 'Payer tests'];
  function payerLearningRows(rollup) {
    return (rollup || []).map((r) => ({
      'NextGen Payer': r.nextgenPayer, 'Sohar payerId': r.payerId, 'Sohar payerName': r.payerName, 'Confidence': r.mapped ? r.confidence : 'Not in Payer Map',
      'Evidence score': r.n ? String(r.score) : '', 'Patients agreed': String(r.agree), 'Payer errors': String(r.payerErrors),
      'Discovery checked': String(r.discovery), 'Discovery other payer': String(r.otherPayer),
      'Most common other payer': r.top ? `${r.top.payerName || r.top.payerId} (${r.top.payerId}): ${r.top.n} of ${r.discovery}` : '',
      'Suggestion': r.suggestion ? `Use ${r.suggestion.payerName || r.suggestion.payerId} (${r.suggestion.payerId}): ${suggestionWhy(r.suggestion, r)}`
        : r.decision === 'flag' ? r.note : '',
      'Last outcome': r.lastOutcome,
      'Discovery other coverage (not evidence)': String(r.unrelated || 0),
      'Payer tests': r.tests ? r.tests.text : '',
    }));
  }

  /**
   * Review rows about payers (MRN blank): High for human-labeled payers Sohar disagrees with (FLAG), Medium for a
   * name on the Payer Map more than once (learning skips it).
   */
  function learningReviewRows(rollup) {
    const flags = (rollup || []).filter((r) => r.decision === 'flag').map((r) => ({
      Severity: 'High', Issue: 'Sohar disagrees with Payer Map', MRN: '',
      Detail: `${r.nextgenPayer} -> ${r.payerName || r.payerId} (${r.payerId}, ${r.confidence}): ${r.note.replace(/^Sohar disagrees with Payer Map: /, '')}`,
      Fix: 'Check the mapping in the pane (Admin, step 1) or on the Payer Map sheet; "Looks right" there marks it reviewed. The add-in never changes a human label.',
    }));
    const dups = (rollup || []).filter((r) => r.duplicate).map((r) => ({
      Severity: 'Medium', Issue: 'Payer Map name listed twice', MRN: '',
      Detail: `${r.nextgenPayer} is on the Payer Map ${r.duplicates} times (names match ignoring case). Payer learning skips it until one row remains.`,
      Fix: 'Delete the extra rows on the Payer Map sheet so the name appears once.',
    }));
    return flags.concat(dups);
  }

  // Payer Map Log sheet: every Payer Map change, append-only. No patient identifiers.
  const PAYER_MAP_LOG_COLS = ['When', 'NextGen Payer', 'Old payerId', 'Old Confidence', 'New payerId', 'New Confidence', 'By', 'Score', 'Evidence'];
  const logRow = (when, name, old, nw, by, score, evidence) => ({
    'When': str(when), 'NextGen Payer': str(name), 'Old payerId': old ? str(old.payerId) : '', 'Old Confidence': old ? str(old.confidence) : '',
    'New payerId': str(nw.payerId), 'New Confidence': str(nw.confidence), 'By': str(by), 'Score': score == null ? '' : String(score), 'Evidence': str(evidence),
  });
  /**
   * Log rows for Payer Map writes. changes: [{entry: {nextgenPayerName, payerId, confidence}, by, score, evidence, old?}];
   * old defaults to the row in oldMap. A write that changes neither payerId nor confidence is not logged, unless
   * ch.always (a person reviewing a Sohar disagreement without changing the mapping).
   */
  function mapChangeLog(oldMap, changes, when) {
    const idx = payerMapIndex(oldMap || []);
    const out = [];
    for (const ch of changes || []) {
      const e = ch.entry || ch, name = str(e.nextgenPayerName);
      if (!name) continue;
      const old = ch.old !== undefined ? ch.old : idx.get(name.toLowerCase()) || null;
      if (!ch.always && old && str(old.payerId) === str(e.payerId) && str(old.confidence).toLowerCase() === str(e.confidence).toLowerCase()) continue;
      out.push(logRow(when, name, old, e, ch.by, ch.score, ch.evidence));
    }
    return out;
  }
  /**
   * Payer Map rows that differ from their last logged state were edited on the sheet: one 'Sheet edit' row each.
   * An empty log starts with one 'Baseline' row per Payer Map row.
   */
  function sheetEditLog(payerMap, logRows, when) {
    const rows = (payerMap || []).filter((r) => str(r.nextgenPayerName));
    const seen = new Set();
    const uniq = rows.filter((r) => { const k = str(r.nextgenPayerName).toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; });
    if (!(logRows || []).some((l) => str(l['NextGen Payer']))) {
      return uniq.map((r) => logRow(when, r.nextgenPayerName, null, r, 'Baseline', '', 'Payer Map when the log started'));
    }
    const last = new Map();
    for (const l of logRows) { const k = str(l['NextGen Payer']).toLowerCase(); if (k) last.set(k, l); }
    const out = [];
    for (const r of uniq) {
      const l = last.get(str(r.nextgenPayerName).toLowerCase());
      if (l && str(l['New payerId']) === str(r.payerId) && str(l['New Confidence']).toLowerCase() === str(r.confidence).toLowerCase()) continue;
      out.push(logRow(when, r.nextgenPayerName, l ? { payerId: l['New payerId'], confidence: l['New Confidence'] } : null, r, 'Sheet edit', '',
        l ? 'Changed on the Payer Map sheet' : 'Added on the Payer Map sheet'));
    }
    return out;
  }

  /** One High Review row for Sohar results with no patient identifiers (Match 'No identifiers'), with the count; else []. */
  function noIdentifierReviewRows(results) {
    const rows = (results || []).filter((r) => r && r.Match === 'No identifiers');
    if (!rows.length) return [];
    const refs = rows.map((r) => str(r['Sohar Ref'])).filter(Boolean);
    const n = rows.length;
    return [{
      Severity: 'High', Issue: 'Sohar results with no patient identifiers', MRN: '',
      Detail: `${plural(n, 'Sohar result')} with no id, member ID, name or DOB${refs.length ? ` (Sohar ref ${refs.slice(0, 5).join(', ')}${refs.length > 5 ? ', ...' : ''})` : ''}. Likely an export anomaly or a check run outside the batch.`,
      Fix: 'Open the Sohar link on those Sohar Results rows (Match = No identifiers) to see who each one is. No NextGen task is made for them.',
    }];
  }

  // ---------- re-verify skip window (approved by Brendan 2026-10-20) ----------
  // A verification-routed policy (payerId + member ID) recently verified eligible is left out of the verification file;
  // the clinic R9 Re-verify flag uses the same due rule (reverifyDue / reverifyCheck).
  const REVERIFY_DEFAULTS = { skip: true, days: 7, newMonth: false };
  const REVERIFY_CONFIG_ROWS = [
    { setting: 'Skip recent verifications', value: 'On', what: 'On or Off. On: a patient whose same insurance (Sohar payerId + member ID) Sohar verified eligible within Re-verify days is left out of the verification file (still on Patients; the pane lists them with a Re-verify anyway button). Failed, ineligible or pending results never skip.' },
    { setting: 'Re-verify days', value: 7, what: 'Calendar days a Sohar eligible verification stays current: skipped until then, and the clinic Re-verify flag after. A next visit in a new month within these days re-verifies sooner.' },
    { setting: 'Re-verify in new month', value: 'Off', what: 'On or Off. On: also re-verify whenever the month changed since the verification, or the next visit is in a different month, whatever the days.' },
  ];
  /** Config {Setting: Value} -> {skip, days, newMonth}. Re-verify days: a whole number >= 0, default 7. */
  function reverifyConfig(cfg) {
    const c = cfg || {};
    const raw = (k) => (c[k] == null ? '' : String(c[k]).trim());
    const n = parseInt(raw('Re-verify days'), 10);
    return {
      skip: !/^(off|no|n|false|0)$/i.test(raw('Skip recent verifications')),
      days: isNaN(n) || n < 0 ? REVERIFY_DEFAULTS.days : n,
      newMonth: /^(on|yes|y|true|1)$/i.test(raw('Re-verify in new month')),
    };
  }
  /** Local calendar day of a Sohar Date (resolved || created) as a UTC-midnight Date, else null. */
  function soharDay(v) { const t = soharTime(v); return isNaN(t) ? null : dayOf(new Date(t)); }
  const isVerResult = (r) => !!r && r['Sent As'] === 'Verification' && r['Sohar Status'] !== 'Not in Sohar file';
  // A name-only match, an unmatched row or one with no identifiers never stands for the patient.
  const untrustedMatch = (r) => /^name|^no match|^no identifiers/i.test(str(r.Match));
  const monthOf = (d) => d.getUTCFullYear() * 12 + d.getUTCMonth();

  /**
   * Is a Sohar re-verification due? rows: Sohar Results rows (verification) for one policy. opts: {nextAppt (Date or
   * MM/DD/YYYY), today (Date, local day used), days (default 7), newMonth (bool)}. Not due only when the latest row
   * (by Sohar Date) is complete.eligible, verified less than days calendar days ago, its Coverage End is not before the
   * next visit, the next visit is not in a different month/year within days of today, and (newMonth On) neither today
   * nor the next visit is in a different month/year than the verification.
   * Returns {due, code ('none'|'not eligible'|'window'|'coverage end'|'new month'|'fresh'), reason, verified, dueDate, daysSince, status}.
   */
  function reverifyDue(rows, opts) {
    const o = opts || {};
    const days = o.days != null && !isNaN(+o.days) ? +o.days : REVERIFY_DEFAULTS.days;
    const today = dayOf(o.today instanceof Date && !isNaN(o.today) ? o.today : new Date());
    let best = null, bt = -Infinity;
    for (const r of rows || []) {
      if (!isVerResult(r)) continue;
      const t = soharTime(r['Sohar Date']);
      if (isNaN(t)) continue;
      if (!best || t > bt || (t === bt && str(r['Sohar Date']) > str(best['Sohar Date']))) { best = r; bt = t; }
    }
    if (!best) return { due: true, code: 'none', reason: 'no Sohar verification on file', verified: '', dueDate: '', daysSince: null, status: '' };
    const vDay = dayOf(new Date(bt));
    const since = Math.round((today - vDay) / MS_DAY);
    const status = str(best['Sohar Status']);
    const base = { verified: fmtDate(vDay), daysSince: since, status };
    const due = (code, reason) => ({ ...base, due: true, code, reason, dueDate: fmtDate(today) });
    if (!/^complete\.eligible/i.test(status)) return due('not eligible', `latest verification ${fmtDate(vDay)} came back ${status || 'pending'}`);
    if (since >= days) return due('window', `verified eligible ${fmtDate(vDay)}, ${since} day${since === 1 ? '' : 's'} ago`);
    const next = o.nextAppt instanceof Date ? dayOf(o.nextAppt) : parseDate(o.nextAppt);
    const end = parseDate(best['Coverage End']);
    if (end && next && end < next) return due('coverage end', `coverage ends ${fmtDate(end)}, before the ${fmtDate(next)} visit`);
    let dueDay = new Date(vDay.getTime() + days * MS_DAY);
    if (next && monthOf(next) !== monthOf(vDay)) {
      if (o.newMonth) return due('new month', `next visit ${fmtDate(next)} is in a different month than the ${fmtDate(vDay)} verification`);
      if ((next - today) / MS_DAY < days) return due('new month', `next visit ${fmtDate(next)} is in a new month, within ${days} days`);
      const sooner = new Date(next.getTime() - (days - 1) * MS_DAY);
      if (sooner < dueDay) dueDay = sooner;
    }
    if (o.newMonth) {
      if (monthOf(today) !== monthOf(vDay)) return due('new month', `verified ${fmtDate(vDay)}, in an earlier month`);
      const first = new Date(Date.UTC(vDay.getUTCFullYear(), vDay.getUTCMonth() + 1, 1));
      if (first < dueDay) dueDay = first;
    }
    return { ...base, due: false, code: 'fresh', reason: `verified eligible ${fmtDate(vDay)}, next check ${fmtDate(dueDay)}`, dueDate: fmtDate(dueDay) };
  }

  /**
   * The Sohar Results rows (verification, trusted match) for one verifyRows entry: same Per Nbr (when given), same
   * member ID (ignoring punctuation and leading zeros; a blank one never matches), same Sohar payerId, and the same
   * Sohar payer name when both are known. Without a recorded payerId sent (row.sentPayerId), both names must be known.
   * v null: every trusted verification row (patients with no verifyRows).
   */
  function policyResults(rows, v, per) {
    return (rows || []).filter((r) => {
      if (!isVerResult(r) || untrustedMatch(r) || (per != null && str(r['Per Nbr']) !== str(per))) return false;
      if (!v) return true;
      const mid = idKey(r['Sohar Member ID']);
      if (!mid || mid !== idKey(v.memberId) || str(r['Sohar payerId']) !== str(v.payerId)) return false;
      const a = str(r['Sohar Payer']).toLowerCase(), b = str(v.payerName).toLowerCase();
      if (a && b && a !== b) return false;
      // Recorded payerId sent: that is the policy. Otherwise Sohar payerId is today's chart (matched by member ID), so the
      // Sohar payer name has to agree too; a blank name on either side means due, never a match.
      return !!str(r.sentPayerId) || (!!a && !!b);
    });
  }
  /** reverifyDue for one policy (verifyRows entry, or null for any verification) from Sohar Results rows. */
  function reverifyCheck(rows, v, opts) { return reverifyDue(policyResults(rows, v, opts && opts.per), opts); }

  // Hidden Re-verify Requests sheet: someone asked to send a recently verified patient anyway. By: who asked (the Clinic
  // tab, a clinic file brought back, or the admin's "re-verify all"). Rows saved before v0.12 have Per Nbr, read as PER- keys.
  const REVERIFY_REQUEST_COLS = ['MRN', 'Requested', 'By'];
  /**
   * Requests still waiting: a request is dropped once that patient has a Sohar verification (any status) dated after it.
   * A request whose time cannot be read is kept. perToKey (optional) converts rows saved before v0.12.
   * Returns {rows (kept), active: Map(key -> Requested), dropped}.
   */
  function pruneReverifyRequests(requests, results, perToKey) {
    const latest = new Map();
    for (const r of results || []) {
      const k = keyOf(r);
      if (!isVerResult(r) || untrustedMatch(r) || !k) continue;
      const t = soharTime(r['Sohar Date']);
      if (!isNaN(t) && t > (latest.get(k) ?? -Infinity)) latest.set(k, t);
    }
    const rows = [], active = new Map(), seen = new Set();
    let dropped = 0;
    for (const q of requests || []) {
      const k = migrateKey(savedKey(q), perToKey), when = str(q.Requested);
      if (!k || seen.has(k)) { if (k) dropped++; continue; }
      const t = soharTime(when);
      if (!isNaN(t) && latest.has(k) && latest.get(k) > t) { dropped++; continue; }
      seen.add(k);
      rows.push({ 'MRN': k, 'Requested': when, 'By': str(q.By) });
      active.set(k, when);
    }
    return { rows, active, dropped };
  }
  /** Requests with more (or refreshed) patient keys: keys a key or a list of keys. by: who asked. */
  function addReverifyRequest(requests, keys, when, by) {
    const add = new Set((Array.isArray(keys) ? keys : [keys]).map(str).filter(Boolean));
    const rows = (requests || []).map((q) => ({ 'MRN': savedKey(q), 'Requested': str(q.Requested), 'By': str(q.By) })).filter((q) => q.MRN && !add.has(q.MRN));
    for (const k of add) rows.push({ 'MRN': k, 'Requested': str(when), 'By': str(by) });
    return rows;
  }
  /** Requests without these patient keys (a clinic unticked "Re-verify anyway", or the admin cancelled all). */
  function removeReverifyRequest(requests, keys) {
    const drop = new Set((Array.isArray(keys) ? keys : [keys]).map(str).filter(Boolean));
    return (requests || []).map((q) => ({ 'MRN': savedKey(q), 'Requested': str(q.Requested), 'By': str(q.By) })).filter((q) => q.MRN && !drop.has(q.MRN));
  }
  /**
   * Requests from a clinic file brought back by the admin: its Re-verify Requests rows (MRN, Requested, Action On/Off).
   * Ticks add a request, unticks remove one. Returns the new request rows and counts.
   */
  function mergeClinicRequests(requests, clinicRows, by) {
    let rows = (requests || []).slice(), added = 0, removed = 0;
    for (const c of clinicRows || []) {
      const k = savedKey(c);
      if (!k) continue;
      const on = !/^(off|no|n|false|0|removed)$/i.test(str(c['Re-verify Anyway'] ?? c.Action ?? 'On'));
      const had = rows.some((q) => savedKey(q) === k);
      if (on && !had) { rows = addReverifyRequest(rows, k, str(c.Requested) || new Date().toISOString(), by); added++; }
      else if (!on && had) { rows = removeReverifyRequest(rows, k); removed++; }
    }
    return { rows, added, removed };
  }

  /**
   * Apply the skip window to built patients (buildPatients). results: Sohar Results rows from the saved uploads
   * (mergeSoharResults). opts: {skip, days, newMonth (reverifyConfig), today, requests (Map or Set of patient keys asking
   * to re-verify anyway)}. Only in-scope verification patients are checked. Each verifyRows entry not due gets reverifySkip
   * {verified, due, reason}; buildExports leaves those rows out. A patient with a skipped row keeps Route verification,
   * gets "; skipped: verified eligible MM/DD/YYYY, next check MM/DD/YYYY" on Route Reason and x.reverifySkip
   * {verified, due, reason, partial} (partial: another insurance still goes). Returns {patients (copies), skipped: [{key,
   * per, mrn, last, first, location, nextAppt, verified, due, reason, partial}], requested: [same shape: patients the skip
   * would have left out but a request sends anyway]}.
   */
  function applyReverifySkips(patients, results, opts) {
    const o = { ...REVERIFY_DEFAULTS, ...(opts || {}) };
    const asked = o.requests instanceof Map || o.requests instanceof Set ? o.requests : new Set(o.requests || []);
    const byKey = new Map();
    for (const r of results || []) { const k = keyOf(r); if (!k) continue; if (!byKey.has(k)) byKey.set(k, []); byKey.get(k).push(r); }
    const skipped = [], requested = [];
    const out = (patients || []).map((p) => {
      const k = keyOf(p);
      const vrs = (p.x && p.x.verifyRows) || [];
      if (!o.skip || p.Route !== 'verification' || p['In Scope'] !== 'Yes' || !vrs.length) return p;
      const rows = byKey.get(k) || [];
      const checks = vrs.map((v) => reverifyCheck(rows, v, { ...o, nextAppt: p['Next Appt'] }));
      if (checks.every((c) => c.due)) return p;
      const first = checks.find((c) => !c.due);
      const partial = checks.some((c) => c.due);
      const skip = { verified: first.verified, due: first.dueDate, reason: first.reason, partial };
      const who = { key: k, per: str(p['Per Nbr']), mrn: str(p.MRN), last: p['Last Name'], first: p['First Name'],
        location: str(p.x && p.x.location), nextAppt: str(p['Next Appt']), ...skip };
      if (asked.has(k)) {
        requested.push({ ...who, requested: asked instanceof Map ? str(asked.get(k)) : '' });
        return { ...p, x: { ...p.x, reverifyRequested: true, reverifySkipped: skip } };
      }
      skipped.push(who);
      return {
        ...p, 'Route Reason': `${p['Route Reason']}; skipped: ${first.reason}${partial ? ' (the other insurance is still sent)' : ''}`,
        x: { ...p.x, reverifySkip: skip, verifyRows: vrs.map((v, i) => (checks[i].due ? v : { ...v, reverifySkip: { verified: checks[i].verified, due: checks[i].dueDate, reason: checks[i].reason } })) },
      };
    });
    return { patients: out, skipped, requested };
  }

  // ---------- guarantor resubmission (v0.12, approved by Brendan 2026-10-20) ----------
  // A patient with a guarantor on the chart (Guar Name, not the patient) and no coverage from Sohar can be searched again
  // with the guarantor as the person (discovery id <key>-G, the guarantor's name and Guar Birth Dt). The admin is asked
  // after an upload; a yes saves a request, and the next discovery file carries the guarantor row.
  const GUARANTOR_REASONS = { none: 'Discovery found nothing', failed: 'Verification ineligible or failed' };
  const GUARANTOR_REQUEST_COLS = ['MRN', 'Requested', 'Reason', 'By'];
  const resolvedResult = (r) => !!r && r['Sohar Status'] !== 'Not in Sohar file' && !untrustedMatch(r) && !/^pending/i.test(str(r['Sohar Status']));
  const latestOf = (rows) => rows.slice().sort((a, b) => (soharTime(b['Sohar Date']) || 0) - (soharTime(a['Sohar Date']) || 0))[0] || null;
  /** A patient's Sohar results by key (trusted matches only). */
  function resultsByKey(results) {
    const m = new Map();
    for (const r of results || []) { const k = keyOf(r); if (!k || untrustedMatch(r)) continue; if (!m.has(k)) m.set(k, []); m.get(k).push(r); }
    return m;
  }

  /**
   * Patients without results who have a guarantor: in scope, a guarantor who is not the patient, and no coverage from Sohar:
   * the latest patient discovery found nothing, or (no discovery found coverage and) every latest verification came back
   * ineligible or with an error. Left out: an eligible verification, coverage found, a guarantor search already run or
   * requested. requests: Guarantor Requests rows (or keys). Returns [{key, per, mrn, last, first, location, nextAppt,
   * reason, hasDob, guarantor}], soonest visit first.
   */
  function guarantorCandidates(patients, results, requests) {
    const byKey = resultsByKey(results);
    const asked = new Set((requests || []).map((q) => (typeof q === 'string' ? q : savedKey(q))));
    const out = [];
    for (const p of patients || []) {
      const g = p && p.x && p.x.guarantor, k = keyOf(p);
      if (!g || p['In Scope'] !== 'Yes' || !k || asked.has(k)) continue;
      const rows = (byKey.get(k) || []).filter(resolvedResult);
      if (rows.some((r) => r['Searched As'] === 'Guarantor')) continue;
      const disc = rows.filter((r) => r['Sent As'] === 'Discovery');
      const ver = rows.filter((r) => r['Sent As'] === 'Verification' && !/^Superseded/.test(str(r.Action)));
      if (disc.some((r) => r['Sohar Status'] === 'Coverage found') || ver.some((r) => /^complete\.eligible/i.test(str(r['Sohar Status'])))) continue;
      const lastDisc = latestOf(disc);
      let reason = '';
      if (lastDisc && lastDisc['Sohar Status'] === 'No coverage found') reason = GUARANTOR_REASONS.none;
      else if (ver.length && ver.every((r) => /^(complete\.ineligible|error)/i.test(str(r['Sohar Status'])))) reason = GUARANTOR_REASONS.failed;
      if (!reason) continue;
      out.push({ key: k, per: str(p['Per Nbr']), mrn: str(p.MRN), last: p['Last Name'], first: p['First Name'], location: str(p.x.location),
        nextAppt: str(p['Next Appt']), reason, hasDob: !!g.dob, guarantor: g });
    }
    const t = (s) => { const d = parseDate(s); return d ? d.getTime() : Infinity; };
    return out.sort((a, b) => t(a.nextAppt) - t(b.nextAppt) || str(a.last).localeCompare(str(b.last)) || str(a.first).localeCompare(str(b.first)));
  }
  /** Requests plus these candidates (guarantorCandidates entries). */
  function addGuarantorRequests(requests, candidates, when, by) {
    const add = new Map((candidates || []).filter((c) => c && c.key).map((c) => [c.key, c]));
    const rows = (requests || []).map((q) => ({ 'MRN': savedKey(q), 'Requested': str(q.Requested), 'Reason': str(q.Reason), 'By': str(q.By) })).filter((q) => q.MRN && !add.has(q.MRN));
    for (const [k, c] of add) rows.push({ 'MRN': k, 'Requested': str(when), 'Reason': str(c.reason), 'By': str(by) });
    return rows;
  }
  /**
   * Requests still waiting: dropped once a guarantor search (<key>-G) for that patient came back after the request.
   * perToKey (optional) converts PER- keys. Returns {rows, active: Map(key -> row), dropped}.
   */
  function pruneGuarantorRequests(requests, results, perToKey) {
    const back = new Map();
    for (const r of results || []) {
      if (r['Searched As'] !== 'Guarantor' || !resolvedResult(r)) continue;
      const k = keyOf(r), t = soharTime(r['Sohar Date']);
      if (k && !isNaN(t) && t > (back.get(k) ?? -Infinity)) back.set(k, t);
    }
    const rows = [], active = new Map();
    let dropped = 0;
    for (const q of requests || []) {
      const k = migrateKey(savedKey(q), perToKey);
      if (!k || active.has(k)) { if (k) dropped++; continue; }
      const t = soharTime(q.Requested);
      if (back.has(k) && (isNaN(t) || back.get(k) > t)) { dropped++; continue; }
      const row = { 'MRN': k, 'Requested': str(q.Requested), 'Reason': str(q.Reason), 'By': str(q.By) };
      rows.push(row); active.set(k, row);
    }
    return { rows, active, dropped };
  }

  /**
   * Themes from guarantor searches: for each patient with a guarantor search back, what the patient's own checks found
   * against what the guarantor search found. Theme: 'Guarantor only', 'Patient only', 'Both, same payer', 'Both,
   * different payers', 'Neither'. Patient side: latest patient discovery with coverage, else latest eligible verification.
   * Returns {patients: [{key, theme, nextgenPayer, patientPayer, guarantorPayer, reason}], rows (GUARANTOR_THEME_COLS,
   * no patient identifiers: one per theme x NextGen payer x Sohar payers), counts: {theme: n}}.
   */
  const GUARANTOR_THEME_COLS = ['Theme', 'NextGen Payer', 'Found for patient', 'Found for guarantor', 'Patients'];
  function guarantorThemes(results) {
    const byKey = resultsByKey(results);
    const patients = [];
    for (const [k, rows0] of byKey) {
      const rows = rows0.filter(resolvedResult);
      const g = latestOf(rows.filter((r) => r['Searched As'] === 'Guarantor'));
      if (!g) continue;
      const own = rows.filter((r) => r['Searched As'] !== 'Guarantor');
      const pd = latestOf(own.filter((r) => r['Sent As'] === 'Discovery' && r['Sohar Status'] === 'Coverage found'));
      const pv = latestOf(own.filter((r) => r['Sent As'] === 'Verification' && /^complete\.eligible/i.test(str(r['Sohar Status'])) && !/^Superseded/.test(str(r.Action))));
      const mine = pd || pv;
      const gFound = g['Sohar Status'] === 'Coverage found';
      const pid = (r) => str(r && r['Sohar payerId']) || str(r && r['Sohar Payer']).toLowerCase();
      const theme = gFound && mine ? (pid(g) === pid(mine) ? 'Both, same payer' : 'Both, different payers') : gFound ? 'Guarantor only' : mine ? 'Patient only' : 'Neither';
      patients.push({ key: k, theme, nextgenPayer: str(g['NextGen Payer']) || 'None on file', patientPayer: mine ? str(mine['Sohar Payer']) : '', guarantorPayer: gFound ? str(g['Sohar Payer']) : '' });
    }
    const counts = {}, agg = new Map();
    for (const p of patients) {
      counts[p.theme] = (counts[p.theme] || 0) + 1;
      const k = [p.theme, p.nextgenPayer, p.patientPayer, p.guarantorPayer].join('|').toLowerCase();
      if (!agg.has(k)) agg.set(k, { 'Theme': p.theme, 'NextGen Payer': p.nextgenPayer, 'Found for patient': p.patientPayer || '—', 'Found for guarantor': p.guarantorPayer || '—', n: 0 });
      agg.get(k).n++;
    }
    const order = ['Guarantor only', 'Both, different payers', 'Both, same payer', 'Patient only', 'Neither'];
    const rows = [...agg.values()].sort((a, b) => order.indexOf(a.Theme) - order.indexOf(b.Theme) || b.n - a.n || a['NextGen Payer'].localeCompare(b['NextGen Payer']))
      .map(({ n, ...r }) => ({ ...r, 'Patients': String(n) }));
    return { patients, rows, counts };
  }

  // ---------- payer tests (v0.12, approved 2026-10-20; Off by default) ----------
  // For a NextGen payer name without a confident Sohar payer, the same policy (member ID) is verified under the most
  // similar Sohar payers (Config "Payer test payers", 3) for a few patients (Config "Payer test patients", 2): ids
  // <key>-T1.. on their own file. A Sohar payer that comes back eligible for every patient tried (at least the patient
  // count) is suggested on the Payer Learning list; a person confirms it. Results never become the patient's benefits.
  const PAYER_TEST_DEFAULTS = { on: false, payers: 3, patients: 2 };
  const PAYER_TEST_COLS = ['Test Id', 'MRN', 'NextGen Payer', 'Candidate payerId', 'Candidate Payer', 'First Built', 'Status', 'Sohar Date', 'Sohar Ref'];
  const PAYER_TEST_CONFIG_ROWS = [
    { setting: 'Payer mapping discovery', value: 'On', what: 'On or Off. On: discovery patients whose NextGen payer has no confident Sohar payer, but who have a member ID, go on a separate Payer Mapping discovery file that includes the member ID (one discovery each). Off: they go on the regular discovery file.' },
    { setting: 'Payer tests', value: 'Off', what: 'On or Off. On: a separate verification file sends the same policy under the most similar Sohar payers, for a few patients per unconfident NextGen payer, to learn which payer sticks. Each row is a Sohar check.' },
    { setting: 'Payer test payers', value: 3, what: 'Similar Sohar payers tried per NextGen payer name.' },
    { setting: 'Payer test patients', value: 2, what: 'Patients tried per NextGen payer name. A Sohar payer eligible for this many is suggested.' },
  ];
  function payerTestConfig(cfg) {
    const c = cfg || {};
    const raw = (k) => (c[k] == null ? '' : String(c[k]).trim());
    const count = (k, d, max) => { const n = parseInt(raw(k), 10); return isNaN(n) || n < 1 ? d : Math.min(n, max); };
    return {
      on: /^(on|yes|y|true|1)$/i.test(raw('Payer tests')),
      mappingDiscovery: !/^(off|no|n|false|0)$/i.test(raw('Payer mapping discovery')),
      payers: count('Payer test payers', PAYER_TEST_DEFAULTS.payers, 9),
      patients: count('Payer test patients', PAYER_TEST_DEFAULTS.patients, 10),
    };
  }

  /**
   * Sohar payers ranked by similarity to a NextGen payer name (autoMatchPayer's scoring): exact name or alias matches
   * first, then the fuzzy score. opts.confirmed as for autoMatchPayer. Returns [{payerId, payerName, score}] (n best).
   */
  function rankPayers(name, soharPayers, n, opts) {
    const q = autoTokens(name);
    if (!q.length) return [];
    const prepared = autoPrepWith(soharPayers, opts && opts.confirmed);
    const qW = q.reduce((s, t) => s + autoWeight(t), 0);
    const qHasStrong = q.some((t) => autoWeight(t) === 1);
    const key = q.join(' ');
    const scored = [];
    for (const p of prepared) {
      if (!p.payerId) continue;
      let best = 0;
      for (const v of p.variants) {
        if (v.toks.join(' ') === key) { best = 2; break; }
        const hit = q.filter((t) => v.toks.includes(t));
        if (!hit.length || (qHasStrong && !hit.some((t) => autoWeight(t) === 1))) continue;
        const hW = hit.reduce((s, t) => s + autoWeight(t), 0), cW = v.toks.reduce((s, t) => s + autoWeight(t), 0);
        best = Math.max(best, 0.7 * (hW / qW) + 0.3 * (hW / cW));
      }
      if (best > 0) scored.push({ payerId: p.payerId, payerName: p.payerName, score: Math.round(best * 100) / 100 });
    }
    scored.sort((a, b) => b.score - a.score || a.payerName.localeCompare(b.payerName));
    return scored.slice(0, n || 3);
  }

  /** NextGen payer names worth testing: no confident Sohar payer (Auto-Low, Auto-Medium, Review, unknown label) or a Sohar payer error on record. */
  function testWorthy(confidence, payerId, payerErrors) {
    const c = str(confidence).toLowerCase();
    if (c === 'not a payer') return false;
    return payerErrors > 0 || c === 'auto-medium' || needsClassification(confidence, payerId);
  }

  /**
   * The payer test file. patients: built patients (after the skip window). ctx: {soharPayers, payerMap, previous (Payer
   * Tests rows), payerErrors (Map(lower name -> patients with a payer error)), payers, patients, specialty, when}.
   * For each test-worthy NextGen name of in-scope patients with a member ID: the patients already tried (still in the
   * report), then the soonest visits, up to ctx.patients; each tried on the ctx.payers most similar Sohar payers. A row
   * whose result came back (Status) is not sent again; a name with a Sohar payer eligible for ctx.patients patients is done.
   * Returns {rows (VERIFICATION_COLS), table (PAYER_TEST_COLS: previous rows kept, new ones added), names: [{name, done, sent}]}.
   */
  function buildPayerTests(patients, ctx) {
    const c = { ...PAYER_TEST_DEFAULTS, ...(ctx || {}) };
    const idx = payerMapIndex(c.payerMap || []);
    const confirmed = confirmedNames(c.payerMap || []);
    const prev = (c.previous || []).map((t) => ({ ...Object.fromEntries(PAYER_TEST_COLS.map((k) => [k, str(t[k])])) }));
    const byName = new Map();
    for (const t of prev) { const k = t['NextGen Payer'].toLowerCase(); if (!byName.has(k)) byName.set(k, []); byName.get(k).push(t); }
    const pool = new Map();
    for (const p of patients || []) {
      const name = str(p['Primary Payer']);
      if (p['In Scope'] !== 'Yes' || !name || !str(p['Member ID']) || isVip(name) || isSelfPay(name) || (p.x && p.x.vip)) continue;
      const h = idx.get(name.toLowerCase());
      const errs = c.payerErrors ? +(c.payerErrors.get(name.toLowerCase()) || 0) : 0;
      if (!testWorthy(h ? h.confidence : 'Unmapped', h ? h.payerId : '', errs)) continue;
      if (!pool.has(name.toLowerCase())) pool.set(name.toLowerCase(), { name, list: [] });
      pool.get(name.toLowerCase()).list.push(p);
    }
    const rows = [], table = prev.slice(), names = [];
    const day = (s) => { const d = parseDate(s); return d ? d.getTime() : Infinity; };
    for (const [k, { name, list }] of pool) {
      const tried = byName.get(k) || [];
      const elig = new Map();
      for (const t of tried) if (/^complete\.eligible/i.test(t.Status)) elig.set(t['Candidate payerId'], (elig.get(t['Candidate payerId']) || 0) + 1);
      if ([...elig.values()].some((n) => n >= c.patients)) { names.push({ name, done: true, sent: 0 }); continue; }
      const cands = rankPayers(name, c.soharPayers || [], c.payers, { confirmed });
      if (!cands.length) { names.push({ name, done: false, sent: 0, none: true }); continue; }
      const triedKeys = new Set(tried.map((t) => t.MRN));
      const chosen = list.slice().sort((a, b) => (triedKeys.has(keyOf(b)) - triedKeys.has(keyOf(a))) || day(a['Next Appt']) - day(b['Next Appt'])).slice(0, c.patients);
      let sent = 0;
      for (const p of chosen) {
        const key = keyOf(p);
        const mine = table.filter((t) => t.MRN === key);
        for (const cand of cands) {
          let t = mine.find((x) => x['Candidate payerId'] === cand.payerId && x['NextGen Payer'].toLowerCase() === k);
          if (!t) {
            const used = new Set(mine.map((x) => x['Test Id']));
            let i = 1; while (used.has(`${key}-T${i}`) && i < 9) i++;
            if (used.has(`${key}-T${i}`)) continue; // nine tests on file for this patient: no room
            t = { 'Test Id': `${key}-T${i}`, 'MRN': key, 'NextGen Payer': name, 'Candidate payerId': cand.payerId, 'Candidate Payer': cand.payerName, 'First Built': str(c.when), 'Status': '', 'Sohar Date': '', 'Sohar Ref': '' };
            table.push(t); mine.push(t);
          }
          if (t.Status) continue;
          rows.push({
            firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, memberId: p['Member ID'],
            payerId: cand.payerId, placeOfServiceCode: p.POS, specialtyCode: c.specialty || 'PSY', id: t['Test Id'],
            subscriberFirstName: p['Subscriber First'], subscriberLastName: p['Subscriber Last'], subscriberDateOfBirth: p['Subscriber DOB'],
          });
          sent++;
        }
      }
      names.push({ name, done: false, sent });
    }
    return { rows, table, names };
  }

  /** Payer Tests rows with the results that came back (mergeSoharResults payerTests): Status, Sohar Date, Sohar Ref (latest wins). */
  function recordPayerTests(table, results) {
    const by = new Map();
    for (const r of results || []) {
      const id = str(r['Test Id']);
      const prev = by.get(id);
      if (!prev || (soharTime(r['Sohar Date']) || 0) >= (soharTime(prev['Sohar Date']) || 0)) by.set(id, r);
    }
    let updated = 0;
    const out = (table || []).map((t) => {
      const r = by.get(str(t['Test Id']));
      if (!r || (str(t['Sohar Ref']) === str(r['Sohar Ref']) && str(t.Status) === str(r['Sohar Status']))) return t;
      updated++;
      return { ...t, 'Status': str(r['Sohar Status']), 'Sohar Date': str(r['Sohar Date']), 'Sohar Ref': str(r['Sohar Ref']) };
    });
    return { table: out, updated };
  }

  /**
   * Per NextGen name: each Sohar payer tried, how many patients were sent, came back, came back eligible. suggestion: the
   * one payer eligible for at least `patients` patients (two payers both doing so: no suggestion, shown as ambiguous).
   * Returns Map(lower name -> {name, candidates: [{payerId, payerName, sent, back, eligible}], suggestion, ambiguous, text}).
   */
  function payerTestSummary(table, cfg) {
    const need = (cfg && cfg.patients) || PAYER_TEST_DEFAULTS.patients;
    const g = new Map();
    for (const t of table || []) {
      const name = str(t['NextGen Payer']), k = name.toLowerCase();
      if (!name) continue;
      if (!g.has(k)) g.set(k, { name, cands: new Map() });
      const cs = g.get(k).cands, id = str(t['Candidate payerId']);
      if (!cs.has(id)) cs.set(id, { payerId: id, payerName: str(t['Candidate Payer']), sent: 0, back: 0, eligible: 0 });
      const e = cs.get(id);
      e.sent++;
      if (str(t.Status)) e.back++;
      if (/^complete\.eligible/i.test(str(t.Status))) e.eligible++;
    }
    const out = new Map();
    for (const [k, { name, cands }] of g) {
      const list = [...cands.values()].sort((a, b) => b.eligible - a.eligible || b.back - a.back || a.payerName.localeCompare(b.payerName));
      const winners = list.filter((x) => x.eligible >= need);
      const suggestion = winners.length === 1 ? { payerId: winners[0].payerId, payerName: winners[0].payerName, n: winners[0].eligible, tried: winners[0].back, source: 'test' } : null;
      const text = list.map((x) => `${x.payerName || x.payerId}: ${x.eligible} of ${x.back} eligible${x.sent > x.back ? ` (${x.sent - x.back} waiting)` : ''}`).join('; ');
      out.set(k, { name, candidates: list, suggestion, ambiguous: winners.length > 1, text });
    }
    return out;
  }
  /** Rollup rows (learnPayers) with payer test results: r.tests, and a test suggestion where discovery gave none. */
  function withTestSuggestions(rollup, summary) {
    const s = summary || new Map();
    const rows = (rollup || []).map((r) => {
      const t = s.get(r.key);
      if (!t) return r;
      const c = str(r.confidence).toLowerCase();
      const may = !r.duplicate && (!r.mapped || c === 'auto-low' || c === 'auto-medium' || c === 'review' || r.payerErrors > 0);
      const sg = !r.suggestion && may && t.suggestion && t.suggestion.payerId !== r.payerId ? t.suggestion : r.suggestion;
      return { ...r, tests: t, suggestion: sg };
    });
    const have = new Set(rows.map((r) => r.key));
    for (const [k, t] of s) {
      if (have.has(k)) continue;
      rows.push({ key: k, nextgenPayer: t.name, mapped: false, payerId: '', payerName: '', confidence: '', notes: '', agree: 0, payerErrors: 0, discovery: 0, otherPayer: 0, unrelated: 0,
        conflicts: 0, score: 50, n: 0, top: null, decision: '', note: '', suggestion: t.suggestion, duplicate: false, duplicates: 0, lastOutcome: '', tests: t });
    }
    return rows;
  }

  // ---------- payer metrics (v0.12) ----------
  // Per NextGen payer, today: in-scope patients (and their upcoming appointments) sent to discovery because the payer
  // has no confident Sohar payer, and those whose latest verification failed; plus payer test results. The history
  // sheet keeps one row per day and payer, so a payer that keeps coming back shows as recurring.
  const PAYER_METRIC_COLS = ['NextGen Payer', 'Confidence', 'Unmapped patients', 'Unmapped appts', 'Failed verification patients', 'Failed verification appts',
    'Most common failure', 'Payer tests sent', 'Payer tests back', 'Payer tests eligible', 'Test success rate', 'Best test payer', 'Days with issues', 'Recurring'];
  const PAYER_METRIC_HISTORY_COLS = ['Date', 'NextGen Payer', 'Unmapped patients', 'Unmapped appts', 'Failed verification patients', 'Failed verification appts', 'Payer tests sent', 'Payer tests eligible'];
  const RECURRING_DAYS = 3;
  /**
   * patients: built patients; results: Sohar Results rows; tests: Payer Tests rows; history: earlier history rows.
   * Returns {rows (PAYER_METRIC_COLS, worst first), totals}.
   */
  function payerMetrics(patients, results, tests, history) {
    const byKey = resultsByKey(results);
    const g = new Map();
    const at = (name, conf) => {
      const k = str(name).toLowerCase() || '(none on file)';
      if (!g.has(k)) g.set(k, { name: str(name) || 'None on file', conf: str(conf), um: 0, ua: 0, fp: 0, fa: 0, fail: new Map(), ts: 0, tb: 0, te: 0, best: null });
      return g.get(k);
    };
    for (const p of patients || []) {
      if (!p || p['In Scope'] !== 'Yes') continue;
      const name = str(p['Primary Payer']);
      if (!name || isSelfPay(name) || isVip(name)) continue;
      const appts = (p.x && p.x.apptCount) || (str(p['Next Appt']) ? 1 : 0);
      if (p.Route === 'discovery' && needsClassification(p.Confidence, p['Sohar payerId'])) { const e = at(name, p.Confidence); e.um++; e.ua += appts; }
      const ver = (byKey.get(keyOf(p)) || []).filter((r) => r['Sent As'] === 'Verification' && resolvedResult(r) && !/^Superseded/.test(str(r.Action)) && r['Searched As'] !== 'Payer test');
      const last = latestOf(ver);
      if (last && /^(complete\.ineligible|error)/i.test(str(last['Sohar Status']))) {
        const e = at(name, p.Confidence); e.fp++; e.fa += appts;
        const why = str(last['Sohar Status']).replace(/\s*\(.*$/, '');
        e.fail.set(why, (e.fail.get(why) || 0) + 1);
      }
    }
    for (const t of payerTestSummary(tests).values()) {
      const e = at(t.name, '');
      for (const c of t.candidates) { e.ts += c.sent; e.tb += c.back; e.te += c.eligible; }
      e.best = t.suggestion || (t.candidates[0] && t.candidates[0].eligible ? t.candidates[0] : null);
    }
    const days = new Map();
    for (const h of history || []) {
      const k = str(h['NextGen Payer']).toLowerCase();
      if (!k || !(+h['Unmapped appts'] || +h['Failed verification appts'])) continue;
      if (!days.has(k)) days.set(k, new Set());
      days.get(k).add(str(h.Date));
    }
    const rows = [...g.entries()].filter(([, e]) => e.um || e.fp || e.ts).map(([k, e]) => {
      const top = [...e.fail].sort((a, b) => b[1] - a[1])[0];
      const d = days.get(k) ? days.get(k).size : 0;
      return {
        'NextGen Payer': e.name, 'Confidence': e.conf, 'Unmapped patients': String(e.um), 'Unmapped appts': String(e.ua),
        'Failed verification patients': String(e.fp), 'Failed verification appts': String(e.fa), 'Most common failure': top ? `${top[0]} (${top[1]})` : '',
        'Payer tests sent': String(e.ts), 'Payer tests back': String(e.tb), 'Payer tests eligible': String(e.te),
        'Test success rate': e.tb ? `${Math.round((100 * e.te) / e.tb)}%` : '', 'Best test payer': e.best ? `${e.best.payerName || e.best.payerId} (${e.best.payerId})` : '',
        'Days with issues': String(d), 'Recurring': d >= RECURRING_DAYS ? 'Yes' : '',
      };
    });
    rows.sort((a, b) => (+b['Unmapped appts'] + +b['Failed verification appts']) - (+a['Unmapped appts'] + +a['Failed verification appts']) || a['NextGen Payer'].localeCompare(b['NextGen Payer']));
    const sum = (c) => rows.reduce((s, r) => s + (+r[c] || 0), 0);
    return { rows, totals: { unmappedPatients: sum('Unmapped patients'), unmappedAppts: sum('Unmapped appts'), failedPatients: sum('Failed verification patients'),
      failedAppts: sum('Failed verification appts'), testsSent: sum('Payer tests sent'), testsBack: sum('Payer tests back'), testsEligible: sum('Payer tests eligible'),
      recurring: rows.filter((r) => r.Recurring).length } };
  }
  /** History with today's metric rows (one row per day and payer; today's replace earlier ones from today). */
  function appendMetricHistory(history, metricRows, date) {
    const d = str(date);
    const keep = (history || []).filter((h) => str(h.Date) !== d).map((h) => Object.fromEntries(PAYER_METRIC_HISTORY_COLS.map((c) => [c, str(h[c])])));
    const add = (metricRows || []).map((r) => ({ 'Date': d, 'NextGen Payer': r['NextGen Payer'], 'Unmapped patients': r['Unmapped patients'], 'Unmapped appts': r['Unmapped appts'],
      'Failed verification patients': r['Failed verification patients'], 'Failed verification appts': r['Failed verification appts'], 'Payer tests sent': r['Payer tests sent'], 'Payer tests eligible': r['Payer tests eligible'] }));
    return keep.concat(add);
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
    resultKey, accumulateResults, stampSentPayerIds, combineResults, keepMemberIds, TRACKER_EXPORT_COLS, serviceLine, trackerExport, TASK_COLS, TASK_TYPES, IR_COLS, HEADER_SYNONYMS,
    irCategory, parseIrPortal, irCrossReference, irMessage,
    normHeader, detectHeader, parseDate, fmtDate, fmtTime, timeMins, soharTime, excelSerialToDate, parseName,
    payerMapIndex, lookupPayer, suggestPayers, buildPatients, buildExports, toCSV, parseCSV, summarize,
    VERIFIABLE, CONFIDENT, AUTO_LABELS, isAutoLabel, needsClassification, confirmedNames, autoMatchPayer, autoTokens, AUTO_MEDIUM, finClassPayerIndex, matchVerifyRow, isVip, isSelfPay,
    LEARN_DEFAULTS, LEARNING_CONFIG_ROWS, learningConfig, mergeNotes, PAYER_OUTCOME_COLS, verificationSignal, payerOutcomes, appendOutcomes,
    learnPayers, learnFinClasses, planLearning, learningEvidence, PAYER_LEARNING_COLS, payerLearningRows, learningReviewRows,
    PAYER_MAP_LOG_COLS, mapChangeLog, sheetEditLog, DEMOTE_MARK, FLAG_ACK, flagAckNote, groupOutcomes,
    SOHAR_REF_COLS, refsOf, refIndex, appendRefs, noIdentifierReviewRows,
    REVERIFY_DEFAULTS, REVERIFY_CONFIG_ROWS, reverifyConfig, reverifyDue, policyResults, reverifyCheck,
    REVERIFY_REQUEST_COLS, pruneReverifyRequests, addReverifyRequest, applyReverifySkips, soharDay,
    PER_PREFIX, keyOf, savedKey, perKey, migrateKey, perKeyMap, splitId, searchedAs, idLabel, MAPPING_DISCOVERY_COLS, isMappingCandidate, subscriberFor,
    removeReverifyRequest, mergeClinicRequests, GUARANTOR_REASONS, GUARANTOR_REQUEST_COLS, guarantorCandidates, addGuarantorRequests, pruneGuarantorRequests,
    GUARANTOR_THEME_COLS, guarantorThemes, PAYER_TEST_DEFAULTS, PAYER_TEST_COLS, PAYER_TEST_CONFIG_ROWS, payerTestConfig, rankPayers, testWorthy, buildPayerTests,
    recordPayerTests, payerTestSummary, withTestSuggestions, PAYER_METRIC_COLS, PAYER_METRIC_HISTORY_COLS, RECURRING_DAYS, payerMetrics, appendMetricHistory, suggestionWhy,
    resultKind, readResultRows, mergeSoharResults, benefitFields, oonBenefitFields, accumulatorStatus, benefitFlag, benefitsExport, buildTasks, taskMessage, keepDone, chunkMessage, messageAssembler, workQueue, lasfir, adminMessage,
  };
});
