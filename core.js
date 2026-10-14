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
    zip: ['Zip', 'Pat Zip', 'Zip Code', 'Postal Code', 'Patient Zip', 'Home Zip'],
    state: ['Pat State', 'Patient State', 'Home State', 'State'],
    guarName: ['Guar Name', 'Guarantor Name', 'Guarantor'],
    rendering: ['Rendering', 'Rendering Provider', 'Provider'],
  };
  const REQUIRED = ['perNbr', 'apptDt', 'payer'];

  function normHeader(s) {
    return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  }
  function isBlank(v) {
    return v == null || String(v).trim() === '';
  }
  function str(v) {
    return isBlank(v) ? '' : String(v).trim();
  }

  /** Find the header row in a pasted report. Returns {headerRow, colMap, headers} or null. */
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
    return best;
  }

  // ---------- dates ----------
  const MS_DAY = 86400000;
  function excelSerialToDate(n) {
    // Excel 1900 system, serial 1 = 1900-01-01; 25569 = 1970-01-01 (UTC)
    return new Date(Math.round((n - 25569) * MS_DAY));
  }
  function parseDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : utcDay(v);
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
  const VERIFIABLE = new Set(['exact', 'probable', 'confirmed']);
  function payerMapIndex(mapRows) {
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

  // ---------- build ----------
  const NOT_KEPT_FUTURE = new Set(['cancelled', 'canceled', 'no show', 'noshow', 'deleted']);

  /**
   * rows: 2D array from the Import sheet (including title rows). cfg:
   *  { asOf: Date, scope: 'upcoming'|'all', locationState: {locName: 'CA'}, defaultState, teleKeywords: ['tele'],
   *    posTele: '10', posOffice: '11', specialty: 'PSY', payerMap: [ {nextgenPayerName,payerId,payerName,confidence,notes} ],
   *    useGuarantorAsSubscriber: bool }
   */
  function buildPatients(rows, cfg) {
    const det = detectHeader(rows);
    if (!det) return { error: 'Could not find the header row. The Import sheet needs the NextGen columns, including "Per Nbr" and "Appt Dt".' };
    if (det.missing.length) return { error: `Header row found but missing: ${det.missing.join(', ')}.`, detection: det };
    const cm = det.colMap;
    const idx = payerMapIndex(cfg.payerMap);
    const asOf = cfg.asOf ? utcDay(cfg.asOf) : utcDay(new Date());
    const tele = (cfg.teleKeywords && cfg.teleKeywords.length ? cfg.teleKeywords : ['tele']).map((k) => k.toLowerCase());
    const locState = {};
    for (const [k, v] of Object.entries(cfg.locationState || {})) locState[str(k).toLowerCase()] = str(v).toUpperCase();

    const get = (row, f) => (cm[f] == null ? '' : row[cm[f]]);
    const byPatient = new Map();
    let dataRows = 0;
    for (let r = det.headerRow + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const per = str(get(row, 'perNbr'));
      const dt = parseDate(get(row, 'apptDt'));
      if (!per || !dt) continue; // footer / blank rows
      dataRows++;
      const rec = {
        dt, status: str(get(row, 'sts')), event: str(get(row, 'event')), loc: str(get(row, 'loc')),
        payer: str(get(row, 'payer')), secPayer: str(get(row, 'secPayer')),
        pol: str(get(row, 'polNbr')), polLst: str(get(row, 'polNbrLstEnc')),
        name: str(get(row, 'patName')), dob: parseDate(get(row, 'dob')), zip: str(get(row, 'zip')),
        state: str(get(row, 'state')), patSts: str(get(row, 'patSts')), mrn: str(get(row, 'mrn')), guar: str(get(row, 'guarName')),
      };
      if (!byPatient.has(per)) byPatient.set(per, []);
      byPatient.get(per).push(rec);
    }
    if (!dataRows) return { error: 'Header row found but no appointment rows under it.', detection: det };

    const patients = [];
    const review = [];
    const addReview = (sev, issue, per, detail, fix) => review.push({ Severity: sev, Issue: issue, 'Per Nbr': per, Detail: detail, Fix: fix });

    for (const [per, recs] of byPatient) {
      recs.sort((a, b) => b.dt - a.dt); // latest first
      const latest = recs[0];
      const pick = (f) => { const h = recs.find((x) => x[f]); return h ? h[f] : ''; };
      const future = recs.filter((x) => x.dt >= asOf && !NOT_KEPT_FUTURE.has(x.status.toLowerCase())).sort((a, b) => a.dt - b.dt);
      const next = future[0] || null;
      const posEvent = (next || latest).event;
      const pos = tele.some((k) => posEvent.toLowerCase().includes(k)) ? (cfg.posTele || '10') : (cfg.posOffice || '11');

      const nm = parseName(pick('name'));
      const dob = recs.map((x) => x.dob).find(Boolean) || null;
      const zip = str(pick('zip')).replace(/[^0-9-]/g, '').slice(0, 10);
      const state = (str(pick('state')) || locState[latest.loc.toLowerCase()] || cfg.defaultState || '').toUpperCase().slice(0, 2);

      const payerName = latest.payer || pick('payer');
      const memberId = latest.pol || pick('pol');
      const memberLst = latest.polLst || pick('polLst');
      const lk = lookupPayer(payerName, idx);
      const secName = latest.secPayer || pick('secPayer');
      const secLk = lookupPayer(secName, idx);

      const distinctPayers = new Set(recs.map((x) => x.payer).filter(Boolean));
      const flags = [];
      let route, reason;
      if (lk.status === 'mapped' && memberId) { route = 'verification'; reason = 'mapped payer + member ID'; }
      else if (lk.status === 'none') { route = 'discovery'; reason = 'no payer on file'; }
      else if (lk.status === 'unmapped') { route = 'discovery'; reason = `payer not in Payer Map: ${payerName}`; addReview('High', 'Unmapped payer', per, payerName, 'Map it in step 2 or add a row on the Payer Map sheet.'); }
      else if (lk.status === 'review') { route = 'discovery'; reason = `payer needs review (${lk.confidence}): ${payerName}`; if (lk.confidence.toLowerCase() === 'not a payer' && payerName.toLowerCase() !== 'self pay') { route = 'exclude'; reason = `not a payer: ${payerName}`; } }
      else { route = 'discovery'; reason = 'mapped payer but no member ID'; flags.push('No member ID'); }

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
        'Primary Payer': payerName, 'Sohar payerId': lk.payerId || '', 'Confidence': lk.confidence || '',
        'Member ID': memberId, 'Member ID (Lst Enc)': memberLst, 'Member ID Mismatch': (memberId && memberLst && memberId !== memberLst) ? 'Yes' : 'No',
        'Secondary Payer': secName, 'Secondary Sohar payerId': secLk.payerId || '',
        'Appt Count': recs.length, 'First Appt': fmtDate(recs[recs.length - 1].dt), 'Last Appt': fmtDate(latest.dt),
        'Subscriber First': subF, 'Subscriber Last': subL, 'Subscriber DOB': subD, 'Flags': flags.join('; '),
      });
    }
    patients.sort((a, b) => (a.Route > b.Route ? -1 : a.Route < b.Route ? 1 : 0) || ((b['In Scope'] === 'Yes') - (a['In Scope'] === 'Yes')) || a['Per Nbr'].localeCompare(b['Per Nbr']));

    const unmapped = {};
    for (const p of patients) {
      if (p.Route === 'discovery' && p['Route Reason'].startsWith('payer not in Payer Map')) unmapped[p['Primary Payer']] = (unmapped[p['Primary Payer']] || 0) + 1;
    }
    return { detection: det, dataRows, patients, review, unmappedPayers: unmapped, asOf };
  }

  function buildExports(patients, cfg) {
    const specialty = (cfg && cfg.specialty) || 'PSY';
    const inScope = patients.filter((p) => p['In Scope'] === 'Yes');
    const verification = inScope.filter((p) => p.Route === 'verification').map((p) => ({
      firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State, memberId: p['Member ID'],
      payerId: p['Sohar payerId'], placeOfServiceCode: p.POS, specialtyCode: specialty, id: p['Per Nbr'],
      subscriberFirstName: p['Subscriber First'], subscriberLastName: p['Subscriber Last'], subscriberDateOfBirth: p['Subscriber DOB'],
    }));
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
  function workQueue(row) {
    if (row.Match === 'No match') return 'Unmatched';
    if (/^Pending/.test(row.Action)) return 'Pending';
    if (/^Eligible\. No action|^Coverage found matches/.test(row.Action)) return 'Clear';
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
    for (let r = det.headerRow + 1; r < importRows.length; r++) {
      const row = importRows[r] || [];
      const per = str(get(row, 'perNbr')); const dt = parseDate(get(row, 'apptDt'));
      if (!per || !dt || dt < asOf || NOT_KEPT_FUTURE.has(str(get(row, 'sts')).toLowerCase())) continue;
      if (!upcoming.has(per)) upcoming.set(per, []);
      upcoming.get(per).push({ dt, tmRaw: get(row, 'begTm'), event: str(get(row, 'event')), loc: str(get(row, 'loc')), rendering: str(get(row, 'rendering')) });
    }
    for (const list of upcoming.values()) list.sort((a, b) => a.dt - b.dt || String(a.tmRaw).localeCompare(String(b.tmRaw)));

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
    for (const d of results) {
      if (d['Sent As'] !== 'Discovery' || d['Sohar Status'] !== 'Coverage found' || !d['Per Nbr'] || !d['Sohar payerId']) continue;
      const p = byPer.get(d['Per Nbr']);
      const payer = d['Sohar Payer'].toLowerCase();
      const member = idKey(d['Sohar Member ID']);
      const samePolicy = (r) => (member && idKey(r['Sohar Member ID']) ? idKey(r['Sohar Member ID']) === member : !!payer && r['Sohar Payer'].toLowerCase() === payer);
      const v = results
        .filter((r) => r['Sent As'] === 'Verification' && r['Per Nbr'] === d['Per Nbr'] && r['Sohar Status'] !== 'Not in Sohar file' && samePolicy(r))
        .sort((a, b) => String(b['Sohar Date']).localeCompare(String(a['Sohar Date'])))[0];
      if (!v) {
        d['Discovery Follow-up'] = 'Ready to verify';
        d.Action = `${d.Action} Copay, deductible and OOP need a Sohar verification: this patient is on the discovery verification import.`;
        discoveryVerification.push({
          firstName: p['First Name'], lastName: p['Last Name'], dateOfBirth: p.DOB, state: p.State,
          // Member ID from discovery; NextGen's only when NextGen already has this payer.
          memberId: d['Sohar Member ID'] || (p['Sohar payerId'] === d['Sohar payerId'] ? p['Member ID'] : ''),
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
    return { results, unknownFiles, counts, asOf, discoveryVerification };
  }

  /** Stable key for one Sohar result, so a re-upload replaces it and a new file adds to the saved set. */
  function resultKey(kind, r) {
    if (kind === 'discovery') return `d|${str(r.id) || str(r.requestid) || `${nameKey(r.firstname, r.lastname)}|${str(r.dateofbirth)}`}`;
    const who = str(r.id) || str(r.verificationid || r.requestid) || `${idKey(r.memberid)}|${nameKey(r.firstname, r.lastname)}`;
    return `v|${who}|${str(r.payername).toLowerCase()}`;
  }

  /**
   * Add newly uploaded Sohar files to the saved result records. Same key: the new upload wins.
   * saved: rows of UPLOAD_COLS. Returns {saved, files: [{name, kind, records}], added, replaced, unknownFiles}.
   */
  function accumulateResults(saved, uploads, uploadedAt) {
    const by = new Map();
    for (const s of saved || []) {
      let rec; try { rec = JSON.parse(s.Record); } catch (e) { continue; }
      by.set(String(s.Key), { Kind: str(s.Kind), Key: String(s.Key), File: str(s.File), Uploaded: str(s.Uploaded), rec });
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
      if (status.startsWith('complete.ineligible')) add(r, 'Update insurance in NextGen', `Sohar: ${r['Sohar Status']} for ${r['Sohar Payer'] || r['NextGen Payer']}. Confirm current coverage with the patient; consider discovery.`);
      else if (status === 'error.member.id') add(r, 'Fix member ID in NextGen', `Sohar rejected member ID ${r['NextGen Member ID'] || '(blank)'} for ${r['NextGen Payer']}. Check the card, then re-verify.`);
      else if (status === 'error.member.name') add(r, 'Fix name in NextGen', 'Sohar rejected the name. Match it to the insurance card, then re-verify.');
      else if (status === 'complete.eligible') {
        if (/^Coverage ends/.test(r.Action)) add(r, 'Update insurance in NextGen', r.Action);
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
      if (r['Sent As'] !== 'Verification' || !r['Per Nbr'] || !r['Sohar Ref']) continue;
      const prev = by.get(r['Per Nbr']);
      if (!prev || String(r['Sohar Date']) > String(prev['Sohar Date'])) by.set(r['Per Nbr'], r);
    }
    return [...by.values()].map((r) => Object.fromEntries(BENEFITS_EXPORT_COLS.map((c) => [c, r[c] == null ? '' : r[c]])));
  }

  // Final file for the TMS Tracker and ESK Tracker imports: the benefits columns, then these. Add columns at the end only.
  const TRACKER_EXPORT_COLS = [...BENEFITS_EXPORT_COLS, 'Service', 'Source', 'Exported'];

  /** "TMS", "ESK", "TMS; ESK" or '' from a patient's upcoming NextGen events. */
  function serviceLine(text) {
    const s = String(text || '');
    return [/\btms\b|magv|\bmt\b/i.test(s) ? 'TMS' : '', /esketamine|spravato|\besk\b/i.test(s) ? 'ESK' : ''].filter(Boolean).join('; ');
  }

  /**
   * One row per patient with a Sohar verification, for the tracker imports. Patients whose discovered coverage
   * is still waiting on a verification are left out and returned as pending, so their benefits are not stale.
   */
  function trackerExport(results, exportedAt) {
    const pendingPer = new Set((results || []).filter((r) => r['Discovery Follow-up'] === 'Ready to verify' && r['Per Nbr']).map((r) => r['Per Nbr']));
    const events = new Map();
    for (const r of results || []) if (r['Per Nbr']) events.set(r['Per Nbr'], `${events.get(r['Per Nbr']) || ''} ${r['Next Event']} ${r['Upcoming Detail']}`);
    const by = new Map();
    for (const r of results || []) {
      if (r['Sent As'] !== 'Verification' || !r['Per Nbr'] || !r['Sohar Ref'] || pendingPer.has(r['Per Nbr'])) continue;
      const prev = by.get(r['Per Nbr']);
      if (!prev || String(r['Sohar Date']) > String(prev['Sohar Date'])) by.set(r['Per Nbr'], r);
    }
    const rows = [...by.values()].map((r) => ({
      ...Object.fromEntries(BENEFITS_EXPORT_COLS.map((c) => [c, r[c] == null ? '' : r[c]])),
      'Service': serviceLine(events.get(r['Per Nbr'])),
      'Source': r['Discovery Follow-up'] === 'From discovery' ? 'Discovery, then verification' : 'Verification',
      'Exported': exportedAt || '',
    }));
    const pending = (results || []).filter((r) => r['Discovery Follow-up'] === 'Ready to verify' && r['Per Nbr']);
    return { rows, pending };
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
    resultKey, accumulateResults, keepMemberIds, TRACKER_EXPORT_COLS, serviceLine, trackerExport, TASK_COLS, TASK_TYPES, IR_COLS, HEADER_SYNONYMS,
    irCategory, parseIrPortal, irCrossReference, irMessage,
    normHeader, detectHeader, parseDate, fmtDate, fmtTime, excelSerialToDate, parseName,
    payerMapIndex, lookupPayer, suggestPayers, buildPatients, buildExports, toCSV, parseCSV, summarize,
    resultKind, readResultRows, mergeSoharResults, benefitFields, accumulatorStatus, benefitFlag, benefitsExport, buildTasks, taskMessage, keepDone, chunkMessage, messageAssembler, workQueue, lasfir, adminMessage,
  };
});
