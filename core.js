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
  const RESULTS_COLS = ['Next Appt', 'Next Time', 'Next Event', 'Next Location', 'Next Provider', 'Upcoming Appts',
    'Per Nbr', 'Last Name', 'First Name', 'DOB', 'Sent As', 'Sohar Status', 'Work Queue', 'Action',
    'Deductible', 'Deductible Remaining', 'Deductible Met', 'OOP Max', 'OOP Remaining', 'OOP Met', 'Copay', 'Coinsurance',
    'NextGen Payer', 'NextGen Member ID', 'Sohar Payer', 'Sohar payerId', 'Sohar Member ID', 'Coverage Differs', 'Match',
    'Sohar Ref', 'Sohar Link', 'Sohar Date'];

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
    if (h.has('verificationid') && h.has('status')) return 'verification';
    if (h.has('requestid') && (h.has('responsepayerid') || h.has('responsepayername'))) return 'discovery';
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

  // Benefit amounts, if a Sohar export carries them. Matched loosely on normalized headers; family and
  // out-of-network columns are skipped so the numbers are the patient's in-network individual accumulators.
  const money = (v) => { if (isBlank(v)) return null; const n = parseFloat(String(v).replace(/[$,\s]/g, '')); return isNaN(n) ? null : n; };
  function benefitFields(rec) {
    const keys = Object.keys(rec).filter((k) => !/family|outofnetwork|oon/.test(k));
    const find = (test) => { const k = keys.find((x) => test(x) && !isBlank(rec[x])); return k ? rec[k] : ''; };
    const oop = (k) => k.includes('outofpocket') || k.includes('oop');
    const rem = (k) => k.includes('remain');
    const met = (k) => /met$|met[^a-z]|metamount/.test(k) || k.endsWith('met');
    const ded = {
      total: find((k) => k.includes('deductible') && !rem(k) && !met(k)),
      remaining: find((k) => k.includes('deductible') && rem(k)),
      met: find((k) => k.includes('deductible') && met(k)),
    };
    const oo = {
      total: find((k) => oop(k) && !rem(k) && !met(k)),
      remaining: find((k) => oop(k) && rem(k)),
      met: find((k) => oop(k) && met(k)),
    };
    const isMet = (b) => {
      if (/^(true|yes|y)$/i.test(str(b.met))) return 'Yes';
      if (/^(false|no|n)$/i.test(str(b.met))) return 'No';
      const r = money(b.remaining), t = money(b.total), m = money(b.met);
      if (r != null) return r <= 0 ? 'Yes' : 'No';
      if (t != null && m != null) return m >= t ? 'Yes' : 'No';
      if (t === 0) return 'Yes';
      return '';
    };
    return {
      'Deductible': str(ded.total), 'Deductible Remaining': str(ded.remaining), 'Deductible Met': isMet(ded),
      'OOP Max': str(oo.total), 'OOP Remaining': str(oo.remaining), 'OOP Met': isMet(oo),
      'Copay': str(find((k) => k.includes('copay'))), 'Coinsurance': str(find((k) => k.includes('coinsurance'))),
    };
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

  /** Plain-text note for admins, grouped by location then appointment. LASFIR + Per Nbr only (minimum necessary). */
  function adminMessage(rows, opts) {
    const o = opts || {};
    const todo = rows.filter((r) => r['Work Queue'] === 'Action needed' || r['Work Queue'] === 'Unmatched');
    if (!todo.length) return 'Sohar eligibility: nothing needs action in this list.';
    const byLoc = new Map();
    for (const r of todo) {
      const loc = r['Next Location'] || 'No upcoming appointment';
      if (!byLoc.has(loc)) byLoc.set(loc, []);
      byLoc.get(loc).push(r);
    }
    const lines = [`Sohar eligibility follow-up${o.asOf ? ` (appointments from ${fmtDate(o.asOf)})` : ''}: ${todo.length} patient${todo.length > 1 ? 's' : ''} need action before their visit.`, ''];
    for (const [loc, list] of [...byLoc].sort((a, b) => a[0].localeCompare(b[0]))) {
      lines.push(`${loc} (${list.length})`);
      for (const r of list) {
        const when = r['Next Appt'] ? `${r['Next Appt'].slice(0, 5)}${r['Next Time'] ? ' ' + r['Next Time'] : ''}` : 'no appt';
        const ben = [r['Deductible Met'] === 'Yes' ? 'deductible met' : '', r['OOP Met'] === 'Yes' ? 'OOP max met' : ''].filter(Boolean).join(', ');
        lines.push(`- ${when} · ${lasfir(r['Last Name'], r['First Name'])} · Per ${r['Per Nbr'] || '?'} · ${r.Action}${ben ? ` (${ben})` : ''}`);
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
   * importRows: Import sheet 2D array. files: [{name, rows}] (2D, header first). cfg: as for buildPatients.
   * Returns {results, unknownFiles, counts} or {error}.
   */
  function mergeSoharResults(importRows, files, cfg) {
    const built = buildPatients(importRows, { ...cfg, scope: 'all' });
    if (built.error) return { error: `The Import sheet has to hold the appointment report to combine results. ${built.error}` };
    const asOf = built.asOf;
    const det = built.detection, cm = det.colMap;
    const get = (row, f) => (cm[f] == null ? '' : row[cm[f]]);

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
    const counts = { matched: 0, unmatched: 0, notReturned: 0, withBenefits: 0, byAction: {}, byQueue: {} };
    const BENEFIT_KEYS = ['Deductible', 'Deductible Remaining', 'OOP Max', 'OOP Remaining', 'Copay', 'Coinsurance'];

    const emit = (p, kind, f) => {
      const ups = p ? upcoming.get(p['Per Nbr']) || [] : [];
      const n = ups[0];
      const row = {
        'Next Appt': n ? fmtDate(n.dt) : '', 'Next Time': n ? fmtTime(n.tmRaw) : '', 'Next Event': n ? n.event : '',
        'Next Location': n ? n.loc : '', 'Next Provider': n ? n.rendering : '', 'Upcoming Appts': p ? ups.length : '',
        'Per Nbr': p ? p['Per Nbr'] : '', 'Last Name': p ? p['Last Name'] : str(f.lastName), 'First Name': p ? p['First Name'] : str(f.firstName),
        'DOB': p ? p.DOB : f.dob || '', 'Sent As': kind === 'verification' ? 'Verification' : 'Discovery',
        'NextGen Payer': p ? p['Primary Payer'] : '', 'NextGen Member ID': p ? p['Member ID'] : '',
        ...(f.benefits || {}), ...f.cols,
      };
      for (const c of RESULTS_COLS) if (row[c] == null) row[c] = '';
      row['Work Queue'] = workQueue(row);
      results.push(row);
      counts.byAction[row.Action] = (counts.byAction[row.Action] || 0) + 1;
      counts.byQueue[row['Work Queue']] = (counts.byQueue[row['Work Queue']] || 0) + 1;
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
      const { kind, records } = readResultRows(file.rows);
      if (!kind) { unknownFiles.push(file.name); continue; }
      kinds.add(kind);
      for (const r of records) {
        let p = null, match = '';
        const first = str(r.firstname), last = str(r.lastname);
        const dob = parseDate(r.dateofbirth);
        const cols = {};
        if (kind === 'verification') {
          p = pickOne(byMember.get(idKey(r.memberid)), last);
          if (p) match = 'Member ID';
          else { p = pickOne(byName.get(nameKey(first, last)), last); if (p) match = 'Name only. Confirm.'; }
          Object.assign(cols, {
            'Sohar Status': str(r.status), 'Action': verificationAction(r.status),
            'Sohar Payer': str(r.payername), 'Sohar payerId': '', 'Sohar Member ID': str(r.memberid), 'Coverage Differs': '',
            'Sohar Ref': str(r.verificationid), 'Sohar Link': '', 'Sohar Date': str(r.created),
          });
        } else {
          p = byPer.get(str(r.id)) || null;
          if (p) match = 'Per Nbr';
          else {
            p = pickOne(byNameDob.get(`${nameKey(first, last)}|${fmtDate(dob)}`), last);
            if (p) match = 'Name + DOB';
          }
          const found = !isBlank(r.responsepayerid) || !isBlank(r.responsepayername);
          const resolved = !isBlank(r.resolved);
          const status = found ? 'Coverage found' : resolved ? 'No coverage found' : 'Pending';
          const differs = found && p && (str(r.responsepayerid) !== p['Sohar payerId'] || (!isBlank(r.memberid) && idKey(r.memberid) !== idKey(p['Member ID'])));
          let action;
          if (!found) action = resolved ? 'No coverage found. Confirm self-pay or ask the patient for a card.' : 'Pending at Sohar. Recheck later.';
          else if (p && !p['Primary Payer']) action = 'Coverage found. Add it to NextGen, then verify.';
          else if (differs) action = 'Different coverage found. Confirm with the patient, update NextGen, then verify.';
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
        const benefits = benefitFields(r);
        if (BENEFIT_KEYS.some((k) => benefits[k])) counts.withBenefits++;
        emit(p, kind, { firstName: first, lastName: last, dob: fmtDate(dob), cols, benefits });
      }
    }

    // Patients who would have been sent for an uploaded file type but came back with nothing.
    for (const p of built.patients) {
      const kind = p.Route;
      if (!kinds.has(kind) || seen[kind].has(p['Per Nbr']) || !upcoming.has(p['Per Nbr'])) continue;
      counts.notReturned++;
      emit(p, kind, { cols: {
        'Sohar Status': 'Not in Sohar file', 'Action': 'Not in the results. Check it was uploaded to Sohar.',
        'Sohar Payer': '', 'Sohar payerId': '', 'Sohar Member ID': '', 'Coverage Differs': '', 'Match': '', 'Sohar Ref': '', 'Sohar Link': '', 'Sohar Date': '',
      } });
    }

    // Soonest appointment first; rows with no upcoming appointment last.
    const t = (r) => (r['Next Appt'] ? parseDate(r['Next Appt']).getTime() : Infinity);
    results.sort((a, b) => t(a) - t(b) || a['Last Name'].localeCompare(b['Last Name']));
    return { results, unknownFiles, counts, asOf };
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
    VERIFICATION_COLS, DISCOVERY_COLS, PATIENT_COLS, REVIEW_COLS, RESULTS_COLS, HEADER_SYNONYMS,
    normHeader, detectHeader, parseDate, fmtDate, fmtTime, excelSerialToDate, parseName,
    payerMapIndex, lookupPayer, suggestPayers, buildPatients, buildExports, toCSV, parseCSV, summarize,
    resultKind, readResultRows, mergeSoharResults, benefitFields, workQueue, lasfir, adminMessage,
  };
});
