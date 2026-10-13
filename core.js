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
    VERIFICATION_COLS, DISCOVERY_COLS, PATIENT_COLS, REVIEW_COLS, HEADER_SYNONYMS,
    normHeader, detectHeader, parseDate, fmtDate, excelSerialToDate, parseName,
    payerMapIndex, lookupPayer, suggestPayers, buildPatients, buildExports, toCSV, summarize,
  };
});
