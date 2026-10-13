/* app.js — wires the pane to core.js (logic) and excel.js (workbook). */
(function () {
  'use strict';
  const VERSION = '0.1.0';
  const C = window.SoharCore;
  let X = null; // SoharExcel when running inside Excel; demo adapter otherwise

  const state = { rows: null, detection: null, cfg: null, payerMap: null, soharPayers: null, patients: null, review: null, exports: null };
  const $ = (id) => document.getElementById(id);

  // ---------- boot ----------
  document.getElementById('version').textContent = `v${VERSION}`;
  if (window.Office && Office.onReady) {
    Office.onReady((info) => { if (info.host === Office.HostType.Excel) startExcel(); else startDemo('Not running inside Excel.'); });
  } else {
    startDemo('Office.js did not load.');
  }

  async function startExcel() {
    X = window.SoharExcel;
    try {
      const missing = await X.templateCheck();
      if (missing.length) {
        setStatus(`This workbook is missing ${missing.join(', ')}. Open the Sohar Eligibility template.`, 'bad');
        disableAll();
        return;
      }
      setStatus('Template workbook detected.', 'ok');
    } catch (e) { setStatus(`Could not read the workbook: ${e.message}`, 'bad'); }
    wire();
  }

  function startDemo(why) {
    X = demoAdapter();
    const b = document.createElement('div');
    b.className = 'banner';
    b.textContent = `Demo mode (${why}) Synthetic data only; nothing is written anywhere.`;
    document.body.insertBefore(b, document.querySelector('main'));
    setStatus('Demo workbook', 'ok');
    wire();
  }

  function wire() {
    $('btnScan').addEventListener('click', () => guard(scan));
    $('btnBuild').addEventListener('click', () => guard(build));
    $('btnDlVer').addEventListener('click', () => download('verification'));
    $('btnDlDisc').addEventListener('click', () => download('discovery'));
    document.querySelectorAll('[data-goto]').forEach((b) => b.addEventListener('click', () => X.activate(b.dataset.goto)));
    $('scope').addEventListener('change', () => X.setConfigValue('Scope', $('scope').value === 'all' ? 'All patients' : 'Upcoming appointments').catch(() => {}));
    $('asOf').addEventListener('change', () => X.setConfigValue('As-of date', $('asOf').value ? isoToUs($('asOf').value) : '').catch(() => {}));
    setCurrent(1);
    loadConfig();
  }

  async function loadConfig() {
    try {
      state.cfg = await X.readConfig();
      $('scope').value = state.cfg.scope;
      $('asOf').value = state.cfg.asOf ? state.cfg.asOf.toISOString().slice(0, 10) : '';
    } catch (e) { setStatus(e.message, 'bad'); }
  }

  // ---------- step 1 ----------
  async function scan() {
    busy('btnScan', 'Scanning…');
    try {
      state.rows = await X.readImportRows();
      const det = C.detectHeader(state.rows);
      const out = $('scanResult'); out.hidden = false;
      if (!det) {
        out.innerHTML = `<p class="bad">No header row found. The Import sheet needs the NextGen columns, including <b>Per Nbr</b> and <b>Appt Dt</b>.</p>`;
        state.detection = null; $('btnBuild').disabled = true; return;
      }
      state.detection = det;
      const dataRows = state.rows.length - det.headerRow - 1;
      const found = Object.keys(det.colMap).map((f) => `<span class="chip">${label(f)}</span>`).join('');
      const miss = det.optionalMissing.map((f) => `<span class="chip miss">${label(f)} missing</span>`).join('');
      let warn = '';
      if (det.colMap.patName == null || det.colMap.dob == null) warn += `<p class="warn">Sohar needs patient name and DOB. This report has ${det.colMap.patName == null ? 'no name column' : ''}${det.colMap.patName == null && det.colMap.dob == null ? ' and ' : ''}${det.colMap.dob == null ? 'no DOB column' : ''}; those fields will be blank until the report includes them.</p>`;
      if (det.colMap.zip == null) warn += `<p class="warn">No ZIP column. Discovery rows will have an empty zipCode.</p>`;
      out.innerHTML = `<p class="ok">Header found on row ${det.headerRow + 1}. About ${dataRows.toLocaleString()} appointment rows.</p><div class="chips">${found}${miss}</div>${warn}`;
      if (det.missing.length) { out.innerHTML += `<p class="bad">Missing required columns: ${det.missing.map(label).join(', ')}.</p>`; $('btnBuild').disabled = true; return; }
      await loadPayerStep();
      $('btnBuild').disabled = false;
      markDone(1); setCurrent(2);
    } finally { unbusy('btnScan', 'Scan Import sheet'); }
  }

  // ---------- step 2 ----------
  async function loadPayerStep() {
    state.payerMap = await X.readPayerMap();
    state.soharPayers = await X.readSoharPayers();
    renderPayers();
  }

  function payersInImport() {
    const det = state.detection; if (!det) return new Map();
    const counts = new Map();
    for (let r = det.headerRow + 1; r < state.rows.length; r++) {
      const row = state.rows[r] || [];
      const per = String(row[det.colMap.perNbr] ?? '').trim();
      const name = String(row[det.colMap.payer] ?? '').trim();
      if (!per || !name) continue;
      if (!counts.has(name)) counts.set(name, new Set());
      counts.get(name).add(per);
    }
    return counts;
  }

  function renderPayers() {
    const idx = C.payerMapIndex(state.payerMap);
    const counts = payersInImport();
    const unmapped = [...counts].filter(([n]) => !idx.has(n.toLowerCase())).sort((a, b) => b[1].size - a[1].size);
    const probable = [...counts].filter(([n]) => { const h = idx.get(n.toLowerCase()); return h && /^probable$/i.test(h.confidence); });
    const box = $('payerList'); box.hidden = false;
    if (!unmapped.length) {
      box.innerHTML = `<p class="ok">All ${counts.size} payer names in this report are mapped.</p>` +
        (probable.length ? `<p class="warn">${probable.length} are marked Probable. They will be verified under the mapped ID; confirm them on the Payer Map sheet after the first Sohar run.</p>` : '');
      markDone(2); setCurrent(3);
      return;
    }
    box.innerHTML = `<p class="warn">${unmapped.length} payer name${unmapped.length > 1 ? 's' : ''} not in the Payer Map. Unmapped payers go to discovery.</p>`;
    for (const [name, pers] of unmapped) {
      const wrap = document.createElement('div');
      wrap.className = 'payer';
      wrap.innerHTML = `<div class="name">${esc(name)} <span class="n">(${pers.size} patient${pers.size > 1 ? 's' : ''})</span></div>
        <input type="search" placeholder="Search Sohar payers" aria-label="Search Sohar payers for ${esc(name)}">
        <div class="sugg"></div>
        <div class="alt-actions"><button class="link" data-act="notpayer">Not a payer</button><button class="link" data-act="skip">Leave unmapped</button></div>`;
      const input = wrap.querySelector('input');
      const sugg = wrap.querySelector('.sugg');
      const show = (q) => {
        const list = C.suggestPayers(q || name, state.soharPayers, 6);
        sugg.innerHTML = list.length ? '' : `<p class="hint">No close match. Try another word, or ask Sohar to add this payer.</p>`;
        for (const s of list) {
          const b = document.createElement('button');
          b.className = 'small';
          b.innerHTML = `<span>${esc(s.payerName)}</span><span class="id">${esc(s.payerId)}</span>`;
          b.addEventListener('click', () => guard(async () => {
            await X.upsertPayerMap({ nextgenPayerName: name, payerId: s.payerId, payerName: s.payerName, confidence: 'Confirmed', notes: 'Mapped in the add-in' });
            await loadPayerStep();
          }));
          sugg.appendChild(b);
        }
      };
      input.addEventListener('input', () => show(input.value));
      wrap.querySelector('[data-act="notpayer"]').addEventListener('click', () => guard(async () => {
        await X.upsertPayerMap({ nextgenPayerName: name, payerId: '', payerName: '', confidence: 'Not a payer', notes: 'Marked in the add-in' });
        await loadPayerStep();
      }));
      wrap.querySelector('[data-act="skip"]').addEventListener('click', () => guard(async () => {
        await X.upsertPayerMap({ nextgenPayerName: name, payerId: '', payerName: '', confidence: 'Review', notes: 'Left unmapped in the add-in; patients go to discovery' });
        await loadPayerStep();
      }));
      show('');
      box.appendChild(wrap);
    }
    setCurrent(2);
  }

  // ---------- step 3 ----------
  async function build() {
    busy('btnBuild', 'Building…');
    try {
      state.cfg = await X.readConfig();
      state.cfg.scope = $('scope').value;
      state.cfg.asOf = $('asOf').value ? new Date($('asOf').value + 'T00:00:00Z') : null;
      state.payerMap = await X.readPayerMap();
      const res = C.buildPatients(state.rows, { ...state.cfg, payerMap: state.payerMap });
      const out = $('buildResult'); out.hidden = false;
      if (res.error) { out.innerHTML = `<p class="bad">${esc(res.error)}</p>`; return; }
      state.patients = res.patients; state.review = res.review;
      state.exports = C.buildExports(res.patients, state.cfg);
      await X.replaceTable(X.TABLES.patients, C.PATIENT_COLS, res.patients);
      await X.replaceTable(X.TABLES.review, C.REVIEW_COLS, res.review);
      await X.replaceTable(X.TABLES.verification, C.VERIFICATION_COLS, state.exports.verification);
      await X.replaceTable(X.TABLES.discovery, C.DISCOVERY_COLS, state.exports.discovery);
      await X.setConfigValue('Last build', `${C.fmtDate(new Date())} · scope: ${state.cfg.scope} · as of ${C.fmtDate(res.asOf)}`).catch(() => {});

      const s = C.summarize(res.patients);
      const total = Math.max(1, s.inScope);
      const pct = (n) => `${(100 * n / total).toFixed(1)}%`;
      const sev = { High: 0, Medium: 0, Low: 0 };
      for (const r of res.review) sev[r.Severity] = (sev[r.Severity] || 0) + 1;
      out.innerHTML = `
        <p>${s.patients.toLocaleString()} patients in the report, ${s.inScope.toLocaleString()} in scope${state.cfg.scope === 'upcoming' ? ` (appointment on or after ${C.fmtDate(res.asOf)})` : ''}.</p>
        <div class="routebar" role="img" aria-label="${s.verification} verification, ${s.discovery} discovery, ${s.exclude || 0} excluded"></div>
        <div class="counts">
          <div class="count v"><div class="n">${s.verification}</div><div class="l">verification</div></div>
          <div class="count d"><div class="n">${s.discovery}</div><div class="l">discovery</div></div>
          <div class="count x"><div class="n">${s.exclude || 0}</div><div class="l">excluded</div></div>
        </div>
        <p class="issues">Review sheet: <span class="hi">${sev.High} high</span> <span class="md">${sev.Medium} medium</span> <span class="lo">${sev.Low} low</span></p>`;
      // widths set through the CSSOM so the strict style-src policy holds
      const bar = out.querySelector('.routebar');
      for (const [cls, n] of [['v', s.verification], ['d', s.discovery], ['x', s.exclude || 0]]) {
        const seg = document.createElement('span'); seg.className = cls; seg.style.width = pct(n); bar.appendChild(seg);
      }
      preflight();
      markDone(3); setCurrent(4);
      $('btnDlVer').disabled = !state.exports.verification.length;
      $('btnDlDisc').disabled = !state.exports.discovery.length;
    } finally { unbusy('btnBuild', 'Build patient list'); }
  }

  // ---------- step 4 ----------
  function preflight() {
    const box = $('preflight'); box.hidden = false;
    const all = [...state.exports.verification, ...state.exports.discovery];
    const noName = all.filter((r) => !r.firstName || !r.lastName).length;
    const noDob = all.filter((r) => !r.dateOfBirth).length;
    const noZip = state.exports.discovery.filter((r) => !r.zipCode).length;
    const parts = [];
    if (noName) parts.push(`<p class="bad">${noName} rows have no first or last name. Sohar will reject them.</p>`);
    if (noDob) parts.push(`<p class="bad">${noDob} rows have no date of birth. Sohar will reject them.</p>`);
    if (noZip) parts.push(`<p class="warn">${noZip} discovery rows have no ZIP.</p>`);
    if (!parts.length) parts.push(`<p class="ok">All rows have name, DOB and state.</p>`);
    box.innerHTML = parts.join('');
  }

  function download(kind) {
    const cols = kind === 'verification' ? C.VERIFICATION_COLS : C.DISCOVERY_COLS;
    const csv = C.toCSV(cols, state.exports[kind]);
    const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const name = `sohar_${kind}_import_${state.cfg.scope}_${stamp}.csv`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  // ---------- helpers ----------
  function setStatus(msg, cls) { const el = $('wbStatus'); el.textContent = msg; el.className = `status ${cls || ''}`; }
  function setCurrent(n) { document.querySelectorAll('.step').forEach((s) => s.classList.toggle('current', +s.dataset.step === n)); }
  function markDone(n) { $(`step${n}`).classList.add('done'); }
  function disableAll() { document.querySelectorAll('button.primary').forEach((b) => (b.disabled = true)); }
  function busy(id, text) { const b = $(id); b.dataset.label = b.textContent; b.textContent = text; b.disabled = true; }
  function unbusy(id, text) { const b = $(id); b.textContent = text || b.dataset.label; b.disabled = false; }
  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function isoToUs(iso) { const [y, m, d] = iso.split('-'); return `${m}/${d}/${y}`; }
  function label(f) {
    return { patName: 'Patient name', dob: 'DOB', perNbr: 'Per Nbr', mrn: 'MRN', apptDt: 'Appt date', begTm: 'Start time', event: 'Event', loc: 'Location', sts: 'Status', patSts: 'Patient status', payer: 'Payer', secPayer: 'Secondary payer', polNbr: 'Member ID', polNbrLstEnc: 'Member ID (last enc)', zip: 'ZIP', state: 'State', guarName: 'Guarantor', rendering: 'Rendering' }[f] || f;
  }
  async function guard(fn) {
    try { await fn(); }
    catch (e) { console.error(e); setStatus(`Something went wrong: ${e.message}`, 'bad'); }
  }

  // ---------- demo adapter (browser preview, synthetic data) ----------
  function demoAdapter() {
    const hdr = ['Appointment Pre-Check Complete', 'Res Name', 'Pat Sts', 'Portal Enrollment', 'Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Details', 'Loc Name', 'Gender Identity', 'Current Gender', 'Preferred Pronoun', 'End Tm', 'Dur', 'Per Nbr', 'Birth Dt', 'Rendering', 'Sts', 'Rschd', 'Rschd Desc', 'Md Rc', 'Guar Name', 'Pol Nbr (Lst Enc)', 'Payer Name (Cht Guar)', 'Sec Payer (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Cht Bal', 'Pat Slf Bal', 'Crt Dt', 'Created By', 'Mod Dt', 'Modified By', 'No Show/Late Cancel Charged?'];
    const names = ['Rivera, Ana', 'Okafor, Chidi', 'Lindqvist, Sven', 'Haddad, Mira', 'Nakamura, Kenji', 'Dubois, Claire', 'Osei, Kwame', 'Petrov, Ilya', 'Alvarez, Sofia', 'Byrne, Niamh', 'Kowalski, Piotr', 'Mensah, Afia'];
    const payers = ['Kaiser Foundation Northern California', 'Blue Shield CA', 'Cigna', 'Anthem Blue Cross', 'Optum Health Behavioral Solutions', '', 'Self Pay', 'Noridian Northern CA', 'Imperial Health CA Medicare Adv', 'Pacific Harbor Health Plan', 'Aetna', 'UMR'];
    const events = ['Telepsych', 'TMS Treatment Unilateral', 'Med Management In-Office', 'Esketamine Follow Up', 'Med Management Telepsych'];
    const rows = [[], ['', '', '', '', '', 'Mindful Health Solutions'], [], ['', '', '', '', '', 'Appointments Listing'], ['', '', '', '', '', 'From 6/1/2026 to 12/31/2026'], [], hdr];
    const base = Date.UTC(2026, 5, 1);
    for (let p = 0; p < names.length; p++) {
      for (let k = 0; k < 3 + (p % 4); k++) {
        const dt = new Date(base + (p * 9 + k * 23) * 86400000);
        const r = new Array(hdr.length).fill('');
        r[4] = names[p]; r[5] = `${dt.getUTCMonth() + 1}/${dt.getUTCDate()}/${dt.getUTCFullYear()}`; r[7] = events[(p + k) % events.length]; r[9] = 'San Rafael';
        r[15] = String(100100 + p); r[16] = `0${(p % 9) + 1}/1${p % 9}/19${70 + p}`; r[18] = k === 0 ? 'Cancelled' : dt > new Date() ? 'Expected' : 'Kept'; r[21] = String(80100 + p);
        r[23] = payers[p] ? `ID${1000 + p}` : ''; r[24] = payers[p]; r[26] = payers[p] && payers[p] !== 'Self Pay' ? `ID${1000 + p}` : '';
        rows.push(r);
      }
    }
    rows.push(['Total Appointments (' + (rows.length - 7) + ')']);
    let payerMap = [
      { nextgenPayerName: 'Kaiser Foundation Northern California', payerId: '94135', payerName: 'Kaiser Foundation Health Plan of Northern California', confidence: 'Exact' },
      { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
      { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
      { nextgenPayerName: 'Anthem Blue Cross', payerId: '47198', payerName: 'Blue Cross of California', confidence: 'Exact' },
      { nextgenPayerName: 'Optum Health Behavioral Solutions', payerId: '87726', payerName: 'UnitedHealthcare', confidence: 'Probable' },
      { nextgenPayerName: 'Self Pay', payerId: '', payerName: '', confidence: 'Not a payer' },
      { nextgenPayerName: 'Noridian Northern CA', payerId: '01111', payerName: 'Medicare of California', confidence: 'Exact' },
      { nextgenPayerName: 'Aetna', payerId: '60054', payerName: 'Aetna', confidence: 'Exact' },
      { nextgenPayerName: 'UMR', payerId: '37237', payerName: 'UMR', confidence: 'Exact' },
    ];
    const sohar = [['94135', 'Kaiser Foundation Health Plan of Northern California'], ['94036', 'Blue Shield of California'], ['62308', 'Cigna'], ['47198', 'Blue Cross of California'], ['87726', 'UnitedHealthcare'], ['01111', 'Medicare of California'], ['60054', 'Aetna'], ['37237', 'UMR'], ['95567', 'Health Net'], ['68039', 'Western Health Advantage'], ['943226498', 'San Francisco Health Plan'], ['PHHP1', 'Pacific Harbor Health Plan']].map(([payerId, payerName]) => ({ payerId, payerName, aliasPayerNames: '' }));
    return {
      TABLES: { patients: 'tblPatients', review: 'tblReview', verification: 'tblVerification', discovery: 'tblDiscovery' },
      async readImportRows() { return rows; },
      async readConfig() { return { asOf: null, scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, defaultState: '', teleKeywords: ['Tele'], posTele: '10', posOffice: '11', specialty: 'PSY', useGuarantorAsSubscriber: false }; },
      async readPayerMap() { return payerMap; },
      async readSoharPayers() { return sohar; },
      async upsertPayerMap(e) { payerMap = payerMap.filter((m) => m.nextgenPayerName.toLowerCase() !== e.nextgenPayerName.toLowerCase()).concat([e]); },
      async replaceTable() {}, async setConfigValue() {}, async activate() {}, async templateCheck() { return []; },
    };
  }
})();
