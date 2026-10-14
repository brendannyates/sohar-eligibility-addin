/* app.js — wires the pane to core.js (logic) and excel.js (workbook). */
(function () {
  'use strict';
  const VERSION = '0.6.0';
  const C = window.SoharCore;
  let X = null; // SoharExcel when running inside Excel; demo adapter otherwise

  const state = { rows: null, detection: null, cfg: null, payerMap: null, soharPayers: null, patients: null, review: null, exports: null, results: null, resultsAsOf: null, tasks: null, ir: null, discoveryVerify: null, pendingFocus: '', demo: false };
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
        // Still wire the buttons so each step can say what is wrong, instead of the whole pane going dead.
        setStatus(`This workbook is missing ${missing.join(', ')}. Steps that use them will fail; open the Sohar Eligibility template.`, 'bad');
        wire();
        return;
      }
      setStatus('Template workbook detected.', 'ok');
    } catch (e) { setStatus(`Could not read the workbook: ${e.message}`, 'bad'); }
    wire();
    await autoRefresh();
  }

  /**
   * On open: if the Import sheet already holds a report, scan it; if a patient list was built before, rebuild it
   * (and the Sohar imports) for today; if Sohar results were saved, match them again so next visits and tasks are current.
   */
  async function autoRefresh() {
    try {
      const rows = await X.readImportRows();
      const det = C.detectHeader(rows);
      if (!det || det.missing.length || rows.length <= det.headerRow + 1) return;
      setStatus('Refreshing from the data already in this workbook…', '');
      await loadConfig();
      await scan();
      const built = (await X.tableRowCount(X.TABLES.patients)) > 0;
      if (built && !$('btnBuild').disabled) await build();
      const saved = (await X.readUploads()).length > 0;
      if (saved) {
        $('resultsResult').innerHTML = '';
        const html = await processResults([]);
        if (html) $('resultsResult').insertAdjacentHTML('beforeend', html);
      }
      const did = [built ? 'patient list and Sohar imports rebuilt' : 'Import sheet scanned', saved ? 'saved Sohar results matched again' : ''].filter(Boolean).join('; ');
      setStatus(`Refreshed on open: ${did}.`, 'ok');
    } catch (e) {
      console.error(e);
      setStatus(`Opened, but the refresh stopped: ${e.message}`, 'bad');
    }
  }

  function startDemo(why) {
    X = demoAdapter();
    state.demo = true;
    const b = document.createElement('div');
    b.className = 'banner';
    b.textContent = `Demo mode (${why}) Synthetic data only; nothing is written anywhere.`;
    document.body.insertBefore(b, document.querySelector('main'));
    setStatus('Demo workbook', 'ok');
    wire();
    autoRefresh();
  }

  function wire() {
    $('btnScan').addEventListener('click', () => guard(scan, 'scanResult'));
    $('btnBuild').addEventListener('click', () => guard(build, 'buildResult'));
    $('btnDlVer').addEventListener('click', () => guard(() => download('verification'), 'preflight'));
    $('btnDlDisc').addEventListener('click', () => guard(() => download('discovery'), 'preflight'));
    $('btnUpload').addEventListener('click', () => $('fileImport').click());
    $('fileImport').addEventListener('change', (e) => guard(() => uploadImport(e.target), 'scanResult'));
    $('btnResults').addEventListener('click', () => $('fileResults').click());
    $('fileResults').addEventListener('change', (e) => guard(() => uploadResults(e.target), 'resultsResult'));
    $('btnCopyMsg').addEventListener('click', () => guard(() => copyAdminMessage('action'), 'resultsResult'));
    $('btnCopyBen').addEventListener('click', () => guard(() => copyAdminMessage('benefits'), 'resultsResult'));
    $('btnDlWork').addEventListener('click', () => guard(async () => saveCSV(`sohar_results_worklist_${stamp()}.csv`, C.RESULTS_COLS, (await currentResults()) || [], 'resultsResult'), 'resultsResult'));
    $('btnDlBen').addEventListener('click', () => guard(async () => saveCSV(`sohar_benefits_${stamp()}.csv`, C.BENEFITS_EXPORT_COLS, C.benefitsExport((await currentResults()) || []), 'resultsResult'), 'resultsResult'));
    $('btnDlDiscVer').addEventListener('click', () => guard(downloadDiscoveryVerify, 'resultsResult'));
    $('btnClearSaved').addEventListener('click', () => guard(clearSaved, 'resultsResult'));
    $('btnDlTasks').addEventListener('click',() => guard(async () => saveCSV(`nextgen_tasks_${stamp()}.csv`, C.TASK_COLS, await currentTasks(), 'resultsResult'), 'resultsResult'));
    $('btnFull').addEventListener('click', () => guard(() => openFull(), 'resultsResult'));
    $('btnFull2').addEventListener('click', () => guard(() => openFull('upload'), 'resultsResult'));
    $('btnIr').addEventListener('click', () => $('fileIr').click());
    $('fileIr').addEventListener('change', (e) => guard(() => uploadIr(e.target), 'irResult'));
    $('btnCopyIr').addEventListener('click', () => guard(async () => {
      const text = C.irMessage(await currentIr());
      if (await copyText(text)) note('irResult', 'IR message copied. Paste it into Outlook or Teams.', 'ok');
      else note('irResult', 'Copy is blocked here. Open full screen > Messages to copy it.', 'warn');
    }, 'irResult'));
    $('btnDlIr').addEventListener('click', () => guard(async () => saveCSV(`ir_portal_cross_reference_${stamp()}.csv`, C.IR_COLS, await currentIr(), 'irResult'), 'irResult'));
    diagnostics();
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
  async function uploadImport(input) {
    const file = input.files && input.files[0];
    input.value = ''; // so choosing the same file again still fires
    if (!file) return;
    busy('btnUpload', 'Loading…');
    try {
      if (isCSV(file)) await X.writeImportRows(C.parseCSV(await readText(file)));
      else if (isXlsx(file)) await X.importXlsx(await readBase64(file));
      else throw new Error(`${file.name} is not .csv or .xlsx. In Excel, use File > Save As to change the type, then upload again.`);
    } finally { unbusy('btnUpload', 'Upload report file'); }
    await scan();
  }

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

  async function download(kind) {
    if (!state.exports) throw new Error('Build the patient list first (step 3).');
    const cols = kind === 'verification' ? C.VERIFICATION_COLS : C.DISCOVERY_COLS;
    await saveCSV(`sohar_${kind}_import_${state.cfg.scope}_${stamp()}.csv`, cols, state.exports[kind], 'preflight');
  }

  // ---------- step 5 ----------
  async function uploadResults(input) {
    const files = [...(input.files || [])];
    input.value = '';
    if (!files.length) return;
    busy('btnResults', 'Matching…');
    $('resultsResult').hidden = false; $('resultsResult').innerHTML = '<ol class="progress"></ol>';
    progress(`Reading ${files.length} file${files.length > 1 ? 's' : ''}…`);
    try {
      const list = [];
      for (const f of files) {
        if (isCSV(f)) list.push({ name: f.name, rows: C.parseCSV(await readText(f)) });
        else if (isXlsx(f)) list.push({ name: f.name, rows: await X.readXlsxRows(await readBase64(f)) });
        else list.push({ name: f.name, rows: [] });
      }
      const html = await processResults(list);
      $('resultsResult').insertAdjacentHTML('beforeend', html);
    } finally { unbusy('btnResults', 'Upload Sohar results'); }
  }

  /** Match Sohar files to the Import report, write Sohar Results + NextGen Tasks, return a summary (HTML). Shared by pane and full screen. */
  async function processResults(list, onStep) {
    const step = (m) => { progress(m); if (onStep) onStep(m); };
    // Say what each file was recognised as, so a wrong or unreadable file is obvious.
    for (const f of list) {
      const { kind, records } = C.readResultRows(f.rows);
      step(`${f.name}: ${kind ? `${kind} output, ${records.length} rows` : 'not recognised as a Sohar verification or discovery output'}`);
    }
    // New files are added to the results saved in the workbook, so the verification of discovery patients
    // can be uploaded on its own later without losing the first round.
    const acc = C.accumulateResults(await X.readUploads(), list, new Date().toISOString());
    if (list.length && !acc.added && !acc.replaced) throw new Error(`None of these files look like a Sohar verification or discovery output${acc.unknownFiles.length ? ` (${acc.unknownFiles.join(', ')})` : ''}.`);
    if (!acc.files.length) return '';
    if (list.length) {
      step(`${acc.added} new result${acc.added === 1 ? '' : 's'}${acc.replaced ? `, ${acc.replaced} replaced` : ''}; ${acc.saved.length} saved in this workbook.`);
      await X.writeUploads(C.UPLOAD_COLS, acc.saved);
    }
    step('Reading the appointment report on the Import sheet…');
    state.rows = await X.readImportRows();
    const cfg = await X.readConfig();
    if ($('asOf').value) cfg.asOf = new Date($('asOf').value + 'T00:00:00Z');
    const res = C.mergeSoharResults(state.rows, acc.files, { ...cfg, payerMap: await X.readPayerMap() });
    if (res.error) throw new Error(res.error);
    step(`Matched ${res.counts.matched} results to patients (${res.counts.unmatched} unmatched). Writing the Sohar Results sheet…`);
    await X.writeResults(C.RESULTS_COLS, res.results);
    const tasks = C.keepDone(C.buildTasks(res.results), await X.readTasks());
    step(`Writing ${tasks.length} rows to the NextGen Tasks sheet…`);
    await X.writeTasks(C.TASK_COLS, tasks);
    const dv = C.keepMemberIds(res.discoveryVerification, await X.readDiscoveryVerify());
    await X.writeDiscoveryVerify(C.VERIFICATION_COLS, dv);
    state.results = res.results; state.resultsAsOf = res.asOf; state.tasks = tasks; state.discoveryVerify = dv;
    for (const id of ['btnCopyMsg', 'btnCopyBen', 'btnDlWork', 'btnDlBen', 'btnDlTasks']) $(id).disabled = false;
    $('btnDlDiscVer').disabled = !dv.length;
    markDone(5);
    const noId = dv.filter((r) => !r.memberId).length;
    const verifiedFromDisc = res.results.filter((r) => r['Discovery Follow-up'] === 'From discovery').length;
    const dvHtml = (dv.length
      ? `<p class="ok"><b>${dv.length} discovery patient${dv.length === 1 ? '' : 's'} ready to verify.</b> Click <b>Download discovery verification import</b> and upload it to Sohar as a verification batch, then bring the verification output back here; it is added to these results.</p>`
        + (noId ? `<p class="warn">${noId} of them ha${noId === 1 ? 's' : 've'} no member ID (Sohar discovery usually leaves it blank). Open the Sohar link, type it in the memberId column on the Discovery Verify sheet, then download.</p>` : '')
      : '')
      + (verifiedFromDisc ? `<p class="ok">${verifiedFromDisc} plan${verifiedFromDisc === 1 ? '' : 's'} found by discovery ${verifiedFromDisc === 1 ? 'has' : 'have'} been verified; benefits are on those rows.</p>` : '');

    const q = res.counts.byQueue, c = res.counts;
    const n = (col, v) => res.results.filter((r) => r[col] === v).length;
    const open = tasks.filter((t) => !t.Done).length;
    const pctTxt = `${Math.round((cfg.approachingPct ?? 0.2) * 100)}%`;
    const html = `
        <div class="counts q4">
          <div class="count act"><div class="n">${q['Action needed'] || 0}</div><div class="l">action needed</div></div>
          <div class="count unm"><div class="n">${q.Unmatched || 0}</div><div class="l">unmatched</div></div>
          <div class="count pen"><div class="n">${q.Pending || 0}</div><div class="l">pending</div></div>
          <div class="count clr"><div class="n">${q.Clear || 0}</div><div class="l">clear</div></div>
        </div>
        <p>${c.matched} result${c.matched === 1 ? '' : 's'} matched to the appointment report${c.notReturned ? `; ${c.notReturned} upcoming patient${c.notReturned === 1 ? ' was' : 's were'} not in the Sohar files` : ''}.</p>
        <p><b>${open} NextGen task${open === 1 ? '' : 's'}</b> on the NextGen Tasks sheet${tasks.length - open ? ` (${tasks.length - open} already done)` : ''}.</p>
        ${dvHtml}
        ${c.withBenefits
          ? `<p class="ok">${c.withBenefits} verified patient${c.withBenefits === 1 ? ' has' : 's have'} benefits. Deductible: ${n('Deductible Status', 'Met')} met, ${n('Deductible Status', 'Approaching')} approaching, ${n('Deductible Status', 'None')} with no deductible. OOP max: ${n('OOP Status', 'Met')} met, ${n('OOP Status', 'Approaching')} approaching.</p>
             <p class="hint">Approaching means ${pctTxt} or less remaining (Config: Benefit approaching %).</p>`
          : `<p class="warn">These files have no deductible or out-of-pocket columns, so benefit statuses are blank. Use Sohar's full verification output, which includes them.</p>`}
        ${acc.unknownFiles.length ? `<p class="warn">Skipped (not a Sohar output): ${acc.unknownFiles.map(esc).join(', ')}</p>` : ''}`;
    return html;
  }

  async function currentDiscoveryVerify() {
    if (state.demo) return state.discoveryVerify || [];
    return X.readDiscoveryVerify(); // the sheet, so member IDs typed there go out
  }
  async function downloadDiscoveryVerify() {
    const rows = await currentDiscoveryVerify();
    if (!rows.length) throw new Error('No discovery patients are waiting for verification. Upload a discovery output with coverage found first.');
    await saveCSV(`sohar_verification_import_from_discovery_${stamp()}.csv`, C.VERIFICATION_COLS, rows, 'resultsResult');
    const noId = rows.filter((r) => !String(r.memberId || '').trim()).length;
    if (noId) note('resultsResult', `${noId} row${noId === 1 ? ' has' : 's have'} no member ID. Sohar may reject ${noId === 1 ? 'it' : 'them'}; fill the memberId column on the Discovery Verify sheet and download again.`, 'warn');
  }

  let clearArmed = 0;
  async function clearSaved() {
    if (Date.now() - clearArmed > 5000) { clearArmed = Date.now(); $('btnClearSaved').textContent = 'Click again to clear'; setTimeout(() => { $('btnClearSaved').textContent = 'Clear saved Sohar results'; }, 5000); return; }
    clearArmed = 0;
    await X.writeUploads(C.UPLOAD_COLS, []);
    $('btnClearSaved').textContent = 'Clear saved Sohar results';
    note('resultsResult', 'Saved Sohar results cleared. The next upload starts fresh; the sheets stay as they are until then.', 'ok');
  }

  // ---------- step 6: IR Portal cross-check ----------
  async function uploadIr(input) {
    const f = input.files && input.files[0];
    input.value = '';
    if (!f) return;
    busy('btnIr', 'Matching…');
    const box = $('irResult'); box.hidden = false; box.innerHTML = '';
    try {
      let rows;
      if (isCSV(f)) rows = C.parseCSV(await readText(f));
      else if (isXlsx(f)) rows = await X.readXlsxRows(await readBase64(f), /^filtered/i);
      else throw new Error(`${f.name} is not .csv or .xlsx.`);
      const parsed = C.parseIrPortal(rows);
      if (parsed.error) throw new Error(parsed.error);
      const results = (await currentResults()) || [];
      if (!results.length) throw new Error('Upload the Sohar results first (step 5); the IR items are matched against them.');
      const ir = C.irCrossReference(parsed.records, results, await currentTasks());
      await X.writeIr(C.IR_COLS, ir);
      state.ir = ir;
      const by = {}; for (const r of ir) by[r.Category] = (by[r.Category] || 0) + 1;
      const insurance = parsed.records.filter((r) => C.irCategory(r.desc)).length;
      box.innerHTML = `<p class="ok">${parsed.records.length} IR Portal rows read; ${insurance} are insurance-related, ${ir.length} of those have a Sohar result and are on the IR Cross-Ref sheet.</p>
        <p>${Object.entries(by).map(([k, v]) => `${esc(k)}: ${v}`).join(' · ') || 'No insurance-related items with a Sohar result.'}</p>`;
      $('btnCopyIr').disabled = false; $('btnDlIr').disabled = false;
      markDone(6);
      if (dlg) await sendToDialog(await payload());
    } finally { unbusy('btnIr', 'Upload IR Portal file'); }
  }
  async function currentIr() {
    if (state.ir) return state.ir;
    return state.demo ? [] : X.readIr();
  }

  async function currentResults() {
    if (state.results) return state.results;
    if (state.demo) return null;
    try { return await X.readTable('tblSoharResults'); } catch (e) { return null; }
  }
  async function currentTasks() {
    if (state.tasks) return state.tasks;
    if (state.demo) return [];
    return X.readTasks();
  }

  async function copyAdminMessage(mode) {
    const rows = await currentResults();
    if (!rows || !rows.length) throw new Error('Upload Sohar results first.');
    const text = C.adminMessage(rows, { asOf: state.resultsAsOf, mode });
    if (await copyText(text)) note('resultsResult', `${mode === 'benefits' ? 'Benefits' : 'Action'} message copied. Paste it into Outlook or Teams.`, 'ok');
    else { note('resultsResult', 'Copy is blocked in this pane. Opening full screen, where you can copy it.', 'warn'); await openFull('messages'); }
  }

  // ---------- full screen ----------
  // Floats over the sheet (displayInIframe). Excel for Windows/Mac always shows dialogs as their own window.
  let dlg = null;
  const take = C.messageAssembler();
  const isWeb = () => !state.demo && Office.context.platform === Office.PlatformType.OfficeOnline;

  async function payload(focus) {
    return {
      type: 'data', focus: focus || '', web: isWeb(), results: (await currentResults()) || [], tasks: await currentTasks(), ir: await currentIr(),
      exports: state.exports || null, discoveryVerify: await currentDiscoveryVerify(), scope: state.cfg ? state.cfg.scope : 'upcoming',
      asOf: state.resultsAsOf ? state.resultsAsOf.toISOString() : null,
    };
  }
  async function sendToDialog(obj) {
    if (!dlg) return;
    for (const piece of C.chunkMessage(obj)) dlg.messageChild(piece);
  }

  async function openFull(focus) {
    const url = new URL('dashboard.html', location.href).href;
    if (state.demo) { window.open(url + (focus ? `#${focus}` : ''), '_blank'); return; }
    if (!Office.context.requirements.isSetSupported('DialogApi', '1.2')) throw new Error('This version of Excel cannot open the full-screen window (needs DialogApi 1.2). Use the Sohar Results and NextGen Tasks sheets.');
    if (dlg) { await sendToDialog(await payload(focus)); return; }
    state.pendingFocus = focus || '';
    Office.context.ui.displayDialogAsync(url, { height: 90, width: 90, displayInIframe: true }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        note('resultsResult', `Could not open full screen: ${r.error.message}. If a pop-up blocker appeared, allow pop-ups for Excel and try again.`, 'bad');
        return;
      }
      dlg = r.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => guard(() => fromDialog(arg.message), 'resultsResult'));
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { dlg = null; });
    });
  }

  async function fromDialog(raw) {
    const m = take(raw);
    if (!m) return;
    if (m.type === 'ready') { await sendToDialog(await payload(state.pendingFocus)); state.pendingFocus = ''; }
    else if (m.type === 'close') { if (dlg) dlg.close(); dlg = null; }
    else if (m.type === 'task-done') {
      await X.setTaskDone(m.key, m.done);
      if (state.tasks) state.tasks = state.tasks.map((t) => (t['Task Key'] === m.key ? { ...t, Done: m.done ? 'Yes' : '' } : t));
    } else if (m.type === 'export') {
      const msg = await saveCSV(m.name, m.cols, m.rows, 'resultsResult');
      await sendToDialog({ type: 'notice', ok: true, text: msg || `Saved ${m.name}.` });
    } else if (m.type === 'upload') {
      // Files picked in the full-screen window: CSVs arrive as text, .xlsx as base64 for Excel to open.
      try {
        const list = [];
        for (const f of m.files || []) {
          if (f.text != null) list.push({ name: f.name, rows: C.parseCSV(f.text) });
          else if (f.base64) list.push({ name: f.name, rows: await X.readXlsxRows(f.base64) });
        }
        $('resultsResult').innerHTML = ''; const html = await processResults(list, (s) => sendToDialog({ type: 'notice', ok: true, text: s }));
        $('resultsResult').insertAdjacentHTML('beforeend', html);
        await sendToDialog({ ...(await payload('prep')), notice: { ok: true, text: `Loaded ${list.map((f) => f.name).join(', ')}.` } });
      } catch (e) {
        await sendToDialog({ type: 'notice', ok: false, text: e.message });
        throw e;
      }
    }
  }

  // Excel on the web runs this pane (and the full-screen view) in a sandboxed frame that blocks downloads.
  // There the rows go to an Export sheet instead, and Excel's own File > Export > Download as CSV saves them.
  async function saveCSV(name, cols, rows, boxId) {
    if (isWeb()) {
      const n = await X.writeExport(cols, rows);
      const msg = `${n} rows are on the Export sheet. In Excel: File > Export > Download as CSV, and name it ${name}.`;
      if (boxId) note(boxId, msg, 'ok');
      return msg;
    }
    const csv = C.toCSV(cols, rows);
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
  function busy(id, text) { const b = $(id); b.dataset.label = b.textContent; b.textContent = text; b.disabled = true; }
  function unbusy(id, text) { const b = $(id); b.textContent = text || b.dataset.label; b.disabled = false; }
  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function stamp() { return new Date().toISOString().slice(0, 10).replace(/-/g, ''); }
  const isCSV = (f) => /\.csv$/i.test(f.name);
  const isXlsx = (f) => /\.xls[xm]$/i.test(f.name);
  async function readText(file) {
    // NextGen and Excel CSVs are often Windows-1252; fall back when the bytes are not valid UTF-8.
    const buf = await file.arrayBuffer();
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { return new TextDecoder('windows-1252').decode(buf); }
  }
  async function readBase64(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; } catch (e) { /* fall through */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.setAttribute('readonly', ''); ta.className = 'offscreen';
    document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    return ok;
  }
  function isoToUs(iso) { const [y, m, d] = iso.split('-'); return `${m}/${d}/${y}`; }
  function label(f) {
    return { patName: 'Patient name', dob: 'DOB', perNbr: 'Per Nbr', mrn: 'MRN', apptDt: 'Appt date', begTm: 'Start time', event: 'Event', loc: 'Location', sts: 'Status', patSts: 'Patient status', payer: 'Payer', secPayer: 'Secondary payer', polNbr: 'Member ID', polNbrLstEnc: 'Member ID (last enc)', zip: 'ZIP', state: 'State', guarName: 'Guarantor', rendering: 'Rendering' }[f] || f;
  }
  /** Run fn; show any error next to the step that failed (boxId) as well as in the header. */
  async function guard(fn, boxId) {
    try { await fn(); }
    catch (e) {
      console.error(e);
      const msg = e && e.message ? e.message : String(e);
      setStatus(`Something went wrong: ${msg}`, 'bad');
      if (boxId) note(boxId, `Something went wrong: ${msg}${e && e.code ? ` (${e.code})` : ''}`, 'bad');
    }
  }
  function progress(msg) {
    const box = $('resultsResult'); box.hidden = false;
    let ol = box.querySelector('ol.progress');
    if (!ol) { ol = document.createElement('ol'); ol.className = 'progress'; box.prepend(ol); }
    const li = document.createElement('li'); li.textContent = msg; ol.appendChild(li);
  }
  function note(boxId, msg, cls) {
    const box = $(boxId); if (!box) return;
    box.hidden = false;
    const p = document.createElement('p'); p.className = cls || ''; p.textContent = msg;
    box.prepend(p);
  }
  // Footer line so problems can be reported precisely: which Excel, and which API sets it supports.
  function diagnostics() {
    const el = $('diag'); if (!el) return;
    if (state.demo) { el.textContent = 'Demo mode (browser preview).'; return; }
    const sup = (set, v) => `${set} ${v} ${Office.context.requirements.isSetSupported(set, v) ? 'yes' : 'no'}`;
    const plat = { OfficeOnline: 'Excel on the web', PC: 'Excel for Windows', Mac: 'Excel for Mac' }[Office.context.platform] || String(Office.context.platform);
    el.textContent = `${plat} · ${sup('ExcelApi', '1.13')} · ${sup('DialogApi', '1.2')}`;
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
      async writeImportRows(r) { rows = r; },
      async importXlsx() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async readXlsxRows() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async writeResults() {}, async writeTasks() {}, async readTasks() { return []; }, async readUploads() { return []; }, async writeUploads() {}, async readDiscoveryVerify() { return []; }, async writeDiscoveryVerify() {}, async tableRowCount() { return 0; }, async setTaskDone() {}, async writeIr() {}, async readIr() { return []; }, async writeExport(c, r) { return r.length; },
      async readConfig() { return { asOf: null, scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, defaultState: '', teleKeywords: ['Tele'], posTele: '10', posOffice: '11', specialty: 'PSY', useGuarantorAsSubscriber: false }; },
      async readPayerMap() { return payerMap; },
      async readSoharPayers() { return sohar; },
      async upsertPayerMap(e) { payerMap = payerMap.filter((m) => m.nextgenPayerName.toLowerCase() !== e.nextgenPayerName.toLowerCase()).concat([e]); },
      async replaceTable() {}, async setConfigValue() {}, async activate() {}, async templateCheck() { return []; },
    };
  }
})();
