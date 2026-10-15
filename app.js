/* app.js — wires the pane to core.js (logic) and excel.js (workbook). */
(function () {
  'use strict';
  const VERSION = '0.7.0';
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
      setStatus('Loading what is already in this workbook…', '');
      await loadConfig();
      await refreshAll();
    } catch (e) {
      console.error(e);
      setStatus(`Opened, but loading stopped: ${e.message}`, 'bad');
    }
  }

  /** Scan the Import sheet, build the patient list and Sohar imports, then re-match saved Sohar and IR uploads. */
  async function refreshAll() {
    await scan();
    if ($('btnBuild').disabled) return;
    await build();
    const saved = await X.readUploads();
    let did = 'patient list and Sohar imports built';
    if (saved.length) {
      $('resultsResult').innerHTML = '';
      const html = await processResults([]);
      if (html) $('resultsResult').insertAdjacentHTML('beforeend', html); foldProgress();
      did += `; ${saved.length} saved Sohar result${saved.length === 1 ? '' : 's'} matched`;
      const ir = await X.readIrUploads();
      if (ir.length) { await runIr(ir); did += '; IR cross-check re-run'; }
    } else if (!state.demo && await X.tableRowCount('tblSoharResults').catch(() => 0)) {
      // Results written by an older version, before uploads were saved in the workbook: they cannot be re-matched.
      note('resultsResult', 'The Sohar Results sheet has results from an earlier version of this add-in, but the Sohar files themselves were not saved in this workbook. Drop the Sohar outputs here again so follow-ups and benefits are current.', 'warn');
    }
    const last = saved.map((s) => String(s.Uploaded || '')).sort().pop();
    await stampUpdated(`${did}${last ? `. Last Sohar upload ${fmtStamp(new Date(last))}` : ''}`);
  }

  /** Record when the workbook was last brought up to date (Config "Last updated") and say so in the header. */
  async function stampUpdated(what) {
    const now = fmtStamp(new Date());
    await X.setConfigValue('Last updated', now).catch(() => {});
    setStatus(`Updated ${now}: ${what}.`, 'ok');
  }
  function fmtStamp(d) {
    if (isNaN(d)) return '';
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
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
    $('btnScan').addEventListener('click', () => guard(refreshAll, 'scanResult'));
    $('btnBuild').addEventListener('click', () => guard(build, 'buildResult'));
    $('btnDlVer').addEventListener('click', () => guard(() => download('verification'), 'preflight'));
    $('btnDlDisc').addEventListener('click', () => guard(() => download('discovery'), 'preflight'));
    dropZone('dropReport', 'fileImport', (files) => guard(() => uploadImport(files), 'scanResult'));
    dropZone('dropResults', 'fileResults', (files) => guard(() => uploadResults(files), 'resultsResult'));
    $('btnCopyMsg').addEventListener('click', () => guard(() => copyAdminMessage('action'), 'resultsResult'));
    $('btnCopyBen').addEventListener('click', () => guard(() => copyAdminMessage('benefits'), 'benBox'));
    $('btnDlWork').addEventListener('click', () => guard(async () => saveCSV(C.fileName('worklist'), C.RESULTS_COLS, (await currentResults()) || [], 'resultsResult'), 'resultsResult'));
    $('btnDlBen').addEventListener('click', () => guard(async () => saveCSV(C.fileName('benefits'), C.BENEFITS_EXPORT_COLS, C.benefitsExport((await currentResults()) || []), 'benBox'), 'benBox'));
    $('btnDlDiscVer').addEventListener('click', () => guard(downloadDiscoveryVerify, 'resultsResult'));
    $('btnDlFinal').addEventListener('click', () => guard(downloadTrackerFile, 'benBox'));
    $('btnClearSaved').addEventListener('click', () => guard(clearSaved, 'resultsResult'));
    $('btnDlTasks').addEventListener('click',() => guard(async () => saveCSV(C.fileName('tasks'), C.TASK_COLS, await currentTasks(), 'resultsResult'), 'resultsResult'));
    $('btnFull').addEventListener('click', () => guard(() => openFull(), 'resultsResult'));
    $('btnIr').addEventListener('click', () => $('fileIr').click());
    $('fileIr').addEventListener('change', (e) => guard(() => uploadIr(e.target), 'irResult'));
    $('btnCopyIr').addEventListener('click', () => guard(async () => {
      const text = C.irMessage(await currentIr());
      if (await copyText(text)) note('irResult', 'IR message copied. Paste it into Outlook or Teams.', 'ok');
      else note('irResult', 'Copy is blocked here. Open full screen > Messages to copy it.', 'warn');
    }, 'irResult'));
    $('btnDlIr').addEventListener('click', () => guard(async () => saveCSV(C.fileName('ir'), C.IR_COLS, await currentIr(), 'irResult'), 'irResult'));
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
  async function uploadImport(files) {
    const file = files[0];
    if (!file) return;
    if (!isCSV(file) && !isXlsx(file)) throw new Error(`${file.name} is not .xlsx or .csv. Open it in Excel, File > Save As > Excel Workbook (.xlsx), then upload that.`);
    dropBusy('dropReport', `Loading ${file.name}…`);
    try {
      if (isCSV(file)) await X.writeImportRows(C.parseCSV(await readText(file)));
      else await X.importXlsx(await readBase64(file));
      await refreshAll();
    } finally { dropBusy('dropReport'); }
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
      out.innerHTML = `<p class="ok">Report loaded: ${dataRows.toLocaleString()} appointment rows.</p>${warn}<details class="more inline"><summary>Columns found</summary><div class="chips">${found}${miss}</div></details>`;
      if (det.missing.length) { out.innerHTML += `<p class="bad">Missing required columns: ${det.missing.map(label).join(', ')}.</p>`; $('btnBuild').disabled = true; return; }
      await loadPayerStep();
      $('btnBuild').disabled = false;
    } finally { unbusy('btnScan', 'Re-read the Import sheet'); }
  }

  // ---------- step 2 ----------
  async function loadPayerStep() {
    state.payerMap = await X.readPayerMap();
    state.soharPayers = await X.readSoharPayers();
    renderPayers();
  }

  /** After a payer is mapped, rebuild so the patient moves from discovery to verification right away. */
  async function afterMap() {
    await loadPayerStep();
    if (state.patients) await build();
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
      box.hidden = !probable.length;
      box.innerHTML = probable.length ? `<p class="hint">${probable.length} payer name${probable.length > 1 ? 's are' : ' is'} mapped as Probable; confirm on the Payer Map sheet after the first Sohar run.</p>` : '';
      return;
    }
    box.innerHTML = `<p class="warn"><b>${unmapped.length} insurance name${unmapped.length > 1 ? 's' : ''} need a Sohar payer.</b> Pick one below and it is remembered. Anything left unmapped goes to discovery.</p>`;
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
            await afterMap();
          }));
          sugg.appendChild(b);
        }
      };
      input.addEventListener('input', () => show(input.value));
      wrap.querySelector('[data-act="notpayer"]').addEventListener('click', () => guard(async () => {
        await X.upsertPayerMap({ nextgenPayerName: name, payerId: '', payerName: '', confidence: 'Not a payer', notes: 'Marked in the add-in' });
        await afterMap();
      }));
      wrap.querySelector('[data-act="skip"]').addEventListener('click', () => guard(async () => {
        await X.upsertPayerMap({ nextgenPayerName: name, payerId: '', payerName: '', confidence: 'Review', notes: 'Left unmapped in the add-in; patients go to discovery' });
        await afterMap();
      }));
      show('');
      box.appendChild(wrap);
    }
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
        <p>${s.inScope.toLocaleString()} patient${s.inScope === 1 ? '' : 's'}${state.cfg.scope === 'upcoming' ? ` with a visit on or after ${C.fmtDate(res.asOf)}` : ' in the report'}.</p>
        <div class="routebar" role="img" aria-label="${s.verification} verification, ${s.discovery} discovery, ${s.exclude || 0} excluded"></div>
        <div class="counts">
          <div class="count v"><div class="n">${s.verification}</div><div class="l">verification<br>(have insurance on file)</div></div>
          <div class="count d"><div class="n">${s.discovery}</div><div class="l">discovery<br>(insurance missing or unmapped)</div></div>
          <div class="count x"><div class="n">${s.exclude || 0}</div><div class="l">left out</div></div>
        </div>
        ${sev.High ? `<p class="issues"><span class="hi">${sev.High} high-priority item${sev.High === 1 ? '' : 's'}</span> on the Review sheet.</p>` : ''}`;
      // widths set through the CSSOM so the strict style-src policy holds
      const bar = out.querySelector('.routebar');
      for (const [cls, n] of [['v', s.verification], ['d', s.discovery], ['x', s.exclude || 0]]) {
        const seg = document.createElement('span'); seg.className = cls; seg.style.width = pct(n); bar.appendChild(seg);
      }
      preflight();
      markDone(1); setCurrent(2);
      const nv = state.exports.verification.length, nd = state.exports.discovery.length;
      $('btnDlVer').disabled = !nv; $('btnDlDisc').disabled = !nd;
      $('fnVer').textContent = `${nv} patient${nv === 1 ? '' : 's'} · ${C.fileName('verification')}`;
      $('fnDisc').textContent = `${nd} patient${nd === 1 ? '' : 's'} · ${C.fileName('discovery')}`;
    } finally { unbusy('btnBuild', 'Rebuild with these options'); }
  }

  // ---------- step 2: save the Sohar imports ----------
  function preflight() {
    const box = $('preflight'); box.hidden = false;
    const all = [...state.exports.verification, ...state.exports.discovery];
    const noName = all.filter((r) => !r.firstName || !r.lastName).length;
    const noDob = all.filter((r) => !r.dateOfBirth).length;
    const noZip = state.exports.discovery.filter((r) => !r.zipCode).length;
    const parts = [];
    if (noName) parts.push(`<p class="bad">${noName} rows have no first or last name. Sohar will reject them.</p>`);
    if (noDob) parts.push(`<p class="bad">${noDob} rows have no date of birth. Sohar will reject them.</p>`);
    if (noZip) parts.push(`<p class="hint">${noZip} discovery row${noZip === 1 ? ' has' : 's have'} no ZIP (the NextGen report has none). Sohar still searches by name, DOB and state.</p>`);
    box.hidden = !parts.length;
    box.innerHTML = parts.join('');
  }

  async function download(kind) {
    if (!state.exports) throw new Error('Upload the NextGen report first (step 1).');
    const cols = kind === 'verification' ? C.VERIFICATION_COLS : C.DISCOVERY_COLS;
    await saveCSV(C.fileName(kind), cols, state.exports[kind], 'preflight');
    markDone(2); setCurrent(3);
  }

  // ---------- step 3: Sohar results ----------
  async function uploadResults(files) {
    files = [...files];
    if (!files.length) return;
    dropBusy('dropResults', `Matching ${files.length} file${files.length > 1 ? 's' : ''}…`);
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
      foldProgress();
    } finally { dropBusy('dropResults'); }
  }
  /** Tuck the file-by-file progress lines under "What was read" once the summary is up. */
  function foldProgress() {
    const ol = $('resultsResult').querySelector('ol.progress');
    if (!ol || ol.parentElement.tagName === 'DETAILS') return;
    const d = document.createElement('details'); d.className = 'more inline';
    d.innerHTML = '<summary>What was read</summary>';
    ol.replaceWith(d); d.appendChild(ol);
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
    for (const id of ['btnDlWork', 'btnDlBen', 'btnDlTasks']) $(id).disabled = false;
    $('fixActions').hidden = false;
    markDone(3); setCurrent(4);

    // Discovery found coverage -> those patients need a verification run before benefits exist.
    const noId = dv.filter((r) => !r.memberId).length;
    $('discVerBox').hidden = !dv.length;
    if (dv.length) {
      $('discVerText').innerHTML = `<b>Discovery found insurance for ${dv.length} patient${dv.length === 1 ? '' : 's'}.</b> Benefits need one more Sohar run: save this verification file, upload it to Sohar, then drop Sohar's output above.`
        + (noId ? ` <span class="warn">${noId} ha${noId === 1 ? 's' : 've'} no member ID yet: type ${noId === 1 ? 'it' : 'them'} in the memberId column on the <b>Discovery Verify</b> sheet first.</span>` : '');
      $('fnDiscVer').textContent = `${dv.length} patient${dv.length === 1 ? '' : 's'} · ${C.fileName('discoveryVerify')}`;
    }

    const q = res.counts.byQueue, c = res.counts;
    const open = tasks.filter((t) => !t.Done).length;
    const kinds = { verification: res.results.filter((r) => r['Sent As'] === 'Verification').length, discovery: res.results.filter((r) => r['Sent As'] === 'Discovery').length, waiting: dv.length };
    const html = `
        <p>${c.matched} Sohar result${c.matched === 1 ? '' : 's'} matched to patients${c.notReturned ? `; ${c.notReturned} upcoming patient${c.notReturned === 1 ? ' was' : 's were'} not in the Sohar files` : ''}.</p>
        <div class="counts q4">
          <div class="count act"><div class="n">${q['Action needed'] || 0}</div><div class="l">need action</div></div>
          <div class="count unm"><div class="n">${q.Unmatched || 0}</div><div class="l">unmatched</div></div>
          <div class="count pen"><div class="n">${q.Pending || 0}</div><div class="l">pending</div></div>
          <div class="count clr"><div class="n">${q.Clear || 0}</div><div class="l">clear</div></div>
        </div>
        ${open ? `<p><b>${open} NextGen fix${open === 1 ? '' : 'es'}</b> to make before visits (NextGen Tasks sheet).</p>` : ''}
        ${acc.unknownFiles.length ? `<p class="warn">Skipped (not a Sohar output): ${acc.unknownFiles.map(esc).join(', ')}</p>` : ''}`;
    await renderBenefits(res.results, cfg, kinds);
    if (list.length) {
      // New Sohar results change the IR cross-check, so re-run it from the saved IR items.
      const ir = await X.readIrUploads();
      if (ir.length) { step('Re-running the IR Portal cross-check…'); await runIr(ir); }
      await stampUpdated(`${list.map((f) => f.name).join(', ')} added to the Sohar results`);
    }
    return html;
  }

  // ---------- step 4: benefits ----------
  const usd = (v) => { const n = parseFloat(String(v ?? '').replace(/[$,]/g, '')); return isNaN(n) ? '' : `$${n.toLocaleString('en-US', { maximumFractionDigits: 0 })}`; };

  /** Write the Benefits sheet and show OOP / deductible progress, with the patients at or near their OOP max listed. */
  async function renderBenefits(results, cfg, kinds) {
    const view = C.benefitsView(results);
    const s = C.benefitsSummary(view);
    const box = $('benBox');
    const pctTxt = `${Math.round((cfg.approachingPct ?? 0.2) * 100)}%`;
    await X.writeBenefits(C.BENEFIT_SHEET_COLS, view.rows, C.BENEFIT_PCT_COLS);
    state.benefits = view;
    $('benActions').hidden = !s.verified;
    if (!s.verified) {
      box.innerHTML = !kinds.verification
        ? `<p class="warn"><b>No benefits yet: only discovery results are loaded.</b> Discovery finds the insurance; copay, deductible and OOP max come from a <b>verification</b> run.</p>
           <p class="hint">${kinds.waiting ? `Save the verification file in step 3 for the ${kinds.waiting} patient${kinds.waiting === 1 ? '' : 's'} discovery found, run it in Sohar, and drop the output in step 3. ` : ''}Patients sent for verification in step 2 show here once that output is uploaded too.</p>`
        : `<p class="warn">${kinds.verification} verification result${kinds.verification === 1 ? '' : 's'} loaded, but none has benefit amounts. Upload Sohar's full verification output (the one with innOutOfPocket and innTotalDeductible columns).</p>`;
      return;
    }
    const tile = (n, l, cls) => `<div class="count ${cls}"><div class="n">${n}</div><div class="l">${l}</div></div>`;
    const bar = (pct) => `<span class="meter" data-pct="${pct === '' ? 0 : pct}"><span></span></span>`;
    const rows = s.watch.slice(0, 25).map((r) => `<tr>
        <td><b>${esc(C.lasfir(r['Last Name'], r['First Name']))}</b><br><span class="hint">Per ${esc(r['Per Nbr'])}${r.Service ? ` · ${esc(r.Service)}` : ''}</span></td>
        <td>${esc(String(r['Next Appt'] || '').slice(0, 5)) || '<span class="hint">no visit</span>'}</td>
        <td>${r['OOP Status'] === 'Met' ? '<b class="okc">Met</b>' : `${esc(usd(r['OOP Remaining']))} left`}<br><span class="hint">of ${esc(usd(r['OOP Max']))}</span></td>
        <td class="meterc">${bar(r['OOP % Met'])}<span class="hint">${r['OOP % Met'] === '' ? '' : Math.round(r['OOP % Met'] * 100) + '%'}</span></td>
      </tr>`).join('');
    box.innerHTML = `
      <p>${s.verified} verified patient${s.verified === 1 ? '' : 's'} on the <b>Benefits</b> sheet, closest to OOP max first.</p>
      <div class="counts q4">
        ${tile(s.oopMet, 'OOP max met', 'clr')}${tile(s.oopApproaching, 'OOP max approaching', 'unm')}
        ${tile(s.dedMet, 'deductible met', 'clr')}${tile(s.dedApproaching, 'deductible approaching', 'unm')}
      </div>
      <h3 class="sub">At or near OOP max</h3>
      ${s.watch.length
        ? `<table class="watch"><thead><tr><th>Patient</th><th>Next</th><th>OOP</th><th>Met</th></tr></thead><tbody>${rows}</tbody></table>
           ${s.watch.length > 25 ? `<p class="hint">${s.watch.length - 25} more on the Benefits sheet.</p>` : ''}`
        : `<p class="hint">No one is within ${pctTxt} of their OOP max.</p>`}
      <p class="hint">Approaching = ${pctTxt} or less left (Config: Benefit approaching %).${s.verified - s.withAmounts ? ` ${s.verified - s.withAmounts} verified patient${s.verified - s.withAmounts === 1 ? ' has' : 's have'} no amounts from Sohar.` : ''}${s.pending ? ` ${s.pending} patient${s.pending === 1 ? ' is' : 's are'} still waiting on the step 3 verification and not counted.` : ''}</p>`;
    // widths set through the CSSOM so the strict style-src policy holds
    box.querySelectorAll('.meter').forEach((m) => { m.firstElementChild.style.width = `${Math.round(Math.min(1, +m.dataset.pct) * 100)}%`; });
  }

  async function currentDiscoveryVerify() {
    if (state.demo) return state.discoveryVerify || [];
    return X.readDiscoveryVerify(); // the sheet, so member IDs typed there go out
  }
  async function downloadDiscoveryVerify() {
    const rows = await currentDiscoveryVerify();
    if (!rows.length) throw new Error('No discovery patients are waiting for verification. Upload a discovery output with coverage found first.');
    await saveCSV(C.fileName('discoveryVerify'), C.VERIFICATION_COLS, rows, 'resultsResult');
    const noId = rows.filter((r) => !String(r.memberId || '').trim()).length;
    if (noId) note('resultsResult', `${noId} row${noId === 1 ? ' has' : 's have'} no member ID. Sohar may reject ${noId === 1 ? 'it' : 'them'}; fill the memberId column on the Discovery Verify sheet and download again.`, 'warn');
  }

  /** Final file for the TMS Tracker and ESK Tracker: verified benefits; patients still waiting on a discovery verification are listed, not included. */
  async function downloadTrackerFile() {
    const rows = (await currentResults()) || [];
    if (!rows.length) throw new Error('Upload Sohar results first.');
    const out = C.trackerExport(rows, fmtStamp(new Date()));
    if (!out.rows.length) throw new Error('No verified patients yet. The final file needs at least one Sohar verification result.');
    await saveCSV(C.fileName('tracker'), C.TRACKER_EXPORT_COLS, out.rows, 'benBox');
    const n = `${out.rows.length} patient${out.rows.length === 1 ? '' : 's'}`;
    if (out.pending.length) {
      const who = out.pending.map((r) => `${C.lasfir(r['Last Name'], r['First Name'])} · Per ${r['Per Nbr']}`).join(', ');
      note('benBox', `${n} in the file. Left out, still waiting on the step 3 verification: ${who}. Save the tracker file again after that output is uploaded.`, 'warn');
    } else note('benBox', `${n} in the tracker file. Import it in the TMS Tracker and the ESK Tracker.`, 'ok');
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
      // Only insurance-related items are kept, so the next open can re-run the cross-check without the file.
      const insurance = parsed.records.filter((r) => C.irCategory(r.desc));
      await X.writeIrUploads(insurance);
      await runIr(insurance, `${parsed.records.length} IR Portal rows read; ${insurance.length} are insurance-related. `);
      await stampUpdated(`IR Portal file ${f.name} cross-checked`);
    } finally { unbusy('btnIr', 'Upload IR Portal file'); }
  }

  /** Cross-check saved IR items against the current Sohar results; writes IR Cross-Ref. */
  async function runIr(records, lead) {
    const ir = C.irCrossReference(records, (await currentResults()) || [], await currentTasks());
    await X.writeIr(C.IR_COLS, ir);
    state.ir = ir;
    const by = {}; for (const r of ir) by[r.Category] = (by[r.Category] || 0) + 1;
    const box = $('irResult'); box.hidden = false;
    box.innerHTML = `<p class="ok">${esc(lead || `${records.length} saved insurance-related IR items. `)}${ir.length} have a Sohar result and are on the IR Cross-Ref sheet.</p>
      <p>${Object.entries(by).map(([k, v]) => `${esc(k)}: ${v}`).join(' · ') || 'No insurance-related items with a Sohar result.'}</p>`;
    $('btnCopyIr').disabled = false; $('btnDlIr').disabled = false;
    if (dlg) await sendToDialog(await payload());
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
  // Floats over the Excel page (displayInIframe: true). Never a separate window.

  // Demo only: full screen as a frame over the pane, with the pane blurred behind it.
  function showFloat(src) {
    hideFloat();
    const bd = document.createElement('div'); bd.className = 'backdrop'; bd.id = 'floatBackdrop'; bd.addEventListener('click', hideFloat);
    const box = document.createElement('div'); box.className = 'float'; box.id = 'floatBox';
    box.setAttribute('role', 'dialog'); box.setAttribute('aria-modal', 'true'); box.setAttribute('aria-label', 'Full screen');
    box.innerHTML = '<div class="floatbar"><button class="small" type="button">Close</button></div><iframe title="Sohar eligibility, full screen"></iframe>';
    box.querySelector('iframe').src = src;
    box.querySelector('button').addEventListener('click', hideFloat);
    document.body.append(bd, box);
    box.querySelector('button').focus();
  }
  function hideFloat() { ['floatBackdrop', 'floatBox'].forEach((id) => { const el = $(id); if (el) el.remove(); }); }
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') hideFloat(); });
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
    if (state.demo) { showFloat(url + (focus ? `#${focus}` : '')); return; } // browser preview: the same float, drawn in the page
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
        $('resultsResult').insertAdjacentHTML('beforeend', html); foldProgress();
        await sendToDialog({ ...(await payload('prep')), notice: { ok: true, text: `Loaded ${list.map((f) => f.name).join(', ')}.` } });
      } catch (e) {
        await sendToDialog({ type: 'notice', ok: false, text: e.message });
        throw e;
      }
    }
  }

  // Saves straight to Downloads. Excel on the web may block downloads inside the add-in frame, so there the note
  // also offers a save tab (save.html, same site, top level) and, as a last resort, the Export sheet.
  async function saveCSV(name, cols, rows, boxId) {
    const csv = C.toCSV(cols, rows);
    downloadNow(name, csv);
    const msg = `Saved ${name} (${rows.length} row${rows.length === 1 ? '' : 's'}) to your Downloads folder.`;
    if (!boxId) return msg;
    if (!isWeb()) { note(boxId, msg, 'ok'); return msg; }
    const p = note(boxId, `${msg} Nothing there? `, 'ok');
    const tab = document.createElement('button'); tab.className = 'link'; tab.textContent = 'Save in a new tab';
    tab.addEventListener('click', () => saveInTab(name, csv, boxId));
    const sheet = document.createElement('button'); sheet.className = 'link'; sheet.textContent = 'put it on the Export sheet';
    sheet.addEventListener('click', () => guard(async () => {
      await X.writeExport(cols, rows);
      note(boxId, `Rows are on the Export sheet. File > Export > Download as CSV, and name it ${name}.`, 'ok');
    }, boxId));
    p.append(tab, ' or ', sheet, '.');
    return msg;
  }
  function downloadNow(name, csv) {
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }
  /** Open save.html and hand it the file by postMessage until it answers. Must run inside the click (pop-up rules). */
  function saveInTab(name, csv, boxId) {
    const w = window.open(new URL('save.html', location.href).href, '_blank');
    if (!w) { note(boxId, 'The browser blocked the new tab. Allow pop-ups for Excel, or use the Export sheet.', 'bad'); return; }
    let got = false, tries = 0;
    const onMsg = (e) => { if (e.origin === location.origin && e.data && e.data.type === 'sohar-save-got') { got = true; window.removeEventListener('message', onMsg); } };
    window.addEventListener('message', onMsg);
    const send = () => {
      if (got || ++tries > 40 || w.closed) { window.removeEventListener('message', onMsg); return; }
      try { w.postMessage({ type: 'sohar-save', name, csv }, location.origin); } catch (e) { /* not loaded yet */ }
      setTimeout(send, 250);
    };
    send();
  }

  // ---------- helpers ----------
  function setStatus(msg, cls) { const el = $('wbStatus'); el.textContent = msg; el.className = `status ${cls || ''}`; }
  function setCurrent(n) { document.querySelectorAll('.step').forEach((s) => s.classList.toggle('current', +s.dataset.step === n)); }
  function markDone(n) { $(`step${n}`).classList.add('done'); }
  function busy(id, text) { const b = $(id); b.dataset.label = b.textContent; b.textContent = text; b.disabled = true; }
  function unbusy(id, text) { const b = $(id); b.textContent = text || b.dataset.label; b.disabled = false; }
  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
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
    const box = $(boxId); if (!box) return null;
    box.hidden = false;
    const p = document.createElement('p'); p.className = cls || ''; p.textContent = msg;
    box.prepend(p);
    return p;
  }
  /** A drop zone that also opens the file picker on click, Enter or Space. */
  function dropZone(id, inputId, onFiles) {
    const z = $(id), input = $(inputId);
    z.addEventListener('click', () => { if (!z.classList.contains('busy')) input.click(); });
    z.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); z.click(); } });
    input.addEventListener('change', () => { const f = [...input.files]; input.value = ''; onFiles(f); });
    z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('over'); });
    z.addEventListener('dragleave', () => z.classList.remove('over'));
    z.addEventListener('drop', (e) => { e.preventDefault(); z.classList.remove('over'); if (!z.classList.contains('busy')) onFiles([...e.dataTransfer.files]); });
  }
  function dropBusy(id, text) {
    const z = $(id);
    if (text) { z.dataset.html = z.dataset.html || z.innerHTML; z.textContent = text; z.classList.add('busy'); }
    else { if (z.dataset.html) z.innerHTML = z.dataset.html; z.classList.remove('busy'); }
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
    // Synthetic Sohar results so the benefits step has something to show. Add ?fresh to the URL to start empty.
    const ver = (id, payer, ded, dedRem, oop, oopRem, copay) => ({ requestid: `ver-${id}`, id: String(id), status: 'complete.eligible', payername: payer, planname: 'SYNTH PPO', networktype: 'PPO', inncopay: copay, inncoinsurance: '0.2', inntotaldeductible: ded, innremainingdeductible: dedRem, innoutofpocket: oop, innoutofpocketremaining: oopRem, created: '2026-10-14T14:05:00.000Z' });
    const demoUploads = /fresh/.test(location.search) ? [] : [
      ver(100100, 'Kaiser Foundation Health Plan of Northern California', '1500', '0', '5000', '400', '40'),
      ver(100101, 'Blue Shield of California', '2000', '300', '6000', '5100', '30'),
      ver(100102, 'Cigna', '0', '0', '3000', '0', '0'),
      ver(100103, 'Blue Cross of California', '1000', '1000', '4000', '3800', '50'),
      ver(100104, 'UnitedHealthcare', '500', '0', '2500', '200', '25'),
      { requestid: 'disc-100106', id: '100106', responsepayerid: '95567', responsepayername: 'Health Net', created: '2026-10-14T14:05:00.000Z' },
    ].map((rec) => { const kind = rec.responsepayerid ? 'discovery' : 'verification'; return { Kind: kind, Key: C.resultKey(kind, rec), File: 'demo', Uploaded: '2026-10-14T14:10:00.000Z', Record: JSON.stringify(rec) }; });
    return {
      TABLES: { patients: 'tblPatients', review: 'tblReview', verification: 'tblVerification', discovery: 'tblDiscovery' },
      async readImportRows() { return rows; },
      async writeImportRows(r) { rows = r; },
      async importXlsx() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async readXlsxRows() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async writeResults() {}, async writeTasks() {}, async writeBenefits() {}, async readTasks() { return []; }, async readUploads() { return demoUploads; }, async readIrUploads() { return []; }, async writeIrUploads() {}, async writeUploads() {}, async readDiscoveryVerify() { return []; }, async writeDiscoveryVerify() {}, async tableRowCount() { return 0; }, async setTaskDone() {}, async writeIr() {}, async readIr() { return []; }, async writeExport(c, r) { return r.length; },
      async readConfig() { return { asOf: null, scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, defaultState: '', teleKeywords: ['Tele'], posTele: '10', posOffice: '11', specialty: 'PSY', useGuarantorAsSubscriber: false }; },
      async readPayerMap() { return payerMap; },
      async readSoharPayers() { return sohar; },
      async upsertPayerMap(e) { payerMap = payerMap.filter((m) => m.nextgenPayerName.toLowerCase() !== e.nextgenPayerName.toLowerCase()).concat([e]); },
      async replaceTable() {}, async setConfigValue() {}, async activate() {}, async templateCheck() { return []; },
    };
  }
})();
