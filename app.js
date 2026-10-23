/* app.js — wires the pane to core.js (logic) and excel.js (workbook). */
(function () {
  'use strict';
  const VERSION = '0.14.0';
  const ADDIN_ID = '9786ae8a-f258-4f2c-9627-4c6f0f25db7e';
  const C = window.SoharCore;
  const R = window.SoharRules;
  const XL = window.SoharXlsx;
  const F = window.SoharFees; // fee estimates (v0.11, a test, Off by default)
  const PRM = window.PayerRules; // esketamine payer credential rules (v0.13, payer-rules.js, shared with other add-ins)
  let X = null; // SoharExcel when running inside Excel; demo adapter otherwise

  const state = { rows: null, detection: null, cfg: null, payerMap: null, soharPayers: null, patients: null, review: null, exports: null, results: null, resultsAsOf: null, tasks: null, ir: null, discoveryVerify: null, pendingFocus: '', demo: false,
    flags: new Map(), clinic: null, evidence: null, clinicFile: false, clinicFilter: '',
    // outcomes / finClasses start null ("not read yet"): build reads the ledger once, and retries after a read error.
    outcomes: null, learning: null, finClasses: null, learnRowsChecked: false,
    // Fee estimates (v0.11): null while Off; the fee sheets are read only when Config "Fee estimates" is On.
    fees: null,
    // v0.12: Per Nbr -> patient key (MRN) from the last build, for converting rows saved before v0.12; skip window
    // lists; guarantor requests (active Map) and candidates; payer tests; metrics; guarantor themes.
    perToKey: new Map(), skipped: [], requestedAnyway: [], guarantorActive: new Map(), guarantorCandidates: [], payerTests: [], metrics: null, themes: null,
    // v0.14: Clinic Notes / Clinic Payer Notes rows brought back from Clinic Eligibility files.
    clinicNotes: [], clinicPayerNotes: [] };
  const $ = (id) => document.getElementById(id);

  // ---------- boot ----------
  document.getElementById('version').textContent = `v${VERSION}`;
  // Header ⟳: a clinic file re-opens from its saved tables; otherwise re-read and rebuild, same as on open.
  $('btnRefresh').addEventListener('click', async (e) => {
    const b = e.currentTarget;
    if (!X || state.clinicFile) { location.reload(); return; }
    b.disabled = true;
    await guard(async () => { setStatus('Refreshing from this workbook…', ''); await loadConfig(); await refreshAll();
      if ($('btnBuild').disabled) setStatus('Refreshed. There is no usable report on the Import sheet yet; see step 1.', 'bad');
    }, 'scanResult');
    b.disabled = false;
  });
  if (window.Office && Office.onReady) {
    Office.onReady((info) => { if (info.host === Office.HostType.Excel) startExcel(); else startDemo('Not running inside Excel.'); });
  } else {
    startDemo('Office.js did not load.');
  }

  async function startExcel() {
    X = window.SoharExcel;
    wireTabs();
    wireClinic();
    // A clinic file (saved from the Admin tab) has tblClinicMeta: open the Clinic tab and read it; no template steps.
    try {
      const cf = await X.readClinicFile();
      if (cf) { startClinicFile(cf); return; }
    } catch (e) { console.error(e); }
    showPaneTab('admin');
    try {
      const missing = await X.templateCheck();
      if (missing.length) {
        // Still wire the buttons so each step can say what is wrong, instead of the whole pane going dead.
        setStatus(`This workbook is missing ${missing.join(', ')}. Steps that use them will fail; open the Sohar Eligibility template.`, 'bad');
        wire();
        return;
      }
      setStatus('Template workbook detected.', 'ok');
      // A save cut off last time (Excel closed, network drop): finish it before anything reads those sheets.
      const fixed = await X.recoverAll().catch((e) => { console.error(e); return []; });
      if (fixed.length) {
        state.recoveredNote = `An earlier save of ${fixed.join(', ')} was interrupted; the add-in finished it. Nothing was lost.`;
        setStatus(state.recoveredNote, 'warn');
      }
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
      if (!det || det.missing.length || det.dataEnd <= det.headerRow + 1) { await loadInbox(); return; }
      setStatus('Loading what is already in this workbook…', '');
      await loadConfig();
      await refreshAll();
    } catch (e) {
      console.error(e);
      setStatus(`Opened, but loading stopped: ${e.message}`, 'bad');
      await loadInbox(); // its own errors stay in its box
    }
  }

  /** Scan the Import sheet, build the patient list and Sohar imports, then re-match saved Sohar and IR uploads. */
  async function refreshAll() {
    if (state.fees) state.fees.loaded = false; // fee estimates (On only): pick up edits made on the fee sheets
    await loadPayerRules(); // esketamine payer credential rules: pick up edits made on the sheet
    await loadClinicNotes();
    await scan();
    if ($('btnBuild').disabled) { await loadInbox(); return; }
    await build({ deferRules: true });
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
    await refreshRules();
    await loadInbox(); // after the tasks are built: its count is a dry run of the merge against them
    const last = saved.map((s) => String(s.Uploaded || '')).sort().pop();
    await stampUpdated(`${did}${last ? `. Last Sohar upload ${fmtStamp(new Date(last))}` : ''}`);
  }

  /** Record when the workbook was last brought up to date (Config "Last updated") and say so in the header. */
  async function stampUpdated(what) {
    const now = fmtStamp(new Date());
    await X.setConfigValue('Last updated', now).catch(() => {});
    // Saves finished by a read or save since open (an earlier one was cut off) are reported once, then cleared.
    const more = X.takeRecovered ? X.takeRecovered() : [];
    const note = [state.recoveredNote, more.length ? `An earlier save of ${more.join(', ')} was interrupted; the add-in finished it. Nothing was lost.` : ''].filter(Boolean).join(' ');
    state.recoveredNote = '';
    setStatus(`${note ? `${note} ` : ''}Updated ${now}: ${what}.`, note ? 'warn' : 'ok');
  }
  /** "2026-10-20 18:05 UTC": readable on the Config sheet, never turned into a date by Excel, parsed by core.soharTime. */
  function utcStamp(d) { return `${d.toISOString().slice(0, 16).replace('T', ' ')} UTC`; }
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
    wireTabs();
    wireClinic();
    // ?tab=clinic or #clinic opens the Clinic tab, as a clinic file would.
    showPaneTab(/(^|[?&])tab=clinic\b/.test(location.search) || location.hash === '#clinic' ? 'clinic' : 'admin');
    wire();
    autoRefresh();
  }

  /** Inside a clinic file: Clinic tab only, from tblClinic / tblClinicReview / tblClinicMeta. */
  function startClinicFile(cf) {
    state.clinicFile = true;
    const m = cf.meta || {};
    state.clinic = { rows: cf.rows || [], review: cf.review || [], runDate: m['Run Date'] || '', location: m.Location || '' };
    // The clinic's own ticks saved in this file win over what the file was made with.
    for (const q of cf.requests || []) {
      const r = state.clinic.rows.find((x) => C.keyOf(x) === String(q.MRN));
      if (r) r['Re-verify Anyway'] = /^(on|yes|y|true|1)$/i.test(String(q['Re-verify Anyway'] || '')) ? 'Yes' : '';
    }
    $('adminSteps').hidden = true;
    document.querySelectorAll('#panelAdmin > details').forEach((d) => { d.hidden = true; });
    $('adminOnlyTemplate').hidden = false;
    $('btnFull').hidden = true;
    setStatus(`Clinic file: ${m.Location || 'all locations'}, run ${m['Run Date'] || '?'}${m['Add-in Version'] ? ` (add-in ${m['Add-in Version']})` : ''}.`, 'ok');
    renderClinic(m.Location || '');
    showPaneTab('clinic');
    diagnostics();
  }

  // ---------- pane tabs: Clinic | Admin ----------
  function wireTabs() {
    const tabs = [...document.querySelectorAll('[data-ptab]')];
    tabs.forEach((t, i) => {
      t.addEventListener('click', () => showPaneTab(t.dataset.ptab));
      t.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft' && e.key !== 'Home' && e.key !== 'End') return;
        e.preventDefault();
        const j = e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : (i + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
        showPaneTab(tabs[j].dataset.ptab); tabs[j].focus();
      });
    });
  }
  function showPaneTab(name) {
    document.querySelectorAll('[data-ptab]').forEach((t) => {
      const on = t.dataset.ptab === name;
      t.setAttribute('aria-selected', String(on)); t.tabIndex = on ? 0 : -1;
    });
    $('panelClinic').hidden = name !== 'clinic';
    $('panelAdmin').hidden = name !== 'admin';
  }

  function wire() {
    $('btnScan').addEventListener('click', () => guard(refreshAll, 'scanResult'));
    $('btnBuild').addEventListener('click', () => guard(build, 'buildResult'));
    $('skipBox').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.id === 'btnReverifyAll') guard(reverifyAll, 'preflight');
      else if (b.id === 'btnReverifyCancel') guard(cancelRequests, 'preflight');
      else if (b.id === 'btnSkipCsv') guard(() => saveCSV(C.fileName('skipped'), SKIP_COLS, skipRows(), 'preflight'), 'preflight');
    });
    $('btnDlMap').addEventListener('click', () => guard(() => saveCSV(C.fileName('mapping'), C.MAPPING_DISCOVERY_COLS, (state.exports && state.exports.mapping) || [], 'preflight'), 'preflight'));
    $('btnDlTest').addEventListener('click', () => guard(() => saveCSV(C.fileName('payerTests'), C.VERIFICATION_COLS, (state.exports && state.exports.payerTests) || [], 'preflight'), 'preflight'));
    $('guarBox').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (b && b.id === 'btnGuarYes') guard(resubmitGuarantors, 'resultsResult');
      else if (b && b.id === 'btnGuarNo') { $('guarBox').hidden = true; state.guarantorDismissed = true; }
    });
    dropZone('dropClinicBack', 'fileClinicBack', (files) => guard(() => bringBackClinicFiles(files), 'clinicFilesNote'));
    $('inboxBox').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      // loadInbox redraws the box and the pressed button with it: put focus where the result is (or the new button).
      const refocus = (id) => () => { const el = $(id) || $('clinicFilesNote'); if (el) el.focus(); };
      if (b.id === 'btnInboxApply') guard(applyInbox, 'clinicFilesNote').then(refocus('clinicFilesNote'));
      else if (b.id === 'btnInboxSetup') guard(setupInbox, 'clinicFilesNote').then(refocus('clinicFilesNote'));
      else if (b.id === 'btnInboxCheck') guard(loadInbox, 'clinicFilesNote').then(refocus('btnInboxCheck'));
      else if (b.dataset.goto) X.activate(b.dataset.goto);
    });
    dropZone('dropCred', 'fileCred', (files) => guard(() => uploadPayerRules(files), 'credStatus'));
    $('btnCredRefresh').addEventListener('click', () => guard(async () => { await loadPayerRules(); await refreshRules(); renderCredStatus(); }, 'credStatus'));
    ['learnBox', 'clinicNotesBox', 'payerNotesBox'].forEach((id) => $(id).addEventListener('click', (e) => { const b = e.target.closest('[data-goto]'); if (b) X.activate(b.dataset.goto); }));
    $('btnDlVer').addEventListener('click', () => guard(() => download('verification'), 'preflight'));
    $('btnDlDisc').addEventListener('click', () => guard(() => download('discovery'), 'preflight'));
    dropZone('dropReport', 'fileImport', (files) => guard(() => uploadImport(files), 'scanResult'));
    dropZone('dropResults', 'fileResults', (files) => guard(() => uploadResults(files), 'resultsResult'));
    $('btnCopyMsg').addEventListener('click', () => guard(() => copyAdminMessage('action'), 'resultsResult'));
    $('btnCopyBen').addEventListener('click', () => guard(() => copyAdminMessage('benefits'), 'benBox'));
    $('btnDlWork').addEventListener('click', () => guard(async () => saveCSV(C.fileName('worklist'), C.RESULTS_COLS, (await currentResults()) || [], 'resultsResult'), 'resultsResult'));
    $('btnDlBen').addEventListener('click', () => guard(async () => saveCSV(C.fileName('benefits'), C.BENEFITS_EXPORT_COLS, C.benefitsExport((await currentResults()) || []), 'benBox'), 'benBox'));
    $('btnDlDiscVer').addEventListener('click', () => guard(downloadDiscoveryVerify, 'resultsResult'));
    $('btnDlVerFail').addEventListener('click', () => guard(downloadVerifyFailed, 'resultsResult'));
    $('btnDlFinal').addEventListener('click', () => guard(downloadTrackerFile, 'benBox'));
    $('btnClearSaved').addEventListener('click', () => guard(clearSaved, 'resultsResult'));
    $('btnDlTasks').addEventListener('click',() => guard(async () => saveCSV(C.fileName('tasks'), C.TASK_COLS, await currentTasks(), 'resultsResult', { safe: true }), 'resultsResult'));
    $('btnFull').addEventListener('click', () => guard(() => openFull(), 'resultsResult'));
    $('btnIr').addEventListener('click', () => $('fileIr').click());
    $('fileIr').addEventListener('change', (e) => guard(() => uploadIr(e.target), 'irResult'));
    $('btnCopyIr').addEventListener('click', () => guard(async () => {
      const text = C.irMessage(await currentIr());
      if (await copyText(text)) note('irResult', 'IR message copied. Paste it into Outlook or Teams.', 'ok');
      else note('irResult', 'Copy is blocked here. Open full screen > Messages to copy it.', 'warn');
    }, 'irResult'));
    $('btnDlIr').addEventListener('click', () => guard(async () => saveCSV(C.fileName('ir'), C.IR_COLS, await currentIr(), 'irResult'), 'irResult'));
    // Fee estimates (shown only when Config "Fee estimates" is On).
    dropZone('dropFee', 'fileFee', (files) => guard(() => uploadFee(files), 'feeStatus'));
    $('btnFeeRefresh').addEventListener('click', () => guard(async () => {
      await loadConfig();
      if (!feesOn()) { note('feeStatus', 'Fee estimates are Off in Config.', 'warn'); return; }
      if (state.fees) state.fees.loaded = false;
      await refreshFees();
      await afterFeeChange();
    }, 'feeStatus'));
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
      showFeeUi();
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
      const rw = $('reportWarn');
      rw.hidden = true; rw.innerHTML = '';
      if (!det) {
        out.innerHTML = `<p class="bad">No header row found. The Import sheet needs the NextGen columns, including <b>Per Nbr</b> and <b>Appt Dt</b>.</p>`;
        state.detection = null; $('btnBuild').disabled = true; return;
      }
      state.detection = det;
      // R1: leftover rows below "Total Appointments" and a payer column holding provider names, shown up front.
      if (det.warnings && det.warnings.length) {
        rw.innerHTML = `<p><b>Check the report.</b></p><ul>${det.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>`;
        rw.hidden = false;
      }
      const dataRows = Math.max(0, (det.dataEnd ?? state.rows.length) - det.headerRow - 1);
      const found = Object.keys(det.colMap).map((f) => `<span class="chip">${label(f)}</span>`).join('');
      const miss = det.optionalMissing.map((f) => `<span class="chip miss">${label(f)} missing</span>`).join('');
      let warn = '';
      if (det.colMap.patName == null || det.colMap.dob == null) warn += `<p class="warn">Sohar needs patient name and DOB. This report has ${det.colMap.patName == null ? 'no name column' : ''}${det.colMap.patName == null && det.colMap.dob == null ? ' and ' : ''}${det.colMap.dob == null ? 'no DOB column' : ''}; those fields will be blank until the report includes them.</p>`;
      out.innerHTML = `<p class="ok">Report loaded: ${dataRows.toLocaleString()} appointment rows${det.footerRow >= 0 ? ' (up to the Total Appointments line)' : ''}.</p>${warn}<details class="more inline"><summary>Columns found</summary><div class="chips">${found}${miss}</div></details>`;
      if (det.missing.length) { out.innerHTML += `<p class="bad">Missing required columns: ${det.missing.map(label).join(', ')}.</p>`; $('btnBuild').disabled = true; return; }
      await loadPayerStep();
      $('btnBuild').disabled = false;
    } finally { unbusy('btnScan', 'Re-read the Import sheet'); }
  }

  // ---------- payers: auto-mapped, with an optional review list ----------
  async function loadPayerStep() {
    state.payerMap = await X.readPayerMap();
    state.soharPayers = await X.readSoharPayers();
    renderPayers();
  }

  /** After a payer is corrected, rebuild so routing follows the Payer Map right away. */
  async function afterMap() {
    await loadPayerStep();
    if (state.patients) await build();
  }

  /** NextGen payer names (primary and secondary) in the report, with the patients that have each. Stops at the report total (R1). */
  function payersInImport() {
    const det = state.detection; if (!det) return new Map();
    const counts = new Map();
    const end = det.dataEnd ?? state.rows.length;
    for (let r = det.headerRow + 1; r < end; r++) {
      const row = state.rows[r] || [];
      const per = String(row[det.colMap.perNbr] ?? '').trim();
      if (!per) continue;
      for (const f of ['payer', 'secPayer']) {
        if (det.colMap[f] == null) continue;
        const name = String(row[det.colMap[f]] ?? '').trim();
        if (!name) continue;
        if (!counts.has(name)) counts.set(name, new Set());
        counts.get(name).add(per);
      }
    }
    return counts;
  }

  const AUTO_ORDER = { 'auto-low': 0, '': 0, 'auto-medium': 1, 'auto-high': 2, 'learned': 3 };
  const HUMAN_LABELS = new Set(['exact', 'probable', 'confirmed']); // "Looks right" keeps these as they are

  /**
   * Write Payer Map rows and log each change on the Payer Map Log. by: Auto-match, Pane, Learning, Re-score, Suggestion.
   * human: a person's write (stamps Last reviewed). changes: [{entry, score, evidence}] or plain entries.
   */
  async function writeMap(changes, by, human) {
    const list = changes.map((c) => (c.entry ? c : { entry: c, evidence: c.notes || '' }));
    const old = state.payerMap || await X.readPayerMap();
    await X.upsertPayerMapMany(list.map((c) => c.entry), { human: !!human });
    // Rows an earlier log save lost go first, so the log keeps the order things happened in.
    const rows = (state.pendingLog || []).concat(C.mapChangeLog(old, list.map((c) => ({ ...c, by: c.by || by })), fmtStamp(new Date())));
    if (!rows.length) return;
    try {
      await X.appendPayerMapLog(C.PAYER_MAP_LOG_COLS, rows);
      state.pendingLog = [];
    } catch (e) {
      // Kept and tried again on the next build (ahead of the sheet-edit check, so the change keeps By = Pane). Said in
      // the status line: the payer list is redrawn right after this and would wipe a note there.
      console.error(e);
      state.pendingLog = rows;
      setStatus(`The Payer Map change is saved, but the Payer Map Log was not updated (${e.message}). The add-in logs it on the next build.`, 'warn');
    }
  }

  /**
   * Payer list for this report. Needs a person (top block): unmapped, Auto-Low, Review or an unrecognized label, a human
   * label Sohar disagrees with (learning FLAG), or a name with a Sohar suggestion. Optional list: Auto-High, Auto-Medium,
   * Learned. Each row shows the Sohar evidence (learnPayers). Nothing here blocks the build; a person's choice always wins.
   */
  function renderPayers() {
    const idx = C.payerMapIndex(state.payerMap || []);
    const learned = new Map((state.learning || []).map((r) => [r.key, r]));
    const counts = payersInImport();
    const item = (name, n) => {
      const h = idx.get(name.toLowerCase());
      const conf = h ? String(h.confidence || '').toLowerCase() : '';
      const lr = learned.get(name.toLowerCase()) || null;
      const flagged = !!(lr && lr.decision === 'flag');
      const needs = !h || C.needsClassification(h.confidence, h.payerId) || flagged || !!(lr && lr.suggestion);
      return { name, n, h, conf, lr, flagged, needs };
    };
    const inReport = new Set([...counts.keys()].map((n) => n.toLowerCase()));
    // Flags and suggestions come from every payer in the ledger, so the build note and Review sheet count names no one
    // in this report carries. List those too, or their "Use" / "Looks right" buttons could not be reached.
    const offReport = (state.learning || []).filter((r) => !inReport.has(r.key) && (r.decision === 'flag' || r.suggestion)).map((r) => item(r.nextgenPayer, 0));
    const items = [...counts].map(([name, pers]) => item(name, pers.size)).concat(offReport)
      .filter((it) => !C.isVip(it.name) && !C.isSelfPay(it.name) && (it.needs || C.isAutoLabel(it.conf)))
      .sort((a, b) => (b.needs - a.needs) || (AUTO_ORDER[a.conf] ?? 0) - (AUTO_ORDER[b.conf] ?? 0) || b.n - a.n || a.name.localeCompare(b.name));
    const box = $('payerList');
    box.hidden = !items.length;
    if (!items.length) { box.innerHTML = ''; return; }
    const isLow = (it) => it.needs;
    const low = items.filter(isLow).length, rest = items.length - low;
    const keepOpen = box.querySelector('details') && box.querySelector('details').open;
    box.innerHTML = (low ? `<div class="needclass"><p class="warn"><b>${low} insurance name${low === 1 ? '' : 's'} need${low === 1 ? 's' : ''} manual classification.</b> Not matched confidently, labeled Review or with a label the add-in does not know, or Sohar disagrees with the mapping. Unclassified names go to discovery (or to both insurances when the last visit was billed differently). Pick the right Sohar payer once and it is remembered.</p><div class="needrows"></div></div>` : '')
      + (rest ? `<details class="more payerreview"${keepOpen ? ' open' : ''}><summary>Review auto-mapped payers (${rest})</summary>
      <p class="hint">Matched confidently (Auto-High / Auto-Medium) or learned from Sohar results (Learned) and sent to verification. Optional: fix a wrong match and it is remembered.</p>
      <div class="payerrows"></div></details>` : '');
    for (const it of items) {
      const list = box.querySelector(isLow(it) ? '.needrows' : '.payerrows');
      const h = it.h || {};
      const label = it.h ? (h.confidence || 'No confidence') : 'Not mapped';
      const cls = it.conf === 'auto-high' || it.conf === 'learned' ? 'ok' : it.conf === 'auto-medium' ? 'mid' : 'low';
      const latestNote = String(h.notes || '').split(' | ')[0].replace(/^Auto-matched by the add-in:\s*/, '');
      const evidence = C.learningEvidence(it.lr);
      const sg = it.lr && it.lr.suggestion;
      const wrap = document.createElement('div');
      wrap.className = 'payer';
      wrap.innerHTML = `<div class="name">${esc(it.name)} <span class="n">(${it.n ? `${it.n} patient${it.n > 1 ? 's' : ''}` : 'not in this report'})</span> <span class="conf ${cls}">${esc(label)}</span></div>
        <div class="guess">${h.payerId ? `→ ${esc(h.payerName || '')} <span class="id">${esc(h.payerId)}</span>` : '<span class="hint">No Sohar payer guessed</span>'}</div>
        ${latestNote ? `<div class="hint ev">${esc(latestNote)}</div>` : ''}
        ${evidence ? `<div class="hint ev">${esc(evidence)}</div>` : ''}
        ${it.flagged ? `<p class="warn">${esc(it.lr.note)}</p>` : ''}
        ${sg ? `<div class="alt-actions"><button class="small" data-act="suggest">Use ${esc(sg.payerName || sg.payerId)} <span class="id">${esc(sg.payerId)}</span></button> <span class="hint">${sg.n} of ${it.lr.discovery} discovery patients</span></div>` : ''}
        <div class="alt-actions">${h.payerId ? '<button class="link" data-act="ok">Looks right</button>' : ''}<button class="link" data-act="change">Change</button><button class="link" data-act="notpayer">Not a payer</button></div>
        <div class="pick" hidden><input type="search" placeholder="Search Sohar payers" aria-label="Search Sohar payers for ${esc(it.name)}"><div class="sugg"></div></div>`;
      const use = wrap.querySelector('[data-act="suggest"]');
      if (use) use.addEventListener('click', () => guard(async () => {
        await writeMap([{ entry: { nextgenPayerName: it.name, payerId: sg.payerId, payerName: sg.payerName, confidence: 'Confirmed', notes: `Suggested by Sohar results (${sg.n} of ${it.lr.discovery} discovery patients), confirmed in the add-in` },
          evidence: `${sg.n} of ${it.lr.discovery} discovery patients returned ${sg.payerName || sg.payerId}` }], 'Suggestion', true);
        await afterMap();
      }, 'payerList'));
      const pick = wrap.querySelector('.pick');
      const input = wrap.querySelector('input');
      const sugg = wrap.querySelector('.sugg');
      const show = (q) => {
        const found = C.suggestPayers(q || it.name, state.soharPayers || [], 6);
        sugg.innerHTML = found.length ? '' : '<p class="hint">No close match. Try another word, or ask Sohar to add this payer.</p>';
        for (const s of found) {
          const b = document.createElement('button');
          b.className = 'small';
          b.innerHTML = `<span>${esc(s.payerName)}</span><span class="id">${esc(s.payerId)}</span>`;
          b.addEventListener('click', () => guard(async () => {
            await writeMap([{ nextgenPayerName: it.name, payerId: s.payerId, payerName: s.payerName, confidence: 'Confirmed', notes: 'Mapped in the add-in' }], 'Pane', true);
            await afterMap();
          }, 'payerList'));
          sugg.appendChild(b);
        }
      };
      input.addEventListener('input', () => show(input.value));
      const ok = wrap.querySelector('[data-act="ok"]');
      if (ok) ok.addEventListener('click', () => guard(async () => {
        // A human label stays as it is (Exact stays Exact); anything else becomes Confirmed. With Sohar conflicts on
        // record, the note marks them reviewed, so the flag stays quiet until more conflicts arrive (core.learnPayers).
        const keep = HUMAN_LABELS.has(it.conf);
        const conflicts = it.lr ? it.lr.conflicts : 0;
        const ack = conflicts > 0 ? C.flagAckNote(conflicts, fmtStamp(new Date())) : '';
        const what = keep ? `Checked in the add-in (${h.confidence})` : `Confirmed in the add-in (was ${h.confidence || 'blank'})`;
        await writeMap([{ entry: { nextgenPayerName: it.name, payerId: h.payerId, payerName: h.payerName, confidence: keep ? h.confidence : 'Confirmed', notes: [ack, what].filter(Boolean).join('; ') },
          evidence: ack ? `${ack}${it.lr.note ? `: ${it.lr.note}` : ''}` : what, always: !!ack }], 'Pane', true);
        await afterMap();
      }, 'payerList'));
      wrap.querySelector('[data-act="change"]').addEventListener('click', () => { pick.hidden = !pick.hidden; if (!pick.hidden) { show(''); input.focus(); } });
      wrap.querySelector('[data-act="notpayer"]').addEventListener('click', () => guard(async () => {
        await writeMap([{ nextgenPayerName: it.name, payerId: '', payerName: '', confidence: 'Not a payer', notes: 'Marked in the add-in' }], 'Pane', true);
        await afterMap();
      }, 'payerList'));
      list.appendChild(wrap);
    }
  }

  // ---------- step 3 ----------
  async function build(opts) {
    busy('btnBuild', 'Building…');
    // Payer Map Log rows collected during this build. They are saved at the end, or in the finally below if the build
    // stops after the Payer Map was already written, so a change is never left to look like a "Sheet edit" later.
    const logRows = [];
    let logSaved = false;
    const saveLog = async () => {
      if (logSaved || !logRows.length) return '';
      logSaved = true;
      const hadPending = !!(state.pendingLog && state.pendingLog.length);
      // appendPayerMapLog re-reads the log before saving, so a failed read earlier never overwrites it.
      try {
        await X.appendPayerMapLog(C.PAYER_MAP_LOG_COLS, logRows);
        state.pendingLog = [];
        if (hadPending) setStatus('Payer Map Log caught up: the pane change that could not be logged earlier is on it now.', 'ok');
        return '';
      } catch (e) {
        console.error(e);
        state.pendingLog = logRows.slice(); // tried again on the next build or pane change
        return `Payer Map Log not updated: ${e.message}. The add-in tries again on the next build.`;
      }
    };
    try {
      state.cfg = await X.readConfig();
      state.cfg.scope = $('scope').value;
      state.cfg.asOf = $('asOf').value ? new Date($('asOf').value + 'T00:00:00Z') : null;
      // Optional Config rows for payer learning, added once per session when missing (never changes existing rows).
      if (!state.learnRowsChecked) {
        state.learnRowsChecked = true;
        // Fee estimates: only the "Fee estimates" (Off) row while Off; its estimate defaults once it is On.
        const feeRows = F ? (feesOn() ? F.FEE_CONFIG_ROWS : F.FEE_CONFIG_ROWS.slice(0, 1)) : [];
        // "MRN ids since": Sohar checks created before this time carry the Per Nbr as their id, later ones the MRN.
        const since = { setting: 'MRN ids since', value: utcStamp(new Date()), what: 'When this workbook started sending the MRN (not the Per Nbr) as the Sohar id. Sohar results created before then are matched by Per Nbr. Set by the add-in; do not change.' };
        await X.ensureConfigRows(C.LEARNING_CONFIG_ROWS.concat(C.REVERIFY_CONFIG_ROWS, C.PAYER_TEST_CONFIG_ROWS, [since], feeRows)).catch((e) => console.warn('Config learning / re-verify rows not added', e));
        if (!String(state.cfg.mrnSince || '').trim()) state.cfg = await X.readConfig().then((c) => ({ ...c, scope: state.cfg.scope, asOf: state.cfg.asOf })).catch(() => ({ ...state.cfg, mrnSince: since.value }));
      }
      const learn = state.cfg.learning || C.LEARN_DEFAULTS;
      state.payerMap = await X.readPayerMap();
      if (!state.soharPayers) state.soharPayers = await X.readSoharPayers();
      const when = fmtStamp(new Date());
      // Payer Map Log: pane changes whose log save failed come first (they explain the map), then edits made on the
      // sheet since the last logged state (or a Baseline), before the add-in writes anything.
      const pending = state.pendingLog || [];
      logRows.push(...pending);
      try { logRows.push(...C.sheetEditLog(state.payerMap, (await X.readPayerMapLog()).concat(pending), when)); } catch (e) { console.error(e); }
      // The outcomes ledger feeds learning. Read once per pane session; processResults keeps state.outcomes current
      // after each upload. If it cannot be read, nothing is learned this time (never a guess from nothing) and the
      // next build tries again.
      if (!state.outcomes) state.outcomes = await X.readPayerOutcomes().catch((e) => { console.error(e); return null; });
      // Ledger rows saved before v0.12 name the patient by Per Nbr: read them as this report's MRN so one patient counts once.
      if (!state.perToKey.size) { const pk = C.buildPatients(state.rows, { ...state.cfg, payerMap: state.payerMap, scope: 'all' }); if (!pk.error) state.perToKey = pk.perToKey; }
      if (state.outcomes) state.outcomes = C.appendOutcomes(state.outcomes, [], state.perToKey).rows;
      const outcomes = state.outcomes || [];
      state.finClasses = learn.on ? C.learnFinClasses(outcomes, learn) : [];
      const cfgFor = () => ({ ...state.cfg, payerMap: state.payerMap, soharPayers: state.soharPayers, learnedFinClasses: state.finClasses });
      let res = C.buildPatients(state.rows, cfgFor());
      const out = $('buildResult'); out.hidden = false;
      if (res.error) { out.innerHTML = `<p class="bad">${esc(res.error)}</p>`; return; }
      // R2: new NextGen payer names are matched to Sohar payers automatically and saved on the Payer Map once
      // (editable there; a manual confidence always wins). Build again so the Fin Class index sees them too.
      let autoNote = '', rebuild = false;
      if (res.autoMapped && res.autoMapped.length) {
        const add = res.autoMapped.map(({ nextgenPayerName, payerId, payerName, confidence, notes }) => ({ nextgenPayerName, payerId, payerName, confidence, notes }));
        await X.upsertPayerMapMany(add);
        logRows.push(...C.mapChangeLog(state.payerMap, res.autoMapped.map((a) => ({ entry: a, by: 'Auto-match', score: a.score,
          evidence: `${String(a.notes).replace(/^Auto-matched by the add-in:\s*/, '')}; ${a.patients} patient${a.patients === 1 ? '' : 's'} in this report` })), when));
        state.payerMap = await X.readPayerMap();
        rebuild = true;
        const lowN = add.filter((a) => a.confidence === 'Auto-Low').length;
        autoNote = `<p class="hint">${add.length} new insurance name${add.length === 1 ? ' was' : 's were'} matched to Sohar payers and added to the Payer Map${lowN ? ` (${lowN} low confidence)` : ''}. ${lowN ? ' The low-confidence ones are listed above for classification.' : ''}</p>`;
      }
      // Payer learning: promote / demote the add-in's own labels from Sohar outcomes, then re-score Auto-Low rows.
      // Human labels are never changed (Config: Learning, Learn: ...). Every change goes on the Payer Map Log.
      const plan = C.planLearning(outcomes, state.payerMap, state.soharPayers, learn);
      if (plan.changes.length) {
        await X.upsertPayerMapMany(plan.changes.map((c) => c.entry));
        logRows.push(...C.mapChangeLog(state.payerMap, plan.changes, when));
        state.payerMap = await X.readPayerMap();
        rebuild = true;
      }
      // Payer tests (Off by default): their results add suggestions where discovery gave none.
      state.payerTests = await X.readPayerTests().catch((e) => { console.error(e); return state.payerTests || []; });
      state.learning = C.withTestSuggestions(plan.rollup, C.payerTestSummary(state.payerTests, state.cfg.payerTests));
      if (rebuild) {
        res = C.buildPatients(state.rows, cfgFor());
        if (res.error) { out.innerHTML = `<p class="bad">${esc(res.error)}</p>`; return; }
      }
      const sheetWarn = [];
      const logWarn = await saveLog();
      if (logWarn) sheetWarn.push(logWarn);
      await X.writePayerLearning(C.PAYER_LEARNING_COLS, C.payerLearningRows(state.learning)).catch((e) => { console.error(e); sheetWarn.push(`Payer Learning sheet not updated: ${e.message}`); });
      const k = plan.counts;
      const learnNote = outcomes.length || plan.changes.length
        ? `<p class="hint">Learning: ${k.promoted} promoted, ${k.demoted} demoted, ${k.suggestions} suggestion${k.suggestions === 1 ? '' : 's'}, ${k.rescored} re-scored${k.flagged ? `; Sohar disagrees with ${k.flagged} human-labeled payer${k.flagged === 1 ? '' : 's'} (Review sheet)` : ''}${state.finClasses.length ? `; ${state.finClasses.length} Fin Class${state.finClasses.length === 1 ? '' : 'es'} learned` : ''}.${learn.on ? '' : ' Learning is Off in Config: outcomes are recorded, nothing changes automatically.'} See the Payer Learning sheet.</p>`
        : '';
      // Re-verify skip window: patients whose same insurance Sohar verified eligible recently leave the verification file.
      const sk = await applySkips(res, cfgFor);
      res.patients = sk.patients;
      if (sk.warn) sheetWarn.push(sk.warn);
      const nSkip = state.skipped.length;
      const skipNote = nSkip ? `<p class="hint">${nSkip} skipped: verified within ${sk.days} days (listed in step 2).</p>` : '';
      state.perToKey = res.perToKey || state.perToKey;
      renderPayers();
      state.patients = res.patients; state.review = res.review.concat(C.learningReviewRows(state.learning));
      const pt = state.cfg.payerTests || C.PAYER_TEST_DEFAULTS;
      state.exports = C.buildExports(res.patients, { ...state.cfg, guarantorRequests: state.guarantorActive, mappingDiscovery: pt.mappingDiscovery !== false });
      await X.writePayerMapping(C.MAPPING_DISCOVERY_COLS, state.exports.mapping).catch((e) => { console.error(e); sheetWarn.push(`Payer Mapping sheet not updated: ${e.message}`); });
      state.exports.payerTests = [];
      if (pt.on) {
        const errs = new Map();
        for (const o of outcomes) if (o.Kind === 'Verification' && o.Signal === 'payer error' && o['NextGen Payer']) errs.set(String(o['NextGen Payer']).toLowerCase(), (errs.get(String(o['NextGen Payer']).toLowerCase()) || 0) + 1);
        const built = C.buildPayerTests(res.patients, { soharPayers: state.soharPayers, payerMap: state.payerMap, previous: state.payerTests, payerErrors: errs,
          payers: pt.payers, patients: pt.patients, specialty: state.cfg.specialty, when: fmtStamp(new Date()) });
        state.exports.payerTests = built.rows;
        if (built.table.length !== state.payerTests.length) {
          try { await X.writePayerTests(C.PAYER_TEST_COLS, built.table); state.payerTests = built.table; } catch (e) { console.error(e); sheetWarn.push(`Payer Tests not saved: ${e.message}`); state.exports.payerTests = []; }
        }
      }
      await X.replaceTable(X.TABLES.patients, C.PATIENT_COLS, res.patients);
      // The Review sheet is written by refreshRules (build review + clinic rule flags).
      await X.replaceTable(X.TABLES.verification, C.VERIFICATION_COLS, state.exports.verification);
      await X.replaceTable(X.TABLES.discovery, C.DISCOVERY_COLS, state.exports.discovery);
      await X.setConfigValue('Last build', `${C.fmtDate(new Date())} · scope: ${state.cfg.scope} · as of ${C.fmtDate(res.asOf)}`).catch(() => {});

      const s = C.summarize(res.patients);
      const total = Math.max(1, s.inScope);
      const pct = (n) => `${(100 * n / total).toFixed(1)}%`;
      out.innerHTML = `
        <p>${s.inScope.toLocaleString()} patient${s.inScope === 1 ? '' : 's'}${state.cfg.scope === 'upcoming' ? ` with a visit on or after ${C.fmtDate(res.asOf)}` : ' in the report'}.</p>
        <div class="routebar" role="img" aria-label="${s.verification} verification, ${s.discovery} discovery, ${s.exclude || 0} excluded"></div>
        <div class="counts">
          <div class="count v"><div class="n">${s.verification}</div><div class="l">verification<br>(have insurance on file)</div></div>
          <div class="count d"><div class="n">${s.discovery}</div><div class="l">discovery<br>(insurance missing or unmapped)</div></div>
          <div class="count x"><div class="n">${s.exclude || 0}</div><div class="l">left out</div></div>
        </div>
        <p class="issues" id="buildIssues"></p>${skipNote}${autoNote}${learnNote}${sheetWarn.map((w) => `<p class="warn">${esc(w)}</p>`).join('')}`;
      // widths set through the CSSOM so the strict style-src policy holds
      const bar = out.querySelector('.routebar');
      for (const [cls, n] of [['v', s.verification], ['d', s.discovery], ['x', s.exclude || 0]]) {
        const seg = document.createElement('span'); seg.className = cls; seg.style.width = pct(n); bar.appendChild(seg);
      }
      preflight();
      markDone(1); setCurrent(2);
      const nv = state.exports.verification.length, nd = state.exports.discovery.length;
      const nvp = new Set(state.exports.verification.map((r) => r.id)).size;
      $('btnDlVer').disabled = !nv; $('btnDlDisc').disabled = !nd;
      $('fnVer').textContent = `${nvp} patient${nvp === 1 ? '' : 's'}${nv > nvp ? ` (${nv} rows: ${nv - nvp} with two insurances)` : ''} · ${C.fileName('verification')}`;
      const ng = state.exports.discovery.filter((r) => /-G$/.test(r.id)).length;
      $('fnDisc').textContent = `${nd - ng} patient${nd - ng === 1 ? '' : 's'}${ng ? ` + ${ng} guarantor${ng === 1 ? '' : 's'}` : ''} · ${C.fileName('discovery')}`;
      renderExtraFiles();
      renderSkipped(sk.days);
      if (!(opts && opts.deferRules)) await refreshRules();
    } finally {
      // The build stopped early (an error or an early return): log the Payer Map changes it already made.
      if (!logSaved && logRows.length) {
        const w = await saveLog();
        if (w) setStatus(w, 'warn');
      }
      unbusy('btnBuild', 'Rebuild with these options');
    }
  }

  /**
   * Re-verify skip window (Config: Skip recent verifications, Re-verify days, Re-verify in new month; core.applyReverifySkips).
   * The history is the saved Sohar uploads matched again (MRN or an old Per Nbr id, Sohar ref or Member ID; never a name match).
   * "Re-verify anyway" requests are dropped once a newer verification exists. If the uploads, refs or requests cannot
   * be read, nobody is skipped this time: a patient is never left out on a guess.
   */
  async function applySkips(res, cfgFor) {
    const rc = (state.cfg && state.cfg.reverify) || C.REVERIFY_DEFAULTS;
    state.skipped = []; state.requestedAnyway = [];
    const pk = res.perToKey || state.perToKey;
    const none = (warn) => ({ patients: res.patients, days: rc.days, warn });
    let saved, refs, requests, gReq;
    try {
      saved = await X.readUploads();
      refs = await X.readSoharRefs();
      requests = await X.readReverifyRequests();
      gReq = await X.readGuarantorRequests();
    } catch (e) { console.error(e); return none(`Recent verifications not checked, so nobody was skipped: ${e.message}`); }
    const merged = saved.length ? C.mergeSoharResults(state.rows, C.accumulateResults(saved, [], '', refs).files, { ...cfgFor(), refs }) : { results: [] };
    if (merged.error) return none('');
    // Guarantor resubmissions still waiting (dropped once the guarantor search came back): the discovery file carries them.
    const gp = C.pruneGuarantorRequests(gReq, merged.results, pk);
    state.guarantorActive = gp.active;
    if (gp.dropped) await X.writeGuarantorRequests(C.GUARANTOR_REQUEST_COLS, gp.rows).catch((e) => console.warn('Guarantor Requests not pruned', e));
    if (!rc.skip || !saved.length) return none('');
    const pr = C.pruneReverifyRequests(requests, merged.results, pk);
    if (pr.dropped || requests.some((q) => !String(q.MRN || '').trim())) await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, pr.rows).catch((e) => console.warn('Re-verify Requests not pruned', e));
    state.reverifyRequests = pr.rows;
    const out = C.applyReverifySkips(res.patients, merged.results, { ...rc, today: new Date(), requests: pr.active });
    state.skipped = out.skipped; state.requestedAnyway = out.requested;
    return { patients: out.patients, days: rc.days, warn: '' };
  }

  /**
   * Step 2: patients left out of the verification file because Sohar verified them eligible recently. Collapsed by
   * default. Clinics pick patients to re-verify anyway (Clinic tab, clinic files); the admin can only send all of them.
   */
  function renderSkipped(days) {
    const box = $('skipBox');
    if (!box) return;
    const list = state.skipped || [], asked = state.requestedAnyway || [];
    box.hidden = !list.length && !asked.length;
    if (box.hidden) { box.innerHTML = ''; return; }
    const short = (d) => esc(String(d || '').slice(0, 5));
    const who = (s) => `<b>${esc(C.lasfir(s.last, s.first))}</b><br><span class="hint">${esc(C.idLabel({ MRN: s.mrn, 'Per Nbr': s.per }))}${s.location ? ` · ${esc(s.location)}` : ''}${s.partial ? ' · other insurance still sent' : ''}</span>`;
    const by = (s) => { const q = (state.reverifyRequests || []).find((r) => r.MRN === s.key); return q && q.By ? q.By : 'request'; };
    box.innerHTML = (list.length ? `<details class="more skiplist"><summary><b>Skipped (verified recently): ${list.length}</b></summary>
      <p class="hint">Sohar verified the same insurance eligible within ${esc(days)} days, so ${list.length === 1 ? 'this patient is' : 'these patients are'} not on the verification file. Clinics can ask for a patient on the Clinic tab or in their clinic file.</p>
      <table class="watch"><thead><tr><th>Patient</th><th>Verified</th><th>Next check</th></tr></thead><tbody>${list.map((s) => `<tr>
        <td>${who(s)}</td><td title="${esc(s.verified)}">${short(s.verified)}</td><td title="${esc(s.due)}">${short(s.due)}</td></tr>`).join('')}</tbody></table>
      <div class="actions"><button class="small" id="btnReverifyAll">Re-verify all ${list.length} anyway</button><button class="link" id="btnSkipCsv">Save this list</button></div>
      </details>` : '')
      + (asked.length ? `<details class="more skiplist"><summary>Sent anyway on request: ${asked.length}</summary>
      <table class="watch"><thead><tr><th>Patient</th><th>Verified</th><th>Asked by</th></tr></thead><tbody>${asked.map((s) => `<tr>
        <td>${who(s)}</td><td title="${esc(s.verified)}">${short(s.verified)}</td><td>${esc(by(s))}</td></tr>`).join('')}</tbody></table>
      <div class="actions"><button class="link" id="btnReverifyCancel">Cancel these requests</button></div></details>` : '');
  }
  const SKIP_COLS = ['Location', 'Next Appt', 'MRN', 'Per Nbr', 'Last Name', 'First Name', 'Verified', 'Next Check', 'Other Insurance Sent'];
  const skipRows = () => (state.skipped || []).map((s) => ({ Location: s.location, 'Next Appt': s.nextAppt, MRN: s.mrn, 'Per Nbr': s.per, 'Last Name': s.last, 'First Name': s.first,
    Verified: s.verified, 'Next Check': s.due, 'Other Insurance Sent': s.partial ? 'Yes' : '' }));

  /** Admin opt-in: every skipped patient goes back on the verification file (until a newer Sohar verification exists). */
  async function reverifyAll() {
    const keys = (state.skipped || []).map((s) => s.key);
    if (!keys.length) return;
    await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, C.addReverifyRequest(await X.readReverifyRequests(), keys, new Date().toISOString(), 'Admin: all skipped'));
    await build();
    note('preflight', `${keys.length} patient${keys.length === 1 ? ' is' : 's are'} back on the verification file. Save it again before uploading to Sohar.`, 'ok');
  }
  async function cancelRequests() {
    const keys = (state.requestedAnyway || []).map((s) => s.key);
    await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, C.removeReverifyRequest(await X.readReverifyRequests(), keys));
    await build();
    note('preflight', 'Requests cancelled: those patients are skipped again until their next check.', 'ok');
  }
  /**
   * A clinic's "Re-verify anyway" tick on the Clinic tab (this workbook): add or drop the request and rebuild. Inside a
   * clinic file the tick is saved in that file (Clinic Requests) for the admin to bring back.
   */
  async function toggleClinicRequest(key, on) {
    const r = ((state.clinic && state.clinic.rows) || []).find((x) => clinicKey(x) === key);
    if (state.clinicFile) {
      const prev = await X.readClinicRequests().catch(() => []);
      const rows = prev.filter((q) => String(q.MRN) !== key);
      rows.push({ MRN: key, 'Last Name': r ? r['Last Name'] : '', 'First Name': r ? r['First Name'] : '', 'Next Appt': r ? r['Next Appt'] : '', 'Re-verify Anyway': on ? 'On' : 'Off', Requested: new Date().toISOString() });
      await X.writeClinicRequests(R.CLINIC_REQUEST_COLS, rows);
      if (r) r['Re-verify Anyway'] = on ? 'Yes' : '';
      renderClinicList();
      note('clinicNote', `Saved in this file. Save it and send it back to the admin team (OneDrive or Teams), who bring the request into the main workbook.`, 'ok');
      return;
    }
    const cur = await X.readReverifyRequests();
    await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, on ? C.addReverifyRequest(cur, key, new Date().toISOString(), 'Clinic tab') : C.removeReverifyRequest(cur, key));
    await build();
    note('clinicNote', on ? `${r ? C.lasfir(r['Last Name'], r['First Name']) + ' is' : 'The patient is'} back on the verification file for the admin's next Sohar run.` : 'Request removed.', 'ok');
  }
  const clinicKey = (r) => C.keyOf(r);

  /** Step 2: the payer mapping discovery (member ID sent) and, when Config "Payer tests" is On, the payer test file. */
  function renderExtraFiles() {
    const m = (state.exports && state.exports.mapping) || [], t = (state.exports && state.exports.payerTests) || [];
    $('mapFile').hidden = !m.length; $('btnDlMap').disabled = !m.length;
    $('fnMap').textContent = `${m.length} patient${m.length === 1 ? '' : 's'} · ${C.fileName('mapping')}`;
    $('testFile').hidden = !t.length; $('btnDlTest').disabled = !t.length;
    const names = new Set((state.payerTests || []).filter((x) => t.some((r) => r.id === x['Test Id'])).map((x) => x['NextGen Payer'].toLowerCase())).size;
    $('fnTest').textContent = `${t.length} check${t.length === 1 ? '' : 's'} for ${names} payer name${names === 1 ? '' : 's'} · ${C.fileName('payerTests')}`;
  }

  /**
   * After results: patients without coverage who have a guarantor. "Yes" saves a request for each; the next discovery
   * file carries a guarantor row (id <MRN>-G) for those whose guarantor DOB is in the report.
   */
  function renderGuarantorPrompt() {
    const box = $('guarBox');
    const c = C.guarantorCandidates(state.patients || [], state.results || [], [...state.guarantorActive.keys()]);
    state.guarantorCandidates = c;
    if (!c.length || state.guarantorDismissed) { box.hidden = true; return; }
    const noDob = c.filter((x) => !x.hasDob).length;
    const none = c.filter((x) => x.reason === C.GUARANTOR_REASONS.none).length;
    box.hidden = false;
    box.innerHTML = `<p><b>${c.length} patient${c.length === 1 ? '' : 's'} without results ${c.length === 1 ? 'has a guarantor' : 'have guarantors'}.</b> Resubmit with the guarantor as the subscriber?</p>
      <p class="hint">${none ? `${none} came back with no coverage from discovery` : ''}${none && c.length - none ? '; ' : ''}${c.length - none ? `${c.length - none} came back ineligible or with an error` : ''}. Their guarantor goes on the next discovery file (one extra row each); results come back tied to the patient and flagged Via guarantor on the Review sheet.</p>
      ${noDob ? `<p class="warn">${noDob} of them ha${noDob === 1 ? 's' : 've'} no guarantor DOB in the report, so ${noDob === 1 ? 'it waits' : 'they wait'} until the NextGen report has a Guar Birth Dt column.</p>` : ''}
      <details class="more inline"><summary>Who</summary><ul class="pc-items">${c.slice(0, 40).map((x) => `<li>${esc(C.lasfir(x.last, x.first))} · ${esc(C.idLabel({ MRN: x.mrn, 'Per Nbr': x.per }))} · ${esc(x.reason)}${x.hasDob ? '' : ' · no guarantor DOB'}</li>`).join('')}${c.length > 40 ? `<li>${c.length - 40} more</li>` : ''}</ul></details>
      <div class="actions"><button class="primary" id="btnGuarYes">Yes, add to the next discovery file</button><button class="link" id="btnGuarNo">Not now</button></div>`;
  }
  async function resubmitGuarantors() {
    const c = state.guarantorCandidates || [];
    if (!c.length) return;
    await X.writeGuarantorRequests(C.GUARANTOR_REQUEST_COLS, C.addGuarantorRequests(await X.readGuarantorRequests(), c, new Date().toISOString(), 'Admin'));
    await build();
    const n = ((state.exports && state.exports.discovery) || []).filter((r) => /-G$/.test(r.id)).length;
    $('guarBox').hidden = true;
    note('preflight', `${n} guarantor row${n === 1 ? ' is' : 's are'} on the discovery file${n < c.length ? ` (${c.length - n} wait for a guarantor DOB in the report)` : ''}. Save the discovery file again (step 2) and upload it to Sohar.`, 'ok');
    setCurrent(2);
  }

  /**
   * Clinic files sent back. A Clinic Eligibility file (Clinic Template Tool) has Clinic Work: rules.mergeClinicWork applies
   * its re-verify ticks, tasks done, notes and payer notes. An older Sohar clinic file has Clinic Requests: its ticks go into
   * the Re-verify Requests. Every file is read before anything is saved, so a bad file leaves the workbook as it was.
   */
  async function bringBackClinicFiles(files) {
    const sheetOf = (wb, name) => wb.sheets.find((x) => x.name.trim().toLowerCase() === name.toLowerCase());
    const s = clinicMerge(await X.readReverifyRequests());
    try {
      for (const f of files) {
        if (!/\.xls[xm]$/i.test(f.name)) throw new Error(`${f.name} is not a clinic file (.xlsx).`);
        const wb = await XL.readWorkbook(new Uint8Array(await f.arrayBuffer()));
        const work = sheetOf(wb, R.SE_SHEETS.work), sheet = sheetOf(wb, 'Clinic Requests');
        if (work) {
          const info = {};
          for (const r of R.tableObjects((sheetOf(wb, R.SE_SHEETS.info) || {}).rows)) info[String(r.Setting ?? '').trim()] = String(r.Value ?? '').trim();
          if (info.Kind !== 'Clinic Eligibility') throw new Error(`${f.name} has a Clinic Work sheet but no Clinic Eligibility info (hidden SE Info sheet).`);
          await mergeClinicRows(s, work.rows, info.Location);
        } else if (sheet) {
          const info = sheetOf(wb, 'Clinic Info');
          const [hdr, ...body] = sheet.rows.filter((r) => (r || []).some((v) => v !== '' && v != null));
          const rows = body.map((r) => Object.fromEntries((hdr || []).map((h, j) => [String(h).trim(), r[j] == null ? '' : String(r[j])])));
          const loc = info && info.rows[1] ? String(info.rows[1][(info.rows[0] || []).indexOf('Location')] || '') : '';
          const m = C.mergeClinicRequests(s.requests, rows, `Clinic file${loc ? ` ${loc}` : ''}`);
          s.requests = m.rows; s.n.reverifyAdded += m.added; s.n.reverifyRemoved += m.removed;
          if (m.added || m.removed) s.changed.add('requests');
        } else throw new Error(`${f.name} has no Clinic Work or Clinic Requests sheet. Use a file saved by the Clinic Template Tool or by this add-in.`);
        s.n.read++;
      }
    } catch (e) { throw new Error(`${e.message} Nothing from these files was saved.`); }
    await saveClinicMerge(s);
    note('clinicFilesNote', bringBackSummary(s.n), 'ok');
  }

  /**
   * Clinic Work rows applied to the admin tables, shared by returned files and the clinic inbox: start with clinicMerge,
   * fold in each location's rows with mergeClinicRows (pure, nothing saved), then saveClinicMerge writes what changed.
   */
  function clinicMerge(requests) {
    return { requests, tasks: null, notes: null, payerNotes: null, applied: null, changed: new Set(), // tasks... read with the first Clinic Work rows
      n: { read: 0, reverifyAdded: 0, reverifyRemoved: 0, doneSet: 0, doneCleared: 0, doneNotes: 0, tasksChanged: 0, notes: 0, payerNotes: 0, skipped: 0, otherClinic: 0, old: 0, applied: 0, locations: 0 } };
  }
  // What the admin is told about: changes that take effect, and ticks that were skipped. A row that changes nothing
  // (a tick already acted on, a blank note over none) is only filed in the brought-back ledger.
  const clinicNews = (n) => R.clinicChanges(n) + n.skipped + n.otherClinic;
  async function mergeClinicRows(s, rows, location) {
    if (!s.tasks) { s.tasks = await liveTasks(); s.notes = await X.readClinicNotes(); s.payerNotes = await X.readClinicPayerNotes(); s.applied = await X.readClinicWorkApplied(); }
    const m = R.mergeClinicWork(rows, { location, requests: s.requests, tasks: s.tasks, notes: s.notes, payerNotes: s.payerNotes, applied: s.applied, now: new Date() });
    s.requests = m.requests; s.notes = m.notes; s.payerNotes = m.payerNotes; s.applied = m.applied;
    // Applied location by location, so the next one sees these Done marks and notes.
    if (m.done.size || m.doneNotes.size) {
      s.tasks = s.tasks.map((t) => {
        const k = String(t['Task Key']);
        if (!m.done.has(k) && !m.doneNotes.has(k)) return t;
        return { ...t, ...(m.done.has(k) ? { Done: m.done.get(k) } : {}), ...(m.doneNotes.has(k) ? { 'Clinic Note': m.doneNotes.get(k) } : {}) };
      });
      s.changed.add('tasks');
    }
    for (const k of Object.keys(m.counts)) s.n[k] += m.counts[k];
    if (m.counts.applied) s.n.locations++;
    if (m.counts.reverifyAdded || m.counts.reverifyRemoved) s.changed.add('requests');
    if (m.counts.notes) s.changed.add('notes');
    if (m.counts.payerNotes) s.changed.add('payerNotes');
    if (m.counts.applied) s.changed.add('applied');
  }
  async function saveClinicMerge(s) {
    if (s.changed.has('requests')) await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, s.requests);
    if (s.changed.has('tasks')) { await X.writeTasks(C.TASK_COLS, s.tasks); if (state.tasks) state.tasks = s.tasks; }
    if (s.changed.has('notes')) { await X.writeClinicNotes(R.CLINIC_NOTE_COLS, s.notes); state.clinicNotes = s.notes; }
    if (s.changed.has('payerNotes')) { await X.writeClinicPayerNotes(R.CLINIC_PAYER_NOTE_COLS, s.payerNotes); state.clinicPayerNotes = s.payerNotes; }
    // Last: if a save above fails, the rows are not marked brought back and the next try applies them.
    if (s.changed.has('applied')) await X.writeClinicWorkApplied(R.CLINIC_APPLIED_COLS, s.applied);
    await build();
    renderClinicNotes();
    await loadInbox();
  }

  /** "2 clinic files read. 3 re-verify requests added, 1 removed. 4 tasks marked done. ..." Zero counts are left out. head replaces the first sentence. */
  function bringBackSummary(n, head) {
    const s = (k, one, many) => `${k} ${k === 1 ? one : many}`;
    const parts = [];
    if (n.reverifyAdded || n.reverifyRemoved) parts.push(`${s(n.reverifyAdded, 're-verify request', 're-verify requests')} added, ${n.reverifyRemoved} removed. Save the verification file again (step 2).`);
    if (n.doneSet || n.doneCleared) parts.push(`${[n.doneSet ? `${s(n.doneSet, 'task', 'tasks')} marked done` : '', n.doneCleared ? `${s(n.doneCleared, 'task', 'tasks')} the clinic had marked done ${n.doneCleared === 1 ? 'is' : 'are'} open again` : ''].filter(Boolean).join('; ')} on NextGen Tasks.`);
    if (n.doneNotes) parts.push(`${s(n.doneNotes, 'done note', 'done notes')} saved in the Clinic Note column on NextGen Tasks.`);
    if (n.notes) parts.push(`${s(n.notes, 'patient note', 'patient notes')} saved on Clinic Notes.`);
    if (n.payerNotes) parts.push(`${s(n.payerNotes, 'payer note', 'payer notes')} saved on Clinic Payer Notes (listed under Payer learning; check the card and fix the Payer Map yourself).`);
    if (n.skipped) parts.push(`${s(n.skipped, 'task the clinic ticked is', 'tasks the clinic ticked are')} no longer on NextGen Tasks, so ${n.skipped === 1 ? 'it was' : 'they were'} skipped.`);
    if (n.otherClinic) parts.push(`${s(n.otherClinic, 'task was', 'tasks were')} ticked in a clinic file for a patient whose next visit is at another clinic, so ${n.otherClinic === 1 ? 'it was' : 'they were'} skipped. Check with that clinic if it looks wrong.`);
    if (n.old) parts.push(`${s(n.old, 'change was', 'changes were')} brought back before or ${n.old === 1 ? 'is' : 'are'} over ${R.APPLIED_DAYS} days old, so ${n.old === 1 ? 'it was' : 'they were'} skipped (your later edits stay).`);
    return `${head || `${s(n.read, 'clinic file', 'clinic files')} read.`} ${parts.join(' ') || 'Nothing new in them.'}`;
  }

  // ---------- the clinic inbox (Clinic Work Inbox, filled by a Power Automate flow; docs/clinic-inbox-flow.md) ----------
  /**
   * The inbox read and merged as a dry run: {missing} when the table is not set up, else {s (a clinicMerge with every
   * location folded in, nothing saved), skipped (blank or unknown rows), rows}. Apply saves that same s.
   */
  async function inboxPlan() {
    const rows = await X.readClinicInbox();
    if (rows == null) return { missing: true };
    // The flow writes text into text columns; a When that Excel turned into a date serial anyway is read back as ISO.
    const norm = rows.map((r) => (typeof r.When === 'number' ? { ...r, When: C.excelSerialToDate(r.When).toISOString() } : r));
    const work = R.inboxWork(norm, new Date());
    const s = clinicMerge(await X.readReverifyRequests());
    for (const [loc, list] of work) await mergeClinicRows(s, list, loc);
    return { s, rows: rows.length, skipped: work.counts.blank + work.counts.unknown };
  }
  /**
   * New inbox rows that change nothing (a note edit that carries a tick already acted on, ticking a task already done) are
   * filed in the brought-back ledger at once, so they are never counted as waiting and never act later (say, on a task the
   * admin reopens). Returns true when it filed some. Only the ledger is written; a failed write is retried on the next read.
   */
  async function fileQuietRows(p) {
    if (p.missing || !p.s.n.applied || clinicNews(p.s.n)) return false;
    await X.writeClinicWorkApplied(R.CLINIC_APPLIED_COLS, p.s.applied);
    return true;
  }
  async function loadInbox() {
    const box = $('inboxBox');
    if (!box) return;
    let p;
    const check = '<button class="link" id="btnInboxCheck">Check again</button>';
    try { p = await inboxPlan(); await fileQuietRows(p).catch((e) => console.warn('Clinic inbox: unchanged rows not filed', e)); }
    catch (e) { console.warn('Clinic inbox not read', e); box.innerHTML = `<p class="warn">Clinic inbox not read: ${esc(e.message)}</p><div class="actions">${check}</div>`; return; }
    const s = (k, one, many) => `${k} ${k === 1 ? one : many}`;
    if (p.missing) {
      box.innerHTML = `<p><b>Clinic inbox</b>: not set up. With it, a Power Automate flow copies each clinic's changes here when the clinic saves its file in the Clinic Eligibility library, so nobody sends files back.</p>
        <div class="actions"><button class="small" id="btnInboxSetup">Set up the clinic inbox</button>${check}</div>`;
      return;
    }
    const k = clinicNews(p.s.n);
    box.innerHTML = `<p><b>Clinic inbox</b>: ${k ? `${s(k, 'new clinic change', 'new clinic changes')} waiting${p.s.n.locations > 1 ? ` from ${p.s.n.locations} clinics` : ''}.` : 'nothing new.'}${p.skipped ? ` <span class="hint">${s(p.skipped, 'row is', 'rows are')} blank or of a kind this version does not know, and ${p.skipped === 1 ? 'is' : 'are'} skipped.</span>` : ''}</p>
      <div class="actions">${k ? '<button class="primary small" id="btnInboxApply">Apply clinic changes</button>' : ''}${check}<button class="link" data-goto="${esc(X.INBOX ? X.INBOX.sheet : 'Clinic Work Inbox')}">Clinic Work Inbox</button></div>`;
  }
  async function applyInbox() {
    const p = await inboxPlan();
    if (p.missing) throw new Error('Set up the clinic inbox first.');
    const n = p.s.n, k = clinicNews(n), s = (x, one, many) => `${x} ${x === 1 ? one : many}`;
    if (!k) { await loadInbox(); note('clinicFilesNote', 'Clinic inbox: nothing to change.', 'ok'); return; }
    await saveClinicMerge(p.s);
    // Rows applied earlier are "old" on every read of the inbox: not news here.
    note('clinicFilesNote', bringBackSummary({ ...n, old: 0 }, `Clinic inbox: ${s(k, 'change', 'changes')} from ${s(n.locations, 'clinic', 'clinics')} applied.`), 'ok');
  }
  async function setupInbox() {
    const made = await X.createClinicInbox(R.INBOX_COLS);
    await loadInbox();
    const t = X.INBOX || { sheet: 'Clinic Work Inbox', table: 'tblClinicWorkInbox' };
    note('clinicFilesNote', `${made ? 'Clinic inbox set up.' : 'The clinic inbox was already set up.'} Give the Power Automate flow: this workbook, the sheet ${t.sheet} and the table ${t.table}. The flow adds rows and the add-in only reads them, so do not sort that sheet or type in it. Build steps: docs/clinic-inbox-flow.md in the add-in repo.`, 'ok');
  }

  // ---------- clinic rules (R5-R10): flags, Review sheet, payer evidence, clinic view ----------
  const inScopePatients = () => (state.patients || []).filter((p) => p['In Scope'] === 'Yes');
  function rulesCfg() {
    const c = state.cfg || {};
    return { today: new Date(), sfhpProviders: c.sfhpProviders || [], reverifyDays: c.reverifyDays != null ? c.reverifyDays : 7, reverifyNewMonth: !!(c.reverify && c.reverify.newMonth),
      payerRules: state.payerRules || [], locationState: c.locationState || {} };
  }

  // ---------- esketamine payer credential rules (R13): the editable Payer Credential Rules sheet ----------
  /** Read the sheet (staff may have edited it). Missing sheet = no rules, and the rule stays silent. */
  async function loadPayerRules() {
    try { state.payerRules = R.payerRulesFrom(await X.readPayerCredRules()); }
    catch (e) { console.warn('Payer Credential Rules not read', e); state.payerRules = []; }
    renderCredStatus();
    return state.payerRules;
  }
  /** Import the shared CSV (or an .xlsx copy) and replace the sheet. The file stays in OneDrive; only the sheet changes. */
  async function uploadPayerRules(files) {
    const file = files[0];
    if (!file) return;
    if (!PRM) throw new Error('payer-rules.js did not load. Close and reopen the pane.');
    dropBusy('dropCred', `Reading ${file.name}…`);
    try {
      let res;
      if (isCSV(file)) res = PRM.read(C.parseCSV(await readText(file)));
      else if (isXlsx(file)) {
        const wb = await XL.readWorkbook(new Uint8Array(await file.arrayBuffer()));
        const tries = wb.sheets.filter((s) => !s.hidden).map((s) => PRM.read(s.rows));
        res = tries.find((r) => r.rules && r.rules.length) || tries.find((r) => r.rules) || tries[0] || { error: 'The workbook has no sheets.' };
      } else throw new Error(`${file.name} is not a .csv or .xlsx file.`);
      if (res.error) throw new Error(`${file.name}: ${res.error} The Payer Credential Rules sheet was not changed.`);
      if (!res.rules.length) throw new Error(`${file.name} has no payer rows. The Payer Credential Rules sheet was not changed.`);
      await X.writePayerCredRules(PRM.HEADERS, R.payerRuleRows(res.rules));
      state.payerRules = res.rules;
      await refreshRules();
      const payers = new Set(res.rules.map((r) => r.payer)).size;
      renderCredStatus();
      note('credStatus', `${file.name}: ${res.rules.length} rows for ${payers} payer${payers === 1 ? '' : 's'} saved to Payer Credential Rules${res.skipped ? `; ${res.skipped} row${res.skipped === 1 ? '' : 's'} skipped (no MD or NP credential)` : ''}.`, 'ok');
    } finally { dropBusy('dropCred'); }
  }
  function renderCredStatus() {
    const box = $('credStatus');
    if (!box) return;
    const rules = state.payerRules || [];
    if (!rules.length) { box.innerHTML = '<p class="hint">No Payer Credential Rules sheet in this workbook yet, so nothing is checked. Import the file to turn it on.</p>'; return; }
    const payers = new Set(rules.map((r) => r.payer)).size;
    const n = { ESK_CRED: 0, ESK_STATE: 0, BESSE_NO: 0, BESSE_PENDING: 0 };
    for (const fl of (state.flags && state.flags.values()) || []) for (const f of fl) if (n[f.code] != null) n[f.code]++;
    const parts = [`<p class="ok">${rules.length} rules for ${payers} payer${payers === 1 ? '' : 's'} on the Payer Credential Rules sheet.</p>`];
    if (state.patients) {
      parts.push(`<p>Upcoming esketamine visits: ${n.ESK_CRED} with a provider the payer doesn't allow, ${n.ESK_STATE} to check the plan's state, ${n.BESSE_NO} patient${n.BESSE_NO === 1 ? '' : 's'} not Besse-approved, ${n.BESSE_PENDING} Besse pending.</p>`);
    }
    box.innerHTML = parts.join('');
  }
  function computeFlags() {
    state.flags = state.patients ? R.allFlags(inScopePatients(), state.results || [], rulesCfg()) : new Map();
    return state.flags;
  }
  /** Per Nbr -> secondary payer (chart, else the one billed at the last encounter), for the tracker file. */
  function secondaryByPer() {
    const m = new Map();
    for (const p of state.patients || []) {
      const sec = String(p['Secondary Payer'] || '').trim() || String((p.x && p.x.secPayerLstEnc) || '').trim();
      if (sec) m.set(String(p['Per Nbr']), sec);
    }
    return m;
  }
  const pillsFor = (per) => [...new Set((state.flags.get(String(per ?? '').trim()) || []).map((f) => f.label))].join('; ');

  /** Recompute flags; write the Review sheet (build review + flags) and Payer Evidence; refresh the Clinic tab. */
  async function refreshRules() {
    if (!state.patients) return;
    computeFlags();
    const pts = inScopePatients();
    const flagRows = R.reviewRows(state.flags, pts);
    // One High row for Sohar results with no patient identifiers (no NextGen task each).
    const review = (state.review || []).concat(C.noIdentifierReviewRows(state.results || []), flagRows);
    await X.replaceTable(X.TABLES.review, C.REVIEW_COLS, review);
    const sev = { High: 0, Medium: 0, Low: 0 };
    for (const r of review) sev[r.Severity] = (sev[r.Severity] || 0) + 1;
    const bi = $('buildIssues');
    if (bi) bi.innerHTML = sev.High ? `<span class="hi">${sev.High} high-priority item${sev.High === 1 ? '' : 's'}</span> on the Review sheet${flagRows.length ? ` (${flagRows.length} clinic flag${flagRows.length === 1 ? '' : 's'} included)` : ''}.` : '';
    try {
      const prev = await X.readPayerEvidence();
      state.evidence = R.payerEvidence(pts, state.results || [], prev, new Date(), state.perToKey);
      await X.writePayerEvidence(R.PAYER_EVIDENCE_COLS, state.evidence);
    } catch (e) { console.error(e); note('clinicFilesNote', `Payer Evidence sheet not updated: ${e.message}`, 'warn'); }
    const run = new Date();
    state.clinic = {
      rows: R.clinicRows('All', pts, state.results || [], state.flags, run),
      review: R.clinicReviewRows('All', pts, state.results || [], state.flags, run),
      runDate: C.fmtDate(new Date(Date.UTC(run.getFullYear(), run.getMonth(), run.getDate()))),
    };
    renderClinic();
    renderClinicFiles();
    renderEvidence();
    renderCredStatus();
    await refreshLearningViews();
    renderGuarantorPrompt();
  }

  /**
   * Payer metrics (today, plus one history row per payer and day), guarantor themes, and the pane's Payer learning
   * section. The sheets have no patient identifiers.
   */
  async function refreshLearningViews() {
    const warn = [];
    try {
      const history = await X.readPayerMetricHistory();
      const today = C.localDate(new Date());
      const first = C.payerMetrics(state.patients || [], state.results || [], state.payerTests || [], history);
      const hist = C.appendMetricHistory(history, first.rows, today);
      state.metrics = C.payerMetrics(state.patients || [], state.results || [], state.payerTests || [], hist);
      await X.writePayerMetrics(C.PAYER_METRIC_COLS, state.metrics.rows);
      if (first.rows.length || history.some((h) => String(h.Date) === today)) await X.writePayerMetricHistory(C.PAYER_METRIC_HISTORY_COLS, hist);
    } catch (e) { console.error(e); warn.push(`Payer Metrics not updated: ${e.message}`); }
    state.themes = C.guarantorThemes(state.results || []);
    if (state.themes.rows.length) await X.writeGuarantorThemes(C.GUARANTOR_THEME_COLS, state.themes.rows).catch((e) => { console.error(e); warn.push(`Guarantor Themes not updated: ${e.message}`); });
    renderLearning(warn);
  }
  function renderLearning(warn) {
    const box = $('learnBox');
    if (!box) return;
    const m = state.metrics, t = m ? m.totals : null, th = state.themes || { counts: {} };
    const pt = (state.cfg && state.cfg.payerTests) || C.PAYER_TEST_DEFAULTS;
    const sugg = (state.learning || []).filter((r) => r.suggestion);
    const top = m ? m.rows.slice(0, 8) : [];
    const themeLine = Object.entries(th.counts || {}).map(([k, v]) => `${esc(k)}: ${v}`).join(' · ');
    box.innerHTML = (t ? `<div class="counts q4">
        <div class="count unm"><div class="n">${t.unmappedAppts}</div><div class="l">appts, payer unmapped<br>(${t.unmappedPatients} patients)</div></div>
        <div class="count act"><div class="n">${t.failedAppts}</div><div class="l">appts, verification failed<br>(${t.failedPatients} patients)</div></div>
        <div class="count pen"><div class="n">${t.testsBack ? `${Math.round((100 * t.testsEligible) / t.testsBack)}%` : '—'}</div><div class="l">payer tests eligible<br>(${t.testsEligible} of ${t.testsBack} back)</div></div>
        <div class="count clr"><div class="n">${t.recurring}</div><div class="l">payers recurring<br>(${C.RECURRING_DAYS}+ days)</div></div>
      </div>` : '')
      + (top.length ? `<table class="watch"><thead><tr><th>NextGen payer</th><th>Unmapped</th><th>Failed</th><th>Tests</th></tr></thead><tbody>${top.map((r) => `<tr>
        <td>${esc(r['NextGen Payer'])}${r.Recurring ? ' <span class="pill p-amber">Recurring</span>' : ''}<br><span class="hint">${esc(r['Most common failure'] || r.Confidence || '')}</span></td>
        <td>${esc(r['Unmapped appts'])}</td><td>${esc(r['Failed verification appts'])}</td><td>${r['Payer tests back'] !== '0' ? `${esc(r['Payer tests eligible'])}/${esc(r['Payer tests back'])}${r['Best test payer'] ? `<br><span class="hint">${esc(r['Best test payer'])}</span>` : ''}` : '—'}</td></tr>`).join('')}</tbody></table>` : '<p class="hint">No unmapped or failed payers in this report.</p>')
      + `<p class="hint">${sugg.length ? `${sugg.length} payer suggestion${sugg.length === 1 ? '' : 's'} wait${sugg.length === 1 ? 's' : ''} for you in step 1 (Use ...). ` : ''}Payer mapping discovery is ${pt.mappingDiscovery !== false ? 'On' : 'Off'}; payer tests are ${pt.on ? `On (${pt.payers} payers × ${pt.patients} patients)` : 'Off (Config "Payer tests")'}. A discovery that finds a different payer only counts when Sohar returns the same member ID.</p>`
      + (themeLine ? `<p><b>Guarantor searches:</b> ${themeLine}.</p>` : '')
      + (warn || []).map((w) => `<p class="warn">${esc(w)}</p>`).join('')
      + '<div class="actions"><button class="link" data-goto="Payer Metrics">Payer Metrics</button><button class="link" data-goto="Payer Metrics History">History</button><button class="link" data-goto="Payer Learning">Payer Learning</button><button class="link" data-goto="Payer Mapping">Payer Mapping</button>' + (themeLine ? '<button class="link" data-goto="Guarantor Themes">Guarantor Themes</button>' : '') + '</div>';
  }

  // ---------- Clinic tab (R11) ----------
  const PILL_BY_LABEL = new Map(R.PILLS.map((p) => [p.label, p]));
  const pillHtml = (label, title) => { const p = PILL_BY_LABEL.get(label); return `<span class="pill p-${p ? p.color : 'grey'}"${title ? ` title="${esc(title)}"` : ''}>${esc(label)}</span>`; };
  const splitList = (s, sep) => String(s ?? '').split(sep).map((x) => x.trim()).filter(Boolean);
  const money = (v) => { const n = parseFloat(String(v ?? '').replace(/[$,]/g, '')); return isNaN(n) ? null : n; };
  const sameLoc = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
  /** Clinic tab only: Sohar status codes in plain words (Admin and full screen keep the raw codes). */
  function plainStatus(s) {
    const t = String(s || '').trim();
    if (!t || t === 'Not in Sohar file') return 'Not checked yet';
    if (/^complete\.eligible/i.test(t)) return 'Eligible';
    if (/^complete\.ineligible/i.test(t)) return 'Not eligible';
    if (/^pending/i.test(t)) return 'Pending';
    const e = t.match(/^error\.?([^(\s]*)/i);
    if (e) return `Error${e[1] ? `: ${e[1].replace(/[._]+/g, ' ')}` : ''}`;
    if (/^coverage found/i.test(t)) return 'Insurance found';
    if (/^no coverage/i.test(t)) return 'No insurance found';
    return t;
  }
  /** Clinic tab only: "Not in Payer Map: Health Net (95567)" becomes words front-desk staff can act on. */
  function plainChoose(c) {
    const m = String(c || '').match(/^Not in Payer Map: (.*?)(?: \([^)]*\))?$/);
    return m ? { text: `Insurance found: ${m[1]}. Ask billing which NextGen plan to pick.`, quoted: false } : { text: c, quoted: true };
  }
  const SEV_WORD = { high: 'Action', medium: 'Check' };
  /** Card review items with severity; per-appointment repeats collapsed (rules.collapseFlags). */
  function cardItems(per, r) {
    const rv = ((state.clinic && state.clinic.review) || []).filter((x) => String(x.MRN || C.perKey(x['Per Nbr'])) === per && x.Severity !== 'Low');
    if (!rv.length) return splitList(r['Review Items'], '|').map((t) => ({ detail: t, severity: '' }));
    return R.collapseFlags(rv.map((x) => ({ code: x.Issue, label: x.Pill, severity: String(x.Severity || '').toLowerCase(), detail: x.Detail,
      appt: { date: x['Appt Date'], event: x.Event, rendering: x.Rendering } })));
  }
  /** "Kaiser/VA MM" is a filter category; on a card it only repeats "Kaiser/VA MM no auth". */
  const cardPills = (pills) => (pills.includes('Kaiser/VA MM no auth') ? pills.filter((l) => l !== 'Kaiser/VA MM') : pills);

  function wireClinic() {
    $('clinicLoc').addEventListener('change', () => {
      renderClinic();
      if (!state.clinicFile && !state.demo) X.setConfigValue('Clinic location', $('clinicLoc').value || 'All').catch(() => {});
    });
    $('clinicText').addEventListener('input', () => renderClinicList());
    $('clinicList').addEventListener('change', (e) => {
      const cb = e.target.closest && e.target.closest('input[data-rv]');
      if (!cb) return;
      cb.disabled = true;
      guard(() => toggleClinicRequest(cb.dataset.rv, cb.checked), 'clinicNote').finally(() => { cb.disabled = false; });
    });
    $('clinicChips').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-pill]'); if (!b) return;
      state.clinicFilter = state.clinicFilter === b.dataset.pill ? '' : b.dataset.pill;
      renderClinicChips(); renderClinicList();
    });
    $('btnClinicEmail').addEventListener('click', () => guard(async () => {
      const loc = $('clinicLoc').value;
      const text = R.clinicEmail(loc, clinicRowsAt(loc), { runDate: runDateOf() });
      await copyOrShow(text, 'clinicNote', `Email summary for ${loc || 'all locations'} copied (LASFIR + MRN only). Paste it into Outlook or Teams.`);
    }, 'clinicNote'));
  }
  function runDateOf() {
    const d = state.clinic && C.parseDate(state.clinic.runDate);
    return d ? new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) : new Date();
  }
  function clinicRowsAt(loc) {
    const rows = (state.clinic && state.clinic.rows) || [];
    return loc ? rows.filter((r) => sameLoc(r.Location, loc)) : rows;
  }

  /** Location picker from the data (remembers Config "Clinic location" in the template), then chips and list. */
  function renderClinic(prefer) {
    const sel = $('clinicLoc');
    const rows = (state.clinic && state.clinic.rows) || [];
    const locs = [...new Set(rows.map((r) => String(r.Location || '').trim()).filter(Boolean))].sort();
    const want = prefer != null ? prefer : (sel.dataset.ready ? sel.value : ((state.cfg && state.cfg.clinicLocation) || ''));
    sel.innerHTML = '<option value="">All locations</option>' + locs.map((l) => `<option>${esc(l)}</option>`).join('');
    const hit = locs.find((l) => sameLoc(l, want));
    sel.value = hit || '';
    sel.dataset.ready = '1';
    renderClinicChips();
    renderClinicList();
  }
  function renderClinicChips() {
    const rows = clinicRowsAt($('clinicLoc').value);
    const count = new Map();
    for (const r of rows) for (const l of splitList(r.Pills, ';')) count.set(l, (count.get(l) || 0) + 1);
    const chips = R.PILLS.filter((p) => count.has(p.label));
    if (state.clinicFilter && !count.has(state.clinicFilter)) state.clinicFilter = '';
    $('clinicChips').innerHTML = rows.length
      ? `<button class="chipbtn" data-pill="" aria-pressed="${!state.clinicFilter}">All ${rows.length}</button>` + chips.map((p) => `<button class="chipbtn p-${p.color}" data-pill="${esc(p.label)}" aria-pressed="${state.clinicFilter === p.label}">${esc(p.label)} ${count.get(p.label)}</button>`).join('')
      : '';
  }
  function meterHtml(label, total, remaining, status) {
    if (status === 'None') return `<div class="mrow"><span class="ml">${label}</span><span class="hint">None</span></div>`;
    const t = money(total), r = money(remaining);
    if (t == null || r == null || t <= 0) return status ? `<div class="mrow"><span class="ml">${label}</span><span>${esc(status)}</span></div>` : '';
    const pct = Math.max(0, Math.min(1, (t - r) / t));
    const cls = status === 'Met' ? 'met' : status === 'Approaching' ? 'appr' : '';
    return `<div class="mrow ${cls}"><span class="ml">${label}</span><span class="meter" data-pct="${pct}"><span></span></span><span>${status === 'Met' ? '<b class="okc">Met</b>' : `${esc(usd(r))} left`} <span class="hint">of ${esc(usd(t))}</span></span></div>`;
  }
  function renderClinicList() {
    const loc = $('clinicLoc').value;
    const q = $('clinicText').value.trim().toLowerCase();
    const all = clinicRowsAt(loc);
    const rows = all.filter((r) => (!state.clinicFilter || splitList(r.Pills, ';').includes(state.clinicFilter))
      && (!q || [r['Last Name'], r['First Name'], r.MRN, r['Per Nbr']].some((v) => String(v || '').toLowerCase().includes(q))));
    const list = $('clinicList');
    const meta = $('clinicMeta');
    if (!state.clinic) { meta.textContent = ''; return; }
    meta.textContent = `${rows.length} of ${all.length} patient${all.length === 1 ? '' : 's'} with an upcoming visit${loc ? ` at ${loc}` : ''}${state.clinic.runDate ? `. Run ${state.clinic.runDate}` : ''}.`;
    if (!rows.length) { list.innerHTML = `<p class="hint">${all.length ? 'No one matches this filter.' : 'No upcoming patients here.'}</p>`; return; }
    const days = new Map();
    for (const r of rows) { const d = r['Next Appt'] || 'No date'; if (!days.has(d)) days.set(d, []); days.get(d).push(r); }
    const wd = (s) => { const d = C.parseDate(s); return d ? d.toLocaleDateString('en-US', { weekday: 'short', timeZone: 'UTC' }) + ' ' : ''; };
    list.innerHTML = [...days].map(([d, rs]) => `<h3 class="cday">${esc(wd(d) + d)} <span class="hint">${rs.length}</span></h3>` + rs.map((r) => {
      const per = clinicKey(r);
      const pills = splitList(r.Pills, ';');
      const items = cardItems(per, r);
      const choose = plainChoose(r['Choose in NextGen']);
      const notes = splitList(r.Notes, '|');
      const high = pills.some((l) => (PILL_BY_LABEL.get(l) || {}).severity === 'high');
      const reverify = r['Re-verify'] === 'Yes';
      const ver = r['Verified Date'] ? `Verified ${esc(r['Verified Date'])}${r['Days Since Verified'] !== '' && r['Days Since Verified'] != null ? ` (${esc(r['Days Since Verified'])}d)` : ''}` : 'Not verified in Sohar';
      const ben = meterHtml('Deductible', r.Deductible, r['Deductible Remaining'], r['Deductible Status']) + meterHtml('OOP max', r['OOP Max'], r['OOP Remaining'], r['OOP Status']);
      return `<article class="pcard${high ? ' hi' : ''}">
        <div class="pc-head"><span class="pc-time">${esc(r['Next Time'] || '')}</span><b>${esc([r['Last Name'], r['First Name']].filter(Boolean).join(', '))}</b><span class="hint">${esc(C.idLabel(r))}</span></div>
        <div class="pc-sub">${esc([r['Next Event'], r['Next Provider']].filter(Boolean).join(' · '))}${!loc && r.Location ? ` <span class="hint">· ${esc(r.Location)}</span>` : ''}</div>
        ${pills.length ? `<div class="pills">${cardPills(pills).map((l) => pillHtml(l)).join(' ')}</div>` : ''}
        <dl class="pc-dl">
          <dt>Payer</dt><dd>${esc(r['Primary Payer'] || 'None on file')}${r['Secondary Payer'] ? ` <span class="hint">+ ${esc(r['Secondary Payer'])}</span>` : ''}</dd>
          ${r['Choose in NextGen'] ? `<dt>NextGen</dt><dd>${choose.quoted ? `Choose "${esc(choose.text)}"` : esc(choose.text)}</dd>` : ''}
          <dt>Sohar</dt><dd>${esc(plainStatus(r['Sohar Status']))}${r.Plan ? ` · ${esc(r.Plan)}` : ''}<br><span class="hint">${ver}</span></dd>
          ${r.Copay !== '' || r.Coinsurance !== '' ? `<dt>Copay</dt><dd>${r.Copay !== '' ? esc(usd(r.Copay) || r.Copay) : '—'}${r.Coinsurance !== '' ? ` · coins ${esc(r.Coinsurance)}` : ''}</dd>` : ''}
          ${feeCardHtml(r['Per Nbr'])}
        </dl>
        ${ben ? `<div class="pc-ben">${ben}</div>` : ''}
        ${r['Verification Skipped'] || r['Re-verify Anyway'] ? `<div class="pc-skip"><span class="hint">${r['Verification Skipped'] ? `${esc(r['Verification Skipped'])}: not on this week's verification file (next check ${esc(r['Next Check'] || '?')}).` : 'Verified recently.'}</span>
          <label class="inline"><input type="checkbox" data-rv="${esc(per)}" ${r['Re-verify Anyway'] ? 'checked' : ''}> Re-verify anyway</label></div>` : ''}
        ${items.length ? `<ul class="pc-items">${items.map((t) => `<li class="sev-${t.severity || 'high'}">${SEV_WORD[t.severity] ? `<b>${SEV_WORD[t.severity]}:</b> ` : ''}${esc(t.detail)}</li>`).join('')}</ul>` : ''}
        ${notes.length ? `<details class="more inline"><summary>Notes (${notes.length})</summary><ul class="pc-notes">${notes.map((t) => `<li>${esc(t)}</li>`).join('')}</ul></details>` : ''}
        ${reverify || /^https:\/\//.test(r['Sohar Link'] || '') ? `<div class="pc-links">${[
          reverify ? `<a href="${esc(R.REVERIFY_URL)}" target="_blank" rel="noopener" class="due">Re-verify in Sohar</a>` : '',
          /^https:\/\//.test(r['Sohar Link'] || '') ? `<a href="${esc(r['Sohar Link'])}" target="_blank" rel="noopener">Sohar result</a>` : '',
        ].filter(Boolean).join(' · ')}</div>` : ''}
      </article>`;
    }).join('')).join('');
    list.querySelectorAll('.meter').forEach((m) => { m.firstElementChild.style.width = `${Math.round(Math.min(1, +m.dataset.pct) * 100)}%`; });
  }

  // ---------- Admin: clinic files, emails, payer evidence (R10, R12) ----------
  function renderClinicFiles() {
    const box = $('clinicFiles');
    const rows = (state.clinic && state.clinic.rows) || [];
    const locs = [...new Set(rows.map((r) => String(r.Location || '').trim()).filter(Boolean))].sort();
    if (!locs.length) { box.innerHTML = '<p class="hint">No patients with an upcoming visit yet.</p>'; return; }
    // The three buttons get their own row under each location so the table fits a 350 px pane.
    box.innerHTML = `<table class="watch clinicfiles"><thead><tr><th>Location</th><th>Patients</th><th>Flagged</th></tr></thead><tbody>${locs.map((l) => {
      const rs = clinicRowsAt(l);
      const flagged = rs.filter((r) => splitList(r['Review Items'], '|').length).length;
      return `<tr class="cf-loc"><td><b>${esc(l)}</b></td><td>${rs.length}</td><td>${flagged ? `<span class="hi">${flagged}</span>` : '0'}</td></tr>
        <tr><td class="acts" colspan="3"><button class="small" data-cfile="${esc(l)}">Save clinic file</button> <button class="small" data-ctt="${esc(l)}">Clinic Template Tool file</button> <button class="small" data-cemail="${esc(l)}">Copy email</button></td></tr>`;
    }).join('')}</tbody></table>`;
    box.querySelectorAll('[data-cfile]').forEach((b) => b.addEventListener('click', () => guard(() => saveClinicFile(b.dataset.cfile), 'clinicFilesNote')));
    box.querySelectorAll('[data-ctt]').forEach((b) => b.addEventListener('click', () => guard(() => saveCttFile(b.dataset.ctt), 'clinicFilesNote')));
    box.querySelectorAll('[data-cemail]').forEach((b) => b.addEventListener('click', () => guard(async () => {
      const l = b.dataset.cemail;
      const text = R.clinicEmail(l, clinicRowsAt(l), { runDate: runDateOf() });
      await copyOrShow(text, 'clinicFilesNote', `Email summary for ${l} copied (LASFIR + MRN only). Paste it into Outlook or Teams.`);
    }, 'clinicFilesNote')));
  }

  /** "Sohar Clinic - <Location> YYYY-MM-DD.xlsx": Clinic, Review and a hidden Clinic Info sheet, plus the pane auto-open parts. */
  function clinicWorkbook(loc) {
    const pts = inScopePatients();
    const run = new Date();
    const rows = R.clinicRows(loc, pts, state.results || [], state.flags, run);
    const review = R.clinicReviewRows(loc, pts, state.results || [], state.flags, run);
    const runDate = C.fmtDate(new Date(Date.UTC(run.getFullYear(), run.getMonth(), run.getDate())));
    const bytes = XL.buildWorkbook({
      sheets: [
        { name: 'Clinic', table: 'tblClinic', columns: R.CLINIC_COLS, rows, linkColumns: ['Re-verify Link', 'Sohar Link'], linkText: { 'Re-verify Link': 'Re-verify in Sohar' } },
        { name: 'Review', table: 'tblClinicReview', columns: R.CLINIC_REVIEW_COLS, rows: review, linkColumns: ['Link'], linkText: { Link: 'Re-verify in Sohar' } },
        { name: 'Clinic Requests', table: 'tblClinicRequests', columns: R.CLINIC_REQUEST_COLS,
          rows: rows.filter((r) => r['Re-verify Anyway']).map((r) => ({ MRN: C.keyOf(r), 'Last Name': r['Last Name'], 'First Name': r['First Name'], 'Next Appt': r['Next Appt'], 'Re-verify Anyway': 'On', Requested: '' })) },
        { name: 'Clinic Info', table: 'tblClinicMeta', columns: ['Kind', 'Location', 'Run Date', 'Add-in Version'], rows: [{ Kind: 'Sohar Clinic', Location: loc, 'Run Date': runDate, 'Add-in Version': VERSION }], hidden: true },
      ],
      webExtension: { id: ADDIN_ID, version: '0.1.1.0' },
      now: run,
    });
    const name = `Sohar Clinic - ${String(loc).replace(/[\\/:*?"<>|]/g, '-')} ${C.localDate(run)}.xlsx`;
    return { name, bytes, count: rows.length, flags: review.length };
  }
  async function saveClinicFile(loc) {
    if (!state.clinic) throw new Error('Upload the NextGen report first.');
    const f = clinicWorkbook(loc);
    await saveBinary(f.name, f.bytes, XL.MIME, 'clinicFilesNote', `${f.count} patient${f.count === 1 ? '' : 's'}, ${f.flags} flag row${f.flags === 1 ? '' : 's'}`);
  }

  /**
   * "Clinic Eligibility - <Location> YYYY-MM-DD.xlsx" (v0.14) for the Clinic Template Tool, which opens it and re-derives
   * the list with the same rules.js. rules.clinicPack keeps only this clinic's patients and strips names and DOBs.
   */
  async function cttWorkbook(loc) {
    const pack = R.clinicPack(loc, {
      patients: inScopePatients(), results: state.results || [], review: state.review || [],
      learningRows: C.payerLearningRows(state.learning || []), metricRows: state.metrics ? state.metrics.rows : [], tasks: await liveTasks(),
      cfg: state.cfg, payerRuleRows: R.payerRuleRows(state.payerRules || []), runAt: new Date(), version: VERSION,
    });
    return { name: pack.name, bytes: XL.buildWorkbook(pack), counts: pack.counts };
  }
  async function saveCttFile(loc) {
    if (!state.clinic) throw new Error('Upload the NextGen report first.');
    const f = await cttWorkbook(loc), c = f.counts;
    await saveBinary(f.name, f.bytes, XL.MIME, 'clinicFilesNote', `${c.patients} patient${c.patients === 1 ? '' : 's'}, ${c.results} Sohar result${c.results === 1 ? '' : 's'}, for the Clinic Template Tool; PHI: share through MHS OneDrive or Teams only`);
  }

  /** Admin: counts from returned Clinic Eligibility files, and the payer notes under Payer learning for a person to act on. */
  async function loadClinicNotes() {
    try { state.clinicNotes = await X.readClinicNotes(); state.clinicPayerNotes = await X.readClinicPayerNotes(); }
    catch (e) { console.warn('Clinic Notes not read', e); }
    renderClinicNotes();
  }
  function renderClinicNotes() {
    const notes = state.clinicNotes || [], pn = state.clinicPayerNotes || [];
    const box = $('clinicNotesBox'), pbox = $('payerNotesBox');
    if (box) {
      box.innerHTML = notes.length || pn.length ? `<p>From returned Clinic Template Tool files: ${notes.length} patient note${notes.length === 1 ? '' : 's'} and ${pn.length} payer note${pn.length === 1 ? '' : 's'}${pn.length ? ' (listed under Payer learning)' : ''}.</p>
        <div class="actions">${notes.length ? '<button class="link" data-goto="Clinic Notes">Clinic Notes</button>' : ''}${pn.length ? '<button class="link" data-goto="Clinic Payer Notes">Clinic Payer Notes</button>' : ''}</div>` : '';
    }
    if (!pbox) return;
    pbox.hidden = !pn.length;
    if (!pn.length) { pbox.innerHTML = ''; return; }
    const when = (w) => { const d = new Date(w); return isNaN(d) ? '' : C.localDate(d); };
    const list = pn.slice().sort((a, b) => String(b.When || '').localeCompare(String(a.When || '')));
    const top = list.slice(0, 12);
    pbox.innerHTML = `<p><b>${pn.length} payer note${pn.length === 1 ? '' : 's'} from clinics.</b> What the clinic saw on the insurance card. Check each one, then fix the Payer Map or the patient's insurance in NextGen yourself; a note never changes the mapping.</p>
      <table class="watch"><thead><tr><th>NextGen payer</th><th>Card says</th><th>From</th></tr></thead><tbody>${top.map((r) => `<tr>
        <td>${esc(r['NextGen Payer'])}</td><td>${esc(r['Card Says'])}${r.Note ? `<br><span class="hint">${esc(r.Note)}</span>` : ''}</td><td>${esc(r.Location)}<br><span class="hint">${esc([r.By, when(r.When)].filter(Boolean).join(' · '))}</span></td></tr>`).join('')}</tbody></table>
      ${list.length > top.length ? `<p class="hint">${list.length - top.length} more on the sheet.</p>` : ''}
      <div class="actions"><button class="link" data-goto="Clinic Payer Notes">Clinic Payer Notes sheet</button></div>`;
  }

  function renderEvidence() {
    const box = $('evidenceBox');
    const ev = state.evidence || [];
    if (!ev.length) { box.innerHTML = '<p class="hint">Nothing recorded yet.</p>'; return; }
    const sum = R.evidenceSummary(ev);
    const top = sum.slice(0, 12);
    box.innerHTML = `<p>${ev.length} observation${ev.length === 1 ? '' : 's'} on the Payer Evidence sheet; ${sum.length} payer × prefix × location group${sum.length === 1 ? '' : 's'}.</p>
      <table class="watch"><thead><tr><th>NextGen payer</th><th>Prefix</th><th>Location</th><th>#</th><th>Sohar said</th></tr></thead><tbody>${top.map((g) => `<tr>
        <td>${esc(g.nextgenPayer)}</td><td>${esc(g.prefix || '—')}</td><td>${esc(g.location || '—')}</td><td>${g.count}</td><td>${esc(g.soharPayers.join('; ') || '—')}</td></tr>`).join('')}</tbody></table>
      ${sum.length > top.length ? `<p class="hint">${sum.length - top.length} more groups on the sheet.</p>` : ''}`;
  }

  /** Copy text; if the pane blocks the clipboard, show it in a box to select by hand. */
  async function copyOrShow(text, boxId, okMsg) {
    if (await copyText(text)) { note(boxId, okMsg, 'ok'); return; }
    const box = $(boxId); box.hidden = false;
    const p = note(boxId, 'Copy is blocked here. Select the text below and press Ctrl+C.', 'warn');
    const ta = document.createElement('textarea'); ta.className = 'copybox'; ta.readOnly = true; ta.rows = 10; ta.value = text;
    p.after(ta); ta.focus(); ta.select();
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

  /**
   * Learned Fin Classes for the merge: the ones the last build used (they made the imports), else (no build yet this
   * session) learned from the ledger now, none when Learning is Off.
   */
  async function learnedFinClassesFor(cfg) {
    if (state.finClasses) return state.finClasses;
    const learn = (cfg && cfg.learning) || C.LEARN_DEFAULTS;
    if (!learn.on) return [];
    if (!state.outcomes) state.outcomes = await X.readPayerOutcomes().catch((e) => { console.error(e); return null; });
    return C.learnFinClasses(state.outcomes || [], learn);
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
    // Sohar ref memory: a re-export with blank identifiers keys and matches like the record its ref was first seen on.
    // A read error means no refs this time (and none saved), never an overwrite.
    let refs = null;
    try { refs = await X.readSoharRefs(); } catch (e) { console.error(e); note('resultsResult', `Sohar Refs not read (rows with blank identifiers may not match this time): ${e.message}`, 'warn'); }
    const saved = await X.readUploads();
    const first = C.accumulateResults(saved, list, '', refs || []);
    if (list.length && !first.added && !first.replaced) throw new Error(`None of these files look like a Sohar verification or discovery output${first.unknownFiles.length ? ` (${first.unknownFiles.join(', ')})` : ''}.`);
    if (!first.files.length) return '';
    step('Reading the appointment report on the Import sheet…');
    state.rows = await X.readImportRows();
    const cfg = await X.readConfig();
    if ($('asOf').value) cfg.asOf = new Date($('asOf').value + 'T00:00:00Z');
    if (!state.soharPayers) state.soharPayers = await X.readSoharPayers();
    // Same learned Fin Classes as the build that made the imports, so a patient sent to verification through a learned
    // Fin Class matches its last-encounter row here (and its outcome is recorded against that Fin Class).
    // combineResults matches again when a blank re-export's ref is only learned in this pass, and records the payerId
    // sent on this upload's verification records (Sohar exports have none) for the skip window.
    const payerTests = await X.readPayerTests().catch(() => state.payerTests || []);
    const cr = C.combineResults(state.rows, saved, list, new Date().toISOString(), refs || [], { ...cfg, payerMap: await X.readPayerMap(), soharPayers: state.soharPayers, learnedFinClasses: await learnedFinClassesFor(cfg), payerTestRows: payerTests });
    const { acc, res } = cr;
    if (list.length) {
      step(`${acc.added} new result${acc.added === 1 ? '' : 's'}${acc.replaced ? `, ${acc.replaced} replaced` : ''}; ${acc.saved.length} saved in this workbook.`);
      await X.writeUploads(C.UPLOAD_COLS, cr.saved); // saved even when the report is missing: it matches once it is back
    }
    if (res.error) throw new Error(res.error);
    const learnedRefs = cr.refs;
    const noIds = res.counts.noIdentifiers || 0;
    step(`Matched ${res.counts.matched} results to patients (${res.counts.unmatched} unmatched${noIds ? `, ${noIds} with no patient identifiers` : ''})${res.counts.twoInsurances ? `; ${res.counts.twoInsurances} patient${res.counts.twoInsurances === 1 ? '' : 's'} had two insurances checked` : ''}. Writing the Sohar Results sheet…`);
    // Remember the refs of records matched by their id or Member ID (append-only; "Clear saved Sohar results" keeps them).
    if (refs && learnedRefs.length) {
      try {
        const kept = C.appendRefs(refs, learnedRefs, fmtStamp(new Date()), res.perToKey);
        if (kept.added || refs.some((r) => !String(r.MRN || '').trim())) await X.writeSoharRefs(C.SOHAR_REF_COLS, kept.rows);
      } catch (e) { console.error(e); note('resultsResult', `Sohar Refs not updated: ${e.message}`, 'warn'); }
    }
    // Flags need the new results (re-verify, Medi-Cal from Sohar); the Results sheet gets a Pills column after RESULTS_COLS.
    state.results = res.results;
    computeFlags();
    const secBy = secondaryByPer();
    // Pills and Secondary Payer ride along for full screen (Secondary Payer is not written to the Results sheet).
    res.results = res.results.map((r) => ({ ...r, Pills: pillsFor(C.keyOf(r)), 'Secondary Payer': secBy.get(String(r['Per Nbr'])) || '' }));
    await X.writeResults(C.RESULTS_COLS.concat('Pills'), res.results);
    const tasks = C.keepDone(C.buildTasks(res.results), await X.readTasks());
    step(`Writing ${tasks.length} rows to the NextGen Tasks sheet…`);
    await X.writeTasks(C.TASK_COLS, tasks);
    if (res.perToKey) state.perToKey = res.perToKey;
    // Payer test results go to the hidden Payer Tests sheet (never the patient's results).
    if (res.payerTests && res.payerTests.length) {
      const pt = C.recordPayerTests(payerTests, res.payerTests);
      if (pt.updated) { await X.writePayerTests(C.PAYER_TEST_COLS, pt.table).catch((e) => note('resultsResult', `Payer Tests not updated: ${e.message}`, 'warn')); state.payerTests = pt.table; step(`${pt.updated} payer test result${pt.updated === 1 ? '' : 's'} recorded.`); }
    }
    const dv = C.keepMemberIds(res.discoveryVerification, await X.readDiscoveryVerify(), state.perToKey);
    await X.writeDiscoveryVerify(C.VERIFICATION_COLS, dv);
    // Keep ZIPs an admin typed on the Discovery from Verify sheet.
    const zips = new Map((await X.readVerifyFailed()).filter((r) => String(r.zipCode || '').trim()).map((r) => [state.perToKey.get(String(r.id)) || String(r.id), String(r.zipCode).trim()]));
    const vf = res.verificationDiscovery.map((r) => (r.zipCode ? r : { ...r, zipCode: zips.get(String(r.id)) || '' }));
    await X.writeVerifyFailed(C.DISCOVERY_COLS, vf);
    state.results = res.results; state.resultsAsOf = res.asOf; state.tasks = tasks; state.discoveryVerify = dv; state.verifyFailed = vf;
    // Payer outcomes ledger (hidden Payer Outcomes sheet): every Sohar outcome, append-only, for payer learning.
    // "Clear saved Sohar results" does not touch it. A read error skips the update, so the ledger is never overwritten.
    let newOutcomes = 0;
    try {
      const before = await X.readPayerOutcomes();
      const ledger = C.appendOutcomes(before, C.payerOutcomes(res.results, res.patients, fmtStamp(new Date())), state.perToKey);
      if (ledger.added || before.some((o) => !String(o.MRN || '').trim())) {
        await X.writePayerOutcomes(C.PAYER_OUTCOME_COLS, ledger.rows);
        step(`${ledger.added} new Sohar outcome${ledger.added === 1 ? '' : 's'} recorded for payer learning.`);
      }
      newOutcomes = ledger.added;
      state.outcomes = ledger.rows;
    } catch (e) { console.error(e); note('resultsResult', `Payer Outcomes not updated (payer learning waits for the next upload): ${e.message}`, 'warn'); }
    // New outcomes can promote, demote or suggest payers: rebuild so the Payer Map, Patients and imports follow.
    // (The Sohar Results sheet picks up a changed mapping on the next refresh.)
    if (newOutcomes && state.patients) await build({ deferRules: true });
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

    // Verification failed (ineligible or error) -> discovery, to find the patient's current coverage.
    $('verFailBox').hidden = !vf.length;
    if (vf.length) {
      $('verFailText').innerHTML = `<b>${vf.length} patient${vf.length === 1 ? '' : 's'} came back ineligible or with an error.</b> Save this discovery file, upload it to Sohar, then drop Sohar's output above. They are listed on the <b>Discovery from Verify</b> sheet.`;
      $('fnVerFail').textContent = `${vf.length} patient${vf.length === 1 ? '' : 's'} · ${C.fileName('verifyFailed')}`;
    }

    const q = res.counts.byQueue, c = res.counts;
    const open = tasks.filter((t) => !t.Done).length, noted = tasks.filter((t) => String(t['Clinic Note'] || '').trim()).length;
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
        ${noted ? `<p class="hint">${noted} task${noted === 1 ? ' has a note' : 's have notes'} from the clinic (Clinic Note column on NextGen Tasks, and in full screen).</p>` : ''}
        ${noIds ? `<p class="warn"><b>${noIds} Sohar result${noIds === 1 ? ' has' : 's have'} no patient identifiers</b> (no id, member ID or name): likely an export anomaly or a check run outside the batch. Open the Sohar link on those rows (Match = No identifiers); one item is on the Review sheet.</p>` : ''}
        ${acc.unknownFiles.length ? `<p class="warn">Skipped (not a Sohar output): ${acc.unknownFiles.map(esc).join(', ')}</p>` : ''}`;
    await renderBenefits(res.results, cfg, kinds);
    if (list.length) {
      // New Sohar results change the IR cross-check, so re-run it from the saved IR items.
      const ir = await X.readIrUploads();
      if (ir.length) { step('Re-running the IR Portal cross-check…'); await runIr(ir); }
      await refreshRules();
      await stampUpdated(`${list.map((f) => f.name).join(', ')} added to the Sohar results`);
    }
    return html;
  }

  // ---------- step 4: benefits ----------
  const usd = (v) => { const n = parseFloat(String(v ?? '').replace(/[$,]/g, '')); return isNaN(n) ? '' : `$${n.toLocaleString('en-US', { maximumFractionDigits: 0 })}`; };

  /** Write the Benefits sheet and show OOP / deductible progress, with the patients at or near their OOP max listed. */
  async function renderBenefits(results, cfg, kinds) {
    // Fee estimates (On only) first, so the Benefits sheet gets its Est MD / Est NP columns.
    if (feesOn()) await refreshFees(results).catch((e) => { console.error(e); note('feeStatus', `Fee estimates not updated: ${e.message}`, 'warn'); });
    const view = C.benefitsView(results);
    const s = C.benefitsSummary(view);
    const box = $('benBox');
    const pctTxt = `${Math.round((cfg.approachingPct ?? 0.2) * 100)}%`;
    await X.writeBenefits(benefitCols(), benefitRows(view), C.BENEFIT_PCT_COLS);
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
        <td><b>${esc(C.lasfir(r['Last Name'], r['First Name']))}</b><br><span class="hint">${esc(C.idLabel(r))}${r.Service ? ` · ${esc(r.Service)}` : ''}</span></td>
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

  // ---------- fee estimates (v0.11, a test, Off by default) ----------
  // Config "Fee estimates" = On turns it on. Off: none of this runs, no fee sheet is read, the Benefits sheet keeps its
  // v0.10 columns and the full-screen window gets no fee data. All fee sheets are add-in owned and safe to delete.
  const feesOn = () => !!(F && state.cfg && state.cfg.fees && state.cfg.fees.on);
  const benefitCols = () => C.BENEFIT_SHEET_COLS.concat('Pills', ...(feesOn() ? ['Est MD', 'Est NP'] : []));
  function benefitRows(view) {
    return view.rows.map((r) => ({ ...r, Pills: pillsFor(C.keyOf(r)), ...(feesOn() ? estCols(r['Per Nbr']) : {}) }));
  }
  const estAmount = (e) => (e && e.patient != null ? `${e.patient.toFixed(2)}${F.partialNote(e) ? ` (${F.partialNote(e)})` : ''}` : '');
  function estCols(per) {
    const e = state.fees && state.fees.est ? state.fees.est.get(String(per)) : null;
    return { 'Est MD': estAmount(e && e.md), 'Est NP': estAmount(e && e.np) };
  }
  function showFeeUi() {
    const box = $('feeSection');
    if (box) box.hidden = !feesOn();
    if (!feesOn()) state.fees = null;
  }
  /** Appointments on the Import sheet ({rendering, loc, event}), for the provider home clinic. */
  function importAppts() {
    const det = state.detection || (state.rows && C.detectHeader(state.rows));
    if (!det || !state.rows) return [];
    const cm = det.colMap, out = [];
    const end = det.dataEnd ?? state.rows.length;
    for (let r = det.headerRow + 1; r < end; r++) {
      const row = state.rows[r] || [];
      const g = (f) => (cm[f] == null ? '' : String(row[cm[f]] ?? '').trim());
      if (g('perNbr')) out.push({ rendering: g('rendering'), loc: g('loc'), event: g('event') });
    }
    return out;
  }
  // The small sheets staff edit. Read again before every refresh, so an edit made on the sheet since the last read is
  // never overwritten from a stale copy when the refresh writes the sheet back.
  const FEE_EDITABLE = ['locations', 'homes', 'payerMap', 'payerLog', 'codes', 'defaults'];
  /** The fee sheets: the schedule once per pane session (and after an upload or "Re-read fee sheets"); editable sheets every time. */
  async function loadFeeData() {
    if (state.fees && state.fees.loaded) {
      for (const k of FEE_EDITABLE) state.fees[k] = await X.readFee(k);
      return state.fees;
    }
    const f = { loaded: true };
    for (const k of ['schedule', 'meta', 'locations', 'homes', 'payerMap', 'payerLog', 'codes', 'defaults']) f[k] = await X.readFee(k);
    if (!state.feeRowsChecked) {
      state.feeRowsChecked = true;
      await X.ensureConfigRows(F.FEE_CONFIG_ROWS).catch((e) => console.warn('Config fee estimate rows not added', e));
    }
    state.fees = f;
    return f;
  }
  const sameRows = (cols, a, b) => JSON.stringify((a || []).map((r) => cols.map((c) => String(r[c] ?? '')))) === JSON.stringify((b || []).map((r) => cols.map((c) => String(r[c] ?? ''))));
  /**
   * Match each verified patient's next visit to a fee tab and payer column, save the mapping sheets (a person's edits
   * on them are kept), and estimate. Nothing is guessed: what cannot be matched is flagged.
   */
  async function refreshFees(results) {
    if (!feesOn()) { state.fees = null; return; }
    const f = await loadFeeData();
    const rs = results || state.results || [];
    f.status = F.feeStatus(f.meta, f.schedule.length, new Date());
    f.warn = [];
    if (!f.schedule.length) { f.prep = null; f.est = new Map(); renderFeeStatus(); return; }
    const prep = F.prepareEstimates({
      // Info saved for a different upload (half-failed save) must not lend its Locations to these rows.
      results: rs, scheduleRows: f.schedule, tabSummary: F.metaMatches(f.meta, f.schedule.length) ? F.readMeta(f.meta).tabSummary : [], locations: f.locations, homes: f.homes,
      payerMap: f.payerMap, payerLog: f.payerLog, codes: f.codes, defaults: f.defaults, appts: importAppts(),
      teleKeywords: (state.cfg && state.cfg.teleKeywords) || ['Tele'], cfgDefaults: state.cfg.fees.defaults, when: fmtStamp(new Date()),
      pers: C.benefitsView(rs).rows.map((r) => r['Per Nbr']),
    });
    const save = async (kind, cols, rows, label) => {
      if (sameRows(cols, rows, f[kind])) return;
      try { await X.writeFee(kind, cols, rows); f[kind] = rows; } catch (e) { console.error(e); f.warn.push(`${label} not saved: ${e.message}`); }
    };
    await save('locations', F.FEE_LOCATION_COLS, prep.locationRows, 'Fee Locations');
    await save('homes', F.PROVIDER_HOME_COLS, prep.homeRows, 'Provider Home Clinic');
    await save('payerMap', F.FEE_PAYER_MAP_COLS, prep.payerRows, 'Fee Payer Map');
    if (prep.logRows.length) await save('payerLog', F.FEE_PAYER_LOG_COLS, f.payerLog.concat(prep.logRows), 'Fee Payer Map Log');
    if (prep.seeded) await save('codes', F.ESTIMATE_CODE_COLS, prep.codeRows, 'Estimate Codes');
    f.prep = prep;
    f.est = new Map(Object.values(prep.contexts).map((ctx) => [ctx.per, F.estimatePatient(ctx, prep.tables)]));
    renderFeeStatus();
  }
  function renderFeeStatus() {
    const box = $('feeStatus');
    if (!box) return;
    const f = state.fees;
    if (!feesOn() || !f) { box.innerHTML = ''; return; }
    const st = f.status || { text: 'No fee schedule uploaded yet.' };
    const parts = [`<p class="${st.loaded ? 'ok' : 'hint'}">${esc(st.text)}${st.loaded ? '.' : ' Upload the fee schedule workbook (.xlsx).'}</p>`];
    if (st.warn) parts.push(`<p class="warn">${esc(st.warn)}</p>`);
    for (const i of st.issues || []) parts.push(`<p class="hint">${esc(i)}</p>`);
    if (f.prep) {
      const groups = new Map();
      for (const e of f.est.values()) for (const fl of e.flags) {
        const k = /^No fee tab|matches more than one fee tab|is not in the uploaded/.test(fl) ? 'need a fee tab (Fee Locations)'
          : /^Map payer|^Fee payer/.test(fl) ? 'need a payer mapped to a fee column (Fee Payer Map)'
            : /^No default codes|^Enter codes/.test(fl) ? 'need estimate codes (Estimate Codes)'
              : /^Telehealth visit/.test(fl) ? 'telehealth with no home clinic (Provider Home Clinic)' : 'other';
        groups.set(k, (groups.get(k) || 0) + 1);
      }
      const n = f.prep.counts;
      parts.push(`<p>${n.ready} of ${n.patients} verified patient${n.patients === 1 ? '' : 's'} have an estimate for the next visit${n.flagged ? `; ${n.flagged} flagged` : ''}. Estimates are not a guarantee.</p>`);
      if (groups.size) parts.push(`<ul class="pc-items">${[...groups].map(([k, v]) => `<li>${v} ${esc(k)}</li>`).join('')}</ul>`);
    }
    for (const w of f.warn || []) parts.push(`<p class="warn">${esc(w)}</p>`);
    box.innerHTML = parts.join('');
    box.hidden = false;
  }
  /** Parse an uploaded fee schedule (pure-JS reader) and replace the saved one completely. Returns a summary line. */
  async function saveFeeSchedule(name, bytes) {
    if (!feesOn()) throw new Error('Fee estimates are Off. Set Config "Fee estimates" to On, then refresh (⟳).');
    const parsed = F.parseFeeSchedule(await XL.readWorkbook(bytes));
    const skipped = parsed.skipped.map((s) => `${s.name} (${s.reason})`).join(', ');
    if (!parsed.counts.tabs) throw new Error(`No fee tabs found in ${name}${skipped ? `: ${skipped}` : ''}. The saved fee schedule was not changed.`);
    const meta = F.feeMetaRows(name, new Date().toISOString(), parsed);
    await X.writeFee('schedule', F.FEE_COLS, parsed.rows);
    try {
      await X.writeFee('meta', F.FEE_META_COLS, meta);
    } catch (e) {
      // Two saves: the rates are already replaced. Say so plainly; the info's row count no longer matches the rows, so
      // the status line warns and its old tab summary is not used until the next upload.
      if (state.fees) state.fees.loaded = false;
      throw new Error(`${name}: the fee schedule rates were replaced, but its info (file name, upload date, tab summary) was not saved: ${e.message}. Upload the fee schedule again.`);
    }
    // The new schedule replaces the old one in memory too (no re-read of a large sheet); nothing of the old one is kept.
    if (state.fees && state.fees.loaded) { state.fees.schedule = parsed.rows; state.fees.meta = meta; }
    await refreshFees();
    await afterFeeChange();
    const issues = F.parseWarnings(parsed.tabs);
    return `${name}: ${parsed.counts.tabs} tab${parsed.counts.tabs === 1 ? '' : 's'}, ${parsed.counts.codes} codes saved${skipped ? `; skipped ${skipped}` : ''}.${issues.length ? ` ${issues.join('. ')}.` : ''}`;
  }
  /** After fee data changed: Benefits sheet columns, Clinic cards and the full-screen window follow. */
  async function afterFeeChange() {
    if (state.results) await X.writeBenefits(benefitCols(), benefitRows(C.benefitsView(state.results)), C.BENEFIT_PCT_COLS).catch((e) => { console.error(e); note('feeStatus', `Benefits sheet not updated: ${e.message}`, 'warn'); });
    renderClinicList();
    if (dlg) await sendToDialog({ type: 'fee-update', fees: feePayload() });
  }
  async function uploadFee(files) {
    const file = files[0];
    if (!file) return;
    if (!/\.xls[xm]$/i.test(file.name)) throw new Error(`${file.name} is not an .xlsx file. The fee schedule must be the Excel workbook.`);
    dropBusy('dropFee', `Reading ${file.name}…`);
    try { note('feeStatus', await saveFeeSchedule(file.name, new Uint8Array(await file.arrayBuffer())), 'ok'); } finally { dropBusy('dropFee'); }
  }
  /** What the full-screen window needs for estimates: contexts, the rate tables they use, status. Never the whole schedule. */
  function feePayload() {
    if (!feesOn() || !state.fees) return null;
    const f = state.fees;
    return { on: true, status: f.status || null, contexts: f.prep ? f.prep.contexts : {}, tables: f.prep ? f.prep.tables : { rates: {}, gross: {}, desc: {} }, toggles: F.TOGGLES };
  }
  /** "Remember these toggles for this fee payer" from the full-screen estimate panel. */
  async function saveFeeDefaults(feePayer, toggles) {
    if (!feesOn() || !feePayer) return;
    const f = await loadFeeData();
    const rows = F.rememberDefaults(f.defaults, feePayer, toggles || {}, fmtStamp(new Date()));
    await X.writeFee('defaults', F.ESTIMATE_DEFAULT_COLS, rows);
    f.defaults = rows;
    await refreshFees();
    await afterFeeChange();
  }
  function feeCardHtml(per) {
    if (!feesOn() || !state.fees || !state.fees.est) return '';
    const e = state.fees.est.get(String(per));
    if (!e) return '';
    const amt = (x, lvl) => (x && x.patient != null ? `${lvl} ${esc(F.usd(x.patient))}` : `${lvl} —`);
    if (!e.md && !e.np) return `<dt>Estimate</dt><dd class="hint">${esc(e.flags[0] || 'No estimate')}</dd>`;
    if ((e.md || e.np).noBenefits) return `<dt>Estimate</dt><dd class="hint">No deductible, coinsurance or copay from Sohar; allowed ${e.md && !e.md.none ? `MD ${esc(F.usd(e.md.allowed))}` : ''}${e.np && !e.np.none ? ` NP ${esc(F.usd(e.np.allowed))}` : ''}</dd>`;
    const partial = [...new Set([F.partialNote(e.md), F.partialNote(e.np)].filter(Boolean))];
    const missing = F.benefitsMissing(e.md) || F.benefitsMissing(e.np);
    return `<dt>Estimate</dt><dd>${amt(e.md, 'MD')} · ${amt(e.np, 'NP')}<br><span class="hint">Next visit, not a guarantee${partial.length ? `; ${esc(partial.join('; '))}` : ''}${missing ? '; some benefits missing' : ''}</span></dd>`;
  }

  async function currentDiscoveryVerify() {
    if (state.demo) return state.discoveryVerify || [];
    return X.readDiscoveryVerify(); // the sheet, so member IDs typed there go out
  }
  async function currentVerifyFailed() {
    if (state.demo) return state.verifyFailed || [];
    return X.readVerifyFailed(); // the sheet, so ZIPs typed there go out
  }
  async function downloadVerifyFailed() {
    const rows = await currentVerifyFailed();
    if (!rows.length) throw new Error('No failed verifications are waiting for discovery.');
    await saveCSV(C.fileName('verifyFailed'), C.DISCOVERY_COLS, rows, 'resultsResult');
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
    const out = C.trackerExport(rows, fmtStamp(new Date()), secondaryByPer());
    if (!out.rows.length) throw new Error('No verified patients yet. The final file needs at least one Sohar verification result.');
    await saveCSV(C.fileName('tracker'), C.TRACKER_EXPORT_COLS, out.rows, 'benBox');
    const n = `${out.rows.length} patient${out.rows.length === 1 ? '' : 's'}`;
    if (out.pending.length) {
      const who = out.pending.map((r) => `${C.lasfir(r['Last Name'], r['First Name'])} · ${C.idLabel(r)}`).join(', ');
      note('benBox', `${n} in the file. Left out, still waiting on the step 3 verification: ${who}. Save the tracker file again after that output is uploaded.`, 'warn');
    } else note('benBox', `${n} in the tracker file. Import it in the TMS Tracker and the ESK Tracker.`, 'ok');
  }

  let clearArmed = 0;
  async function clearSaved() {
    if (Date.now() - clearArmed > 5000) { clearArmed = Date.now(); $('btnClearSaved').textContent = 'Click again to clear'; setTimeout(() => { $('btnClearSaved').textContent = 'Clear saved Sohar results'; }, 5000); return; }
    clearArmed = 0;
    await X.writeUploads(C.UPLOAD_COLS, []);
    $('btnClearSaved').textContent = 'Clear saved Sohar results';
    note('resultsResult', 'Saved Sohar results cleared. The next upload starts fresh; the sheets stay as they are until then. The payer outcomes ledger (payer learning) and the Sohar ref memory are kept.', 'ok');
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
  /**
   * Tasks with Done and Clinic Note as the NextGen Tasks sheet has them now: a Done or note typed on the sheet after the
   * upload is not in state.tasks, and a bring-back or inbox save would otherwise put the old value back.
   */
  async function liveTasks() {
    const tasks = (await currentTasks()) || [];
    if (!state.tasks || state.demo) return tasks;
    const cell = (t, c) => String(t[c] ?? '').trim();
    const sheet = new Map(((await X.readTasks()) || []).map((t) => [String(t['Task Key']), t]));
    state.tasks = tasks.map((t) => { const s = sheet.get(String(t['Task Key'])); return s ? { ...t, Done: cell(s, 'Done'), 'Clinic Note': cell(s, 'Clinic Note') } : t; });
    return state.tasks;
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
    const base = {
      // x (patient extras) stays in the pane; the window gets the Pills column instead. oon (out-of-network benefits)
      // stays too: fee estimates send what they need in fees.contexts.
      type: 'data', focus: focus || '', web: isWeb(), results: ((await currentResults()) || []).map(({ x, resultKey, chooseGuess, oon, ...r }) => r), tasks: await currentTasks(), ir: await currentIr(),
      exports: state.exports || null, discoveryVerify: await currentDiscoveryVerify(), verifyFailed: await currentVerifyFailed(), scope: state.cfg ? state.cfg.scope : 'upcoming',
      asOf: state.resultsAsOf ? state.resultsAsOf.toISOString() : null,
      // v0.12: clinic lists (with the skip window) for the Clinic lists tab, and the Payer learning views.
      clinic: state.clinic ? { rows: state.clinic.rows, review: state.clinic.review, runDate: state.clinic.runDate } : null,
      skipped: skipRows(), skipCols: SKIP_COLS,
      learning: { rows: C.payerLearningRows(state.learning || []), metrics: state.metrics ? state.metrics.rows : [], totals: state.metrics ? state.metrics.totals : null,
        themes: state.themes ? state.themes.rows : [], themeCounts: state.themes ? state.themes.counts : {}, tests: state.payerTests || [],
        mapping: (state.exports && state.exports.mapping) || [], payerTestRows: (state.exports && state.exports.payerTests) || [] },
      guarantor: { count: (state.guarantorCandidates || []).length, noDob: (state.guarantorCandidates || []).filter((x) => !x.hasDob).length },
    };
    if (feesOn()) base.fees = feePayload(); // fee estimates (v0.11): only when On
    return base;
  }
  async function sendToDialog(obj) {
    if (!dlg) return;
    for (const piece of C.chunkMessage(obj)) dlg.messageChild(piece);
  }

  async function openFull(focus) {
    const url = new URL(`dashboard.html?v=${VERSION}`, location.href).href;
    if (state.demo) { showFloat(url + (feesOn() ? '&fees=1' : '') + (focus ? `#${focus}` : '')); return; } // browser preview: the same float, drawn in the page
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
    } else if (m.type === 'fee-upload' && feesOn()) {
      // Fee schedule picked on the full-screen Benefits tab: it arrives as base64 and is read here (pure JS, no Excel sheets).
      try {
        const text = await saveFeeSchedule(m.name, base64Bytes(m.base64));
        note('feeStatus', text, 'ok');
        await sendToDialog({ ...(await payload('benefits')), notice: { ok: true, text } });
      } catch (e) {
        await sendToDialog({ type: 'notice', ok: false, text: e.message });
        throw e;
      }
    } else if (m.type === 'clinic-file') {
      await saveClinicFile(m.location);
      await sendToDialog({ type: 'notice', ok: true, text: `Clinic file for ${m.location} saved to Downloads (Excel on the web: see the pane for "Save in a new tab").` });
    } else if (m.type === 'ctt-file') {
      // clinicPack throws a readable error (one record too big for a cell): show it in the window too.
      try { await saveCttFile(m.location); } catch (e) { await sendToDialog({ type: 'notice', ok: false, text: e.message }); throw e; }
      await sendToDialog({ type: 'notice', ok: true, text: `Clinic Template Tool file for ${m.location} saved to Downloads (Excel on the web: see the pane for "Save in a new tab"). It has PHI: share it through MHS OneDrive or Teams only.` });
    } else if (m.type === 'guarantor-resubmit') {
      await resubmitGuarantors();
      await sendToDialog({ ...(await payload('upload')), notice: { ok: true, text: 'Guarantors added to the discovery file. Save it again on this tab.' } });
    } else if (m.type === 'fee-defaults' && feesOn()) {
      await saveFeeDefaults(m.feePayer, m.toggles);
    }
  }
  function base64Bytes(b64) {
    const bin = atob(String(b64 || ''));
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  // Saves straight to Downloads. Excel on the web may block downloads inside the add-in frame, so there the note
  // also offers a save tab (save.html, same site, top level) and, as a last resort, the Export sheet.
  /** Binary files (clinic .xlsx): same download, same save-tab fallback on the web (save.html takes base64). */
  async function saveBinary(name, bytes, mime, boxId, what) {
    downloadBlob(name, new Blob([bytes], { type: mime }));
    const msg = `Saved ${name}${what ? ` (${what})` : ''} to your Downloads folder.`;
    if (!isWeb()) { note(boxId, msg, 'ok'); return msg; }
    const p = note(boxId, `${msg} Nothing there? `, 'ok');
    const tab = document.createElement('button'); tab.className = 'link'; tab.textContent = 'Save in a new tab';
    tab.addEventListener('click', () => saveInTab(name, { base64: XL.toBase64(bytes), mime }, boxId));
    p.append(tab, '.');
    return msg;
  }

  async function saveCSV(name, cols, rows, boxId, opts) {
    const csv = C.toCSV(cols, rows, opts);
    downloadNow(name, csv);
    const msg = `Saved ${name} (${rows.length} row${rows.length === 1 ? '' : 's'}) to your Downloads folder.`;
    if (!boxId) return msg;
    if (!isWeb()) { note(boxId, msg, 'ok'); return msg; }
    const p = note(boxId, `${msg} Nothing there? `, 'ok');
    const tab = document.createElement('button'); tab.className = 'link'; tab.textContent = 'Save in a new tab';
    tab.addEventListener('click', () => saveInTab(name, { csv }, boxId));
    const sheet = document.createElement('button'); sheet.className = 'link'; sheet.textContent = 'put it on the Export sheet';
    sheet.addEventListener('click', () => guard(async () => {
      await X.writeExport(cols, rows);
      note(boxId, `Rows are on the Export sheet. File > Export > Download as CSV, and name it ${name}.`, 'ok');
    }, boxId));
    p.append(tab, ' or ', sheet, '.');
    return msg;
  }
  function downloadNow(name, csv) { downloadBlob(name, new Blob([csv], { type: 'text/csv;charset=utf-8' })); }
  function downloadBlob(name, blob) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }
  /** Open save.html and hand it the file ({csv} or {base64, mime}) by postMessage until it answers. Must run inside the click (pop-up rules). */
  function saveInTab(name, file, boxId) {
    const w = window.open(new URL(`save.html?v=${VERSION}`, location.href).href, '_blank');
    if (!w) { note(boxId, 'The browser blocked the new tab. Allow pop-ups for Excel, or use the Export sheet.', 'bad'); return; }
    let got = false, tries = 0;
    const onMsg = (e) => { if (e.origin === location.origin && e.data && e.data.type === 'sohar-save-got') { got = true; window.removeEventListener('message', onMsg); } };
    window.addEventListener('message', onMsg);
    const send = () => {
      if (got || ++tries > 40 || w.closed) { window.removeEventListener('message', onMsg); return; }
      try { w.postMessage({ type: 'sohar-save', name, ...file }, location.origin); } catch (e) { /* not loaded yet */ }
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
    return { patName: 'Patient name', dob: 'DOB', perNbr: 'Per Nbr', mrn: 'MRN', apptDt: 'Appt date', begTm: 'Start time', event: 'Event', loc: 'Location', sts: 'Status', patSts: 'Patient status', payer: 'Payer', secPayer: 'Secondary payer', polNbr: 'Member ID', polNbrLstEnc: 'Member ID (last enc)', zip: 'ZIP', state: 'State', guarName: 'Guarantor', guarDob: 'Guarantor DOB', rendering: 'Rendering',
      resName: 'Resource', finClass: 'Fin class', finClassLstEnc: 'Fin class (last enc)', secPayerLstEnc: 'Secondary (last enc)', authCd: 'Auth code', authCdLstEnc: 'Auth code (last enc)' }[f] || f;
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
  // Every name, MRN, member ID and auth code below is made up. Dates are relative to today so the demo stays current.
  function demoAdapter() {
    const hdr = ['Appointment Pre-Check Complete', 'Res Name', 'Pat Sts', 'Portal Enrollment', 'Pat Name', 'Appt Date', 'Beg Tm', 'Event', 'Details', 'Loc Name',
      'Gender Identity', 'Current Gender', 'Preferred Pronoun', 'MRN', 'Status', 'Dur', 'Birth Dt', 'Rendering', 'End Tm', 'Per Nbr', 'Nxt Appt Dt', 'Nxt Beg Tm',
      'Nxt Event', 'Rschd', 'Rschd Desc', 'Guar Zip', 'Chkin Dt', 'Chkout Dt', 'Sec Payer (Lst Enc)', 'Fin Class (Lst Enc)', 'Pol Nbr (Lst Enc)', 'Auth Cd (Lst Enc)',
      'Payer Name (Cht Guar)', 'Sec Payer (Cht Guar)', 'Fin Class (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Auth Cd (Cht Guar)', 'Other ID Nbr', 'Crt Dt', 'Created By',
      'Mod Dt', 'Modified By', 'Measures Needed?', 'No Show/Late Cancel Charged?', 'Guar Name', 'Guar Birth Dt'];
    const col = (h) => hdr.indexOf(h);
    const now = new Date();
    const day = (off) => { const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() + off); return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`; };
    const iso = (off) => new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate() + off, 15, 0)).toISOString();
    const MD = 'Stone, Avery MD', DO = 'Reyes, Dana DO', PMHNP = 'Lark, Jordan PMHNP', NP = 'Quill, Morgan NP';
    // [name, per, dob, zip, payer, sec, fin, pol, auth, finL, polL, authL, appts: [dayOffset, time, event, loc, rendering, status?]]
    const P = [
      ['Rivera, Ana', '200101', '03/04/1980', '92660', 'Blue Shield CA', '', 'Blue Shield', 'XEA200101', '', 'Blue Shield', 'XEA200101', '',
        [[-21, '9:00 AM', 'TMS Treatment Unilateral', 'Newport Beach', MD, 'Kept'], [1, '9:00 AM', 'TMS Treatment Unilateral', 'Newport Beach', MD], [2, '9:00 AM', 'TMS Treatment Unilateral', 'Newport Beach', MD], [3, '9:00 AM', 'TMS Treatment Unilateral', 'Newport Beach', MD, 'Cancelled']]],
      ['Okafor, Chidi', '200102', '07/19/1975', '92663', 'VIP', '', 'VIP', '', '', 'Anthem', 'SYA200102', '',
        [[-10, '10:00 AM', 'Esketamine Follow Up', 'Newport Beach', NP, 'Kept'], [1, '10:00 AM', 'Esketamine Follow Up', 'Newport Beach', NP]]],
      ['Lindqvist, Sven', '200103', '11/02/1968', '94901', 'Partial VIP', '', 'VIP', '', '', 'VIP', '', '',
        [[3, '8:30 AM', 'IV Ketamine Follow Up', 'San Rafael', DO]]],
      ['Haddad, Mira', '200104', '01/30/1990', '94903', 'Kaiser Foundation Northern California', '', 'Kaiser', 'KP200104', '', 'Kaiser', 'KP200104', '',
        [[2, '1:00 PM', 'Med Management Telepsych', 'San Rafael', PMHNP]]],
      ['Nakamura, Kenji', '200105', '05/12/1983', '94904', 'Kaiser Foundation Northern California', '', 'Kaiser', 'KP200105', 'KPAUTH-0105', 'Kaiser', 'KP200105', 'KPAUTH-0105',
        [[4, '11:00 AM', 'Telepsych', 'San Rafael', DO], [5, '9:30 AM', 'TMS Treatment Unilateral', 'San Rafael', MD]]],
      ['Dubois, Claire', '200106', '09/09/1979', '92691', 'UMR', '', 'United Healthcare/Optum', 'UMR200106', '', 'United Healthcare/Optum', 'UMR200106', '',
        [[1, '11:00 AM', 'TMS Treatment Unilateral', 'Mission Viejo', PMHNP], [2, '11:00 AM', 'TMS Treatment Unilateral', 'Mission Viejo', PMHNP]]],
      ['Osei, Kwame', '200107', '12/24/1971', '92692', 'Carelon Behavioral Health', '', 'Beacon', 'CBH200107', '', 'Beacon', 'CBH200107', '',
        [[2, '2:00 PM', 'TMS Theta Tx Bilateral', 'Mission Viejo', MD], [6, '3:00 PM', 'Esketamine Follow Up', 'Mission Viejo', NP]]],
      ['Petrov, Ilya', '200108', '04/15/1986', '94941', 'Partnership HealthPlan Medi-Cal', '', 'Non-Contracted Commercial', 'PHP200108', '', 'Blue Shield', 'XEA200108', '',
        [[-20, '10:00 AM', 'Med Management In-Office', 'San Rafael', DO, 'Kept'], [3, '10:00 AM', 'Med Management In-Office', 'San Rafael', DO]]],
      ['Alvarez, Sofia', '200109', '06/06/1994', '94114', 'San Francisco Health Plan', '', 'Contracted Commercial', 'SFHP200109', '', 'Contracted Commercial', 'SFHP200109', '',
        [[2, '9:00 AM', 'Telepsych', 'Castro Street', PMHNP]]],
      ['Byrne, Niamh', '200110', '02/28/1952', '94110', 'QMB Noridian Northern CA', 'Blue Shield CA', 'Medicare', 'SYN1EG4TE5', '', 'Medicare', 'SYN1EG4TE5', '',
        [[1, '2:00 PM', 'TMS Treatment Unilateral', 'Castro Street', MD]]],
      ['Kowalski, Piotr', '200111', '08/08/1958', '92657', 'Imperial Health CA Medicare Adv', '', 'Medicare Advantage', 'IMP200111', '', 'Aetna', 'W200111A', '',
        [[-14, '1:00 PM', 'Telepsych', 'Newport Beach', DO, 'Kept'], [5, '1:00 PM', 'Telepsych TMS F/U', 'Newport Beach', DO]]],
      ['Mensah, Afia', '200112', '10/10/1988', '92694', 'Aetna', '', 'Aetna', 'W200112', '', 'Aetna', 'W200112', '',
        [[3, '4:00 PM', 'Med Management Telepsych', 'Mission Viejo', DO]]],
      ['Moreau, Lucie', '200113', '03/21/1977', '92625', 'Cigna', '', 'Cigna', 'U200113', '', 'Cigna', 'U200113', '',
        [[2, '8:00 AM', 'TMS Follow Up', 'Newport Beach', MD]]],
      ['Tanaka, Yui', '200114', '12/01/1991', '94939', 'Anthem Blue Cross', '', 'Anthem', 'SYA200114', '', 'Anthem', 'SYA200114', '',
        [[4, '1:30 PM', 'Esketamine Tx Start', 'San Rafael', NP]]],
      ['Ibarra, Luis', '200115', '05/05/1999', '94117', 'Self Pay', '', 'Self Pay', '', '', 'Self Pay', '', '',
        [[6, '10:30 AM', 'Tel New Patient - Mature', 'Castro Street', PMHNP]]],
      ['Fischer, Hanna', '200116', '07/07/1984', '92692', 'Blue Shield CA Blue Card', '', 'Blue Shield', 'XEB200116', '', 'Blue Shield', 'XEB200116', '',
        [[7, '9:00 AM', 'TMS Treatment Unilateral', 'Mission Viejo', MD]]],
      ['Grant, Theo', '200117', '01/17/1949', '94949', 'Kaiser Medicare Northern CA', '', 'Kaiser', 'KPM200117', '', 'Kaiser', 'KPM200117', '',
        [[8, '11:30 AM', 'Med Management In-Office', 'San Rafael', DO]]],
      // Esketamine payer credential rules (R13) against the made-up demo rules below: an NP booked for a first treatment
      // the payer only lets an MD bill (and not Besse-approved), and a payer whose Besse approval is pending.
      ['Castillo, Rosa', '200118', '04/04/1987', '92660', 'Zephyr Health PPO', '', 'Contracted Commercial', 'ZPH200118', '', 'Contracted Commercial', 'ZPH200118', '',
        [[3, '9:00 AM', 'Esketamine First Treatment', 'Newport Beach', NP], [5, '9:00 AM', 'Esketamine Follow Up', 'Newport Beach', NP]]],
      ['Brandt, Felix', '200119', '09/19/1973', '92691', 'Lumen Care HMO', '', 'Contracted Commercial', 'LUM200119', '', 'Contracted Commercial', 'LUM200119', '',
        [[5, '2:00 PM', 'Esketamine Follow Up', 'Mission Viejo', PMHNP]]],
    ];
    // Made-up guarantors: Lindqvist's search found nothing (guarantor DOB on file); Ibarra's guarantor has no DOB in the report.
    const DEMO_GUAR = { 200103: ['Lindqvist, Freya', '02/14/1970'], 200115: ['Ibarra, Rosa', ''], 200116: ['Fischer, Jonas', '05/05/1960'] };
    let rows = [[], ['', '', '', '', 'Mindful Health Solutions'], [], ['', '', '', '', 'Appointments Listing'], ['', '', '', '', `From ${day(-30)} to ${day(30)}`], [], hdr];
    let n = 0;
    for (const [name, per, dob, zip, payer, sec, fin, pol, auth, finL, polL, authL, appts] of P) {
      for (const [off, tm, ev, loc, rend, sts] of appts) {
        const r = new Array(hdr.length).fill('');
        r[col('Res Name')] = rend; r[col('Pat Sts')] = 'Active'; r[col('Pat Name')] = name; r[col('Appt Date')] = day(off); r[col('Beg Tm')] = tm;
        r[col('Event')] = ev; r[col('Loc Name')] = loc; r[col('MRN')] = `8${per.slice(1)}`; r[col('Status')] = sts || (off < 0 ? 'Kept' : 'Expected');
        r[col('Birth Dt')] = dob; r[col('Rendering')] = rend; r[col('Per Nbr')] = per; r[col('Guar Zip')] = zip;
        r[col('Fin Class (Lst Enc)')] = finL; r[col('Pol Nbr (Lst Enc)')] = polL; r[col('Auth Cd (Lst Enc)')] = authL;
        r[col('Payer Name (Cht Guar)')] = payer; r[col('Sec Payer (Cht Guar)')] = sec; r[col('Fin Class (Cht Guar)')] = fin;
        r[col('Pol Nbr (Cht Guar)')] = pol; r[col('Auth Cd (Cht Guar)')] = auth; r[col('Created By')] = 'synthetic';
        const g = DEMO_GUAR[per];
        if (g) { r[col('Guar Name')] = g[0]; r[col('Guar Birth Dt')] = g[1]; }
        rows.push(r); n++;
      }
    }
    // A staff meeting row (no patient), as NextGen lists them.
    { const r = new Array(hdr.length).fill(''); r[col('Res Name')] = MD; r[col('Appt Date')] = day(1); r[col('Beg Tm')] = '12:00 PM'; r[col('Event')] = 'Meeting'; r[col('Loc Name')] = 'Newport Beach'; rows.push(r); n++; }
    rows.push([`Total Appointments (${n})`]);
    // Leftovers from pasting over an older, longer report: staff names sit where the payer column is. They must be ignored (R1).
    for (let i = 0; i < 4; i++) {
      const r = new Array(hdr.length).fill('');
      r[col('Per Nbr')] = String(299000 + i); r[col('Appt Date')] = day(i + 2); r[col('Payer Name (Cht Guar)')] = [MD, DO, PMHNP, NP][i]; r[col('Pat Name')] = 'Old, Layout';
      rows.push(r);
    }
    let payerMap = [
      { nextgenPayerName: 'Kaiser Foundation Northern California', payerId: '94135', payerName: 'Kaiser Foundation Health Plan of Northern California', confidence: 'Exact' },
      { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
      { nextgenPayerName: 'Aetna', payerId: '60054', payerName: 'Aetna', confidence: 'Exact' },
      { nextgenPayerName: 'Self Pay', payerId: '', payerName: '', confidence: 'Not a payer' },
    ];
    const sohar = [['94135', 'Kaiser Foundation Health Plan of Northern California', ''], ['94036', 'Blue Shield of California', 'Blue Shield CA'], ['62308', 'Cigna', ''],
      ['47198', 'Blue Cross of California', 'Anthem Blue Cross; Anthem Blue Cross California'], ['87726', 'UnitedHealthcare', 'UHC'],
      ['01111', 'Medicare of California', 'Noridian Northern CA; Noridian Southern CA'], ['60054', 'Aetna', ''], ['37237', 'UMR', ''],
      ['CBH01', 'Carelon Behavioral Health', 'Beacon Health Options'], ['95567', 'Health Net', ''], ['PHPCA', 'Partnership HealthPlan of California', ''],
      ['SFHP1', 'San Francisco Health Plan', 'SFHP'], ['68039', 'Western Health Advantage', ''], ['KPSA1', 'Kaiser Permanente Senior Advantage', '']]
      .map(([payerId, payerName, aliasPayerNames]) => ({ payerId, payerName, aliasPayerNames }));
    // Synthetic Sohar results so the benefits and clinic views have something to show. Add ?fresh to the URL to start empty.
    const ver = (id, payer, member, status, ded, dedRem, oop, oopRem, copay, coins, off, plan) => ({ requestid: `ver-${id}-${member}`, id: String(id), status, payername: payer, memberid: member,
      planname: plan || 'SYNTH PPO', networktype: 'PPO', inncopay: copay, inncoinsurance: coins, inntotaldeductible: ded, innremainingdeductible: dedRem, innoutofpocket: oop, innoutofpocketremaining: oopRem, created: iso(off) });
    const demoUploads = /fresh/.test(location.search) ? [] : [
      ver(200101, 'Blue Shield of California', 'XEA200101', 'complete.eligible', '1500', '0', '5000', '400', '40', '0.2', -2),
      ver(200113, 'Cigna', 'U200113', 'complete.eligible', '2000', '300', '6000', '5100', '30', '0.1', -20),
      ver(200110, 'Medicare of California', 'SYN1EG4TE5', 'complete.eligible', '257', '0', '', '', '0', '0.2', -1, 'MEDICARE PART B'),
      ver(200112, 'Aetna', 'W200112', 'complete.eligible', '0', '0', '3000', '0', '0', '0', -3),
      ver(200111, 'Medicare of California', 'IMP200111', 'complete.ineligible', '', '', '', '', '', '', -1, ''),
      ver(200111, 'Aetna', 'W200111A', 'complete.eligible', '1000', '1000', '4000', '3800', '50', '0.2', -1),
      ver(200102, 'Blue Cross of California', 'SYA200102', 'complete.eligible', '750', '750', '3500', '3100', '35', '0.2', -1),
      ver(200106, 'UMR', 'UMR200106', 'complete.eligible', '500', '0', '2500', '200', '25', '0.1', -2),
      ver(200108, 'Partnership HealthPlan of California', 'PHP200108', 'complete.eligible', '0', '0', '0', '0', '0', '0', -2, 'MEDI-CAL MANAGED CARE'),
      ver(200108, 'Blue Shield of California', 'XEA200108', 'complete.ineligible', '', '', '', '', '', '', -2, ''),
      ver(200107, 'Carelon Behavioral Health', 'CBH200107', 'error.member.id', '', '', '', '', '', '', -1, ''),
      { requestid: 'disc-200103', id: '200103', resolved: iso(-2), created: iso(-2) },
      { requestid: 'disc-200115', id: '200115', resolved: iso(-2), created: iso(-2) },
      { requestid: 'disc-200116-g', id: '800116-G', responsepayerid: '94036', responsepayername: 'Blue Shield of California', memberid: 'XEB200116', resolved: iso(-1), created: iso(-1) },
    ].map((rec) => { const kind = rec.responsepayerid || /^disc/.test(rec.requestid) ? 'discovery' : 'verification'; return { Kind: kind, Key: C.resultKey(kind, rec), File: 'demo', Uploaded: iso(-1), Record: JSON.stringify(rec) }; });
    // Synthetic payer-learning history (made-up Per Nbrs, earlier weeks): discovery keeps agreeing with the Partnership guess
    // (promoted to Learned), keeps finding Senior Advantage for "Kaiser Medicare Northern CA" (a suggestion), and two
    // Blue Card verifications came back eligible (evidence only).
    const past = (per, kind, payer, sent, ret, retName, signal, off) => ({ Key: `demo|${per}|${payer}`, 'Per Nbr': per, Kind: kind, 'NextGen Payer': payer, 'Fin Class': '',
      Source: kind === 'Discovery' ? 'discovery' : 'chart', 'Sent payerId': sent, 'Sent Confidence': '', 'Sohar Status': kind === 'Discovery' ? 'Coverage found' : 'complete.eligible',
      'Returned payerId': ret, 'Returned Payer': retName, Signal: signal, 'Sohar Date': iso(off), Recorded: 'demo' });
    let outcomes = /fresh/.test(location.search) ? [] : [
      past('290001', 'Discovery', 'Partnership HealthPlan Medi-Cal', 'PHPCA', 'PHPCA', 'Partnership HealthPlan of California', 'agree', -40),
      past('290002', 'Discovery', 'Partnership HealthPlan Medi-Cal', 'PHPCA', 'PHPCA', 'Partnership HealthPlan of California', 'agree', -33),
      past('290003', 'Discovery', 'Partnership HealthPlan Medi-Cal', 'PHPCA', 'PHPCA', 'Partnership HealthPlan of California', 'agree', -26),
      past('290004', 'Discovery', 'Kaiser Medicare Northern CA', '94135', 'KPSA1', 'Kaiser Permanente Senior Advantage', 'other payer', -40),
      past('290005', 'Discovery', 'Kaiser Medicare Northern CA', '94135', 'KPSA1', 'Kaiser Permanente Senior Advantage', 'other payer', -33),
      past('290006', 'Discovery', 'Kaiser Medicare Northern CA', '94135', 'KPSA1', 'Kaiser Permanente Senior Advantage', 'other payer', -26),
      past('290007', 'Verification', 'Blue Shield CA Blue Card', '94036', '', 'Blue Shield of California', 'agree', -30),
      past('290008', 'Verification', 'Blue Shield CA Blue Card', '94036', '', 'Blue Shield of California', 'agree', -23),
    ];
    let mapLog = [], soharRefs = [], reverifyRequests = [];
    // v0.12 tables, in memory. Fischer's guarantor search was asked for earlier and found Blue Shield (Via guarantor).
    let guarantorRequests = [{ MRN: '800116', Requested: iso(-3), Reason: 'Discovery found nothing', By: 'demo' }], payerTests = [], metricHistory = [], clinicRequests = [];
    // A made-up payer note, so the Payer learning list shows in the demo.
    let clinicNotes = [], clinicApplied = [], clinicPayerNotes = [{ 'NextGen Payer': 'Zephyr Health', Location: 'San Rafael', 'Card Says': 'Zephyr Health PPO, group 0000', Note: 'Card shows a new plan name', By: 'demo', When: iso(-1) }];
    // The clinic inbox (the flow fills it in a real workbook): two made-up rows, made on first read against the demo's own
    // tasks so they apply: a task done with a note, and a patient note. ?noinbox starts without it (shows the set-up button).
    let inbox = /[?&]noinbox\b/.test(location.search) ? null : undefined;
    const inboxFirstRead = () => {
      const t = (state.tasks || []).find((x) => !x.Done && x['Next Location']);
      if (!t) return [];
      const at = (h) => new Date(Date.now() - h * 36e5).toISOString(), file = (l) => `Clinic Eligibility - ${l} ${C.localDate(new Date())}.xlsx`;
      return [
        { Location: t['Next Location'], Kind: 'Task done', Key: t['Task Key'], MRN: t.MRN, LASFIR: C.lasfir(t['Last Name'], t['First Name']), Value: 'Yes',
          Note: 'Called the patient, new card scanned', By: 'Front desk', When: at(3), File: file(t['Next Location']), Received: at(2.9) },
        { Location: 'San Rafael', Kind: 'Note', Key: '800104', MRN: '800104', LASFIR: C.lasfir('Haddad', 'Mira'), Value: '',
          Note: 'Bringing the new Kaiser card on Monday', By: 'Front desk', When: at(2), File: file('San Rafael'), Received: at(1.9) },
      ];
    };
    // Fee estimates demo: ?fees=1 turns them on with a made-up fee schedule (fees.js DEMO_WORKBOOK), in memory.
    const feesDemo = /[?&]fees=1\b/.test(location.search) && !!F;
    const feeStore = {};
    if (feesDemo) {
      const s = F.demoSchedule();
      feeStore.schedule = s.rows;
      feeStore.meta = F.feeMetaRows('Demo fee schedule (made-up rates).xlsx', new Date().toISOString(), s);
    }
    // Payer Credential Rules: made-up payers only (the real list is imported in each workbook, never committed).
    let credRules = PRM ? R.payerRuleRows(PRM.read([PRM.HEADERS,
      ['Zephyr Health', '', 'NP', 'No', 'No', 'No', 'No', 'Yes', 'zephyr health', 'NP may do observation only'],
      ['Zephyr Health', '', 'MD', 'No', 'N/A', 'Yes', 'Yes', 'Yes', 'zephyr health', ''],
      ['Lumen Care', '', 'NP', '', 'Yes', 'Yes', 'Yes', 'Yes', 'lumen care', ''],
      ['Lumen Care', '', 'MD', '', 'N/A', 'Yes', 'Yes', 'Yes', 'lumen care', ''],
    ]).rules) : [];
    return {
      async readPayerCredRules() { return credRules; }, async writePayerCredRules(c, r) { credRules = r; return r.length; },
      async readFee(kind) { return feeStore[kind] || []; }, async writeFee(kind, c, r) { feeStore[kind] = r; return r.length; },
      TABLES: { patients: 'tblPatients', review: 'tblReview', verification: 'tblVerification', discovery: 'tblDiscovery' },
      async readImportRows() { return rows; },
      async writeImportRows(r) { rows = r; },
      async importXlsx() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async readXlsxRows() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async writeResults() {}, async writeTasks() {}, async writeBenefits() {}, async readTasks() { return []; }, async readUploads() { return demoUploads; }, async readIrUploads() { return []; }, async writeIrUploads() {}, async writeUploads() {}, async readDiscoveryVerify() { return []; }, async writeDiscoveryVerify() {}, async readVerifyFailed() { return []; }, async writeVerifyFailed() {}, async tableRowCount() { return 0; }, async setTaskDone() {}, async writeIr() {}, async readIr() { return []; }, async writeExport(c, r) { return r.length; },
      async readPayerEvidence() { return []; }, async writePayerEvidence() {}, async tableExists() { return false; }, async readClinicFile() { return null; },
      // Payer learning: the ledger and the log live in memory; the ledger starts with a few synthetic outcomes (below).
      async readPayerOutcomes() { return outcomes; }, async writePayerOutcomes(c, r) { outcomes = r; },
      async readPayerMapLog() { return mapLog; }, async writePayerMapLog(c, r) { mapLog = r; },
      async appendPayerMapLog(c, r) { mapLog = mapLog.concat(r || []); return (r || []).length; },
      async writePayerLearning() {}, async ensureConfigRows() { return []; },
      // Sohar ref memory and Re-verify anyway requests, in memory.
      async readSoharRefs() { return soharRefs; }, async writeSoharRefs(c, r) { soharRefs = r; },
      async readReverifyRequests() { return reverifyRequests; }, async writeReverifyRequests(c, r) { reverifyRequests = r; },
      async readGuarantorRequests() { return guarantorRequests; }, async writeGuarantorRequests(c, r) { guarantorRequests = r; },
      async readPayerTests() { return payerTests; }, async writePayerTests(c, r) { payerTests = r; },
      async writePayerMapping() {}, async readPayerMapping() { return []; }, async writePayerMetrics() {}, async writeGuarantorThemes() {},
      async readPayerMetricHistory() { return metricHistory; }, async writePayerMetricHistory(c, r) { metricHistory = r; },
      async readClinicRequests() { return clinicRequests; }, async writeClinicRequests(c, r) { clinicRequests = r; },
      async readClinicNotes() { return clinicNotes; }, async writeClinicNotes(c, r) { clinicNotes = r; },
      async readClinicPayerNotes() { return clinicPayerNotes; }, async writeClinicPayerNotes(c, r) { clinicPayerNotes = r; },
      async readClinicWorkApplied() { return clinicApplied; }, async writeClinicWorkApplied(c, r) { clinicApplied = r; },
      INBOX: { sheet: 'Clinic Work Inbox', table: 'tblClinicWorkInbox' },
      async readClinicInbox() { if (inbox === undefined) { const r = inboxFirstRead(); if (r.length) inbox = r; return r; } return inbox; },
      async createClinicInbox() { if (inbox) return false; inbox = []; return true; },
      async readConfig() { return { asOf: null, scope: 'upcoming', locationState: { 'San Rafael': 'CA', 'Newport Beach': 'CA', 'Mission Viejo': 'CA', 'Castro Street': 'CA' }, defaultState: 'CA', teleKeywords: ['Tele'], posTele: '10', posOffice: '11', specialty: 'PSY', useGuarantorAsSubscriber: false, approachingPct: 0.2, sfhpProviders: ['Stone', 'Reyes'], reverifyDays: 7, reverify: C.reverifyConfig({}), clinicLocation: '', payerTests: C.payerTestConfig({ 'Payer tests': /[?&]tests=1\b/.test(location.search) ? 'On' : 'Off' }), mrnSince: '', learning: C.learningConfig({}), fees: F ? F.feeConfig({ 'Fee estimates': feesDemo ? 'On' : 'Off' }) : { on: false } }; },
      async readPayerMap() { return payerMap; },
      async readSoharPayers() { return sohar; },
      async upsertPayerMap(e, o) { return this.upsertPayerMapMany([e], o); },
      async upsertPayerMapMany(list) {
        for (const e of list) {
          const prev = payerMap.find((m) => m.nextgenPayerName.toLowerCase() === e.nextgenPayerName.toLowerCase());
          payerMap = payerMap.filter((m) => m !== prev).concat([{ ...e, notes: C.mergeNotes(e.notes, prev && prev.notes) }]);
        }
        return list.length;
      },
      async replaceTable() {}, async setConfigValue() {}, async activate() {}, async templateCheck() { return []; }, async recoverAll() { return []; }, takeRecovered() { return []; },
    };
  }
})();
