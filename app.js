/* app.js — wires the pane to core.js (logic) and excel.js (workbook). */
(function () {
  'use strict';
  const VERSION = '0.8.3';
  const ADDIN_ID = '9786ae8a-f258-4f2c-9627-4c6f0f25db7e';
  const C = window.SoharCore;
  const R = window.SoharRules;
  const XL = window.SoharXlsx;
  let X = null; // SoharExcel when running inside Excel; demo adapter otherwise

  const state = { rows: null, detection: null, cfg: null, payerMap: null, soharPayers: null, patients: null, review: null, exports: null, results: null, resultsAsOf: null, tasks: null, ir: null, discoveryVerify: null, pendingFocus: '', demo: false,
    flags: new Map(), clinic: null, evidence: null, clinicFile: false, clinicFilter: '' };
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
      if (!det || det.missing.length || det.dataEnd <= det.headerRow + 1) return;
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

  const AUTO_ORDER = { 'auto-low': 0, '': 0, 'auto-medium': 1, 'auto-high': 2 };
  /**
   * Optional review list: every NextGen name in this report that the add-in mapped itself (Auto-Low first).
   * Nothing here blocks the build. Confirming or changing a row writes a manual confidence, which always wins.
   */
  function renderPayers() {
    const idx = C.payerMapIndex(state.payerMap || []);
    const counts = payersInImport();
    const items = [...counts].map(([name, pers]) => {
      const h = idx.get(name.toLowerCase());
      const conf = h ? String(h.confidence || '').toLowerCase() : '';
      return { name, n: pers.size, h, conf };
    }).filter((it) => (!it.h && !C.isVip(it.name) && !/self\s*-?\s*pay/i.test(it.name)) || it.conf.startsWith('auto-'))
      .sort((a, b) => (AUTO_ORDER[a.conf] ?? 0) - (AUTO_ORDER[b.conf] ?? 0) || b.n - a.n || a.name.localeCompare(b.name));
    const box = $('payerList');
    box.hidden = !items.length;
    if (!items.length) { box.innerHTML = ''; return; }
    // Auto-Low and unmapped names need a person to classify them; Auto-High / Auto-Medium stay in an optional list.
    const isLow = (it) => it.conf === 'auto-low' || !it.h;
    const low = items.filter(isLow).length, rest = items.length - low;
    const keepOpen = box.querySelector('details') && box.querySelector('details').open;
    box.innerHTML = (low ? `<div class="needclass"><p class="warn"><b>${low} insurance name${low === 1 ? '' : 's'} need${low === 1 ? 's' : ''} manual classification.</b> The add-in could not match ${low === 1 ? 'it' : 'them'} confidently, so those patients go to discovery (or to both insurances when the last visit was billed differently). Pick the right Sohar payer once and it is remembered.</p><div class="needrows"></div></div>` : '')
      + (rest ? `<details class="more payerreview"${keepOpen ? ' open' : ''}><summary>Review auto-mapped payers (${rest})</summary>
      <p class="hint">Matched confidently (Auto-High / Auto-Medium) and sent to verification. Optional: fix a wrong match and it is remembered.</p>
      <div class="payerrows"></div></details>` : '');
    for (const it of items) {
      const list = box.querySelector(isLow(it) ? '.needrows' : '.payerrows');
      const h = it.h || {};
      const label = it.h ? h.confidence : 'Not mapped';
      const cls = it.conf === 'auto-high' ? 'ok' : it.conf === 'auto-medium' ? 'mid' : 'low';
      const wrap = document.createElement('div');
      wrap.className = 'payer';
      wrap.innerHTML = `<div class="name">${esc(it.name)} <span class="n">(${it.n} patient${it.n > 1 ? 's' : ''})</span> <span class="conf ${cls}">${esc(label)}</span></div>
        <div class="guess">${h.payerId ? `→ ${esc(h.payerName || '')} <span class="id">${esc(h.payerId)}</span>` : '<span class="hint">No Sohar payer guessed</span>'}</div>
        ${h.notes ? `<div class="hint ev">${esc(String(h.notes).replace(/^Auto-matched by the add-in:\s*/, ''))}</div>` : ''}
        <div class="alt-actions">${h.payerId ? '<button class="link" data-act="ok">Looks right</button>' : ''}<button class="link" data-act="change">Change</button><button class="link" data-act="notpayer">Not a payer</button></div>
        <div class="pick" hidden><input type="search" placeholder="Search Sohar payers" aria-label="Search Sohar payers for ${esc(it.name)}"><div class="sugg"></div></div>`;
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
            await X.upsertPayerMap({ nextgenPayerName: it.name, payerId: s.payerId, payerName: s.payerName, confidence: 'Confirmed', notes: 'Mapped in the add-in' });
            await afterMap();
          }, 'payerList'));
          sugg.appendChild(b);
        }
      };
      input.addEventListener('input', () => show(input.value));
      const ok = wrap.querySelector('[data-act="ok"]');
      if (ok) ok.addEventListener('click', () => guard(async () => {
        await X.upsertPayerMap({ nextgenPayerName: it.name, payerId: h.payerId, payerName: h.payerName, confidence: 'Confirmed', notes: `Confirmed in the add-in (was ${h.confidence})` });
        await afterMap();
      }, 'payerList'));
      wrap.querySelector('[data-act="change"]').addEventListener('click', () => { pick.hidden = !pick.hidden; if (!pick.hidden) { show(''); input.focus(); } });
      wrap.querySelector('[data-act="notpayer"]').addEventListener('click', () => guard(async () => {
        await X.upsertPayerMap({ nextgenPayerName: it.name, payerId: '', payerName: '', confidence: 'Not a payer', notes: 'Marked in the add-in' });
        await afterMap();
      }, 'payerList'));
      list.appendChild(wrap);
    }
  }

  // ---------- step 3 ----------
  async function build(opts) {
    busy('btnBuild', 'Building…');
    try {
      state.cfg = await X.readConfig();
      state.cfg.scope = $('scope').value;
      state.cfg.asOf = $('asOf').value ? new Date($('asOf').value + 'T00:00:00Z') : null;
      state.payerMap = await X.readPayerMap();
      if (!state.soharPayers) state.soharPayers = await X.readSoharPayers();
      let res = C.buildPatients(state.rows, { ...state.cfg, payerMap: state.payerMap, soharPayers: state.soharPayers });
      const out = $('buildResult'); out.hidden = false;
      if (res.error) { out.innerHTML = `<p class="bad">${esc(res.error)}</p>`; return; }
      // R2: new NextGen payer names are matched to Sohar payers automatically and saved on the Payer Map once
      // (editable there; a manual confidence always wins). Build again so the Fin Class index sees them too.
      let autoNote = '';
      if (res.autoMapped && res.autoMapped.length) {
        const add = res.autoMapped.map(({ nextgenPayerName, payerId, payerName, confidence, notes }) => ({ nextgenPayerName, payerId, payerName, confidence, notes }));
        await X.upsertPayerMapMany(add);
        state.payerMap = await X.readPayerMap();
        res = C.buildPatients(state.rows, { ...state.cfg, payerMap: state.payerMap, soharPayers: state.soharPayers });
        if (res.error) { out.innerHTML = `<p class="bad">${esc(res.error)}</p>`; return; }
        const lowN = add.filter((a) => a.confidence === 'Auto-Low').length;
        autoNote = `<p class="hint">${add.length} new insurance name${add.length === 1 ? ' was' : 's were'} matched to Sohar payers and added to the Payer Map${lowN ? ` (${lowN} low confidence)` : ''}. ${lowN ? ' The low-confidence ones are listed above for classification.' : ''}</p>`;
      }
      renderPayers();
      state.patients = res.patients; state.review = res.review;
      state.exports = C.buildExports(res.patients, state.cfg);
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
        <p class="issues" id="buildIssues"></p>${autoNote}`;
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
      $('fnDisc').textContent = `${nd} patient${nd === 1 ? '' : 's'} · ${C.fileName('discovery')}`;
      if (!(opts && opts.deferRules)) await refreshRules();
    } finally { unbusy('btnBuild', 'Rebuild with these options'); }
  }

  // ---------- clinic rules (R5-R10): flags, Review sheet, payer evidence, clinic view ----------
  const inScopePatients = () => (state.patients || []).filter((p) => p['In Scope'] === 'Yes');
  function rulesCfg() {
    const c = state.cfg || {};
    return { today: new Date(), sfhpProviders: c.sfhpProviders || [], reverifyDays: c.reverifyDays != null ? c.reverifyDays : 7 };
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
    const review = (state.review || []).concat(flagRows);
    await X.replaceTable(X.TABLES.review, C.REVIEW_COLS, review);
    const sev = { High: 0, Medium: 0, Low: 0 };
    for (const r of review) sev[r.Severity] = (sev[r.Severity] || 0) + 1;
    const bi = $('buildIssues');
    if (bi) bi.innerHTML = sev.High ? `<span class="hi">${sev.High} high-priority item${sev.High === 1 ? '' : 's'}</span> on the Review sheet${flagRows.length ? ` (${flagRows.length} clinic flag${flagRows.length === 1 ? '' : 's'} included)` : ''}.` : '';
    try {
      const prev = await X.readPayerEvidence();
      state.evidence = R.payerEvidence(pts, state.results || [], prev, new Date());
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
    const rv = ((state.clinic && state.clinic.review) || []).filter((x) => String(x['Per Nbr']) === per && x.Severity !== 'Low');
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
    $('clinicChips').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-pill]'); if (!b) return;
      state.clinicFilter = state.clinicFilter === b.dataset.pill ? '' : b.dataset.pill;
      renderClinicChips(); renderClinicList();
    });
    $('btnClinicEmail').addEventListener('click', () => guard(async () => {
      const loc = $('clinicLoc').value;
      const text = R.clinicEmail(loc, clinicRowsAt(loc), { runDate: runDateOf() });
      await copyOrShow(text, 'clinicNote', `Email summary for ${loc || 'all locations'} copied (LASFIR + Per Nbr only). Paste it into Outlook or Teams.`);
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
      && (!q || [r['Last Name'], r['First Name'], r['Per Nbr'], r.MRN].some((v) => String(v || '').toLowerCase().includes(q))));
    const list = $('clinicList');
    const meta = $('clinicMeta');
    if (!state.clinic) { meta.textContent = ''; return; }
    meta.textContent = `${rows.length} of ${all.length} patient${all.length === 1 ? '' : 's'} with an upcoming visit${loc ? ` at ${loc}` : ''}${state.clinic.runDate ? `. Run ${state.clinic.runDate}` : ''}.`;
    if (!rows.length) { list.innerHTML = `<p class="hint">${all.length ? 'No one matches this filter.' : 'No upcoming patients here.'}</p>`; return; }
    const days = new Map();
    for (const r of rows) { const d = r['Next Appt'] || 'No date'; if (!days.has(d)) days.set(d, []); days.get(d).push(r); }
    const wd = (s) => { const d = C.parseDate(s); return d ? d.toLocaleDateString('en-US', { weekday: 'short', timeZone: 'UTC' }) + ' ' : ''; };
    list.innerHTML = [...days].map(([d, rs]) => `<h3 class="cday">${esc(wd(d) + d)} <span class="hint">${rs.length}</span></h3>` + rs.map((r) => {
      const per = String(r['Per Nbr']);
      const pills = splitList(r.Pills, ';');
      const items = cardItems(per, r);
      const choose = plainChoose(r['Choose in NextGen']);
      const notes = splitList(r.Notes, '|');
      const high = pills.some((l) => (PILL_BY_LABEL.get(l) || {}).severity === 'high');
      const reverify = r['Re-verify'] === 'Yes';
      const ver = r['Verified Date'] ? `Verified ${esc(r['Verified Date'])}${r['Days Since Verified'] !== '' && r['Days Since Verified'] != null ? ` (${esc(r['Days Since Verified'])}d)` : ''}` : 'Not verified in Sohar';
      const ben = meterHtml('Deductible', r.Deductible, r['Deductible Remaining'], r['Deductible Status']) + meterHtml('OOP max', r['OOP Max'], r['OOP Remaining'], r['OOP Status']);
      return `<article class="pcard${high ? ' hi' : ''}">
        <div class="pc-head"><span class="pc-time">${esc(r['Next Time'] || '')}</span><b>${esc([r['Last Name'], r['First Name']].filter(Boolean).join(', '))}</b><span class="hint">Per ${esc(per)}</span></div>
        <div class="pc-sub">${esc([r['Next Event'], r['Next Provider']].filter(Boolean).join(' · '))}${!loc && r.Location ? ` <span class="hint">· ${esc(r.Location)}</span>` : ''}</div>
        ${pills.length ? `<div class="pills">${cardPills(pills).map((l) => pillHtml(l)).join(' ')}</div>` : ''}
        <dl class="pc-dl">
          <dt>Payer</dt><dd>${esc(r['Primary Payer'] || 'None on file')}${r['Secondary Payer'] ? ` <span class="hint">+ ${esc(r['Secondary Payer'])}</span>` : ''}</dd>
          ${r['Choose in NextGen'] ? `<dt>NextGen</dt><dd>${choose.quoted ? `Choose "${esc(choose.text)}"` : esc(choose.text)}</dd>` : ''}
          <dt>Sohar</dt><dd>${esc(plainStatus(r['Sohar Status']))}${r.Plan ? ` · ${esc(r.Plan)}` : ''}<br><span class="hint">${ver}</span></dd>
          ${r.Copay !== '' || r.Coinsurance !== '' ? `<dt>Copay</dt><dd>${r.Copay !== '' ? esc(usd(r.Copay) || r.Copay) : '—'}${r.Coinsurance !== '' ? ` · coins ${esc(r.Coinsurance)}` : ''}</dd>` : ''}
        </dl>
        ${ben ? `<div class="pc-ben">${ben}</div>` : ''}
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
    box.innerHTML = `<table class="watch clinicfiles"><thead><tr><th>Location</th><th>Patients</th><th>Flagged</th><th></th></tr></thead><tbody>${locs.map((l) => {
      const rs = clinicRowsAt(l);
      const flagged = rs.filter((r) => splitList(r['Review Items'], '|').length).length;
      return `<tr><td><b>${esc(l)}</b></td><td>${rs.length}</td><td>${flagged ? `<span class="hi">${flagged}</span>` : '0'}</td>
        <td class="acts"><button class="small" data-cfile="${esc(l)}">Save clinic file</button> <button class="small" data-cemail="${esc(l)}">Copy email</button></td></tr>`;
    }).join('')}</tbody></table>`;
    box.querySelectorAll('[data-cfile]').forEach((b) => b.addEventListener('click', () => guard(() => saveClinicFile(b.dataset.cfile), 'clinicFilesNote')));
    box.querySelectorAll('[data-cemail]').forEach((b) => b.addEventListener('click', () => guard(async () => {
      const l = b.dataset.cemail;
      const text = R.clinicEmail(l, clinicRowsAt(l), { runDate: runDateOf() });
      await copyOrShow(text, 'clinicFilesNote', `Email summary for ${l} copied (LASFIR + Per Nbr only). Paste it into Outlook or Teams.`);
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
    if (!state.soharPayers) state.soharPayers = await X.readSoharPayers();
    const res = C.mergeSoharResults(state.rows, acc.files, { ...cfg, payerMap: await X.readPayerMap(), soharPayers: state.soharPayers });
    if (res.error) throw new Error(res.error);
    step(`Matched ${res.counts.matched} results to patients (${res.counts.unmatched} unmatched)${res.counts.twoInsurances ? `; ${res.counts.twoInsurances} patient${res.counts.twoInsurances === 1 ? '' : 's'} had two insurances checked` : ''}. Writing the Sohar Results sheet…`);
    // Flags need the new results (re-verify, Medi-Cal from Sohar); the Results sheet gets a Pills column after RESULTS_COLS.
    state.results = res.results;
    computeFlags();
    const secBy = secondaryByPer();
    // Pills and Secondary Payer ride along for full screen (Secondary Payer is not written to the Results sheet).
    res.results = res.results.map((r) => ({ ...r, Pills: pillsFor(r['Per Nbr']), 'Secondary Payer': secBy.get(String(r['Per Nbr'])) || '' }));
    await X.writeResults(C.RESULTS_COLS.concat('Pills'), res.results);
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
      await refreshRules();
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
    await X.writeBenefits(C.BENEFIT_SHEET_COLS.concat('Pills'), view.rows.map((r) => ({ ...r, Pills: pillsFor(r['Per Nbr']) })), C.BENEFIT_PCT_COLS);
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
    const out = C.trackerExport(rows, fmtStamp(new Date()), secondaryByPer());
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
      // x (patient extras) stays in the pane; the window gets the Pills column instead.
      type: 'data', focus: focus || '', web: isWeb(), results: ((await currentResults()) || []).map(({ x, ...r }) => r), tasks: await currentTasks(), ir: await currentIr(),
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

  async function saveCSV(name, cols, rows, boxId) {
    const csv = C.toCSV(cols, rows);
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
    const w = window.open(new URL('save.html', location.href).href, '_blank');
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
    return { patName: 'Patient name', dob: 'DOB', perNbr: 'Per Nbr', mrn: 'MRN', apptDt: 'Appt date', begTm: 'Start time', event: 'Event', loc: 'Location', sts: 'Status', patSts: 'Patient status', payer: 'Payer', secPayer: 'Secondary payer', polNbr: 'Member ID', polNbrLstEnc: 'Member ID (last enc)', zip: 'ZIP', state: 'State', guarName: 'Guarantor', rendering: 'Rendering',
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
      'Mod Dt', 'Modified By', 'Measures Needed?', 'No Show/Late Cancel Charged?'];
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
    ];
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
      { requestid: 'disc-200103', id: '200103', responsepayerid: '95567', responsepayername: 'Health Net', created: iso(-2) },
    ].map((rec) => { const kind = rec.responsepayerid ? 'discovery' : 'verification'; return { Kind: kind, Key: C.resultKey(kind, rec), File: 'demo', Uploaded: iso(-1), Record: JSON.stringify(rec) }; });
    return {
      TABLES: { patients: 'tblPatients', review: 'tblReview', verification: 'tblVerification', discovery: 'tblDiscovery' },
      async readImportRows() { return rows; },
      async writeImportRows(r) { rows = r; },
      async importXlsx() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async readXlsxRows() { throw new Error('Demo mode reads .csv only. Inside Excel, .xlsx works too.'); },
      async writeResults() {}, async writeTasks() {}, async writeBenefits() {}, async readTasks() { return []; }, async readUploads() { return demoUploads; }, async readIrUploads() { return []; }, async writeIrUploads() {}, async writeUploads() {}, async readDiscoveryVerify() { return []; }, async writeDiscoveryVerify() {}, async tableRowCount() { return 0; }, async setTaskDone() {}, async writeIr() {}, async readIr() { return []; }, async writeExport(c, r) { return r.length; },
      async readPayerEvidence() { return []; }, async writePayerEvidence() {}, async tableExists() { return false; }, async readClinicFile() { return null; },
      async readConfig() { return { asOf: null, scope: 'upcoming', locationState: { 'San Rafael': 'CA', 'Newport Beach': 'CA', 'Mission Viejo': 'CA', 'Castro Street': 'CA' }, defaultState: 'CA', teleKeywords: ['Tele'], posTele: '10', posOffice: '11', specialty: 'PSY', useGuarantorAsSubscriber: false, approachingPct: 0.2, sfhpProviders: ['Stone', 'Reyes'], reverifyDays: 7, clinicLocation: '' }; },
      async readPayerMap() { return payerMap; },
      async readSoharPayers() { return sohar; },
      async upsertPayerMap(e) { return this.upsertPayerMapMany([e]); },
      async upsertPayerMapMany(list) { for (const e of list) payerMap = payerMap.filter((m) => m.nextgenPayerName.toLowerCase() !== e.nextgenPayerName.toLowerCase()).concat([e]); return list.length; },
      async replaceTable() {}, async setConfigValue() {}, async activate() {}, async templateCheck() { return []; },
    };
  }
})();
