/* dashboard.js — full-screen window. Runs in an Office dialog (its own window, so uploads and downloads work on the web).
   Data comes from the pane via messageChild; uploads and Done ticks go back via messageParent. */
(function () {
  'use strict';
  const C = window.SoharCore;
  const $ = (id) => document.getElementById(id);
  let rows = [], tasks = [], irRows = [], exportsData = null, scope = 'upcoming', asOf = null, list = [], tab = 'prep';
  let inOffice = false, web = false;
  const take = C.messageAssembler();

  const BEN = {
    dedMet: (r) => r['Deductible Status'] === 'Met',
    dedAppr: (r) => r['Deductible Status'] === 'Approaching',
    oopMet: (r) => r['OOP Status'] === 'Met',
    oopAppr: (r) => r['OOP Status'] === 'Approaching',
    anyFlag: (r) => /met|approaching/i.test(r['Benefit Flag'] || ''),
    noDed: (r) => r['Deductible Status'] === 'None',
    unknown: (r) => !r['Deductible Status'] && !r['OOP Status'],
  };

  // ---------- messages from the pane ----------
  function onParent(arg) {
    const m = take(arg.message);
    if (!m) return;
    if (m.type === 'data') {
      load(m);
      if (m.notice) notice(m.notice.text, m.notice.ok);
      if (m.focus) showTab(m.focus === 'exports' ? 'upload' : m.focus);
    } else if (m.type === 'notice') { notice(m.text, m.ok); $('upState').textContent = ''; }
  }
  function toPane(obj) {
    if (!inOffice) { notice('Demo mode: nothing is sent to Excel.', false); return; }
    for (const piece of C.chunkMessage(obj, 12000)) Office.context.ui.messageParent(piece);
  }

  function load(m) {
    rows = m.results || []; tasks = m.tasks || []; irRows = m.ir || []; exportsData = m.exports || null; scope = m.scope || 'upcoming'; web = !!m.web;
    asOf = m.asOf ? new Date(m.asOf) : null;
    $('dStatus').textContent = rows.length
      ? `${rows.length} Sohar results, ${tasks.filter(open).length} open NextGen tasks${asOf ? `, appointments from ${C.fmtDate(asOf)}` : ''}.`
      : 'No Sohar results yet. Upload the verification and discovery outputs on the Upload & save tab.';
    $('dStatus').className = `status ${rows.length ? 'ok' : 'bad'}`;
    const q = (name) => rows.filter((x) => x['Work Queue'] === name).length;
    $('tAct').textContent = q('Action needed'); $('tUnm').textContent = q('Unmatched');
    $('tPen').textContent = q('Pending'); $('tClr').textContent = q('Clear');
    $('tDedMet').textContent = rows.filter(BEN.dedMet).length; $('tDedAppr').textContent = rows.filter(BEN.dedAppr).length;
    $('tOopMet').textContent = rows.filter(BEN.oopMet).length; $('tOopAppr').textContent = rows.filter(BEN.oopAppr).length;
    $('benNote').hidden = !rows.length || rows.some((x) => x['Deductible Status'] || x['OOP Status']);
    const locs = [...new Set([...rows, ...tasks].map((x) => x['Next Location']).filter(Boolean))].sort();
    const keepLoc = $('fLoc').value;
    $('fLoc').innerHTML = '<option value="">All</option>' + locs.map((l) => `<option>${esc(l)}</option>`).join('');
    $('fLoc').value = locs.includes(keepLoc) ? keepLoc : '';
    const types = [...new Set(tasks.map((t) => t.Task))].sort((a, b) => C.TASK_TYPES.indexOf(a) - C.TASK_TYPES.indexOf(b));
    $('fTask').innerHTML = '<option value="">All tasks</option>' + types.map((t) => `<option>${esc(t)}</option>`).join('');
    if (!rows.length && tab === 'prep') showTab('upload');
    render();
  }

  const open = (t) => !String(t.Done || '').trim();

  // ---------- filters ----------
  function shared(r) {
    const sv = $('fSent').value, lv = $('fLoc').value, tv = $('fText').value.trim().toLowerCase();
    const from = $('fFrom').value ? C.parseDate($('fFrom').value) : null;
    const to = $('fTo').value ? C.parseDate($('fTo').value) : null;
    if (sv && (r['Sent As'] || r['Result Type']) !== sv) return false;
    if (lv && r['Next Location'] !== lv) return false;
    if (from || to) {
      const d = C.parseDate(r['Next Appt']);
      if (!d || (from && d < from) || (to && d > to)) return false;
    }
    if (tv && !['Last Name', 'First Name', 'Per Nbr', 'MRN', 'NextGen Payer', 'Sohar Payer', 'Plan', 'Choose in NextGen', 'Task', 'Details'].some((k) => String(r[k] || '').toLowerCase().includes(tv))) return false;
    return true;
  }
  function filteredRows() {
    const qv = $('fQueue').value, bv = $('fBen').value;
    return rows.filter((r) => shared(r) && (!qv || r['Work Queue'] === qv) && (!bv || BEN[bv](r)));
  }
  function filteredTasks() {
    const dv = $('fDone').value, kv = $('fTask').value;
    return tasks.filter((t) => shared(t) && (!kv || t.Task === kv) && (dv === '' || (dv === 'open' ? open(t) : !open(t))));
  }

  // ---------- render ----------
  function render() {
    $('taskCount').textContent = tasks.filter(open).length || '';
    renderPrep(); renderWorklist(); renderTasks(); renderMessage();
  }

  const cell = (v, cls) => `<td${cls ? ` class="${cls}"` : ''}>${esc(v)}</td>`;
  const left = (rem, total) => (rem === '' || rem == null ? '' : total !== '' && total != null ? `${rem} of ${total}` : String(rem));
  const name = (r) => [r['Last Name'], r['First Name']].filter(Boolean).join(', ') || '(not matched)';
  function link(u) { return /^https:\/\//.test(u || '') ? `<a href="${esc(u)}" target="_blank" rel="noopener">Sohar</a>` : ''; }

  function renderPrep() {
    const rs = rows.filter(shared);
    const byPer = new Map();
    for (const t of tasks) if (open(t)) { const k = t['Per Nbr'] || `ref:${t['Sohar Ref']}`; byPer.set(k, (byPer.get(k) || 0) + 1); }
    const days = new Map();
    for (const r of rs) {
      const d = r['Next Appt'] || 'No upcoming appointment';
      if (!days.has(d)) days.set(d, []);
      days.get(d).push(r);
    }
    const order = [...days.keys()].sort((a, b) => (C.parseDate(a) || new Date(8.64e15)) - (C.parseDate(b) || new Date(8.64e15)));
    const wd = (s) => { const d = C.parseDate(s); return d ? d.toLocaleDateString('en-US', { weekday: 'short', timeZone: 'UTC' }) + ' ' : ''; };
    const minutes = (t) => { const m = String(t || '').match(/(\d+):(\d+)\s*(AM|PM)/i); return m ? ((+m[1] % 12) + (/pm/i.test(m[3]) ? 12 : 0)) * 60 + +m[2] : 9999; };
    $('prepList').innerHTML = order.length ? order.map((d) => {
      const day = days.get(d).sort((a, b) => minutes(a['Next Time']) - minutes(b['Next Time']) || name(a).localeCompare(name(b)));
      const need = day.filter((r) => byPer.get(r['Per Nbr'] || `ref:${r['Sohar Ref']}`)).length;
      return `<div class="day"><h3>${esc(wd(d) + d)} <span class="hint">${day.length} patient${day.length > 1 ? 's' : ''}${need ? ` · ${need} with NextGen tasks` : ''}</span></h3>
        <table class="prep"><tbody>${day.map((r) => {
          const n = byPer.get(r['Per Nbr'] || `ref:${r['Sohar Ref']}`) || 0;
          return `<tr class="q-${queueClass(r['Work Queue'])}" data-key="${esc(r['Per Nbr'] + '|' + r['Sohar Ref'])}" tabindex="0">
            ${cell(r['Next Time'])}<td><b>${esc(name(r))}</b><br><span class="hint">Per ${esc(r['Per Nbr'] || '?')}${r.MRN ? ` · MRN ${esc(r.MRN)}` : ''}</span></td>
            <td>${esc(r['Next Event'])}<br><span class="hint">${esc([r['Next Location'], r['Next Provider']].filter(Boolean).join(' · '))}</span></td>
            <td><span class="pill ${r['Sent As'] === 'Discovery' ? 'd' : 'v'}">${esc(r['Sent As'])}</span><br><span class="hint">${esc(r['Sohar Status'])}</span></td>
            <td>${esc(r['Benefit Flag'] || (r['Sent As'] === 'Verification' ? '' : ''))}${r.Copay !== '' && r.Copay != null ? `<br><span class="hint">Copay ${esc(r.Copay)}${r.Coinsurance ? ` · ${esc(r.Coinsurance)}` : ''}</span>` : ''}</td>
            <td class="wide">${n ? `<span class="pill t">${n} task${n > 1 ? 's' : ''}</span> ` : ''}${esc(r.Action)}${r['Choose in NextGen'] ? `<br><span class="hint">NextGen: ${esc(r['Choose in NextGen'])}</span>` : ''}</td>
          </tr>`;
        }).join('')}</tbody></table></div>`;
    }).join('') : '<p class="hint">No patients match these filters.</p>';
  }

  function renderWorklist() {
    list = filteredRows();
    $('shown').textContent = `Showing ${list.length} of ${rows.length}`;
    const st = (s) => ({ Met: 'met', Approaching: 'appr' }[s] || '');
    $('tbl').querySelector('tbody').innerHTML = list.map((r) => `<tr class="q-${queueClass(r['Work Queue'])}" data-key="${esc(r['Per Nbr'] + '|' + r['Sohar Ref'])}" tabindex="0">
      ${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r['Next Location'])}${cell(name(r))}${cell(r['Per Nbr'])}${cell(r['Sent As'])}
      ${cell(r['Sohar Status'])}${cell(r['Work Queue'])}${cell(r.Action, 'wide')}${cell(r['Choose in NextGen'], 'mid')}
      ${cell(r['Benefit Flag'])}${cell(r.Copay)}${cell(r.Coinsurance)}
      ${cell(left(r['Deductible Remaining'], r.Deductible), st(r['Deductible Status']))}${cell(left(r['OOP Remaining'], r['OOP Max']), st(r['OOP Status']))}
      ${cell(r.Plan, 'mid')}<td>${link(r['Sohar Link'])}</td></tr>`).join('');
  }

  function renderTasks() {
    const ts = filteredTasks();
    $('taskShown').textContent = `Showing ${ts.length} of ${tasks.length} (${tasks.filter(open).length} open)`;
    $('tasksTbl').querySelector('tbody').innerHTML = ts.map((t) => `<tr class="${open(t) ? '' : 'done'}">
      <td><input type="checkbox" data-task="${esc(t['Task Key'])}" ${open(t) ? '' : 'checked'} aria-label="Done: ${esc(t.Task)} for ${esc(name(t))}"></td>
      ${cell([t['Due By'], t['Next Time']].filter(Boolean).join(' '))}${cell(t['Next Location'])}${cell(name(t))}${cell(t['Per Nbr'])}${cell(t['Result Type'])}
      ${cell(t.Task, 'mid')}${cell(t.Details, 'wide')}${cell(t['Choose in NextGen'], 'mid')}<td>${link(t['Sohar Link'])}</td></tr>`).join('');
  }

  function renderMessage() {
    const mode = $('msgMode').value;
    $('msg').value = mode === 'tasks' ? C.taskMessage(filteredTasks().filter(open), { asOf })
      : mode === 'ir' ? (irRows.length ? C.irMessage(irRows) : 'No IR Portal cross-check yet. Upload the IR Portal file in step 6 of the pane.\n')
      : C.adminMessage(rows.filter(shared), { asOf, mode });
  }

  // ---------- patient view ----------
  function showDetail(key) {
    const r = rows.find((x) => `${x['Per Nbr']}|${x['Sohar Ref']}` === key);
    if (!r) return;
    const its = tasks.filter((t) => (r['Per Nbr'] ? t['Per Nbr'] === r['Per Nbr'] : t['Sohar Ref'] === r['Sohar Ref']));
    const others = r['Per Nbr'] ? rows.filter((x) => x !== r && x['Per Nbr'] === r['Per Nbr']) : [];
    const grp = (title, keys, src) => `<div class="dgroup"><h3>${title}</h3><dl>${keys.filter((k) => (src || r)[k] !== '' && (src || r)[k] != null).map((k) => `<dt>${esc(k)}</dt><dd>${k === 'Sohar Link' ? link((src || r)[k]) : esc((src || r)[k])}</dd>`).join('') || '<dd class="hint">None</dd>'}</dl></div>`;
    const box = $('detail');
    box.innerHTML = `<div class="dhead"><h2>${esc(name(r))}</h2>
        <span class="hint">Per Nbr ${esc(r['Per Nbr'] || '?')}${r.MRN ? ` · MRN ${esc(r.MRN)}` : ''}${r.DOB ? ` · DOB ${esc(r.DOB)}` : ''}</span>
        <button class="small" id="btnHide">Close</button></div>
      <div class="dgrid">
        ${grp('Upcoming visits', ['Next Appt', 'Next Time', 'Next Event', 'Next Location', 'Next Provider', 'Upcoming Appts', 'Upcoming Detail'])}
        <div class="dgroup"><h3>NextGen tasks</h3>${its.length ? `<ul class="tasklist">${its.map((t) => `<li class="${open(t) ? '' : 'done'}"><label><input type="checkbox" data-task="${esc(t['Task Key'])}" ${open(t) ? '' : 'checked'}> <b>${esc(t.Task)}</b></label><br>${esc(t.Details)}</li>`).join('')}</ul>` : '<p class="hint">None</p>'}</div>
        ${grp('Benefits (in-network)', ['Benefit Flag', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status', 'Coverage Level'])}
        ${grp('Coverage', ['Sent As', 'Sohar Status', 'Action', 'Choose in NextGen', 'NextGen Payer', 'NextGen Member ID', 'Sohar Payer', 'Sohar payerId', 'Sohar Member ID', 'Plan', 'Network', 'Coverage Start', 'Coverage End', 'Match', 'Sohar Date', 'Sohar Link'])}
        ${others.map((o) => grp(`Also: ${esc(o['Sent As'])} result`, ['Sohar Status', 'Action', 'Sohar Payer', 'Benefit Flag', 'Sohar Date'], o)).join('')}
      </div>`;
    box.hidden = false;
    $('btnHide').addEventListener('click', () => { box.hidden = true; });
    box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function queueClass(q) { return { 'Action needed': 'act', Unmatched: 'unm', Pending: 'pen', Clear: 'clr' }[q] || ''; }

  // ---------- tabs ----------
  function showTab(name) {
    tab = name;
    document.querySelectorAll('[data-tab]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === name)));
    document.querySelectorAll('[data-panel]').forEach((p) => { p.hidden = p.dataset.panel !== name; });
    $('sharedFilters').hidden = name === 'upload';
    $('detail').hidden = true;
  }

  // ---------- uploads ----------
  async function readFiles(files) {
    const out = [];
    for (const f of files) {
      if (/\.csv$/i.test(f.name)) {
        const buf = await f.arrayBuffer();
        let text; try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { text = new TextDecoder('windows-1252').decode(buf); }
        out.push({ name: f.name, text });
      } else if (/\.xls[xm]$/i.test(f.name)) {
        const bytes = new Uint8Array(await f.arrayBuffer());
        let bin = '';
        for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        out.push({ name: f.name, base64: btoa(bin) });
      } else throw new Error(`${f.name} is not .csv or .xlsx.`);
    }
    return out;
  }
  async function upload(files) {
    if (!files.length) return;
    try {
      const fl = await readFiles([...files]);
      const kinds = fl.filter((f) => f.text != null).map((f) => `${f.name}: ${C.resultKind(C.parseCSV(f.text.split(/\r?\n/, 1)[0])[0] || []) || 'not a Sohar output'}`);
      $('upState').textContent = `Sending ${fl.length} file${fl.length > 1 ? 's' : ''} to Excel… ${kinds.join('; ')}`;
      toPane({ type: 'upload', files: fl });
    } catch (e) { notice(e.message, false); }
  }

  // ---------- saving ----------
  const stamp = () => new Date().toISOString().slice(0, 10).replace(/-/g, '');
  function save(kind) {
    let name, cols, data;
    if (kind === 'verification' || kind === 'discovery') {
      if (!exportsData) { notice('No Sohar import built yet. Run step 3 (Build) in the pane first.', false); return; }
      name = `sohar_${kind}_import_${scope}_${stamp()}.csv`;
      cols = kind === 'verification' ? C.VERIFICATION_COLS : C.DISCOVERY_COLS; data = exportsData[kind] || [];
    } else if (kind === 'tasks') { name = `nextgen_tasks_open_${stamp()}.csv`; cols = C.TASK_COLS; data = tasks.filter(open); }
    else if (kind === 'allTasks') { name = `nextgen_tasks_all_${stamp()}.csv`; cols = C.TASK_COLS; data = tasks; }
    else if (kind === 'benefits') { name = `sohar_benefits_${stamp()}.csv`; cols = C.BENEFITS_EXPORT_COLS; data = C.benefitsExport(rows); }
    else if (kind === 'ir') {
      if (!irRows.length) { notice('No IR Portal cross-check yet. Upload the IR Portal file in step 6 of the pane.', false); return; }
      name = `ir_portal_cross_reference_${stamp()}.csv`; cols = C.IR_COLS; data = irRows;
    }
    else { name = `sohar_results_filtered_${stamp()}.csv`; cols = C.RESULTS_COLS; data = filteredRows(); }
    // Excel on the web blocks downloads here too, so the pane writes the rows to the Export sheet instead.
    if (inOffice && web) { notice('Writing to the Export sheet…', true); toPane({ type: 'export', name, cols, rows: data }); return; }
    const csv = C.toCSV(cols, data);
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    notice(`Saved ${name}.`, true);
  }

  function notice(text, ok) { const n = $('notice'); n.hidden = false; n.textContent = text; n.className = `notice ${ok ? 'ok' : 'bad'}`; }
  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

  // ---------- wiring ----------
  document.querySelectorAll('[data-tab]').forEach((b) => b.addEventListener('click', () => showTab(b.dataset.tab)));
  ['fSent', 'fLoc', 'fFrom', 'fTo', 'fQueue', 'fBen', 'fDone', 'fTask', 'msgMode'].forEach((id) => $(id).addEventListener('change', render));
  $('fText').addEventListener('input', render);
  document.querySelectorAll('[data-queue]').forEach((b) => b.addEventListener('click', () => { $('fQueue').value = b.dataset.queue; $('fBen').value = ''; render(); }));
  document.querySelectorAll('[data-benefit]').forEach((b) => b.addEventListener('click', () => { $('fBen').value = b.dataset.benefit; $('fQueue').value = ''; render(); }));
  document.addEventListener('click', (e) => {
    if (e.target.closest('a, input, button')) return;
    const tr = e.target.closest('tr[data-key]');
    if (tr) showDetail(tr.dataset.key);
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Enter') { const tr = e.target.closest && e.target.closest('tr[data-key]'); if (tr) showDetail(tr.dataset.key); } });
  document.addEventListener('change', (e) => {
    const cb = e.target.closest && e.target.closest('input[data-task]');
    if (!cb) return;
    const key = cb.dataset.task;
    tasks = tasks.map((t) => (t['Task Key'] === key ? { ...t, Done: cb.checked ? 'Yes' : '' } : t));
    toPane({ type: 'task-done', key, done: cb.checked });
    document.querySelectorAll(`input[data-task="${CSS.escape(key)}"]`).forEach((x) => { x.checked = cb.checked; });
    $('taskCount').textContent = tasks.filter(open).length || '';
    renderTasks(); renderPrep(); renderMessage();
  });
  $('drop').addEventListener('click', () => $('fileUp').click());
  $('drop').addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileUp').click(); } });
  $('fileUp').addEventListener('change', (e) => { upload(e.target.files); e.target.value = ''; });
  $('drop').addEventListener('dragover', (e) => { e.preventDefault(); $('drop').classList.add('over'); });
  $('drop').addEventListener('dragleave', () => $('drop').classList.remove('over'));
  $('drop').addEventListener('drop', (e) => { e.preventDefault(); $('drop').classList.remove('over'); upload(e.dataTransfer.files); });
  document.querySelectorAll('[data-save]').forEach((b) => b.addEventListener('click', () => save(b.dataset.save)));
  $('btnCopy').addEventListener('click', async () => {
    const ta = $('msg');
    let ok = false;
    try { await navigator.clipboard.writeText(ta.value); ok = true; } catch (e) {
      ta.focus(); ta.select(); try { ok = document.execCommand('copy'); } catch (e2) { ok = false; }
    }
    notice(ok ? 'Message copied. Paste it into Outlook or Teams.' : 'Copy was blocked. Select the message text and press Ctrl+C.', ok);
  });

  // Browser preview outside Office: synthetic rows only.
  function demo() {
    const b = document.createElement('div');
    b.className = 'banner';
    b.textContent = 'Demo mode. Synthetic data only.';
    document.body.insertBefore(b, document.querySelector('nav'));
    const mk = (o) => { const r = Object.fromEntries(C.RESULTS_COLS.map((c) => [c, ''])); return Object.assign(r, o); };
    const res = [
      mk({ 'Next Appt': '10/20/2026', 'Next Time': '9:00 AM', 'Next Event': 'TMS Treatment Unilateral', 'Next Location': 'San Rafael', 'Next Provider': 'Dr Test', 'Upcoming Appts': 12, 'Upcoming Detail': '10/20 TMS Treatment Unilateral; 10/21 TMS Treatment Unilateral; +10 more', 'Per Nbr': '100100', MRN: '80100', 'Last Name': 'Rivera', 'First Name': 'Ana', DOB: '03/04/1980', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', Action: 'Eligible. No action.', Plan: 'SYNTH GOLD PPO', Network: 'PR', 'Coverage Level': 'IND', Copay: '40', Coinsurance: '20%', Deductible: '1500', 'Deductible Remaining': '0', 'Deductible Status': 'Met', 'OOP Max': '5000', 'OOP Remaining': '800', 'OOP Status': 'Approaching', 'NextGen Payer': 'Blue Shield CA', 'Sohar Payer': 'Blue Shield of California', Match: 'Per Nbr', 'Sohar Ref': 'q-1' }),
      mk({ 'Next Appt': '10/20/2026', 'Next Time': '1:30 PM', 'Next Event': 'Telepsych', 'Next Location': 'San Rafael', 'Upcoming Appts': 1, 'Per Nbr': '100101', 'Last Name': 'Okafor', 'First Name': 'Chidi', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', Action: 'Coverage found. Add it to NextGen, then verify. Choose "Cigna" in NextGen. Get the member ID from the Sohar link.', 'Choose in NextGen': 'Cigna', 'Sohar Payer': 'Cigna', 'Sohar payerId': '62308', 'Coverage Differs': 'Yes', Match: 'Per Nbr', 'Sohar Link': 'https://example.test/r-1', 'Sohar Ref': 'r-1' }),
      mk({ 'Next Appt': '10/22/2026', 'Next Time': '10:00 AM', 'Next Event': 'Med Management In-Office', 'Next Location': 'Walnut Creek', 'Upcoming Appts': 2, 'Per Nbr': '100102', 'Last Name': 'Lindqvist', 'First Name': 'Sven', 'Sent As': 'Verification', 'Sohar Status': 'error.member.id', Action: 'Member ID rejected. Check the card, fix in NextGen, re-verify.', 'NextGen Payer': 'Aetna', 'NextGen Member ID': 'SYN555', Match: 'Per Nbr', 'Sohar Ref': 'q-2' }),
      mk({ 'Next Appt': '10/23/2026', 'Next Time': '11:00 AM', 'Next Event': 'Telepsych', 'Next Location': 'Walnut Creek', 'Upcoming Appts': 4, 'Per Nbr': '100103', 'Last Name': 'Mensah', 'First Name': 'Afia', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', Action: 'Eligible. No action.', Copay: '0', Coinsurance: '0%', Deductible: '0', 'Deductible Remaining': '0', 'Deductible Status': 'None', 'OOP Max': '3000', 'OOP Remaining': '0', 'OOP Status': 'Met', 'NextGen Payer': 'Cigna', 'Sohar Payer': 'Cigna', Plan: 'OPEN ACCESS PLUS', Match: 'Per Nbr', 'Sohar Ref': 'q-3' }),
      mk({ 'Sent As': 'Verification', 'Sohar Status': 'error.member.id', Action: 'Not matched to a patient in the Import report. Member ID rejected. Check the card, fix in NextGen, re-verify.', Match: 'No match', 'Sohar Ref': 'q-9' }),
    ].map((r) => ({ ...r, 'Benefit Flag': C.benefitFlag(r), 'Work Queue': C.workQueue(r) }));
    load({ results: res, tasks: C.buildTasks(res), exports: { verification: [], discovery: [] }, asOf: '2026-10-13T00:00:00.000Z' });
    if (location.hash) showTab(location.hash.slice(1) === 'exports' ? 'upload' : location.hash.slice(1));
  }

  // Boot last, so everything above is initialized before demo() or the first message can use it.
  if (window.Office && Office.onReady) {
    Office.onReady(() => {
      if (Office.context && Office.context.ui && Office.context.ui.messageParent) {
        inOffice = true;
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParent, () => {
          Office.context.ui.messageParent(JSON.stringify({ type: 'ready' }));
        });
        $('btnClose').hidden = false;
        $('btnClose').addEventListener('click', () => Office.context.ui.messageParent(JSON.stringify({ type: 'close' })));
      } else demo();
    });
  } else demo();
})();
