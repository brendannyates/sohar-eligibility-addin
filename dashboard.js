/* dashboard.js — full screen. Runs in an Office dialog that floats over the Excel page (displayInIframe: true).
   Data comes from the pane via messageChild; uploads and Done ticks go back via messageParent. */
(function () {
  'use strict';
  const C = window.SoharCore;
  const $ = (id) => document.getElementById(id);
  let rows = [], tasks = [], irRows = [], exportsData = null, discVerify = [], verFailed = [], scope = 'upcoming', asOf = null, list = [], tab = 'prep';
  let inOffice = false, web = false;
  // v0.12: clinic lists, the skipped list and the Payer learning views (from the pane).
  let clinic = null, skipped = [], skipCols = [], learning = null, guarantor = { count: 0, noDob: 0 };
  const R = window.SoharRules;
  // Patient key (core.keyOf): the MRN, else PER-<Per Nbr>. Rows, tasks and the detail card join on it.
  const keyOf = (r) => C.keyOf(r);
  const idText = (r) => C.idLabel(r) || '?';
  // Fee estimates (v0.11, a test): null unless the pane sends them (Config "Fee estimates" On). estOver: the open
  // estimate panel's codes and toggles.
  const F = window.SoharFees;
  let fees = null, estOver = null;
  const take = C.messageAssembler();
  // Flag pills come from the pane as a Pills column ("Medi-Cal; Re-verify"). Colors from rules.js.
  const PILL_COLOR = new Map(((window.SoharRules && window.SoharRules.PILLS) || []).map((p) => [p.label, p.color]));
  const pillsOf = (r) => String(r.Pills || '').split(';').map((s) => s.trim()).filter(Boolean);
  // "Kaiser/VA MM" is a filter category; next to "Kaiser/VA MM no auth" it only repeats it, so rows hide it.
  const shownPills = (r) => { const p = pillsOf(r); return p.includes('Kaiser/VA MM no auth') ? p.filter((l) => l !== 'Kaiser/VA MM') : p; };
  const pillsHtml = (r) => shownPills(r).map((l) => `<span class="pill p-${PILL_COLOR.get(l) || 'grey'}">${esc(l)}</span>`).join(' ');

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
    else if (m.type === 'fee-update') { fees = m.fees || null; renderFeeBar(); renderBenefits(); refreshEstimate(); }
  }
  function toPane(obj) {
    if (!inOffice) { notice('Demo mode: nothing is sent to Excel.', false); return; }
    for (const piece of C.chunkMessage(obj, 12000)) Office.context.ui.messageParent(piece);
  }

  function load(m) {
    clinic = m.clinic || null; skipped = m.skipped || []; skipCols = m.skipCols || []; learning = m.learning || null; guarantor = m.guarantor || { count: 0, noDob: 0 };
    rows = m.results || []; tasks = m.tasks || []; irRows = m.ir || []; exportsData = m.exports || null; discVerify = m.discoveryVerify || []; verFailed = m.verifyFailed || []; scope = m.scope || 'upcoming'; web = !!m.web;
    asOf = m.asOf ? new Date(m.asOf) : null;
    fees = F && m.fees && m.fees.on ? m.fees : null;
    renderFeeBar();
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
    const pillSet = new Set(rows.flatMap(pillsOf));
    const order = [...PILL_COLOR.keys()];
    const pl = [...pillSet].sort((a, b) => (order.indexOf(a) + 1 || 99) - (order.indexOf(b) + 1 || 99));
    const keepPill = $('fPill').value;
    $('fPill').innerHTML = '<option value="">Any</option>' + pl.map((l) => `<option>${esc(l)}</option>`).join('');
    $('fPill').value = pl.includes(keepPill) ? keepPill : '';
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
    const pv = $('fPill').value;
    // Results rows carry Pills; tasks and benefits rows borrow them from the patient's results by patient key (MRN).
    if (pv && (r.Pills != null ? !pillsOf(r).includes(pv) : !rows.some((x) => keyOf(x) && keyOf(x) === keyOf(r) && pillsOf(x).includes(pv)))) return false;
    if (from || to) {
      const d = C.parseDate(r['Next Appt']);
      if (!d || (from && d < from) || (to && d > to)) return false;
    }
    if (tv && !['Last Name', 'First Name', 'MRN', 'Per Nbr', 'NextGen Payer', 'Sohar Payer', 'Plan', 'Choose in NextGen', 'Task', 'Details'].some((k) => String(r[k] || '').toLowerCase().includes(tv))) return false;
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
    renderPrep(); renderBenefits(); renderWorklist(); renderTasks(); renderMessage(); renderClinicLists(); renderLearning(); renderGuarantorCall();
  }

  const usd = (v) => { const n = parseFloat(String(v ?? '').replace(/[$,]/g, '')); return isNaN(n) ? '' : `$${n.toLocaleString('en-US', { maximumFractionDigits: 0 })}`; };
  /** "$2,700 of $3,000" with a bar, or "Met" / "None". */
  function progressCell(met, total, pct, status) {
    if (status === 'None') return '<td class="hint">None</td>';
    if (pct === '' || pct == null) return `<td>${esc(status || '')}</td>`;
    const cls = status === 'Met' ? 'met' : status === 'Approaching' ? 'appr' : '';
    return `<td class="${cls}"><span class="meter" data-pct="${pct}"><span></span></span> ${Math.round(pct * 100)}%<br><span class="hint">${esc(usd(met))} of ${esc(usd(total))}</span></td>`;
  }
  function renderBenefits() {
    const all = C.benefitsView(rows).rows;
    const v = $('fBenView').value;
    const bs = all.filter((r) => shared(r)
      && (!v || (v === 'oop' ? /^(Met|Approaching)$/.test(r['OOP Status']) : /^(Met|Approaching)$/.test(r['Deductible Status']))));
    $('benShown').textContent = `Showing ${bs.length} of ${all.length} verified`;
    const tb = $('benTbl').querySelector('tbody');
    tb.innerHTML = bs.map((r) => `<tr data-key="${esc(keyOf(r) + '|' + ((rows.find((x) => keyOf(x) === keyOf(r) && x['Sent As'] === 'Verification') || {})['Sohar Ref'] || ''))}" tabindex="0">
      <td><b>${esc(name(r))}</b></td>${cell(idText(r))}${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r['Next Location'])}${cell(r.Service)}
      ${cell([r['Sohar Payer'], r.Plan].filter(Boolean).join(' · '), 'mid')}${cell(r.Copay === '' ? '' : usd(r.Copay))}${cell(r.Coinsurance)}
      ${progressCell(r['OOP Met'], r['OOP Max'], r['OOP % Met'], r['OOP Status'])}${cell(r['OOP Remaining'] === '' ? '' : usd(r['OOP Remaining']))}
      ${progressCell(r['Deductible Met'], r.Deductible, r['Deductible % Met'], r['Deductible Status'])}${cell(r['Deductible Remaining'] === '' ? '' : usd(r['Deductible Remaining']))}
      ${r.Flag ? cell(r.Flag, r.Flag === 'OOP max met' ? 'flag met' : r.Flag.startsWith('OOP') ? 'flag' : '') : r['OOP Max'] === '' && r.Deductible === '' ? '<td class="hint">No amounts from Sohar</td>' : cell('')}${fees ? estCells(r['Per Nbr']) : ''}</tr>`).join('') || `<tr><td colspan="${fees ? 15 : 13}" class="hint">${all.length ? 'No one matches these filters.' : 'No verified benefits yet. Benefits come from Sohar\'s verification output; discovery alone has none.'}</td></tr>`;
    tb.querySelectorAll('.meter').forEach((m) => { m.firstElementChild.style.width = `${Math.round(Math.min(1, +m.dataset.pct) * 100)}%`; });
  }

  const cell = (v, cls) => `<td${cls ? ` class="${cls}"` : ''}>${esc(v)}</td>`;
  const left = (rem, total) => (rem === '' || rem == null ? '' : total !== '' && total != null ? `${rem} of ${total}` : String(rem));
  const name = (r) => [r['Last Name'], r['First Name']].filter(Boolean).join(', ') || '(not matched)';
  function link(u) { return /^https:\/\//.test(u || '') ? `<a href="${esc(u)}" target="_blank" rel="noopener">Sohar</a>` : ''; }

  function renderPrep() {
    const rs = rows.filter(shared);
    const byPer = new Map();
    for (const t of tasks) if (open(t)) { const k = keyOf(t) || `ref:${t['Sohar Ref']}`; byPer.set(k, (byPer.get(k) || 0) + 1); }
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
      const need = day.filter((r) => byPer.get(keyOf(r) || `ref:${r['Sohar Ref']}`)).length;
      return `<div class="day"><h3>${esc(wd(d) + d)} <span class="hint">${day.length} patient${day.length > 1 ? 's' : ''}${need ? ` · ${need} with NextGen tasks` : ''}</span></h3>
        <table class="prep"><tbody>${day.map((r) => {
          const n = byPer.get(keyOf(r) || `ref:${r['Sohar Ref']}`) || 0;
          return `<tr class="q-${queueClass(r['Work Queue'])}" data-key="${esc(keyOf(r) + '|' + r['Sohar Ref'])}" tabindex="0">
            ${cell(r['Next Time'])}<td><b>${esc(name(r))}</b><br><span class="hint">${esc(idText(r))}${r['Searched As'] && r['Searched As'] !== 'Patient' ? ` · searched as ${esc(r['Searched As'].toLowerCase())}` : ''}</span>${pillsOf(r).length ? `<div class="pills">${pillsHtml(r)}</div>` : ''}</td>
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
    $('tbl').querySelector('tbody').innerHTML = list.map((r) => `<tr class="q-${queueClass(r['Work Queue'])}" data-key="${esc(keyOf(r) + '|' + r['Sohar Ref'])}" tabindex="0">
      ${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r['Next Location'])}${cell(name(r))}${cell(idText(r))}<td class="pillc">${pillsHtml(r)}</td>${cell(r['Sent As'])}
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
      ${cell([t['Due By'], t['Next Time']].filter(Boolean).join(' '))}${cell(t['Next Location'])}${cell(name(t))}${cell(idText(t))}${cell(t['Result Type'])}
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
    const r = rows.find((x) => `${keyOf(x)}|${x['Sohar Ref']}` === key);
    if (!r) return;
    const k = keyOf(r);
    const its = tasks.filter((t) => (k ? keyOf(t) === k : t['Sohar Ref'] === r['Sohar Ref']));
    const others = k ? rows.filter((x) => x !== r && keyOf(x) === k) : [];
    const grp = (title, keys, src) => `<div class="dgroup"><h3>${title}</h3><dl>${keys.filter((k) => (src || r)[k] !== '' && (src || r)[k] != null).map((k) => `<dt>${esc(k)}</dt><dd>${k === 'Sohar Link' ? link((src || r)[k]) : esc((src || r)[k])}</dd>`).join('') || '<dd class="hint">None</dd>'}</dl></div>`;
    const box = $('detail');
    box.innerHTML = `<div class="dhead"><h2>${esc(name(r))}</h2>
        <span class="hint">${esc(idText(r))}${r['Per Nbr'] && r.MRN ? ` · Per Nbr ${esc(r['Per Nbr'])}` : ''}${r.DOB ? ` · DOB ${esc(r.DOB)}` : ''}</span>
        <button class="small" id="btnHide">Close</button></div>
      <div class="dgrid">
        ${grp('Upcoming visits', ['Next Appt', 'Next Time', 'Next Event', 'Next Location', 'Next Provider', 'Upcoming Appts', 'Upcoming Detail'])}
        <div class="dgroup"><h3>NextGen tasks</h3>${its.length ? `<ul class="tasklist">${its.map((t) => `<li class="${open(t) ? '' : 'done'}"><label><input type="checkbox" data-task="${esc(t['Task Key'])}" ${open(t) ? '' : 'checked'}> <b>${esc(t.Task)}</b></label><br>${esc(t.Details)}</li>`).join('')}</ul>` : '<p class="hint">None</p>'}</div>
        ${grp('Benefits (in-network)', ['Benefit Flag', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status', 'Coverage Level'])}
        ${grp('Coverage', ['Sent As', 'Searched As', 'Sohar Status', 'Action', 'Choose in NextGen', 'NextGen Payer', 'NextGen Member ID', 'Sohar Payer', 'Sohar payerId', 'Sohar Member ID', 'Plan', 'Network', 'Coverage Start', 'Coverage End', 'Match', 'Sohar Date', 'Sohar Link'])}
        ${others.map((o) => grp(`Also: ${esc(o['Sent As'])} result`, ['Sohar Status', 'Action', 'Sohar Payer', 'Benefit Flag', 'Sohar Date'], o)).join('')}
        ${fees ? `<div class="dgroup est" id="estPanel" data-per="${esc(r['Per Nbr'])}"></div>` : ''}
      </div>`;
    box.hidden = false;
    if (fees) renderEstimatePanel(r['Per Nbr']);
    $('btnHide').addEventListener('click', hideDetail);
    if (!$('detailBackdrop')){ const bd = document.createElement('div'); bd.className = 'backdrop'; bd.id = 'detailBackdrop'; bd.addEventListener('click', hideDetail); document.body.append(bd); }
    $('btnHide').focus();
  }
  // the detail card floats over the dashboard; everything behind it is blurred
  function hideDetail() { $('detail').hidden = true; estOver = null; const bd = $('detailBackdrop'); if (bd) bd.remove(); }

  // ---------- fee estimates (v0.11, a test; only when the pane sends fee data) ----------
  const money2 = (n) => (n == null ? '—' : F.usd(n));
  function renderFeeBar() {
    const bar = $('feeBar');
    document.querySelectorAll('.feecol').forEach((th) => { th.hidden = !fees; });
    bar.hidden = !fees;
    if (!fees) return;
    const st = fees.status || {};
    $('feeState').innerHTML = `${esc(st.text || 'No fee schedule uploaded yet.')}${st.warn ? ` <b class="warn-text">${esc(st.warn)}</b>` : ''}`;
  }
  /** Default estimate (codes and toggles as saved) for the Benefits table. */
  function estimateFor(per, over) {
    const ctx = fees && fees.contexts ? fees.contexts[per] : null;
    return ctx ? { ctx, e: F.estimatePatient(ctx, fees.tables, over || null) } : null;
  }
  function estCells(per) {
    const x = estimateFor(per);
    if (!x) return '<td class="hint">—</td><td class="hint">—</td>';
    const { e } = x;
    if (!e.md && !e.np) return `<td class="hint" colspan="2" title="${esc(e.flags.join('; '))}">${esc(e.flags[0] || 'No estimate')}</td>`;
    const cell = (est) => { const p = F.partialNote(est); return `<td${p ? ` title="${esc(p)}"` : ''}>${esc(money2(est && est.patient))}${p ? ` <span class="hint">${esc(p)}</span>` : ''}</td>`; };
    return `${cell(e.md)}${cell(e.np)}`;
  }
  /** The estimate panel in the patient card: code + units editors, the three toggles, MD and NP math. */
  function renderEstimatePanel(per) {
    const box = $('estPanel');
    if (!box) return;
    const x = estimateFor(per);
    if (!x) {
      box.innerHTML = `<h3>Estimate</h3><p class="hint">${fees.status && fees.status.loaded ? 'No estimate: no verification with benefits for this patient.' : 'Upload the fee schedule on the Benefits tab first.'}</p>`;
      return;
    }
    const { ctx } = x;
    if (!estOver || estOver.per !== per) estOver = { per, codes: (ctx.codes || []).map((c) => ({ ...c })), toggles: { ...ctx.defaults } };
    const T = fees.toggles || F.TOGGLES;
    const sel = (kind, label) => `<label>${label}<select data-etoggle="${kind}">${T[kind].map((o) => `<option${estOver.toggles[kind] === o ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select></label>`;
    box.innerHTML = `<h3>Estimate, not a guarantee <span class="hint">next visit${ctx.appt ? ` ${esc(ctx.appt)}` : ''}${ctx.event ? ` · ${esc(ctx.event)}` : ''}</span></h3>
      <p class="hint">${esc([ctx.telehealth ? `Telehealth: home clinic ${ctx.home || 'unknown'}` : ctx.location, ctx.payer ? `${ctx.payer} (${ctx.payerSource})` : '', ctx.codeKeyword ? `codes from "${ctx.codeKeyword}"${ctx.codeNote ? `: ${ctx.codeNote}` : ''}` : ''].filter(Boolean).join(' · '))}</p>
      <div class="estcodes">${estOver.codes.map((c, i) => `<span class="estcode"><input data-ecode="${i}" value="${esc(c.code)}" aria-label="Code ${i + 1}" maxlength="8"> x <input data-eunits="${i}" type="number" min="1" step="1" value="${esc(c.units)}" aria-label="Units for code ${i + 1}"> <button class="link" data-edel="${i}" aria-label="Remove code ${i + 1}">Remove</button></span>`).join('')}
        <button class="small" data-eadd="1">Add code</button> <button class="link" data-ereset="1">Reset</button></div>
      <div class="filters esttoggles">${sel('deductible', 'Deductible')}${sel('coinsurance', 'Coinsurance')}${sel('copay', 'Copay')}</div>
      <p class="hint">${ctx.feePayer ? `Toggle changes are remembered for ${esc(ctx.feePayer)} (Estimate Defaults sheet).` : 'Toggle changes are for this estimate only: no fee payer column is mapped for this patient, so they are not remembered.'} Code changes are for this estimate only.</p>
      <div id="estResults"></div>`;
    renderEstimateResults();
  }
  function renderEstimateResults() {
    const out = $('estResults');
    if (!out || !estOver) return;
    const x = estimateFor(estOver.per, { codes: estOver.codes.filter((c) => c.code), toggles: estOver.toggles });
    if (!x) { out.innerHTML = ''; return; }
    const { ctx, e } = x;
    const col = (lbl, est) => `<div class="estcol"><h4>${lbl}: ${esc(est && est.patient != null ? F.usd(est.patient) : '—')}</h4>${est && est.none ? `<p class="hint">${esc(est.none)}</p>` : est ? `<ul class="estlines">${F.estimateText(est).map((t) => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}</div>`;
    out.innerHTML = `${e.flags.length ? `<ul class="pc-items">${e.flags.map((f) => `<li class="sev-medium">${esc(f)}</li>`).join('')}</ul>` : ''}
      ${e.basis ? `<p class="hint">${esc(e.basis)}${ctx.oon ? ' (out of network)' : ''}</p>` : ''}
      ${e.md || e.np ? `<div class="estcols">${col('MD', e.md)}${col('NP', e.np)}</div>` : ''}`;
  }
  function refreshEstimate() {
    const p = $('estPanel');
    if (!p || $('detail').hidden) return;
    if (!fees) { p.remove(); return; }
    // New fee data (an upload or saved defaults): toggles follow the saved defaults, codes typed here are kept.
    const keep = estOver && estOver.per === p.dataset.per ? estOver.codes : null;
    estOver = null;
    renderEstimatePanel(p.dataset.per);
    if (keep && estOver) { estOver.codes = keep; renderEstimatePanel(p.dataset.per); }
  }
  $('detail').addEventListener('input', (e) => {
    const c = e.target.closest('[data-ecode],[data-eunits]');
    if (!c || !estOver) return;
    const i = +(c.dataset.ecode ?? c.dataset.eunits);
    if (!estOver.codes[i]) return;
    if (c.dataset.ecode != null) estOver.codes[i].code = c.value.trim().toUpperCase();
    else estOver.codes[i].units = Math.max(1, Number(c.value) || 1);
    renderEstimateResults();
  });
  $('detail').addEventListener('change', (e) => {
    const s = e.target.closest('[data-etoggle]');
    if (!s || !estOver) return;
    estOver.toggles[s.dataset.etoggle] = s.value;
    renderEstimateResults();
    const x = estimateFor(estOver.per);
    // Only the toggle staff changed is remembered; the others shown here may be this patient's defaults (e.g. copay
    // None for a $0 copay) and must not be saved for every patient on the fee payer.
    if (x && x.ctx.feePayer) toPane({ type: 'fee-defaults', feePayer: x.ctx.feePayer, toggles: { [s.dataset.etoggle]: s.value } });
  });
  $('detail').addEventListener('click', (e) => {
    const b = e.target.closest('[data-eadd],[data-edel],[data-ereset]');
    if (!b || !estOver) return;
    const per = estOver.per;
    if (b.dataset.eadd) estOver.codes.push({ code: '', units: 1 });
    else if (b.dataset.edel != null) estOver.codes.splice(+b.dataset.edel, 1);
    else estOver = null;
    renderEstimatePanel(per);
  });
  async function uploadFee(files) {
    const f = files && files[0];
    if (!f) return;
    if (!/\.xls[xm]$/i.test(f.name)) { notice(`${f.name} is not an .xlsx file. The fee schedule must be the Excel workbook.`, false); return; }
    try {
      const [file] = await readFiles([f]);
      $('feeState').textContent = `Sending ${f.name} to Excel…`;
      toPane({ type: 'fee-upload', name: file.name, base64: file.base64 });
    } catch (e) { notice(e.message, false); }
  }

  // ---------- clinic lists (v0.12) ----------
  const cSame = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
  function clinicFiltered() {
    const all = (clinic && clinic.rows) || [];
    const loc = $('cLoc').value, show = $('cShow').value;
    return all.filter((r) => (!loc || cSame(r.Location, loc))
      && (!show || (show === 'flagged' ? !!String(r['Review Items'] || '').trim() : show === 'skipped' ? !!r['Verification Skipped'] && !r['Re-verify Anyway']
        : show === 'asked' ? !!r['Re-verify Anyway'] : show === 'guar' ? !!r['Found Via Guarantor'] : true)));
  }
  function renderClinicLists() {
    const all = (clinic && clinic.rows) || [];
    const locs = [...new Set(all.map((r) => String(r.Location || '').trim()).filter(Boolean))].sort();
    const keep = $('cLoc').value;
    $('cLoc').innerHTML = '<option value="">All locations</option>' + locs.map((l) => `<option>${esc(l)}</option>`).join('');
    $('cLoc').value = locs.includes(keep) ? keep : '';
    const list = clinicFiltered();
    $('cShown').textContent = clinic ? `Showing ${list.length} of ${all.length}${clinic.runDate ? `, run ${clinic.runDate}` : ''}` : 'Upload the NextGen report in the pane first.';
    $('btnCFile').disabled = !$('cLoc').value;
    $('btnCFile').title = $('cLoc').value ? '' : 'Pick a clinic first';
    $('cTbl').querySelector('tbody').innerHTML = list.map((r) => `<tr>
      ${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r.Location)}<td><b>${esc(name(r))}</b></td>${cell(idText(r))}
      <td>${esc(r['Next Event'])}<br><span class="hint">${esc(r['Next Provider'])}</span></td>
      <td>${esc(r['Primary Payer'] || 'None on file')}${r['Choose in NextGen'] ? `<br><span class="hint">NextGen: ${esc(r['Choose in NextGen'])}</span>` : ''}</td>
      <td>${esc(r['Sohar Status'])}${r.Plan ? `<br><span class="hint">${esc(r.Plan)}</span>` : ''}</td>
      <td class="pillc">${pillsHtml(r)}</td>
      <td>${r['Verification Skipped'] ? `${esc(r['Verification Skipped'])}<br><span class="hint">next check ${esc(r['Next Check'])}${r['Re-verify Anyway'] ? ' · re-verify asked' : ''}</span>` : r['Re-verify Anyway'] ? 'Re-verify asked' : ''}</td>
      <td>${esc(r['Benefit Flag'])}${r.Copay !== '' && r.Copay != null ? `<br><span class="hint">Copay ${esc(r.Copay)}</span>` : ''}</td></tr>`).join('')
      || `<tr><td colspan="10" class="hint">${all.length ? 'No one matches these filters.' : 'No patients with an upcoming visit yet.'}</td></tr>`;
  }
  /** Lists for export from the Clinic lists and Payer learning tabs (as shown, after the filters). */
  function listFor(kind) {
    const loc = $('cLoc').value;
    const tag = loc ? ` - ${loc.replace(/[\\/:*?"<>|]/g, '-')}` : '';
    if (kind === 'clinicList') return { name: C.fileName('clinicList').replace(/ (\d{4})/, `${tag} $1`), cols: R.CLINIC_COLS, data: clinicFiltered() };
    if (kind === 'clinicFlags') {
      const keys = new Set(clinicFiltered().map(keyOf));
      return { name: C.fileName('clinicFlags').replace(/ (\d{4})/, `${tag} $1`), cols: R.CLINIC_REVIEW_COLS, data: ((clinic && clinic.review) || []).filter((r) => keys.has(String(r.MRN || C.perKey(r['Per Nbr'])))) };
    }
    if (kind === 'skipped') return { name: C.fileName('skipped').replace(/ (\d{4})/, `${tag} $1`), cols: skipCols, data: skipped.filter((r) => !loc || cSame(r.Location, loc)) };
    if (kind === 'metrics') return { name: C.fileName('metrics'), cols: C.PAYER_METRIC_COLS, data: (learning && learning.metrics) || [] };
    if (kind === 'themes') return { name: C.fileName('themes'), cols: C.GUARANTOR_THEME_COLS, data: (learning && learning.themes) || [] };
    return { name: kind, cols: [], data: [] };
  }

  // ---------- payer learning (v0.12) ----------
  function tableInto(id, cols, data, empty) {
    const t = $(id);
    t.querySelector('thead').innerHTML = `<tr>${cols.map((c) => `<th>${esc(c)}</th>`).join('')}</tr>`;
    t.querySelector('tbody').innerHTML = data.length ? data.map((r) => `<tr>${cols.map((c) => cell(r[c])).join('')}</tr>`).join('') : `<tr><td colspan="${cols.length}" class="hint">${esc(empty)}</td></tr>`;
  }
  function renderLearning() {
    const L = learning || {};
    const t = L.totals;
    $('lUnm').textContent = t ? t.unmappedAppts : 0;
    $('lFail').textContent = t ? t.failedAppts : 0;
    $('lTest').textContent = t && t.testsBack ? `${Math.round((100 * t.testsEligible) / t.testsBack)}%` : '—';
    $('lRec').textContent = t ? t.recurring : 0;
    tableInto('lMetrics', C.PAYER_METRIC_COLS, L.metrics || [], 'No unmapped or failed payers in this report.');
    const sum = C.payerTestSummary(L.tests || []);
    const testRows = [];
    for (const v of sum.values()) for (const c of v.candidates) testRows.push({ name: v.name, payer: `${c.payerName || c.payerId} (${c.payerId})${v.suggestion && v.suggestion.payerId === c.payerId ? ' · suggested' : ''}`, sent: c.sent, back: c.back, eligible: c.eligible });
    $('lTests').querySelector('tbody').innerHTML = testRows.length ? testRows.map((r) => `<tr>${cell(r.name)}${cell(r.payer)}${cell(r.sent)}${cell(r.back)}${cell(r.eligible)}</tr>`).join('') : '<tr><td colspan="5" class="hint">No payer tests yet. Config "Payer tests" = On sends them; the file is on the Upload &amp; save tab.</td></tr>';
    $('lTestNote').textContent = (L.payerTestRows || []).length ? `${L.payerTestRows.length} test check${L.payerTestRows.length === 1 ? '' : 's'} ready to send (Upload & save).` : 'Each test is one Sohar check: the same member ID under a similar Sohar payer.';
    const mp = L.mapping || [];
    $('lMapNote').textContent = mp.length ? `${mp.length} patient${mp.length === 1 ? '' : 's'} on the payer mapping discovery file (member ID sent so Sohar can name the payer). Save it on the Upload & save tab.` : 'No patients need the payer mapping discovery in this build.';
    tableInto('lThemes', C.GUARANTOR_THEME_COLS, L.themes || [], 'No guarantor searches back yet.');
    tableInto('lRows', C.PAYER_LEARNING_COLS, L.rows || [], 'Nothing learned yet.');
  }
  function renderGuarantorCall() {
    const n = guarantor.count || 0;
    $('guarCall').hidden = !n;
    if (n) $('guarText').textContent = `${n} patient${n === 1 ? '' : 's'} without results ${n === 1 ? 'has a guarantor' : 'have guarantors'}. Resubmit with the guarantor as the subscriber?${guarantor.noDob ? ` ${guarantor.noDob} wait for a guarantor DOB in the NextGen report.` : ''}`;
  }

  function queueClass(q) { return { 'Action needed': 'act', Unmatched: 'unm', Pending: 'pen', Clear: 'clr' }[q] || ''; }

  // ---------- tabs ----------
  function showTab(name) {
    tab = name;
    document.querySelectorAll('[data-tab]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === name)));
    document.querySelectorAll('[data-panel]').forEach((p) => { p.hidden = p.dataset.panel !== name; });
    $('sharedFilters').hidden = name === 'upload' || name === 'clinic' || name === 'learning';
    hideDetail();
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
  function save(kind) {
    let name, cols, data, extra = '';
    if (kind === 'verification' || kind === 'discovery') {
      if (!exportsData) { notice('No Sohar import built yet. Upload the NextGen report in the pane first.', false); return; }
      name = C.fileName(kind);
      cols = kind === 'verification' ? C.VERIFICATION_COLS : C.DISCOVERY_COLS; data = exportsData[kind] || [];
    } else if (kind === 'discoveryVerify') {
      if (!discVerify.length) { notice('No discovery patients are waiting for verification. Upload a discovery output with coverage found first.', false); return; }
      name = C.fileName('discoveryVerify'); cols = C.VERIFICATION_COLS; data = discVerify;
    } else if (kind === 'verifyFailed') {
      if (!verFailed.length) { notice('No failed verifications are waiting for discovery.', false); return; }
      name = C.fileName('verifyFailed'); cols = C.DISCOVERY_COLS; data = verFailed;
    } else if (kind === 'mapping') {
      const d = (learning && learning.mapping) || [];
      if (!d.length) { notice('No payer mapping discovery rows in this build.', false); return; }
      name = C.fileName('mapping'); cols = C.MAPPING_DISCOVERY_COLS; data = d;
      extra = ' memberId is added at the end of the Sohar discovery template: check that Sohar accepts it.';
    } else if (kind === 'payerTests') {
      const d = (learning && learning.payerTestRows) || [];
      if (!d.length) { notice('No payer tests to send. Turn them on with Config "Payer tests" = On and rebuild.', false); return; }
      name = C.fileName('payerTests'); cols = C.VERIFICATION_COLS; data = d;
    } else if (kind === 'clinicList' || kind === 'clinicFlags' || kind === 'skipped' || kind === 'metrics' || kind === 'themes') {
      const x = listFor(kind);
      if (!x.data.length) { notice('Nothing in this list.', false); return; }
      name = x.name; cols = x.cols; data = x.data;
    } else if (kind === 'tasks') { name = C.fileName('tasks').replace('Tasks', 'Tasks (open)'); cols = C.TASK_COLS; data = tasks.filter(open); }
    else if (kind === 'allTasks') { name = C.fileName('tasks'); cols = C.TASK_COLS; data = tasks; }
    else if (kind === 'tracker') {
      const out = C.trackerExport(rows, new Date().toLocaleString());
      if (!out.rows.length) { notice('No verified patients yet. The tracker file needs at least one Sohar verification result.', false); return; }
      name = C.fileName('tracker'); cols = C.TRACKER_EXPORT_COLS; data = out.rows;
      if (out.pending.length) extra = ` ${out.pending.length} left out, still waiting on the verification of what discovery found.`;
    }
    else if (kind === 'benefits') { name = C.fileName('benefits'); cols = C.BENEFIT_SHEET_COLS; data = C.benefitsView(rows).rows; }
    else if (kind === 'ir') {
      if (!irRows.length) { notice('No IR Portal cross-check yet. Upload the IR Portal file in the pane.', false); return; }
      name = C.fileName('ir'); cols = C.IR_COLS; data = irRows;
    }
    else { name = C.fileName('worklist'); cols = C.RESULTS_COLS; data = filteredRows(); }
    const csv = C.toCSV(cols, data);
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    notice(`Saved ${name} (${data.length} rows) to your Downloads folder.${extra}`, true);
    // Excel on the web may block downloads in this frame; offer a top-level save tab, then the Export sheet.
    if (inOffice && web) {
      const n = $('notice');
      const tab = document.createElement('button'); tab.className = 'link'; tab.textContent = 'Save in a new tab';
      tab.addEventListener('click', () => saveInTab(name, csv));
      const sheet = document.createElement('button'); sheet.className = 'link'; sheet.textContent = 'put it on the Export sheet';
      sheet.addEventListener('click', () => { notice('Writing to the Export sheet…', true); toPane({ type: 'export', name, cols, rows: data }); });
      n.append(' Nothing there? ', tab, ' or ', sheet, '.');
    }
  }
  function saveInTab(name, csv) {
    const w = window.open(new URL('save.html', location.href).href, '_blank');
    if (!w) { notice('The browser blocked the new tab. Allow pop-ups for Excel, or use the Export sheet.', false); return; }
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

  function notice(text, ok) { const n = $('notice'); n.hidden = false; n.textContent = text; n.className = `notice ${ok ? 'ok' : 'bad'}`; }
  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

  // ---------- wiring ----------
  document.querySelectorAll('[data-tab]').forEach((b) => b.addEventListener('click', () => showTab(b.dataset.tab)));
  ['fSent', 'fLoc', 'fPill', 'fFrom', 'fTo', 'fQueue', 'fBen', 'fBenView', 'fDone', 'fTask', 'msgMode', 'cLoc', 'cShow'].forEach((id) => $(id).addEventListener('change', render));
  document.querySelectorAll('[data-csave]').forEach((b) => b.addEventListener('click', () => save({ list: 'clinicList', flags: 'clinicFlags', skipped: 'skipped' }[b.dataset.csave])));
  document.querySelectorAll('[data-lsave]').forEach((b) => b.addEventListener('click', () => save(b.dataset.lsave)));
  $('btnCFile').addEventListener('click', () => { if ($('cLoc').value) { notice(`Saving the clinic file for ${$('cLoc').value}…`, true); toPane({ type: 'clinic-file', location: $('cLoc').value }); } });
  $('btnGuar').addEventListener('click', () => { notice('Adding guarantors to the discovery file…', true); toPane({ type: 'guarantor-resubmit' }); });
  $('fText').addEventListener('input', render);
  document.querySelectorAll('[data-queue]').forEach((b) => b.addEventListener('click', () => { $('fQueue').value = b.dataset.queue; $('fBen').value = ''; render(); }));
  document.querySelectorAll('[data-benefit]').forEach((b) => b.addEventListener('click', () => { $('fBen').value = b.dataset.benefit; $('fQueue').value = ''; render(); }));
  document.addEventListener('click', (e) => {
    if (e.target.closest('a, input, button')) return;
    const tr = e.target.closest('tr[data-key]');
    if (tr) showDetail(tr.dataset.key);
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('detail').hidden) { hideDetail(); return; } if (e.key === 'Enter') { const tr = e.target.closest && e.target.closest('tr[data-key]'); if (tr) showDetail(tr.dataset.key); } });
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
  $('btnFee').addEventListener('click', () => $('fileFee').click());
  $('fileFee').addEventListener('change', (e) => { uploadFee(e.target.files); e.target.value = ''; });
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
      mk({ 'Next Appt': '10/20/2026', 'Next Time': '9:00 AM', 'Next Event': 'TMS Treatment Unilateral', 'Next Location': 'San Rafael', 'Next Provider': 'Lark, Jordan PMHNP', 'Upcoming Appts': 12, 'Upcoming Detail': '10/20 TMS Treatment Unilateral; 10/21 TMS Treatment Unilateral; +10 more', 'Per Nbr': '100100', MRN: '80100', 'Last Name': 'Rivera', 'First Name': 'Ana', DOB: '03/04/1980', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', Action: 'Eligible. No action.', Plan: 'SYNTH GOLD PPO', Network: 'PR', 'Coverage Level': 'IND', Copay: '40', Coinsurance: '20%', Deductible: '1500', 'Deductible Remaining': '0', 'Deductible Status': 'Met', 'OOP Max': '5000', 'OOP Remaining': '800', 'OOP Status': 'Approaching', 'NextGen Payer': 'UMR', 'Sohar Payer': 'UMR', Match: 'Per Nbr', 'Sohar Ref': 'q-1' }),
      mk({ 'Next Appt': '10/20/2026', 'Next Time': '1:30 PM', 'Next Event': 'Telepsych', 'Next Location': 'San Rafael', 'Upcoming Appts': 1, 'Per Nbr': '100101', 'Last Name': 'Okafor', 'First Name': 'Chidi', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', Action: 'Coverage found. Add it to NextGen, then verify. Get the member ID from the Sohar link.', 'Choose in NextGen': 'Not in Payer Map: Partnership HealthPlan Medi-Cal (SYN01)', 'NextGen Payer': 'Cigna', 'Sohar Payer': 'Partnership HealthPlan Medi-Cal', 'Sohar payerId': 'SYN01', 'Coverage Differs': 'Yes', Match: 'Per Nbr', 'Sohar Link': 'https://example.test/r-1', 'Sohar Ref': 'r-1' }),
      mk({ 'Next Appt': '10/22/2026', 'Next Time': '10:00 AM', 'Next Event': 'Med Management In-Office', 'Next Location': 'Walnut Creek', 'Upcoming Appts': 2, 'Per Nbr': '100102', 'Last Name': 'Lindqvist', 'First Name': 'Sven', 'Sent As': 'Verification', 'Sohar Status': 'error.member.id', Action: 'Member ID rejected. Check the card, fix in NextGen, re-verify.', 'NextGen Payer': 'Kaiser Foundation Northern California', 'NextGen Member ID': 'SYN555', Match: 'Per Nbr', 'Sohar Ref': 'q-2' }),
      mk({ 'Next Appt': '10/23/2026', 'Next Time': '11:00 AM', 'Next Event': 'Telepsych', 'Next Location': 'Walnut Creek', 'Upcoming Appts': 4, 'Per Nbr': '100103', 'Last Name': 'Mensah', 'First Name': 'Afia', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', Action: 'Eligible. No action.', Copay: '0', Coinsurance: '0%', Deductible: '0', 'Deductible Remaining': '0', 'Deductible Status': 'None', 'OOP Max': '3000', 'OOP Remaining': '0', 'OOP Status': 'Met', 'NextGen Payer': 'QMB Noridian Northern CA', 'Sohar Payer': 'Medicare of California', Plan: 'MEDICARE PART B', Match: 'Per Nbr', 'Sohar Ref': 'q-3' }),
      mk({ 'Sent As': 'Verification', 'Sohar Status': 'error.member.id', Action: 'Not matched to a patient in the Import report. Member ID rejected. Check the card, fix in NextGen, re-verify.', Match: 'No match', 'Sohar Ref': 'q-9' }),
    ].map((r) => ({ ...r, 'Benefit Flag': C.benefitFlag(r), 'Work Queue': C.workQueue(r) }));
    const demoPills = { 100100: 'NP-TMS', 100101: 'Medi-Cal; New Medi-Cal; Re-verify', 100102: 'Kaiser/VA MM no auth; Kaiser/VA MM', 100103: 'QMB' };
    for (const r of res) r.Pills = demoPills[r['Per Nbr']] || '';
    // ?fees=1: fee estimates with the made-up demo fee schedule (fees.js), as the pane would send them.
    let demoFees = null;
    if (F && /[?&]fees=1\b/.test(location.search)) {
      const s = F.demoSchedule();
      const prep = F.prepareEstimates({ results: res, scheduleRows: s.rows, tabSummary: s.tabs, appts: [], teleKeywords: ['Tele'], cfgDefaults: F.TOGGLE_DEFAULTS, when: '' });
      demoFees = { on: true, status: F.feeStatus(F.feeMetaRows('Demo fee schedule (made-up rates).xlsx', new Date().toISOString(), s), s.rows.length, new Date()), contexts: prep.contexts, tables: prep.tables, toggles: F.TOGGLES };
    }
    // Clinic lists and payer learning, as the pane would send them (synthetic).
    const crow = (r, o) => Object.assign(Object.fromEntries(R.CLINIC_COLS.map((c) => [c, r[c] == null ? '' : r[c]])), { Location: r['Next Location'], 'Primary Payer': r['NextGen Payer'] }, o || {});
    const demoClinic = { runDate: '10/20/2026', rows: [
      crow(res[0], { 'Verification Skipped': 'Verified eligible 10/18/2026', 'Next Check': '10/25/2026' }),
      crow(res[1], { 'Review Items': 'Medi-Cal is not accepted' }),
      crow(res[2], { 'Verification Skipped': 'Verified eligible 10/19/2026', 'Next Check': '10/26/2026', 'Re-verify Anyway': 'Yes' }),
      crow(res[3], { 'Found Via Guarantor': 'Yes', Pills: 'QMB; Via guarantor' }),
    ], review: [] };
    const demoSkipped = [{ Location: 'San Rafael', 'Next Appt': '10/20/2026', MRN: '80100', 'Per Nbr': '100100', 'Last Name': 'Rivera', 'First Name': 'Ana', Verified: '10/18/2026', 'Next Check': '10/25/2026', 'Other Insurance Sent': '' }];
    const metric = (name, conf, um, ua, fp, fa, ts, tb, te) => ({ 'NextGen Payer': name, Confidence: conf, 'Unmapped patients': String(um), 'Unmapped appts': String(ua), 'Failed verification patients': String(fp), 'Failed verification appts': String(fa),
      'Most common failure': fp ? 'error.payer (1)' : '', 'Payer tests sent': String(ts), 'Payer tests back': String(tb), 'Payer tests eligible': String(te), 'Test success rate': tb ? `${Math.round(100 * te / tb)}%` : '', 'Best test payer': te ? 'Blue Shield of California (94036)' : '', 'Days with issues': '4', Recurring: 'Yes' });
    const demoLearning = { totals: { unmappedPatients: 3, unmappedAppts: 14, failedPatients: 1, failedAppts: 2, testsSent: 6, testsBack: 4, testsEligible: 2, recurring: 1 },
      metrics: [metric('Blue Shield CA Blue Card', 'Auto-Low', 2, 12, 0, 0, 6, 4, 2), metric('Odd Local Plan', 'Unmapped', 1, 2, 1, 2, 0, 0, 0)],
      tests: [['94036', 'Blue Shield of California', 'complete.eligible'], ['94036', 'Blue Shield of California', 'complete.eligible'], ['47198', 'Blue Cross of California', 'complete.ineligible'], ['47198', 'Blue Cross of California', 'error.payer']]
        .map(([id, nm, st], i) => ({ 'Test Id': `8${i}-T1`, MRN: `8${i}`, 'NextGen Payer': 'Blue Shield CA Blue Card', 'Candidate payerId': id, 'Candidate Payer': nm, Status: st })),
      themes: [{ Theme: 'Guarantor only', 'NextGen Payer': 'None on file', 'Found for patient': '—', 'Found for guarantor': 'Blue Shield of California', Patients: '1' }],
      rows: [], mapping: [{ id: '80101-M' }], payerTestRows: [] };
    load({ results: res, tasks: C.buildTasks(res), exports: { verification: [], discovery: [] }, asOf: '2026-10-13T00:00:00.000Z', fees: demoFees,
      clinic: demoClinic, skipped: demoSkipped, skipCols: ['Location', 'Next Appt', 'MRN', 'Per Nbr', 'Last Name', 'First Name', 'Verified', 'Next Check', 'Other Insurance Sent'], learning: demoLearning, guarantor: { count: 2, noDob: 1 } });
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
