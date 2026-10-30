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
  // v0.20 patient record: Sohar Flow results, Payer Outcomes ledger rows (timeline) and report profiles (Per Nbr ->
  // pronouns, gender, guarantor LASFIR and checks), all from the pane.
  const REC = window.SoharRecord;
  let flow = [], outcomes = [], profiles = {};
  // Per Nbr -> upcoming visits with the report's Pre-Check Complete and the checklist; when each source last came in.
  let visits = {}, lastUpdated = null, admins = []; // admins: Config "Admin names"
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
    flow = m.flow || []; outcomes = m.outcomes || []; profiles = m.profiles || {};
    visits = m.visits || {}; lastUpdated = m.lastUpdated || null; admins = m.admins || [];
    fees = F && m.fees && m.fees.on ? m.fees : null;
    renderFeeBar(); renderUpdated();
    const people = REC.byPatient(rows).length;
    const gFlags = Object.values(profiles).filter((p) => (p.checks || []).some((c) => /^(high|medium)$/.test(c.severity))).length;
    $('dStatus').textContent = rows.length
      ? `${people} patient${people === 1 ? '' : 's'} (${rows.length} Sohar results), ${tasks.filter(open).length} open NextGen tasks${asOf ? `, appointments from ${C.fmtDate(asOf)}` : ''}${flow.length ? `, ${flow.length} Sohar Flow visit${flow.length === 1 ? '' : 's'}` : ''}${gFlags ? `, ${gFlags} guarantor / subscriber check${gFlags === 1 ? '' : 's'}` : ''}.`
      : 'No Sohar results yet. Upload the verification and discovery outputs on the Upload & save tab.';
    $('dStatus').className = `status ${rows.length ? 'ok' : 'bad'}`;
    // Tiles count patients (v0.20: one record per MRN): a patient counts in each queue any of their results is in.
    const q = (name) => REC.byPatient(rows).filter((g) => g.queues.includes(name)).length;
    $('tAct').textContent = q('Action needed'); $('tUnm').textContent = q('Unmatched');
    $('tPen').textContent = q('Pending'); $('tClr').textContent = q('Clear');
    const leads = REC.byPatient(rows).map((g) => g.row); // benefit tiles: each patient's most recent result
    $('tDedMet').textContent = leads.filter(BEN.dedMet).length; $('tDedAppr').textContent = leads.filter(BEN.dedAppr).length;
    $('tOopMet').textContent = leads.filter(BEN.oopMet).length; $('tOopAppr').textContent = leads.filter(BEN.oopAppr).length;
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

  /** "Last updated" pill: the newest source time, colored by freshness; the title lists each source. */
  function renderUpdated() {
    const el = $('dUpdated');
    const f = REC.freshness(lastUpdated || {}, new Date());
    if (f.level === 'none') { el.hidden = true; return; }
    const when = (t) => new Date(t).toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
    // Not fresh: name the oldest of the report and the Sohar results, and how old it is.
    const old = f.parts.filter((p) => p.key !== 'flow').sort((a, b) => (a.time ?? -Infinity) - (b.time ?? -Infinity))[0];
    const days = old && old.time != null ? Math.floor((Date.now() - old.time) / 864e5) : null;
    const word = f.level === 'fresh' ? '' : ` · ${old.label} ${days == null ? 'never uploaded' : `${days} day${days === 1 ? '' : 's'} old`}${f.level === 'stale' ? ', upload again' : ''}`;
    el.className = `pill upd ${{ fresh: 'p-blue', aging: 'p-amber', stale: 'p-red' }[f.level]}`;
    el.textContent = `Last updated ${when(f.newest)}${word}`;
    const detail = f.parts.map((p) => `${p.label}: ${p.time == null ? 'never' : when(p.time)}`).join('\n');
    el.title = detail;
    el.setAttribute('aria-label', `Last updated ${when(f.newest)}${word}. ${detail.replace(/\n/g, '. ')}`);
    el.hidden = false;
  }
  /** The next visit's pre-check state for the lists: 'done', 'open' or '' (no report visit). */
  function precheckOf(r) {
    const v = (visits[String(r['Per Nbr'] || '')] || [])[0];
    if (!v) return '';
    const pc = v.items.find((i) => i.id === 'precheck');
    if (pc && pc.done) return 'done';
    // A report without the Pre-Check column says nothing: no pill until someone ticks it.
    return v.hasPrecheck || (pc && pc.source) ? 'open' : '';
  }
  const precheckPill = (r) => ({ done: ' <span class="pill p-blue">Pre-check ✓</span>', open: ' <span class="pill p-amber">Pre-check open</span>' }[precheckOf(r)] || '');
  /** A red pill when the NextGen guarantor and a Sohar subscriber are different people (record.guarantorCheck, high). */
  // Amber when the patient is the guarantor but Sohar's subscriber is someone else (check the policy's insured).
  const guarFlag = (r) => {
    const cs = (profiles[String(r['Per Nbr'] || '')] || {}).checks || [];
    if (cs.some((c) => c.severity === 'high')) return ' <span class="pill p-red">Guarantor ≠ subscriber</span>';
    return cs.some((c) => c.severity === 'medium') ? ' <span class="pill p-amber">Check insured</span>' : '';
  };

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
    if (tv && !['LASFIR', 'MRN', 'Per Nbr', 'NextGen Payer', 'Sohar Payer', 'Plan', 'Choose in NextGen', 'Task', 'Details'].some((k) => String(C.cellOf(r, k) || '').toLowerCase().includes(tv))) return false;
    return true;
  }
  function filteredRows() {
    const qv = $('fQueue').value, bv = $('fBen').value;
    return rows.filter((r) => shared(r) && (!qv || r['Work Queue'] === qv) && (!bv || BEN[bv](r)));
  }
  /** One entry per patient (record.byPatient) when any of the patient's results passes `pass`. */
  const patientsWhere = (pass) => REC.byPatient(rows).filter((g) => g.rows.some(pass));
  const moreText = (g) => (g.resultCount > 1 ? ` · ${g.resultCount} results` : '');
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
  const name = (r) => C.tagOf(r) || '(not matched)'; // LASFIR: names are not shown (v0.19)
  function link(u) { return /^https:\/\//.test(u || '') ? `<a href="${esc(u)}" target="_blank" rel="noopener">Sohar</a>` : ''; }

  function renderPrep() {
    // One row per patient: the most recent result stands for them (v0.20).
    const groups = patientsWhere(shared), gOf = new Map(groups.map((g) => [g.row, g]));
    const rs = groups.map((g) => g.row);
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
            ${cell(r['Next Time'])}<td><b>${esc(name(r))}</b><br><span class="hint">${esc(idText(r))}${esc(moreText(gOf.get(r)))}${r['Searched As'] && r['Searched As'] !== 'Patient' ? ` · searched as ${esc(r['Searched As'].toLowerCase())}` : ''}</span>${pillsOf(r).length || guarFlag(r) || precheckPill(r) ? `<div class="pills">${pillsHtml(r)}${guarFlag(r)}${precheckPill(r)}</div>` : ''}</td>
            <td>${esc(r['Next Event'])}<br><span class="hint">${esc([r['Next Location'], r['Next Provider']].filter(Boolean).join(' · '))}</span></td>
            <td><span class="pill ${r['Sent As'] === 'Discovery' ? 'd' : 'v'}">${esc(r['Sent As'])}</span><br><span class="hint">${esc(r['Sohar Status'])}</span></td>
            <td>${esc(r['Benefit Flag'] || (r['Sent As'] === 'Verification' ? '' : ''))}${r.Copay !== '' && r.Copay != null ? `<br><span class="hint">Copay ${esc(r.Copay)}${r.Coinsurance ? ` · ${esc(r.Coinsurance)}` : ''}</span>` : ''}</td>
            <td class="wide">${n ? `<span class="pill t">${n} task${n > 1 ? 's' : ''}</span> ` : ''}${esc(r.Action)}${r['Choose in NextGen'] ? `<br><span class="hint">NextGen: ${esc(r['Choose in NextGen'])}</span>` : ''}</td>
          </tr>`;
        }).join('')}</tbody></table></div>`;
    }).join('') : '<p class="hint">No patients match these filters.</p>';
  }

  function renderWorklist() {
    const qv = $('fQueue').value, bv = $('fBen').value;
    const groups = patientsWhere((r) => shared(r) && (!qv || r['Work Queue'] === qv) && (!bv || BEN[bv](r))), gOf = new Map(groups.map((g) => [g.row, g]));
    list = groups.map((g) => g.row);
    $('shown').textContent = `Showing ${list.length} of ${REC.byPatient(rows).length} patients (${rows.length} Sohar results; the most recent result stands for each patient)`;
    const st = (s) => ({ Met: 'met', Approaching: 'appr' }[s] || '');
    $('tbl').querySelector('tbody').innerHTML = list.map((r) => `<tr class="q-${queueClass(r['Work Queue'])}" data-key="${esc(keyOf(r) + '|' + r['Sohar Ref'])}" tabindex="0">
      ${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r['Next Location'])}${cell(name(r))}${cell(idText(r))}<td class="pillc">${pillsHtml(r)}${guarFlag(r)}${precheckPill(r)}</td>${cell(r['Sent As'] + moreText(gOf.get(r)))}
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
      ${cell(t.Task, 'mid')}${cell(t.Details, 'wide')}${cell(t['Clinic Note'], 'mid cnote')}${cell(t['Choose in NextGen'], 'mid')}<td>${link(t['Sohar Link'])}</td></tr>`).join('');
  }

  function renderMessage() {
    const mode = $('msgMode').value;
    $('msg').value = mode === 'tasks' ? C.taskMessage(filteredTasks().filter(open), { asOf })
      : mode === 'ir' ? (irRows.length ? C.irMessage(irRows) : 'No IR Portal cross-check yet. Upload the IR Portal file in step 6 of the pane.\n')
      : C.adminMessage(rows.filter(shared), { asOf, mode });
  }

  // ---------- patient record (v0.20: one record per MRN) ----------
  // Top down: the Sohar Flow cost estimates, then the most recent result, the upcoming visits and the financial progress,
  // then the guarantor check, the calculator, NextGen tasks, and every earlier verification or discovery, newest first.
  let detailKey = '';
  const keyOfDetail = (dk) => { const r = rows.find((x) => `${keyOf(x)}|${x['Sohar Ref']}` === dk); return r ? (keyOf(r) || `ref:${r['Sohar Ref']}`) : ''; };
  function showDetail(dk) {
    const k = keyOfDetail(dk);
    if (!k) return;
    detailKey = k;
    const rec = k.startsWith('ref:') ? null : REC.patientRecord(rows, k, { outcomes, flow, asOf });
    const r = rec ? rec.lead : rows.find((x) => `ref:${x['Sohar Ref']}` === k);
    if (!r) return;
    const its = tasks.filter((t) => (rec ? keyOf(t) === k : t['Sohar Ref'] === r['Sohar Ref']));
    const prof = profiles[String(r['Per Nbr'] || '')] || null;
    const dl = (pairs) => { const p = pairs.filter(([, v]) => v !== '' && v != null); return p.length ? `<dl>${p.map(([t, v, html]) => `<dt>${esc(t)}</dt><dd>${html ? v : esc(v)}</dd>`).join('')}</dl>` : '<p class="hint">None</p>'; };
    const who = [prof && prof.pronouns ? `Pronouns: ${prof.pronouns}` : '', prof && prof.genderIdentity ? `Gender identity: ${prof.genderIdentity}` : '', prof && prof.currentGender ? `Current gender: ${prof.currentGender}` : ''].filter(Boolean);
    const box = $('detail');
    box.innerHTML = `<div class="dhead"><h2>${esc(name(r))}</h2>
        <span class="hint">${esc(idText(r))}${r['Per Nbr'] && r.MRN ? ` · Per Nbr ${esc(r['Per Nbr'])}` : ''}${rec && rec.results.length > 1 ? ` · ${rec.results.length} Sohar results` : ''}</span>
        <button class="small" id="btnHide">Close</button></div>
      ${who.length ? `<p class="hint who">${who.map(esc).join(' · ')}</p>` : ''}
      ${rec ? flowHtml(rec.flow) : ''}
      <div class="dgrid rtop">
        <div class="dgroup lead"><h3>Most recent result</h3>
          <p><span class="pill ${r['Sent As'] === 'Discovery' ? 'd' : 'v'}">${esc(r['Sent As'])}</span> <b>${esc(r['Sohar Status'])}</b> <span class="hint">${esc(r['Sohar Date'] ? fmtWhen({ time: C.soharTime(r['Sohar Date']), date: r['Sohar Date'] }) : '')}</span></p>
          ${dl([['Action', r.Action], ['Choose in NextGen', r['Choose in NextGen']], ['Sohar Payer', r['Sohar Payer']], ['NextGen Payer', r['NextGen Payer']], ['Plan', r.Plan], ['Network', r.Network],
            ['Coverage', [r['Coverage Start'], r['Coverage End']].some(Boolean) ? `${r['Coverage Start'] || '?'} to ${r['Coverage End'] || 'open'}` : ''], ['Member ID', r['Sohar Member ID'] || r['NextGen Member ID']],
            ['Searched as', r['Searched As'] && r['Searched As'] !== 'Patient' ? r['Searched As'] : ''], ['Match', r.Match], ['Sohar', link(r['Sohar Link']), true]])}
          ${pillsOf(r).length ? `<div class="pills">${pillsHtml(r)}</div>` : ''}</div>
        <div class="dgroup" id="visitBox"><h3>Upcoming visits and pre-check</h3>${visitsHtml(r, rec)}
          ${r['Next Provider'] ? `<p class="hint">Next with ${esc(r['Next Provider'])}</p>` : ''}</div>
        <div class="dgroup"><h3>Financial progress</h3>${financialHtml(rec && rec.financial, r)}</div>
      </div>
      <div class="dgrid">
        ${guarantorHtml(prof)}
        <div class="dgroup"><h3>NextGen tasks</h3>${its.length ? `<ul class="tasklist">${its.map((t) => `<li class="${open(t) ? '' : 'done'}"><label><input type="checkbox" data-task="${esc(t['Task Key'])}" ${open(t) ? '' : 'checked'}> <b>${esc(t.Task)}</b></label><br>${esc(t.Details)}${t['Clinic Note'] ? `<br><span class="cnote">Clinic: ${esc(t['Clinic Note'])}</span>` : ''}</li>`).join('')}</ul>` : '<p class="hint">None</p>'}</div>
        ${fees ? `<div class="dgroup est" id="estPanel" data-per="${esc(r['Per Nbr'])}"></div>` : ''}
      </div>
      <div class="dgroup timeline"><h3>Earlier verifications and discoveries <span class="hint">newest first</span></h3>${timelineHtml(rec ? rec.timeline : [])}</div>`;
    box.hidden = false;
    if (fees) renderEstimatePanel(r['Per Nbr'], r);
    box.querySelectorAll('.meter').forEach((m) => { m.firstElementChild.style.width = `${Math.round(Math.min(1, +m.dataset.pct) * 100)}%`; });
    $('btnHide').addEventListener('click', hideDetail);
    if (!$('detailBackdrop')){ const bd = document.createElement('div'); bd.className = 'backdrop'; bd.id = 'detailBackdrop'; bd.addEventListener('click', hideDetail); document.body.append(bd); }
    $('btnHide').focus();
  }
  // the detail card floats over the dashboard; everything behind it is blurred
  function hideDetail() { $('detail').hidden = true; estOver = null; detailKey = ''; const bd = $('detailBackdrop'); if (bd) bd.remove(); }

  /**
   * Upcoming visits from the report, each with the NextGen Pre-Check Complete value, the checklist and notes admins add
   * here (saved on the hidden Pre-Check sheet with their name). Without report visits: the visit summary from the Sohar
   * result. The admin's name comes from Config "Admin names" and is remembered on this computer.
   */
  const NAME_KEY = 'sohar.precheck.by';
  let adminName = (() => { try { return localStorage.getItem(NAME_KEY) || ''; } catch (e) { return ''; } })();
  function visitsHtml(r, rec) {
    const vs = visits[String(r['Per Nbr'] || '')] || [];
    if (!vs.length) return rec && rec.upcoming.length ? `<ul class="visits">${rec.upcoming.map((v) => `<li>${esc(v)}</li>`).join('')}</ul>` : '<p class="hint">None scheduled</p>';
    const names = admins.includes(adminName) || !adminName ? admins : [adminName, ...admins];
    const picker = `<label class="inline pcby">Your name
      ${names.length ? `<select id="pcBy"><option value="">Pick your name</option>${names.map((n) => `<option${n === adminName ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select>`
        : `<input id="pcBy" value="${esc(adminName)}" placeholder="Your name" maxlength="60">`}</label>
      ${names.length ? '' : '<span class="hint">Add the admin names to Config "Admin names" (separated by ;) to pick from a list.</span>'}`;
    return `${picker}<ul class="visits pc">${vs.map((v, i) => `<li>
      <b>${esc([v.date, v.time].filter(Boolean).join(' '))}</b> ${esc([v.event, v.loc].filter(Boolean).join(' · '))}${v.provider ? ` <span class="hint">${esc(v.provider)}</span>` : ''}
      <br><span class="hint">${esc(v.lasfir || name(r))} · NextGen pre-check: ${esc(v.precheck || 'not on the report')}</span>
      <div class="pcheck">${v.items.map((it) => `<label class="inline"><input type="checkbox" data-pc="${i}" data-pc-item="${esc(it.id)}"${it.done ? ' checked' : ''}> ${esc(it.label)}
        <span class="hint" data-pc-note="${i}|${esc(it.id)}">${esc(pcNote(it))}</span></label>`).join('')}</div>
      <div class="pnotes">${(v.notes || []).map((n) => `<p class="pnote">${esc(n.note)} <span class="hint">${esc([n.by, n.when].filter(Boolean).join(', '))}</span></p>`).join('')}
        <div class="pnote-add"><input data-pn="${i}" maxlength="${REC.NOTE_MAX}" placeholder="Add a note for this visit" aria-label="Note for the ${esc(v.date)} visit"> <button class="small" data-pn-add="${i}">Add note</button></div></div></li>`).join('')}</ul>
      <p class="hint">Ticks and notes are saved in this workbook with your name. The add-in does not change NextGen: "Updated in NextGen" means you made the change there.</p>`;
  }
  const pcNote = (it) => (it.source === 'NextGen report' ? '(from the report)' : it.source === 'Marked here' ? `(${[it.by, it.when].filter(Boolean).join(', ')})` : '');
  const stamp = () => { const now = new Date(); return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()} ${now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`; };
  const detailVisit = (i) => {
    const r = rows.find((x) => keyOf(x) === detailKey) || rows.find((x) => `ref:${x['Sohar Ref']}` === detailKey);
    return r ? { r, v: (visits[String(r['Per Nbr'] || '')] || [])[+i] } : {};
  };
  /** The admin's name, or a notice and null when none is picked. */
  function byOrAsk() {
    if (adminName) return adminName;
    notice('Pick your name first (top of Upcoming visits). It is saved with each tick and note.', false);
    const el = $('pcBy'); if (el) el.focus();
    return null;
  }
  $('detail').addEventListener('change', (e) => {
    if (e.target.id === 'pcBy') {
      adminName = e.target.value.trim();
      try { localStorage.setItem(NAME_KEY, adminName); } catch (err) { /* storage blocked: remembered for this window only */ }
      return;
    }
    const c = e.target.closest('input[data-pc]');
    if (!c || !detailKey) return;
    const { r, v } = detailVisit(c.dataset.pc);
    const it = v && v.items.find((x) => x.id === c.dataset.pcItem);
    if (!it) return;
    const by = byOrAsk();
    if (!by) { c.checked = !c.checked; return; }
    Object.assign(it, { done: c.checked, source: 'Marked here', when: stamp(), by });
    const note = $('detail').querySelector(`[data-pc-note="${c.dataset.pc}|${it.id}"]`);
    if (note) note.textContent = pcNote(it);
    toPane({ type: 'precheck', key: v.key, lasfir: v.lasfir || name(r), date: v.date, time: v.time, item: it.id, done: c.checked, by });
    renderPrep(); renderWorklist(); // the list pills follow
  });
  $('detail').addEventListener('click', (e) => {
    const b = e.target.closest('[data-pn-add]');
    if (!b || !detailKey) return;
    const { r, v } = detailVisit(b.dataset.pnAdd);
    const input = $('detail').querySelector(`input[data-pn="${b.dataset.pnAdd}"]`);
    const text = input ? input.value.replace(/\s+/g, ' ').trim() : '';
    if (!v || !text) return;
    const by = byOrAsk();
    if (!by) return;
    const n = { note: text.slice(0, REC.NOTE_MAX), by, when: stamp() };
    v.notes = [n, ...(v.notes || [])];
    input.value = '';
    input.closest('.pnotes').insertAdjacentHTML('afterbegin', `<p class="pnote">${esc(n.note)} <span class="hint">${esc(`${n.by}, ${n.when}`)}</span></p>`);
    toPane({ type: 'precheck-note', key: v.key, lasfir: v.lasfir || name(r), date: v.date, time: v.time, note: n.note, by });
  });
  $('detail').addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const i = e.target.closest('input[data-pn]');
    if (i) { e.preventDefault(); const b = $('detail').querySelector(`[data-pn-add="${i.dataset.pn}"]`); if (b) b.click(); }
  });
  // Free-text name (no Config list yet): remembered as it is typed.
  $('detail').addEventListener('input', (e) => {
    if (e.target.id !== 'pcBy') return;
    adminName = e.target.value.trim();
    try { localStorage.setItem(NAME_KEY, adminName); } catch (err) { /* storage blocked */ }
  });
  /** Sohar Flow cost estimates, one card per visit (soonest first). Sandbox files say so. */
  function flowHtml(list) {
    if (!list || !list.length) return '';
    const card = (f) => {
      const pt = REC.flowAmount(f['Sohar Est Patient']), plan = REC.flowAmount(f['Sohar Est Plan Pays']);
      const done = /^complete/i.test(f['Run Status']);
      return `<div class="fcard">
        <div class="fhead"><b>${esc([f['Appt Date'], f.Time].filter(Boolean).join(' '))}</b> <span class="hint">${esc([f.Event, f.Location, f.Provider].filter(Boolean).join(' · '))}</span></div>
        <div class="famt">${pt ? `<span class="big">${esc(pt)}</span> <span class="hint">patient${plan ? ` · plan pays ${esc(plan)}` : ''}</span>` : `<span class="hint">${esc(done ? f['Estimate Note'] || 'No estimate' : `Flow ${String(f['Run Status'] || 'not started').toLowerCase().replace(/_/g, ' ')}`)}</span>`}</div>
        <p class="hint">${esc([f['Est Codes'] ? `Codes ${f['Est Codes']}` : '', f['Rate Level'], f['Fee Tab'], f['Estimate Policy'] ? `${f['Estimate Policy']} policy` : '', f.Network, f['Sohar Payer']].filter(Boolean).join(' · '))}</p>
        ${f['Policy vs NextGen'] ? `<p class="fpol">${esc(f['Policy vs NextGen'])}</p>` : ''}
        ${pt && f['Estimate Note'] ? `<p class="hint">${esc(f['Estimate Note'])}</p>` : ''}
        ${f['Run Error'] ? `<p class="warn-text">${esc(f['Run Error'])}</p>` : ''}
      </div>`;
    };
    const sandbox = list.some(REC.isSandbox);
    return `<div class="dgroup flow"><h3>Sohar Flow cost estimate${list.length > 1 ? 's' : ''} ${sandbox ? '<span class="pill p-amber">Sandbox: test data</span>' : ''} <span class="hint">estimates, not a guarantee</span></h3>
      <div class="fcards">${list.map(card).join('')}</div></div>`;
  }
  /** Deductible and OOP max: spent so far of the total, with a bar (core.benefitsView row). */
  function financialHtml(b, r) {
    if (!b) return `<p class="hint">${r['Sent As'] === 'Discovery' ? 'No benefits yet: discovery finds coverage; the verification of it brings the amounts.' : 'No benefit amounts from Sohar.'}</p>`;
    const bar = (label, met, total, pct, status, rem) => {
      if (status === 'None') return `<div class="frow"><span>${label}</span> <span class="hint">None</span></div>`;
      if (pct === '' || pct == null) return `<div class="frow"><span>${label}</span> <span class="hint">${esc(status || 'Unknown')}</span></div>`;
      const cls = status === 'Met' ? 'met' : status === 'Approaching' ? 'appr' : '';
      return `<div class="frow ${cls}"><span>${label}</span> <span class="meter wide" data-pct="${pct}"><span></span></span> <b>${Math.round(pct * 100)}%</b>
        <br><span class="hint">${esc(usd(met))} of ${esc(usd(total))} met${rem !== '' && rem != null ? ` · ${esc(usd(rem))} left` : ''}${status ? ` · ${esc(status)}` : ''}</span></div>`;
    };
    return `${bar('Deductible', b['Deductible Met'], b.Deductible, b['Deductible % Met'], b['Deductible Status'], b['Deductible Remaining'])}
      ${bar('OOP max', b['OOP Met'], b['OOP Max'], b['OOP % Met'], b['OOP Status'], b['OOP Remaining'])}
      <p class="hint">${esc([b.Copay !== '' && b.Copay != null ? `Copay ${usd(b.Copay)}` : '', b.Coinsurance ? `Coinsurance ${b.Coinsurance}` : '', b.Plan].filter(Boolean).join(' · '))}</p>
      ${b.Flag ? `<p class="${b.Flag === 'OOP max met' ? 'ok-text' : 'warn-text'}">${esc(b.Flag)}</p>` : ''}`;
  }
  /** NextGen guarantor (LASFIR) and the guarantor vs Sohar subscriber check for each result. */
  function guarantorHtml(prof) {
    if (!prof) return '';
    const checks = prof.checks || [];
    if (!prof.guarantor && !checks.length) return '';
    const sev = (s) => (s === 'high' ? 'sev-high' : s === 'medium' ? 'sev-medium' : s === 'low' ? 'sev-low' : '');
    return `<div class="dgroup"><h3>Guarantor</h3>
      ${dl2([['NextGen guarantor', prof.guarantor || 'Not on the report'], ['Guarantor DOB', prof.guarantorDob], ['Previous last name', prof.prevLastName]])}
      ${checks.length ? `<ul class="pc-items gchecks">${checks.map((c) => `<li class="${sev(c.severity)}"><b>${esc(c.status)}</b> <span class="hint">${esc([c.sentAs, c.payer, c.date].filter(Boolean).join(' · '))}</span><br>${esc(c.text)}</li>`).join('')}</ul>` : ''}</div>`;
  }
  const dl2 = (pairs) => { const p = pairs.filter(([, v]) => v !== '' && v != null); return p.length ? `<dl>${p.map(([t, v]) => `<dt>${esc(t)}</dt><dd>${esc(v)}</dd>`).join('')}</dl>` : ''; };
  /** Earlier results and ledger rows, newest first. */
  function timelineHtml(list) {
    if (!list.length) return '<p class="hint">No earlier Sohar checks for this patient.</p>';
    return `<ol class="tline">${list.map((e) => `<li>
      <span class="tdate">${esc(e.date ? fmtWhen(e) : 'Date unknown')}</span>
      <span class="pill ${e.kind === 'Discovery' ? 'd' : 'v'}">${esc(e.kind)}</span> <b>${esc(e.status || '—')}</b>${e.payer ? ` · ${esc(e.payer)}` : ''}${e.plan ? ` · ${esc(e.plan)}` : ''}
      ${e.source === 'ledger' ? ' <span class="hint">(from the payer outcomes history; details not kept)</span>' : ''}
      ${[e.action, e.coverage ? `Coverage ${e.coverage}` : '', e.benefits, e.searched && e.searched !== 'Patient' ? `Searched as ${e.searched.toLowerCase()}` : ''].filter(Boolean).length
        ? `<br><span class="hint">${esc([e.action, e.coverage ? `Coverage ${e.coverage}` : '', e.benefits, e.searched && e.searched !== 'Patient' ? `Searched as ${e.searched.toLowerCase()}` : ''].filter(Boolean).join(' · '))}</span>` : ''}
      ${e.link ? ` ${link(e.link)}` : ''}</li>`).join('')}</ol>`;
  }
  const fmtWhen = (e) => {
    if (!isFinite(e.time)) return e.date;
    const d = new Date(e.time);
    return d.toLocaleString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
  };

  // ---------- patient responsibility calculator (v0.20; the West Coast calculator's math, pre-filled from Sohar) ----------
  // Shown when the pane sends fee data (Config "Fee estimates" On). estOver: the open panel's inputs and codes.
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
  /** The calculator panel in the patient record: plan, benefits and codes pre-filled, all editable; MD and NP results. */
  function renderEstimatePanel(per, row) {
    const box = $('estPanel');
    if (!box) return;
    const ctx = fees && fees.contexts ? fees.contexts[per] : null;
    if (!ctx) {
      box.innerHTML = `<h3>Patient responsibility calculator</h3><p class="hint">${fees.status && fees.status.loaded ? 'No estimate: no verification with benefits for this patient.' : 'Drop the fee schedule or the West Coast Patient Responsibility Calculator on the Benefits tab first.'}</p>`;
      return;
    }
    if (!estOver || estOver.per !== per) estOver = { per, row: row || estOver?.row || {}, inputs: F.calcInputs(ctx, row), codes: (ctx.codes || []).map((c) => ({ ...c, coins: '' })) };
    const b = estOver.inputs;
    const opt = (list, v) => list.map(([k, l]) => `<option value="${esc(k)}"${k === v ? ' selected' : ''}>${esc(l)}</option>`).join('');
    const numIn = (k, label) => `<label>${label}<input type="number" min="0" step="any" data-cin="${k}" value="${esc(b[k])}"></label>`;
    const med = b.planType === 'medicare';
    box.innerHTML = `<h3>Patient responsibility calculator <span class="hint">next visit${ctx.appt ? ` ${esc(ctx.appt)}` : ''}${ctx.event ? ` · ${esc(ctx.event)}` : ''} · not a guarantee</span></h3>
      <p class="hint">${esc([ctx.telehealth ? `Telehealth: home clinic ${ctx.home || 'unknown'}` : ctx.location, ctx.tab ? `fee tab ${ctx.tab}` : '', ctx.payer ? `${ctx.payer} (${ctx.payerSource})` : '', ctx.codeKeyword ? `codes from "${ctx.codeKeyword}"${ctx.codeNote ? `: ${ctx.codeNote}` : ''}` : ''].filter(Boolean).join(' · '))}</p>
      <div class="filters calcgrid">
        <label>Plan type<select data-cin="planType">${opt(F.CALC_PLANS, b.planType)}</select></label>
        <label>Network<select data-cin="network">${opt([['in', 'In-network'], ['out', 'Out-of-network']], b.network)}</select></label>
        <label>How does this plan charge the patient?<select data-cin="mode"${b.planType === 'self_pay' ? ' disabled' : ''}>${opt(F.CALC_MODES, b.mode)}</select></label>
      </div>
      ${med ? `<div class="filters"><label class="inline"><input type="checkbox" data-cin="secCoins"${b.secCoins ? ' checked' : ''}> Secondary / Medigap covers coinsurance (also zeroes any copay)</label>
        <label class="inline"><input type="checkbox" data-cin="secDed"${b.secDed ? ' checked' : ''}> Secondary also covers the deductible</label></div>` : ''}
      <div class="filters calcgrid${b.planType === 'self_pay' ? ' off' : ''}">${numIn('deductible', 'Annual deductible ($)')}${numIn('deductibleMet', 'Deductible met so far ($)')}${numIn('coins', 'Coinsurance, patient pays (%)')}
        ${numIn('copay', 'Copay per visit ($)')}${numIn('oopMax', 'OOP max ($, 0 = none)')}${numIn('oopMet', 'OOP met so far ($)')}</div>
      <p class="hint">Pre-filled from Sohar${b.missing.length ? `; not returned: ${esc(b.missing.join(', '))} (counted as $0)` : ''}. ${ctx.feePayer ? `A change to "How does this plan charge" is remembered for ${esc(ctx.feePayer)} (Estimate Defaults sheet); other changes are for this estimate only.` : 'No fee payer column is mapped for this patient, so changes are for this estimate only and are not remembered.'}</p>
      <div class="estcodes">${estOver.codes.map((c, i) => `<span class="estcode"><input data-ecode="${i}" value="${esc(c.code)}" aria-label="Code ${i + 1}" maxlength="8"> x <input data-eunits="${i}" type="number" min="1" step="1" value="${esc(c.units)}" aria-label="Units for code ${i + 1}">
        <input data-ecoins="${i}" type="number" min="0" max="100" step="any" value="${esc(c.coins)}" placeholder="coins %" aria-label="Coinsurance % for code ${i + 1} (blank: the plan's)"> <button class="link" data-edel="${i}" aria-label="Remove code ${i + 1}">Remove</button></span>`).join('')}
        <button class="small" data-eadd="1">Add code</button> <button class="link" data-ereset="1">Reset</button></div>
      <div id="estResults"></div>`;
    renderEstimateResults();
  }
  function renderEstimateResults() {
    const out = $('estResults');
    if (!out || !estOver) return;
    const ctx = fees && fees.contexts ? fees.contexts[estOver.per] : null;
    if (!ctx) { out.innerHTML = ''; return; }
    const b = estOver.inputs;
    const e = F.calcPatient(ctx, fees.tables, b, estOver.codes);
    const col = (lbl, est) => `<div class="estcol"><h4>${lbl}: ${esc(est && est.patient != null ? F.usd(est.patient) : '—')}</h4>${est ? `<ul class="estlines">${F.calcText(est, b).map((t) => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}</div>`;
    out.innerHTML = `${e.flags.length ? `<ul class="pc-items">${e.flags.map((f) => `<li class="sev-medium">${esc(f)}</li>`).join('')}</ul>` : ''}
      ${e.basis ? `<p class="hint">${esc(e.basis)}</p>` : ''}
      ${e.md || e.np ? `<div class="estcols">${e.same ? col(b.planType === 'self_pay' ? 'Self pay' : 'MD and NP', e.md) : `${col('MD', e.md)}${col('NP', e.np)}`}</div>` : ''}`;
  }
  function refreshEstimate() {
    const p = $('estPanel');
    if (!p || $('detail').hidden) return;
    if (!fees) { p.remove(); return; }
    // New fee data (an upload or saved defaults): what staff typed in the open panel is kept; rates follow the new data.
    const keep = estOver && estOver.per === p.dataset.per ? estOver : null;
    estOver = keep;
    renderEstimatePanel(p.dataset.per, keep ? keep.row : null);
  }
  // Plan type presets, as on the calculator page: Medicare 20% after the deductible; Medicaid and VA copay only.
  function planPreset(b) {
    if (b.planType !== 'medicare') { b.secCoins = false; b.secDed = false; }
    if (b.planType === 'medicare') { b.mode = 'ded_coins'; b.coins = 20; }
    if (b.planType === 'medicaid' || b.planType === 'va') { b.mode = 'copay_only'; b.coins = 0; b.deductible = 0; }
  }
  $('detail').addEventListener('input', (e) => {
    const c = e.target.closest('[data-ecode],[data-eunits],[data-ecoins],input[data-cin]');
    if (!c || !estOver) return;
    if (c.dataset.cin) {
      if (c.type === 'checkbox') return; // change event
      estOver.inputs[c.dataset.cin] = c.value === '' ? '' : Math.max(0, Number(c.value) || 0);
      renderEstimateResults();
      return;
    }
    const i = +(c.dataset.ecode ?? c.dataset.eunits ?? c.dataset.ecoins);
    if (!estOver.codes[i]) return;
    if (c.dataset.ecode != null) estOver.codes[i].code = c.value.trim().toUpperCase();
    else if (c.dataset.eunits != null) estOver.codes[i].units = Math.max(1, Number(c.value) || 1);
    else estOver.codes[i].coins = c.value;
    renderEstimateResults();
  });
  $('detail').addEventListener('change', (e) => {
    const s = e.target.closest('select[data-cin],input[type=checkbox][data-cin]');
    if (!s || !estOver) return;
    const k = s.dataset.cin, b = estOver.inputs;
    b[k] = s.type === 'checkbox' ? s.checked : s.value;
    if (k === 'planType') { planPreset(b); renderEstimatePanel(estOver.per, estOver.row); return; }
    if (k === 'mode') {
      // Only the charge mode is remembered, as the matching Estimate Defaults toggles, so the Benefits table agrees.
      const ctx = fees && fees.contexts ? fees.contexts[estOver.per] : null;
      if (ctx && ctx.feePayer) toPane({ type: 'fee-defaults', feePayer: ctx.feePayer, toggles: F.CALC_MODE_TOGGLES[s.value] });
    }
    renderEstimateResults();
  });
  $('detail').addEventListener('click', (e) => {
    const b = e.target.closest('[data-eadd],[data-edel],[data-ereset]');
    if (!b || !estOver) return;
    const { per, row } = estOver;
    if (b.dataset.eadd) estOver.codes.push({ code: '', units: 1, coins: '' });
    else if (b.dataset.edel != null) estOver.codes.splice(+b.dataset.edel, 1);
    else estOver = null;
    renderEstimatePanel(per, row);
  });
  async function uploadFee(files) {
    const f = files && files[0];
    if (!f) return;
    if (/\.html?$/i.test(f.name)) {
      // The West Coast Patient Responsibility Calculator (v0.20): its built-in rates complete the saved schedule.
      try {
        const html = await f.text();
        if (!F.isCalculatorHtml(html)) { notice(`${f.name} has no built-in fee schedule. Drop the West Coast Patient Responsibility Calculator, or the fee schedule .xlsx.`, false); return; }
        $('feeState').textContent = `Sending ${f.name} to Excel…`;
        toPane({ type: 'fee-upload', name: f.name, html });
      } catch (e) { notice(e.message, false); }
      return;
    }
    if (!/\.xls[xm]$/i.test(f.name)) { notice(`${f.name} is not an .xlsx or .html file. Drop the fee schedule workbook or the West Coast Patient Responsibility Calculator.`, false); return; }
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
    for (const id of ['btnCFile', 'btnCttFile']) { $(id).disabled = !$('cLoc').value; $(id).title = $('cLoc').value ? '' : 'Pick a clinic first'; }
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
    t.querySelector('tbody').innerHTML = data.length ? data.map((r) => `<tr>${cols.map((c) => cell(C.cellOf(r, c))).join('')}</tr>`).join('') : `<tr><td colspan="${cols.length}" class="hint">${esc(empty)}</td></tr>`;
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
    let name, cols, data, extra = '', safe = false; // safe: typed notes, so guard against Excel formulas
    if (kind === 'verification' || kind === 'discovery') {
      if (!exportsData) { notice('No Sohar import built yet. Upload the NextGen report in the pane first.', false); return; }
      name = C.fileName(kind);
      cols = kind === 'verification' ? C.VERIFICATION_COLS : C.DISCOVERY_COLS; data = exportsData[kind] || [];
    } else if (kind === 'discoveryVerify') {
      if (!discVerify.length) { notice('No discovery patients are waiting for verification. Upload a discovery output with coverage found first.', false); return; }
      try { C.namesMissing(discVerify); } catch (e) { notice(e.message, false); return; }
      name = C.fileName('discoveryVerify'); cols = C.VERIFICATION_COLS; data = discVerify;
    } else if (kind === 'verifyFailed') {
      if (!verFailed.length) { notice('No failed verifications are waiting for discovery.', false); return; }
      try { C.namesMissing(verFailed); } catch (e) { notice(e.message, false); return; }
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
    } else if (kind === 'tasks') { name = C.fileName('tasks').replace('Tasks', 'Tasks (open)'); cols = C.TASK_COLS; data = tasks.filter(open); safe = true; }
    else if (kind === 'allTasks') { name = C.fileName('tasks'); cols = C.TASK_COLS; data = tasks; safe = true; }
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
    const csv = C.toCSV(cols, data, { safe });
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
  $('btnCttFile').addEventListener('click', () => { if ($('cLoc').value) { notice(`Saving the Clinic Template Tool file for ${$('cLoc').value}…`, true); toPane({ type: 'ctt-file', location: $('cLoc').value }); } });
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
    // v0.20 patient record: Rivera's earlier checks (a discovery, then a verification with less of the deductible met),
    // one older check from the payer outcomes history, Flow estimates, pronouns and a guarantor check. All made up.
    res[0]['Sohar Date'] = '2026-10-18T16:05:00Z';
    res.push(
      { ...res[0], 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-09-12T15:30:00Z', 'Deductible Remaining': '600', 'Deductible Status': 'Not met', 'OOP Remaining': '3100', 'OOP Status': 'Not met', Action: 'Superseded by a later result', 'Sohar Ref': 'q-1a', 'Work Queue': 'Clear' },
      { ...res[0], 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Date': '2026-08-30T17:10:00Z', Plan: '', Deductible: '', 'Deductible Remaining': '', 'OOP Max': '', 'OOP Remaining': '', Copay: '', Coinsurance: '', 'Deductible Status': '', 'OOP Status': '', Action: 'Coverage found: verify it.', 'Sohar Ref': 'q-1b', 'Work Queue': 'Clear' });
    const demoOutcomes = [{ MRN: '80100', Kind: 'Verification', 'Sohar Status': 'complete.ineligible', 'Sohar Date': '2026-06-02T15:00:00Z', 'Returned Payer': 'Old Synthetic Plan', 'NextGen Payer': 'Old Synthetic Plan', Signal: '' }];
    const fl = (o) => ({ ...Object.fromEntries(REC.FLOW_SHEET_COLS.map((c) => [c, ''])), MRN: '80100', LASFIR: C.lasfir('Rivera', 'Ana'), Location: 'San Rafael', 'Run Status': 'COMPLETE', 'Discovery Status': 'complete',
      'Policy vs NextGen': 'Same policy as NextGen', 'Estimate Policy': 'NextGen', Network: 'INN', 'Sohar Payer': 'UMR', 'Fee Tab': 'Marin', File: 'Flow Results (SANDBOX).csv', ...o });
    const demoFlow = [fl({ Key: 'd1', 'Appt Date': '10/20/2026', Time: '9:00 AM', Event: 'TMS Treatment Unilateral', 'Rate Level': 'NP', 'Est Codes': '', 'Estimate Note': 'TMS (intervention): no estimate; verify deductible, OOP max and insurance status' }),
      fl({ Key: 'd2', 'Appt Date': '10/27/2026', Time: '9:00 AM', Event: 'Med Management In-Office', 'Rate Level': 'NP', 'Est Codes': '99214', 'Sohar Est Patient': '71.11', 'Sohar Est Plan Pays': '100.00' })];
    const demoProfiles = { 100100: { pronouns: 'They, Them, Theirs', genderIdentity: '', currentGender: 'Female', guarantor: 'RIVANA (the patient)', checks: [
      { sentAs: 'Verification', payer: 'UMR', date: '2026-10-18', ...REC.guarantorCheck({ patient: { firstName: 'Ana', lastName: 'Rivera' }, guarantor: { firstName: 'Ana', lastName: 'Rivera' }, guarantorSelf: true }, { firstName: 'Luis', lastName: 'Rivera' }) }] } };
    // Report visits with Pre-Check Complete (made up): Rivera's first visit done in NextGen, Lindqvist's not.
    const dv = (key, date, time, event, loc, pc, ticks) => {
      const v = { date, time, event, loc, provider: '', status: '', precheck: pc, hasPrecheck: true, key };
      return { ...v, lasfir: '', items: REC.precheckItems(v, key, ticks || []), notes: REC.precheckNotes(v, key, ticks || []) };
    };
    const demoVisits = {
      100100: [dv('80100', '10/20/2026', '9:00 AM', 'TMS Treatment Unilateral', 'San Rafael', 'Yes', [{ Key: '80100|10/20/2026|9:00 AM', Item: 'nextgen', Done: 'Yes', By: 'Jordan Lee', When: '10/19/2026 4:10 PM' },
        { Key: '80100|10/20/2026|9:00 AM', Item: 'note', Note: 'New card scanned; plan unchanged', By: 'Jordan Lee', When: '10/19/2026 4:12 PM' }]),
        dv('80100', '10/27/2026', '9:00 AM', 'Med Management In-Office', 'San Rafael', '')],
      100102: [dv('PER-100102', '10/22/2026', '10:00 AM', 'Med Management In-Office', 'Walnut Creek', 'No')],
    };
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
    // One task the clinic ticked done with a note (Clinic Template Tool file or the clinic inbox), made up.
    const demoTasks = C.buildTasks(res).map((t, i) => (i === 0 ? { ...t, Done: `Yes (clinic ${t['Next Location']})`, 'Clinic Note': `Called the patient, new card scanned (Front desk, ${t['Next Location']}, 10/19)` } : t));
    load({ results: res, tasks: demoTasks, exports: { verification: [], discovery: [] }, asOf: '2026-10-13T00:00:00.000Z', fees: demoFees,
      clinic: demoClinic, skipped: demoSkipped, skipCols: ['Location', 'Next Appt', 'MRN', 'Per Nbr', 'Last Name', 'First Name', 'Verified', 'Next Check', 'Other Insurance Sent'], learning: demoLearning, guarantor: { count: 2, noDob: 1 },
      flow: demoFlow, outcomes: demoOutcomes, profiles: demoProfiles, visits: demoVisits, admins: ['Jordan Lee', 'Sam Ortiz'],
      lastUpdated: { report: new Date(Date.now() - 2 * 36e5).toISOString(), sohar: new Date(Date.now() - 3 * 864e5).toISOString(), flow: new Date().toISOString() } });
    // #patient opens the first demo patient's record (for checking the layout); other hashes pick a tab.
    if (location.hash === '#patient') showDetail(`${keyOf(res[0])}|${res[0]['Sohar Ref']}`);
    else if (location.hash) showTab(location.hash.slice(1) === 'exports' ? 'upload' : location.hash.slice(1));
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
