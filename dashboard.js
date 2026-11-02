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
  let notes = {}; // patient key (MRN) -> the patient's notes, newest first (record.patientNotes)
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
    else if (m.type === 'fee-update') { fees = m.fees || null; renderFeeBar(); renderBenefits(); if (!$('detail').hidden && detailKey) renderDetail(); }
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
    visits = m.visits || {}; lastUpdated = m.lastUpdated || null; admins = m.admins || []; notes = m.notes || {};
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
  const precheckPill = (r) => ({ done: ' <span class="pill p-blue">Pre-check done</span>', open: ' <span class="pill p-amber">Pre-check open</span>' }[precheckOf(r)] || '');
  /** A red pill when the NextGen guarantor and a Sohar subscriber are different people (record.guarantorCheck, high). */
  // Amber when the patient is the guarantor but Sohar's subscriber is someone else (check the policy's insured).
  const guarFlag = (r) => {
    const cs = (profiles[String(r['Per Nbr'] || '')] || {}).checks || [];
    if (cs.some((c) => c.severity === 'high')) return ' <span class="pill p-red">Guarantor: Different people</span>';
    return cs.some((c) => c.severity === 'medium') ? ' <span class="pill p-amber">Guarantor: Check insured</span>' : '';
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

  // ---------- patient view (v0.22: one record per MRN, sidebar + one section at a time) ----------
  // A floating window over the dashboard: the sidebar has identity, freshness, the guarantor check, the sections, the
  // admin's name and previous / next patient; the main area shows one section. The derivations are in record.js
  // (plansOf, milestones, sinceLast, projection, actionItems, leadLine); this lays them out.
  const SECTIONS = [['overview', 'Overview'], ['visits', 'Visits and pre-check'], ['benefits', 'Benefits projection'], ['activity', 'Activity and plans'],
    ['guarantor', 'Guarantor'], ['calc', 'Calculator'], ['tasks', 'NextGen tasks']];
  let detailKey = '', detailDk = '', sec = 'overview', order = [], animKey = '';
  const ui = {}; // patient key -> {open: visit row open on Visits, sel: timeline row open, todoOpen: all action items shown}
  const uiOf = (k) => (ui[k] = ui[k] || { open: 0, sel: '', todoOpen: false });
  const keyOfDetail = (dk) => { const r = rows.find((x) => `${keyOf(x)}|${x['Sohar Ref']}` === dk); return r ? (keyOf(r) || `ref:${r['Sohar Ref']}`) : ''; };
  const num = (v) => { const n = parseFloat(String(v ?? '').replace(/[$,%\s]/g, '')); return isNaN(n) ? null : n; };
  const money = (n) => `$${(Math.round(n * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const money0 = (n) => `$${Math.round(n).toLocaleString('en-US')}`;
  const sd = (v) => REC.shortDate(v);
  const minsOf = (t) => { const m = String(t || '').match(/(\d+):(\d+)\s*(AM|PM)/i); return m ? ((+m[1] % 12) + (/pm/i.test(m[3]) ? 12 : 0)) * 60 + +m[2] : 0; };
  const tag = (tone, label) => `<span class="tag t-${tone}">${esc(label)}</span>`;
  const mark = (tone, label) => `<span class="tag mark t-${tone}"><i aria-hidden="true"></i>${esc(label)}</span>`;
  /** A tick or note stamp ("11/2/2026 1:30 PM") as "Nov 2, 1:30 PM". */
  const niceStamp = (s) => { const m = String(s || '').match(/^(\d{1,2}\/\d{1,2}\/\d{4})\s+(.+)$/); return m ? `${sd(m[1])}, ${m[2]}` : String(s || ''); };
  /** A time in ms as "Oct 29, 4:40 PM" (the year too when it isn't this year). */
  const whenOf = (t) => {
    if (!isFinite(t)) return '';
    const d = new Date(t);
    return d.toLocaleString('en-US', { month: 'short', day: 'numeric', year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
  };
  const fmtWhen = (e) => (isFinite(e.time) ? whenOf(e.time) : e.date);
  const planTitle = (r) => [r['Sohar Payer'] || r['NextGen Payer'], r.Plan].filter(Boolean).join(' · ');
  const nameOfDk = (dk) => { const r = rows.find((x) => `${keyOf(x)}|${x['Sohar Ref']}` === dk); return r ? name(r) : ''; };

  /** Everything the patient view shows for one patient key. */
  function patientCtx(k) {
    const rec = k.startsWith('ref:') ? null : REC.patientRecord(rows, k, { outcomes, flow, asOf });
    const r = rec ? rec.lead : rows.find((x) => `ref:${x['Sohar Ref']}` === k);
    if (!r) return null;
    const per = String(r['Per Nbr'] || ''), prof = profiles[per] || null, checks = (prof && prof.checks) || [];
    const its = tasks.filter((t) => (rec ? keyOf(t) === k : t['Sohar Ref'] === r['Sohar Ref']));
    const vs = visits[per] || [], fl = rec ? rec.flow : [];
    const fresh = REC.freshness(lastUpdated || {}, new Date());
    const plans = rec ? REC.plansOf(rec.results, r) : [];
    // The fee context (fees.js prepareEstimates) carries the payer's per-code rules; a visit counts toward the
    // deductible per its codes' rules when staff classified them (v0.23).
    const fx = fees && fees.contexts ? fees.contexts[per] || null : null;
    const proj = REC.projection(rec && rec.financial, fl, { dedApplies: (f) => dedAppliesOf(fx, codesOfVisit({ f, event: f.Event })) });
    const items = visitItems(vs, fl);
    const next = items[0] ? items[0].date : r['Next Appt'];
    return { k, rec, r, per, prof, checks, its, vs, fl, fx, fresh, plans, proj, items, next, marks: REC.milestones(rec ? rec.results : [r], plans),
      todo: REC.actionItems({ lead: r, fresh, checks, visits: vs, flow: fl, tasks: its.filter(open), proj }),
      lead: REC.leadLine({ lead: r, checks, fresh, proj, next }) };
  }
  /** A visit's codes: the Flow estimate's codes, else the Estimate Codes row for its event. [{code, units}] */
  function codesOfVisit(x) {
    const fromFlow = x.f && x.f['Est Codes'] ? F.parseCodeList(x.f['Est Codes']).codes : [];
    if (fromFlow.length) return fromFlow;
    const ce = fees && x.event ? F.codesForEvent(x.event, fees.codeRows || []) : null;
    return ce && !ce.none ? ce.codes : [];
  }
  /**
   * Whether the visit counts toward the deductible by its codes' saved rules: true when any ruled code counts, false
   * when every code is ruled and none counts, null (the projection's copay rule decides) otherwise.
   */
  function dedAppliesOf(fx, codes) {
    const rules = (fx && fx.codeRules) || {};
    const ruled = codes.map((c) => rules[c.code]).filter(Boolean);
    if (ruled.some(F.countsTowardDeductible)) return true;
    return ruled.length && ruled.length === codes.length ? false : null;
  }
  /** Report visits (with their ticks) and Flow rows joined on date and time, soonest first. */
  function visitItems(vs, fl) {
    const key = (d, t) => { const x = C.parseDate(d); return `${x ? x.getTime() : d}|${minsOf(t)}`; };
    const out = vs.map((v, i) => ({ v, i, f: null, date: v.date, time: v.time, event: v.event, loc: v.loc, prov: v.provider }));
    for (const f of fl) {
      const m = out.find((x) => x.v && !x.f && key(x.date, x.time) === key(f['Appt Date'], f.Time));
      if (m) m.f = f; else out.push({ v: null, i: -1, f, date: f['Appt Date'], time: f.Time, event: f.Event, loc: f.Location, prov: f.Provider });
    }
    const at = (x) => { const d = C.parseDate(x.date); return (d ? d.getTime() : 8.64e15) + minsOf(x.time) * 6e4; };
    return out.sort((a, b) => at(a) - at(b));
  }
  /** The freshness pill: "Up to date", or the oldest source and its age; the title lists every source. */
  function freshPill(f) {
    if (f.level === 'none') return '';
    const old = f.parts.filter((p) => p.key !== 'flow').sort((a, b) => (a.time ?? -Infinity) - (b.time ?? -Infinity))[0];
    const days = old && old.time != null ? Math.floor((Date.now() - old.time) / 864e5) : null;
    const label = f.level === 'fresh' ? 'Up to date' : `${old.label} ${days == null ? 'never uploaded' : `${days} day${days === 1 ? '' : 's'} old`}`;
    const title = f.parts.map((p) => `${p.label}: ${p.time == null ? 'never' : whenOf(p.time)}`).join(' · ');
    return `<span class="tag lg t-${{ fresh: 'blue', aging: 'amber', stale: 'red' }[f.level]}" title="${esc(title)}" aria-label="${esc(`${label}. ${title}`)}"><i class="dot" aria-hidden="true"></i>${esc(label)}</span>`;
  }
  /** [tone, label] for the patient's guarantor check: the worst of the per-result checks. */
  function guarState(cs) {
    if (cs.some((x) => x.severity === 'high')) return ['red', 'Different people'];
    if (cs.some((x) => x.severity === 'medium')) return ['amber', 'Check insured'];
    if (cs.some((x) => x.status === 'Match')) return ['green', 'Match'];
    if (cs.some((x) => x.severity === 'low')) return ['grey', 'Other guarantor'];
    return ['grey', 'Not checked'];
  }
  const checkTone = (x) => (x.severity === 'high' ? 'red' : x.severity === 'medium' ? 'amber' : x.status === 'Match' ? 'green' : 'grey');
  const preDone = (v) => { const pc = v && v.items.find((i) => i.id === 'precheck'); return !!(pc && pc.done); };

  function showDetail(dk, from) {
    const k = keyOfDetail(dk);
    if (!k) return;
    // Previous / next follow the list the patient was opened from (one entry per patient).
    if (from) {
      const panel = from.closest('[data-panel]') || document;
      const seen = new Set();
      order = [...panel.querySelectorAll('tr[data-key]')].map((tr) => tr.dataset.key).filter((x) => { const pk = keyOfDetail(x); if (!pk || seen.has(pk)) return false; seen.add(pk); return true; });
      sec = 'overview';
    }
    detailKey = k; detailDk = dk;
    renderDetail();
    if (!$('detailBackdrop')) { const bd = document.createElement('div'); bd.className = 'backdrop'; bd.id = 'detailBackdrop'; bd.addEventListener('click', hideDetail); document.body.append(bd); }
    $('btnHide').focus();
  }
  // the patient view floats over the dashboard; everything behind it is blurred
  function hideDetail() { $('detail').hidden = true; estOver = null; detailKey = ''; detailDk = ''; const bd = $('detailBackdrop'); if (bd) bd.remove(); }

  /** Draw the open patient's view. focus: a selector to focus afterwards (a tick that was just changed). */
  function renderDetail(focus) {
    const c = detailKey ? patientCtx(detailKey) : null;
    const box = $('detail');
    if (!c) { hideDetail(); return; }
    const r = c.r, p = c.prof || {};
    const ids = [p.pronouns, p.currentGender, p.genderIdentity ? `Identity: ${p.genderIdentity}` : ''].filter(Boolean);
    const [gt, gl] = guarState(c.checks);
    const openPre = c.vs.filter((v) => !preDone(v)).length, openTasks = c.its.filter(open).length;
    const badge = { visits: openPre ? tag('amber', String(openPre)) : '', guarantor: gt === 'red' ? tag('red', 'Fix') : gt === 'amber' ? tag('amber', 'Check') : '', tasks: openTasks ? tag('grey', String(openTasks)) : '' };
    const at = order.indexOf(detailDk);
    const step = (d) => {
      if (order.length < 2 || at < 0) return '';
      const nm = nameOfDk(order[(at + d + order.length) % order.length]);
      return `<button type="button" class="pv-btn" data-step="${d}" aria-label="${d < 0 ? 'Previous' : 'Next'} patient, ${esc(nm)}">${d < 0 ? '↑' : '↓'} ${esc(nm)}</button>`;
    };
    const anim = sec === 'overview' && animKey !== c.k;
    if (anim) animKey = c.k;
    box.innerHTML = `<div class="pv">
      <aside class="pv-side" aria-label="Patient">
        <button type="button" class="pv-back" id="btnHide">← All patients</button>
        <div class="pv-who">
          <div class="pv-name"><h2 id="pvTitle">${esc(name(r))}</h2><span class="pv-mrn">${esc(idText(r))}</span></div>
          ${ids.length ? `<div class="pv-ids">${ids.map((x) => `<span>${esc(x)}</span>`).join('')}</div>` : ''}
          ${planTitle(r) ? `<div class="pv-plan">${esc(planTitle(r))}</div>` : ''}
          <div class="pv-pills">${freshPill(c.fresh)}${tag(gt, `Guarantor: ${gl}`)}</div>
        </div>
        <nav class="pv-nav" aria-label="Record sections">${SECTIONS.map(([k, l]) => `<button type="button" data-sec="${k}"${sec === k ? ' aria-current="page"' : ''}><span>${esc(l)}</span>${badge[k] || ''}</button>`).join('')}</nav>
        <div class="pv-foot">${whoHtml()}<div class="pv-steps">${step(-1)}${step(1)}</div></div>
      </aside>
      <div class="pv-main" id="pvMain" role="region" aria-label="${esc(SECTIONS.find(([k]) => k === sec)[1])}">${sectionHtml(c, anim)}</div></div>`;
    box.setAttribute('role', 'dialog'); box.setAttribute('aria-modal', 'true'); box.setAttribute('aria-labelledby', 'pvTitle');
    box.hidden = false;
    if (sec === 'calc' && fees) renderEstimatePanel(r['Per Nbr'], r);
    fitDetail();
    if (focus) { const el = box.querySelector(focus); if (el) el.focus(); }
  }
  // Narrow (under 760px, e.g. the 350px pane): the sidebar stacks on top and the sections become a scrolling row.
  // A ResizeObserver, not a media query, because the window width varies with the pane.
  function fitDetail() { const box = $('detail'); box.classList.toggle('narrow', box.getBoundingClientRect().width < 760); }
  if (window.ResizeObserver) new ResizeObserver(() => { if (!$('detail').hidden) fitDetail(); }).observe($('detail'));

  function sectionHtml(c, anim) {
    switch (sec) {
      case 'visits': return visitsSection(c);
      case 'benefits': return benefitsSection(c);
      case 'activity': return activitySection(c);
      case 'guarantor': return guarantorSection(c);
      case 'calc': return calcSection(c);
      case 'tasks': return tasksSection(c);
      default: return overviewSection(c, anim);
    }
  }

  // ----- Overview: one line, four numbers, action items, the next visit's pre-check and the patient's notes -----
  function overviewSection(c, anim) {
    const r = c.r, u = uiOf(c.k);
    const tile = (t, i) => `<div class="pv-tile">
        <div class="pv-tile-top"><span>${esc(t.label)}</span>${tag(t.tag[0], t.tag[1])}</div>
        <div class="pv-num"><span class="${anim ? 'rise' : ''}" style="animation-delay:${i * 90}ms">${esc(t.num)}</span></div>
        <div class="pv-sub">${esc(t.sub)}</div>
        ${t.bar != null ? `<div class="pv-bar" role="img" aria-label="${esc(t.barAria)}"><span class="f-${t.tone}" style="width:${Math.round(Math.min(1, t.bar) * 100)}%"></span></div>` : ''}
        ${(t.lines || []).filter(Boolean).map((l) => `<div class="pv-line">${esc(l)}</div>`).join('')}</div>`;
    const todo = c.todo, more = todo.length > 4;
    const v0 = c.items[0];
    return `<p class="pv-lead tone-${c.lead.tone}">${esc(c.lead.text)}</p>
      <div class="pv-tiles">${[coverageTile(c), nextTile(c), benTile(c, 'ded'), benTile(c, 'oop')].map(tile).join('')}</div>
      <div class="pv-row">
        <section class="pv-card pv-todo" aria-labelledby="todoH">
          <div class="pv-card-h"><h3 id="todoH">Action items · ${todo.length}</h3>${more ? `<button type="button" class="pv-link" data-todo="1" aria-expanded="${u.todoOpen}">${u.todoOpen ? 'Show fewer' : `Show all ${todo.length}`}</button>` : ''}</div>
          ${todo.length ? `<div class="pv-scroll${u.todoOpen ? ' open' : ''}${more && !u.todoOpen ? ' fade' : ''}"><ul tabindex="0" aria-label="Action items, scrollable">${todo.map((t) => `<li>${tag(t.tone, t.tag)}<span class="pv-todo-t">${esc(t.text)}</span><span class="pv-due">${esc(t.due)}</span></li>`).join('')}</ul></div>`
            : '<p class="hint">Nothing to do for this patient.</p>'}
        </section>
        <section class="pv-card pv-next" aria-labelledby="nextH">
          <div class="pv-card-h"><h3 id="nextH">Next visit · pre-check</h3>${c.items.length ? `<button type="button" class="pv-link" data-sec="visits">All ${c.items.length} visit${c.items.length === 1 ? '' : 's'}</button>` : ''}</div>
          ${v0 ? `${visitHead(v0)}${v0.v ? ticksHtml(v0) : '<p class="hint">This visit is not on the NextGen report, so there is no pre-check to tick.</p>'}`
            : `<p class="hint">${esc(r['Next Appt'] ? `${[r['Next Appt'], r['Next Time'], r['Next Event']].filter(Boolean).join(' · ')}. Not on the NextGen report, so there is no pre-check to tick.` : 'No upcoming visit.')}</p>`}
          ${c.rec ? notesHtml(c.k) : ''}
        </section>
      </div>`;
  }
  function coverageTile(c) {
    const r = c.r, st = String(r['Sohar Status'] || '');
    const ver = r['Sent As'] === 'Verification';
    const lastVer = c.rec ? c.rec.results.find((x) => x['Sent As'] === 'Verification') : ver ? r : null;
    const lines = [lastVer ? `Last verification: ${whenOf(C.soharTime(lastVer['Sohar Date'])) || lastVer['Sohar Date'] || 'date unknown'}` : 'No verification yet'];
    const sub = [r['Sohar Payer'] || r['NextGen Payer'], r.Plan, r.Network].filter(Boolean).join(' · ');
    if (ver && /^complete\.eligible/.test(st)) return { label: 'Coverage', tag: ['green', 'Eligible'], num: 'Active', sub, lines };
    if (r['Sent As'] === 'Discovery') return { label: 'Coverage', tag: ['amber', 'Discovery only'], num: /found/i.test(st) ? 'Not verified' : 'Not found', sub: sub || st, lines };
    if (r['Work Queue'] === 'Pending') return { label: 'Coverage', tag: ['grey', 'Pending'], num: 'Waiting', sub: st, lines };
    return { label: 'Coverage', tag: ['red', /ineligible|inactive/.test(st) ? 'Not eligible' : 'Check'], num: /ineligible|inactive/.test(st) ? 'Inactive' : 'Check', sub: st, lines: [r.Action, ...lines] };
  }
  function nextTile(c) {
    const v = c.items[0], f = v && v.f;
    const sub = v ? [sd(v.date), v.time].filter(Boolean).join(', ') : 'No upcoming visit';
    if (!f) return { label: 'Next visit owes', tag: ['grey', 'No estimate'], num: '—', sub, lines: [v ? v.event : '', v ? 'No Sohar Flow run for this visit yet' : ''] };
    const amt = num(f['Sohar Est Patient']);
    if (amt == null) return { label: 'Next visit owes', tag: ['grey', 'No estimate'], num: '—', sub, lines: [v.event, f['Estimate Note'] || 'Verify deductible, OOP max and insurance status'] };
    return { label: 'Next visit owes', tag: REC.isSandbox(f) ? ['purple', 'Sandbox'] : ['grey', 'Estimate'], num: money(amt), sub, lines: [v.event, REC.isSandbox(f) ? 'Sandbox: don\'t quote or collect this' : ''] };
  }
  function benTile(c, key) {
    const label = key === 'ded' ? 'Deductible left' : 'Out-of-pocket max left';
    const fin = c.rec && c.rec.financial;
    const cols = key === 'ded' ? ['Deductible', 'Deductible Remaining', 'Deductible Status'] : ['OOP Max', 'OOP Remaining', 'OOP Status'];
    if (!fin) return { label, tag: ['grey', 'Unknown'], num: '—', sub: 'Not known yet', lines: [c.r['Sent As'] === 'Discovery' ? 'Run verification to get it' : 'Sohar returned no amounts'] };
    const T = num(fin[cols[0]]), R = num(fin[cols[1]]), status = fin[cols[2]];
    if (status === 'None') return { label, tag: ['grey', 'None'], num: '—', sub: key === 'ded' ? 'No deductible on this plan' : 'No OOP max on this plan', lines: [] };
    if (T == null || R == null) return { label, tag: ['grey', 'Unknown'], num: '—', sub: 'Not returned by Sohar', lines: [] };
    const left = Math.max(0, R), met = T - R;
    const st = left <= 0 ? 'Met' : status === 'Approaching' || left <= T * 0.2 ? 'Approaching' : 'Not met';
    const tone = { Met: 'green', Approaching: 'amber', 'Not met': 'grey' }[st];
    const atV = c.proj && c.proj[key] && c.proj[key].at;
    const first = [...c.marks.entries()].find(([, ls]) => ls.includes(key === 'ded' ? 'Deductible met, first seen' : 'OOP max met, first seen'));
    const sub = left <= 0 ? `Met${first ? ` ${sd(first[0]['Sohar Date'])}` : ''} · of ${money0(T)}` : atV ? `of ${money0(T)} · met ${sd(atV['Appt Date'])} as booked` : `of ${money0(T)} · not met by booked visits`;
    return { label, tag: [tone, st], num: money(left), sub, tone, bar: T > 0 ? met / T : 0, barAria: `${label.replace(' left', '')}: ${money(met)} of ${money(T)} met`, lines: [`${money(met)} met so far`] };
  }
  const visitHead = (x) => `<div class="pv-visit"><b>${esc([sd(x.date), x.time].filter(Boolean).join(', '))}${x.event ? ` · ${esc(x.event)}` : ''}</b>
    ${[x.loc, x.prov].some(Boolean) ? `<span class="hint">${esc([x.loc, x.prov].filter(Boolean).join(' · '))}</span>` : ''}</div>`;
  /** The visit's two ticks, each with who and when; data-pc is the visit's index in the report visits. */
  function ticksHtml(x) {
    const v = x.v;
    return `<p class="hint">NextGen "Pre-Check Complete": ${esc(v.precheck || 'blank')}. A tick here wins over the report.</p>
      ${v.items.map((it) => `<label class="pv-tick"><input type="checkbox" data-pc="${x.i}" data-pc-item="${esc(it.id)}"${it.done ? ' checked' : ''}>
        <span><b>${esc(it.label)}</b><span class="hint" data-pc-note="${x.i}|${esc(it.id)}">${esc(pcNote(it, v))}</span></span></label>`).join('')}
      <p class="hint">The add-in doesn't change NextGen. Tick "Updated in NextGen" after you make the change there.</p>`;
  }
  const pcNote = (it, v) => (it.source === 'Marked here' ? `Marked here · ${[it.by, niceStamp(it.when)].filter(Boolean).join(' · ')}` : it.source === 'NextGen report' ? `From NextGen report: ${(v && v.precheck) || (it.done ? 'Yes' : 'No')}` : 'Not ticked yet');
  /** The patient's notes, newest first, with an add box. A note written against a visit before v0.21.1 says which. */
  const noteLine = (n) => `<div class="pv-note"><div class="hint">${esc([n.by, niceStamp(n.when), n.visit ? `visit ${n.visit}` : ''].filter(Boolean).join(' · '))}</div><div>${esc(n.note)}</div></div>`;
  function notesHtml(k) {
    return `<div class="pv-notes"><h4>Patient notes</h4>
      <div class="pv-note-add"><textarea id="pnText" rows="2" maxlength="${REC.NOTE_MAX}" placeholder="Add a note for this patient" aria-label="Note for this patient"></textarea><button type="button" class="pv-btn" id="pnAdd">Add note</button></div>
      <div id="pnList">${(notes[k] || []).map(noteLine).join('') || '<p class="hint" id="pnNone">No notes yet.</p>'}</div></div>`;
  }

  // ----- Visits and pre-check: one row per booked visit, the Flow estimate and the ticks inside -----
  function visitsSection(c) {
    const u = uiOf(c.k);
    if (!c.items.length) return `<h3 class="pv-h">Visits and pre-check</h3><p class="hint">${esc(c.rec && c.rec.upcoming.length ? c.rec.upcoming.join('; ') : 'No upcoming visits.')}</p>`;
    return `<h3 class="pv-h">Visits and pre-check</h3>
      <section class="pv-card pv-list">${c.items.map((x, j) => {
        const f = x.f, amt = f ? num(f['Sohar Est Patient']) : null, opened = u.open === j;
        const owes = amt == null ? 'No estimate' : `${money(amt)}${REC.isSandbox(f) ? ' sandbox' : ''}`;
        const pre = !x.v ? tag('grey', 'Not on report') : preDone(x.v) ? tag('green', 'Pre-check done') : tag('amber', 'Pre-check open');
        const da = dedAppliesOf(c.fx, codesOfVisit(x));
        const dedTag = da === true ? tag('grey', 'Counts toward deductible') : da === false ? tag('grey', 'Deductible waived') : '';
        return `<div class="pv-vrow"><button type="button" class="pv-vhead" data-vopen="${j}" aria-expanded="${opened}">
            <span class="pv-when"><b>${esc(sd(x.date))}</b><span class="hint">${esc(x.time || '')}</span></span>
            <span class="pv-what"><span>${esc(x.event || '')}</span><span class="hint">${esc([x.loc, x.prov].filter(Boolean).join(' · '))}</span></span>
            <span class="pv-owes">${esc(owes)}</span>${dedTag}${pre}</button>
          ${opened ? `<div class="pv-vbody"><div class="pv-est">${estimateHtml(x)}${codeRulesHtml(c, x)}</div><div class="pv-pre">${x.v ? ticksHtml(x) : '<p class="hint">Not on the NextGen report, so there is no pre-check to tick.</p>'}</div></div>` : ''}</div>`;
      }).join('')}</section>`;
  }
  /**
   * How each of the visit's codes counts for the payer (v0.23): the plan's charge mode, or a class staff pick here.
   * A pick is saved for the payer on Estimate Defaults with the admin's name; estimates and the projection follow it.
   */
  function codeRulesHtml(c, x) {
    const codes = codesOfVisit(x);
    if (!codes.length) return '';
    const head = '<div class="pv-est-h pv-codes-h">How the codes count</div>';
    const fx = c.fx;
    if (!fx) return `${head}<p class="hint">${fees ? 'Needs a verification with benefits for this patient.' : 'Turn on fee estimates (Config "Fee estimates") to classify codes.'}</p>`;
    if (!fx.ruleKey) return `${head}<p class="hint">No payer for this patient, so the codes can't be classified.</p>`;
    const mode = (F.CALC_MODES.find(([k]) => k === F.calcModeOf(fx.defaults)) || ['', 'the plan default'])[1];
    return `${head}${codes.map((cd) => {
      const t = (fx.codeRules || {})[cd.code], id = t ? F.codeClassOf(t) : '';
      return `<label class="pv-crule"><b>${esc(cd.code)}</b><select data-crule="${esc(cd.code)}" aria-label="How ${esc(cd.code)} counts for ${esc(fx.ruleKey)}">
        <option value=""${id ? '' : ' selected'}>Plan charge mode: ${esc(mode)}</option>
        ${F.CODE_CLASSES.map((k) => `<option value="${k.id}"${k.id === id ? ' selected' : ''}>${esc(k.label)}</option>`).join('')}
        ${id === 'custom' ? `<option value="custom" selected disabled>As typed on Estimate Defaults: ${esc(F.ruleLabel(t))}</option>` : ''}</select></label>`;
    }).join('')}<p class="hint">Saved for ${esc(fx.ruleKey)} with your name. Estimates and the projection follow it.</p>`;
  }
  function estimateHtml(x) {
    const f = x.f;
    if (!f) return '<div class="pv-est-h">Sohar Flow estimate</div><p class="hint">No Sohar Flow run for this visit yet.</p>';
    const amt = num(f['Sohar Est Patient']), plan = num(f['Sohar Est Plan Pays']);
    const done = /^complete/i.test(f['Run Status']);
    return `<div class="pv-est-h">Sohar Flow estimate</div>
      ${REC.isSandbox(f) ? '<div class="pv-sandbox">Sandbox estimate. Don\'t quote or collect this amount.</div>' : ''}
      ${amt != null ? `<div class="pv-est-amt"><span class="pv-big">${esc(money(amt))}</span><span class="hint">patient owes${plan != null ? ` · plan pays ${esc(money(plan))}` : ''}</span></div>
        <div class="hint">${esc([f['Est Codes'], f['Rate Level'] ? `${f['Rate Level']} rate` : '', f['Fee Tab'], f['Estimate Policy'] ? `${f['Estimate Policy']} policy` : ''].filter(Boolean).join(' · '))}</div>
        ${f['Estimate Note'] ? `<div class="hint">${esc(f['Estimate Note'])}</div>` : ''}`
        : `<div><b>No estimate</b></div><div class="hint">${esc(done ? f['Estimate Note'] || 'Verify deductible, OOP max and insurance status.' : `Flow ${String(f['Run Status'] || 'not started').toLowerCase().replace(/_/g, ' ')}.`)}</div>`}
      ${f['Policy vs NextGen'] ? `<div>Policy vs NextGen: ${esc(f['Policy vs NextGen'])}</div>` : ''}
      ${f['Run Error'] ? `<div class="warn-text">${esc(f['Run Error'])}</div>` : ''}`;
  }

  // ----- Benefits projection: the booked visits' estimates applied to the deductible and the OOP max -----
  function benefitsSection(c) {
    const head = '<div class="pv-hrow"><h3 class="pv-h">Benefits projection</h3><span class="hint">From booked visits and their Flow estimates. A projection, not a quote.</span></div>';
    const p = c.proj;
    if (!p) return `${head}<p class="hint">No projection yet. It needs the deductible and OOP max from a verification.</p>`;
    const meter = (label, m) => {
      if (!m) return '';
      const already = m.met >= m.total || m.left <= 0, appr = !already && m.left <= m.total * 0.2;
      const tone = already || m.at ? 'green' : appr ? 'amber' : 'grey';
      const metPct = m.total > 0 ? Math.min(100, (m.met / m.total) * 100) : 100, projPct = m.total > 0 ? Math.min(100 - metPct, (m.applied / m.total) * 100) : 0;
      const right = already ? 'Met' : m.at ? `${money(m.left)} left · met ${sd(m.at['Appt Date'])} as booked` : appr ? `${money(m.left)} left · approaching` : `${money(m.left)} left`;
      return `<div class="pv-meter"><div class="pv-meter-h"><b>${esc(label)} ${esc(money0(m.total))}</b><b class="tone-${tone}">${esc(right)}</b></div>
        <div class="pv-mbar" role="img" aria-label="${esc(`${label}: ${money(m.met)} of ${money(m.total)} met; the striped part is projected from booked visits`)}"><span class="met f-${tone}" style="width:${metPct}%"></span><span class="proj f-${tone}" style="left:${metPct}%;width:${projPct}%"></span></div>
        <div class="hint">${esc(`${money(m.met)} met · ${money(m.left)} left`)}</div></div>`;
    };
    const missing = c.items.filter((x) => !x.f || num(x.f['Sohar Est Patient']) == null || REC.isSandbox(x.f));
    const missText = (x) => `${sd(x.date)} ${x.event || ''}${x.f && REC.isSandbox(x.f) && num(x.f['Sohar Est Patient']) != null ? ' (sandbox)' : ''}`.trim();
    const line = (m, what) => (!m ? '' : m.left <= 0 ? `${what} already met.` : m.at ? `${what} projected to be met at the ${sd(m.at['Appt Date'])} visit.` : `${what} not met by booked visits. ${money(Math.max(0, m.left - m.applied))} left after the last estimate.`);
    return `${head}<section class="pv-card pv-proj">
        ${meter('Deductible', p.ded)}${meter('Out-of-pocket max', p.oop)}
        ${p.rows.length ? `<div class="pv-prows">${p.rows.map((x) => `<div class="pv-prow"><b class="pv-pdate">${esc(sd(x.flow['Appt Date']))}</b><span class="pv-pev">${esc(x.flow.Event || '')}</span>
          <b class="pv-pamt">${esc(x.applied == null ? (REC.isSandbox(x.flow) && x.est != null ? 'Sandbox' : 'No estimate') : x.est != null && x.applied < x.est ? `${money(x.applied)} (est. ${money(x.est)})` : money(x.applied))}</b>
          <span class="hint pv-pleft">${x.oopLeft == null ? '' : `OOP left ${x.applied == null ? '—' : x.oopLeft === 0 ? 'Met' : money(x.oopLeft)}`}</span>
          ${x.ruled && x.applied != null ? `<span class="hint">${x.toward ? 'Counts toward deductible' : 'Deductible waived'}</span>` : ''}
          ${x.marks.map((l) => mark(/approaching/i.test(l) ? 'amber' : 'green', l)).join('')}</div>`).join('')}</div>` : '<p class="hint">No Sohar Flow estimates for the booked visits yet.</p>'}
        <div class="pv-psum">${[line(p.ded, 'Deductible'), line(p.oop, 'OOP max')].filter(Boolean).map((t) => `<div>${esc(t)}</div>`).join('')}
          ${missing.length ? `<div class="tone-amber"><b>Incomplete: ${esc(missing.map(missText).join(', '))} ${missing.length > 1 ? 'have' : 'has'} no estimate and ${missing.length > 1 ? 'are' : 'is'} left out.</b></div>` : c.items.length ? '<div class="hint">Every booked visit has an estimate.</div>' : ''}</div>
      </section>`;
  }

  // ----- Activity and plans: what changed, the current plan's timeline, then earlier plans -----
  function activitySection(c) {
    const u = uiOf(c.k);
    const since = REC.sinceLast(c.plans);
    const cur = c.plans.find((g) => g.current), prev = c.plans.filter((g) => !g.current);
    const ledger = c.rec ? c.rec.timeline.filter((e) => e.source === 'ledger') : [];
    // Ledger rows keep only the payer: one for a payer with no saved result is an earlier plan of its own.
    const same = (a, b) => !!a && !!b && a.toLowerCase() === b.toLowerCase();
    for (const e of ledger) {
      if (!e.payer || same(e.payer, cur && cur.payer) || prev.some((g) => same(g.payer, e.payer))) continue;
      prev.push({ payer: e.payer, memberId: '', results: [], current: false, superseded: false });
    }
    const planOfLedger = (e) => prev.find((g) => same(g.payer, e.payer)) || cur;
    const evs = [...(cur ? cur.results : [c.r]).map((r) => resultEvent(r, c.marks)), ...ledger.filter((e) => planOfLedger(e) === cur).map(ledgerEvent), ...c.fl.map(flowEvent)]
      .sort((a, b) => (isFinite(b.time) ? b.time : -Infinity) - (isFinite(a.time) ? a.time : -Infinity));
    const evHtml = (e) => {
      const opened = u.sel === e.id;
      return `<li><button type="button" class="pv-ev${opened ? ' sel' : ''}" data-tl="${esc(e.id)}" aria-expanded="${opened}" title="${esc(`${e.kind} · ${e.when} · ${e.hover || e.title}`)}">
          <i class="pv-dot${e.solid ? ' solid' : ''}" aria-hidden="true"></i>
          <span class="pv-ev-b"><span class="pv-ev-top"><b>${esc(e.kind)}</b><span class="hint">${esc(e.when)}</span><span class="pv-ev-st tone-${e.tone}">${esc(e.status)}</span></span>
            <span>${esc(e.title)}</span>${e.chips.map((l) => mark(/met/i.test(l) && !/approaching/i.test(l) ? 'green' : /approaching/i.test(l) ? 'amber' : 'grey', l)).join('')}</span></button>
        ${opened ? `<div class="pv-ev-d">${e.rows.filter(([, v]) => v).map(([k, v]) => `<div><span class="hint">${esc(k)}</span><span>${esc(v)}</span></div>`).join('')}
          ${e.ledger ? '<p class="hint">History ledger entry. Only date, kind, status and payer were kept.</p>' : ''}${/^https:\/\//.test(e.link || '') ? `<a href="${esc(e.link)}" target="_blank" rel="noopener">Open in Sohar</a>` : ''}</div>` : ''}</li>`;
    };
    const prevHtml = (g) => {
      const top = g.results[0] || {};
      const amt = (a, b) => (num(top[a]) == null ? '' : `${money0(num(top[a]))}${num(top[b]) == null ? '' : ` · ${money(num(top[a]) - num(top[b]))} met at the last check`}`);
      const facts = [['Deductible', amt('Deductible', 'Deductible Remaining')], ['Out-of-pocket max', amt('OOP Max', 'OOP Remaining')], ['Superseded', g.superseded ? 'A later result replaced this plan' : '']];
      const led = ledger.filter((e) => planOfLedger(e) === g);
      const last = top['Sohar Date'] ? sd(top['Sohar Date']) : led[0] ? sd(led[0].date) : '';
      return `<details open><summary><b>${esc(g.payer || 'Unknown payer')}</b> <span class="hint">${esc([g.memberId, `${g.results.length + led.length} result${g.results.length + led.length === 1 ? '' : 's'}`, last ? `last ${last}` : ''].filter(Boolean).join(' · '))}</span></summary>
        ${facts.filter(([, v]) => v).map(([k, v]) => `<div class="pv-fact"><span class="hint">${esc(k)}</span><span>${esc(v)}</span></div>`).join('')}
        ${[...g.results.map((r) => resultEvent(r, c.marks)), ...led.map(ledgerEvent)].map((e) => `<div class="pv-pitem"><b>${esc(e.kind)}</b> · ${esc(e.when)} · ${esc(e.status)}<div class="hint">${esc(e.title)}</div></div>`).join('')}</details>`;
    };
    return `<h3 class="pv-h">Activity and previous plans</h3>
      <div class="pv-card pv-since"><div class="hint">Since last check</div><b>${esc(since.a)}</b><div class="hint">${esc(since.sub)}</div></div>
      <section class="pv-card pv-tl"><div class="hint pv-tl-h">${esc(cur ? [cur.payer, cur.memberId].filter(Boolean).join(' · ') : planTitle(c.r))} · current · newest first</div>
        <ol>${evs.map(evHtml).join('')}</ol></section>
      <section class="pv-card pv-prev"><h4>Previous plans</h4>${prev.length ? prev.map(prevHtml).join('') : '<p class="hint">No earlier plans for this MRN.</p>'}</section>`;
  }
  const statusTone = (s) => (/eligible$|found|active/i.test(s) && !/ineligible|not found/i.test(s) ? 'green' : /error|ineligible|inactive|not found/i.test(s) ? 'red' : 'grey');
  function resultEvent(r, marks) {
    const t = C.soharTime(r['Sohar Date']);
    const ben = [num(r['Deductible Remaining']) != null ? `deductible ${money(num(r['Deductible Remaining']))} left` : '', num(r['OOP Remaining']) != null ? `OOP max ${money(num(r['OOP Remaining']))} left` : ''].filter(Boolean).join(' · ');
    return { id: `r:${r['Sohar Ref']}`, kind: r['Sent As'] || 'Verification', time: t, when: whenOf(t) || r['Sohar Date'] || 'Date unknown', status: r['Sohar Status'] || '—', tone: statusTone(r['Sohar Status'] || ''), solid: true,
      title: [r['Sohar Payer'] || r['NextGen Payer'], r.Plan, ben].filter(Boolean).join(' · ') || r.Action || '', hover: r.Action, chips: marks.get(r) || [], link: r['Sohar Link'], ledger: false,
      rows: [['Status', [r['Sohar Status'], r['Sohar Ref'] ? `ref …${String(r['Sohar Ref']).slice(-4)}` : ''].filter(Boolean).join(' · ')], ['Network', r.Network],
        ['Coverage', [r['Coverage Start'], r['Coverage End']].some(Boolean) ? `${r['Coverage Start'] || '?'} to ${r['Coverage End'] || 'open'}` : ''], ['Member ID', r['Sohar Member ID'] || r['NextGen Member ID']],
        ['Action', r.Action], ['Flags', pillsOf(r).join(', ')], ['Copay', r.Copay !== '' && r.Copay != null ? usd(r.Copay) : ''], ['Coinsurance', r.Coinsurance], ['Searched as', r['Searched As'] && r['Searched As'] !== 'Patient' ? r['Searched As'] : '']] };
  }
  const ledgerEvent = (e) => ({ id: `l:${e.kind}|${e.date}|${e.status}`, kind: `${e.kind} · ledger`, time: e.time, when: fmtWhen(e) || 'Date unknown', status: e.status || '—', tone: statusTone(e.status || ''), solid: true,
    title: e.payer, chips: [], link: '', ledger: true, rows: [] });
  function flowEvent(f) {
    const t = Date.parse(f['Last Updated'] || f.Started), amt = num(f['Sohar Est Patient']), sb = REC.isSandbox(f);
    return { id: `f:${f.Key}`, kind: amt != null ? 'Cost estimate' : 'Flow run', time: t, when: whenOf(t) || 'Date unknown', status: sb ? 'Sandbox' : String(f['Run Status'] || '').toLowerCase().replace(/_/g, ' ') || '—', tone: sb ? 'purple' : 'grey', solid: false,
      title: [sd(f['Appt Date']), f.Event, amt != null ? money(amt) : f['Estimate Note'] || 'no estimate'].filter(Boolean).join(' · '), hover: f['Policy vs NextGen'], chips: [], ledger: false,
      link: [f['Discovered Verification Link'], f['NextGen Verification Link'], f['Discovery Link']].find((l) => /^https:\/\//.test(l || '')) || '',
      rows: [['Visit', [f['Appt Date'], f.Time, f.Event].filter(Boolean).join(' ')], ['Codes', [f['Est Codes'], f['Rate Level'] ? `${f['Rate Level']} rate` : ''].filter(Boolean).join(' · ')], ['Fee tab', f['Fee Tab']],
        ['Policy', f['Estimate Policy']], ['Policy vs NextGen', f['Policy vs NextGen']], ['Plan pays', num(f['Sohar Est Plan Pays']) != null ? money(num(f['Sohar Est Plan Pays'])) : ''], ['Error', f['Run Error']], ['File', f.File]] };
  }

  // ----- Guarantor: NextGen's guarantor against each result's Sohar subscriber -----
  function guarantorSection(c) {
    const p = c.prof || {}, cs = c.checks;
    const worst = cs.find((x) => x.severity === 'high') || cs.find((x) => x.severity === 'medium') || cs[0];
    const g = String(p.guarantor || '').replace(' (the patient)', ', the patient');
    const note = worst ? worst.text : p.guarantor ? `NextGen guarantor is ${g}. Sohar returned no subscriber to compare yet.` : 'Not checked. The NextGen report has no guarantor for this patient.';
    return `<h3 class="pv-h">Guarantor check</h3>
      <p class="pv-gnote${worst && worst.severity === 'high' ? ' tone-red' : ''}">${esc(note)}</p>
      ${cs.length ? `<section class="pv-card pv-grows">${cs.map((x) => `<div class="pv-grow"><span class="hint pv-gres">${esc([x.sentAs, whenOf(C.soharTime(x.date)) || x.date, x.payer].filter(Boolean).join(' · '))}</span>
          <span><span class="hint">NextGen guarantor</span><b>${esc(g || 'Not on the report')}</b></span>
          <span><span class="hint">Sohar subscriber</span><b>${esc(x.subscriber || 'Not returned')}</b></span>${tag(checkTone(x), x.status)}</div>`).join('')}</section>` : ''}
      ${[p.guarantorDob ? 'Guarantor DOB on file' : '', p.prevLastName ? 'Previous last name on file' : ''].filter(Boolean).length ? `<p class="hint">${esc([p.guarantorDob ? 'Guarantor DOB on file' : '', p.prevLastName ? 'Previous last name on file' : ''].filter(Boolean).join(' · '))}.</p>` : ''}`;
  }

  // ----- Calculator: the West Coast calculator panel (renderEstimatePanel) -----
  function calcSection(c) {
    if (!fees) return '<h3 class="pv-h">Calculator</h3><p class="hint">The calculator needs fee estimates turned on: set Config "Fee estimates" to On, then drop the fee schedule or the West Coast Patient Responsibility Calculator on the Benefits tab.</p>';
    return `<div class="pv-hrow"><h3 class="pv-h">Calculator</h3><span class="hint">Pre-filled from Sohar and the payer's saved code rules. Pick a service and a number of visits to price care not yet booked. Edits are for this estimate only, unless you save a code's rule for the payer.</span></div>
      <section class="pv-card dgroup est" id="estPanel" data-per="${esc(c.r['Per Nbr'])}"></section>`;
  }

  // ----- NextGen tasks: Done ticks stay in the workbook -----
  function tasksSection(c) {
    return `<h3 class="pv-h">NextGen tasks</h3>
      <section class="pv-card pv-tasks">${c.its.length ? c.its.map((t) => `<label class="pv-task${open(t) ? '' : ' done'}"><input type="checkbox" data-task="${esc(t['Task Key'])}"${open(t) ? '' : ' checked'} aria-label="Done: ${esc(t.Task)}">
          <span class="pv-task-t"><b>${esc(t.Task)}</b><span>${esc(t.Details || '')}</span>${t['Clinic Note'] ? `<span class="cnote">Clinic: ${esc(t['Clinic Note'])}</span>` : ''}<span class="hint">${esc(open(t) ? (t['Due By'] ? `Due ${sd(t['Due By'])}` : '') : String(t.Done))}</span></span>
          ${open(t) ? tag('amber', 'Open') : tag('green', 'Done')}</label>`).join('') : '<p class="hint">No NextGen tasks for this patient.</p>'}
        <p class="hint">Tasks come from the Sohar results. Done ticks stay in this workbook.</p></section>`;
  }

  // ----- the admin's name, ticks and notes (saved on the hidden Pre-Check sheet by the pane) -----
  // The name is picked once from Config "Admin names" and remembered on this computer; every tick and note records it.
  const NAME_KEY = 'sohar.precheck.by';
  let adminName = (() => { try { return localStorage.getItem(NAME_KEY) || ''; } catch (e) { return ''; } })();
  function whoHtml() {
    const names = admins.includes(adminName) || !adminName ? admins : [adminName, ...admins];
    return `<label class="pv-as">Working as
      ${names.length ? `<select id="pcBy"><option value="">Pick your name</option>${names.map((n) => `<option${n === adminName ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select>`
        : `<input id="pcBy" value="${esc(adminName)}" placeholder="Your name" maxlength="60">`}</label>
      ${names.length ? '' : '<span class="hint">Add the admin names to Config "Admin names" (separated by ;) to pick from a list.</span>'}`;
  }
  const stamp = () => { const now = new Date(); return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()} ${now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`; };
  const detailRow = () => rows.find((x) => keyOf(x) === detailKey) || rows.find((x) => `ref:${x['Sohar Ref']}` === detailKey);
  /** The admin's name, or a notice and null when none is picked. */
  function byOrAsk() {
    if (adminName) return adminName;
    notice('Pick your name first ("Working as", bottom of the sidebar). It is saved with each tick and note.', false);
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
    const r = detailRow(), v = r && (visits[String(r['Per Nbr'] || '')] || [])[+c.dataset.pc];
    const it = v && v.items.find((x) => x.id === c.dataset.pcItem);
    if (!it) return;
    const by = byOrAsk();
    if (!by) { c.checked = !c.checked; return; }
    Object.assign(it, { done: c.checked, source: 'Marked here', when: stamp(), by });
    toPane({ type: 'precheck', key: v.key, lasfir: v.lasfir || name(r), date: v.date, time: v.time, item: it.id, done: c.checked, by });
    renderDetail(`input[data-pc="${c.dataset.pc}"][data-pc-item="${CSS.escape(it.id)}"]`); // badges and action items follow
    renderPrep(); renderWorklist(); // the list pills follow
  });
  // How a code counts for the payer (Visits section): saved by the pane on Estimate Defaults; the new fee data comes back
  // as a fee-update and the view redraws.
  $('detail').addEventListener('change', (e) => {
    const s = e.target.closest('select[data-crule]');
    if (!s || !detailKey) return;
    const r = detailRow(), fx = r && fees && fees.contexts ? fees.contexts[String(r['Per Nbr'] || '')] : null;
    if (!fx || !fx.ruleKey) return;
    const by = byOrAsk();
    if (!by) { renderDetail(); return; }
    notice(`Saving how ${s.dataset.crule} counts for ${fx.ruleKey}…`, true);
    toPane({ type: 'fee-code-rule', ruleKey: fx.ruleKey, code: s.dataset.crule, cls: s.value, by });
  });
  function addNote() {
    const r = detailRow(), input = $('pnText');
    const k = r && keyOf(r), text = input ? input.value.replace(/\s+/g, ' ').trim() : '';
    if (!k || !text) return;
    const by = byOrAsk();
    if (!by) return;
    const n = { note: text.slice(0, REC.NOTE_MAX), by, when: stamp(), visit: '' };
    notes[k] = [n, ...(notes[k] || [])];
    input.value = '';
    const none = $('pnNone'); if (none) none.remove();
    $('pnList').insertAdjacentHTML('afterbegin', noteLine(n));
    toPane({ type: 'precheck-note', key: k, lasfir: name(r), note: n.note, by });
  }
  $('detail').addEventListener('click', (e) => {
    if (e.target.closest('#pnAdd')) { addNote(); return; }
    if (e.target.closest('#btnHide')) { hideDetail(); return; }
    const b = e.target.closest('[data-sec],[data-step],[data-vopen],[data-tl],[data-todo]');
    if (!b || !detailKey) return;
    const u = uiOf(detailKey);
    if (b.dataset.sec) { sec = b.dataset.sec; renderDetail(`.pv-nav [data-sec="${sec}"]`); $('pvMain').scrollTop = 0; return; }
    if (b.dataset.step) {
      const dk = order[(order.indexOf(detailDk) + +b.dataset.step + order.length) % order.length];
      if (dk) { showDetail(dk); const s = $('detail').querySelector(`[data-step="${b.dataset.step}"]`); if (s) s.focus(); }
      return;
    }
    if (b.dataset.vopen != null) { u.open = u.open === +b.dataset.vopen ? -1 : +b.dataset.vopen; renderDetail(`[data-vopen="${b.dataset.vopen}"]`); return; }
    if (b.dataset.tl) { u.sel = u.sel === b.dataset.tl ? '' : b.dataset.tl; renderDetail(`[data-tl="${CSS.escape(b.dataset.tl)}"]`); return; }
    if (b.dataset.todo) { u.todoOpen = !u.todoOpen; renderDetail('[data-todo]'); }
  });
  // Ctrl+Enter adds the note (Enter makes a new line in the box).
  $('detail').addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && e.target.id === 'pnText') { e.preventDefault(); addNote(); } });
  // Free-text name (no Config list yet): remembered as it is typed.
  $('detail').addEventListener('input', (e) => {
    if (e.target.id !== 'pcBy') return;
    adminName = e.target.value.trim();
    try { localStorage.setItem(NAME_KEY, adminName); } catch (err) { /* storage blocked */ }
  });

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
    if (!estOver || estOver.per !== per) estOver = { per, row: row || estOver?.row || {}, inputs: F.calcInputs(ctx, row), codes: (ctx.codes || []).map((c) => ({ ...c, coins: '' })), service: '', visits: 1, start: 'now' };
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
      ${serviceHtml(ctx)}
      <div class="estcodes">${estOver.codes.map((c, i) => `<span class="estcode"><input data-ecode="${i}" value="${esc(c.code)}" aria-label="Code ${i + 1}" maxlength="8"> x <input data-eunits="${i}" type="number" min="1" step="1" value="${esc(c.units)}" aria-label="Units for code ${i + 1}">
        <input data-ecoins="${i}" type="number" min="0" max="100" step="any" value="${esc(c.coins)}" placeholder="coins %" aria-label="Coinsurance % for code ${i + 1} (blank: the plan's)"> ${ruleSelect(ctx, c, i)} <button class="link" data-edel="${i}" aria-label="Remove code ${i + 1}">Remove</button></span>`).join('')}
        <button class="small" data-eadd="1">Add code</button> <button class="link" data-ereset="1">Reset</button></div>
      <div id="estResults"></div>`;
    renderEstimateResults();
  }
  function renderEstimateResults() {
    const out = $('estResults');
    if (!out || !estOver) return;
    const ctx = fees && fees.contexts ? fees.contexts[estOver.per] : null;
    if (!ctx) { out.innerHTML = ''; return; }
    const b = startInputs(estOver.inputs);
    const codes = calcCodes(estOver.codes);
    if (estOver.visits > 1) { out.innerHTML = seriesHtml(F.calcSeries(ctx, fees.tables, b, codes, estOver.visits), b); return; }
    const e = F.calcPatient(ctx, fees.tables, b, codes);
    const col = (lbl, est) => `<div class="estcol"><h4>${lbl}: ${esc(est && est.patient != null ? F.usd(est.patient) : '—')}</h4>${est ? `<ul class="estlines">${F.calcText(est, b).map((t) => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}</div>`;
    out.innerHTML = `${e.flags.length ? `<ul class="pc-items">${e.flags.map((f) => `<li class="sev-medium">${esc(f)}</li>`).join('')}</ul>` : ''}
      ${e.basis ? `<p class="hint">${esc(e.basis)}</p>` : ''}
      ${e.md || e.np ? `<div class="estcols">${e.same ? col(b.planType === 'self_pay' ? 'Self pay' : 'MD and NP', e.md) : `${col('MD', e.md)}${col('NP', e.np)}`}</div>` : ''}`;
  }
  // ----- pricing services not yet booked (v0.23) -----
  // A service from Estimate Codes (or typed codes), repeated over a number of future visits, from today's benefits or
  // from where the booked visits leave them (the projection). Each code counts by its saved rule unless one is picked.
  /** Estimate Codes rows that have codes (interventions are "none"): [{i, label, codes}]. */
  const pricedServices = () => ((fees && fees.codeRows) || []).map((r, i) => ({ i, r, p: F.parseCodeList(r.Codes) }))
    .filter((x) => !x.p.none && x.p.codes.length).map((x) => ({ i: x.i, label: `${x.r['Event keyword']} (${x.p.codes.map((c) => `${c.code}${c.units > 1 ? ` x${c.units}` : ''}`).join(' + ')})`, codes: x.p.codes }));
  function serviceHtml(ctx) {
    const opts = pricedServices();
    return `<div class="filters calcgrid">
        <label>Service<select data-esvc="1"><option value=""${estOver.service === '' ? ' selected' : ''}>${esc(ctx.event ? `Booked: ${ctx.event}` : 'Codes below')}</option>
          ${opts.map((o) => `<option value="${o.i}"${String(o.i) === String(estOver.service) ? ' selected' : ''}>${esc(o.label)}</option>`).join('')}</select></label>
        <label>Visits<input type="number" min="1" max="52" step="1" data-evisits="1" value="${esc(estOver.visits)}"></label>
        <label>Start from<select data-estart="1"><option value="now"${estOver.start === 'now' ? ' selected' : ''}>Today's benefits</option>
          <option value="booked"${estOver.start === 'booked' ? ' selected' : ''}>After the booked visits (projection)</option></select></label>
      </div>`;
  }
  /** The rule picker on a code row: the saved rule (or the plan's charge mode), or a class for this estimate. */
  function ruleSelect(ctx, c, i) {
    const saved = (ctx.codeRules || {})[String(c.code || '').toUpperCase()];
    const cur = c.cls == null ? 'saved' : c.cls;
    const savedLabel = saved ? `Saved: ${F.ruleLabel(saved)}` : 'Plan charge mode';
    const save = ctx.ruleKey && c.code && ((cur !== 'saved' && cur !== 'mode') || (cur === 'mode' && saved))
      ? ` <button class="link" data-erulesave="${i}">${cur === 'mode' ? `Clear the saved rule for ${esc(ctx.ruleKey)}` : `Save for ${esc(ctx.ruleKey)}`}</button>` : '';
    return `<select data-erule="${i}" aria-label="How code ${i + 1} counts"><option value="saved"${cur === 'saved' ? ' selected' : ''}>${esc(savedLabel)}</option>
      ${saved ? `<option value="mode"${cur === 'mode' ? ' selected' : ''}>Plan charge mode</option>` : ''}
      ${F.CODE_CLASSES.map((k) => `<option value="${k.id}"${cur === k.id ? ' selected' : ''}>${esc(k.label)}</option>`).join('')}</select>${save}`;
  }
  /** The panel's codes for fees.js: rule undefined = the saved rule, '' = the plan's charge mode, else the class's toggles. */
  const calcCodes = (codes) => codes.map((c) => ({ ...c, rule: c.cls == null ? undefined : c.cls === 'mode' ? '' : (F.CODE_CLASSES.find((k) => k.id === c.cls) || {}).toggles }));
  /** The inputs, moved past the booked visits when "After the booked visits" is picked. */
  function startInputs(b) {
    if (estOver.start !== 'booked') return b;
    const c = detailKey ? patientCtx(detailKey) : null, p = c && c.proj;
    if (!p) return b;
    const add = (v, a) => (v === '' || v == null ? v : Math.round((Number(v) + (a || 0)) * 100) / 100);
    return { ...b, deductibleMet: add(b.deductibleMet, p.ded && p.ded.applied), oopMet: add(b.oopMet, p.oop && p.oop.applied) };
  }
  function seriesHtml(sr, b) {
    if (!sr.md && !sr.np) return `${sr.flags.length ? `<ul class="pc-items">${sr.flags.map((f) => `<li class="sev-medium">${esc(f)}</li>`).join('')}</ul>` : ''}`;
    const lv = (label, x) => {
      if (!x || x.none) return `<div class="estcol"><h4>${esc(label)}: —</h4><p class="hint">${esc((x && x.none) || 'No estimate')}</p></div>`;
      return `<div class="estcol"><h4>${esc(label)}: ${esc(F.usd(x.total))} over ${x.visits.length} visit${x.visits.length === 1 ? '' : 's'}</h4>
        <ul class="estlines">${x.visits.map((v) => `<li>Visit ${v.n}: ${esc(F.usd(v.patient))}${v.towardDeductible ? ` (${esc(F.usd(v.towardDeductible))} toward the deductible)` : ''}</li>`).join('')}</ul>
        <p class="hint">${esc([x.dedAt ? `Deductible met at visit ${x.dedAt}` : '', x.oopAt ? `OOP max met at visit ${x.oopAt}` : ''].filter(Boolean).join(' · ') || 'Neither the deductible nor the OOP max is met by these visits.')}</p></div>`;
    };
    return `${sr.flags.length ? `<ul class="pc-items">${sr.flags.map((f) => `<li class="sev-medium">${esc(f)}</li>`).join('')}</ul>` : ''}
      ${sr.basis ? `<p class="hint">${esc(sr.basis)}${estOver.start === 'booked' ? ' · starting after the booked visits' : ''}</p>` : ''}
      <div class="estcols">${sr.same ? lv(b.planType === 'self_pay' ? 'Self pay' : 'MD and NP', sr.md) : `${lv('MD', sr.md)}${lv('NP', sr.np)}`}</div>`;
  }
  $('detail').addEventListener('change', (e) => {
    const t = e.target;
    if (!estOver || !t.closest || !t.closest('#estPanel')) return;
    const ctx = fees && fees.contexts ? fees.contexts[estOver.per] : null;
    if (!ctx) return;
    if (t.dataset.esvc) {
      estOver.service = t.value;
      const svc = pricedServices().find((o) => String(o.i) === t.value);
      estOver.codes = (svc ? svc.codes : ctx.codes || []).map((c) => ({ ...c, coins: '' }));
      renderEstimatePanel(estOver.per, estOver.row);
    } else if (t.dataset.evisits) { estOver.visits = Math.max(1, Math.min(52, Math.floor(Number(t.value) || 1))); renderEstimateResults(); }
    else if (t.dataset.estart) { estOver.start = t.value; renderEstimateResults(); }
    else if (t.dataset.erule != null) {
      const c = estOver.codes[+t.dataset.erule];
      if (c) c.cls = t.value === 'saved' ? undefined : t.value;
      renderEstimatePanel(estOver.per, estOver.row);
    }
  });
  $('detail').addEventListener('click', (e) => {
    const b = e.target.closest('[data-erulesave]');
    if (!b || !estOver) return;
    const ctx = fees && fees.contexts ? fees.contexts[estOver.per] : null, c = estOver.codes[+b.dataset.erulesave];
    if (!ctx || !c || !ctx.ruleKey) return;
    const by = byOrAsk();
    if (!by) return;
    const cls = c.cls === 'mode' ? '' : c.cls;
    c.cls = undefined; // after the save, the saved rule is this one
    notice(`Saving how ${c.code} counts for ${ctx.ruleKey}…`, true);
    toPane({ type: 'fee-code-rule', ruleKey: ctx.ruleKey, code: String(c.code).toUpperCase(), cls, by });
  });
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
      if (ctx && ctx.feePayer) toPane({ type: 'fee-defaults', feePayer: ctx.feePayer, toggles: F.CALC_MODE_TOGGLES[s.value], by: adminName });
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
    } else if (kind === 'flow') {
      // v0.21: the Sohar Flow import (record.FLOW_IMPORT_COLS), built by the pane with the other imports.
      if (!exportsData || !(exportsData.flow || []).length) { notice('No Sohar Flow file built yet. Upload the NextGen report in the pane first.', false); return; }
      name = REC.flowImportFileName(); cols = REC.FLOW_IMPORT_COLS; data = exportsData.flow;
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
    if (tr) showDetail(tr.dataset.key, tr);
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('detail').hidden) { hideDetail(); return; } if (e.key === 'Enter') { const tr = e.target.closest && e.target.closest('tr[data-key]'); if (tr) showDetail(tr.dataset.key, tr); } });
  document.addEventListener('change', (e) => {
    const cb = e.target.closest && e.target.closest('input[data-task]');
    if (!cb) return;
    const key = cb.dataset.task;
    tasks = tasks.map((t) => (t['Task Key'] === key ? { ...t, Done: cb.checked ? 'Yes' : '' } : t));
    toPane({ type: 'task-done', key, done: cb.checked });
    document.querySelectorAll(`input[data-task="${CSS.escape(key)}"]`).forEach((x) => { x.checked = cb.checked; });
    $('taskCount').textContent = tasks.filter(open).length || '';
    renderTasks(); renderPrep(); renderMessage();
    if (!$('detail').hidden && $('detail').contains(cb)) renderDetail(`input[data-task="${CSS.escape(key)}"]`);
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
      'Policy vs NextGen': 'Same policy as NextGen', 'Estimate Policy': 'NextGen', Network: 'INN', 'Sohar Payer': 'UMR', 'Fee Tab': 'Marin', File: 'Flow Results.csv',
      Started: '2026-10-18T16:20:00Z', 'Last Updated': '2026-10-18T16:24:00Z', ...o });
    // Four booked visits: TMS (no estimate), two priced visits that reach the OOP max, and a sandbox one (left out).
    const demoFlow = [fl({ Key: 'd1', 'Appt Date': '10/20/2026', Time: '9:00 AM', Event: 'TMS Treatment Unilateral', 'Rate Level': 'NP', 'Est Codes': '', 'Estimate Note': 'TMS (intervention): no estimate; verify deductible, OOP max and insurance status' }),
      fl({ Key: 'd2', 'Appt Date': '10/27/2026', Time: '9:00 AM', Event: 'Med Management In-Office', 'Rate Level': 'NP', 'Est Codes': '99214 + 90836', 'Sohar Est Patient': '420.00', 'Sohar Est Plan Pays': '100.00' }),
      fl({ Key: 'd3', 'Appt Date': '11/03/2026', Time: '9:00 AM', Event: 'Med Management In-Office', 'Rate Level': 'NP', 'Est Codes': '99214 + 90836', 'Sohar Est Patient': '420.00', 'Sohar Est Plan Pays': '100.00' }),
      fl({ Key: 'd4', 'Appt Date': '11/10/2026', Time: '9:00 AM', Event: 'Med Management In-Office', 'Rate Level': 'NP', 'Est Codes': '99214', 'Sohar Est Patient': '71.11', 'Sohar Est Plan Pays': '100.00', File: 'Flow Results (SANDBOX).csv' })];
    const demoProfiles = { 100100: { pronouns: 'They, Them, Theirs', genderIdentity: '', currentGender: 'Female', guarantor: 'RIVANA (the patient)', checks: [
      { sentAs: 'Verification', payer: 'UMR', date: '2026-10-18', subscriber: 'RIVLUI', ...REC.guarantorCheck({ patient: { firstName: 'Ana', lastName: 'Rivera' }, guarantor: { firstName: 'Ana', lastName: 'Rivera' }, guarantorSelf: true }, { firstName: 'Luis', lastName: 'Rivera' }) }] } };
    // Report visits with Pre-Check Complete (made up): Rivera's first visit done in NextGen, Lindqvist's not.
    const dv = (key, date, time, event, loc, pc, ticks) => {
      const v = { date, time, event, loc, provider: '', status: '', precheck: pc, hasPrecheck: true, key };
      return { ...v, lasfir: '', items: REC.precheckItems(v, key, ticks || []) };
    };
    const demoVisits = {
      100100: [dv('80100', '10/20/2026', '9:00 AM', 'TMS Treatment Unilateral', 'San Rafael', 'Yes', [{ Key: '80100|10/20/2026|9:00 AM', Item: 'nextgen', Done: 'Yes', By: 'Jordan Lee', When: '10/19/2026 4:10 PM' }]),
        dv('80100', '10/27/2026', '9:00 AM', 'Med Management In-Office', 'San Rafael', ''), dv('80100', '11/03/2026', '9:00 AM', 'Med Management In-Office', 'San Rafael', '')],
      100102: [dv('PER-100102', '10/22/2026', '10:00 AM', 'Med Management In-Office', 'Walnut Creek', 'No')],
    };
    const demoPills = { 100100: 'NP-TMS', 100101: 'Medi-Cal; New Medi-Cal; Re-verify', 100102: 'Kaiser/VA MM no auth; Kaiser/VA MM', 100103: 'QMB' };
    for (const r of res) r.Pills = demoPills[r['Per Nbr']] || '';
    // ?fees=1: fee estimates with the made-up demo fee schedule (fees.js), as the pane would send them.
    let demoFees = null;
    if (F && /[?&]fees=1\b/.test(location.search)) {
      const s = F.demoSchedule();
      // UMR is mapped on Marin (as another patient's visit would have done) and 99214 is classified copay only for it.
      const payerMap = [{ Payer: 'UMR', 'Fee Tab': 'Marin', 'Fee Payer': 'UMR', Confidence: 'High', Source: 'manual', 'Auto Match': '' }];
      const defaults = F.rememberCodeRule([], 'UMR', '99214', F.CODE_CLASSES.find((k) => k.id === 'copay_waived').toggles, '10/19/2026 4:15 PM', 'Jordan Lee');
      const prep = F.prepareEstimates({ results: res, scheduleRows: s.rows, tabSummary: s.tabs, appts: [], teleKeywords: ['Tele'], cfgDefaults: F.TOGGLE_DEFAULTS, when: '', payerMap, defaults });
      demoFees = { on: true, status: F.feeStatus(F.feeMetaRows('Demo fee schedule (made-up rates).xlsx', new Date().toISOString(), s), s.rows.length, new Date()), contexts: prep.contexts, tables: prep.tables, toggles: F.TOGGLES,
        codeRows: prep.codeRows.map((r) => ({ 'Event keyword': r['Event keyword'], Codes: r.Codes, Note: r.Note })) };
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
      notes: { 80100: [{ note: 'Called the patient; new card scanned, plan unchanged', by: 'Jordan Lee', when: '10/19/2026 4:12 PM', visit: '' }] },
      lastUpdated: { report: new Date(Date.now() - 2 * 36e5).toISOString(), sohar: new Date(Date.now() - 3 * 864e5).toISOString(), flow: new Date().toISOString() } });
    // #patient opens the first demo patient's record (for checking the layout), #patient/<section> a section of it;
    // other hashes pick a tab.
    // #patient/visits/2 also opens the third visit row.
    const ph = location.hash.match(/^#patient(?:\/(\w+))?(?:\/(\d+))?$/);
    if (ph) {
      const dk = `${keyOf(res[0])}|${res[0]['Sohar Ref']}`;
      showDetail(dk, document.querySelector(`#prepList tr[data-key="${CSS.escape(dk)}"]`));
      if (ph[2]) uiOf(keyOf(res[0])).open = +ph[2];
      if (ph[1] && SECTIONS.some(([k]) => k === ph[1])) { sec = ph[1]; renderDetail(); }
      // ?fees=1&svc=med management&visits=4&start=booked#patient/calc prices a service in the calculator.
      const qs = new URLSearchParams(location.search);
      if (sec === 'calc' && estOver && qs.get('svc')) {
        const svc = pricedServices().find((o) => fees.codeRows[o.i]['Event keyword'] === qs.get('svc'));
        if (svc) { estOver.service = String(svc.i); estOver.codes = svc.codes.map((c) => ({ ...c, coins: '' })); }
        estOver.visits = Math.max(1, Math.min(52, +qs.get('visits') || 1)); estOver.start = qs.get('start') === 'booked' ? 'booked' : 'now';
        renderEstimatePanel(estOver.per, estOver.row);
      }
    }
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
