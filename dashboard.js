/* dashboard.js — full-screen worklist. Runs in an Office dialog; data arrives from the pane via messageChild. */
(function () {
  'use strict';
  const C = window.SoharCore;
  const $ = (id) => document.getElementById(id);
  let rows = [], asOf = null, list = [];
  const chunks = [];

  function onParent(arg) {
    let m; try { m = JSON.parse(arg.message); } catch (e) { return; }
    if (m.type !== 'chunk') return;
    chunks[m.i] = m.data;
    if (chunks.filter((x) => x != null).length < m.n) return;
    const data = JSON.parse(chunks.join(''));
    load(data.results || [], data.asOf ? new Date(data.asOf) : null);
  }

  const BEN = {
    dedMet: (r) => r['Deductible Status'] === 'Met',
    dedAppr: (r) => r['Deductible Status'] === 'Approaching',
    oopMet: (r) => r['OOP Status'] === 'Met',
    oopAppr: (r) => r['OOP Status'] === 'Approaching',
    anyFlag: (r) => /met|approaching/i.test(r['Benefit Flag'] || ''),
    noDed: (r) => r['Deductible Status'] === 'None',
    unknown: (r) => !r['Deductible Status'] && !r['OOP Status'],
  };

  function load(r, d) {
    rows = r; asOf = d;
    $('dStatus').textContent = rows.length ? `${rows.length} result rows${asOf ? `, appointments from ${C.fmtDate(asOf)}` : ''}. Click a row for the patient view.` : 'No Sohar results yet. Upload them in step 5 of the pane, then open full screen again.';
    $('dStatus').className = `status ${rows.length ? 'ok' : 'bad'}`;
    const q = (name) => rows.filter((x) => x['Work Queue'] === name).length;
    $('tAct').textContent = q('Action needed'); $('tUnm').textContent = q('Unmatched');
    $('tPen').textContent = q('Pending'); $('tClr').textContent = q('Clear');
    $('tDedMet').textContent = rows.filter(BEN.dedMet).length; $('tDedAppr').textContent = rows.filter(BEN.dedAppr).length;
    $('tOopMet').textContent = rows.filter(BEN.oopMet).length; $('tOopAppr').textContent = rows.filter(BEN.oopAppr).length;
    $('benNote').hidden = !rows.length || rows.some((x) => x['Deductible Status'] || x['OOP Status']);
    const locs = [...new Set(rows.map((x) => x['Next Location']).filter(Boolean))].sort();
    $('fLoc').innerHTML = '<option value="">All</option>' + locs.map((l) => `<option>${esc(l)}</option>`).join('');
    render();
  }

  function filtered() {
    const qv = $('fQueue').value, bv = $('fBen').value, sv = $('fSent').value, lv = $('fLoc').value, tv = $('fText').value.trim().toLowerCase();
    const from = $('fFrom').value ? C.parseDate($('fFrom').value) : null;
    const to = $('fTo').value ? C.parseDate($('fTo').value) : null;
    return rows.filter((r) => {
      if (qv && r['Work Queue'] !== qv) return false;
      if (sv && r['Sent As'] !== sv) return false;
      if (lv && r['Next Location'] !== lv) return false;
      if (bv && !BEN[bv](r)) return false;
      if (from || to) {
        const d = C.parseDate(r['Next Appt']);
        if (!d || (from && d < from) || (to && d > to)) return false;
      }
      if (tv && !['Last Name', 'First Name', 'Per Nbr', 'MRN', 'NextGen Payer', 'Sohar Payer', 'Plan', 'Choose in NextGen'].some((k) => String(r[k] || '').toLowerCase().includes(tv))) return false;
      return true;
    });
  }

  const left = (rem, total) => (rem === '' || rem == null ? '' : total !== '' && total != null ? `${rem} of ${total}` : String(rem));

  function render() {
    list = filtered();
    $('shown').textContent = `Showing ${list.length} of ${rows.length}`;
    const cell = (v, cls) => `<td${cls ? ` class="${cls}"` : ''}>${esc(v)}</td>`;
    const st = (s) => ({ Met: 'met', Approaching: 'appr' }[s] || '');
    $('tbl').querySelector('tbody').innerHTML = list.map((r, i) => `<tr class="q-${queueClass(r['Work Queue'])}" data-i="${i}" tabindex="0">
      ${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r['Next Location'])}
      ${cell([r['Last Name'], r['First Name']].filter(Boolean).join(', '))}${cell(r['Per Nbr'])}${cell(r['Sent As'])}
      ${cell(r['Sohar Status'])}${cell(r['Work Queue'])}${cell(r.Action, 'wide')}${cell(r['Choose in NextGen'], 'mid')}
      ${cell(r['Benefit Flag'])}${cell(r.Copay)}${cell(r.Coinsurance)}
      ${cell(left(r['Deductible Remaining'], r.Deductible), st(r['Deductible Status']))}${cell(left(r['OOP Remaining'], r['OOP Max']), st(r['OOP Status']))}
      ${cell(r.Plan, 'mid')}<td>${link(r['Sohar Link'])}</td></tr>`).join('');
    $('msg').value = C.adminMessage(list, { asOf, mode: $('msgMode').value });
  }

  function link(u) { return /^https:\/\//.test(u || '') ? `<a href="${esc(u)}" target="_blank" rel="noopener">Open</a>` : ''; }

  // Patient view: everything we know for one row, grouped for visit prep.
  function showDetail(r) {
    const box = $('detail');
    const grp = (title, keys) => `<div class="dgroup"><h3>${title}</h3><dl>${keys.filter((k) => r[k] !== '' && r[k] != null).map((k) => `<dt>${esc(k)}</dt><dd>${k === 'Sohar Link' ? link(r[k]) : esc(r[k])}</dd>`).join('') || '<dd class="hint">None</dd>'}</dl></div>`;
    box.innerHTML = `<div class="dhead"><h2>${esc([r['Last Name'], r['First Name']].filter(Boolean).join(', ') || 'Unmatched result')}</h2>
        <span class="hint">Per Nbr ${esc(r['Per Nbr'] || '?')}${r.MRN ? ` · MRN ${esc(r.MRN)}` : ''}${r.DOB ? ` · DOB ${esc(r.DOB)}` : ''}</span>
        <button class="small" id="btnHide">Close</button></div>
      <div class="dgrid">
        ${grp('Upcoming visits', ['Next Appt', 'Next Time', 'Next Event', 'Next Location', 'Next Provider', 'Upcoming Appts', 'Upcoming Detail'])}
        ${grp('What to do', ['Work Queue', 'Action', 'Choose in NextGen', 'Sohar Status', 'Match', 'Coverage Differs'])}
        ${grp('Benefits (in-network)', ['Benefit Flag', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status', 'Coverage Level'])}
        ${grp('Coverage', ['NextGen Payer', 'NextGen Member ID', 'Sohar Payer', 'Sohar payerId', 'Sohar Member ID', 'Plan', 'Network', 'Coverage Start', 'Coverage End', 'Sent As', 'Sohar Date', 'Sohar Link'])}
      </div>`;
    box.hidden = false;
    $('btnHide').addEventListener('click', () => { box.hidden = true; });
    box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function queueClass(q) { return { 'Action needed': 'act', Unmatched: 'unm', Pending: 'pen', Clear: 'clr' }[q] || ''; }

  ['fQueue', 'fBen', 'fSent', 'fLoc', 'fFrom', 'fTo', 'msgMode'].forEach((id) => $(id).addEventListener('change', render));
  $('fText').addEventListener('input', render);
  const tbody = $('tbl').querySelector('tbody');
  tbody.addEventListener('click', (e) => { if (e.target.closest('a')) return; const tr = e.target.closest('tr'); if (tr) showDetail(list[+tr.dataset.i]); });
  tbody.addEventListener('keydown', (e) => { if (e.key === 'Enter') { const tr = e.target.closest('tr'); if (tr) showDetail(list[+tr.dataset.i]); } });
  document.querySelectorAll('[data-queue]').forEach((b) => b.addEventListener('click', () => { $('fQueue').value = b.dataset.queue; $('fBen').value = ''; $('msgMode').value = 'action'; render(); }));
  document.querySelectorAll('[data-benefit]').forEach((b) => b.addEventListener('click', () => { $('fBen').value = b.dataset.benefit; $('fQueue').value = ''; $('msgMode').value = 'benefits'; render(); }));

  $('btnCopy').addEventListener('click', async () => {
    const ta = $('msg');
    let ok = false;
    try { await navigator.clipboard.writeText(ta.value); ok = true; } catch (e) {
      ta.focus(); ta.select(); try { ok = document.execCommand('copy'); } catch (e2) { ok = false; }
    }
    $('dStatus').textContent = ok ? 'Message copied. Paste it into Outlook or Teams.' : 'Copy was blocked. Select the message box text and press Ctrl+C.';
  });
  const stamp = () => new Date().toISOString().slice(0, 10).replace(/-/g, '');
  $('btnCsv').addEventListener('click', () => save(`sohar_results_filtered_${stamp()}.csv`, C.toCSV(C.RESULTS_COLS, filtered())));
  $('btnBen').addEventListener('click', () => save(`sohar_benefits_${stamp()}.csv`, C.toCSV(C.BENEFITS_EXPORT_COLS, C.benefitsExport(rows))));
  function save(name, csv) {
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

  // Browser preview outside Office: synthetic rows only.
  function demo() {
    const b = document.createElement('div');
    b.className = 'banner';
    b.textContent = 'Demo mode. Synthetic data only.';
    document.body.insertBefore(b, document.querySelector('main'));
    const mk = (o) => { const r = Object.fromEntries(C.RESULTS_COLS.map((c) => [c, ''])); return Object.assign(r, o); };
    load([
      mk({ 'Next Appt': '10/20/2026', 'Next Time': '9:00 AM', 'Next Event': 'TMS Treatment Unilateral', 'Next Location': 'San Rafael', 'Next Provider': 'Dr Test', 'Upcoming Appts': 12, 'Upcoming Detail': '10/20 TMS Treatment Unilateral; 10/21 TMS Treatment Unilateral; +10 more', 'Per Nbr': '100100', MRN: '80100', 'Last Name': 'Rivera', 'First Name': 'Ana', DOB: '03/04/1980', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Work Queue': 'Clear', Action: 'Eligible. No action.', 'Benefit Flag': 'Deductible met; OOP max approaching', Plan: 'SYNTH GOLD PPO', Network: 'PR', 'Coverage Level': 'IND', Copay: '40', Coinsurance: '20%', Deductible: '1500', 'Deductible Remaining': '0', 'Deductible Status': 'Met', 'OOP Max': '5000', 'OOP Remaining': '800', 'OOP Status': 'Approaching', 'NextGen Payer': 'Blue Shield CA', 'Sohar Payer': 'Blue Shield of California', Match: 'Per Nbr' }),
      mk({ 'Next Appt': '10/21/2026', 'Next Time': '1:30 PM', 'Next Event': 'Telepsych', 'Next Location': 'San Rafael', 'Upcoming Appts': 1, 'Per Nbr': '100101', 'Last Name': 'Okafor', 'First Name': 'Chidi', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Work Queue': 'Action needed', Action: 'Coverage found. Add it to NextGen, then verify. Choose "Cigna" in NextGen. Get the member ID from the Sohar link.', 'Choose in NextGen': 'Cigna', 'Sohar Payer': 'Cigna', 'Sohar payerId': '62308', Match: 'Per Nbr', 'Sohar Link': 'https://example.test/r-1' }),
      mk({ 'Next Appt': '10/22/2026', 'Next Time': '10:00 AM', 'Next Location': 'Walnut Creek', 'Upcoming Appts': 2, 'Per Nbr': '100102', 'Last Name': 'Lindqvist', 'First Name': 'Sven', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Work Queue': 'Action needed', Action: 'Different coverage found. Confirm with the patient, update NextGen, then verify. Choose one of "Optum Health Behavioral Solutions", "UHC Commercial" in NextGen; match the plan type on the card.', 'Choose in NextGen': 'One of: Optum Health Behavioral Solutions | UHC Commercial', 'NextGen Payer': 'Self Pay', 'Sohar Payer': 'UnitedHealthcare', Match: 'Per Nbr' }),
      mk({ 'Next Appt': '10/23/2026', 'Next Time': '11:00 AM', 'Next Location': 'Walnut Creek', 'Upcoming Appts': 4, 'Per Nbr': '100103', 'Last Name': 'Mensah', 'First Name': 'Afia', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Work Queue': 'Clear', Action: 'Eligible. No action.', 'Benefit Flag': 'OOP max met', Copay: '0', Coinsurance: '0%', Deductible: '0', 'Deductible Remaining': '0', 'Deductible Status': 'None', 'OOP Max': '3000', 'OOP Remaining': '0', 'OOP Status': 'Met', 'NextGen Payer': 'Cigna', 'Sohar Payer': 'Cigna', Plan: 'OPEN ACCESS PLUS', Match: 'Per Nbr' }),
      mk({ 'Per Nbr': '', 'Last Name': '', 'First Name': '', 'Sent As': 'Verification', 'Sohar Status': 'error.member.id', 'Work Queue': 'Unmatched', Action: 'Not matched to a patient in the Import report. Member ID rejected. Check the card, fix in NextGen, re-verify.', Match: 'No match' }),
    ], new Date(Date.UTC(2026, 9, 13)));
  }
  // Boot last, so every const above is initialized before demo() or the first message can use it.
  if (window.Office && Office.onReady) {
    Office.onReady(() => {
      if (Office.context && Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParent, () => {
          Office.context.ui.messageParent(JSON.stringify({ type: 'ready' }));
        });
        $('btnClose').hidden = false;
        $('btnClose').addEventListener('click', () => Office.context.ui.messageParent(JSON.stringify({ type: 'close' })));
      } else demo();
    });
  } else demo();
})();
