/* dashboard.js — full-screen worklist. Runs in an Office dialog; data arrives from the pane via messageChild. */
(function () {
  'use strict';
  const C = window.SoharCore;
  const $ = (id) => document.getElementById(id);
  let rows = [], asOf = null;
  const chunks = [];

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

  function onParent(arg) {
    let m; try { m = JSON.parse(arg.message); } catch (e) { return; }
    if (m.type !== 'chunk') return;
    chunks[m.i] = m.data;
    if (chunks.filter((x) => x != null).length < m.n) return;
    const data = JSON.parse(chunks.join(''));
    load(data.results || [], data.asOf ? new Date(data.asOf) : null);
  }

  function load(r, d) {
    rows = r; asOf = d;
    $('dStatus').textContent = rows.length ? `${rows.length} result rows${asOf ? `, appointments from ${C.fmtDate(asOf)}` : ''}.` : 'No Sohar results yet. Upload them in step 5 of the pane, then open full screen again.';
    $('dStatus').className = `status ${rows.length ? 'ok' : 'bad'}`;
    const q = (name) => rows.filter((x) => x['Work Queue'] === name).length;
    $('tAct').textContent = q('Action needed'); $('tUnm').textContent = q('Unmatched');
    $('tPen').textContent = q('Pending'); $('tClr').textContent = q('Clear');
    $('tDed').textContent = rows.filter((x) => x['Deductible Met'] === 'Yes').length;
    $('tOop').textContent = rows.filter((x) => x['OOP Met'] === 'Yes').length;
    $('benNote').hidden = !rows.length || rows.some((x) => x['Deductible Met'] || x['OOP Met']);
    const locs = [...new Set(rows.map((x) => x['Next Location']).filter(Boolean))].sort();
    $('fLoc').innerHTML = '<option value="">All</option>' + locs.map((l) => `<option>${esc(l)}</option>`).join('');
    render();
  }

  function filtered() {
    const qv = $('fQueue').value, bv = $('fBen').value, lv = $('fLoc').value, tv = $('fText').value.trim().toLowerCase();
    const from = $('fFrom').value ? C.parseDate($('fFrom').value) : null;
    const to = $('fTo').value ? C.parseDate($('fTo').value) : null;
    return rows.filter((r) => {
      if (qv && r['Work Queue'] !== qv) return false;
      if (lv && r['Next Location'] !== lv) return false;
      const ded = r['Deductible Met'], oop = r['OOP Met'];
      if (bv === 'ded' && ded !== 'Yes') return false;
      if (bv === 'oop' && oop !== 'Yes') return false;
      if (bv === 'either' && ded !== 'Yes' && oop !== 'Yes') return false;
      if (bv === 'notmet' && !(ded !== 'Yes' && oop !== 'Yes' && (ded === 'No' || oop === 'No'))) return false;
      if (bv === 'unknown' && (ded || oop)) return false;
      if (from || to) {
        const d = C.parseDate(r['Next Appt']);
        if (!d || (from && d < from) || (to && d > to)) return false;
      }
      if (tv && ![r['Last Name'], r['First Name'], r['Per Nbr'], r['NextGen Payer'], r['Sohar Payer']].some((v) => String(v || '').toLowerCase().includes(tv))) return false;
      return true;
    });
  }

  function render() {
    const list = filtered();
    $('shown').textContent = `Showing ${list.length} of ${rows.length}`;
    const cell = (v) => `<td>${esc(v)}</td>`;
    $('tbl').querySelector('tbody').innerHTML = list.map((r) => `<tr class="q-${queueClass(r['Work Queue'])}">
      ${cell([r['Next Appt'], r['Next Time']].filter(Boolean).join(' '))}${cell(r['Next Location'])}${cell(r['Next Provider'])}
      ${cell([r['Last Name'], r['First Name']].filter(Boolean).join(', '))}${cell(r['Per Nbr'])}${cell(r['Sent As'])}
      ${cell(r['Sohar Status'])}${cell(r['Work Queue'])}<td class="wide">${esc(r.Action)}</td>
      ${cell(r['Deductible Met'])}${cell(r['Deductible Remaining'])}${cell(r['OOP Met'])}${cell(r['OOP Remaining'])}
      ${cell(r['NextGen Payer'])}${cell(r['Sohar Payer'])}
      <td>${/^https:\/\//.test(r['Sohar Link'] || '') ? `<a href="${esc(r['Sohar Link'])}" target="_blank" rel="noopener">Open</a>` : ''}</td></tr>`).join('');
    $('msg').value = C.adminMessage(list, { asOf });
  }

  function queueClass(q) { return { 'Action needed': 'act', Unmatched: 'unm', Pending: 'pen', Clear: 'clr' }[q] || ''; }

  ['fQueue', 'fBen', 'fLoc', 'fFrom', 'fTo'].forEach((id) => $(id).addEventListener('change', render));
  $('fText').addEventListener('input', render);
  document.querySelectorAll('[data-queue]').forEach((b) => b.addEventListener('click', () => { $('fQueue').value = b.dataset.queue; $('fBen').value = ''; render(); }));
  document.querySelectorAll('[data-benefit]').forEach((b) => b.addEventListener('click', () => { $('fBen').value = b.dataset.benefit; $('fQueue').value = ''; render(); }));

  $('btnCopy').addEventListener('click', async () => {
    const ta = $('msg');
    let ok = false;
    try { await navigator.clipboard.writeText(ta.value); ok = true; } catch (e) {
      ta.focus(); ta.select(); try { ok = document.execCommand('copy'); } catch (e2) { ok = false; }
    }
    $('dStatus').textContent = ok ? 'Admin message copied. Paste it into Outlook or Teams.' : 'Copy was blocked. Select the message box text and press Ctrl+C.';
  });
  $('btnCsv').addEventListener('click', () => {
    const csv = C.toCSV(C.RESULTS_COLS, filtered());
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = `sohar_results_filtered_${new Date().toISOString().slice(0, 10).replace(/-/g, '')}.csv`;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  });

  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

  // Browser preview outside Office: synthetic rows only.
  function demo() {
    const b = document.createElement('div');
    b.className = 'banner';
    b.textContent = 'Demo mode. Synthetic data only.';
    document.body.insertBefore(b, document.querySelector('main'));
    const base = { 'Upcoming Appts': 1, 'Sohar payerId': '', 'Sohar Member ID': '', 'Coverage Differs': '', 'Sohar Ref': '', 'Sohar Link': '', 'Sohar Date': '', 'Deductible': '', 'OOP Max': '', 'Copay': '', 'Coinsurance': '', 'DOB': '', 'NextGen Member ID': '' };
    const mk = (o) => ({ ...base, ...o });
    load([
      mk({ 'Next Appt': '10/20/2026', 'Next Time': '9:00 AM', 'Next Event': 'Telepsych', 'Next Location': 'San Rafael', 'Next Provider': 'Dr Test', 'Per Nbr': '100100', 'Last Name': 'Rivera', 'First Name': 'Ana', 'Sent As': 'Verification', 'Sohar Status': 'complete.ineligible', 'Work Queue': 'Action needed', Action: 'Ineligible. Confirm coverage with the patient before the visit; consider discovery.', 'Deductible Met': 'Yes', 'Deductible Remaining': '0', 'OOP Met': 'No', 'OOP Remaining': '1200', 'NextGen Payer': 'Blue Shield CA', 'Sohar Payer': 'Blue Shield of California', Match: 'Member ID' }),
      mk({ 'Next Appt': '10/21/2026', 'Next Time': '1:30 PM', 'Next Event': 'TMS Treatment Unilateral', 'Next Location': 'San Rafael', 'Per Nbr': '100101', 'Last Name': 'Okafor', 'First Name': 'Chidi', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Work Queue': 'Action needed', Action: 'Coverage found. Add it to NextGen, then verify.', 'Deductible Met': '', 'OOP Met': '', 'NextGen Payer': '', 'Sohar Payer': 'Cigna', Match: 'Per Nbr', 'Sohar Link': 'https://example.test/r-1' }),
      mk({ 'Next Appt': '10/22/2026', 'Next Time': '10:00 AM', 'Next Location': 'Walnut Creek', 'Per Nbr': '100102', 'Last Name': 'Lindqvist', 'First Name': 'Sven', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Work Queue': 'Clear', Action: 'Eligible. No action.', 'Deductible Met': 'Yes', 'OOP Met': 'Yes', 'OOP Remaining': '0', 'NextGen Payer': 'Aetna', 'Sohar Payer': 'Aetna', Match: 'Member ID' }),
      mk({ 'Next Appt': '', 'Per Nbr': '', 'Last Name': 'Synthetic', 'First Name': 'Nobody', 'Sent As': 'Verification', 'Sohar Status': 'pending', 'Work Queue': 'Unmatched', Action: 'Not matched to a patient in the Import report. Pending at Sohar. Recheck later.', 'Upcoming Appts': '', Match: 'No match' }),
    ], new Date(Date.UTC(2026, 9, 13)));
  }
})();
