/* config-ui.js — the Config section (v0.31), drawn the same way in the pane (Admin tab, at the bottom) and in full screen
   (Config tab). UI only: the model comes from core.configModel and every edit goes to the pane, which checks it
   (core.configEdit) and writes it with excel.js setConfigValue. No network, no workbook access here. */
(function (root) {
  'use strict';
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const RULES = ['Window', 'Weekly'];
  const onOff = (setting, on, label) => `<label class="cfg-switch"><input type="checkbox" data-cfg="${esc(setting)}" ${on ? 'checked' : ''}> ${esc(label)}</label>`;
  function ruleSelect(l) {
    const days = RULES.includes(l.rule) ? '' : l.rule;
    return `<select data-cfg-rule="${esc(l.ruleSetting)}" aria-label="Re-verify rule for ${esc(l.line)}">
      ${RULES.map((r) => `<option value="${r}"${l.rule === r ? ' selected' : ''}>${r === 'Window' ? 'Window (and the 1st)' : 'Weekly, from Monday'}</option>`).join('')}
      <option value="days"${days ? ' selected' : ''}>Days</option></select>
      <input type="number" min="0" step="1" class="cfg-days" data-cfg-days="${esc(l.ruleSetting)}" value="${esc(days)}" aria-label="Days for ${esc(l.line)}"${days ? '' : ' hidden'}>`;
  }
  /** One list for Fee Payer Map or Estimate Codes: read-only, filterable, at most `max` rows shown. */
  function lookupTable(id, title, cols, rows, sheet, max) {
    const list = rows || [];
    const shown = list.slice(0, max || 150);
    return `<details class="more cfg-look" data-look="${esc(id)}"><summary>${esc(title)} (${list.length})</summary>
      <div class="cfg-lookbar"><input type="search" data-look-filter="${esc(id)}" placeholder="Filter" aria-label="Filter ${esc(title)}">
      <button class="link" data-goto="${esc(sheet)}">Edit on the ${esc(sheet)} sheet</button></div>
      ${list.length ? `<div class="cfg-scroll"><table class="watch"><thead><tr>${cols.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>
      ${shown.map((r) => `<tr>${cols.map((c) => `<td>${esc(r[c])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
      ${list.length > shown.length ? `<p class="hint">First ${shown.length} of ${list.length}; the sheet has them all.</p>` : ''}` : '<p class="hint">No rows yet.</p>'}
      </details>`;
  }
  /**
   * The Config section's HTML. model: core.configModel. look: {feePayerMap, estimateCodes, feesOn} (rows without rates).
   * The payer mapping list goes in the element with data-cfg-payers (the pane fills it; full screen draws its own).
   */
  function html(model, look) {
    const m = model || { lines: [], extra: [] };
    const lk = look || {};
    return `<div class="cfg">
      <h3>Re-verify rules</h3>
      <p class="hint">A visit is covered by an eligible verification of the same insurance made no earlier than the window before it, and never before the 1st of the visit's month (00:00 Pacific). Covered visits leave the verification file; the 7 AM rotation fills the gaps with the same rule.</p>
      <div class="cfg-row">${onOff('Skip recent verifications', m.skip, 'Leave covered visits out of the verification file')}</div>
      <div class="cfg-row"><label>Window (hours) <input type="number" min="0" step="1" class="cfg-days" data-cfg="Re-verify window (hours)" value="${esc(m.windowHours)}"></label>
        ${onOff('Re-verify on the 1st', m.onFirst, 'Never before the 1st of the month')}</div>
      <table class="watch cfg-lines"><thead><tr><th>Service line</th><th>Rule</th><th>In the next import</th></tr></thead><tbody>
      ${(m.lines || []).map((l) => `<tr><td>${esc(l.line)}</td><td>${ruleSelect(l)}</td><td>${onOff(l.includeSetting, l.include, l.include ? 'Included' : 'Left out')}</td></tr>`).join('')}
      </tbody></table>
      <p class="hint">Provider covers consults, follow-ups, TMS F/U and MT / MTR. Left out: that line leaves the import files at the next build, and the rotation still covers it.</p>
      ${(m.extra || []).length ? `<h3>Plan-year estimate</h3>${m.extra.map((r) => `<div class="cfg-row"><label>${esc(r.setting)} <input type="text" data-cfg="${esc(r.setting)}" value="${esc(r.value)}"></label></div><p class="hint">${esc(r.what)}</p>`).join('')}` : ''}
      <h3>Payer mapping</h3>
      <div data-cfg-payers></div>
      <h3>Fee lookups</h3>
      ${lk.feesOn ? '' : '<p class="hint">Config <b>Fee estimates</b> is Off, so the fee sheets are not read.</p>'}
      ${lookupTable('feemap', 'Fee Payer Map', ['Payer', 'Sohar payerId', 'Fee Tab', 'Fee Payer', 'Confidence', 'Source'], lk.feePayerMap, 'Fee Payer Map')}
      ${lookupTable('codes', 'Estimate Codes', ['Event keyword', 'Codes', 'Note'], lk.estimateCodes, 'Estimate Codes')}
      <p class="cfg-msg hint" role="status" aria-live="polite"></p>
    </div>`;
  }
  /**
   * Wire the inputs in `box`: each change calls send(setting, value). goto(sheet) for the edit links. Filters run here.
   */
  function wire(box, send, goto) {
    if (!box || box.dataset.cfgWired) return;
    box.dataset.cfgWired = '1';
    box.addEventListener('change', (e) => {
      const t = e.target;
      if (t.matches('input[type="checkbox"][data-cfg]')) send(t.dataset.cfg, t.checked ? 'On' : 'Off');
      else if (t.matches('[data-cfg]')) send(t.dataset.cfg, t.value);
      else if (t.matches('[data-cfg-rule]')) {
        const days = box.querySelector(`[data-cfg-days="${CSS.escape(t.dataset.cfgRule)}"]`);
        if (t.value === 'days') { if (days) { days.hidden = false; days.focus(); } return; }
        if (days) days.hidden = true;
        send(t.dataset.cfgRule, t.value);
      } else if (t.matches('[data-cfg-days]')) send(t.dataset.cfgDays, t.value);
    });
    box.addEventListener('input', (e) => {
      const f = e.target.closest('[data-look-filter]');
      if (!f) return;
      const q = f.value.trim().toLowerCase();
      const look = f.closest('[data-look]');
      look.querySelectorAll('tbody tr').forEach((tr) => { tr.hidden = !!q && !tr.textContent.toLowerCase().includes(q); });
    });
    box.addEventListener('click', (e) => { const b = e.target.closest('[data-goto]'); if (b && box.contains(b) && goto) { e.preventDefault(); goto(b.dataset.goto); } });
  }
  function say(box, text, ok) {
    const p = box && box.querySelector('.cfg-msg');
    if (p) { p.textContent = text || ''; p.className = `cfg-msg ${ok === false ? 'warn' : 'hint'}`; }
  }
  root.SoharConfigUI = { html, wire, say };
})(typeof self !== 'undefined' ? self : this);
