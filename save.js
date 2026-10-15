/* save.js — fallback save tab. Excel on the web can block downloads inside the add-in frame; this top-level tab
   (same site as the add-in) receives the file from the pane by postMessage and saves it. No network calls.
   Accepts {type:'sohar-save', name, csv} or {type:'sohar-save', name, base64, mime} (clinic .xlsx files). */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  let file = null;

  function blobOf(f) {
    if (typeof f.csv === 'string') return new Blob([f.csv], { type: 'text/csv;charset=utf-8' });
    const bin = atob(f.base64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new Blob([bytes], { type: f.mime || 'application/octet-stream' });
  }

  function save() {
    if (!file) return;
    const url = URL.createObjectURL(blobOf(file));
    const a = document.createElement('a');
    a.href = url; a.download = file.name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    $('saveMsg').textContent = `Saved ${file.name} to your Downloads folder. You can close this tab.`;
    $('saveMsg').className = 'status ok';
    $('saveAgain').hidden = false;
  }

  window.addEventListener('message', (e) => {
    if (e.origin !== location.origin) return;
    const m = e.data;
    if (!m || m.type !== 'sohar-save' || typeof m.name !== 'string') return;
    const isCsv = typeof m.csv === 'string';
    if (!isCsv && typeof m.base64 !== 'string') return;
    if (e.source) e.source.postMessage({ type: 'sohar-save-got' }, location.origin);
    if (file && file.name === m.name) return; // the pane repeats until it hears back
    file = { name: m.name.replace(/[\\/:*?"<>|]/g, '-'), csv: isCsv ? m.csv : null, base64: isCsv ? null : m.base64, mime: typeof m.mime === 'string' ? m.mime : '' };
    try { save(); } catch (err) { $('saveMsg').textContent = `Could not save ${file.name}: ${err.message}`; $('saveMsg').className = 'status bad'; }
  });
  $('saveAgain').addEventListener('click', save);
  if (window.opener) window.opener.postMessage({ type: 'sohar-save-ready' }, location.origin);
})();
