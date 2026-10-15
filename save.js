/* save.js — fallback save tab. Excel on the web can block downloads inside the add-in frame; this top-level tab
   (same site as the add-in) receives the CSV from the pane by postMessage and saves it. No network calls. */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  let file = null;

  function save() {
    if (!file) return;
    const url = URL.createObjectURL(new Blob([file.csv], { type: 'text/csv;charset=utf-8' }));
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
    if (!m || m.type !== 'sohar-save' || typeof m.csv !== 'string' || typeof m.name !== 'string') return;
    if (e.source) e.source.postMessage({ type: 'sohar-save-got' }, location.origin);
    if (file && file.name === m.name) return; // the pane repeats until it hears back
    file = { name: m.name.replace(/[\\/:*?"<>|]/g, '-'), csv: m.csv };
    save();
  });
  $('saveAgain').addEventListener('click', save);
  if (window.opener) window.opener.postMessage({ type: 'sohar-save-ready' }, location.origin);
})();
