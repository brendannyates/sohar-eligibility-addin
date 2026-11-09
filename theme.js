/* Light or dark (v0.28.2), loaded in <head> so the page never flashes the wrong one. The choice is kept per computer
   (localStorage "soharTheme"); with none saved it follows Windows. The pane adds ?theme= to the full-screen URL and sends
   { type: 'theme' } when it changes, because the dialog may not share the pane's storage. Sets data-theme on <html>. */
(function () {
  'use strict';
  const KEY = 'soharTheme';
  const ok = (t) => t === 'dark' || t === 'light';
  const saved = () => { try { const t = localStorage.getItem(KEY); return ok(t) ? t : ''; } catch (e) { return ''; } };
  const os = () => { try { return window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'; } catch (e) { return 'light'; } };
  const fromUrl = (() => { try { const t = new URLSearchParams(location.search).get('theme'); return ok(t) ? t : ''; } catch (e) { return ''; } })();
  let cur = fromUrl || saved() || os();
  document.documentElement.dataset.theme = cur;
  const SUN = '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
  const MOON = '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/></svg>';
  /* the toolbar button (the IVK Tracker's): the current symbol faded, a slash, the one a press switches to; the label is that mode */
  function paint(btn, swap) {
    if (!btn) return;
    const dark = cur === 'dark', tl = dark ? 'Light mode' : 'Dark mode';
    btn.className = `ibtn theme${dark ? ' is-dark' : ''}${swap ? ' swap' : ''}`;
    btn.dataset.tip = tl;
    btn.setAttribute('aria-label', `Switch to ${tl.toLowerCase()}`);
    btn.setAttribute('aria-pressed', String(dark));
    btn.innerHTML = dark ? `<span class="tcur">${MOON}</span><span class="tsl">/</span><span class="tgo">${SUN}</span>`
      : `<span class="tcur">${SUN}</span><span class="tsl">/</span><span class="tgo">${MOON}</span>`;
  }
  function set(t, opts) {
    if (!ok(t)) return;
    const changed = t !== cur; cur = t;
    document.documentElement.dataset.theme = t;
    if (!opts || opts.save !== false) { try { localStorage.setItem(KEY, t); } catch (e) { /* this session only */ } }
    document.querySelectorAll('button[data-theme-btn]').forEach((b) => paint(b, changed && !(opts && opts.quiet)));
  }
  window.SoharTheme = {
    get: () => cur,
    set,
    toggle: () => { set(cur === 'dark' ? 'light' : 'dark'); return cur; },
    paint,
    /* the full-screen URL carries the pane's choice */
    url: (href) => { const u = new URL(href, location.href); u.searchParams.set('theme', cur); return u.href; },
  };
})();
