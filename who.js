/* who.js — who is using the add-in (v0.28, the IVK Tracker's "Enter your name" screen). Pure: no Office.js, no DOM, no
   network. Loaded by index.html and dashboard.html, and by tests/.
   The workbook's list of people lives in its Office document settings (USERS_SETTING: [{name, added}]), never on a sheet.
   Config "Admin names" (the older list) is offered too, so nobody has to add themselves twice. Today's pick is kept on
   this computer per workbook (localStorage, whoKey): { day, name, used }, and lasts the day. The picked name is the one
   pre-check ticks, notes, code rules and saved audits record. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SoharWho = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const USERS_SETTING = 'soharUsers';
  const WB_ID_SETTING = 'soharWbId'; // the workbook's own id, when Excel gives no file URL
  const MAX_USED = 20;
  const txt = (v) => String(v ?? '').trim();
  const userKey = (n) => txt(n).toLowerCase();
  /** The name as typed to register: spaces tidied, at most 30 characters. */
  const cleanName = (s) => txt(s).replace(/\s+/g, ' ').slice(0, 30).trim();

  /** The saved list (an array, or its JSON), cleaned: one row per name, first spelling wins. */
  function usersFrom(v) {
    let list = v;
    if (typeof v === 'string') { try { list = JSON.parse(v || '[]'); } catch (e) { list = []; } }
    if (!Array.isArray(list)) return [];
    const seen = new Set(), out = [];
    for (const u of list) {
      const name = cleanName(u && typeof u === 'object' ? u.name : u), k = userKey(name);
      if (!name || seen.has(k)) continue;
      seen.add(k);
      out.push({ name, added: txt(u && u.added) });
    }
    return out;
  }
  /** Add a name (no duplicates, any case). Returns the new list and the name as the list spells it. */
  function addUser(users, name, when) {
    const list = usersFrom(users), n = cleanName(name);
    if (!n) return { users: list, name: '' };
    const had = list.find((u) => userKey(u.name) === userKey(n));
    if (had) return { users: list, name: had.name };
    return { users: list.concat({ name: n, added: txt(when) }), name: n };
  }
  /** Who the name screen offers: the workbook's list, then Config "Admin names" not already on it. */
  function listFor(users, adminNames) {
    const out = usersFrom(users), seen = new Set(out.map((u) => userKey(u.name)));
    for (const a of adminNames || []) {
      const name = cleanName(a), k = userKey(name);
      if (name && !seen.has(k)) { seen.add(k); out.push({ name, added: '', config: true }); }
    }
    return out;
  }
  /** The screen's order: people who used this computer on this workbook first (most recent first), then A-Z. */
  function pickOrder(users, here) {
    const H = (here || []).map(userKey), rank = (u) => { const i = H.indexOf(userKey(u.name)); return i < 0 ? Infinity : i; };
    return (users || []).map((u) => ({ ...u, here: rank(u) !== Infinity }))
      .sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  }
  /** A short key for one workbook, from its URL (or the id saved in it), for this computer's memory. */
  function wbKeyOf(s) {
    let h = 0x811c9dc5; const t = String(s || '').trim().toLowerCase().replace(/[?#].*$/, '');
    for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return t ? `w${h.toString(36)}` : '';
  }
  const whoKey = (wb) => (wb ? `sohar.who:${wb}` : '');
  /** What this computer remembers for one workbook: { day: "2026-11-07", name: "Sam", used: ["Sam", "Ana"] }. */
  function whoRec(json) {
    try {
      const o = JSON.parse(json || '{}') || {};
      return { day: txt(o.day), name: txt(o.name), used: Array.isArray(o.used) ? o.used.map(txt).filter(Boolean) : [] };
    } catch (e) { return { day: '', name: '', used: [] }; }
  }
  const whoToday = (rec, dayIso) => (rec && rec.day && rec.day === dayIso ? rec.name : '');
  const whoPick = (rec, name, dayIso) => ({ day: dayIso, name, used: [name].concat(((rec && rec.used) || []).filter((n) => userKey(n) !== userKey(name))).slice(0, MAX_USED) });
  /** Today's pick, but only while that name is still offered (someone removed from the list is asked again). */
  function pickedToday(rec, dayIso, list) {
    const n = whoToday(rec, dayIso);
    const u = n && (list || []).find((x) => userKey(x.name) === userKey(n));
    return u ? u.name : '';
  }
  /** The local day as yyyy-mm-dd (the pick lasts until midnight on this computer). */
  const dayIso = (d) => { const t = d || new Date(); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`; };

  return { USERS_SETTING, WB_ID_SETTING, userKey, cleanName, usersFrom, addUser, listFor, pickOrder, wbKeyOf, whoKey, whoRec, whoToday, whoPick, pickedToday, dayIso };
});
