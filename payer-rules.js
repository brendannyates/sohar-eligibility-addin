// Payer credential rules for esketamine: whether a payer lets an MD or an NP bill each ESK visit (first treatment, start,
// observation), and whether the payer is approved for Besse stock. Pure functions, no Office.js or DOM.
// The table itself lives in each workbook on a "Payer Credential Rules" sheet, imported from the shared CSV in
// OneDrive (Add-In Manifests\Payer Credential Rules.csv). It is never committed: these repos are public.
// The same file ships in the Clinic Template Tool, Sohar Eligibility and ESK Tracker repos. Change all three together.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PayerRules = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SHEET = 'Payer Credential Rules';
  const HEADERS = ['Payer', 'State', 'Credential', 'Besse Approved', 'NP Fully Approved', 'Esketamine First Treatment', 'Esketamine Start', 'Esketamine Observation', 'Match Names', 'Notes'];
  const STEPS = { first: 'first treatment', start: 'start', obs: 'observation' };
  const COLS = {
    payer: /^(insurance\s*provider|payer|insurance)$/i, state: /^state$/i, cred: /^(provider\s*type|credential)$/i, besse: /besse/i,
    npFull: /np.*fully|fully\s*approved/i, first: /first\s*treatment/i, start: /^(esketamine\s*)?start$/i, obs: /observation/i, match: /match|alias/i, notes: /^notes?$/i,
  };
  const STATES = { AL: 'alabama', AK: 'alaska', AZ: 'arizona', AR: 'arkansas', CA: 'california', CO: 'colorado', CT: 'connecticut', DE: 'delaware', FL: 'florida', GA: 'georgia',
    HI: 'hawaii', ID: 'idaho', IL: 'illinois', IN: 'indiana', IA: 'iowa', KS: 'kansas', KY: 'kentucky', LA: 'louisiana', ME: 'maine', MD: 'maryland', MA: 'massachusetts',
    MI: 'michigan', MN: 'minnesota', MS: 'mississippi', MO: 'missouri', MT: 'montana', NE: 'nebraska', NV: 'nevada', NH: 'new hampshire', NJ: 'new jersey', NM: 'new mexico',
    NY: 'new york', NC: 'north carolina', ND: 'north dakota', OH: 'ohio', OK: 'oklahoma', OR: 'oregon', PA: 'pennsylvania', RI: 'rhode island', SC: 'south carolina',
    SD: 'south dakota', TN: 'tennessee', TX: 'texas', UT: 'utah', VT: 'vermont', VA: 'virginia', WA: 'washington', WV: 'west virginia', WI: 'wisconsin', WY: 'wyoming' };
  const txt = s => String(s == null ? '' : s).trim();
  // "Blue Shield of CA - HMO" -> "blue shield of ca hmo"
  const norm = s => ' ' + txt(s).toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  function yn(v){ const s = txt(v).toLowerCase(); if (/^(y|yes|true|approved)$/.test(s)) return 'yes'; if (/^(n|no|false|not approved)$/.test(s)) return 'no'; if (/^n\/?a$/.test(s)) return 'na'; return ''; }
  // NP first: "Do" is also a name ("Do, Minh NP", "Nguyen, Do PMHNP"), so DO only counts when no NP credential is present
  const credOf = v => { const s = txt(v).toUpperCase().replace(/[^A-Z]/g, ' '); if (/\b(NP|PMHNP|FNP|APRN|DNP|ARNP|CRNP)\b/.test(s)) return 'NP'; if (/\b(MD|DO)\b/.test(s)) return 'MD'; if (/\b(PA|PAC)\b/.test(s)) return 'PA'; return ''; };

  // Two-letter state in a payer, rule or clinic name: "Anthem CA", "Medicare WA (Noridian)", "Castro - CA", "Blue Shield of California".
  function stateOf(s){
    const raw = txt(s); if (!raw) return '';
    const codes = raw.match(/\b[A-Z]{2}\b/g) || []; const hit = codes.find(c => STATES[c]); if (hit) return hit;
    const n = norm(raw); for (const [c, name] of Object.entries(STATES)) if (n.includes(' ' + name + ' ')) return c;
    return '';
  }

  // Rows of the sheet or CSV (array of arrays, header row anywhere in the first 10) -> { rules, skipped } or { error }.
  function read(aoa){
    for (let r = 0; r < Math.min((aoa || []).length, 10); r++){
      const h = (aoa[r] || []).map(txt); const idx = {};
      for (const [k, re] of Object.entries(COLS)){ const i = h.findIndex(x => re.test(x)); if (i >= 0 && idx[k] == null) idx[k] = i; }
      if (idx.payer == null || idx.cred == null || idx.obs == null) continue;
      const rules = []; let skipped = 0;
      for (const v of aoa.slice(r + 1)){
        const payer = txt(v[idx.payer]); if (!payer){ continue; }
        const cred = credOf(v[idx.cred]); if (!cred){ skipped++; continue; }
        const at = k => idx[k] != null ? v[idx[k]] : '';
        const match = txt(at('match')).split(/[;\n]/).map(x => norm(x).trim()).filter(Boolean);
        if (!match.length) match.push(norm(payer.replace(/\([^)]*\)/g, ' ').replace(/\b[A-Z]{2}\b/g, ' ')).trim());
        rules.push({ payer, state: txt(at('state')).toUpperCase() || stateOf(payer), cred, besse: yn(at('besse')), npFull: yn(at('npFull')),
          first: yn(at('first')), start: yn(at('start')), obs: yn(at('obs')), match, notes: txt(at('notes')) });
      }
      return { rules, skipped };
    }
    return { error: 'Couldn’t find Payer, Credential and Esketamine Observation columns. Use the Payer Credential Rules file.' };
  }

  // Which rule group a payer name belongs to. hint = a state ('CA') or a clinic/location name to take one from.
  // The longest matching name wins; when the same name fits several states, the patient's state decides, else a rule
  // with no state; if that still leaves rules that disagree, the result is ambiguous.
  // -> null (not on the list) | { payer, state, MD, NP, ambiguous: [payer names] }
  function match(rules, payerName, hint){
    const n = norm(payerName); if (!(rules && rules.length) || !n.trim()) return null;
    let best = 0; const hits = [];
    for (const r of rules) for (const p of r.match){ if (!n.includes(' ' + p + ' ')) continue; if (p.length > best){ best = p.length; hits.length = 0; } if (p.length === best) hits.push(r); }
    if (!hits.length) return null;
    const groups = new Map(); for (const r of hits){ if (!groups.has(r.payer)) groups.set(r.payer, { payer: r.payer, state: r.state, MD: null, NP: null, PA: null }); groups.get(r.payer)[r.cred] = r; }
    let list = [...groups.values()];
    if (list.length > 1){
      const st = stateOf(payerName) || (hint && hint.length === 2 && STATES[hint.toUpperCase()] ? hint.toUpperCase() : stateOf(hint));
      const same = st ? list.filter(g => g.state === st) : [];
      list = same.length ? same : list.filter(g => !g.state).length ? list.filter(g => !g.state) : list;
    }
    const g = list[0];
    const sig = x => ['MD', 'NP'].map(c => x[c] ? [x[c].besse, x[c].first, x[c].start, x[c].obs].join() : '-').join('|');
    const ambiguous = list.length > 1 && list.some(x => sig(x) !== sig(g)) ? list.map(x => x.payer) : [];
    return Object.assign({}, g, { ambiguous, alts: ambiguous.length ? list : [g] });
  }

  // ESK visit kind from the event name: "Esketamine Start" -> start; otherwise the patient's first dosing visit is
  // the first treatment and later ones are observation.
  function step(event, isFirst){ return /\bstart\b|induction/i.test(txt(event)) ? 'start' : isFirst ? 'first' : 'obs'; }

  // Can this credential bill this ESK step for this payer? -> { ok: true | false | null, why }
  // null = unknown: payer not on the list, state rules that disagree on this answer, no row for the credential, or a blank cell.
  function check(m, cred, stp){
    const c = credOf(cred) || txt(cred).toUpperCase();
    if (!m) return { ok: null, why: 'Payer isn’t on the Payer Credential Rules list' };
    if (m.ambiguous.length){
      const each = m.alts.map(g => check(Object.assign({}, g, { ambiguous: [] }), c, stp));
      if (each.every(x => x.ok === each[0].ok) && each[0].ok !== null) return each[0].ok ? each[0] : { ok: false, why: each[0].why.replace(m.alts[0].payer, m.ambiguous.join(' / ')) };
      return { ok: null, why: `Matches ${m.ambiguous.join(' and ')}; check which state the plan is from` };
    }
    const row = m[c]; if (!row) return { ok: null, why: `No ${c || 'credential'} row for ${m.payer}` };
    const v = row[stp];
    if (v === 'yes') return { ok: true, why: '' };
    if (v === 'no') return { ok: false, why: `${m.payer} doesn’t allow ${c === 'NP' ? 'an NP' : c === 'MD' ? 'an MD' : c} to bill the esketamine ${STEPS[stp]}${row.notes ? ` (${row.notes})` : ''}` };
    return { ok: null, why: `${m.payer}: ${c} ${STEPS[stp]} not on the list yet` };
  }
  // Besse approval for the payer: 'approved' | 'not approved' | 'pending' (blank on the list) | 'unknown' (payer not on it, or ambiguous)
  function besse(m){
    if (!m) return 'unknown';
    const one = g => { const v = (g.MD && g.MD.besse) || (g.NP && g.NP.besse) || ''; return v === 'yes' ? 'approved' : v === 'no' ? 'not approved' : 'pending'; };
    const all = m.alts.map(one);
    return all.every(x => x === all[0]) ? all[0] : 'unknown';
  }
  // Rows for writing the sheet back (same columns as the CSV).
  const toRows = rules => [HEADERS].concat(rules.map(r => {
    const f = v => v === 'yes' ? 'Yes' : v === 'no' ? 'No' : v === 'na' ? 'N/A' : '';
    return [r.payer, r.state, r.cred, f(r.besse), f(r.npFull), f(r.first), f(r.start), f(r.obs), r.match.join('; '), r.notes];
  }));

  return { SHEET, HEADERS, STEPS, read, match, step, check, besse, stateOf, credOf, norm, toRows };
});
