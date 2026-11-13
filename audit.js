/* audit.js — patient account audit (v0.23; saved audits and the export v0.23.1). Pure: no Office.js, no DOM, no network. Loaded after core.js and fees.js
   by dashboard.html, and by tests/.
   Staff drop up to three files in a patient's record: the InstaMed payment transactions report (.xls), the Arietis
   account activity detail (.xlsx) and a fresh NextGen appointments report with the plan year's kept visits. This:
   1. reads each by its header (detectKind, parseInstamed, parseArietis, parseAppointments) and checks it is this
      patient (LASFIR, MRN);
   2. checks every charge (codes billed against the visit's event, allowed against the fee schedule, plan paid plus
      patient share against allowed, balance arithmetic, claims not processed);
   3. matches InstaMed payments to Arietis patient payments (card last 4, amount, posting window);
   4. learns the benefits from how claims processed (copay, coinsurance, deductible, OOP met) and compares them with
      Sohar's;
   5. re-quotes the pending claims and the booked visits from those benefits, next to any Sohar Flow estimate;
   6. adds it up: the Arietis balance, the true patient balance, and every anomaly by system.
   Nothing here is saved: the files and results live in the full-screen window's memory only (they carry names). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'), require('./fees.js'));
  else root.SoharAudit = factory(root.SoharCore, root.SoharFees);
})(typeof self !== 'undefined' ? self : this, function (C, F) {
  'use strict';

  const str = (v) => (v == null ? '' : String(v).trim());
  const num = (v) => { if (typeof v === 'number') return isFinite(v) ? v : null; const s = str(v).replace(/[$,\s]/g, ''); if (!s) return null; const neg = /^\(.*\)$/.test(s); const n = parseFloat(s.replace(/[()]/g, '')); return isNaN(n) ? null : neg ? -n : n; };
  const r2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
  const sum = (xs, f) => r2(xs.reduce((a, x) => a + (f ? f(x) : x), 0));
  const near = (a, b, tol) => Math.abs((a || 0) - (b || 0)) <= (tol == null ? 0.01 : tol);
  const usd = (n) => (n == null ? '—' : `${n < 0 ? '-' : ''}$${Math.abs(r2(n)).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  const DAY = 864e5;
  const norm = (s) => str(s).toLowerCase().replace(/[^a-z0-9#]+/g, ' ').trim();

  // ---------- days ----------
  /** A cell as a UTC day in ms (Excel serials keep their day; text dates go through core.parseDate), or null. */
  function dayOf(v) {
    if (typeof v === 'number') return v > 59 ? Math.round((Math.floor(v) - 25569) * DAY) : null;
    const d = C.parseDate(v);
    return d ? Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) : null;
  }
  /** "h:mm AM" from an Excel serial's fraction, or ''. */
  function timeOf(v) {
    if (typeof v !== 'number') { const m = str(v).match(/\b(\d{1,2}):(\d{2})\s*([AP])\.?M?/i); return m ? `${+m[1]}:${m[2]} ${m[3].toUpperCase()}M` : ''; }
    const mins = Math.round((v - Math.floor(v)) * 1440);
    if (!mins) return '';
    const h = Math.floor(mins / 60) % 24, m = mins % 60;
    return `${h % 12 || 12}:${String(m).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`;
  }
  const iso = (t) => (t == null ? '' : new Date(t).toISOString().slice(0, 10));
  const us = (t) => { if (t == null) return ''; const d = new Date(t); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`; };

  // ---------- reading the files ----------
  const cells = (r) => (r || []).map((v) => str(v).replace(/^﻿/, ''));
  /** The first row (of the first 25) holding every header in need; {i, col(name)} or null. */
  function headerRow(rows, need) {
    for (let i = 0; i < Math.min(25, rows.length); i++) {
      const h = cells(rows[i]).map((x) => x.toLowerCase());
      if (need.every((n) => h.includes(n.toLowerCase()))) {
        const col = (name) => h.indexOf(String(name).toLowerCase());
        return { i, col, names: cells(rows[i]) };
      }
    }
    return null;
  }
  const INSTAMED_HEAD = ['Patient Account #', 'Transaction Date', 'Amount', 'Card # - Last 4 Digits'];
  const ARIETIS_HEAD = ['Date', 'Code', 'Description', 'Billed', 'Expected', 'Amount', 'Balance'];
  const APPT_HEAD = ['Appt Date', 'Event', 'Status'];
  /** Which file this is, from its header: 'instamed', 'arietis', 'appointments' or ''. */
  function detectKind(rows) {
    if (headerRow(rows, INSTAMED_HEAD)) return 'instamed';
    if (headerRow(rows, ARIETIS_HEAD)) return 'arietis';
    if (headerRow(rows, APPT_HEAD)) return 'appointments';
    return '';
  }
  /** "LAST, FIRST" or {last, first} -> LASFIR. */
  const lasfirOf = (last, first) => C.lasfir(last, first);
  const lasfirComma = (s) => { const [l, f] = str(s).split(','); return f == null ? '' : lasfirOf(l, f); };

  /**
   * InstaMed "Payment Transactions": one row per card transaction. Counted payments are settled or complete sales
   * (a refund or return counts negative); voided or declined ones are listed but not counted. amount = Effective Amount.
   */
  function parseInstamed(rows) {
    const h = headerRow(rows, INSTAMED_HEAD);
    if (!h) throw new Error('Not an InstaMed payment transactions report (no "Patient Account #" / "Card # - Last 4 Digits" header).');
    const get = (r, n) => { const c = h.col(n); return c < 0 ? '' : r[c]; };
    const txns = [];
    let reportTotal = null;
    for (let i = h.i + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const first = str(r[0]);
      if (/^total/i.test(first)) { reportTotal = num(get(r, 'Effective Amount')) ?? num(get(r, 'Amount')); break; }
      const raw = get(r, 'Transaction Date');
      const day = dayOf(raw);
      const amt = num(get(r, 'Effective Amount')) ?? num(get(r, 'Amount'));
      if (day == null || amt == null) continue;
      const action = str(get(r, 'Action')), status = str(get(r, 'Current Status')), wf = str(get(r, 'Workflow Status'));
      const refund = /refund|return|credit/i.test(action);
      const dead = /void|declin|fail|cancel|reject|error/i.test(`${status} ${wf}`) || /void/i.test(action);
      const last = str(get(r, 'Patient Last Name')), firstName = str(get(r, 'Patient First Name'));
      const holder = str(get(r, 'Card Holder Name'));
      txns.push({
        row: i + 1, day, time: timeOf(raw), account: str(get(r, 'Patient Account #')), lasfir: last || firstName ? lasfirOf(last, firstName) : '',
        holderLasfir: holder ? (() => { const p = holder.split(/\s+/); return p.length > 1 ? lasfirOf(p[p.length - 1], p[0]) : ''; })() : '',
        type: str(get(r, 'Type')), action, status: status || wf, last4: str(get(r, 'Card # - Last 4 Digits')).replace(/\D/g, '').slice(-4),
        outlet: str(get(r, 'Outlet Desc')), amount: r2(refund ? -Math.abs(amt) : amt), counted: !dead,
      });
    }
    txns.sort((a, b) => a.day - b.day || a.row - b.row);
    const counted = txns.filter((t) => t.counted);
    return {
      txns, reportTotal, total: sum(counted, (t) => t.amount),
      accounts: [...new Set(txns.map((t) => t.account).filter(Boolean))], lasfirs: [...new Set(txns.map((t) => t.lasfir).filter(Boolean))],
      start: txns.length ? txns[0].day : null, end: txns.length ? txns[txns.length - 1].day : null,
    };
  }

  // Arietis transaction codes under a charge. Patient responsibility lines (COPAY / COINS / DEDUCT) carry $0 and say the
  // amount in their text; the money lines carry signed amounts (payments and write-offs negative, reversals positive).
  const TXN_TYPES = [
    [/^claim/i, 'claim'], [/^inspmt\s*-\s*rev|^ins\s*pmt\s*-\s*rev/i, 'insRev'], [/^inspmt|^ins\s*pmt|^inspay/i, 'insPmt'],
    [/^(guar|pat|pt)\s*pmt\s*-\s*rev/i, 'patRev'], [/^(guar|pat|pt)\s*pmt/i, 'patPmt'], [/^contwoff|^contract/i, 'contractual'],
    [/^copay/i, 'copay'], [/^coins/i, 'coins'], [/^deduct|^ded\b/i, 'ded'], [/^stmnt|^statement/i, 'statement'], [/^refund/i, 'refund'],
  ];
  const typeOf = (code) => { for (const [re, t] of TXN_TYPES) if (re.test(code)) return t; return 'other'; };
  const CODE_RE = /^([A-Z]\d{4}|\d{4}[A-Z0-9])\b(.*)$/i;
  /** "99214 - 95" -> {code: '99214', mods: ['95']}, or null for a transaction code. */
  function chargeCode(v) {
    const m = CODE_RE.exec(str(v));
    if (!m || typeOf(str(v)) !== 'other') return null;
    return { code: m[1].toUpperCase(), mods: m[2].split(/[\s,-]+/).map((x) => x.trim().toUpperCase()).filter((x) => /^[A-Z0-9]{2}$/.test(x)) };
  }
  const RESP_RE = { copay: /co-?\s*pay(?:ment)?\s*:?\s*\$?(-?[\d,]+\.?\d*)/i, coins: /co-?\s*insurance\s*:?\s*\$?(-?[\d,]+\.?\d*)/i, ded: /deductible\s*:?\s*\$?(-?[\d,]+\.?\d*)/i };
  /**
   * Arietis "Account Detail": each charge row (Billed filled, a CPT/HCPCS code) followed by its transactions.
   * Per charge: billed, Arietis expected, the balance it shows, claims sent, plan paid, contractual and other
   * adjustments (with CARCs), patient share (copay, coinsurance, deductible) and patient payments with their refs.
   */
  function parseArietis(rows) {
    const h = headerRow(rows, ARIETIS_HEAD);
    if (!h) throw new Error('Not an Arietis account activity detail (no Date / Code / Billed / Expected / Balance header).');
    const get = (r, n) => { const c = h.col(n); return c < 0 ? '' : r[c]; };
    const charges = [];
    let cur = null, lastPosted = null;
    const lasfirs = new Set(), payers = new Set();
    for (let i = h.i + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const codeCell = str(get(r, 'Code'));
      if (!codeCell && !str(get(r, 'Description'))) continue;
      const billed = num(get(r, 'Billed'));
      const cc = billed != null ? chargeCode(codeCell) : null;
      if (cc) {
        const pt = str(get(r, 'Patient'));
        if (pt) lasfirs.add(lasfirComma(pt));
        cur = { row: i + 1, dos: dayOf(get(r, 'Date')), code: cc.code, mods: cc.mods, codeText: codeCell, desc: str(get(r, 'Description')), provider: str(get(r, 'Provider')),
          ins: str(get(r, 'Ins')), billed: r2(billed), expected: num(get(r, 'Expected')), balance: num(get(r, 'Balance')) ?? 0, lines: [] };
        charges.push(cur);
        continue;
      }
      if (!cur) continue;
      const type = typeOf(codeCell);
      const text = [get(r, 'Info'), get(r, 'User'), get(r, 'Description')].map(str).filter(Boolean).join(' ');
      const desc = str(get(r, 'Description')) || text;
      const posted = dayOf(get(r, 'Date'));
      if (posted != null && (lastPosted == null || posted > lastPosted)) lastPosted = posted;
      const line = { row: i + 1, type, code: codeCell, posted, time: timeOf(get(r, 'Date')), info: str(get(r, 'Info')), desc, amount: num(get(r, 'Amount')) ?? 0 };
      if (RESP_RE[type]) { const m = RESP_RE[type].exec(text); line.resp = m ? Math.abs(num(m[1]) || 0) : 0; }
      if (type === 'claim') { const m = /claim\s+(\S+)/i.exec(desc); line.claim = m ? m[1] : ''; const p = /\bto\s+(.+)$/i.exec(desc); if (p) payers.add(p[1].trim()); }
      if (type === 'patPmt' || type === 'patRev' || type === 'refund') {
        const m = /\b(IM|CC|CK|ET|EFT|CASH|CHK)\s*Ref\s*#\s*([A-Za-z0-9]+)/i.exec(desc);
        line.refType = m ? m[1].toUpperCase() : ''; line.ref = m ? m[2] : '';
        const b = /batch\s*#\s*(\d+)/i.exec(desc); line.batch = b ? b[1] : '';
      }
      if (type === 'insPmt' || type === 'insRev') { const m = /\b(ET|CK|EFT)\s*Ref\s*#\s*([A-Za-z0-9]+)/i.exec(desc); line.ref = m ? m[2] : ''; }
      const carc = /^(CO|PR|OA|PI|CR)\s*-?\s*\d{1,3}[A-Z]?$/i.exec(line.info);
      if (carc) line.carc = line.info.toUpperCase().replace(/[\s-]/g, '');
      cur.lines.push(line);
    }
    for (const c of charges) Object.assign(c, totalsOf(c));
    charges.sort((a, b) => (a.dos ?? 0) - (b.dos ?? 0) || a.row - b.row);
    return { charges, lasfirs: [...lasfirs].filter(Boolean), payers: [...payers], asOf: lastPosted,
      start: charges.length ? charges[0].dos : null, end: charges.length ? charges[charges.length - 1].dos : null };
  }
  /** One charge's money: plan paid, write-offs, patient share and payments; processed when the plan acted on it. */
  function totalsOf(c) {
    const L = c.lines, of = (t) => L.filter((l) => l.type === t);
    const insPaid = r2(-sum(of('insPmt'), (l) => l.amount) - sum(of('insRev'), (l) => l.amount));
    const contractual = r2(-sum(of('contractual'), (l) => l.amount));
    const otherAdj = r2(-sum(of('other'), (l) => l.amount) - sum(of('refund'), (l) => l.amount));
    const copay = sum(of('copay'), (l) => l.resp), coins = sum(of('coins'), (l) => l.resp), ded = sum(of('ded'), (l) => l.resp);
    const patPaid = r2(-sum(of('patPmt'), (l) => l.amount) - sum(of('patRev'), (l) => l.amount));
    const carcs = [...new Set(L.map((l) => l.carc).filter(Boolean))];
    const claims = of('claim');
    const processed = of('insPmt').length > 0 || contractual > 0 || copay + coins + ded > 0 || carcs.some((x) => x !== 'CO45');
    const respPosted = L.filter((l) => RESP_RE[l.type] && l.resp > 0).map((l) => l.posted).filter((x) => x != null);
    return { insPaid, contractual, otherAdj, copay, coins, ded, resp: r2(copay + coins + ded), patPaid, carcs, claims, processed,
      firstClaim: claims.length ? Math.min(...claims.map((l) => l.posted ?? Infinity)) : null, respPosted: respPosted.length ? Math.min(...respPosted) : null,
      // Balance from the lines: billed plus every signed amount. Arietis's Balance column should say the same.
      calcBalance: r2(c.billed + sum(L.filter((l) => l.type !== 'statement' && !RESP_RE[l.type]), (l) => l.amount)) };
  }

  // The patient's share on a remit is group code PR with a reason: PR-1 deductible, PR-2 coinsurance, PR-3 copay.
  const PR_PARTS = [['PR1', 'Deductible', 'ded'], ['PR2', 'Coinsurance', 'coins'], ['PR3', 'Copay', 'copay']];
  const PR_NAMES = { PR27: 'Coverage ended before this date', PR96: 'Non-covered charge', PR204: 'Not covered by the current plan' };
  /**
   * One charge for a balance's list (v0.24): date of service, code, billed, plan paid, its PR amounts (PR-1, PR-2,
   * PR-3 from the patient-share lines, then any other PR code on its adjustment lines), patient paid and balance.
   */
  function chargeRow(c) {
    const pr = PR_PARTS.map(([code, label, k]) => ({ code, label, amount: r2(c[k] || 0) })).filter((p) => p.amount);
    const other = new Map();
    for (const l of c.lines || []) {
      if (!l.carc || !/^PR/.test(l.carc) || /^PR[123]$/.test(l.carc)) continue;
      other.set(l.carc, r2((other.get(l.carc) || 0) + Math.abs(l.resp || l.amount || 0)));
    }
    for (const [code, amount] of other) pr.push({ code, label: PR_NAMES[code] || 'Other patient responsibility: see the remit', amount });
    return { dos: c.dos, code: c.code, mods: c.mods || [], desc: c.desc, billed: c.billed, insPaid: c.insPaid, contractual: c.contractual, pr, resp: c.resp,
      patPaid: c.patPaid, balance: c.balance, processed: c.processed };
  }

  const KEPT_RE = /^(kept|checked\s*out|check\s*out|complete|completed|arrived|seen)$/i;
  /** NextGen appointments listing (any of its layouts): one row per appointment, kept or not. */
  function parseAppointments(rows) {
    const h = headerRow(rows, APPT_HEAD);
    if (!h) throw new Error('Not a NextGen appointments report (no Appt Date / Event / Status header).');
    const get = (r, ...ns) => { for (const n of ns) { const c = h.col(n); if (c >= 0 && str(r[c])) return r[c]; } return ''; };
    const visits = [];
    for (let i = h.i + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const day = dayOf(get(r, 'Appt Date'));
      if (day == null || !str(get(r, 'Event'))) continue;
      const status = str(get(r, 'Status', 'Appt Status'));
      const pat = str(get(r, 'Pat Name', 'Patient Name', 'Patient'));
      visits.push({ row: i + 1, day, time: timeOf(get(r, 'Beg Tm', 'Start Time', 'Time')), event: str(get(r, 'Event')), status, kept: KEPT_RE.test(status),
        loc: str(get(r, 'Loc Name', 'Location')), provider: str(get(r, 'Rendering', 'Res Name', 'Provider')), mrn: str(get(r, 'MRN', 'Md Rc', 'Med Rec Nbr')).replace(/^0+(?=\d)/, ''),
        per: str(get(r, 'Per Nbr')), lasfir: pat ? lasfirComma(pat) || lasfirOf(pat.split(/\s+/).pop(), pat.split(/\s+/)[0]) : '',
        payer: str(get(r, 'Payer Name (Cht Guar)', 'Payer Name (Lst Enc)', 'Payer', 'Payer Name')) });
    }
    visits.sort((a, b) => a.day - b.day);
    return { visits, start: visits.length ? visits[0].day : null, end: visits.length ? visits[visits.length - 1].day : null };
  }

  /**
   * Is this file this patient's? who: {mrn, lasfir}. Returns {ok, text}. An appointments report keeps only the patient's
   * rows (by MRN, else LASFIR), so a whole-clinic report works too.
   */
  function identify(kind, parsed, who) {
    const L = str(who.lasfir).toUpperCase(), M = str(who.mrn).replace(/^0+(?=\d)/, '');
    if (kind === 'appointments') {
      const mine = parsed.visits.filter((v) => (M && v.mrn ? v.mrn === M : L && v.lasfir === L));
      if (!mine.length) return { ok: false, text: `No appointments for ${L || M} in this report (${parsed.visits.length} rows read).`, visits: [] };
      return { ok: true, text: `${mine.length} appointment${mine.length === 1 ? '' : 's'} for ${L}${parsed.visits.length > mine.length ? ` (of ${parsed.visits.length} in the report)` : ''}`, visits: mine };
    }
    const ls = parsed.lasfirs || [];
    if (!ls.length) return { ok: true, text: 'No patient name in the file to check: make sure it is this patient\'s.' };
    if (kind === 'instamed' && (parsed.accounts || []).length > 1) return { ok: false, text: `This InstaMed report has ${parsed.accounts.length} patient accounts. Run it for one patient.` };
    if (ls.length > 1) return { ok: false, text: `This file has more than one patient (${ls.join(', ')}). Export one patient's account.` };
    if (L && ls[0] !== L) return { ok: false, text: `This file is for ${ls[0]}, not ${L}.` };
    return { ok: true, text: `Patient ${ls[0]} matches.` };
  }

  // ---------- codes ----------
  const FAMILY = [
    [/^9921[1-5]$/, 'E/M established'], [/^9920[2-5]$/, 'E/M new'], [/^9079[12]$/, 'Psychiatric evaluation'], [/^(9083[368]|90785)$/, 'Psychotherapy add-on'],
    // 90867 is the MT (motor threshold) visit and 90869 the MTR (MT redo), Brendan 11/2: one family, so one billed for
    // the other reads as a level mismatch.
    [/^9083[247]$/, 'Psychotherapy'], [/^90868$/, 'TMS'], [/^9086[79]$/, 'TMS motor threshold'], [/^G208[23]$/, 'Esketamine'], [/^961[3-4]\d$/, 'Testing'],
  ];
  const familyOf = (code) => { for (const [re, f] of FAMILY) if (re.test(code)) return f; return 'Other'; };
  const isAddon = (code) => familyOf(code) === 'Psychotherapy add-on';
  const classOf = (code) => (/^9086[789]$/.test(code) ? 'TMS' : /^G208[23]$/.test(code) ? 'Esketamine' : 'Visit');
  // Interventions: the Estimate Codes sheet says "none" for them (no cost estimate), but the audit still knows what a
  // TMS or esketamine visit bills. First match wins.
  const AUDIT_EVENTS = [
    // MTR (MT redo) before MT: "MT-Redo" also has the word MT.
    [/\bmtr\b|\bmt\s*-?\s*re-?do\b|motor\s*threshold\s*re-?do|re-?determin/i, ['90869'], 'TMS motor threshold redo (MTR)'],
    [/motor\s*threshold|\bmt\b/i, ['90867'], 'TMS motor threshold (MT)'],
    [/\btms\b.*\b(mapping|initial)\b|\bmapping\b/i, ['90867'], 'TMS initial session with mapping'],
    [/\b(esketamine|spravato|esk)\b(?!.*consult)/i, ['G2082|G2083'], 'Esketamine session (dose code G2082 or G2083)'],
    [/\btms\b(?!.*\b(f\/?u|follow|consult)\b)/i, ['90868'], 'TMS treatment session'],
  ];
  /** Expected codes for a NextGen event: [{codes: ['99214'] or ['G2082', 'G2083'] (any one), note}], plus the note. */
  function expectedFor(event, codeRows) {
    for (const [re, codes, note] of AUDIT_EVENTS) if (re.test(str(event))) return { lines: codes.map((c) => c.split('|')), note, source: 'audit' };
    const rows = codeRows && codeRows.length ? codeRows : F ? F.ESTIMATE_CODE_SEEDS : [];
    const ce = F ? F.codesForEvent(event, rows) : { codes: [] };
    if (ce.none || !ce.codes.length) return { lines: [], note: ce.note || `No expected codes for "${str(event)}" (Estimate Codes)`, source: '' };
    return { lines: ce.codes.map((c) => [c.code]), note: ce.note, source: 'Estimate Codes' };
  }
  /** Billed codes on a day against the visits' expected codes: [{sev, text}] (empty when they agree). */
  function codeChecks(visits, charges, codeRows) {
    const out = [];
    const billed = charges.map((c) => c.code);
    const used = new Set();
    const exp = visits.map((v) => ({ v, e: expectedFor(v.event, codeRows) }));
    for (const { v, e } of exp) {
      if (!e.lines.length) { out.push({ sev: 'low', text: `${v.event}: ${e.note}` }); continue; }
      for (const any of e.lines) {
        const hit = charges.find((c, i) => !used.has(i) && any.includes(c.code));
        if (hit) { used.add(charges.indexOf(hit)); continue; }
        const fam = familyOf(any[0]);
        const same = charges.findIndex((c, i) => !used.has(i) && familyOf(c.code) === fam && fam !== 'Other');
        if (same >= 0) { used.add(same); out.push({ sev: 'medium', text: `${v.event}: expected ${any.join(' or ')}, billed ${charges[same].code} (different ${fam} level)` }); continue; }
        out.push({ sev: 'high', text: `${v.event}: expected ${any.join(' or ')} (${e.note || fam}), not billed${billed.length ? `; billed ${billed.join(', ')}` : ''}` });
      }
    }
    charges.forEach((c, i) => {
      if (used.has(i)) return;
      if (isAddon(c.code) && charges.some((x) => /^E\/M/.test(familyOf(x.code)))) out.push({ sev: 'low', text: `Psychotherapy add-on ${c.code} billed with the E/M visit: confirm the note supports it` });
      else out.push({ sev: 'medium', text: `${c.code} billed${visits.length ? `, not expected for ${visits.map((v) => v.event).join(' + ')}` : ''}` });
    });
    const tele = visits.some((v) => /tele|video|virtual/i.test(`${v.event} ${v.loc}`));
    const em = charges.filter((c) => /^E\/M|Psych/.test(familyOf(c.code)));
    if (visits.length && em.length) {
      if (tele && em.some((c) => !c.mods.includes('95') && !c.mods.includes('GT'))) out.push({ sev: 'low', text: 'Telehealth visit billed without modifier 95' });
      if (!tele && em.some((c) => c.mods.includes('95'))) out.push({ sev: 'low', text: 'Modifier 95 (telehealth) on an in-office visit' });
    }
    return out;
  }

  // ---------- payments: InstaMed against Arietis ----------
  const refMatches = (ref, last4) => !!ref && !!last4 && (ref === last4 || ref.endsWith(last4));
  /** Subset of items whose amounts add up to target (cents), earliest first; null when none. Small lists only. */
  function subsetSum(items, target, amt) {
    const t = Math.round(target * 100);
    if (t <= 0 || items.length > 24) return null;
    const best = new Map([[0, []]]);
    for (let i = 0; i < items.length; i++) {
      const a = Math.round(amt(items[i]) * 100);
      if (a <= 0) continue;
      for (const [s, idx] of [...best]) { const n = s + a; if (n <= t && !best.has(n)) best.set(n, [...idx, i]); }
      if (best.has(t)) break;
    }
    return best.has(t) ? best.get(t).map((i) => items[i]) : null;
  }
  /**
   * Match InstaMed transactions to Arietis patient payments. Arietis posts a card payment as GUARPMT "IM Ref # <last 4>
   * (Batch # n)", often split across charges. Order: a batch adding up to the transaction, one line, several lines,
   * then several transactions posted in one batch. Window: posted 2 days before to 45 days after the transaction.
   * A payment and its reversal (same ref, batch and amount) cancel and are left out.
   */
  function matchPayments(im, ar) {
    const lines = [];
    for (const c of ar ? ar.charges : []) for (const l of c.lines) if (l.type === 'patPmt' || l.type === 'patRev') lines.push({ ...l, dos: c.dos, code: c.code, pay: r2(-l.amount) });
    // reversal pairs
    for (const rev of lines.filter((l) => l.type === 'patRev')) {
      const p = lines.find((l) => l.type === 'patPmt' && !l.reversed && l.ref === rev.ref && l.batch === rev.batch && near(l.pay, -rev.pay));
      rev.reversed = true; if (p) p.reversed = true;
    }
    const live = lines.filter((l) => l.type === 'patPmt' && !l.reversed);
    const txns = im ? im.txns.filter((t) => t.counted && t.amount > 0) : [];
    const used = new Set(), done = new Set(), matches = [];
    const inWin = (t, l) => l.posted != null && l.posted >= t.day - 2 * DAY && l.posted <= t.day + 45 * DAY;
    const cands = (t) => live.filter((l) => !used.has(l) && refMatches(l.ref, t.last4) && inWin(t, l));
    const take = (ts, ls, how) => { ts.forEach((t) => done.add(t)); ls.forEach((l) => used.add(l)); matches.push({ txns: ts, lines: ls, how, amount: sum(ts, (t) => t.amount) }); };
    // In date order, each transaction takes the nearest posting (one posted before the transaction counts double). Arietis
    // posts late, so when several payments share a card and an amount, the most recent ones are the ones left unposted.
    const gap = (t, posted) => (posted >= t.day ? posted - t.day : 2 * (t.day - posted));
    for (const t of txns) {
      const cs = cands(t);
      const batches = new Map();
      for (const l of cs) { const k = l.batch || `row${l.row}`; if (!batches.has(k)) batches.set(k, []); batches.get(k).push(l); }
      const b = [...batches.values()].filter((ls) => near(sum(ls, (l) => l.pay), t.amount)).sort((x, y) => gap(t, Math.min(...x.map((l) => l.posted))) - gap(t, Math.min(...y.map((l) => l.posted))))[0];
      if (b) { take([t], b, 'batch'); continue; }
      const one = cs.filter((l) => near(l.pay, t.amount)).sort((x, y) => gap(t, x.posted) - gap(t, y.posted))[0];
      if (one) { take([t], [one], 'line'); continue; }
      const some = subsetSum(cs.sort((x, y) => x.posted - y.posted || x.row - y.row), t.amount, (l) => l.pay);
      if (some) take([t], some, 'lines');
    }
    // several transactions in one batch
    const groups = new Map();
    for (const l of live) if (!used.has(l) && l.batch) { const k = `${l.ref}|${l.batch}`; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(l); }
    for (const ls of groups.values()) {
      const posted = Math.min(...ls.map((l) => l.posted));
      const ts = txns.filter((t) => !done.has(t) && refMatches(ls[0].ref, t.last4) && posted >= t.day - 2 * DAY && posted <= t.day + 45 * DAY);
      const pick = subsetSum(ts, sum(ls, (l) => l.pay), (t) => t.amount);
      if (pick) take(pick, ls, 'batch of several');
    }
    // Several transactions against several lines of one card, when nothing else fits: same card, window, equal totals.
    const left = txns.filter((t) => !done.has(t));
    for (const last4 of [...new Set(left.map((t) => t.last4))]) {
      const ts = left.filter((t) => t.last4 === last4 && !done.has(t));
      const ls = live.filter((l) => !used.has(l) && refMatches(l.ref, last4) && ts.some((t) => inWin(t, l)));
      if (ts.length && ls.length && near(sum(ts, (t) => t.amount), sum(ls, (l) => l.pay))) take(ts, ls, 'card total');
    }
    const unposted = txns.filter((t) => !done.has(t));
    // Same card and amount as a payment that did post: which one is missing can't be told from the files.
    for (const t of unposted) t.ambiguous = txns.some((x) => done.has(x) && x.last4 === t.last4 && near(x.amount, t.amount));
    const imStart = im && im.start, imEnd = im && im.end;
    const inReport = (l) => im && l.posted != null && l.posted >= imStart - 2 * DAY && l.posted <= imEnd + 45 * DAY;
    const extra = live.filter((l) => !used.has(l));
    return {
      matches, unposted,
      // Posted in Arietis with an InstaMed ref, inside the InstaMed report's dates, but no InstaMed transaction for it.
      notInInstamed: extra.filter((l) => l.refType === 'IM' && inReport(l)),
      // Other patient payments (cash, check, another card system, or before the InstaMed report starts).
      other: extra.filter((l) => !(l.refType === 'IM' && inReport(l))),
      reversed: lines.filter((l) => l.reversed),
      postedTotal: sum(live, (l) => l.pay),
    };
  }

  // ---------- benefits learned from processed claims ----------
  const median = (xs) => { if (!xs.length) return null; const s = [...xs].sort((a, b) => a - b); const m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : r2((s[m - 1] + s[m]) / 2); };
  const mode = (xs) => { const c = new Map(); for (const x of xs) c.set(x, (c.get(x) || 0) + 1); return [...c].sort((a, b) => b[1] - a[1])[0]; };
  /**
   * How the plan treated each class of service (TMS, Esketamine, Visit) in the plan year: copay per visit day, the
   * coinsurance rate (coinsurance / (allowed - deductible)), how many were fully covered, plus deductible applied and
   * patient share (OOP) so far. allowed = billed - contractual write-off.
   */
  function learnBenefits(charges) {
    const by = {};
    for (const c of charges.filter((x) => x.processed)) {
      const k = classOf(c.code);
      const b = by[k] || (by[k] = { cls: k, n: 0, copays: new Map(), rates: [], covered: 0, allowed: {} });
      b.n++;
      const allowed = r2(c.billed - c.contractual);
      (b.allowed[c.code] = b.allowed[c.code] || []).push({ allowed, dos: c.dos });
      if (c.copay > 0) b.copays.set(c.dos, r2((b.copays.get(c.dos) || 0) + c.copay));
      if (c.coins > 0 && allowed - c.ded > 0) b.rates.push(Math.round((c.coins / (allowed - c.ded)) * 1000) / 1000);
      if (c.resp === 0 && c.insPaid > 0) b.covered++;
    }
    const classes = {};
    for (const b of Object.values(by)) {
      const cp = mode([...b.copays.values()]);
      const rate = b.rates.length ? median(b.rates) : null;
      classes[b.cls] = {
        cls: b.cls, claims: b.n, copay: cp ? cp[0] : null, copayDays: b.copays.size, coins: rate, coinsClaims: b.rates.length, covered: b.covered,
        rule: cp && cp[1] >= b.rates.length ? 'copay' : rate != null ? 'coins' : b.covered ? 'covered' : 'unknown',
        allowed: Object.fromEntries(Object.entries(b.allowed).map(([code, xs]) => [code, xs.sort((a, z) => z.dos - a.dos)[0].allowed])),
        allowedAll: Object.fromEntries(Object.entries(b.allowed).map(([code, xs]) => [code, [...new Set(xs.map((x) => x.allowed))]])),
      };
    }
    const proc = charges.filter((x) => x.processed);
    return { classes, dedApplied: sum(proc, (c) => c.ded), oopMet: sum(proc, (c) => c.resp), processed: proc.length };
  }

  // ---------- plan year ----------
  /** The plan year holding `day`: from Sohar's coverage start (month and day) if known, else January 1. */
  function planYearOf(day, coverageStart) {
    const d = new Date(day);
    const cs = coverageStart ? C.parseDate(coverageStart) : null;
    const m = cs ? cs.getUTCMonth() : 0, dd = cs ? cs.getUTCDate() : 1;
    let start = Date.UTC(d.getUTCFullYear(), m, dd);
    if (start > day) start = Date.UTC(d.getUTCFullYear() - 1, m, dd);
    const end = Date.UTC(new Date(start).getUTCFullYear() + 1, m, dd) - DAY;
    return { start, end };
  }

  // ---------- re-quote ----------
  /**
   * Patient share for visits in date order, carrying the deductible and OOP max forward. items: [{day, lines: [{code,
   * allowed}]}]. inp: {deductible, dedMet, oopMax (0 = none), oopMet, cls: {TMS|Visit|Esketamine: {rule, copay, coins}}}.
   * A copay class pays its copay once per day (no deductible); a coinsurance class pays the deductible first, then
   * coinsurance; a covered class pays nothing; an unknown class uses the plan's deductible and coinsurance (fallback).
   */
  function requote(items, inp) {
    let dedLeft = Math.max(0, (num(inp.deductible) || 0) - (num(inp.dedMet) || 0));
    const oopMax = num(inp.oopMax) || 0;
    let oopLeft = oopMax > 0 ? Math.max(0, oopMax - (num(inp.oopMet) || 0)) : Infinity;
    const out = [];
    for (const it of [...items].sort((a, b) => a.day - b.day)) {
      const copaid = new Set();
      let ded = 0, coins = 0, copay = 0, unknown = [];
      for (const l of it.lines) {
        if (l.allowed == null) { unknown.push(l.code); continue; }
        const k = classOf(l.code), c = (inp.cls && inp.cls[k]) || { rule: 'unknown' };
        const rule = c.rule === 'unknown' || !c.rule ? (inp.fallback || { rule: 'coins', coins: 0 }).rule : c.rule;
        const rate = c.coins != null ? c.coins : (inp.fallback && inp.fallback.coins) || 0;
        if (rule === 'copay') { if (!copaid.has(k)) { copaid.add(k); copay += Math.min(l.allowed, num(c.copay ?? (inp.fallback && inp.fallback.copay)) || 0); } continue; }
        if (rule === 'covered') continue;
        const d = Math.min(l.allowed, dedLeft); dedLeft = r2(dedLeft - d); ded += d;
        coins += (l.allowed - d) * rate;
      }
      let patient = r2(ded + coins + copay);
      const capped = patient > oopLeft;
      if (capped) patient = r2(oopLeft);
      if (oopLeft !== Infinity) oopLeft = r2(oopLeft - patient);
      out.push({ ...it, ded: r2(ded), coins: r2(coins), copay: r2(copay), patient, capped, oopLeft: oopLeft === Infinity ? null : oopLeft, dedLeft, unknown });
    }
    return out;
  }

  // ---------- the audit ----------
  const SEV = { high: 0, medium: 1, low: 2, info: 3 };
  /**
   * The whole audit for one patient. o: {who: {mrn, lasfir}, instamed, arietis, appointments (parsed, already this
   * patient's visits), sohar (the record's financial row: Deductible, Deductible Remaining, OOP Max, OOP Remaining,
   * Coinsurance, Copay, Sohar Date, Coverage Start), flow (Flow rows), booked ([{day, time, event}] upcoming), codeRows,
   * rateOf(code) -> {rate, basis} | null, today (ms), plan: {start, end} override, inputs (overrides), confirms
   * ({isoDay: true}), quotes ({isoDay: amount})}.
   */
  function runAudit(o) {
    const today = o.today != null ? Date.UTC(new Date(o.today).getUTCFullYear(), new Date(o.today).getUTCMonth(), new Date(o.today).getUTCDate()) : null;
    const ar = o.arietis || null, im = o.instamed || null, ap = o.appointments || null;
    const sohar = o.sohar || {};
    const anomalies = [];
    const flag = (sev, source, text, day, extra) => anomalies.push({ sev, source, text, day: day ?? null, ...(extra || {}) });
    const ref = today ?? (ar && ar.asOf) ?? Date.now();
    const plan = o.plan && o.plan.start != null ? { start: o.plan.start, end: o.plan.end } : planYearOf(ref, sohar['Coverage Start']);
    const inPlan = (d) => d != null && d >= plan.start && d <= plan.end;
    const charges = ar ? ar.charges : [];
    const asOf = (ar && ar.asOf) || ref;
    const rate = (code) => (o.rateOf ? o.rateOf(code) : null);

    // 1. each charge
    for (const c of charges) {
      c.checks = [];
      const chk = (sev, text, source) => { c.checks.push({ sev, text }); flag(sev, source || 'Arietis', `${c.code}: ${text}`, c.dos); };
      if (!near(c.calcBalance, c.balance, 0.01)) chk('high', `balance ${usd(c.balance)} doesn't add up: billed ${usd(c.billed)} less posted amounts is ${usd(c.calcBalance)}`);
      if (c.balance < -0.005) chk('medium', `credit of ${usd(-c.balance)} on this charge`);
      if (c.processed) {
        const allowed = r2(c.billed - c.contractual);
        c.allowed = allowed;
        if (c.expected != null && !near(allowed, c.expected, 0.5) && c.insPaid > 0) chk('medium', `processed at ${usd(allowed)} allowed; Arietis expected ${usd(c.expected)}`);
        const gap = r2(allowed - c.insPaid - c.resp - c.otherAdj);
        // No copay / coinsurance line, but the patient paid exactly what the plan left: a share that wasn't itemized.
        const unitemized = Math.abs(gap) > 0.5 && c.resp === 0 && c.insPaid > 0 && near(c.insPaid + c.patPaid, allowed, 0.5);
        if (unitemized) chk('low', `no patient share line, but the plan paid ${usd(c.insPaid)} and the patient paid ${usd(c.patPaid)} of the ${usd(allowed)} allowed: likely a ${usd(c.patPaid)} copay that wasn't itemized`);
        else if (c.insPaid + c.resp > 0 && Math.abs(gap) > 0.5) chk(Math.abs(gap) < 1 ? 'low' : 'medium', `plan paid ${usd(c.insPaid)} + patient share ${usd(c.resp)} is ${usd(c.insPaid + c.resp)}, not the ${usd(allowed)} allowed`);
        if (c.insPaid === 0 && c.resp === 0) chk('high', `no payment and no patient share: denied or written off${c.carcs.length ? ` (${c.carcs.join(', ')})` : ''}`);
        const odd = c.carcs.filter((x) => x !== 'CO45');
        if (odd.length && c.insPaid > 0) chk('medium', `adjustment reason ${odd.join(', ')}: check the remit`);
        if (!unitemized && c.patPaid > c.resp + 0.005) chk('medium', `patient paid ${usd(c.patPaid)} on a ${usd(c.resp)} patient share`);
        const due = r2(c.resp - c.patPaid);
        if (!unitemized && !near(due, c.balance, 0.01) && near(c.calcBalance, c.balance, 0.01)) chk(Math.abs(due - c.balance) < 1 ? 'low' : 'medium', `balance ${usd(c.balance)}, but patient share less patient payments is ${usd(due)}`);
      } else {
        c.allowed = null;
        const age = c.firstClaim != null ? Math.floor((asOf - c.firstClaim) / DAY) : null;
        if (!c.claims.length) { if (c.dos != null && asOf - c.dos > 7 * DAY) chk('high', 'no claim sent'); else c.checks.push({ sev: 'info', text: 'no claim sent yet' }); }
        else if (age > 30) chk('medium', `claim sent ${age} days ago, not processed`);
        else c.checks.push({ sev: 'info', text: `claim sent ${age == null ? '' : `${age} day${age === 1 ? '' : 's'} ago`}, not processed yet` });
      }
      const fs = rate(c.code);
      // rateOf gives the MD rate and the NP rate (alt): Arietis matching either one is fine.
      c.feeRate = fs ? (fs.alt != null && c.expected != null && near(fs.alt, c.expected, 0.5) ? fs.alt : fs.rate) : null;
      if (fs && fs.rate != null && c.expected != null && !near(fs.rate, c.expected, 0.5) && !(fs.alt != null && near(fs.alt, c.expected, 0.5))) {
        chk('medium', `Arietis expects ${usd(c.expected)}; the fee schedule says ${usd(fs.rate)}${fs.alt != null && !near(fs.alt, fs.rate) ? ` (MD) or ${usd(fs.alt)} (NP)` : ''} (${fs.basis})`, 'Fee schedule');
      }
    }

    // 2. kept visits against charges, by date of service
    const visits = ap ? ap.visits : [];
    const keptDays = new Set(visits.filter((v) => v.kept).map((v) => v.day));
    const dayKeys = new Set([...charges.map((c) => c.dos), ...visits.filter((v) => v.kept || charges.some((c) => c.dos === v.day)).map((v) => v.day)].filter((d) => d != null && (d <= (today ?? Infinity))));
    const days = [...dayKeys].sort((a, b) => b - a).map((day) => {
      const vs = visits.filter((v) => v.day === day), cs = charges.filter((c) => c.dos === day);
      const kept = vs.filter((v) => v.kept);
      const out = { day, iso: iso(day), visits: vs, charges: cs, inPlan: inPlan(day), checks: [], confirmed: !!(o.confirms && o.confirms[iso(day)]) };
      const chk = (sev, source, text) => { out.checks.push({ sev, source, text }); flag(out.confirmed && source === 'Codes' ? 'info' : sev, source, text, day, { confirmed: out.confirmed && source === 'Codes' }); };
      if (ap && ar) {
        const inAppts = day >= ap.start && day <= ap.end;
        if (cs.length && !kept.length && inAppts) {
          const other = vs.find((v) => !v.kept);
          chk('high', 'NextGen', other ? `billed (${cs.map((c) => c.code).join(', ')}) for a visit marked "${other.status}" in NextGen` : `billed (${cs.map((c) => c.code).join(', ')}) with no kept visit on the NextGen report`);
        }
        if (kept.length && !cs.length) {
          const recent = asOf - day <= 7 * DAY;
          chk(recent ? 'low' : 'high', 'NextGen', `kept ${kept.map((v) => v.event).join(' + ')}, ${recent ? 'not billed yet' : 'not billed in Arietis'}`);
        }
      }
      if (kept.length && cs.length) for (const x of codeChecks(kept, cs, o.codeRows)) chk(x.sev, 'Codes', x.text);
      // as quoted: the staff's quote for the day, else a Sohar Flow estimate for it
      const fl = (o.flow || []).find((f) => dayOf(f['Appt Date']) === day && num(f['Sohar Est Patient']) != null);
      const typed = o.quotes && num(o.quotes[iso(day)]);
      out.quote = typed != null ? { amount: typed, source: 'typed' } : fl ? { amount: num(fl['Sohar Est Patient']), source: /sandbox/i.test(str(fl.Note) + str(fl['Run Status'])) ? 'Sohar Flow (sandbox)' : 'Sohar Flow' } : null;
      out.processed = cs.length > 0 && cs.every((c) => c.processed);
      out.resp = sum(cs, (c) => c.resp); out.patPaid = sum(cs, (c) => c.patPaid); out.balance = sum(cs, (c) => c.balance);
      if (out.quote && out.processed && !near(out.quote.amount, out.resp, 1)) chk('medium', 'Quote', `quoted ${usd(out.quote.amount)} (${out.quote.source}); the plan's patient share is ${usd(out.resp)}`);
      return out;
    });
    // kept visits in the plan year outside the Arietis export's dates are not checked
    if (ap && !ar) flag('info', 'NextGen', `${keptDays.size} kept visit day${keptDays.size === 1 ? '' : 's'} read; add the Arietis account detail to check what was billed.`);

    // 3. payments
    const pay = matchPayments(im, ar);
    if (im) {
      if (im.reportTotal != null && !near(im.reportTotal, sum(im.txns, (t) => t.amount), 0.01)) flag('low', 'InstaMed', `report total ${usd(im.reportTotal)} differs from its rows (${usd(sum(im.txns, (t) => t.amount))})`);
      for (const t of im.txns.filter((x) => !x.counted)) flag('info', 'InstaMed', `${usd(t.amount)} ${t.status || t.action}: not counted`, t.day);
      // Arietis posts card payments a few days late: a payment in the export's last 7 days may just not be posted yet.
      if (ar) for (const t of pay.unposted) {
        const recent = ar.asOf != null && t.day > ar.asOf - 7 * DAY;
        t.recent = recent;
        flag(recent ? 'low' : 'high', 'InstaMed', `payment of ${usd(t.amount)} (card ending ${t.last4}) is not posted in Arietis${recent ? ` yet (the Arietis export runs to ${us(ar.asOf)})` : ''}${t.ambiguous ? '; other payments of the same amount on this card did post, so the missing one may be another date (the latest is assumed)' : ''}`, t.day);
      }
      for (const l of pay.notInInstamed) flag('medium', 'Arietis', `patient payment ${usd(l.pay)} posted as InstaMed (ref ${l.ref}, batch ${l.batch}) has no InstaMed transaction`, l.posted);
    }

    // 4. benefits
    const planCharges = charges.filter((c) => inPlan(c.dos));
    const learned = learnBenefits(planCharges);
    const sDate = dayOf(sohar['Sohar Date']);
    const sohOop = num(sohar['OOP Max']) != null && num(sohar['OOP Remaining']) != null ? r2(num(sohar['OOP Max']) - num(sohar['OOP Remaining'])) : null;
    const sohDed = num(sohar.Deductible) != null && num(sohar['Deductible Remaining']) != null ? r2(num(sohar.Deductible) - num(sohar['Deductible Remaining'])) : null;
    // Patient share our claims added after Sohar's check (by when Arietis posted it): Sohar's accumulator can't have it.
    const after = sDate != null ? planCharges.filter((c) => c.processed && c.respPosted != null && c.respPosted > sDate) : [];
    const oopAfter = sum(after, (c) => c.resp), dedAfter = sum(after, (c) => c.ded);
    const claimsBefore = sDate != null ? sum(planCharges.filter((c) => c.processed && c.respPosted != null && c.respPosted <= sDate), (c) => c.resp) : null;
    if (sohOop != null && claimsBefore != null && claimsBefore > sohOop + 1) flag('medium', 'Sohar', `Sohar showed ${usd(sohOop)} of the OOP max met on ${us(sDate)}, but our claims posted by then already add ${usd(claimsBefore)}: re-verify or check the posting`, sDate);
    const sohCoins = F ? F.coinsRate(sohar.Coinsurance) : null;
    for (const k of Object.keys(learned.classes)) {
      const c = learned.classes[k];
      if (c.rule === 'coins' && sohCoins != null && Math.abs(c.coins - sohCoins) > 0.011) flag('medium', 'Sohar', `${k} claims processed at ${Math.round(c.coins * 1000) / 10}% coinsurance; Sohar says ${Math.round(sohCoins * 1000) / 10}%`);
      if (c.rule === 'copay' && num(sohar.Copay) != null && !near(c.copay, num(sohar.Copay), 0.5)) flag('low', 'Sohar', `${k} claims took a ${usd(c.copay)} copay; Sohar's copay is ${usd(num(sohar.Copay))}`);
      const multi = Object.entries(c.allowedAll).filter(([, xs]) => xs.length > 1);
      for (const [code, xs] of multi) flag('low', 'Arietis', `${code} processed at more than one allowed amount this plan year: ${xs.map(usd).join(', ')}`);
    }
    const dflt = {
      deductible: num(sohar.Deductible) ?? '', dedMet: Math.max(learned.dedApplied, sohDed != null ? r2(sohDed + dedAfter) : 0),
      oopMax: num(sohar['OOP Max']) ?? '', oopMet: Math.max(learned.oopMet, sohOop != null ? r2(sohOop + oopAfter) : 0),
    };
    const inputs = { ...dflt, ...Object.fromEntries(Object.entries(o.inputs || {}).filter(([k, v]) => k in dflt && v !== '' && v != null).map(([k, v]) => [k, num(v)])) };
    const cls = {};
    for (const k of ['TMS', 'Visit', 'Esketamine']) {
      const l = learned.classes[k], ov = (o.inputs && o.inputs.cls && o.inputs.cls[k]) || {};
      cls[k] = { rule: ov.rule || (l ? l.rule : 'unknown'), copay: ov.copay != null && ov.copay !== '' ? num(ov.copay) : l && l.copay != null ? l.copay : l && l.rule !== 'copay' ? null : num(sohar.Copay),
        coins: ov.coins != null && ov.coins !== '' ? num(ov.coins) / 100 : l && l.coins != null ? l.coins : sohCoins, learned: !!l, from: ov.rule || ov.copay != null || ov.coins != null ? 'typed' : l ? 'claims' : 'Sohar' };
    }
    inputs.cls = cls;
    inputs.fallback = { rule: sohCoins ? 'coins' : num(sohar.Copay) ? 'copay' : 'coins', coins: sohCoins || 0, copay: num(sohar.Copay) || 0 };

    // 5. calculator check: the West Coast calculator's math (fees.calcEstimate) with Sohar's benefits and the charge mode
    // staff pick, on each processed visit day's allowed amounts, carrying the deductible and OOP met from the plan year
    // start. One finding sums up the days where the plan charged differently; the table shows each day.
    const soharKnown = sohCoins != null || num(sohar.Copay) != null;
    const calcMode = (o.inputs && o.inputs.calcMode) || o.calcMode || 'all';
    const procDays = days.filter((d) => d.inPlan && d.processed).sort((a, b) => a.day - b.day);
    const calc = { mode: calcMode, days: 0, off: 0, calcTotal: 0, planTotal: 0 };
    if (soharKnown && F) {
      let dedMet = 0, oopMet = 0;
      for (const d of procDays) {
        const b = { planType: 'commercial', mode: calcMode, deductible: num(sohar.Deductible) || 0, deductibleMet: dedMet, coins: sohCoins || 0, copay: num(sohar.Copay) || 0,
          oopMax: num(sohar['OOP Max']) || 0, oopMet };
        const e = F.calcEstimate(d.charges.map((c) => ({ allowed: c.allowed, svcCoins: null })), b);
        d.calc = e.patient;
        dedMet = r2(dedMet + e.totalDed); oopMet = r2(oopMet + e.patient);
        calc.days++; calc.calcTotal = r2(calc.calcTotal + e.patient); calc.planTotal = r2(calc.planTotal + d.resp);
        if (!near(e.patient, d.resp, 1)) calc.off++;
      }
      if (calc.off) flag(calc.off > calc.days / 2 ? 'medium' : 'low', 'Calculator', `the calculator with Sohar's benefits (${(F.CALC_MODES.find(([k]) => k === calcMode) || [, calcMode])[1].toLowerCase()}) differs from the plan on ${calc.off} of ${calc.days} processed visit days: ${usd(calc.calcTotal)} quoted vs ${usd(calc.planTotal)} charged`);
    }

    // 6. re-quote: pending claims and booked visits, from what the claims show
    const pending = charges.filter((c) => !c.processed && inPlan(c.dos));
    const learnedAllowed = (code) => { const k = classOf(code); const a = learned.classes[k] && learned.classes[k].allowed[code]; if (a != null) return { allowed: a, src: 'processed claims' }; const fs = rate(code); return fs && fs.rate != null ? { allowed: fs.rate, src: 'fee schedule' } : null; };
    const items = [];
    for (const d of [...new Set(pending.map((c) => c.dos))]) {
      const cs = pending.filter((c) => c.dos === d);
      items.push({ day: d, kind: 'Pending claim', event: cs.map((c) => c.desc || c.code).join(' + '), lines: cs.map((c) => { const a = learnedAllowed(c.code); return { code: c.code, allowed: c.expected ?? (a && a.allowed), src: c.expected != null ? 'Arietis expected' : a ? a.src : '' }; }), paid: sum(cs, (c) => c.patPaid) });
    }
    const billedDays = new Set(charges.map((c) => c.dos));
    for (const b of o.booked || []) {
      const day = dayOf(b.date ?? b.day);
      if (day == null || billedDays.has(day) || !inPlan(day) || (today != null && day < today)) continue;
      const e = expectedFor(b.event, o.codeRows);
      items.push({ day, time: b.time, kind: 'Booked', event: b.event, lines: e.lines.map((any) => { const a = learnedAllowed(any[0]); return { code: any.join('/'), allowed: a ? a.allowed : null, src: a ? a.src : 'no rate' }; }).map((l) => ({ ...l, code: l.code.split('/')[0] })), paid: 0, note: e.lines.length ? '' : e.note });
    }
    const quoted = requote(items, inputs).map((q) => {
      const fl = (o.flow || []).find((f) => dayOf(f['Appt Date']) === q.day && num(f['Sohar Est Patient']) != null);
      const flowAmt = fl ? num(fl['Sohar Est Patient']) : null;
      const sandbox = fl ? /sandbox/i.test(str(fl.Note) + str(fl['Run Status']) + str(fl['Estimate Note'])) : false;
      if (flowAmt != null && !sandbox && !near(flowAmt, q.patient, 1)) flag('medium', 'Sohar', `re-quote ${usd(q.patient)} for the ${us(q.day)} ${q.kind === 'Booked' ? 'visit' : 'claim'}; Sohar Flow estimated ${usd(flowAmt)}`, q.day);
      return { ...q, flow: flowAmt, sandbox, due: r2(q.patient - (q.paid || 0)) };
    });

    // 7. balances
    const proc = charges.filter((c) => c.processed), pend = charges.filter((c) => !c.processed);
    const pendingPatient = sum(quoted.filter((q) => q.kind === 'Pending claim'), (q) => q.due);
    const pendOutside = pend.filter((c) => !inPlan(c.dos));
    if (pendOutside.length) flag('medium', 'Arietis', `${pendOutside.length} unprocessed charge${pendOutside.length === 1 ? '' : 's'} from before this plan year (${usd(sum(pendOutside, (c) => c.balance))}): not re-quoted`);
    const unposted = ar ? sum(pay.unposted, (t) => t.amount) : 0;
    const balance = {
      arietis: sum(charges, (c) => c.balance),
      processed: sum(proc, (c) => c.balance), earlier: sum(proc.filter((c) => !inPlan(c.dos)), (c) => c.balance),
      pendingBilled: sum(pend, (c) => c.balance), pendingPatient, pendingCount: pend.length,
      unposted, true: r2(sum(proc, (c) => c.balance) + pendingPatient - unposted),
      booked: sum(quoted.filter((q) => q.kind === 'Booked'), (q) => q.patient),
      instamedCollected: im ? im.total : null, arietisPatientPaid: pay.postedTotal,
      // What each balance is made of (v0.24), for opening it in the patient view: the charges with their dates of
      // service and PR amounts; pending claims with their re-quote; the InstaMed payments not posted (no card digits).
      lines: {
        processed: proc.filter((c) => Math.abs(c.balance) > 0.005).map(chargeRow),
        earlier: proc.filter((c) => !inPlan(c.dos) && Math.abs(c.balance) > 0.005).map(chargeRow),
        pending: pend.map((c) => { const q = quoted.find((x) => x.kind === 'Pending claim' && x.day === c.dos); return { ...chargeRow(c), quote: q ? { patient: q.patient, paid: q.paid || 0, due: q.due } : null }; }),
        unposted: ar ? pay.unposted.map((t) => ({ day: t.day, amount: t.amount, type: t.type })) : [],
      },
    };
    anomalies.sort((a, b) => SEV[a.sev] - SEV[b.sev] || (b.day ?? 0) - (a.day ?? 0));
    return { plan, asOf, charges, days, payments: pay, learned, inputs, defaults: dflt, calc, sohar: { oopMet: sohOop, dedMet: sohDed, date: sDate, coins: sohCoins, copay: num(sohar.Copay), oopAfter, dedAfter },
      requote: quoted, balance, anomalies, counts: { high: anomalies.filter((a) => a.sev === 'high').length, medium: anomalies.filter((a) => a.sev === 'medium').length, low: anomalies.filter((a) => a.sev === 'low').length } };
  }

  /** Plain text for a ticket or note: LASFIR and MRN, no names, no card numbers. */
  function summaryText(a, who, sources) {
    const L = [];
    L.push(`Account audit: ${who.lasfir}${who.mrn ? ` (MRN ${who.mrn})` : ''}, plan year ${us(a.plan.start)} to ${us(a.plan.end)}`);
    L.push(`Sources: ${sources.join('; ')}`);
    const b = a.balance;
    L.push(`Arietis balance ${usd(b.arietis)}. True patient balance ${usd(b.true)} = processed ${usd(b.processed)}${b.earlier ? ` (incl. ${usd(b.earlier)} from earlier plan years)` : ''} + pending claims patient share ${usd(b.pendingPatient)} (re-quote; Arietis shows ${usd(b.pendingBilled)} billed on ${b.pendingCount}) - InstaMed payments not posted ${usd(b.unposted)}.`);
    if (b.booked) L.push(`Booked visits re-quote: ${usd(b.booked)}.`);
    const i = a.inputs;
    L.push(`Benefits used: deductible ${i.deductible === '' ? 'unknown' : usd(i.deductible)} (met ${usd(i.dedMet)}), OOP max ${i.oopMax === '' ? 'unknown' : usd(i.oopMax)} (met ${usd(i.oopMet)}); ${Object.entries(i.cls).map(([k, c]) => `${k}: ${c.rule === 'copay' ? `${usd(c.copay)} copay` : c.rule === 'coins' ? `${Math.round((c.coins || 0) * 1000) / 10}% coinsurance` : c.rule} (${c.from})`).join(', ')}.`);
    const top = a.anomalies.filter((x) => x.sev !== 'info');
    L.push(`Anomalies: ${a.counts.high} high, ${a.counts.medium} medium, ${a.counts.low} low.`);
    for (const x of top.slice(0, 40)) L.push(`- [${x.sev}] ${x.source}${x.day != null ? ` ${us(x.day)}` : ''}: ${x.text.replace(/card ending \d+/g, 'card on file')}`);
    if (top.length > 40) L.push(`- and ${top.length - 40} more`);
    return L.join('\n');
  }

  // ---------- saved audits and the export (v0.23.1) ----------
  // The last audit per patient is saved on the hidden Account Audits sheet (v0.26; two sheets before): one summary row (AUDIT_COLS) and its lines
  // (AUDIT_LINE_COLS: findings, visits, re-quote, payments). LASFIR and MRN only: no names, file names or card numbers.
  // Saving again replaces that patient's rows. The export compiles every saved audit into one workbook.
  const AUDIT_COLS = ['Key', 'MRN', 'LASFIR', 'Saved', 'By', 'Plan Year', 'Sources', 'Arietis Through', 'Arietis Balance', 'Processed Balance',
    'Earlier Plan Years', 'Pending Claims', 'Pending Billed', 'Pending Patient Share', 'InstaMed Not Posted', 'True Balance', 'Booked Re-quote',
    'InstaMed Collected', 'Arietis Patient Paid', 'Deductible', 'Deductible Met', 'OOP Max', 'OOP Met', 'TMS', 'Visit', 'Esketamine',
    'Calculator', 'High', 'Medium', 'Low', 'Codes Confirmed'];
  const AUDIT_LINE_COLS = ['Key', 'MRN', 'LASFIR', 'Saved', 'Part', 'Date', 'What', 'Codes', 'Allowed', 'Plan Paid', 'Patient Share', 'Calculator',
    'Quoted', 'Paid', 'Balance', 'Status', 'Severity', 'Source', 'Detail'];
  const AUDIT_MONEY = new Set(['Arietis Balance', 'Processed Balance', 'Earlier Plan Years', 'Pending Billed', 'Pending Patient Share', 'InstaMed Not Posted',
    'True Balance', 'Booked Re-quote', 'InstaMed Collected', 'Arietis Patient Paid', 'Deductible', 'Deductible Met', 'OOP Max', 'OOP Met',
    'Allowed', 'Plan Paid', 'Patient Share', 'Calculator', 'Quoted', 'Paid', 'Balance']);
  /** Card digits out of any text that is saved or exported. */
  const scrub = (s) => str(s).replace(/card ending \d+/gi, 'card on file').replace(/(?:…|\.\.\.)\d{4}\b/g, '').replace(/\b(?:IM|CC)\s*Ref\s*#\s*\w+/gi, 'card ref');
  const m2 = (n) => (n == null || n === '' || !isFinite(n) ? '' : r2(n));
  const clsText = (k, c) => (!c ? '' : c.rule === 'copay' ? `${usd(c.copay)} copay (${c.from})` : c.rule === 'coins' ? `${Math.round((c.coins || 0) * 1000) / 10}% coinsurance (${c.from})` : `${c.rule} (${c.from})`);
  /**
   * The rows to save for one audit: {summary, lines}. who: {key, mrn, lasfir}; o: {by, when, sources}.
   */
  function auditRecord(a, who, o) {
    const base = { Key: str(who.key || who.mrn || who.lasfir), MRN: str(who.mrn), LASFIR: str(who.lasfir), Saved: str(o.when) };
    const b = a.balance, i = a.inputs;
    const summary = { ...base, By: str(o.by), 'Plan Year': `${us(a.plan.start)} to ${us(a.plan.end)}`, Sources: (o.sources || []).join('; '),
      'Arietis Through': a.charges.length ? us(a.asOf) : '', 'Arietis Balance': m2(b.arietis), 'Processed Balance': m2(b.processed), 'Earlier Plan Years': m2(b.earlier),
      'Pending Claims': b.pendingCount, 'Pending Billed': m2(b.pendingBilled), 'Pending Patient Share': m2(b.pendingPatient), 'InstaMed Not Posted': m2(b.unposted),
      'True Balance': m2(b.true), 'Booked Re-quote': m2(b.booked), 'InstaMed Collected': m2(b.instamedCollected), 'Arietis Patient Paid': m2(b.arietisPatientPaid),
      Deductible: m2(num(i.deductible)), 'Deductible Met': m2(num(i.dedMet)), 'OOP Max': m2(num(i.oopMax)), 'OOP Met': m2(num(i.oopMet)),
      TMS: clsText('TMS', i.cls.TMS), Visit: clsText('Visit', i.cls.Visit), Esketamine: a.learned.classes.Esketamine ? clsText('Esketamine', i.cls.Esketamine) : '',
      Calculator: a.calc.days ? `${(F && (F.CALC_MODES.find(([k]) => k === a.calc.mode) || [])[1]) || a.calc.mode}: ${usd(a.calc.calcTotal)} vs ${usd(a.calc.planTotal)} charged, ${a.calc.off} of ${a.calc.days} days differ` : '',
      High: a.counts.high, Medium: a.counts.medium, Low: a.counts.low, 'Codes Confirmed': a.days.filter((d) => d.confirmed).length };
    const lines = [];
    const add = (part, o2) => lines.push({ ...Object.fromEntries(AUDIT_LINE_COLS.map((c) => [c, ''])), ...base, Part: part, ...o2 });
    for (const x of a.anomalies.filter((y) => y.sev !== 'info' || y.confirmed)) add('Finding', { Date: x.day != null ? us(x.day) : '', What: scrub(x.text), Severity: x.confirmed ? 'confirmed' : x.sev, Source: x.source, Status: x.confirmed ? 'Codes confirmed' : '' });
    for (const d of a.days.filter((y) => y.inPlan || y.checks.length)) {
      const cs = d.charges;
      add('Visit', { Date: us(d.day), What: d.visits.map((v) => `${v.event} (${v.status})`).join('; '), Codes: cs.map((c) => `${c.code}${c.mods.length ? `-${c.mods.join('-')}` : ''}`).join(', '),
        Allowed: cs.length ? m2(cs.reduce((t, c) => t + (c.allowed ?? c.expected ?? 0), 0)) : '', 'Plan Paid': d.processed ? m2(cs.reduce((t, c) => t + c.insPaid, 0)) : '',
        'Patient Share': d.processed ? m2(d.resp) : '', Calculator: m2(d.calc), Quoted: d.quote ? m2(d.quote.amount) : '', Paid: m2(d.patPaid), Balance: m2(d.balance),
        Status: `${!cs.length ? 'Not billed' : d.processed ? 'Processed' : 'Pending'}${d.confirmed ? ' · codes OK' : ''}`,
        Detail: scrub([...d.checks.map((x) => x.text), ...cs.flatMap((c) => c.checks.filter((x) => x.sev !== 'info').map((x) => `${c.code}: ${x.text}`))].join('; ')) });
    }
    for (const q of a.requote) {
      add('Re-quote', { Date: us(q.day), What: `${q.kind}: ${q.event || ''}`.trim(), Codes: q.lines.map((l) => `${l.code} ${l.allowed == null ? 'no rate' : usd(l.allowed)} (${l.src})`).join(', '),
        'Patient Share': m2(q.patient), Paid: m2(q.paid), Balance: m2(q.due),
        Status: q.flow == null ? '' : q.sandbox ? 'Sohar Flow (sandbox)' : Math.abs(q.flow - q.patient) > 1 ? 'Sohar Flow differs' : 'Sohar Flow same',
        Detail: [q.flow != null ? `Sohar Flow ${usd(q.flow)}` : '', q.capped ? 'reaches the OOP max' : '', q.oopLeft != null ? `OOP left ${usd(q.oopLeft)}` : ''].filter(Boolean).join('; ') });
    }
    const p = a.payments;
    for (const m of p.matches) for (const t of m.txns) {
      add('Payment', { Date: us(t.day), What: 'InstaMed payment', Paid: m2(t.amount), Status: 'Posted in Arietis',
        Detail: `posted ${[...new Set(m.lines.map((l) => us(l.posted)))].join(', ')}; batch ${[...new Set(m.lines.map((l) => l.batch))].join(', ')}${m.lines.length > 1 ? `; split across ${m.lines.length} charges` : ''}` });
    }
    for (const t of p.unposted) add('Payment', { Date: us(t.day), What: 'InstaMed payment', Paid: m2(t.amount), Status: t.recent ? 'Not posted yet' : 'Not posted', Detail: t.ambiguous ? 'same amount as other payments on this card' : '' });
    for (const l of p.notInInstamed) add('Payment', { Date: us(l.posted), What: 'Arietis patient payment', Paid: m2(l.pay), Status: 'Not in the InstaMed report', Detail: `batch ${l.batch}` });
    lines.sort((x, y) => ['Finding', 'Visit', 'Re-quote', 'Payment'].indexOf(x.Part) - ['Finding', 'Visit', 'Re-quote', 'Payment'].indexOf(y.Part));
    return { summary, lines };
  }
  /** Saved audits with one patient's replaced by rec (its Key). saved: {summary, lines}. */
  function mergeAudits(saved, rec) {
    const k = str(rec.summary.Key);
    const keep = (rows) => (rows || []).filter((r) => str(r.Key) !== k);
    return { summary: [...keep(saved && saved.summary), rec.summary].sort((x, y) => str(x.LASFIR).localeCompare(str(y.LASFIR))), lines: [...keep(saved && saved.lines), ...rec.lines] };
  }
  // v0.26 (Brendan, 11/3: fewer tabs): both go on the one hidden Account Audits sheet. Part says what a row is: "Summary"
  // for the audit's summary row, else the line's part (Finding, Visit, Re-quote, Payment).
  const AUDIT_SHEET_COLS = [...new Set(['Key', 'MRN', 'LASFIR', 'Saved', 'Part', ...AUDIT_COLS, ...AUDIT_LINE_COLS])];
  /** Saved audits ({summary, lines}) as Account Audits sheet rows: each summary, then its lines. */
  function auditSheetRows(saved) {
    const blank = Object.fromEntries(AUDIT_SHEET_COLS.map((c) => [c, '']));
    const lines = (saved && saved.lines) || [];
    return ((saved && saved.summary) || []).flatMap((s) => [{ ...blank, ...s, Part: 'Summary' },
      ...lines.filter((l) => str(l.Key) === str(s.Key)).map((l) => ({ ...blank, ...l }))]);
  }
  /**
   * Account Audits sheet rows back to {summary, lines}. A row with no Part is a summary saved before v0.26; oldLines are
   * the rows of the old Account Audit Lines sheet, if it is still there.
   */
  function auditsFromSheet(rows, oldLines) {
    const pick = (cols, r) => Object.fromEntries(cols.map((c) => [c, r[c] == null ? '' : r[c]]));
    const isSummary = (r) => !str(r.Part) || str(r.Part) === 'Summary';
    const summary = (rows || []).filter(isSummary).map((r) => pick(AUDIT_COLS, r));
    const keys = new Set(summary.map((r) => str(r.Key)));
    const lines = (rows || []).filter((r) => !isSummary(r)).concat((oldLines || []).filter((r) => str(r.Part)))
      .filter((r) => keys.has(str(r.Key))).map((r) => pick(AUDIT_LINE_COLS, r));
    return { summary, lines };
  }
  /** A saved summary row for a patient key, or null. */
  const savedFor = (saved, key) => ((saved && saved.summary) || []).find((r) => str(r.Key) === str(key)) || null;
  /**
   * The export workbook spec (xlsx.buildWorkbook): Summary, Findings, Visits, Re-quote, Payments. keys: only these
   * patients (blank = every saved audit). Money columns go out as numbers.
   */
  function exportSpec(saved, keys) {
    const want = keys && keys.length ? new Set(keys.map(str)) : null;
    const pick = (rows) => (rows || []).filter((r) => !want || want.has(str(r.Key)));
    const money = (cols) => (r) => Object.fromEntries(cols.map((c) => [c, AUDIT_MONEY.has(c) && str(r[c]) !== '' && num(r[c]) != null ? num(r[c]) : str(r[c]) === '' ? '' : (/^(High|Medium|Low|Pending Claims|Codes Confirmed)$/.test(c) && num(r[c]) != null ? num(r[c]) : str(r[c]))]));
    const sCols = AUDIT_COLS.filter((c) => c !== 'Key');
    const part = (sheet, kind, cols) => { const cs = ['LASFIR', 'MRN', ...cols]; return { name: sheet, columns: cs, rows: pick(saved && saved.lines).filter((r) => r.Part === kind).map(money(cs)) }; };
    return { sheets: [
      { name: 'Summary', columns: sCols, rows: pick(saved && saved.summary).map(money(sCols)) },
      part('Findings', 'Finding', ['Date', 'Severity', 'Source', 'What', 'Status']),
      part('Visits', 'Visit', ['Date', 'What', 'Codes', 'Allowed', 'Plan Paid', 'Patient Share', 'Calculator', 'Quoted', 'Paid', 'Balance', 'Status', 'Detail']),
      part('Re-quote', 'Re-quote', ['Date', 'What', 'Codes', 'Patient Share', 'Paid', 'Balance', 'Status', 'Detail']),
      part('Payments', 'Payment', ['Date', 'What', 'Paid', 'Status', 'Detail']),
    ] };
  }
  /** "Account Audits 2026-11-02.xlsx", or "Account Audit - TESPAT 2026-11-02.xlsx" for one patient. */
  function exportName(lasfir, d) {
    const x = d || new Date();
    const day = `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`;
    return lasfir ? `Account Audit - ${str(lasfir).replace(/[^A-Z0-9]/gi, '')} ${day}.xlsx` : `Account Audits ${day}.xlsx`;
  }

  // ---------- demo (made-up patient, codes and amounts; the layouts mimic the real exports) ----------
  /**
   * Rows for the three files, shaped like the real exports, for demo mode and tests: {instamed, arietis, appointments}.
   * TMS daily with 30% coinsurance, a med visit with a $41 copay, pending claims, one unposted card payment, one charge
   * on a no-show day and one kept visit not billed.
   */
  function demoRows(last, first, mrn, todayMs) {
    const t0 = Math.floor(((todayMs != null ? todayMs : Date.now()) / DAY) + 25569);
    const ser = (back, frac) => t0 - back + (frac || 0);
    const nm = `${String(last).toUpperCase()}, ${String(first).toUpperCase()} `;
    const A = [['', '', '', '', '', '', 'Date', 'Patient', 'Provider', 'Ins', 'Code', 'Info', 'User', 'Description', '', 'Billed', 'Expected', 'Amount', 'Balance']];
    const line = (d, code, info, desc, amt, ins) => A.push(['', '', '', '', '', '', d, '', '', ins || '', code, info || '', 'DEMO/DEMO', desc, '', '', '', amt, '']);
    const charge = (back, code, billed, expected, balance) => A.push(['-', '', '', '', '', '', ser(back), nm, 'PROVIDER, DEMO', 'DEM101', code, '', 'DEMO', code === '90868' ? 'TCRANIAL MAGN STIM TX DELI' : 'OFFICE O/P EST MOD 30 MIN', '', billed.toFixed(2), expected.toFixed(2), '', balance.toFixed(2)]);
    const tms = [40, 39, 38, 37, 36, 33, 32, 31, 30, 29];
    tms.forEach((b, i) => {
      const paid = i < 7;
      const bal = paid ? (i < 4 ? 0 : i === 6 ? 105.57 : 0) : 2451;
      charge(b, '90868', 2451, 351.91, bal);
      line(ser(b - 2, 0.9), 'CLAIM-1E', '', `Claim DEMO${100 + i} (Batch # 9${i}) to DEMO HEALTH PLAN`, '0.00');
      if (!paid) return;
      line(ser(b - 9, 0.2), 'INSPMT', '', `ET Ref # DEMO${i} (Batch # 71${i}) from DEMO HEALTH PLAN`, '-246.34', 'DEM101 PMT');
      line(ser(b - 9, 0.2), 'COINS', '', 'CoInsurance  105.57', '0.00');
      line(ser(b - 9, 0.2), 'CONTWOFF', 'CO45', 'Charge exceeds fee schedule/maximum allowable or contracted/legi 2099.09', '-2099.09');
      if (i < 6) line(ser(b - 12, 0.5), 'GUARPMT', '', `IM Ref # 4242 (Batch # 81${i < 4 ? i : 4}) from ${nm.trim()}`, '-105.57');
    });
    // one charge whose balance doesn't add up (a mistyped adjustment)
    charge(28, '90868', 2451, 351.91, 2400);
    line(ser(26, 0.9), 'CLAIM-1E', '', 'Claim DEMO120 (Batch # 120) to DEMO HEALTH PLAN', '0.00');
    // med visit: 99214 + add-on, copay on the add-on line
    charge(35, '99214 - 95', 251.71, 131.91, 0);
    line(ser(33, 0.9), 'CLAIM-1E', '', 'Claim DEMO130 (Batch # 130) to DEMO HEALTH PLAN', '0.00');
    line(ser(25, 0.2), 'INSPMT', '', 'ET Ref # DEMO30 (Batch # 730) from DEMO HEALTH PLAN', '-131.91', 'DEM101 PMT');
    line(ser(25, 0.2), 'CONTWOFF', 'CO45', 'Charge exceeds fee schedule/maximum allowable or contracted/legi 119.80', '-119.80');
    charge(35, '90833 - 95', 151.91, 91.71, 0);
    line(ser(25, 0.2), 'INSPMT', '', 'ET Ref # DEMO30 (Batch # 730) from DEMO HEALTH PLAN', '-50.71', 'DEM101 PMT');
    line(ser(25, 0.2), 'COPAY', '', 'CoPayment 41.00', '0.00');
    line(ser(25, 0.2), 'CONTWOFF', 'CO45', 'Charge exceeds fee schedule/maximum allowable or contracted/legi 60.20', '-60.20');
    line(ser(24, 0.5), 'GUARPMT', '', `IM Ref # 4242 (Batch # 840) from ${nm.trim()}`, '-41.00');
    const I = [[], ['', '', '', 'Title'], ['', '', '', 'Payment Transactions'], [],
      ['Transaction ID Secondary', 'User', 'Patient Account #', 'Patient Last Name', 'Patient First Name', 'Workflow Status', 'Transaction Date', 'Type', 'Action', 'Amount', 'Effective Amount', 'Current Status', 'Card # - Last 4 Digits', 'Card Holder Name', 'Outlet Desc', 'Outlet']];
    const tx = (back, amt, status) => I.push(['', 'DEMO.api', '0009000000000001', String(last).toUpperCase(), String(first).toUpperCase(), 'Complete', ser(back, 0.45), 'Credit Card', 'Sale', amt, amt, status || 'SETTLED', '4242', `${first} ${last}`, 'DEMO PATIENT PORTAL', '0000-0001']);
    [0, 1, 2, 3].forEach((i) => tx(tms[i] - 12 + 1, 105.57)); // each posted in its own batch
    tx(tms[4] - 12 + 1, 211.14); // one payment posted as two lines (two charges, one batch)
    tx(25, 41);
    tx(34, 71.51); // never posted (older than the export, so a real gap)
    tx(18, 12.51, 'VOIDED');
    I.push(['Total:', '', '', '', '', '', '', '', '', 1, 1]);
    I[I.length - 1][9] = I[I.length - 1][10] = r2(I.slice(5, -1).reduce((a, r) => a + r[10], 0));
    const P = [['Pat Name', 'MRN', 'Appt Date', 'Beg Tm', 'Event', 'Status', 'Loc Name', 'Rendering']];
    const ap = (back, ev, st, tm) => P.push([`${last}, ${first}`, mrn, ser(back), tm || '10:00 A', ev, st, 'Demo Clinic', 'Provider Md, Demo']);
    tms.forEach((b) => ap(b, 'TMS Treatment Unilateral', 'Kept'));
    ap(35, 'Telepsych TMS F/U', 'Kept', '9:00 A');
    ap(28, 'TMS Treatment Unilateral', 'No Show');
    ap(26, 'TMS Treatment Unilateral', 'Kept');
    for (const b of [-1, -2, -3]) ap(b, 'TMS Treatment Unilateral', 'Expected');
    return { instamed: I, arietis: A, appointments: P };
  }

  // v0.31: every audit run, appended (hidden Account Audit History); Account Audits keeps the latest per patient.
  const AUDIT_HISTORY_COLS = [...AUDIT_COLS, 'Run At', 'Run By'];
  /**
   * Saved Account Audit History rows + one run's summary (AUDIT_COLS): appended, never replaced, oldest first by Run At.
   * runAt: an ISO time or a readable stamp; runBy: who ran it. The same patient and Saved stamp twice is one run.
   */
  function appendAuditHistory(prev, summary, runAt, runBy) {
    const norm = (r) => Object.fromEntries(AUDIT_HISTORY_COLS.map((c) => [c, r[c] == null ? '' : r[c]]));
    const rows = (prev || []).map(norm);
    const next = norm({ ...(summary || {}), 'Run At': str(runAt) || str(summary && summary.Saved), 'Run By': str(runBy) || str(summary && summary.By) });
    const same = (r) => str(r.Key) === str(next.Key) && str(r.Saved) === str(next.Saved);
    if (str(next.Key) && !rows.some(same)) rows.push(next);
    const t = (r) => { const x = Date.parse(str(r['Run At']).replace(/ UTC$/, 'Z')); return isNaN(x) ? 0 : x; };
    return rows.map((r, i) => ({ r, i })).sort((a, b) => t(a.r) - t(b.r) || a.i - b.i).map((x) => x.r);
  }

  return { AUDIT_HISTORY_COLS, appendAuditHistory, AUDIT_COLS, AUDIT_LINE_COLS, AUDIT_SHEET_COLS, auditSheetRows, auditsFromSheet, auditRecord, mergeAudits, savedFor, exportSpec, exportName,
    detectKind, parseInstamed, demoRows, parseArietis, parseAppointments, identify, chargeCode, expectedFor, codeChecks, familyOf, classOf,
    matchPayments, subsetSum, learnBenefits, planYearOf, requote, runAudit, summaryText, dayOf, iso, us, usd, KEPT_RE, chargeRow, PR_PARTS };
});
