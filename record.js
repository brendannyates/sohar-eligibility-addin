/* record.js — one record per patient (v0.20). Pure: no Office.js, no DOM, no network. Loaded after core.js by
   index.html and dashboard.html, and by tests/.
   1. Sohar Flow results (the Flow's results CSV, "Flow Results*.csv"): read, keep the latest per visit, list a patient's.
   2. patientRecord: every Sohar result for one MRN as one record. The most recent result, the upcoming visits, the
      financial progress (deductible and OOP max), then the earlier results and Payer Outcomes ledger rows as a
      timeline, newest first.
   3. byPatient: one row per MRN for the full-screen lists (the most recent result stands for the patient). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.SoharRecord = factory(root.SoharCore);
})(typeof self !== 'undefined' ? self : this, function (C) {
  'use strict';

  const str = (v) => (v == null ? '' : String(v).trim());
  const num = (v) => { const n = parseFloat(str(v).replace(/[$,%\s]/g, '')); return isNaN(n) ? null : n; };

  // ---------- Sohar Flow results ----------
  // The Flow receiver's CSV (eligibility-check skill, scripts/flow/receiver.js), in its order. LASFIR and MRN only: no
  // names or DOBs, so it can sit on a sheet. File and Uploaded are added here.
  const FLOW_COLS = ['Key', 'LASFIR', 'MRN', 'Appt Date', 'Time', 'Event', 'Location', 'Provider', 'Flow Run ID', 'Run Status', 'Run Error',
    'Discovery Status', 'Discovered Payer', 'Policy vs NextGen', 'NextGen Payer ID', 'NextGen Policy Status', 'Discovered Policy Status',
    'Estimate Policy', 'Network', 'Plan', 'Sohar Payer', 'Deductible', 'Deductible Remaining', 'OOP Max', 'OOP Remaining', 'Copay',
    'Coinsurance', 'Fee Tab', 'Rate Level', 'Est Codes', 'Sohar Est Patient', 'Sohar Est Plan Pays', 'Estimate Note', 'Discovery Link',
    'Discovered Verification Link', 'NextGen Verification Link', 'Started', 'Last Updated', 'Note',
    // v0.31: where the check came from (Flow), the picked policy (payerId#last 4), its Reset Basis, and the plan-year lines' PR.
    'Source', 'Policy Key', 'Reset Basis', 'Est 90792', 'Est 99214', 'Est 99215'];
  const FLOW_SHEET_COLS = FLOW_COLS.concat('File', 'Uploaded');
  const FLOW_MARKS = ['Flow Run ID', 'Sohar Est Patient'];

  // The header: the first row (of the first 10) with the Flow columns, so a title row above it is skipped.
  const headerOf = (rows) => {
    const cells = (r) => (r || []).map((v) => str(v).replace(/^﻿/, ''));
    const i = (rows || []).slice(0, 10).findIndex((r) => FLOW_MARKS.every((m) => cells(r).includes(m)));
    return i < 0 ? null : { i, h: cells(rows[i]) };
  };
  /** True when parsed CSV rows (arrays) are a Flow results file. */
  function isFlowFile(rows) {
    const hd = headerOf(rows);
    return !!hd && FLOW_MARKS.every((m) => hd.h.includes(m));
  }
  /** Flow results rows (arrays, header first) -> objects with FLOW_COLS. Rows with no MRN and no Key are left out. */
  function flowRecords(rows) {
    const hd = headerOf(rows);
    if (!hd) return [];
    const out = [];
    for (const r of rows.slice(hd.i + 1)) {
      if (!(r || []).some((v) => str(v))) continue;
      const o = {};
      for (const c of FLOW_COLS) { const j = hd.h.indexOf(c); o[c] = j < 0 ? '' : str(r[j]); }
      if (!o.Source) o.Source = 'Flow'; // files from before v0.31 have no Source column
      if (!o.Key) o.Key = o['Flow Run ID'] || (o.MRN && o['Appt Date'] ? `${o.MRN}|${o['Appt Date']}|${o.Time}` : '');
      if (o.Key) out.push(o);
    }
    return out;
  }
  const updatedMs = (r) => { const t = Date.parse(str(r['Last Updated']) || str(r.Started)); return isNaN(t) ? 0 : t; };
  /**
   * Add uploaded Flow files ({name, rows}) to the saved Flow results (FLOW_SHEET_COLS rows). One row per Key (one visit's
   * Run); the one updated last wins, so an older file uploaded again never replaces a newer result.
   * Returns {rows, added, replaced, files}.
   */
  function mergeFlow(saved, files, uploadedAt) {
    const by = new Map((saved || []).filter((r) => str(r.Key)).map((r) => [str(r.Key), r]));
    let added = 0, replaced = 0, read = 0;
    for (const f of files || []) {
      for (const rec of flowRecords(f.rows)) {
        read++;
        const prev = by.get(rec.Key);
        if (prev && updatedMs(prev) > updatedMs(rec)) continue;
        if (prev) replaced++; else added++;
        by.set(rec.Key, { ...rec, File: str(f.name), Uploaded: str(uploadedAt) });
      }
    }
    return { rows: [...by.values()], added, replaced, read };
  }
  const isSandbox = (r) => /sandbox/i.test(str(r && r.File));
  const visitMs = (r) => {
    const d = C.parseDate(r['Appt Date']);
    if (!d) return NaN;
    const m = str(r.Time).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    const mins = m ? ((+m[1] % 12) + (m[3] && /pm/i.test(m[3]) ? 12 : 0)) * 60 + +m[2] : 0;
    return d.getTime() + mins * 60000;
  };
  /**
   * One patient's Flow results, soonest visit first. Visits before asOf (a day) are left out while the patient has an
   * upcoming one; with none upcoming, the latest past visit is kept so the last estimate still shows.
   */
  function flowFor(flowRows, key, asOf) {
    const k = str(key);
    if (!k) return [];
    const mine = (flowRows || []).filter((r) => str(r.MRN) === k || (k.startsWith('PER-') ? false : str(r.MRN).replace(/^0+/, '') === k.replace(/^0+/, '')));
    const day = asOf ? C.parseDate(asOf instanceof Date ? asOf : String(asOf)) : null;
    const t0 = day ? day.getTime() : -Infinity;
    const up = mine.filter((r) => !(visitMs(r) < t0)).sort((a, b) => (visitMs(a) || 0) - (visitMs(b) || 0));
    if (up.length) return up;
    return mine.sort((a, b) => (visitMs(b) || 0) - (visitMs(a) || 0)).slice(0, 1);
  }
  /** "$45.00" style amounts from a Flow row, '' when blank. */
  const flowAmount = (v) => { const n = num(v); return n == null ? '' : `$${n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; };
  // v0.31 (11/13): plan-year cost estimates on the Sohar Results rows, from Flow Results (the production Flow, or the 7 AM
  // rotation's "Sohar Flow Results - Plan-Year Estimates … (rotation).csv"). Matched on MRN + visit date for the visit's own code
  // (Est Code / Est Patient); otherwise the patient's latest per-code estimate fills Est 90792 / 99214 / 99215 only. Sandbox rows never.
  const ESTIMATE_RESULT_COLS = ['Est Code', 'Est Patient', 'Est 90792', 'Est 99214', 'Est 99215', 'Estimate Note', 'Estimate Source', 'Estimate Date'];
  const PER_CODE_COLS = ['Est 90792', 'Est 99214', 'Est 99215'];
  function withEstimates(results, flowRows) {
    const mk = (m) => str(m).replace(/^0+/, '');
    const dk = (d) => { const t = C.parseDate(str(d)); return t ? C.fmtDate(t) : str(d); };
    const perCode = (f) => PER_CODE_COLS.some((c) => num(f[c]) != null);
    const usable = (flowRows || []).filter((f) => !isSandbox(f) && (num(f['Sohar Est Patient']) != null || perCode(f))).sort((a, b) => updatedMs(a) - updatedMs(b));
    const byVisit = new Map(), byMrn = new Map();
    for (const f of usable) { byVisit.set(mk(f.MRN) + '|' + dk(f['Appt Date']), f); if (perCode(f)) byMrn.set(mk(f.MRN), f); }
    return (results || []).map((r) => {
      const o = { ...r, ...Object.fromEntries(ESTIMATE_RESULT_COLS.map((c) => [c, ''])) };
      const v = str(r.MRN) && str(r['Next Appt']) ? byVisit.get(mk(r.MRN) + '|' + dk(r['Next Appt'])) : null, f = v || byMrn.get(mk(r.MRN));
      if (!f) return o;
      return Object.assign(o, { 'Est Code': v ? str(f['Est Codes']) : '', 'Est Patient': v ? flowAmount(f['Sohar Est Patient']) : '',
        ...Object.fromEntries(PER_CODE_COLS.map((c) => [c, flowAmount(f[c])])), 'Estimate Note': str(f['Estimate Note']),
        'Estimate Source': estimateSourceOf(f), 'Estimate Date': str(f['Last Updated'] || f.Started).slice(0, 10) });
    });
  }
  // v0.32 (Brendan, 11/13): where an estimate's number came from. The fee schedule holds the contracted rates Sohar prices
  // with; when Sohar could not run the estimate, the rotation (or the add-in) priced it with the PR calculator's math.
  const SOHAR_ESTIMATE = 'Sohar Cost Estimate', PR_CALC_ESTIMATE = 'PR Calculator Estimate';
  const estimateSourceOf = (f) => (/fee math|pr calculator/i.test(str(f && f['Estimate Note'])) ? PR_CALC_ESTIMATE : SOHAR_ESTIMATE);

  // ---------- v0.32: the rotation's checks (tblEligibilityChecks, Power Query) on Sohar Results ----------
  // Brendan (11/13): every verification, discovery and cost estimate the eligibility-check rotation ran joins Sohar Results,
  // Verification History, the patient view and the exports, so a clinic the NextGen report on the Import sheet leaves out
  // (Austin, for one) still shows up. One result per Sohar check (Verification ID, else Discovery ID), on the patient's
  // soonest visit on or after asOf that the check covers. A check the Import already matched (same Sohar ID or link) is
  // left out. Source Rotation: never a NextGen task or a payer outcome (the rotation writes its own NextGen tasks).
  const LEDGER_RESULT_COLS = ['Source', 'Discovery Link', 'Cost Estimate Link'];
  // The rotation's Eligibility Checks.csv columns the add-in reads beyond rules.LEDGER_COLS (rules.js is shared with CTT):
  // qLedger.pq brings them in after those (make-pq.mjs), so cost estimate links reach the workbook.
  const LEDGER_LINK_COLS = ['Cost Estimate ID', 'Cost Estimate Link'];
  const LEDGER_SAME = ['MRN', 'LASFIR', 'Sent As', 'Sohar Status', 'Action', 'Benefit Flag', 'Plan', 'Network', 'Coverage Level', 'Coverage Start',
    'Coverage End', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status',
    'Sohar Payer', 'Sohar Date', 'Discovery Link', 'Cost Estimate Link'];
  /** "2026-11-13" or "11/13/2026" -> "11/13/2026" ('' when it isn't a date). */
  const usDay = (v) => { const d = C.parseDate(str(v)); return d ? C.fmtDate(d) : ''; };
  function ledgerResults(ledger, results, opts) {
    const o = opts || {};
    const t0 = o.asOf ? C.parseDate(o.asOf instanceof Date ? o.asOf : str(o.asOf)) : null;
    const low = (v) => str(v).toLowerCase();
    const known = new Set();
    for (const r of results || []) for (const k of ['Sohar Ref', 'Sohar Link']) if (str(r[k])) known.add(low(r[k]));
    const groups = new Map();
    for (const l of ledger || []) {
      const id = str(l['Verification ID']) || str(l['Discovery ID']);
      if (!id || !str(l['Sohar Status']) || !str(l.MRN)) continue;
      if (known.has(low(id)) || (str(l['Sohar Link']) && known.has(low(l['Sohar Link'])))) continue;
      if (!groups.has(id)) groups.set(id, []);
      groups.get(id).push(l);
    }
    const out = [];
    for (const [id, rows] of groups) {
      const dated = rows.map((l) => ({ l, d: C.parseDate(str(l['Appt Date'])) })).sort((a, b) => ((a.d ? a.d.getTime() : Infinity) - (b.d ? b.d.getTime() : Infinity)) || (C.timeMins(a.l.Time) - C.timeMins(b.l.Time)));
      const upcoming = dated.filter((x) => x.d && (!t0 || x.d >= t0));
      const next = (upcoming[0] || dated[dated.length - 1]).l;
      const newest = rows.slice().sort((a, b) => low(b['Last Checked']).localeCompare(low(a['Last Checked'])))[0];
      const r = Object.fromEntries(LEDGER_SAME.map((k) => [k, str(newest[k])]));
      const verified = !!str(newest['Verification ID']);
      Object.assign(r, {
        'Next Appt': usDay(next['Appt Date']), 'Next Time': str(next.Time), 'Next Event': str(next.Event), 'Next Location': str(next.Location), 'Next Provider': str(next.Provider),
        'Upcoming Appts': String(upcoming.length), 'Upcoming Detail': upcoming.map(({ l }) => [usDay(l['Appt Date']), str(l.Time), str(l.Event), str(l.Location) ? `@ ${str(l.Location)}` : ''].filter(Boolean).join(' ')).join('; '),
        'Sent As': str(newest['Sent As']) || (verified ? 'Verification' : 'Discovery'), 'Sohar Ref': id,
        'Sohar Link': str(newest['Sohar Link']) || str(newest['Discovery Link']), 'Discovery Link': verified ? str(newest['Discovery Link']) : '',
        Match: 'Rotation ledger', Source: 'Rotation',
      });
      out.push(r);
    }
    return out;
  }

  // ---------- v0.32: every booked visit priced, soonest first (Brendan, 11/13) ----------
  // Per patient, in visit order, the visit's code (fees.visitCode, passed in as codeOf), then:
  //   1. Sohar's quote for that code (the visit's own Sohar Est Patient, else the plan-year Est <code>) on the first visit
  //      with that code: Sohar priced it against the deductible and OOP max as they stood, so it stands for that visit;
  //   2. each later visit through the PR calculator's math (fees.estimateMath, passed in as math): the code's contracted
  //      rate, the copay and coinsurance, against the deductible and OOP max left after the visits before it.
  // How a code counts: togglesOf(mrn, code) when staff classified it (Estimate Defaults); else a Sohar quote equal to the
  // copay means copay only; else a copay on file means copay only; else deductible, then coinsurance.
  // A visit with a quote but no rate keeps the quote (capped at the OOP max left). The IVK medication is self pay: added to
  // the visit, never toward the deductible or OOP max. Rows carry LASFIR and MRN only.
  const VISIT_PLAN_COLS = ['Appt Date', 'Time', 'LASFIR', 'MRN', 'Event', 'Service Line', 'Location', 'Provider', 'Sohar Status', 'Sohar Payer', 'Plan', 'Network',
    'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'OOP Max', 'OOP Remaining', 'Est Code', 'Est Patient', 'Estimate Source', 'Allowed',
    'Ded Left After', 'OOP Left After', 'Milestone', 'Estimate Note', 'Checked', 'Sohar Link', 'Discovery Link', 'Cost Estimate Link'];
  const COPAY_ONLY = { deductible: 'Waived', coinsurance: 'None', copay: 'Added' };
  const DED_COINS = { deductible: 'Applies', coinsurance: 'After deductible', copay: 'None' };
  function visitPlan(visits, opts) {
    const o = opts || {};
    const codeOf = o.codeOf || (() => ({ code: '', none: true, note: '' }));
    const benefitsOf = o.benefitsOf || (() => null), quotesOf = o.quotesOf || (() => ({})), rateOf = o.rateOf || (() => null);
    const togglesOf = o.togglesOf || (() => null), medOf = o.medOf || (() => null), math = o.math;
    const amt = (n) => (n == null ? '' : money(n));
    const byMrn = new Map();
    for (const v of visits || []) { const k = str(v.MRN); if (!k) continue; if (!byMrn.has(k)) byMrn.set(k, []); byMrn.get(k).push(v); }
    const out = [];
    for (const [mrn, vs] of byMrn) {
      vs.sort((a, b) => (C.parseDate(str(a.date)) - C.parseDate(str(b.date))) || (C.timeMins(a.time) - C.timeMins(b.time)));
      const fin = benefitsOf(mrn) || {};
      const quotes = quotesOf(mrn) || {};
      const dT = num(fin.Deductible), dR = num(fin['Deductible Remaining']), oT = num(fin['OOP Max']), oR = num(fin['OOP Remaining']);
      const copay = num(fin.Copay);
      let dl = dR == null ? null : Math.max(0, dR), ol = oR == null ? null : Math.max(0, oR);
      const used = new Set();
      let dedMarked = dl === 0, oopMarked = ol === 0;
      // The next reset after the check: January 1, or the coverage start's anniversary when Reset Basis says Plan. From
      // then the deductible and OOP max start over and Sohar's quotes (priced against this year's) are re-priced.
      const start = C.parseDate(str(fin['Coverage Start'])), checked = C.parseDate(str(fin['Sohar Date']).slice(0, 10)) || C.parseDate(new Date());
      const nextReset = (after) => {
        const y = after.getUTCFullYear();
        if (str(fin['Reset Basis']) === 'Plan' && start) { const a = Date.UTC(y, start.getUTCMonth(), start.getUTCDate()); return new Date(a > after.getTime() ? a : Date.UTC(y + 1, start.getUTCMonth(), start.getUTCDate())); }
        return new Date(Date.UTC(y + 1, 0, 1));
      };
      let reset = nextReset(checked), stale = false;
      for (const v of vs) {
        const vd = C.parseDate(str(v.date));
        let newYear = false;
        if (vd && vd >= reset) { reset = nextReset(vd); stale = true; newYear = true; dl = dT == null ? null : dT; ol = oT == null ? null : oT; dedMarked = dl === 0; oopMarked = ol === 0; }
        const line = v.line || C.visitLine(v.event);
        const vc = codeOf(v, line, fin['Sohar Payer'] || v.payer) || { code: '', none: true };
        const row = {
          'Appt Date': usDay(v.date) || str(v.date), Time: str(v.time), LASFIR: str(v.lasfir || v.LASFIR), MRN: mrn, Event: str(v.event), 'Service Line': line,
          Location: str(v.location), Provider: str(v.provider), 'Sohar Status': str(fin['Sohar Status']), 'Sohar Payer': str(fin['Sohar Payer']), Plan: str(fin.Plan),
          Network: str(fin.Network), Copay: str(fin.Copay), Coinsurance: str(fin.Coinsurance), Deductible: str(fin.Deductible), 'Deductible Remaining': dl == null ? str(fin['Deductible Remaining']) : amt(dl),
          'OOP Max': str(fin['OOP Max']), 'OOP Remaining': ol == null ? str(fin['OOP Remaining']) : amt(ol), 'Est Code': vc.code || '', 'Est Patient': '', 'Estimate Source': '', Allowed: '',
          'Ded Left After': '', 'OOP Left After': '', Milestone: '', 'Estimate Note': '', Checked: str(fin['Sohar Date']).slice(0, 10),
          'Sohar Link': str(fin['Sohar Link']), 'Discovery Link': str(fin['Discovery Link']), 'Cost Estimate Link': str(fin['Cost Estimate Link']),
        };
        out.push(row);
        if (vc.none || !vc.code) { row['Estimate Note'] = str(vc.note) || 'No plan-year code for this visit'; continue; }
        if (!/^complete\.eligible/i.test(str(fin['Sohar Status']))) { row['Estimate Note'] = fin['Sohar Status'] ? `Not eligible (${str(fin['Sohar Status'])}): no estimate` : 'No Sohar check for this patient yet'; continue; }
        const q0 = quotes[vc.code] || null, q = stale ? null : q0, rate = num(rateOf(mrn, vc.code, v));
        const isCopayQuote = q0 && copay != null && copay > 0 && Math.abs(num(q0.amount) - copay) < 0.005;
        const tg = togglesOf(mrn, vc.code) || (isCopayQuote || (!q0 && copay != null && copay > 0) ? COPAY_ONLY : DED_COINS);
        let patient = null, toward = 0;
        const notes = newYear ? ['New plan year: deductible and OOP max start over'] : [];
        if (ol === 0) { patient = 0; row['Estimate Source'] = q ? q.source : PR_CALC_ESTIMATE; notes.push('OOP max met: $0'); }
        else if (q && !used.has(vc.code)) {
          used.add(vc.code);
          patient = ol == null ? num(q.amount) : Math.min(num(q.amount), ol);
          toward = tg.deductible === 'Applies' && dl != null ? Math.min(patient, dl) : 0;
          row['Estimate Source'] = q.source || SOHAR_ESTIMATE;
          notes.push(`${q.source || SOHAR_ESTIMATE} for ${vc.code}${q.date ? ` (${str(q.date).slice(0, 10)})` : ''}`);
        } else if (rate != null && math) {
          const e = math([{ code: vc.code, units: 1, allowed: rate }], { deductibleRemaining: dl, coinsurance: fin.Coinsurance, copay: fin.Copay, oopRemaining: ol }, tg);
          if (e.patient != null) {
            patient = e.patient; toward = dl == null || e.deductibleAfter == null ? 0 : c2(dl - e.deductibleAfter);
            row['Estimate Source'] = PR_CALC_ESTIMATE; row.Allowed = amt(rate);
            notes.push(`${q ? 'Re-priced after the earlier visits' : 'No Sohar quote'}: ${vc.code} at ${money(rate)}, ${tg === COPAY_ONLY || tg.deductible !== 'Applies' ? 'copay' : 'deductible then coinsurance'}`);
          } else notes.push(`PR calculator: ${e.missing.join(', ') || 'no benefits from Sohar'}`);
        } else if (q) {
          patient = ol == null ? num(q.amount) : Math.min(num(q.amount), ol);
          toward = tg.deductible === 'Applies' && dl != null ? Math.min(patient, dl) : 0;
          row['Estimate Source'] = q.source || SOHAR_ESTIMATE;
          notes.push(`${q.source || SOHAR_ESTIMATE} for ${vc.code}, not re-priced (no contracted rate for ${vc.code})`);
        } else notes.push(`No Sohar quote and no contracted rate for ${vc.code}`);
        if (patient != null) {
          patient = c2(Math.max(0, patient));
          if (dl != null) dl = c2(Math.max(0, dl - toward));
          if (ol != null) ol = c2(Math.max(0, ol - patient));
          const med = vc.medication ? num(medOf(v, mrn)) : null;
          row['Est Patient'] = amt(c2(patient + (med || 0)));
          if (vc.medication) notes.push(med != null ? `includes ${money(med)} IV ketamine medication (self pay)` : 'plus the IV ketamine medication (self pay), not in the total');
          row['Ded Left After'] = dl == null ? '' : amt(dl); row['OOP Left After'] = ol == null ? '' : amt(ol);
          const marks = [];
          if (!dedMarked && dl === 0 && dT) { dedMarked = true; marks.push('Deductible met at this visit'); }
          if (!oopMarked && ol === 0 && oT) { oopMarked = true; marks.push('OOP max met at this visit'); }
          row.Milestone = marks.join('; ');
        }
        row['Estimate Note'] = notes.join('; ');
      }
    }
    return out.sort((a, b) => (C.parseDate(a['Appt Date']) - C.parseDate(b['Appt Date'])) || (C.timeMins(a.Time) - C.timeMins(b.Time)) || a.LASFIR.localeCompare(b.LASFIR));
  }
  /**
   * MRN -> {code: {amount, source, date}}: each patient's newest non-sandbox quote per plan-year code, from Flow Results
   * rows (Est 90792 / 99214 / 99215, and the visit's own Est Codes + Sohar Est Patient) and the rotation ledger
   * (Est Codes + Sohar Est Patient or Projected Patient Responsibility).
   */
  function quotesByMrn(flowRows, ledger) {
    const out = new Map(), mk = (m) => str(m).replace(/^0+/, '');
    const put = (mrn, code, amount, source, date) => {
      const n = num(amount); if (!mrn || !code || n == null) return;
      if (!out.has(mrn)) out.set(mrn, {});
      const q = out.get(mrn), prev = q[code];
      if (!prev || str(date) >= str(prev.date)) q[code] = { amount: n, source, date: str(date) };
    };
    for (const f of flowRows || []) {
      if (isSandbox(f)) continue;
      const m = mk(f.MRN), src = estimateSourceOf(f), d = str(f['Last Updated'] || f.Started);
      for (const c of PER_CODE_COLS) put(m, c.replace(/^Est /, ''), f[c], src, d);
      const own = str(f['Est Codes']);
      if (/^\d{5}$/.test(own)) put(m, own, f['Sohar Est Patient'], src, d);
    }
    for (const l of ledger || []) {
      const own = str(l['Est Codes']);
      if (!/^\d{5}$/.test(own)) continue;
      put(mk(l.MRN), own, str(l['Sohar Est Patient']) || str(l['Projected Patient Responsibility']), estimateSourceOf(l), str(l['Last Checked']));
    }
    return out;
  }
  /** Sohar's own dashboard export of Flow runs: run IDs and inputs only, no outputs (the receiver fetches those). */
  function isSoharRunExport(rows) {
    if (isFlowFile(rows)) return false;
    return (rows || []).slice(0, 10).some((r) => (r || []).some((c) => /^(flowrunid|runid)$/i.test(str(c).replace(/[\s_-]/g, ''))));
  }

  // ---------- one record per patient ----------
  const timeOf = (r) => { const t = C.soharTime(r['Sohar Date']); return isNaN(t) ? -Infinity : t; };
  const newestFirst = (a, b) => timeOf(b) - timeOf(a);
  // Upcoming visits are per patient (every result row carries the same ones); "Upcoming Detail" lists them as
  // "10/30/2026 9:00 AM Event @ Location; ...".
  function upcomingOf(r) {
    if (!r) return [];
    const list = str(r['Upcoming Detail']).split(/\s*;\s*/).filter(Boolean);
    if (list.length) return list;
    const next = [r['Next Appt'], r['Next Time'], r['Next Event'], r['Next Location'] ? `@ ${r['Next Location']}` : ''].map(str).filter(Boolean).join(' ');
    return next ? [next] : [];
  }
  /**
   * The result to lead with: the most recent Sohar Date. A superseded result (an older check of a policy since
   * replaced) never leads while there is another.
   */
  function leadOf(results) {
    const live = results.filter((r) => !/^Superseded/.test(str(r.Action)));
    return (live.length ? live : results).slice().sort(newestFirst)[0] || null;
  }
  /** A timeline entry from a Sohar result row. */
  const resultEntry = (r) => ({
    source: 'result', time: timeOf(r), date: str(r['Sohar Date']), kind: str(r['Sent As']) || 'Verification', status: str(r['Sohar Status']),
    payer: str(r['Sohar Payer']) || str(r['NextGen Payer']), plan: str(r.Plan), network: str(r.Network), action: str(r.Action),
    coverage: [str(r['Coverage Start']), str(r['Coverage End'])].some(Boolean) ? `${str(r['Coverage Start']) || '?'} to ${str(r['Coverage End']) || 'open'}` : '',
    benefits: [r.Deductible !== '' && r.Deductible != null ? `Deductible ${str(r['Deductible Remaining']) || '?'} left of ${str(r.Deductible)}` : '',
      r['OOP Max'] !== '' && r['OOP Max'] != null ? `OOP ${str(r['OOP Remaining']) || '?'} left of ${str(r['OOP Max'])}` : '',
      str(r.Copay) ? `Copay ${str(r.Copay)}` : '', str(r.Coinsurance) ? `Coins ${str(r.Coinsurance)}` : ''].filter(Boolean).join(' · '),
    searched: str(r['Searched As']), link: str(r['Sohar Link']), row: r,
  });
  /** A timeline entry from a Payer Outcomes ledger row (older checks the saved results no longer hold). */
  const outcomeEntry = (o) => {
    const t = C.soharTime(o['Sohar Date']);
    return { source: 'ledger', time: isNaN(t) ? -Infinity : t, date: str(o['Sohar Date']), kind: str(o.Kind) || 'Verification', status: str(o['Sohar Status']),
      payer: str(o['Returned Payer']) || str(o['NextGen Payer']), plan: '', network: '', action: '', coverage: '', benefits: '', searched: '', link: '', signal: str(o.Signal) };
  };
  const dayKey = (e) => (isFinite(e.time) ? new Date(e.time).toISOString().slice(0, 10) : e.date);
  /**
   * Everything for one patient (key = core.keyOf), newest first.
   * rows: Sohar Results rows. outcomes: Payer Outcomes ledger rows (optional). flow: saved Flow results (optional).
   * Returns null when the patient has no result. {key, lead, results, upcoming, financial, flow, timeline}:
   * - financial: the Benefits view row (core.benefitsView) for the patient, or null (no verified benefits).
   * - timeline: every result but the lead, plus ledger rows with no result of the same kind, status and day.
   */
  function patientRecord(rows, key, opts) {
    const o = opts || {};
    const k = str(key);
    const results = (rows || []).filter((r) => str(C.keyOf(r)) === k).sort(newestFirst);
    if (!k || !results.length) return null;
    const lead = leadOf(results);
    const entries = results.filter((r) => r !== lead).map(resultEntry);
    const seen = new Set(results.map((r) => { const e = resultEntry(r); return `${e.kind}|${e.status}|${dayKey(e)}`; }));
    for (const x of o.outcomes || []) {
      if (str(x.MRN) !== k) continue;
      const e = outcomeEntry(x);
      const id = `${e.kind}|${e.status}|${dayKey(e)}`;
      if (seen.has(id)) continue;
      seen.add(id);
      entries.push(e);
    }
    entries.sort((a, b) => b.time - a.time);
    const ben = C.benefitsView(results).rows;
    const financial = ben.find((b) => str(C.keyOf(b)) === k) || ben[0] || null;
    return { key: k, lead, results, upcoming: upcomingOf(lead), financial, flow: flowFor(o.flow, k, o.asOf), timeline: entries };
  }

  /**
   * One row per patient for the lists: the lead result (row), every result for the MRN (rows, newest first) and
   * queues (each Work Queue those results are in). Lists show a patient when any of their results passes the filters,
   * so a filter on Action needed still finds a patient whose lead result is Clear. Rows with no key (unmatched Sohar
   * rows) stay one per result.
   */
  function byPatient(rows) {
    const groups = new Map(), out = [];
    for (const r of rows || []) {
      const k = str(C.keyOf(r));
      if (!k) { out.push({ row: r, rows: [r], resultCount: 1, queues: [str(r['Work Queue'])] }); continue; }
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    for (const rs of groups.values()) {
      rs.sort(newestFirst);
      out.push({ row: leadOf(rs), rows: rs, resultCount: rs.length, queues: [...new Set(rs.map((r) => str(r['Work Queue'])))] });
    }
    return out;
  }

  // ---------- report profile: gender, pronouns, guarantor (v0.20) ----------
  // Read by header name from the Import rows (never by column position), like core.detectHeader. These are shown in the
  // full-screen patient record only: never written to a sheet or a download (as with names and DOBs).
  const PROFILE_HEADERS = {
    genderIdentity: ['Gender Identity'],
    currentGender: ['Current Gender', 'Gender', 'Sex', 'Birth Sex'],
    pronouns: ['Preferred Pronoun', 'Preferred Pronouns', 'Pronouns', 'Pronoun'],
    prevLastName: ['Prev Lst Name', 'Previous Last Name', 'Prev Last Name', 'Prior Last Name'],
  };
  const norm = (s) => C.normHeader(s);
  const letters = (s) => str(s).toLowerCase().replace(/[^a-z]/g, '');
  /**
   * Per Nbr -> {genderIdentity, currentGender, pronouns, prevLastName, patient: {firstName, lastName}, guarantor:
   * {firstName, lastName, dob} | null, guarantorSelf}. rows: the Import rows (arrays); det: core.detectHeader(rows).
   * The latest non-blank value per patient wins. Only Per Nbrs in `pers` (a Set) when given.
   */
  function profiles(rows, det, pers) {
    const out = new Map();
    if (!det || !rows) return out;
    const hdr = (rows[det.headerRow] || []).map(norm);
    const col = (names) => { for (const n of names) { const i = hdr.indexOf(norm(n)); if (i >= 0) return i; } return -1; };
    const cols = Object.fromEntries(Object.entries(PROFILE_HEADERS).map(([k, v]) => [k, col(v)]));
    const cm = det.colMap || {};
    for (let r = det.headerRow + 1; r < (det.dataEnd || rows.length); r++) {
      const row = rows[r] || [];
      const per = str(row[cm.perNbr]);
      if (!per || (pers && !pers.has(per))) continue;
      const p = out.get(per) || { genderIdentity: '', currentGender: '', pronouns: '', prevLastName: '', patient: null, guarantor: null, guarantorSelf: false, guarRaw: '' };
      for (const [k, c] of Object.entries(cols)) if (c >= 0 && str(row[c])) p[k] = str(row[c]);
      if (cm.patName != null && str(row[cm.patName])) {
        const n = C.parseName(row[cm.patName]);
        p.patient = { firstName: n.firstName, lastName: n.lastName, dob: cm.dob != null ? C.fmtDate(C.parseDate(row[cm.dob])) : '' };
      }
      if (cm.guarName != null && str(row[cm.guarName])) {
        p.guarRaw = str(row[cm.guarName]);
        const g = C.parseName(p.guarRaw);
        p.guarantor = { firstName: g.firstName, lastName: g.lastName, dob: cm.guarDob != null ? C.fmtDate(C.parseDate(row[cm.guarDob])) : '' };
      }
      out.set(per, p);
    }
    for (const p of out.values()) p.guarantorSelf = !!(p.guarantor && p.patient && samePerson(p.guarantor, p.patient, p.prevLastName));
    return out;
  }
  /**
   * Same person by name: last names equal (or one is the patient's previous last name) and first names agree on their
   * first three letters (Kate / Katherine); a DOB on both sides must match too.
   */
  function samePerson(a, b, prevLast) {
    if (!a || !b) return false;
    const al = letters(a.lastName), bl = letters(b.lastName), pl = letters(prevLast);
    if (!al || !bl) return false;
    if (al !== bl && !(pl && (al === pl || bl === pl))) return false;
    const af = letters(a.firstName).slice(0, 3), bf = letters(b.firstName).slice(0, 3);
    if (af && bf && af !== bf) return false;
    const ad = str(a.dob), bd = str(b.dob);
    if (ad && bd) { const x = C.parseDate(ad), y = C.parseDate(bd); if (x && y && x.getTime() !== y.getTime()) return false; }
    return true;
  }
  /**
   * NextGen guarantor vs the subscriber Sohar returned. prof: profiles() entry; sub: {firstName, lastName, dob} from the
   * Sohar record (blank when Sohar sent none). Returns {status, severity, text}, LASFIR only:
   * - Match: the guarantor is the subscriber.
   * - Subscriber is the patient, guarantor is not: normal when a parent or partner pays; confirm the NextGen policy's
   *   insured is the patient.
   * - Guarantor is the patient, subscriber is not: the patient is a dependent on someone else's plan; NextGen's
   *   insured must be that subscriber, not the patient.
   * - Different people: neither is the other; check both in NextGen.
   * - No subscriber from Sohar / No guarantor on the report: nothing to compare.
   */
  function guarantorCheck(prof, sub) {
    const tag = (n) => (n && (n.lastName || n.firstName) ? C.lasfir(n.lastName, n.firstName) : '');
    const s = sub && letters(sub.lastName) ? sub : null;
    const g = prof && prof.guarantor && letters(prof.guarantor.lastName) ? prof.guarantor : null;
    const pt = prof && prof.patient;
    // Always LASFIR: the patient is named too ("RIVANA, the patient").
    const me = tag(pt) ? `${tag(pt)}, the patient` : 'the patient';
    if (!s) return { status: 'No subscriber from Sohar', severity: '', text: g ? `NextGen guarantor is ${prof.guarantorSelf ? me : tag(g)}; Sohar returned no subscriber to compare.` : '' };
    if (!g) return { status: 'No guarantor on the report', severity: '', text: `Sohar subscriber is ${samePerson(s, pt, prof && prof.prevLastName) ? me : tag(s)}; the report has no guarantor to compare.` };
    const subIsPt = samePerson(s, pt, prof.prevLastName), gIsPt = prof.guarantorSelf;
    if (samePerson(g, s, prof.prevLastName)) return { status: 'Match', severity: '', text: `NextGen guarantor and Sohar subscriber match (${gIsPt ? me : tag(g)}).` };
    if (subIsPt) return { status: 'Subscriber is the patient, guarantor is not', severity: 'low', text: `Sohar subscriber is ${me}; NextGen guarantor is ${tag(g)}. Normal when someone else pays; confirm the policy's insured is ${tag(pt) || 'the patient'} in NextGen.` };
    // Common and often fine (an adult on a partner's plan pays their own bill), but the NextGen policy's insured must be
    // the subscriber, and the report has no insured column: a check, not an error.
    if (gIsPt) return { status: 'Guarantor is the patient, subscriber is not', severity: 'medium', text: `Sohar subscriber is ${tag(s)}, not ${me}; NextGen's guarantor is ${tag(pt) || 'the patient'}. Check that the NextGen policy's insured is ${tag(s)} (the guarantor can stay ${tag(pt) || 'the patient'}).` };
    return { status: 'Different people', severity: 'high', text: `NextGen guarantor ${tag(g)} and Sohar subscriber ${tag(s)} are different people (patient ${tag(pt) || 'unknown'}). Check the guarantor and the policy's insured in NextGen.` };
  }
  /** Subscriber from a saved Sohar record (lower-case Sohar headers), or null. */
  function subscriberOf(rec) {
    if (!rec) return null;
    const firstName = str(rec.subscriberfirstname), lastName = str(rec.subscriberlastname);
    return firstName || lastName ? { firstName, lastName, dob: str(rec.subscriberdateofbirth) } : null;
  }

  // ---------- visit pre-check (v0.20) ----------
  // The report's "Appointment Pre-Check Complete" (Yes / No / blank) per visit, plus a checklist and notes staff add in
  // full screen (notes are per patient). Saved on the hidden Pre-Check sheet (PRECHECK_COLS: MRN key, the patient's LASFIR, the admin's name);
  // a tick wins over the report's value. The add-in never writes to NextGen: "Updated in NextGen" is the person saying
  // they did it. Admins pick their name from Config "Admin names" (Excel does not give the add-in the signed-in user).
  const PRECHECK_HEADERS = ['Appointment Pre-Check Complete', 'Pre-Check Complete', 'Precheck Complete', 'Pre-Check', 'Precheck'];
  const PRECHECK_ITEMS = [
    { id: 'precheck', label: 'Pre-check complete' },
    { id: 'nextgen', label: 'Updated in NextGen' },
  ];
  const PRECHECK_COLS = ['Key', 'MRN', 'LASFIR', 'Appt Date', 'Time', 'Item', 'Done', 'Note', 'By', 'When'];
  const PRECHECK_CONFIG_ROWS = [
    { setting: 'Admin names', value: '', what: 'Pre-check (full screen): the names admins pick from when they tick a visit or add a note, separated by ; (e.g. Jordan Lee; Sam Ortiz). Saved with each tick and note.' },
  ];
  const NOTE_MAX = 500;
  const yes = (v) => /^(y|yes|true|1|x|complete|completed|done)$/i.test(str(v));
  const no = (v) => /^(n|no|false|0)$/i.test(str(v));
  /** One visit's key: patient key (MRN) | appointment date | time. */
  const visitKey = (key, date, time) => `${str(key)}|${str(date)}|${str(time)}`;
  /** Config "Admin names" -> a clean, de-duplicated list. */
  const adminNames = (v) => [...new Set(str(v).split(/[;\n]/).map((s) => s.trim()).filter(Boolean))];
  /**
   * Per Nbr -> visits on the report from `from` (a day; all when null), soonest first:
   * [{date, time, event, loc, provider, status, precheck: 'Yes' | 'No' | '', hasPrecheck}]. Read by header name; hasPrecheck is
   * false when the report has no Pre-Check column.
   */
  function reportVisits(rows, det, pers, from) {
    const out = new Map();
    if (!det || !rows) return out;
    const hdr = (rows[det.headerRow] || []).map(norm);
    const pc = PRECHECK_HEADERS.map(norm).map((h) => hdr.indexOf(h)).find((i) => i >= 0);
    const cm = det.colMap || {};
    const t0 = from ? C.parseDate(from instanceof Date ? from : String(from)) : null;
    for (let r = det.headerRow + 1; r < (det.dataEnd || rows.length); r++) {
      const row = rows[r] || [];
      const per = str(row[cm.perNbr]);
      if (!per || (pers && !pers.has(per))) continue;
      const d = C.parseDate(row[cm.apptDt]);
      if (!d || (t0 && d < t0)) continue;
      const raw = pc == null ? '' : row[pc];
      const v = { date: C.fmtDate(d), time: cm.begTm != null ? C.fmtTime(row[cm.begTm]) : '', event: str(row[cm.event]), loc: str(row[cm.loc]),
        provider: str(row[cm.rendering]), status: str(row[cm.sts]), precheck: yes(raw) ? 'Yes' : no(raw) ? 'No' : '', hasPrecheck: pc != null };
      if (!out.has(per)) out.set(per, []);
      out.get(per).push(v);
    }
    for (const list of out.values()) list.sort((a, b) => (C.parseDate(a.date) - C.parseDate(b.date)) || (C.timeMins(a.time) - C.timeMins(b.time)));
    return out;
  }
  /**
   * The checklist for one visit: [{id, label, done, source, when, by}]. A saved tick wins; with none, "Pre-check
   * complete" follows the report (source 'NextGen report'). saved: Pre-Check sheet rows.
   */
  function precheckItems(visit, key, saved) {
    const vk = visitKey(key, visit.date, visit.time);
    const mine = new Map((saved || []).filter((r) => str(r.Key) === vk && str(r.Item) !== 'note').map((r) => [str(r.Item), r]));
    return PRECHECK_ITEMS.map((it) => {
      const s = mine.get(it.id);
      if (s) return { ...it, done: yes(s.Done), source: 'Marked here', when: str(s.When), by: str(s.By) };
      if (it.id === 'precheck' && visit.precheck) return { ...it, done: visit.precheck === 'Yes', source: 'NextGen report', when: '', by: '' };
      return { ...it, done: false, source: '', when: '', by: '' };
    });
  }
  /**
   * The patient's notes (one notes list per patient, not per visit), newest first: [{note, by, when, visit}]. visit is
   * the visit date and time a note was written against before v0.21.1 (per-visit notes), else ''.
   */
  function patientNotes(key, saved) {
    const k = str(key);
    return (saved || []).map((r, i) => ({ r, i })).filter(({ r }) => str(r.MRN) === k && str(r.Item) === 'note' && str(r.Note))
      .sort((a, b) => b.i - a.i).map(({ r }) => ({ note: str(r.Note), by: str(r.By), when: str(r.When), visit: [str(r['Appt Date']), str(r.Time)].filter(Boolean).join(' ') }));
  }
  const rowOf = (t, item, extra) => ({ Key: visitKey(t.key, t.date, t.time), MRN: str(t.key), LASFIR: str(t.lasfir), 'Appt Date': str(t.date), Time: str(t.time),
    Item: item, Done: '', Note: '', By: str(t.by), When: str(t.when), ...extra });
  const needBy = (t) => { if (!str(t.by)) throw new Error('Pick your name first (Config "Admin names" lists them).'); };
  /** Save one tick: replaces the visit's row for that item. t: {key, lasfir, date, time, item, done, by, when}. */
  function setPrecheck(saved, t) {
    if (!PRECHECK_ITEMS.some((i) => i.id === t.item)) throw new Error(`Unknown pre-check item "${t.item}".`);
    needBy(t);
    const vk = visitKey(t.key, t.date, t.time);
    const rows = (saved || []).filter((r) => !(str(r.Key) === vk && str(r.Item) === t.item));
    rows.push(rowOf(t, t.item, { Done: t.done ? 'Yes' : 'No' }));
    return rows;
  }
  /**
   * Add a note to the patient (kept with every earlier note). t: {key, lasfir, note, by, when}; no visit. Trimmed, at
   * most NOTE_MAX characters.
   */
  function addPatientNote(saved, t) {
    needBy(t);
    const note = str(t.note).replace(/\s+/g, ' ').slice(0, NOTE_MAX);
    if (!note) throw new Error('The note is empty.');
    return (saved || []).concat({ ...rowOf({ ...t, date: '', time: '' }, 'note', { Note: note }), Key: str(t.key) });
  }
  /** Drop ticks for visits before `from` (a day), so the sheet does not grow forever. Notes are the patient's: always kept. */
  function prunePrecheck(saved, from) {
    const t0 = from ? C.parseDate(from instanceof Date ? from : String(from)) : null;
    if (!t0) return saved || [];
    return (saved || []).filter((r) => { if (str(r.Item) === 'note') return true; const d = C.parseDate(r['Appt Date']); return !d || d >= t0; });
  }

  // ---------- Sohar Flow import (v0.21; v0.31 columns per docs/ARCHITECTURE.md section 5) ----------
  // The CSV template of the production "MHS Eligibility - Discovery vs NextGen" Flow (template v27), exactly, in order.
  // One row per patient. No tin: the Flow template looks the TIN up by providerGroupName. estimate is "yes" only when
  // the policy's plan-year estimate is due (fees.planYearEstimateDue); rateTable is a JSON string the Flow parses
  // ($eval): {"<Sohar payerId>": [{"code": "99214", "contractedRate": 0.00}, ...]}.
  const FLOW_IMPORT_COLS = ['id', 'patient.firstName', 'patient.lastName', 'patient.dateOfBirth', 'patient.state', 'placeOfServiceCode',
    'specialtyCode', 'chartPayerId', 'chartMemberId', 'providerGroupName', 'estimate', 'rateTable'];
  // Sohar's provider group names, by the patient's state (exactly as Sohar has them).
  const PROVIDER_GROUPS = { CA: 'MINDFUL HEALTH SOLUTIONS A PROFESSIONAL MEDICAL CORPORATION', TX: 'Mindful Health Solutions TX PA', WA: 'Mindful Health Solutions Wa Pc' };
  // v0.31: the TIN and Sohar provider group rows are gone (the Flow has them); the plan-year codes row is fees.PLAN_YEAR_CONFIG_ROWS.
  const FLOW_IMPORT_CONFIG_ROWS = [];
  /**
   * Flow import rows from the built patients (core.buildPatients after the freshness rule). In scope only. chartPayerId
   * and chartMemberId are the NextGen policy the Flow checks discovery against: filled only for patients routed to
   * verification (a confident payer), from their first insurance still due; discovery patients get blanks, so a guessed
   * payer is never verified as the chart policy. Left out: a service line with Include in import Off (x.excluded), every
   * insurance not due under the freshness rule, and rows the Flow would reject (no name, DOB or state).
   * cfg: {specialty, estimateOf(p, v) -> true/'yes' when the plan-year estimate is due, rateTableOf(p, v) -> {payerId:
   * lines} or a JSON string}. Returns {rows, skipped: [{key, why}], estimates (rows with estimate yes)}.
   */
  function flowImportRows(patients, cfg) {
    const c = cfg || {};
    const rows = [], skipped = [];
    let estimates = 0;
    for (const p of patients || []) {
      if (p['In Scope'] !== 'Yes') continue;
      const key = C.keyOf(p);
      const vrs = (p.x && p.x.verifyRows) || [];
      if (p.x && p.x.excluded) { skipped.push({ key, why: p.x.excluded.reason || 'service line Off in Config' }); continue; }
      if (p.Route === 'verification' && vrs.length && vrs.every((v) => v.reverifySkip)) { skipped.push({ key, why: 'already covered, or too early to re-verify' }); continue; }
      if (!str(p['First Name']) || !str(p['Last Name']) || !str(p.DOB) || !str(p.State)) { skipped.push({ key, why: 'no name, DOB or state' }); continue; }
      const v = p.Route === 'verification' ? (vrs.find((x) => !x.reverifySkip) || { payerId: p['Sohar payerId'], memberId: p['Member ID'] }) : null;
      const est = typeof c.estimateOf === 'function' ? c.estimateOf(p, v) : false;
      const yes = est === true || /^yes$/i.test(str(est)) || !!(est && est.due);
      let rt = yes && typeof c.rateTableOf === 'function' ? c.rateTableOf(p, v) : null;
      if (rt && typeof rt === 'object') rt = Object.keys(rt).length ? JSON.stringify(rt) : '';
      if (yes) estimates++;
      rows.push({ id: key, 'patient.firstName': str(p['First Name']), 'patient.lastName': str(p['Last Name']), 'patient.dateOfBirth': str(p.DOB),
        'patient.state': str(p.State), placeOfServiceCode: str(p.POS), specialtyCode: str(c.specialty) || 'PSY',
        chartPayerId: v ? str(v.payerId) : '', chartMemberId: v ? str(v.memberId) : '',
        providerGroupName: PROVIDER_GROUPS[str(p.State).toUpperCase()] || '', estimate: yes ? 'yes' : 'no', rateTable: str(rt) });
    }
    return { rows, skipped, estimates };
  }
  // ---------- Verification History (v0.31, docs/ARCHITECTURE.md section 5) ----------
  // One row per Sohar check the add-in processed (Source Import, Flow or Dashboard), never replaced, sorted by Pulled At.
  // The rotation's Eligibility History.csv columns, so a policy's timeline merges both (policyTimeline). No member ID:
  // Policy Key is the Sohar payerId + '#' + the member ID's last 4 characters.
  const VERIFICATION_HISTORY_COLS = ['Pulled At', 'LASFIR', 'MRN', 'Per Nbr', 'Clinic', 'Kind', 'Sohar Status', 'Sohar Payer', 'Sohar payerId', 'Plan',
    'Network', 'Coverage Level', 'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status',
    'OOP Max', 'OOP Remaining', 'OOP Status', 'Benefit Flag', 'Sohar ID', 'Sohar Link', 'Related Sohar IDs', 'Policy Key',
    'Projected Patient Responsibility', 'Estimate Note', 'Source'];
  const HISTORY_SOURCES = ['Import', 'Flow', 'Rotation', 'Dashboard', 'Webhook'];
  /** "94036#7890" from a Sohar payerId and a member ID (letters and digits, last 4, upper case); '' without a payerId. */
  function policyKey(payerId, memberId) {
    const id = str(payerId);
    if (!id) return '';
    return `${id}#${str(memberId).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(-4)}`;
  }
  const isoOf = (v) => { const t = C.soharTime(v); return isNaN(t) ? str(v) : new Date(t).toISOString(); };
  const linkId = (v) => { const m = /\/(?:verifications|discovery)\/([^/?#"]+)/.exec(str(v)); return m ? m[1] : ''; };
  // An id match (MRN, Per Nbr, Sohar ref) is a check the add-in's own import sent; anything else came from the dashboard.
  const importMatch = (m) => /^(mrn|per nbr|id|sohar ref)/i.test(str(m));
  /**
   * Sohar Results rows (mergeSoharResults) -> Verification History rows. Rows with no Sohar ref, not in the Sohar
   * file, or matched by name only / not at all are left out (they may be someone else). source: 'Import' or 'Dashboard';
   * blank = by match (an id match is Import, else Dashboard).
   */
  /**
   * The Sohar ids in a set of uploaded result files that the 7 AM rotation wrote ("… (rotation).csv", lib.js writeImports):
   * their history rows are Source Rotation, not Import, even though they match by MRN. files: [{name, records}].
   */
  function rotationIds(files) {
    const ids = new Set();
    for (const f of files || []) { if (!/\(rotation\)/i.test(str(f && f.name))) continue;
      for (const r of (f && f.records) || []) for (const k of ['verificationId', 'requestId', 'discoveryId']) if (str(r && r[k])) ids.add(str(r[k])); }
    return ids;
  }
  /** History rows with Source set to `source` for the Sohar ids in `ids` (a Set); other rows unchanged. */
  function withSource(rows, ids, source) { return (rows || []).map((r) => (ids && ids.has(str(r['Sohar ID'])) ? { ...r, Source: source } : r)); }
  function historyFromResults(results, source) {
    const out = [];
    for (const r of results || []) {
      const ref = str(r['Sohar Ref']);
      if (!ref || r['Sohar Status'] === 'Not in Sohar file' || /^name|^no match|^no identifiers/i.test(str(r.Match))) continue;
      const row = {};
      for (const k of VERIFICATION_HISTORY_COLS) row[k] = r[k] == null ? '' : str(r[k]);
      Object.assign(row, {
        'Pulled At': isoOf(r['Sohar Date']), Clinic: str(r['Next Location']), Kind: str(r['Sent As']) || 'Verification',
        'Sohar ID': ref, 'Policy Key': policyKey(r['Sohar payerId'], r['Sohar Member ID'] || r['NextGen Member ID']),
        'Related Sohar IDs': '', 'Projected Patient Responsibility': '', 'Estimate Note': '',
        Source: str(source) || (importMatch(r.Match) ? 'Import' : 'Dashboard'),
      });
      out.push(row);
    }
    return out;
  }
  /** Flow Results rows (FLOW_COLS) -> Verification History rows, Source Flow. A Run with no verification is left out. */
  function historyFromFlow(flowRows) {
    const out = [];
    for (const r of flowRows || []) {
      // The rotation's estimates file re-states verifications it did days earlier, stamped with the estimate's time. Those
      // checks reach history from the rotation's own verification import files with their real dates; a copy dated today would
      // make the visit look freshly verified. Its estimates stay on Flow Results (estimateDueFor reads them there).
      if (/^rotation estimate$/i.test(str(r['Run Status']))) continue;
      const disc = /discover/i.test(str(r['Estimate Policy']));
      const id = linkId(disc ? r['Discovered Verification Link'] : r['NextGen Verification Link']) || linkId(r['Discovered Verification Link']) || linkId(r['NextGen Verification Link']);
      if (!id) continue;
      out.push({
        ...Object.fromEntries(VERIFICATION_HISTORY_COLS.map((k) => [k, ''])),
        'Pulled At': isoOf(str(r['Last Updated']) || str(r.Started)), LASFIR: str(r.LASFIR), MRN: str(r.MRN), Clinic: str(r.Location), Kind: 'Verification',
        'Sohar Status': str(disc ? r['Discovered Policy Status'] : r['NextGen Policy Status']) || str(r['NextGen Policy Status']),
        'Sohar Payer': str(r['Sohar Payer']), 'Sohar payerId': str(r['Policy Key']).split('#')[0] || (disc ? '' : str(r['NextGen Payer ID'])), Plan: str(r.Plan), Network: str(r.Network),
        Copay: str(r.Copay), Coinsurance: str(r.Coinsurance), Deductible: str(r.Deductible), 'Deductible Remaining': str(r['Deductible Remaining']),
        'OOP Max': str(r['OOP Max']), 'OOP Remaining': str(r['OOP Remaining']), 'Sohar ID': id,
        'Sohar Link': `https://app.soharhealth.com/verifications/${id}`, 'Related Sohar IDs': linkId(r['Discovery Link']),
        'Policy Key': str(r['Policy Key']), 'Projected Patient Responsibility': str(r['Sohar Est Patient']), 'Estimate Note': str(r['Estimate Note']), Source: str(r.Source) || 'Flow',
      });
    }
    return out;
  }
  const pulledMs = (r) => { const t = C.soharTime(r && r['Pulled At']); return isNaN(t) ? 0 : t; };
  const histKey = (r) => str(r['Sohar ID']) || `${str(r.MRN)}|${str(r['Pulled At'])}|${str(r.Kind)}|${str(r['Policy Key'])}`;
  /**
   * Saved Verification History + new rows: append-only. A row already saved (same Sohar ID) has its fields updated by
   * the new read (blank new fields keep the saved value) and keeps its first Source; nothing is ever dropped. Sorted by
   * Pulled At, oldest first. Returns {rows, added, updated}.
   */
  function mergeHistory(saved, fresh) {
    const by = new Map(), order = [];
    const norm = (r) => Object.fromEntries(VERIFICATION_HISTORY_COLS.map((k) => [k, r[k] == null ? '' : str(r[k])]));
    for (const r of saved || []) { const k = histKey(r); if (!by.has(k)) order.push(k); by.set(k, norm(r)); }
    let added = 0, updated = 0;
    for (const r of fresh || []) {
      const k = histKey(r), n = norm(r), prev = by.get(k);
      if (!prev) { by.set(k, n); order.push(k); added++; continue; }
      const next = { ...prev };
      for (const c of VERIFICATION_HISTORY_COLS) if (c !== 'Source' && n[c] !== '') next[c] = n[c];
      if (VERIFICATION_HISTORY_COLS.some((c) => next[c] !== prev[c])) { by.set(k, next); updated++; }
    }
    const rows = order.map((k, i) => ({ r: by.get(k), i })).sort((a, b) => pulledMs(a.r) - pulledMs(b.r) || a.i - b.i).map((x) => x.r);
    return { rows, added, updated };
  }
  const mrnSame = (a, b) => { const x = str(a).replace(/^0+/, ''), y = str(b).replace(/^0+/, ''); return !!x && x === y; };
  /**
   * One patient's checks as one timeline per policy, newest first: Verification History rows (the add-in) merged with
   * tblPatientHistory rows (the rotation and the webhook: Row Key, MRN, When, Kind, Location, Detail "status · payer",
   * Source, Ref). The same Sohar ID once (Verification History wins). A Patient History row joins the policy whose Sohar
   * Payer it names (only one such policy), else a group of its own by payer name.
   * Returns [{policyKey, payer, payerId, last4, latest, rows}], the policy with the newest check first.
   */
  function policyTimeline(historyRows, patientHistoryRows, mrn) {
    const mine = (historyRows || []).filter((r) => mrnSame(r.MRN, mrn));
    const ids = new Set(mine.map((r) => str(r['Sohar ID'])).filter(Boolean));
    const rot = (patientHistoryRows || []).filter((r) => mrnSame(r.MRN, mrn) && !ids.has(str(r.Ref) || str(r['Row Key']))).map((r) => {
      const [status, payer] = str(r.Detail).split(' · ');
      return { 'Pulled At': isoOf(r.When), MRN: str(r.MRN), LASFIR: str(r.LASFIR), Clinic: str(r.Location), Kind: str(r.Kind), 'Sohar Status': str(status),
        'Sohar Payer': str(payer), 'Sohar ID': str(r.Ref) || str(r['Row Key']), 'Policy Key': '', Source: str(r.Source) || 'Rotation', Amount: str(r.Amount) };
    });
    const groups = new Map();
    const group = (k, r) => {
      if (!groups.has(k)) groups.set(k, { policyKey: k.startsWith('payer:') ? '' : k, payer: '', payerId: '', last4: '', rows: [] });
      const g = groups.get(k);
      g.rows.push(r);
      if (!g.payer && str(r['Sohar Payer'])) g.payer = str(r['Sohar Payer']);
    };
    for (const r of mine) group(str(r['Policy Key']) || `payer:${str(r['Sohar Payer']).toLowerCase()}`, r);
    for (const r of rot) {
      const name = str(r['Sohar Payer']).toLowerCase();
      const hits = [...groups.entries()].filter(([, g]) => g.policyKey && g.rows.some((x) => str(x['Sohar Payer']).toLowerCase() === name));
      group(hits.length === 1 ? hits[0][0] : `payer:${name}`, r);
    }
    const out = [...groups.values()];
    for (const g of out) {
      g.rows.sort((a, b) => pulledMs(b) - pulledMs(a));
      g.latest = g.rows.length ? str(g.rows[0]['Pulled At']) : '';
      const [id, l4] = g.policyKey.split('#');
      g.payerId = str(id); g.last4 = str(l4);
    }
    return out.sort((a, b) => pulledMs(b.rows[0]) - pulledMs(a.rows[0]));
  }
  /**
   * Is the plan-year estimate due for a patient's Flow import row (section 4)? p: a built patient; v: the chart policy
   * sent (verifyRows entry) or null (discovery). ctx: {results (Sohar Results rows), history (Verification History),
   * flow (Flow Results), today, planYear (fees.planYearEstimateDue)}. The policy is the patient's latest trusted
   * verification of that policy (any policy for discovery). With none eligible on file, the Flow decides: due (the
   * Flow estimates only an eligible picked policy), except a VA / TriWest / TRICARE chart payer.
   * Returns {due, basis, policyKey, reason}.
   */
  function estimateDueFor(p, v, ctx) {
    const c = ctx || {};
    const key = C.keyOf(p);
    const mine = (c.results || []).filter((r) => mrnSame(C.keyOf(r), key));
    const pol = C.pickVerification(C.policyResults(mine, v || null));
    const govt = /veterans?'?s*affairs|va|tris*-?west|tris*-?care|champva/i;
    if (!pol || !/^complete.eligible/i.test(str(pol['Sohar Status']))) {
      if (govt.test(str(p && p['Primary Payer']))) return { due: false, basis: '', policyKey: '', reason: 'VA, TriWest or TRICARE: $0 expected' };
      return { due: true, basis: 'Unknown', policyKey: '', reason: 'no eligible verification on file: the Flow estimates the policy it picks if eligible' };
    }
    const pk = policyKey(pol['Sohar payerId'], pol['Sohar Member ID'] || (v && v.memberId));
    const hist = (c.history || []).filter((h) => mrnSame(h.MRN, key) && str(h['Policy Key']) === pk)
      .concat((c.flow || []).filter((f) => mrnSame(f.MRN, key) && (!str(f['Policy Key']) || str(f['Policy Key']) === pk)));
    const basis = resetBasisFor(c.flow, pk);
    const plan = typeof c.planYear === 'function' ? c.planYear({ ...pol, 'Reset Basis': basis || pol['Reset Basis'] }, hist, c.today || new Date()) : { due: true, basis: 'Unknown', reason: '' };
    return { ...plan, policyKey: pk };
  }
  /** The newest Reset Basis recorded for a policy in Flow Results rows ('' when none). */
  function resetBasisFor(flowRows, key) {
    const k = str(key);
    const r = (flowRows || []).filter((x) => k && str(x['Policy Key']) === k && /^(calendar|plan|unknown)$/i.test(str(x['Reset Basis'])))
      .sort((a, b) => updatedMs(b) - updatedMs(a))[0];
    return r ? str(r['Reset Basis']) : '';
  }

  // ---------- interactive HTML export (v0.31; v0.32 one page, docs/ARCHITECTURE.md section 7) ----------
  // One self-contained page in the add-in's design (its tokens, light and dark). v0.32 (Brendan, 11/13): one comprehensive
  // page, a tab per section; one large search right-aligned in the header that searches every section; Clinic and Service
  // line dropdowns for the whole page; on every column header a menu to sort A to Z or Z to A and filter by values or text;
  // the section's own dropdowns with counts. Appointments is the patient view per visit: its estimate and where it came
  // from, the deductible and OOP max before and after it, the patient's booked visits as a projection, and Verification,
  // Discovery and Cost estimate links. No network: the page's CSP allows only its inline style and script. Sohar import
  // files stay CSV. Callers pass sheet-safe columns only (LASFIR and MRN, never names, DOBs or member IDs).
  const HTML_EXPORT_COLS = {
    patients: ['LASFIR', 'MRN', 'State', 'Route', 'In Scope', 'Next Appt', 'Next Event', 'Service Line', 'POS', 'Primary Payer', 'Sohar payerId', 'Confidence', 'Secondary Payer', 'Flags'],
    appointments: VISIT_PLAN_COLS,
    results: ['LASFIR', 'Next Appt', 'Next Time', 'Next Provider', 'Next Event', 'Service Line', 'Next Location', 'MRN', 'Source', 'Sent As', 'Sohar Status', 'Work Queue', 'Action', 'Benefit Flag',
      'Plan', 'Network', 'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'OOP Max', 'OOP Remaining',
      'Est Code', 'Est Patient', 'Est 90792', 'Est 99214', 'Est 99215', 'Estimate Source', 'Estimate Date', 'Estimate Note', 'NextGen Payer',
      'Sohar Payer', 'Sohar payerId', 'Match', 'Discovery Follow-up', 'Sohar Link', 'Discovery Link', 'Cost Estimate Link', 'Sohar Date'],
    estimates: ['LASFIR', 'MRN', 'Appt Date', 'Time', 'Event', 'Service Line', 'Location', 'Provider', 'Sohar Payer', 'Plan', 'Network', 'Est Codes', 'Sohar Est Patient', 'Est 90792',
      'Est 99214', 'Est 99215', 'Estimate Source', 'Reset Basis', 'Source', 'Estimate Note', 'NextGen Verification Link', 'Discovered Verification Link', 'Last Updated'],
  };
  const BENEFIT_BARS = [['Deductible', 'Deductible', 'Deductible Remaining'], ['OOP max', 'OOP Max', 'OOP Remaining']];
  const AFTER_BARS = [['Deductible after this visit', 'Deductible', 'Ded Left After'], ['OOP max after this visit', 'OOP Max', 'OOP Left After']];
  // Per section: details (fields, bars, the per-patient projection) and schedule (which columns place a card).
  const HTML_EXPORT_VIEWS = {
    appointments: {
      details: { fields: ['Event', 'Service Line', 'Location', 'Provider', 'Sohar Status', 'Sohar Payer', 'Plan', 'Network', 'Copay', 'Coinsurance', 'Est Code', 'Est Patient',
        'Estimate Source', 'Allowed', 'Milestone', 'Estimate Note', 'Checked'], bars: BENEFIT_BARS.concat(AFTER_BARS), projection: true },
      schedule: { date: 'Appt Date', time: 'Time', column: 'Provider', person: 'LASFIR', id: 'MRN', visit: ['Event', 'Location'], kind: 'Service Line', status: 'Sohar Status',
        payer: ['Sohar Payer'], facts: [['Copay', 'Copay'], ['Coins', 'Coinsurance'], ['', 'Network']], bars: BENEFIT_BARS, milestone: 'Milestone',
        estimate: { code: 'Est Code', amount: 'Est Patient', perCode: [], source: 'Estimate Source' } },
    },
    results: {
      details: { fields: ['Source', 'Sent As', 'Sohar Status', 'Plan', 'Network', 'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance', 'NextGen Payer', 'Sohar Payer',
        'Sohar payerId', 'Match', 'Benefit Flag', 'Work Queue', 'Action', 'Discovery Follow-up', 'Sohar Date', 'Est Code', 'Est Patient', 'Est 90792', 'Est 99214', 'Est 99215',
        'Estimate Source', 'Estimate Date', 'Estimate Note'], bars: BENEFIT_BARS },
      schedule: { date: 'Next Appt', time: 'Next Time', column: 'Next Provider', person: 'LASFIR', id: 'MRN', visit: ['Next Event', 'Next Location'],
        kind: 'Sent As', status: 'Sohar Status', payer: ['Sohar Payer', 'NextGen Payer'], queue: 'Work Queue', action: 'Action',
        facts: [['Copay', 'Copay'], ['Coins', 'Coinsurance'], ['', 'Network'], ['', 'Benefit Flag']], bars: BENEFIT_BARS,
        estimate: { code: 'Est Code', amount: 'Est Patient', perCode: PER_CODE_COLS, source: 'Estimate Source' } },
    },
    estimates: {
      details: { fields: ['Appt Date', 'Event', 'Provider', 'Sohar Payer', 'Plan', 'Network', 'Est Codes', 'Sohar Est Patient', 'Est 90792', 'Est 99214', 'Est 99215',
        'Estimate Source', 'Reset Basis', 'Source', 'Estimate Note', 'Last Updated'], bars: [] },
    },
    history: {
      details: { fields: ['Pulled At', 'Kind', 'Sohar Status', 'Sohar Payer', 'Plan', 'Network', 'Coverage Level', 'Coverage Start', 'Coverage End', 'Copay',
        'Coinsurance', 'Benefit Flag', 'Policy Key', 'Projected Patient Responsibility', 'Estimate Note', 'Source', 'Related Sohar IDs'], bars: BENEFIT_BARS },
    },
  };
  // The page-wide dropdowns: each reads the first of its columns a section has (a section with none is not filtered by it).
  const HTML_GLOBAL_FILTERS = [{ id: 'clinic', label: 'Clinic', cols: ['Location', 'Next Location', 'Clinic'] }, { id: 'line', label: 'Service line', cols: ['Service Line'] }];
  const NO_EXPORT = /^(last name|first name|dob|date of birth|subscriber (first|last|dob)|member id|nextgen member id|sohar member id|chartmemberid|memberid|firstname|lastname|dateofbirth)$/i;
  const SOHAR_LINK = /^https:\/\/app\.soharhealth\.com\/[^\s"'<>]*$/i;
  const escH = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const REPORT_CSS = String.raw`:root{color-scheme:light;--page:#F3F4F1;--raised:#FFFFFF;--sunken:#F1F2EE;--line:#E3E6E1;--fg:#1B211F;--muted:#5E6863;--primary:#1E6B59;--on:#FFFFFF;--red:#B3261E;--amber:#8F5400;--blue:#2B5C8A}
:root[data-theme="dark"]{color-scheme:dark;--page:#131716;--raised:#1B2220;--sunken:#182019;--line:#2A3330;--fg:#ECEFED;--muted:#9BA59F;--primary:#6DC4A6;--on:#0E2A22;--red:#F19189;--amber:#E9B150;--blue:#8FB8E6}
*{box-sizing:border-box}[hidden]{display:none!important}body{margin:0;font:400 14px/1.4 "Segoe UI Variable Text","Segoe UI",system-ui,-apple-system,sans-serif;color:var(--fg);background:var(--page)}
:focus-visible{outline:2px solid var(--primary);outline-offset:2px}a{color:var(--primary)}.muted{color:var(--muted);font-size:12px}.n{font-variant-numeric:tabular-nums}
header{position:sticky;top:0;z-index:5;display:flex;flex-wrap:wrap;gap:10px 16px;align-items:center;padding:12px 20px;border-bottom:1px solid var(--line);background:var(--raised)}
.ttl{flex:0 1 auto;min-width:0}h1{margin:0;font-size:17px;font-weight:600}h2{font-size:16px;margin:0;font-weight:600}
.search{flex:1 1 560px;max-width:820px;margin-left:auto;position:relative;display:flex;align-items:center}
.search input{width:100%;height:52px;font-size:17px;padding:0 48px 0 46px;border-radius:10px;background:var(--sunken);border:1.5px solid var(--line);text-align:left}
.search input:focus{border-color:var(--primary);background:var(--raised)}
.search svg{position:absolute;left:15px;width:20px;height:20px;fill:none;stroke:var(--muted);stroke-width:1.7;stroke-linecap:round;pointer-events:none}
.search kbd{position:absolute;right:12px;font:600 12px/1 inherit;color:var(--muted);border:1px solid var(--line);border-radius:4px;padding:4px 7px;pointer-events:none}
.search input:focus+kbd,.search input:not(:placeholder-shown)+kbd{display:none}
.tools{display:flex;flex-wrap:wrap;gap:8px 12px;align-items:center;padding:10px 20px;border-bottom:1px solid var(--line);background:var(--raised)}
.tabs{display:flex;flex-wrap:wrap;gap:4px;flex:1 1 auto}.tabs button{border:1px solid transparent;background:transparent;border-radius:8px;padding:7px 12px;font-weight:600;color:var(--muted)}
.tabs button[aria-selected="true"]{background:color-mix(in srgb,var(--primary) 12%,transparent);border-color:color-mix(in srgb,var(--primary) 40%,transparent);color:var(--fg)}
.tabs .tc{font-weight:400;color:var(--muted);margin-left:6px;font-variant-numeric:tabular-nums}
.gf{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-left:auto}
.gf .flt select{font-size:13px;min-width:150px}
main{padding:16px 20px;display:grid;gap:18px}section{background:var(--raised);border:1px solid var(--line);border-radius:8px;padding:14px 16px;min-width:0}
input,select,button{font:inherit;font-size:12.5px;color:var(--fg);background:var(--sunken);border:1px solid var(--line);border-radius:6px;padding:5px 9px}
button{cursor:pointer}button:hover,input:focus,select:focus{border-color:var(--primary);outline:none}
.bar{display:flex;flex-wrap:wrap;align-items:center;gap:10px 14px}.lead{flex:1 1 auto;display:flex;flex-wrap:wrap;align-items:baseline;gap:4px 12px;min-width:0}
.count{color:var(--muted);font-size:12.5px;font-variant-numeric:tabular-nums}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:8px;overflow:hidden;background:var(--sunken)}
.seg button{border:0;border-radius:0;background:transparent;padding:8px 14px;font-weight:600;color:var(--muted)}.seg button[aria-pressed="true"]{background:var(--primary);color:var(--on)}
.note{margin:6px 0 0}.filters{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:12px 0}
.flt{display:inline-flex;align-items:center;gap:4px;font-size:12px;color:var(--muted);background:var(--sunken);border:1px solid var(--line);border-radius:999px;padding:1px 4px 1px 11px}
.flt select{border:0;background:transparent;padding:4px 4px;border-radius:999px;max-width:240px}.flt.on{border-color:var(--primary);color:var(--fg);background:color-mix(in srgb,var(--primary) 12%,transparent)}
.clear{border-radius:999px;background:transparent;color:var(--primary);font-weight:600;border-color:transparent}.clear:hover{border-color:var(--primary)}
.wrap{overflow:auto;max-height:72vh;border:1px solid var(--line);border-radius:6px}
table{border-collapse:collapse;width:100%;font-size:12px}th{position:sticky;top:0;z-index:1;background:var(--sunken);text-align:left;font-weight:600;color:var(--muted);border-bottom:1px solid var(--line);white-space:nowrap;padding:0}
th .hc{display:flex;align-items:stretch}th .sort{all:unset;box-sizing:border-box;display:block;flex:1 1 auto;padding:7px 16px 7px 7px;cursor:pointer;position:relative}th .sort:hover{color:var(--fg)}th .sort:focus-visible,th .cm:focus-visible{outline:2px solid var(--primary);outline-offset:-2px}
th .sort::after{content:"\2195";position:absolute;right:3px;opacity:.35}th[aria-sort="ascending"] .sort::after{content:"\25B2";opacity:1;color:var(--primary)}th[aria-sort="descending"] .sort::after{content:"\25BC";opacity:1;color:var(--primary)}
th .cm{all:unset;box-sizing:border-box;cursor:pointer;padding:0 7px;display:flex;align-items:center;color:var(--muted);border-left:1px solid transparent}th .cm::before{content:"\25BE"}th .cm:hover{color:var(--fg);border-left-color:var(--line)}
th.fl{color:var(--primary)}th.fl .cm{color:var(--primary)}th.fl .cm::before{content:"\25BE\2022"}
th[aria-sort]{color:var(--fg)}td{padding:6px 7px;border-bottom:1px solid var(--line);vertical-align:top;white-space:nowrap;background:var(--raised)}td.wr{white-space:normal;min-width:200px}tr[data-r]:hover td{background:color-mix(in srgb,var(--primary) 6%,var(--raised))}
.c1{position:sticky;left:0;z-index:1;box-shadow:1px 0 0 var(--line)}.hd .c1{left:34px}td.tgc{position:sticky;left:0;z-index:1}th.tgc{left:0}th.c1,th.tgc{z-index:3}
.tgc{width:34px;padding:4px 4px 4px 6px}.tg{width:24px;height:24px;padding:0;line-height:1;color:var(--muted)}.tg::before{content:"\25B8"}.tg[aria-expanded="true"]::before{content:"\25BE";color:var(--primary)}
tr.dt>td{background:var(--sunken);padding:10px 12px 12px 40px;white-space:normal}tr.dt .panel{position:sticky;left:40px;width:min(1100px,calc(100vw - 110px))}
.panel{display:grid;grid-template-columns:minmax(200px,280px) 1fr;gap:10px 22px;align-items:start}.panel+.panel{margin-top:10px;padding-top:10px;border-top:1px dashed var(--line)}
.panel dl{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px 16px;margin:0}.panel dt{font-size:11px;color:var(--muted)}.panel dd{margin:0;overflow-wrap:anywhere}
.meters{display:grid;gap:9px}.meter{display:grid;gap:3px}.ml{display:flex;justify-content:space-between;gap:8px;font-size:11.5px}.ml span:first-child{color:var(--muted)}
.track{height:6px;border-radius:3px;background:var(--line);overflow:hidden}.track i{display:block;height:100%;background:var(--blue)}.meter.met .track i{background:var(--primary)}
.links{display:flex;flex-wrap:wrap;gap:4px 12px}.lk{font-weight:600;text-decoration:none;white-space:nowrap}.lk:hover{text-decoration:underline}.panel .links{grid-column:1/-1}
.proj{grid-column:1/-1}.proj h4{margin:4px 0 6px;font-size:12.5px}.proj table{font-size:11.5px;background:var(--raised);border:1px solid var(--line);border-radius:6px}.proj th{position:static;padding:5px 7px}.proj td{padding:4px 7px;background:transparent}.proj tr.me td{background:color-mix(in srgb,var(--primary) 10%,transparent);font-weight:600}
.badge{display:inline-block;font-size:11px;font-weight:600;padding:1px 8px;border-radius:999px;border:1px solid var(--line);background:var(--sunken);color:var(--muted);max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.badge.ok{color:var(--primary);border-color:color-mix(in srgb,var(--primary) 45%,transparent);background:color-mix(in srgb,var(--primary) 10%,transparent)}
.badge.bad{color:var(--red);border-color:color-mix(in srgb,var(--red) 45%,transparent);background:color-mix(in srgb,var(--red) 9%,transparent)}
.badge.warn{color:var(--amber);border-color:color-mix(in srgb,var(--amber) 45%,transparent);background:color-mix(in srgb,var(--amber) 10%,transparent)}
.day{margin:14px 0 4px}.day h3{display:flex;gap:10px;align-items:baseline;margin:0 0 8px;font-size:14px;font-weight:600}
.cols{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(260px,1fr);gap:12px;overflow-x:auto;padding-bottom:6px}
.pcol{background:var(--sunken);border:1px solid var(--line);border-radius:8px;padding:8px;display:flex;flex-direction:column;gap:8px;min-width:0}
.pcol h4{display:flex;justify-content:space-between;gap:8px;margin:0 2px;font-size:12.5px;font-weight:600}
.card{background:var(--raised);border:1px solid var(--line);border-left:3px solid var(--line);border-radius:6px;padding:8px 10px;display:grid;gap:6px;font-size:12.5px;min-width:0}
.card.ok{border-left-color:var(--primary)}.card.bad{border-left-color:var(--red)}.card.warn{border-left-color:var(--amber)}
.top{display:flex;gap:8px;align-items:baseline}.tm{font-weight:600;font-variant-numeric:tabular-nums;min-width:62px}.who{font-weight:700;letter-spacing:.03em}.top .muted{margin-left:auto}
.todo{background:color-mix(in srgb,var(--amber) 12%,transparent);border-radius:4px;padding:4px 7px;font-size:12px}.todo b{color:var(--amber)}
.ms{background:color-mix(in srgb,var(--primary) 12%,transparent);border-radius:4px;padding:4px 7px;font-size:12px;font-weight:600}
.chk{display:grid;gap:5px;border-top:1px dashed var(--line);padding-top:6px}.ch{display:flex;flex-wrap:wrap;gap:4px 6px;align-items:center}.kind{font-size:11.5px;color:var(--muted)}
.facts{color:var(--muted);font-size:11.5px}.est{font-size:12px;background:color-mix(in srgb,var(--blue) 9%,transparent);border-radius:4px;padding:4px 7px}.est b{font-variant-numeric:tabular-nums}.card details summary{cursor:pointer;color:var(--muted);font-size:11.5px}.card details .panel{display:block;margin-top:6px}.card details .meters,.card details .links{display:none}
.cmenu{position:absolute;z-index:20;width:290px;max-height:420px;display:flex;flex-direction:column;gap:8px;background:var(--raised);border:1px solid var(--line);border-radius:8px;padding:10px;box-shadow:0 8px 24px rgba(0,0,0,.18);font-size:12.5px}
.cmenu .mb{display:flex;gap:6px}.cmenu .mb button{flex:1 1 0}.cmenu input[type=search]{width:100%}.cmenu .mv{overflow:auto;display:grid;gap:2px;border:1px solid var(--line);border-radius:6px;padding:6px;min-height:60px}
.cmenu label{display:flex;gap:6px;align-items:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.cmenu label .muted{margin-left:auto}
.empty{padding:18px 4px;text-align:center}.sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
footer{padding:10px 20px 20px;color:var(--muted);font-size:11.5px}
@media (max-width:700px){.search{flex-basis:100%;max-width:none}.panel{grid-template-columns:1fr}.gf{margin-left:0}}
@media print{header{position:static}.seg,.search,.tools,.filters,#theme,.tg,th .cm{display:none}section[hidden]{display:block}.wrap{max-height:none;overflow:visible;border:0}body{background:#fff}}`;
  // The page's script: one search and the page dropdowns for every section, tabs, column menus (sort A to Z / Z to A,
  // filter by values or text), section dropdowns, details rows, table / schedule, light / dark. No network, no eval.
  const REPORT_JS = String.raw`(function(){var d=document,r=d.documentElement,each=function(l,f){Array.prototype.forEach.call(l,f)};
try{var t=localStorage.getItem("soharTheme");r.dataset.theme=t==="dark"||(!t&&window.matchMedia&&matchMedia("(prefers-color-scheme: dark)").matches)?"dark":"light"}catch(e){}
d.getElementById("theme").addEventListener("click",function(){r.dataset.theme=r.dataset.theme==="dark"?"light":"dark";try{localStorage.setItem("soharTheme",r.dataset.theme)}catch(e){}});
var Q=d.getElementById("q"),gs=d.querySelectorAll("select[data-g]"),gclr=d.getElementById("gclear"),tabs=d.querySelectorAll("[data-tab]"),secs=[];
function esc(x){return String(x).replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]})}
function txt(tr,i){var c=tr.cells[i];return c?c.textContent.trim():""}
function key(tr,i){var x=txt(tr,i),m;if(x==="")return null;
m=x.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);if(m)return m[3]+("0"+m[1]).slice(-2)+("0"+m[2]).slice(-2)+x.slice(m[0].length);if(/^\d{4}-\d{2}-\d{2}/.test(x))return x;
m=x.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap])\.?m?\.?$/i)||x.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);if(m){var h=+m[1];if(m[3])h=h%12+(/p/i.test(m[3])?12:0);return h*60+(+m[2])}
var n=parseFloat(x.replace(/[$,%\s]/g,""));return /^[-$\d.,%\s]+$/.test(x)&&!isNaN(n)?n:x.toLowerCase()}
var M=d.createElement("div");M.className="cmenu";M.hidden=true;M.setAttribute("role","dialog");M.setAttribute("aria-label","Sort and filter");d.body.appendChild(M);var cur=null;
function closeMenu(){if(M.hidden)return;M.hidden=true;if(cur&&cur.btn){cur.btn.setAttribute("aria-expanded","false");cur.btn.focus()}cur=null}
d.addEventListener("keydown",function(e){var a=d.activeElement;if(e.key==="/"&&!(a&&/^(INPUT|SELECT|TEXTAREA)$/.test(a.tagName))){e.preventDefault();Q.focus()}if(e.key==="Escape")closeMenu()});
d.addEventListener("click",function(e){if(!M.hidden&&!M.contains(e.target)&&!(e.target.closest&&e.target.closest(".cm")))closeMenu()});
function openMenu(S,i,btn){var was=cur&&cur.btn===btn;closeMenu();if(was)return;cur={S:S,i:i,btn:btn};var f=S.cf[i]||(S.cf[i]={text:"",off:{}}),counts={},vals=[];
S.rows.forEach(function(tr){var x=txt(tr,i);if(!(x in counts)){counts[x]=0;vals.push(x)}counts[x]++});
vals.sort(function(a,b){return a===""?1:b===""?-1:a.localeCompare(b,undefined,{numeric:true})});var shown=vals.slice(0,500);
M.innerHTML='<div class="mb"><button type="button" data-s="a">Sort A to Z</button><button type="button" data-s="d">Sort Z to A</button></div><input type="search" data-t placeholder="Contains" aria-label="Show rows that contain"><div class="mv"><label><input type="checkbox" data-all> <b>Select all</b></label>'
+shown.map(function(v,k){return '<label title="'+esc(v)+'"><input type="checkbox" data-k="'+k+'"'+(f.off[v]?"":" checked")+"> "+esc(v||"(blank)")+' <span class="muted">'+counts[v]+"</span></label>"}).join("")+(vals.length>shown.length?'<span class="muted">First 500 values; use Contains for the rest.</span>':"")
+'</div><div class="mb"><button type="button" data-x>Clear filter</button><button type="button" data-c>Close</button></div>';
var all=M.querySelector("[data-all]"),boxes=M.querySelectorAll("[data-k]"),T=M.querySelector("[data-t]");
function syncAll(){var on=0;each(boxes,function(b){if(b.checked)on++});all.checked=on===boxes.length;all.indeterminate=on>0&&on<boxes.length}syncAll();T.value=f.text;
each(M.querySelectorAll("[data-s]"),function(b){b.addEventListener("click",function(){S.sort(i,b.dataset.s==="a");closeMenu()})});
T.addEventListener("input",function(){f.text=T.value.trim().toLowerCase();S.apply()});
each(boxes,function(b){b.addEventListener("change",function(){var v=shown[+b.dataset.k];if(b.checked)delete f.off[v];else f.off[v]=1;syncAll();S.apply()})});
all.addEventListener("change",function(){each(boxes,function(b){b.checked=all.checked;var v=shown[+b.dataset.k];if(all.checked)delete f.off[v];else f.off[v]=1});S.apply()});
M.querySelector("[data-x]").addEventListener("click",function(){f.text="";f.off={};T.value="";each(boxes,function(b){b.checked=true});syncAll();S.apply()});
M.querySelector("[data-c]").addEventListener("click",closeMenu);
var b=btn.getBoundingClientRect(),w=d.documentElement.clientWidth;M.style.top=(b.bottom+window.scrollY+4)+"px";M.style.left=Math.max(8,Math.min(b.right-290,w-300))+window.scrollX+"px";
M.hidden=false;btn.setAttribute("aria-expanded","true");T.focus()}
each(d.querySelectorAll("section[data-sec]"),function(s){var tb=s.querySelector("tbody"),rows=Array.prototype.slice.call(tb.querySelectorAll("tr[data-r]")),det={},
fs=s.querySelectorAll("select[data-f]"),cnt=s.querySelector("[data-n]"),clr=s.querySelector("[data-clear]"),empty=s.querySelector("[data-empty]"),
cards=s.querySelectorAll(".card[data-rows]"),seg=s.querySelectorAll("[data-v]"),view="table",gcol={},tc=d.querySelector('[data-tc="'+s.id+'"]');
each(gs,function(g){var c=s.getAttribute("data-g-"+g.dataset.g);if(c!==null&&c!=="")gcol[g.dataset.g]=+c});
each(tb.querySelectorAll("tr[data-d]"),function(x){det[x.dataset.d]=x});
var S={s:s,rows:rows,cf:{}};
S.apply=function(){var v=(Q.value||"").trim().toLowerCase(),n=0,c=0,vis={},on=false,gv=[];
each(gs,function(g){if(g.value!==""&&gcol[g.dataset.g]!=null)gv.push([gcol[g.dataset.g],g.value])});
each(fs,function(f){var a=f.value!=="";f.parentNode.classList.toggle("on",a);if(a)on=true});
var cfs=[];for(var k in S.cf){var f=S.cf[k],any=false;for(var z in f.off){any=true;break}if(f.text||any)cfs.push([+k,f]);var th=s.querySelector('th[data-c="'+k+'"]');if(th)th.classList.toggle("fl",!!(f.text||any))}if(cfs.length)on=true;
rows.forEach(function(tr){var ok=!v||tr.textContent.toLowerCase().indexOf(v)>=0;
for(var j=0;ok&&j<gv.length;j++)if(txt(tr,gv[j][0])!==gv[j][1])ok=false;
each(fs,function(f){if(ok&&f.value!==""&&txt(tr,+f.dataset.f)!==f.value)ok=false});
for(var q=0;ok&&q<cfs.length;q++){var x=txt(tr,cfs[q][0]),F=cfs[q][1];if(F.off[x]||(F.text&&x.toLowerCase().indexOf(F.text)<0))ok=false}
tr.hidden=!ok;var dt=det[tr.dataset.r],b=tr.querySelector(".tg");if(dt)dt.hidden=!ok||!b||b.getAttribute("aria-expanded")!=="true";if(ok){n++;vis[tr.dataset.r]=1}});
each(cards,function(k){var ok=k.dataset.rows.split(" ").some(function(i){return vis[i]});k.hidden=!ok;if(ok)c++});
each(s.querySelectorAll(".pcol"),function(p){var m=p.querySelectorAll(".card:not([hidden])").length;p.hidden=!m;p.querySelector("[data-pc]").textContent=m});
each(s.querySelectorAll(".day"),function(y){var m=y.querySelectorAll(".card:not([hidden])").length;y.hidden=!m;y.querySelector("[data-dc]").textContent=m+(m===1?" visit":" visits")});
var sc=view==="schedule";cnt.textContent=sc?c+" of "+cards.length+(cards.length===1?" visit":" visits"):n+" of "+rows.length+(rows.length===1?" row":" rows");
if(tc)tc.textContent=n===rows.length?String(rows.length):n+" of "+rows.length;if(clr)clr.hidden=!on;empty.hidden=sc?c>0:n>0};
S.sort=function(i,asc){var th=s.querySelector('th[data-c="'+i+'"]');each(s.querySelectorAll("th[data-c]"),function(h){h.removeAttribute("aria-sort")});if(th)th.setAttribute("aria-sort",asc?"ascending":"descending");
rows.sort(function(a,b){var x=key(a,i),y=key(b,i);if(x===null||y===null)return x===y?0:x===null?1:-1;if(typeof x==="number"&&typeof y==="number")return asc?x-y:y-x;x=String(x);y=String(y);return asc?x.localeCompare(y,undefined,{numeric:true}):y.localeCompare(x,undefined,{numeric:true})});
rows.forEach(function(tr){tb.appendChild(tr);var dt=det[tr.dataset.r];if(dt)tb.appendChild(dt)})};
function setView(x){view=x;each(seg,function(b){b.setAttribute("aria-pressed",b.dataset.v===x?"true":"false")});each(s.querySelectorAll("[data-view]"),function(p){p.hidden=p.dataset.view!==x});try{localStorage.setItem("soharView",x)}catch(e){}S.apply()}
each(fs,function(f){f.addEventListener("change",S.apply)});
if(clr)clr.addEventListener("click",function(){each(fs,function(f){f.value=""});S.cf={};S.apply()});
each(seg,function(b){b.addEventListener("click",function(){setView(b.dataset.v)})});
each(tb.querySelectorAll(".tg"),function(b){b.addEventListener("click",function(){var o=b.getAttribute("aria-expanded")!=="true",dt=det[b.closest("tr").dataset.r];b.setAttribute("aria-expanded",o?"true":"false");if(dt)dt.hidden=!o})});
each(s.querySelectorAll("th[data-c]"),function(th){var i=+th.dataset.c;th.querySelector(".sort").addEventListener("click",function(){S.sort(i,th.getAttribute("aria-sort")!=="ascending")});
th.querySelector(".cm").addEventListener("click",function(e){e.stopPropagation();openMenu(S,i,e.currentTarget)})});
S.init=function(){if(seg.length){var sv="schedule";try{sv=localStorage.getItem("soharView")||sv}catch(e){}setView(sv==="table"?"table":"schedule")}else S.apply()};
secs.push(S)});
function applyAll(){var on=!!Q.value;each(gs,function(g){var a=g.value!=="";g.parentNode.classList.toggle("on",a);if(a)on=true});secs.forEach(function(S){S.apply()});if(gclr)gclr.hidden=!on}
Q.addEventListener("input",applyAll);each(gs,function(g){g.addEventListener("change",applyAll)});
if(gclr)gclr.addEventListener("click",function(){Q.value="";each(gs,function(g){g.value=""});applyAll();Q.focus()});
function showTab(k){var hit=false;each(tabs,function(b){var on=b.dataset.tab===k;if(on)hit=true;b.setAttribute("aria-selected",on?"true":"false")});if(!hit&&tabs.length)return showTab(tabs[0].dataset.tab);
secs.forEach(function(S){S.s.hidden=tabs.length>0&&S.s.id!==k});try{localStorage.setItem("soharTab",k)}catch(e){}}
each(tabs,function(b){b.addEventListener("click",function(){closeMenu();showTab(b.dataset.tab)})});
secs.forEach(function(S){S.init()});applyAll();var st="";try{st=localStorage.getItem("soharTab")||""}catch(e){}if(tabs.length)showTab(st||tabs[0].dataset.tab)})();`;
  const numOf = (v) => { const t = str(v).replace(/[$,\s]/g, ''); return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : null; };
  const moneyOf = (n) => `$${Math.round(n).toLocaleString('en-US')}`;
  /** ok / bad / warn / '' for a Sohar status. */
  function statusTone(s) {
    const t = str(s).toLowerCase();
    if (!t) return '';
    if (/inelig|inactive|error|fail|not found|terminat|denied|no coverage|invalid/.test(t)) return 'bad';
    if (/pending|progress|queued|submitted|processing/.test(t)) return 'warn';
    return /eligib|active|complete|verified/.test(t) ? 'ok' : '';
  }
  /** "Verification", "Discovery" or "Cost estimate", from the Sohar URL (then the row's Sent As / Kind). */
  function soharLinkLabel(url, row) {
    const p = str(url).toLowerCase();
    if (/discover/.test(p)) return 'Discovery';
    if (/estimate|cost/.test(p)) return 'Cost estimate';
    if (/verif/.test(p)) return 'Verification';
    return /discover/i.test(str(row && (row['Sent As'] || row.Kind))) ? 'Discovery' : 'Sohar';
  }
  // v0.32: a link column says what it is; Sohar Link (and others) fall back to the URL.
  const LINK_LABELS = { 'Discovery Link': 'Discovery', 'Cost Estimate Link': 'Cost estimate', 'NextGen Verification Link': 'Verification', 'Discovered Verification Link': 'Verification (discovered plan)' };
  const linkLabelOf = (col, u, row) => LINK_LABELS[col] || soharLinkLabel(u, row);
  const soharA = (u, row, col) => `<a class="lk" href="${escH(u)}" target="_blank" rel="noopener noreferrer">${escH(linkLabelOf(col, u, row))} &#8599;</a>`;
  /** A labelled bar: how much of the total is met, and what's left. '' when the row has neither figure. */
  function benefitMeter(label, total, left) {
    const T = numOf(total), L = numOf(left);
    if (T == null && L == null) return '';
    const met = T != null && L != null ? Math.max(0, Math.min(T, T - L)) : null;
    const pct = T > 0 && met != null ? Math.round((met / T) * 100) : 0;
    const done = T !== 0 && L != null && L <= 0;
    const txt = T === 0 ? 'None' : done ? `Met${T != null ? ` (${moneyOf(T)})` : ''}` : L != null && T != null ? `${moneyOf(L)} left of ${moneyOf(T)}` : L != null ? `${moneyOf(L)} left` : moneyOf(T);
    return `<div class="meter${done ? ' met' : ''}"><div class="ml"><span>${escH(label)}</span><span class="n">${escH(txt)}</span></div><div class="track" role="img" aria-label="${escH(`${label}: ${txt}`)}"><i style="width:${done ? 100 : pct}%"></i></div></div>`;
  }
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  /** {key, label} for MM/DD/YYYY or YYYY-MM-DD; null otherwise. */
  function reportDay(v) {
    const t = str(v);
    let m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/), y, mo, dd;
    if (m) [y, mo, dd] = [+m[3], +m[1], +m[2]];
    else if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) [y, mo, dd] = [+m[1], +m[2], +m[3]];
    else return null;
    return { key: y * 10000 + mo * 100 + dd, label: `${DOW[new Date(y, mo - 1, dd).getDay()]} ${mo}/${dd}` };
  }
  /** Minutes after midnight for "9:00 AM" / "14:30"; late for anything else, so it sorts last. */
  function reportMinutes(v) {
    const m = str(v).match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap])?/i);
    if (!m) return 1e6;
    let h = +m[1];
    if (m[3]) h = (h % 12) + (/p/i.test(m[3]) ? 12 : 0);
    return h * 60 + +m[2];
  }
  const SEARCH_SVG = '<svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="7" cy="7" r="5"></circle><path d="M11 11l4 4"></path></svg>';
  /**
   * A self-contained interactive HTML page. o: {title, subtitle, generated (text), version, globals (page dropdowns,
   * default HTML_GLOBAL_FILTERS; [] for none), sections: [{id, title, cols, rows, filters (column names with a dropdown),
   * note, view ({details, schedule}, see HTML_EXPORT_VIEWS)}]}. Two or more sections become tabs. Columns that would carry a
   * name, DOB or member ID are dropped, and the details and schedule read only the columns kept. Cells holding a Sohar URL
   * become labelled links (Verification, Discovery, Cost estimate). Returns the HTML text.
   */
  function htmlReport(o) {
    const c = o || {};
    const globals = (c.globals || HTML_GLOBAL_FILTERS).map((g) => ({ ...g, values: new Map() }));
    const tabbed = (c.sections || []).length > 1;
    const secs = (c.sections || []).map((s, si) => {
      const cols = (s.cols || []).filter((k) => !NO_EXPORT.test(str(k)));
      const has = (k) => !!k && cols.includes(k);
      const g = (r, k) => (has(k) ? str(r[k]) : '');
      const rows = s.rows || [];
      const sid = escH(s.id || `s${si}`);
      const view = s.view || {}, det = view.details, sp = view.schedule && has(view.schedule.date) ? view.schedule : null;
      const off = det ? 1 : 0;
      // Page dropdowns: the first of each one's columns this section has, and its values (counted once per section row).
      const gattr = globals.map((gl) => {
        const col = gl.cols.find(has);
        if (!col) return '';
        for (const r of rows) { const v = g(r, col); if (v) gl.values.set(v, (gl.values.get(v) || 0) + 1); }
        return ` data-g-${gl.id}="${cols.indexOf(col) + off}"`;
      }).join('');
      const linksOf = (r) => { const seen = new Set(); return cols.map((k) => [k, str(r[k])]).filter(([, u]) => SOHAR_LINK.test(u) && !seen.has(u) && seen.add(u)).map(([k, u]) => soharA(u, r, k)).join(''); };
      const meters = (r, bars) => (bars || []).filter(([, a, b]) => has(a) || has(b)).map(([l, a, b]) => benefitMeter(l, g(r, a), g(r, b))).join('');
      // v0.32: the patient's booked visits in order, this one highlighted (Appointments).
      const byMrn = new Map();
      if (det && det.projection) for (const r of rows) { const k = g(r, 'MRN'); if (!byMrn.has(k)) byMrn.set(k, []); byMrn.get(k).push(r); }
      const projection = (r) => {
        const list = byMrn.get(g(r, 'MRN')) || [];
        if (list.length < 1) return '';
        const total = list.reduce((a, x) => a + (numOf(g(x, 'Est Patient')) || 0), 0), priced = list.filter((x) => numOf(g(x, 'Est Patient')) != null).length;
        const tr = (x) => `<tr${x === r ? ' class="me"' : ''}><td>${escH(g(x, 'Appt Date'))} ${escH(g(x, 'Time'))}</td><td>${escH(g(x, 'Event'))}</td><td>${escH(g(x, 'Est Code'))}</td><td class="n">${escH(g(x, 'Est Patient') || '—')}</td><td>${escH(g(x, 'Estimate Source'))}</td><td class="n">${escH(g(x, 'Ded Left After'))}</td><td class="n">${escH(g(x, 'OOP Left After'))}</td><td>${escH(g(x, 'Milestone'))}</td></tr>`;
        return `<div class="proj"><h4>Booked visits: ${list.length}${priced ? ` · estimated ${escH(money(total))} for ${priced}` : ''}</h4><table><thead><tr><th>Visit</th><th>Event</th><th>Code</th><th>Estimate</th><th>Source</th><th>Deductible left</th><th>OOP left</th><th></th></tr></thead><tbody>${list.map(tr).join('')}</tbody></table></div>`;
      };
      const panel = (r) => {
        const f = (det.fields || []).filter(has).map((k) => [k, str(r[k])]).filter(([, v]) => v);
        const m = meters(r, det.bars), l = linksOf(r);
        return `<div class="panel"><div class="meters">${m || '<span class="muted">No deductible or OOP figures.</span>'}</div><dl>${f.map(([k, v]) => `<div><dt>${escH(k)}</dt><dd>${escH(v)}</dd></div>`).join('')}</dl>${l ? `<div class="links">${l}</div>` : ''}${det.projection ? projection(r) : ''}</div>`;
      };
      const cell = (r, k) => { const t = str(r[k]); return SOHAR_LINK.test(t) ? soharA(t, r, k) : escH(t); };
      const isNum = (k) => /amount|balance|copay|deductible|oop|remaining|max|responsibility|est |paid|billed|share|allowed|left after/i.test(k);
      const tdc = (k, i) => [i === 0 ? 'c1' : '', isNum(k) ? 'n' : '', /action|flag|note|detail|follow-up|reason|task|milestone/i.test(k) ? 'wr' : ''].filter(Boolean).join(' ');
      const filters = (s.filters || []).filter(has).map((k) => {
        const counts = new Map();
        for (const r of rows) { const v = str(r[k]); if (v) counts.set(v, (counts.get(v) || 0) + 1); }
        const vals = [...counts.keys()].sort().slice(0, 300);
        return `<label class="flt">${escH(k)}<select data-f="${cols.indexOf(k) + off}"><option value="">All</option>${vals.map((v) => `<option value="${escH(v)}">${escH(v)} (${counts.get(v)})</option>`).join('')}</select></label>`;
      }).join('');
      const who = (r, i) => escH(str(r.LASFIR) || `row ${i + 1}`);
      const body = rows.map((r, i) => `<tr data-r="${i}">${det ? `<td class="tgc"><button type="button" class="tg" aria-expanded="false" aria-controls="${sid}-d${i}" aria-label="Details for ${who(r, i)}"></button></td>` : ''}${cols.map((k, ci) => `<td${tdc(k, ci) ? ` class="${tdc(k, ci)}"` : ''}>${cell(r, k)}</td>`).join('')}</tr>`
        + (det ? `<tr class="dt" data-d="${i}" id="${sid}-d${i}" hidden><td colspan="${cols.length + 1}">${panel(r)}</td></tr>` : '')).join('\n');
      const th = (k, i) => `<th scope="col" data-c="${i + off}"${i === 0 ? ' class="c1"' : ''}><div class="hc"><button type="button" class="sort">${escH(k)}</button><button type="button" class="cm" aria-haspopup="dialog" aria-expanded="false" aria-label="Sort and filter ${escH(k)}"></button></div></th>`;
      const table = `<div class="wrap"><table${det ? ' class="hd"' : ''}><thead><tr>${det ? '<th class="tgc"><span class="sr">Details</span></th>' : ''}${cols.map(th).join('')}</tr></thead><tbody>
${body}
</tbody></table></div>`;
      let schedule = '';
      if (sp) {
        const visits = new Map();
        rows.forEach((r, i) => {
          const k = [g(r, sp.date), g(r, sp.time), g(r, sp.column), g(r, sp.id) || g(r, sp.person) || `#${i}`].join('|');
          if (!visits.has(k)) visits.set(k, { r, idx: [] });
          visits.get(k).idx.push(i);
        });
        const days = new Map();
        for (const v of visits.values()) {
          const day = reportDay(g(v.r, sp.date)) || { key: 99999999, label: 'No upcoming visit' };
          if (!days.has(day.key)) days.set(day.key, { label: day.label, cols: new Map() });
          const col = g(v.r, sp.column) || 'No provider', D = days.get(day.key);
          if (!D.cols.has(col)) D.cols.set(col, []);
          D.cols.get(col).push(v);
        }
        const card = (v) => {
          const rs = v.idx.map((i) => rows[i]);
          const queue = rs.map((r) => g(r, sp.queue)).find(Boolean) || '', action = rs.map((r) => g(r, sp.action)).find(Boolean) || '';
          const ms = sp.milestone ? rs.map((r) => g(r, sp.milestone)).find(Boolean) || '' : '';
          const tones = rs.map((r) => statusTone(g(r, sp.status)));
          const tone = queue ? 'warn' : ['ok', 'bad', 'warn'].find((t) => tones.includes(t)) || '';
          const visit = (sp.visit || []).map((k) => g(v.r, k)).filter(Boolean).join(' · ');
          const es = sp.estimate, eRow = es ? rs.find((r) => g(r, es.amount) || es.perCode.some((k) => g(r, k))) : null;
          // The visit's own code first, then the other plan-year quotes (patient share).
          const codesOf = (r) => es.perCode.filter((k) => g(r, k) && !(g(r, es.amount) && k === `Est ${g(r, es.code)}`)).map((k) => escH(k.replace(/^Est /, '') + ' ' + g(r, k))).join(' · ');
          const estLine = !eRow ? '' : '<div class="est">' + (g(eRow, es.amount) ? 'Estimate ' + escH(g(eRow, es.code)) + ' <b>' + escH(g(eRow, es.amount)) + '</b>' : 'Plan-year estimate')
            + (codesOf(eRow) ? '<span class="muted"> · plan year ' + codesOf(eRow) + '</span>' : '') + (g(eRow, es.source) ? '<span class="muted"> · ' + escH(g(eRow, es.source)) + '</span>' : '') + '</div>';
          const checks = rs.map((r) => {
            const st = g(r, sp.status), payer = (sp.payer || []).map((k) => g(r, k)).find(Boolean) || '';
            const facts = (sp.facts || []).map(([l, k]) => [l, g(r, k)]).filter(([, x]) => x).map(([l, x]) => escH(l ? `${l} ${x}` : x)).join(' · ');
            const m = meters(r, sp.bars), l = linksOf(r);
            return `<div class="chk"><div class="ch"><span class="badge ${statusTone(st)}">${escH(st || 'No result')}</span>${g(r, sp.kind) ? `<span class="kind">${escH(g(r, sp.kind))}</span>` : ''}${payer ? `<span>${escH(payer)}</span>` : ''}</div>`
              + `${facts ? `<div class="facts">${facts}</div>` : ''}${m ? `<div class="meters">${m}</div>` : ''}${l ? `<div class="links">${l}</div>` : ''}</div>`;
          }).join('');
          return `<article class="card ${tone}" data-rows="${v.idx.join(' ')}"><div class="top"><span class="tm">${escH(g(v.r, sp.time) || '—')}</span><span class="who">${escH(g(v.r, sp.person))}</span><span class="muted">${escH(g(v.r, sp.id))}</span></div>`
            + `${visit ? `<div class="muted">${escH(visit)}</div>` : ''}${queue ? `<div class="todo"><b>${escH(queue)}</b>${action ? `: ${escH(action)}` : ''}</div>` : ''}${estLine}${ms ? `<div class="ms">${escH(ms)}</div>` : ''}${checks}`
            + `${det ? `<details><summary>Details</summary>${rs.map(panel).join('')}</details>` : ''}</article>`;
        };
        const byTime = (a, b) => reportMinutes(g(a.r, sp.time)) - reportMinutes(g(b.r, sp.time)) || g(a.r, sp.person).localeCompare(g(b.r, sp.person));
        schedule = [...days.entries()].sort((a, b) => a[0] - b[0]).map(([, D]) => `<div class="day"><h3>${escH(D.label)}<span class="muted" data-dc></span></h3><div class="cols">`
          + [...D.cols.entries()].sort((a, b) => (a[0] === 'No provider') - (b[0] === 'No provider') || a[0].localeCompare(b[0]))
            .map(([name, vs]) => `<div class="pcol"><h4><span>${escH(name)}</span><span class="muted" data-pc></span></h4>${vs.sort(byTime).map(card).join('')}</div>`).join('')
          + '</div></div>').join('\n');
      }
      const seg = sp ? '<div class="seg" role="group" aria-label="View"><button type="button" data-v="schedule" aria-pressed="true">Schedule</button><button type="button" data-v="table" aria-pressed="false">Table</button></div>' : '';
      return { id: sid, title: s.title, n: rows.length, html: `<section data-sec="${si}" id="${sid}" aria-labelledby="${sid}-h"${gattr}${tabbed ? ' role="tabpanel"' : ''}>
<div class="bar"><div class="lead"><h2 id="${sid}-h">${escH(s.title)}</h2><span class="count" data-n aria-live="polite"></span></div>${seg}</div>
${s.note ? `<p class="muted note">${escH(s.note)}</p>` : ''}<div class="filters">${filters}<button type="button" class="clear" data-clear hidden>Clear section filters</button></div>
<div data-view="table"${sp ? ' hidden' : ''}>${table}</div>${sp ? `\n<div data-view="schedule">${schedule}</div>` : ''}
<p class="empty muted" data-empty hidden>Nothing matches. Clear the search or filters.</p></section>` };
    });
    const gsel = globals.filter((gl) => gl.values.size).map((gl) => `<label class="flt">${escH(gl.label)}<select data-g="${escH(gl.id)}"><option value="">All</option>${[...gl.values.keys()].sort().map((v) => `<option value="${escH(v)}">${escH(v)}</option>`).join('')}</select></label>`).join('');
    const tabs = tabbed ? `<nav class="tabs" role="tablist" aria-label="Sections">${secs.map((s) => `<button type="button" role="tab" data-tab="${s.id}" aria-controls="${s.id}" aria-selected="false">${escH(s.title)}<span class="tc" data-tc="${s.id}">${s.n}</span></button>`).join('')}</nav>` : '<div class="tabs"></div>';
    return `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'">
<meta name="robots" content="noindex"><title>${escH(c.title)}</title><style>${REPORT_CSS}</style></head>
<body><header><div class="ttl"><h1>${escH(c.title)}</h1><div class="muted">${escH([c.subtitle, c.generated ? `Saved ${c.generated}` : '', c.version ? `Sohar Eligibility v${c.version}` : ''].filter(Boolean).join(' · '))}</div></div>
<label class="search">${SEARCH_SVG}<input type="search" id="q" data-q placeholder="Search patients: LASFIR, MRN, payer, provider, status" aria-label="Search every section"><kbd>/</kbd></label>
<button id="theme" type="button" aria-label="Light or dark">Light / dark</button></header>
<div class="tools">${tabs}<div class="gf">${gsel}<button type="button" class="clear" id="gclear" hidden>Clear search and filters</button></div></div>
<main>${secs.map((s) => s.html).join('\n')}</main>
<footer>LASFIR and MRN only. Contains PHI: keep it in MHS OneDrive or SharePoint, shared with invited people only. Never on a public site.</footer>
<script>${REPORT_JS}</script></body></html>
`;
  }
  /** "Sohar Results 2026-11-12.html" (local date). */
  const htmlReportFileName = (label, d) => `${str(label).replace(/[\\/:*?"<>|]+/g, ' ').trim()} ${C.localDate(d)}.html`;

  /** "Sohar Ready Import - Flow 2026-10-30.csv" (local date), like the other import files. */
  const flowImportFileName = (d) => `Sohar Ready Import - Flow ${C.localDate(d)}.csv`;

  // ---------- last updated (v0.20) ----------
  /**
   * The "Last updated" pill: each source's time, the newest overall, and how fresh the data is.
   * stamps: {report, sohar, flow} (ISO or readable dates; blank = never). Fresh when the report and the Sohar results
   * are both from the last day; stale after `staleDays` (default 7) for either; aging between.
   */
  function freshness(stamps, now, staleDays) {
    const t = (v) => { const n = C.soharTime(v); return isNaN(n) ? null : n; };
    const n = now ? now.getTime() : Date.now(), day = 86400000, limit = (staleDays == null ? 7 : staleDays) * day;
    const parts = [['report', 'NextGen report'], ['sohar', 'Sohar results'], ['flow', 'Sohar Flow']].map(([k, label]) => ({ key: k, label, time: t(stamps && stamps[k]) }));
    const known = parts.filter((p) => p.time != null);
    const newest = known.length ? Math.max(...known.map((p) => p.time)) : null;
    const core = parts.filter((p) => p.key !== 'flow');
    const ages = core.map((p) => (p.time == null ? Infinity : n - p.time));
    const level = !known.length ? 'none' : ages.some((a) => a > limit) ? 'stale' : ages.every((a) => a <= day) ? 'fresh' : 'aging';
    return { parts, newest, level };
  }

  // ---------- patient view (v0.22): plans, milestones, projection, action items ----------
  const money = (n) => `$${(Math.round(n * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const c2 = (n) => Math.round(n * 100) / 100;
  /** "Nov 4" from a date the report or Sohar wrote; '' when it isn't a date. */
  const shortDate = (v) => { const d = C.parseDate(v); return d ? d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' }) : str(v); };
  const payerOf = (r) => str(r['Sohar Payer']) || str(r['NextGen Payer']);
  const memberOf = (r) => str(r['Sohar Member ID']) || str(r['NextGen Member ID']);
  /**
   * A patient's results grouped by plan (payer + member ID), newest plan first. A result without a member ID joins the
   * newest plan with the same payer. The plan holding the lead result is current; the rest are previous plans.
   * Returns [{payer, memberId, results (newest first), current, superseded}].
   */
  function plansOf(results, lead) {
    const plans = [];
    const low = (s) => s.toLowerCase();
    for (const r of (results || []).slice().sort(newestFirst)) {
      const p = payerOf(r), m = memberOf(r);
      let g = plans.find((x) => low(x.payer) === low(p) && (!m || !x.memberId || x.memberId === m));
      if (!g) { g = { payer: p, memberId: m, results: [], current: false, superseded: false }; plans.push(g); }
      if (!g.memberId && m) g.memberId = m;
      g.results.push(r);
      if (/^Superseded/.test(str(r.Action))) g.superseded = true;
    }
    const cur = plans.find((g) => g.results.includes(lead)) || plans[0];
    if (cur) cur.current = true;
    return cur ? [cur, ...plans.filter((g) => g !== cur)] : plans;
  }
  /**
   * First-seen milestones, oldest result first: the first result each was seen on, keyed by the result row.
   * Deductible approaching / met, OOP max approaching / met, and a new plan (a plan seen after another, or one Sohar
   * found that NextGen doesn't have).
   */
  function milestones(results, plans) {
    const out = new Map(), seen = new Set();
    const add = (r, label) => { if (seen.has(label)) return; seen.add(label); out.set(r, [...(out.get(r) || []), label]); };
    const oldest = (results || []).slice().sort((a, b) => timeOf(a) - timeOf(b));
    const firstOf = new Map((plans || []).map((g) => [g, g.results[g.results.length - 1]]));
    const planOrder = [...firstOf.entries()].sort((a, b) => timeOf(a[1]) - timeOf(b[1]));
    for (const r of oldest) {
      if (r['Deductible Status'] === 'Approaching') add(r, 'Approaching deductible, first seen');
      if (r['Deductible Status'] === 'Met') add(r, 'Deductible met, first seen');
      if (r['OOP Status'] === 'Approaching') add(r, 'Approaching OOP max, first seen');
      if (r['OOP Status'] === 'Met') add(r, 'OOP max met, first seen');
    }
    planOrder.forEach(([g, r], i) => { if (i > 0 || str(r['Coverage Differs']) === 'Yes') { const l = out.get(r) || []; out.set(r, [...l, 'New plan, first seen']); } });
    return out;
  }
  /**
   * What changed between the lead result and the one before it on the same plan. {a, sub}: a short headline and a line.
   */
  function sinceLast(plans) {
    const cur = (plans || []).find((g) => g.current);
    if (!cur) return { a: 'No results', sub: '' };
    const [now, before] = cur.results;
    const prev = (plans || []).filter((g) => !g.current);
    if (!before) {
      return prev.length ? { a: 'New plan', sub: `Replaces ${prev[0].payer || 'an earlier plan'}${prev[0].memberId ? ` (${prev[0].memberId})` : ''}.` }
        : { a: 'First check', sub: 'No earlier Sohar result for this MRN.' };
    }
    const diff = [];
    const amt = (v) => { const n = num(v); return n == null ? '?' : money(n); };
    if (str(now['Sohar Status']) !== str(before['Sohar Status'])) diff.push(`Status ${str(before['Sohar Status']) || '?'} → ${str(now['Sohar Status']) || '?'}`);
    if (str(now['Deductible Remaining']) !== str(before['Deductible Remaining']) && (now['Deductible Remaining'] !== '' || before['Deductible Remaining'] !== '')) diff.push(`Deductible left ${amt(before['Deductible Remaining'])} → ${amt(now['Deductible Remaining'])}`);
    if (str(now['OOP Remaining']) !== str(before['OOP Remaining']) && (now['OOP Remaining'] !== '' || before['OOP Remaining'] !== '')) diff.push(`OOP max left ${amt(before['OOP Remaining'])} → ${amt(now['OOP Remaining'])}`);
    const since = shortDate(before['Sohar Date']);
    if (!diff.length) return { a: 'No change', sub: `Same as the ${str(before['Sent As']).toLowerCase() || 'check'}${since ? ` on ${since}` : ''}.` };
    const benefitsOnly = diff.every((d) => !d.startsWith('Status'));
    return { a: benefitsOnly ? 'Benefits only' : 'Status changed', sub: `${diff.join('; ')}${since ? ` since ${since}` : ''}.` };
  }
  /**
   * The benefits projection: walk the booked visits (Flow rows, soonest first) and apply each visit's estimate to the
   * deductible and the OOP max, capped at what's left. fin: the patient's Benefits view row. Visits with no estimate,
   * or a sandbox one, are left out and listed as missing. Copays stay in the projection: they count toward the OOP max.
   * Whether a visit counts toward the deductible: opts.dedApplies(flowRow) when it says (true / false, from the codes'
   * rules staff classified, v0.23); otherwise an estimate equal to the copay goes to the OOP max only and anything else
   * counts toward both. Null when Sohar returned neither amount.
   * Returns {ded, oop: {total, met, left, applied, at}, rows: [{flow, est, applied, dedLeft, oopLeft, marks, toward}], missing}.
   */
  /**
   * v0.32: the same shape as projection(), from visitPlan rows (every booked visit priced: Sohar's quote, then the PR
   * calculator). Rows past a plan-year reset are listed but not applied to this year's bars. Null without benefit figures.
   */
  function projectionFromPlan(fin, plan) {
    if (!fin || !plan || !plan.length) return null;
    const dT = num(fin.Deductible), dR = num(fin['Deductible Remaining']), oT = num(fin['OOP Max']), oR = num(fin['OOP Remaining']);
    if ((dT == null || dR == null) && (oT == null || oR == null)) return null;
    const ded = dT == null || dR == null ? null : { total: dT, met: c2(dT - dR), left: c2(Math.max(0, dR)), applied: 0, at: null };
    const oop = oT == null || oR == null ? null : { total: oT, met: c2(oT - oR), left: c2(Math.max(0, oR)), applied: 0, at: null };
    const rows = [], missing = [];
    let thisYear = true;
    for (const p of plan) {
      if (/New plan year/.test(str(p['Estimate Note']))) thisYear = false;
      const flow = { 'Appt Date': str(p['Appt Date']), Event: str(p.Event), Time: str(p.Time) };
      const est = num(p['Est Patient']), dl = num(p['Ded Left After']), ol = num(p['OOP Left After']);
      const base = { flow, est, source: str(p['Estimate Source']), note: str(p['Estimate Note']), nextYear: !thisYear, ruled: false };
      if (est == null) { missing.push(p); rows.push({ ...base, applied: null, dedLeft: null, oopLeft: null, marks: [] }); continue; }
      const marks = str(p.Milestone) ? str(p.Milestone).split(/\s*;\s*/) : [];
      if (thisYear) {
        if (ded && dl != null) { ded.applied = c2(Math.max(0, ded.left - dl)); if (!ded.at && marks.some((m) => /deductible met/i.test(m))) ded.at = flow; }
        if (oop && ol != null) { oop.applied = c2(Math.max(0, oop.left - ol)); if (!oop.at && marks.some((m) => /oop max met/i.test(m))) oop.at = flow; }
      }
      rows.push({ ...base, applied: est, dedLeft: dl, oopLeft: ol, marks });
    }
    return { ded, oop, rows, missing, plan: true };
  }
  function projection(fin, flowRows, opts) {
    const dedApplies = (opts && opts.dedApplies) || (() => null);
    if (!fin) return null;
    const dT = num(fin.Deductible), dR = num(fin['Deductible Remaining']), oT = num(fin['OOP Max']), oR = num(fin['OOP Remaining']);
    if ((dT == null || dR == null) && (oT == null || oR == null)) return null;
    const ded = dT == null || dR == null ? null : { total: dT, met: c2(dT - dR), left: c2(Math.max(0, dR)), applied: 0, at: null };
    const oop = oT == null || oR == null ? null : { total: oT, met: c2(oT - oR), left: c2(Math.max(0, oR)), applied: 0, at: null };
    const copay = num(fin.Copay);
    let dl = ded ? ded.left : 0, ol = oop ? oop.left : Infinity, apprMarked = !oop || ol <= oop.total * 0.2;
    const rows = [], missing = [];
    for (const f of flowRows || []) {
      const est = num(f['Sohar Est Patient']);
      if (est == null || isSandbox(f)) { missing.push(f); rows.push({ flow: f, est, applied: null, dedLeft: ded ? dl : null, oopLeft: oop ? ol : null, marks: [] }); continue; }
      const pay = c2(oop ? Math.min(est, ol) : est), marks = [];
      const fCopay = num(f.Copay);
      const isCopay = [fCopay, copay].some((c) => c != null && c > 0 && Math.abs(c - est) < 0.005);
      const rule = dedApplies(f);
      const toward = rule == null ? !isCopay : !!rule;
      if (ded && dl > 0 && toward) {
        const a = Math.min(pay, dl); dl = c2(dl - a); ded.applied = c2(ded.applied + a);
        if (dl === 0 && !ded.at) { ded.at = f; marks.push('Deductible met at this visit'); }
      }
      if (oop && ol > 0) {
        ol = c2(Math.max(0, ol - pay)); oop.applied = c2(oop.applied + pay);
        if (ol === 0 && !oop.at) { oop.at = f; marks.push('OOP max met at this visit'); }
        else if (!apprMarked && ol <= oop.total * 0.2) { apprMarked = true; marks.push('OOP max approaching'); }
      }
      rows.push({ flow: f, est, applied: pay, dedLeft: ded ? dl : null, oopLeft: oop ? ol : null, marks, toward, ruled: rule != null });
    }
    return { ded, oop, rows, missing };
  }
  const daysOld = (t, now) => Math.floor(((now ? now.getTime() : Date.now()) - t) / 86400000);
  /**
   * What to do next for one patient, most urgent first: [{tone: 'red'|'amber'|'grey', tag, text, due}].
   * ctx: {lead, fresh (freshness()), checks (guarantor checks), visits (report visits with items), flow (Flow rows),
   * tasks (open NextGen tasks), proj (projection()), now}.
   */
  function actionItems(ctx) {
    const o = ctx || {}, out = [], lead = o.lead || {}, now = o.now || new Date();
    const day0 = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    const soon = (d) => { const t = C.parseDate(d); return !t || t.getTime() - day0 <= 7 * 86400000; };
    const next = (o.visits || [])[0] || (o.flow || [])[0];
    const nextDate = next ? shortDate(next.date || next['Appt Date']) : '';
    if (o.fresh && o.fresh.level !== 'fresh' && o.fresh.level !== 'none') {
      const old = o.fresh.parts.filter((p) => p.key !== 'flow').sort((a, b) => (a.time ?? -Infinity) - (b.time ?? -Infinity))[0];
      const age = old && old.time != null ? `It is ${daysOld(old.time, now)} days old.` : 'It was never uploaded.';
      if (old) out.push(o.fresh.level === 'stale'
        ? { tone: 'red', tag: 'Fix now', text: old.key === 'report' ? `Upload a new NextGen report. ${age}` : `Re-run Sohar. The results are ${old.time != null ? `${daysOld(old.time, now)} days old` : 'missing'}.`, due: 'Today' }
        : { tone: 'amber', tag: 'Before visit', text: old.key === 'report' ? `Upload a new NextGen report. ${age}` : `Re-run Sohar. The results are ${old.time != null ? `${daysOld(old.time, now)} days old` : 'missing'}.`, due: nextDate ? `By ${nextDate}` : '' });
    }
    const checks = o.checks || [];
    if (checks.some((c) => c.severity === 'high')) out.push({ tone: 'red', tag: 'Fix now', text: 'Fix the guarantor in NextGen: the guarantor and the Sohar subscriber are different people.', due: nextDate ? `Before ${nextDate}` : '' });
    else if (checks.some((c) => c.severity === 'medium')) out.push({ tone: 'amber', tag: 'Before visit', text: 'Check the policy\'s insured in NextGen: Sohar\'s subscriber is not the patient.', due: nextDate ? `By ${nextDate}` : '' });
    if (str(lead['Sent As']) === 'Discovery' && /found/i.test(str(lead['Sohar Status']))) out.push({ tone: 'amber', tag: 'Before visit', text: `Run verification on ${payerOf(lead) || 'the plan discovery found'}.`, due: nextDate ? `By ${nextDate}` : '' });
    else if (str(lead['Work Queue']) === 'Action needed' && str(lead.Action)) out.push({ tone: 'amber', tag: 'Before visit', text: str(lead.Action), due: nextDate ? `By ${nextDate}` : '' });
    for (const v of o.visits || []) {
      const pc = (v.items || []).find((i) => i.id === 'precheck');
      if (!pc || pc.done) continue;
      const near = soon(v.date);
      out.push({ tone: near ? 'amber' : 'grey', tag: near ? 'Pre-check' : 'Later', text: `Pre-check the ${shortDate(v.date)} visit`, due: shortDate(v.date) });
    }
    for (const f of o.flow || []) {
      if (num(f['Sohar Est Patient']) != null || !/^complete/i.test(str(f['Run Status']))) continue;
      out.push({ tone: 'grey', tag: 'Verify', text: `Check deductible, OOP max and status for ${shortDate(f['Appt Date'])} ${str(f.Event)}`.trim(), due: 'No estimate' });
    }
    for (const t of o.tasks || []) out.push({ tone: soon(t['Due By']) ? 'amber' : 'grey', tag: 'NextGen task', text: [str(t.Task), str(t.Details)].filter(Boolean).join(': '), due: t['Due By'] ? `Due ${shortDate(t['Due By'])}` : '' });
    const rank = { red: 0, amber: 1, grey: 2 };
    return out.map((x, i) => ({ ...x, i })).sort((a, b) => rank[a.tone] - rank[b.tone] || a.i - b.i).map(({ i, ...x }) => x);
  }
  /**
   * The one plain line at the top of the Overview: {tone: 'red'|'amber'|'green'|'grey', text}.
   * ctx: {lead, checks, fresh, proj, next (the next visit's date)}.
   */
  function leadLine(ctx) {
    const o = ctx || {}, lead = o.lead || {}, by = o.next ? ` before the ${shortDate(o.next)} visit` : '';
    if ((o.checks || []).some((c) => c.severity === 'high')) return { tone: 'red', text: `Guarantor and subscriber are different people. Fix the guarantor in NextGen${by}.` };
    if (o.fresh && o.fresh.level === 'stale') {
      const old = o.fresh.parts.filter((p) => p.key !== 'flow').sort((a, b) => (a.time ?? -Infinity) - (b.time ?? -Infinity))[0];
      return { tone: 'red', text: `${old ? old.label : 'The data'} is out of date. Upload it again${by}.` };
    }
    if (str(lead['Sent As']) === 'Discovery' && /found/i.test(str(lead['Sohar Status']))) return { tone: 'amber', text: `Discovery found ${payerOf(lead) || 'a plan'}. Run verification to get benefits${by}.` };
    if (str(lead['Work Queue']) === 'Action needed') return { tone: 'amber', text: str(lead.Action) || `Check this patient${by}.` };
    const p = o.proj;
    if (p && p.oop && p.oop.left > 0 && p.oop.left <= p.oop.total * 0.2) {
      const i = p.oop.at ? p.rows.findIndex((r) => r.flow === p.oop.at) : -1;
      return { tone: 'amber', text: `Approaching OOP max: ${money(p.oop.left)} left.${p.oop.at ? ` Met at the ${shortDate(p.oop.at['Appt Date'])} visit if visits go as booked (visit ${i + 1} of ${p.rows.length}).` : ''}` };
    }
    if (p && p.oop && p.oop.left === 0) return { tone: 'green', text: 'OOP max met for this plan year.' };
    if (str(lead['Work Queue']) === 'Pending') return { tone: 'grey', text: 'Waiting on Sohar. Check back after the next upload.' };
    return { tone: 'green', text: 'Coverage looks good. Nothing to fix before the next visit.' };
  }

  // ---------- v0.30: billing tickets, balances and financial notes (the master as the central workbook) ----------
  // tblOpenTickets and tblPatientBalances are Power Query loads (docs/powerquery); the add-in reads them. Financial notes are
  // the add-in's own hidden table. Patients match on the MRN without leading zeros (Excel and the billing form drop them).
  const FINANCIAL_NOTE_COLS = ['MRN', 'LASFIR', 'Note', 'Category', 'Ticket Id', 'By', 'When'];
  const FINANCIAL_CATEGORIES = ['Balance', 'Copay', 'Ticket', 'Other'];
  const mrnOf = (v) => str(v).replace(/\.0+$/, '').replace(/^0+(?=\d)/, ''); // a numeric MRN cell reads back as 1234.0 sometimes
  const sameMrn = (a, b) => !!mrnOf(a) && mrnOf(a) === mrnOf(b);
  // Card numbers (12-19 digits, with spaces or dashes) never go in a note.
  const scrubCards = (s) => str(s).replace(/\b(?:\d[ -]?){11,18}\d\b/g, '[card number removed]');
  const dayOfAny = (v) => (v == null || v === '' ? null : C.parseDate(v));
  /** Business days from a to b (Mon-Fri; a's day not counted), negative when b is before a. Days are parseDate values. */
  function businessDays(a, b) {
    if (!a || !b) return null;
    const sign = b < a ? -1 : 1, [lo, hi] = sign > 0 ? [a, b] : [b, a];
    let n = 0;
    for (let t = lo.getTime() + 864e5; t <= hi.getTime(); t += 864e5) { const w = new Date(t).getUTCDay(); if (w !== 0 && w !== 6) n++; }
    return sign * n;
  }
  /**
   * The patient's open billing tickets (tblOpenTickets rows), oldest first, each with age (business days since Submitted)
   * and sla: 'Past due' (after Resolution Due), 'Due soon' (Resolution Due within 1 business day), 'Waiting on Arietis'
   * (Confirmation Due passed with no reply), else 'On time'. today: a Date.
   */
  function openTicketsOf(rows, mrn, today) {
    return ticketList((rows || []).filter((r) => r && sameMrn(r.MRN, mrn)), today);
  }
  /** Every open ticket row with its age and SLA (openTicketsOf without the patient filter), oldest first. */
  function ticketList(rows, today) {
    const day = C.parseDate(today instanceof Date ? today : new Date());
    return (rows || []).filter((r) => r && str(r['Ticket Id'])).map((r) => {
      const sub = dayOfAny(r.Submitted), res = dayOfAny(r['Resolution Due']), conf = dayOfAny(r['Confirmation Due']);
      const toRes = res ? businessDays(day, res) : null;
      const sla = res && day > res ? 'Past due' : toRes != null && toRes <= 1 ? 'Due soon' : conf && day > conf && !str(r['Last Reply At']) ? 'Waiting on Arietis' : 'On time';
      return { id: str(r['Ticket Id']), clinic: str(r.Clinic), status: str(r.Status), task: str(r['Task Type']), urgency: str(r.Urgency), amount: str(r.Amount),
        submitted: sub ? shortDate(sub) : '', resolutionDue: res ? shortDate(res) : '', confirmationDue: conf ? shortDate(conf) : '',
        age: sub ? businessDays(sub, day) : null, sla, escalation: str(r['Escalation Flag']), lastReply: str(r['Last Reply At']) };
    }).sort((a, b) => (b.age || 0) - (a.age || 0));
  }
  /** The patient's account balance (tblPatientBalances: MRN, Per Nbr, Balance, As Of): {balance, asOf} or null. */
  function balanceOf(rows, mrn, perNbr) {
    const r = (rows || []).find((x) => x && (sameMrn(x.MRN, mrn) || (!str(x.MRN) && str(perNbr) && str(x['Per Nbr']) === str(perNbr))));
    if (!r) return null;
    const n = Number(str(r.Balance).replace(/[$,\s]/g, ''));
    const asOf = dayOfAny(r['As Of']);
    return { balance: isFinite(n) ? n : null, asOf: asOf ? shortDate(asOf) : str(r['As Of']) };
  }
  /** The patient's financial notes, newest first. */
  function financialNotesOf(rows, mrn) {
    return (rows || []).filter((r) => r && sameMrn(r.MRN, mrn)).map((r) => ({ note: str(r.Note), category: str(r.Category), ticket: str(r['Ticket Id']), by: str(r.By), when: str(r.When) }))
      .sort((a, b) => (a.when < b.when ? 1 : -1));
  }
  /**
   * Add one financial note (kept with every earlier note). t: {mrn, lasfir, note, category, ticket, by, when}. Trimmed to
   * NOTE_MAX characters, card numbers taken out, the category one of FINANCIAL_CATEGORIES (else Other).
   */
  function addFinancialNote(saved, t) {
    needBy(t);
    if (!mrnOf(t.mrn)) throw new Error('This patient has no MRN, so the note has nowhere to go.');
    const note = scrubCards(str(t.note).replace(/\s+/g, ' ')).slice(0, NOTE_MAX);
    if (!note) throw new Error('The note is empty.');
    const category = FINANCIAL_CATEGORIES.includes(str(t.category)) ? str(t.category) : 'Other';
    return (saved || []).map((r) => Object.fromEntries(FINANCIAL_NOTE_COLS.map((c) => [c, str(r[c])])))
      .concat({ MRN: str(t.mrn), LASFIR: str(t.lasfir), Note: note, Category: category, 'Ticket Id': str(t.ticket), By: str(t.by), When: str(t.when) || new Date().toISOString() });
  }
  /**
   * The Power Query data's freshness (queryStatus: {at (ISO, the oldest query's refresh), errors, count}) -> {level, text}:
   * fresh within a day, aging up to 3 days, stale after that or on any query error; none when there are no queries.
   */
  function dataFreshness(st, now) {
    if (!st || !st.count) return { level: 'none', text: '' };
    if (st.errors && st.errors.length) return { level: 'stale', text: `A query failed (${st.errors[0]}). Press Data > Refresh All and check the source file.` };
    const t = Date.parse(st.at || '');
    if (isNaN(t)) return { level: 'stale', text: 'Never refreshed. Press Data > Refresh All.' };
    const days = ((now ? now.getTime() : Date.now()) - t) / 864e5;
    if (days <= 1) return { level: 'fresh', text: 'Up to date.' };
    const d = Math.floor(days);
    return { level: days <= 3 ? 'aging' : 'stale', text: `${d} day${d === 1 ? '' : 's'} old. Press Data > Refresh All.` };
  }
  /**
   * Dashboard counts for the open tickets: {total, byClinic, byStatus, byAge (0-2, 3-5, 6-10, 11+ business days), pastDue,
   * noMrn}. Rows are tblOpenTickets as Power Query loads them.
   */
  function ticketCounts(rows, today) {
    const day = C.parseDate(today instanceof Date ? today : new Date());
    const out = { total: 0, byClinic: {}, byStatus: {}, byAge: { '0-2': 0, '3-5': 0, '6-10': 0, '11+': 0 }, pastDue: 0, noMrn: 0 };
    for (const r of rows || []) {
      if (!r || !str(r['Ticket Id'])) continue;
      out.total++;
      const c = str(r.Clinic) || '(no clinic)', s = str(r.Status) || '(no status)';
      out.byClinic[c] = (out.byClinic[c] || 0) + 1; out.byStatus[s] = (out.byStatus[s] || 0) + 1;
      const sub = dayOfAny(r.Submitted), age = sub ? businessDays(sub, day) : null;
      if (age != null) out.byAge[age <= 2 ? '0-2' : age <= 5 ? '3-5' : age <= 10 ? '6-10' : '11+']++;
      const res = dayOfAny(r['Resolution Due']);
      if (res && day > res) out.pastDue++;
      if (!mrnOf(r.MRN)) out.noMrn++;
    }
    return out;
  }

  return { FLOW_COLS, FLOW_SHEET_COLS, isFlowFile, flowRecords, mergeFlow, flowFor, flowAmount, isSandbox, patientRecord, byPatient, upcomingOf, leadOf,
    FINANCIAL_NOTE_COLS, FINANCIAL_CATEGORIES, businessDays, dataFreshness, openTicketsOf, ticketList, balanceOf, financialNotesOf, addFinancialNote, ticketCounts, scrubCards,
    shortDate, plansOf, milestones, sinceLast, projection, projectionFromPlan, actionItems, leadLine,
    PROFILE_HEADERS, profiles, samePerson, guarantorCheck, subscriberOf,
    PRECHECK_HEADERS, PRECHECK_ITEMS, PRECHECK_COLS, PRECHECK_CONFIG_ROWS, NOTE_MAX, visitKey, adminNames, reportVisits, precheckItems, patientNotes, setPrecheck,
    addPatientNote, prunePrecheck, freshness, FLOW_IMPORT_COLS, FLOW_IMPORT_CONFIG_ROWS, flowImportRows, flowImportFileName,
    // v0.31
    PROVIDER_GROUPS, VERIFICATION_HISTORY_COLS, HISTORY_SOURCES, policyKey, historyFromResults, historyFromFlow, rotationIds, withSource, mergeHistory, policyTimeline, resetBasisFor, estimateDueFor, HTML_EXPORT_COLS, HTML_EXPORT_VIEWS, HTML_GLOBAL_FILTERS, ESTIMATE_RESULT_COLS, withEstimates, SOHAR_ESTIMATE, PR_CALC_ESTIMATE, estimateSourceOf, LEDGER_RESULT_COLS, LEDGER_LINK_COLS, ledgerResults, VISIT_PLAN_COLS, visitPlan, quotesByMrn, isSoharRunExport, htmlReport, htmlReportFileName, statusTone, soharLinkLabel, benefitMeter };
});
