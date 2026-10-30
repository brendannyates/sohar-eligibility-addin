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
    'Discovered Verification Link', 'NextGen Verification Link', 'Started', 'Last Updated', 'Note'];
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
  // full screen. Saved on the hidden Pre-Check sheet (PRECHECK_COLS: MRN key, the patient's LASFIR, the admin's name);
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
  /** One visit's notes, newest first: [{note, by, when}]. */
  function precheckNotes(visit, key, saved) {
    const vk = visitKey(key, visit.date, visit.time);
    return (saved || []).map((r, i) => ({ r, i })).filter(({ r }) => str(r.Key) === vk && str(r.Item) === 'note' && str(r.Note))
      .sort((a, b) => b.i - a.i).map(({ r }) => ({ note: str(r.Note), by: str(r.By), when: str(r.When) }));
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
  /** Add a note to a visit (kept with every earlier note). Trimmed, at most NOTE_MAX characters. */
  function addPrecheckNote(saved, t) {
    needBy(t);
    const note = str(t.note).replace(/\s+/g, ' ').slice(0, NOTE_MAX);
    if (!note) throw new Error('The note is empty.');
    return (saved || []).concat(rowOf(t, 'note', { Note: note }));
  }
  /** Drop ticks for visits before `from` (a day), so the sheet does not grow forever. */
  function prunePrecheck(saved, from) {
    const t0 = from ? C.parseDate(from instanceof Date ? from : String(from)) : null;
    if (!t0) return saved || [];
    return (saved || []).filter((r) => { const d = C.parseDate(r['Appt Date']); return !d || d >= t0; });
  }

  // ---------- Sohar Flow import (v0.21) ----------
  // The CSV template Sohar gives for the "MHS Eligibility - Discovery vs NextGen" Flow, exactly, in order. One row per
  // patient. The Flow's rateTable is an object and has no CSV column, so Flow runs from this file get discovery, the
  // NextGen policy check and network status, but no cost estimate.
  const FLOW_IMPORT_COLS = ['id', 'patient.firstName', 'patient.lastName', 'patient.dateOfBirth', 'patient.state', 'placeOfServiceCode',
    'specialtyCode', 'chartPayerId', 'chartMemberId', 'tin', 'providerGroupName'];
  const FLOW_IMPORT_CONFIG_ROWS = [
    { setting: 'TIN', value: '', what: 'Sohar Flow import: the practice tax ID (tin column). Blank: left blank on every row.' },
    { setting: 'Sohar provider group', value: '', what: 'Sohar Flow import: the provider group name the Flow checks network status for (providerGroupName). Blank: the overall network status.' },
  ];
  /**
   * Flow import rows from the built patients (core.buildPatients after the re-verify skip). In scope only. chartPayerId
   * and chartMemberId are the NextGen policy the Flow checks discovery against: filled only for patients routed to
   * verification (a confident payer), from their first insurance; discovery patients get blanks, so a guessed payer is
   * never verified as the chart policy. Left out: every insurance verified eligible inside the skip window (no repeat
   * Sohar calls), and rows the Flow would reject (no name, DOB or state).
   * Returns {rows, skipped: [{key, why}]}.
   */
  function flowImportRows(patients, cfg) {
    const c = cfg || {};
    const rows = [], skipped = [];
    for (const p of patients || []) {
      if (p['In Scope'] !== 'Yes') continue;
      const key = C.keyOf(p);
      const vrs = (p.x && p.x.verifyRows) || [];
      if (p.Route === 'verification' && vrs.length && vrs.every((v) => v.reverifySkip)) { skipped.push({ key, why: 'verified eligible recently (skip window)' }); continue; }
      if (!str(p['First Name']) || !str(p['Last Name']) || !str(p.DOB) || !str(p.State)) { skipped.push({ key, why: 'no name, DOB or state' }); continue; }
      const v = p.Route === 'verification' ? (vrs.find((x) => !x.reverifySkip) || { payerId: p['Sohar payerId'], memberId: p['Member ID'] }) : null;
      rows.push({ id: key, 'patient.firstName': str(p['First Name']), 'patient.lastName': str(p['Last Name']), 'patient.dateOfBirth': str(p.DOB),
        'patient.state': str(p.State), placeOfServiceCode: str(p.POS), specialtyCode: str(c.specialty) || 'PSY',
        chartPayerId: v ? str(v.payerId) : '', chartMemberId: v ? str(v.memberId) : '', tin: str(c.tin), providerGroupName: str(c.providerGroup) });
    }
    return { rows, skipped };
  }
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

  return { FLOW_COLS, FLOW_SHEET_COLS, isFlowFile, flowRecords, mergeFlow, flowFor, flowAmount, isSandbox, patientRecord, byPatient, upcomingOf, leadOf,
    PROFILE_HEADERS, profiles, samePerson, guarantorCheck, subscriberOf,
    PRECHECK_HEADERS, PRECHECK_ITEMS, PRECHECK_COLS, PRECHECK_CONFIG_ROWS, NOTE_MAX, visitKey, adminNames, reportVisits, precheckItems, precheckNotes, setPrecheck,
    addPrecheckNote, prunePrecheck, freshness, FLOW_IMPORT_COLS, FLOW_IMPORT_CONFIG_ROWS, flowImportRows, flowImportFileName };
});
