/* rules.js — pure clinic rules (R5-R12), no Office.js, no DOM. Loaded by index.html, dashboard.html and tests/.
   Needs core.js loaded first (R9 shares core.reverifyCheck with the build's re-verify skip). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.SoharRules = factory(root.SoharCore);
})(typeof self !== 'undefined' ? self : this, function (C) {
  'use strict';

  const REVERIFY_URL = 'https://app.soharhealth.com/create?type=verification';

  // Display order = flag sort order.
  const PILLS = [
    { code: 'MEDICAL', label: 'Medi-Cal', severity: 'high', color: 'red' },
    { code: 'MEDICAL_NEW', label: 'New Medi-Cal', severity: 'high', color: 'red' },
    { code: 'SFHP_PROVIDER', label: 'SFHP provider', severity: 'high', color: 'red' },
    { code: 'KAISER_MM_NOAUTH', label: 'Kaiser/VA MM no auth', severity: 'high', color: 'red' },
    { code: 'NP_TMS', label: 'NP-TMS', severity: 'high', color: 'red' },
    { code: 'REVERIFY', label: 'Re-verify', severity: 'medium', color: 'amber' },
    { code: 'QMB', label: 'QMB', severity: 'medium', color: 'amber' },
    { code: 'SECONDARY', label: 'Secondary may cover', severity: 'info', color: 'blue' },
    { code: 'VIP', label: 'VIP', severity: 'info', color: 'blue' },
    { code: 'DUAL', label: 'Two insurances sent', severity: 'info', color: 'blue' },
    // v0.12: Sohar found coverage when it searched with the guarantor as the person (discovery id <MRN>-G).
    { code: 'GUARANTOR', label: 'Via guarantor', severity: 'medium', color: 'blue' },
    { code: 'KAISER_MM', label: 'Kaiser/VA MM', severity: 'info', color: 'grey' },
  ];
  const PILL = {};
  PILLS.forEach((p, i) => { PILL[p.code] = Object.assign({ order: i }, p); });

  // ---------- small helpers ----------
  function str(v) { return v == null ? '' : String(v).trim(); }
  function low(v) { return str(v).toLowerCase(); }
  function words(v) { return ' ' + low(v).replace(/[^a-z0-9]+/g, ' ').trim() + ' '; }

  function parseDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    if (typeof v === 'number') return v > 59 ? new Date(Math.round((v - 25569) * 86400000)) : null;
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) { let y = +m[3]; if (y < 100) y += y < 50 ? 2000 : 1900; return new Date(Date.UTC(y, +m[1] - 1, +m[2])); }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return null;
  }
  function fmtDate(d) {
    if (!d) return '';
    return `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${String(d.getUTCDate()).padStart(2, '0')}/${d.getUTCFullYear()}`;
  }
  /** Local calendar day of a Date (or now) as a UTC-midnight Date, comparable with parseDate output. */
  function dayOf(d) {
    const x = d instanceof Date && !isNaN(d) ? d : new Date();
    return new Date(Date.UTC(x.getFullYear(), x.getMonth(), x.getDate()));
  }
  function localDate(d) {
    const x = d instanceof Date && !isNaN(d) ? d : new Date();
    return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`;
  }
  const daysBetween = (a, b) => Math.round((b - a) / 86400000);
  /** "9:00 AM" -> 540, "13:30" -> 810; blank or unreadable sorts last. Text order put "10:00 AM" before "9:00 AM". */
  function minutes(t) {
    const m = str(t).match(/^(\d{1,2}):(\d{2})\s*([AP]M)?/i);
    if (!m) return 9999;
    let h = +m[1] % (m[3] ? 12 : 24);
    if (m[3] && /pm/i.test(m[3])) h += 12;
    return h * 60 + +m[2];
  }

  /** LASFIR: first three letters of last name + first three of first name (copied from core.lasfir). */
  function lasfir(last, first) {
    const l = (s) => str(s).toUpperCase().replace(/[^A-Z]/g, '').slice(0, 3);
    return l(last) + l(first);
  }

  // ---------- classifiers (exported, unit-tested) ----------
  const isQMB = (t) => /\bqmb\b/i.test(str(t));
  const isMediCal = (t) => /medi[\s-]?cal\b|medicaid/i.test(str(t));
  const isKaiser = (t) => /kaiser|\bkp\b/i.test(str(t));
  // VA and TriWest (VA community care) share one bucket for rules; flag text uses the payer's own name.
  const isVA = (t) => /veterans?'?\s*affairs|\bva\b|tri\s*-?west/i.test(str(t));
  const isSFHP = (t) => /san\s+francisco\s+health\s+plan|\bsfhp\b/i.test(str(t));
  const isCarelon = (t) => /carelon|beacon/i.test(str(t)); // Beacon Health Options is Carelon's former name
  const isUMR = (t) => /\bumr\b/i.test(str(t));
  const isVIP = (t) => /\bvip\b/i.test(str(t));
  const isMedicare = (t) => /medicare|noridian|\bqmb\b/i.test(str(t));
  const isSelfPay = (t) => /self\s*-?\s*pay/i.test(str(t));
  /** Rendering like "Last, First PMHNP" -> true. MD/DO/PA/PhD -> false. */
  const isNP = (r) => /\b(np|pmhnp|fnp|aprn|dnp|anp|agnp|crnp|arnp)\b/i.test(str(r));
  /** TMS family: TMS, MT (Redo), MagV, Theta. "Telepsych TMS F/U" is TMS. */
  const isTMSEvent = (e) => /\btms\b|\bmt\b|mag\s*v|theta/i.test(str(e));
  const NOT_MM = /\btms\b|\bmt\b|mag\s*v|theta|esketamine|spravato|ketamine|\bivk\b|\bivn\b|meeting|training|on\s*call|peer\s*to\s*peer|adhd|testing|psychotherapy/i;
  /** Medication management: "Med Management ...", plain "Telepsych", "Tel/IP New Patient", "Transition Telepsych Patient". */
  const isMedMgmtEvent = (e) => {
    const s = str(e);
    if (!s || NOT_MM.test(s)) return false;
    return /med(ication)?\s*(management|mgmt)|telepsych|new\s+patient/i.test(s);
  };

  // ---------- inputs ----------
  function xOf(p) { return (p && p.x && typeof p.x === 'object') ? p.x : {}; }
  const INACTIVE = /cancel|no\s*show|resched|deleted/i;

  /** Upcoming appointments (date >= today), soonest first. Falls back to Next Appt / Next Event. */
  function upcoming(p, today) {
    const x = xOf(p);
    const t = dayOf(today);
    let list = Array.isArray(x.appointments) ? x.appointments.slice() : [];
    if (!list.length && p && str(p['Next Appt'])) {
      list = [{ date: p['Next Appt'], event: p['Next Event'], loc: x.location, rendering: x.rendering, authCd: x.authCd, authCdLstEnc: x.authCdLstEnc }];
    }
    return list
      .map((a) => Object.assign({}, a || {}, { _d: parseDate(a && a.date) }))
      .filter((a) => a._d && a._d >= t && !INACTIVE.test(str(a.status)))
      .sort((a, b) => a._d - b._d || minutes(a.time) - minutes(b.time));
  }
  function apptAuth(a, x) {
    const own = Object.prototype.hasOwnProperty;
    if (own.call(a, 'authCd') || own.call(a, 'authCdLstEnc')) return str(a.authCd) || str(a.authCdLstEnc);
    return str(x.authCd) || str(x.authCdLstEnc);
  }
  const apptInfo = (a, x) => ({ date: fmtDate(a._d), event: str(a.event), rendering: str(a.rendering) || str(x.rendering) });

  const isVerification = (r) => r && r['Sent As'] === 'Verification' && r['Sohar Status'] !== 'Not in Sohar file';
  const isEligible = (r) => /^complete\.eligible/i.test(str(r && r['Sohar Status']));
  /** Latest by Sohar Date; same day breaks on the full timestamp text (as core does). */
  function latest(rows) {
    let best = null, bd = null;
    for (const r of rows) {
      const d = parseDate(r['Sohar Date']);
      const later = d && (!bd || d > bd || (+d === +bd && str(r['Sohar Date']) > str(best['Sohar Date'])));
      if (!best || later) { best = r; bd = d; }
    }
    return best;
  }
  const isSuperseded = (r) => /^Superseded/.test(str(r && r.Action));
  /**
   * Latest eligible verification from the latest batch (within 3 days of the newest verification, as core's two-insurance
   * check), else latest verification, else latest row; null if none. Superseded rows are skipped, so an old eligible
   * result on a past policy never hides a newer ineligible one.
   */
  function bestResult(rows) {
    const rs = (rows || []).filter((r) => r && !isSuperseded(r));
    const vers = rs.filter(isVerification);
    const newest = latest(vers);
    const nd = newest && parseDate(newest['Sohar Date']);
    const batch = nd ? vers.filter((r) => { const d = parseDate(r['Sohar Date']); return !d || nd - d <= 3 * 86400000; }) : vers;
    return latest(batch.filter(isEligible)) || newest || latest(rs);
  }
  // Patient key (core.keyOf): the MRN, else PER-<Per Nbr>. Flags, results and clinic rows group on it.
  const keyOf = (r) => C.keyOf(r);
  function resultsFor(results, key) {
    if (!results) return [];
    if (results instanceof Map) return results.get(key) || [];
    if (Array.isArray(results)) return results.filter((r) => r && keyOf(r) === key);
    return results[key] || [];
  }

  // ---------- flags ----------
  function mk(code, detail, extra) {
    const p = PILL[code];
    return Object.assign({ code, label: p.label, severity: p.severity, detail }, extra || {});
  }

  /**
   * Flags for one patient. results = Sohar result rows for this patient (same key).
   * cfg = { today: Date, sfhpProviders: string[], reverifyDays: 7 }.
   */
  function patientFlags(patient, results, cfg) {
    const c = cfg || {};
    const p = patient || {};
    const x = xOf(p);
    const today = dayOf(c.today);
    const rows = (results || []).filter(Boolean);
    const up = upcoming(p, today);
    const primary = str(p['Primary Payer']);
    const secChart = str(p['Secondary Payer']) || str(x.secPayer);
    const secLst = str(x.secPayerLstEnc);
    const sec = secChart || secLst; // last-encounter secondary when the chart has none
    const fin = str(x.finClass), finL = str(x.finClassLstEnc);
    const guar = [primary, fin].filter(Boolean);
    const out = [];

    // R5 Kaiser / VA medication management needs an auth.
    if (guar.some((t) => isKaiser(t) || isVA(t))) {
      const who = guar.find((t) => isKaiser(t) || isVA(t)); // formal payer name, e.g. "Veterans Affairs Nor Cal", "TriWest"
      for (const a of up) {
        if (!isMedMgmtEvent(a.event)) continue;
        const auth = apptAuth(a, x);
        const appt = apptInfo(a, x);
        out.push(mk('KAISER_MM', `${who} med management ${appt.date} (${appt.event}): ${auth ? `auth ${auth}` : 'no auth on file'}`, { appt }));
        if (!auth) out.push(mk('KAISER_MM_NOAUTH', `${who} med management on ${appt.date} (${appt.event}) has no Auth Cd on the guarantor or last encounter`, { appt }));
      }
    }

    // R6 Medi-Cal: only Kaiser and SFHP Medi-Cal accepted. QMB is handled separately.
    const sources = [
      { where: 'NextGen payer', text: [primary, fin].filter(Boolean).join(' / ') },
      { where: 'secondary payer', text: secChart },
      { where: 'last encounter secondary', text: secLst && low(secLst) !== low(secChart) ? secLst : '' },
      { where: 'last encounter Fin Class', text: finL, lstEnc: true },
      ...rows.map((r) => ({ where: r['Sent As'] === 'Discovery' ? 'Sohar discovery' : 'Sohar verification', text: [str(r['Sohar Payer']), str(r.Plan)].filter(Boolean).join(' / ') })),
    ];
    const badMediCal = sources.filter((s) => s.text && isMediCal(s.text) && !isQMB(s.text) && !isKaiser(s.text) && !isSFHP(s.text));
    if (badMediCal.length) {
      // "Had it last visit" means a not-accepted Medi-Cal; Kaiser / SFHP Medi-Cal last visit still makes this new.
      const lastHad = finL && isMediCal(finL) && !isKaiser(finL) && !isSFHP(finL) && !isQMB(finL);
      const isNew = !!finL && !lastHad;
      const detail = `Medi-Cal is not accepted (only Kaiser or SFHP Medi-Cal): ${badMediCal.map((s) => `${s.where} "${s.text}"`).join('; ')}`;
      out.push(mk('MEDICAL', detail));
      if (isNew) out.push(mk('MEDICAL_NEW', isMediCal(finL)
        ? `Last visit was billed as "${finL}", an accepted Medi-Cal plan: patient looks newly on a Medi-Cal plan MHS does not accept`
        : `Last visit was billed as "${finL}", not Medi-Cal: patient looks newly on Medi-Cal`));
    }

    // R6 SFHP provider check.
    if (sources.some((s) => s.text && isSFHP(s.text)) && up.length) {
      const list = (c.sfhpProviders || []).map(words).filter((w) => w.trim());
      if (!list.length) {
        const appt = apptInfo(up[0], x);
        out.push(mk('SFHP_PROVIDER', `SFHP: confirm provider is approved (${appt.rendering || 'no rendering provider'} on ${appt.date}). Add approved names to Config "SFHP providers".`, { severity: 'medium', appt }));
      } else {
        for (const a of up) {
          const appt = apptInfo(a, x);
          const w = words(appt.rendering);
          if (list.some((e) => w.includes(e))) continue;
          out.push(mk('SFHP_PROVIDER', `SFHP patient booked with ${appt.rendering || 'no rendering provider'} on ${appt.date}, who is not on the SFHP providers list`, { appt }));
        }
      }
    }

    // R7 UMR / Carelon (Beacon): no NP rendering for TMS.
    if (guar.some((t) => isUMR(t) || isCarelon(t))) {
      const who = guar.some(isUMR) ? 'UMR' : 'Carelon';
      for (const a of up) {
        const appt = apptInfo(a, x);
        if (isTMSEvent(a.event) && isNP(appt.rendering)) {
          out.push(mk('NP_TMS', `${who} does not allow an NP rendering provider for TMS: ${appt.event} on ${appt.date} with ${appt.rendering}`, { appt }));
        }
      }
    }

    // R8 Secondary insurance and QMB.
    if (sec && !isSelfPay(sec) && low(sec) !== low(primary)) {
      const best = bestResult(rows);
      const copay = best ? str(best.Copay) : '';
      const coins = best ? str(best.Coinsurance) : '';
      const medicare = guar.some(isMedicare) || isMedicare(sec);
      const who = secChart ? `secondary ${sec}` : `the secondary billed at the last visit (${sec})`;
      let detail = medicare
        ? `Medicare/Noridian involved: ${who} likely covers the deductible, copay or coinsurance. Confirm with billing before planning or collecting a copay.`
        : `${who.charAt(0).toUpperCase()}${who.slice(1)} may cover the deductible, copay or coinsurance.`;
      if (copay || coins) detail += ` Plan shows${copay ? ` copay ${copay}` : ''}${copay && coins ? ',' : ''}${coins ? ` coinsurance ${coins}` : ''}.`;
      out.push(mk('SECONDARY', detail, { severity: medicare ? 'medium' : 'info' }));
    }
    const qmbAt = [primary, secChart, secLst, fin, finL, ...rows.map((r) => `${str(r['Sohar Payer'])} ${str(r.Plan)}`)].find(isQMB);
    // Medicare primary with Medi-Cal behind it (dual eligible) is treated like QMB; the Medi-Cal flag above still shows.
    const dualMediCal = guar.some((t) => isMedicare(t) && !isMediCal(t)) && [secChart, secLst].find((t) => t && isMediCal(t) && !isQMB(t));
    if (qmbAt) out.push(mk('QMB', 'QMB: Medicare cost-sharing generally cannot be billed to the patient — confirm with billing before collecting'));
    else if (dualMediCal) out.push(mk('QMB', `Medicare with Medi-Cal secondary (${dualMediCal}): treat like QMB — Medicare cost-sharing generally cannot be billed to the patient. Confirm with billing before collecting`));

    // R4 VIP pill.
    const vipText = [primary, fin, finL, sec].find(isVIP);
    if (x.vip || vipText) {
      const vr = (Array.isArray(x.verifyRows) ? x.verifyRows : []).filter((v) => v && !isVIP(v.nextgenPayer) && !isVIP(v.payerName));
      out.push(mk('VIP', vr.length
        ? `MHS-sponsored (${vipText || 'VIP'}). Verifying other insurance: ${vr.map((v) => str(v.nextgenPayer) || str(v.payerName)).join(', ')}`
        : sec && !isVIP(sec) && !isSelfPay(sec)
          ? `MHS-sponsored (${vipText || 'VIP'}). Secondary ${sec} has no policy number in the report: discovery`
          : `MHS-sponsored (${vipText || 'VIP'}). No other insurance on file: discovery only`));
    }

    // R3 Two insurances sent.
    const vrows = Array.isArray(x.verifyRows) ? x.verifyRows.filter(Boolean) : [];
    if (vrows.length >= 2) {
      out.push(mk('DUAL', `Sent both to verification: ${vrows.map((v) => `${str(v.nextgenPayer) || str(v.payerName)}${v.source ? ` (${v.source})` : ''}`).join(' and ')}. The Sohar results decide which to choose in NextGen.`));
    }

    // R9 Re-verify: the same due rule as the build's skip window (core.reverifyCheck), per insurance sent (verifyRows),
    // else from any verification. Due when missing, not eligible, Re-verify days or more old, coverage ending before the
    // visit, or a next visit in a new month within Re-verify days (Re-verify in new month: any month change).
    if (up.length) {
      const opts = { today, nextAppt: up[0]._d, days: c.reverifyDays != null ? +c.reverifyDays : 7, newMonth: !!c.reverifyNewMonth };
      const checks = vrows.length ? vrows.map((v) => C.reverifyCheck(rows, v, opts)) : [C.reverifyCheck(rows, null, opts)];
      const due = checks.find((k) => k.due);
      if (due) {
        const visit = fmtDate(up[0]._d);
        let detail;
        if (due.code === 'none') detail = `No Sohar verification on file${vrows.length > 1 ? ' for one of the insurances' : ''}. Submit before the ${visit} visit.`;
        else if (due.code === 'window') detail = `Last Sohar verification ${due.verified} (${due.daysSince} days ago). Resubmit before the ${visit} visit.`;
        else detail = `Last Sohar verification ${due.verified}: ${due.reason}. Resubmit before the ${visit} visit.`;
        out.push(mk('REVERIFY', detail, { link: REVERIFY_URL }));
      }
    }

    // Guarantor search found coverage: the policy is the guarantor's (the patient a dependent on it).
    const viaG = latest(rows.filter((r) => r['Searched As'] === 'Guarantor' && r['Sohar Status'] === 'Coverage found'));
    if (viaG) out.push(mk('GUARANTOR', `Coverage found through the guarantor: ${str(viaG['Sohar Payer']) || 'a payer'}${str(viaG['Sohar Date']) ? ` (${fmtDate(parseDate(viaG['Sohar Date']))})` : ''}. Add the guarantor's policy in NextGen with the patient as a dependent, then verify.`));

    return out.sort((a, b) => PILL[a.code].order - PILL[b.code].order || str(a.appt && a.appt.date).localeCompare(str(b.appt && b.appt.date)));
  }

  /** Map(patient key -> flags[]). results: array of rows (any patient), Map, or {key: rows}. */
  function allFlags(patients, results, cfg) {
    const byKey = groupResults(results);
    const m = new Map();
    for (const p of patients || []) {
      const k = keyOf(p);
      if (!k) continue;
      m.set(k, (m.get(k) || []).concat(patientFlags(p, byKey.get(k) || [], cfg)));
    }
    return m;
  }
  function groupResults(results) {
    if (results instanceof Map) return results;
    const m = new Map();
    if (Array.isArray(results)) {
      for (const r of results) {
        const k = keyOf(r);
        if (!k) continue;
        if (!m.has(k)) m.set(k, []);
        m.get(k).push(r);
      }
    } else if (results && typeof results === 'object') {
      for (const k of Object.keys(results)) m.set(k, results[k]);
    }
    return m;
  }

  // ---------- Review sheet rows (core.REVIEW_COLS) ----------
  const FIX = {
    MEDICAL: 'Medi-Cal (other than Kaiser or SFHP) is not accepted. Contact the patient about self-pay or other coverage before the visit.',
    MEDICAL_NEW: 'Coverage changed since the last visit. Confirm the new plan with the patient and update NextGen.',
    SFHP_PROVIDER: 'Move the visit to an SFHP-approved provider, or confirm approval and add the provider to Config "SFHP providers".',
    KAISER_MM_NOAUTH: 'Get the Kaiser, VA or TriWest authorization on file (Auth Cd) before the med management visit.',
    NP_TMS: 'Reassign the TMS session to an MD/DO rendering provider.',
    REVERIFY: `Resubmit in Sohar: ${REVERIFY_URL}`,
    QMB: 'Do not collect Medicare cost-sharing from the patient until billing confirms.',
    SECONDARY: 'Check the secondary before planning or collecting a copay, coinsurance or deductible.',
    VIP: 'MHS-sponsored treatment. Verify any other insurance; otherwise discovery only.',
    DUAL: 'Choose the insurance Sohar returned eligible in NextGen.',
    KAISER_MM: 'Confirm the authorization covers this visit.',
    GUARANTOR: 'Add the guarantor\'s policy in NextGen with the patient as a dependent, then send the patient for verification.',
  };
  const REVIEW_SEV = { high: 'High', medium: 'Medium', info: 'Low' };

  /** Review rows for every flag except the KAISER_MM drill-down marker. */
  function reviewRows(flagsMap, patients) {
    const order = new Map();
    (patients || []).forEach((p, i) => { const k = keyOf(p); if (k && !order.has(k)) order.set(k, i); });
    const out = [];
    const entries = flagsMap instanceof Map ? [...flagsMap.entries()] : Object.entries(flagsMap || {});
    for (const [per, flags] of entries) {
      for (const f of flags || []) {
        if (f.code === 'KAISER_MM') continue;
        out.push({ Severity: REVIEW_SEV[f.severity] || 'Low', Issue: f.label, MRN: per, Detail: f.detail || '', Fix: FIX[f.code] || '', _o: order.has(per) ? order.get(per) : 1e9 });
      }
    }
    const rank = { High: 0, Medium: 1, Low: 2 };
    out.sort((a, b) => rank[a.Severity] - rank[b.Severity] || a._o - b._o);
    return out.map(({ _o, ...r }) => r);
  }

  // ---------- collapsing per-appointment flags for display ----------
  // One flag per appointment stays on the Review sheets; cards and Review Items text show one line per group.
  const PER_APPT = {
    NP_TMS: { byProvider: true, lead: (d) => `${(d.match(/^(\w+) does not/) || [0, 'This payer'])[1]} does not allow an NP rendering provider for TMS` },
    SFHP_PROVIDER: { byProvider: true, lead: () => 'SFHP patient booked with a provider not on the SFHP providers list' },
    KAISER_MM_NOAUTH: { lead: (d) => `${(d.match(/^(.+?) med management/) || [0, 'Kaiser/VA'])[1]} med management with no Auth Cd on file` },
    KAISER_MM: { lead: (d) => `${(d.match(/^(.+?) med management/) || [0, 'Kaiser/VA'])[1]} med management` },
  };
  const shortDate = (s) => str(s).replace(/^(\d{1,2})\/(\d{1,2})\/\d{2,4}$/, '$1/$2');
  /**
   * items: [{code, label, severity, detail, appt: {date, event, rendering}}] (flags, or clinic review rows mapped to that shape).
   * Returns the same shape; repeated per-appointment flags of one code (and one provider, for NP-TMS / SFHP) become one
   * item: "UMR does not allow an NP rendering provider for TMS: 2 visits with Lark, Jordan PMHNP (10/16, 10/17)".
   */
  function collapseFlags(items) {
    const out = [];
    const groups = new Map();
    for (const f of items || []) {
      if (!f) continue;
      const spec = PER_APPT[f.code];
      const a = f.appt || {};
      if (!spec || !str(a.date) || (f.code === 'SFHP_PROVIDER' && f.severity === 'medium')) { out.push(f); continue; }
      const k = `${f.code}|${spec.byProvider ? low(a.rendering) : ''}`;
      if (!groups.has(k)) { const g = { first: f, spec, list: [] }; groups.set(k, g); out.push(g); }
      groups.get(k).list.push(f);
    }
    return out.map((g) => {
      if (!g.list) return g;
      if (g.list.length === 1) return g.first;
      const f = g.first;
      const provs = [...new Set(g.list.map((x) => str(x.appt.rendering)).filter(Boolean))];
      const dates = g.list.map((x) => shortDate(x.appt.date)).join(', ');
      return Object.assign({}, f, { detail: `${g.spec.lead(str(f.detail))}: ${g.list.length} visits${provs.length === 1 ? ` with ${provs[0]}` : ''} (${dates})`, count: g.list.length });
    });
  }

  // ---------- Clinic sheet (R11 / R12) ----------
  const CLINIC_COLS = ['Next Appt', 'Next Time', 'Next Event', 'Next Provider', 'Location', 'Per Nbr', 'MRN', 'Last Name', 'First Name', 'DOB',
    'Primary Payer', 'Choose in NextGen', 'Secondary Payer', 'Sohar Status', 'Sohar Payer', 'Plan', 'Verified Date', 'Days Since Verified',
    'Re-verify', 'Re-verify Link', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status',
    'OOP Max', 'OOP Remaining', 'OOP Status', 'Benefit Flag', 'Pills', 'Review Items', 'Notes', 'Upcoming Appts', 'Sohar Link', 'Run Date',
    // v0.12: the re-verify skip window (left off the verification file until Next Check) and the clinic's "Re-verify anyway".
    'Verification Skipped', 'Next Check', 'Re-verify Anyway', 'Found Via Guarantor'];
  const CLINIC_REVIEW_COLS = ['Appt Date', 'Severity', 'Pill', 'Issue', 'MRN', 'Last Name', 'First Name', 'Event', 'Rendering', 'Location',
    'Detail', 'Fix', 'Link', 'Run Date'];
  // Clinic file: the clinic's own "Re-verify anyway" ticks, brought back to the admin workbook (core.mergeClinicRequests).
  const CLINIC_REQUEST_COLS = ['MRN', 'Last Name', 'First Name', 'Next Appt', 'Re-verify Anyway', 'Requested'];
  const BENEFIT_KEYS = ['Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status', 'OOP Max', 'OOP Remaining', 'OOP Status', 'Benefit Flag'];

  /** Next visit for the clinic view: upcoming appointment, else the result row / patient fields. */
  function nextVisit(p, rows, runDay) {
    const x = xOf(p);
    const up = upcoming(p, runDay);
    if (up.length) {
      const a = up[0];
      return { d: a._d, date: fmtDate(a._d), time: str(a.time), event: str(a.event), provider: str(a.rendering) || str(x.rendering), loc: str(a.loc) || str(x.location), count: up.length };
    }
    const r = rows.find((y) => str(y['Next Appt'])) || {};
    const date = str(p['Next Appt']) || str(r['Next Appt']);
    const d = parseDate(date);
    if (!d || d < runDay) return null;
    return { d, date: fmtDate(d), time: str(r['Next Time']), event: str(p['Next Event']) || str(r['Next Event']), provider: str(r['Next Provider']) || str(x.rendering), loc: str(r['Next Location']) || str(x.location), count: 1 };
  }
  const sameLoc = (a, b) => low(a) === low(b);
  const allLocations = (loc) => !str(loc) || /^all$/i.test(str(loc));
  function inScope(p) { return !p || p['In Scope'] == null || p['In Scope'] === '' || p['In Scope'] === 'Yes'; }

  function clinicSet(location, patients, results, runDate) {
    const runDay = dayOf(runDate);
    const byPer = groupResults(results);
    const set = [];
    for (const p of patients || []) {
      if (!p || !inScope(p)) continue;
      const per = keyOf(p);
      const rows = (byPer.get(per) || []).filter(Boolean);
      const nv = nextVisit(p, rows, runDay);
      if (!nv) continue;
      if (!allLocations(location) && !sameLoc(nv.loc, location)) continue;
      set.push({ p, per, rows, nv });
    }
    set.sort((a, b) => a.nv.d - b.nv.d || minutes(a.nv.time) - minutes(b.nv.time) ||str(a.p['Last Name']).localeCompare(str(b.p['Last Name'])));
    return set;
  }

  /** One row per in-scope patient whose next visit is at `location` ('' or 'All' = every location), soonest first. */
  function clinicRows(location, patients, results, flagsMap, runDate) {
    const run = fmtDate(dayOf(runDate));
    const runDay = dayOf(runDate);
    const fm = flagsMap instanceof Map ? flagsMap : new Map(Object.entries(flagsMap || {}));
    return clinicSet(location, patients, results, runDate).map(({ p, per, rows, nv }) => {
      const best = bestResult(rows.filter((r) => r['Searched As'] !== 'Guarantor')) || bestResult(rows) || {};
      const x = xOf(p), skip = x.reverifySkip || x.reverifySkipped || null;
      const ver = rows.filter(isVerification).map((r) => parseDate(r['Sohar Date'])).filter(Boolean).sort((a, b) => b - a)[0] || null;
      const flags = fm.get(per) || [];
      const rv = flags.find((f) => f.code === 'REVERIFY');
      const pills = [];
      for (const f of flags) if (!pills.includes(f.label)) pills.push(f.label);
      const row = {
        'Next Appt': nv.date, 'Next Time': nv.time, 'Next Event': nv.event, 'Next Provider': nv.provider, 'Location': nv.loc,
        'Per Nbr': str(p['Per Nbr']), 'MRN': str(p.MRN), 'Last Name': str(p['Last Name']), 'First Name': str(p['First Name']), 'DOB': str(p.DOB),
        'Primary Payer': str(p['Primary Payer']), 'Choose in NextGen': str(best['Choose in NextGen']),
        'Secondary Payer': str(p['Secondary Payer']) || str(xOf(p).secPayer),
        'Sohar Status': str(best['Sohar Status']), 'Sohar Payer': str(best['Sohar Payer']), 'Plan': str(best.Plan),
        'Verified Date': fmtDate(ver), 'Days Since Verified': ver ? daysBetween(ver, runDay) : '',
        'Re-verify': rv ? 'Yes' : '', 'Re-verify Link': rv ? REVERIFY_URL : '',
        'Pills': pills.join('; '),
        'Review Items': collapseFlags(flags.filter((f) => f.severity !== 'info')).map((f) => f.detail).join(' | '),
        'Notes': flags.filter((f) => f.severity === 'info' && f.code !== 'KAISER_MM').map((f) => f.detail).concat(str(best.Action) ? [str(best.Action)] : []).join(' | '),
        'Upcoming Appts': nv.count, 'Sohar Link': str(best['Sohar Link']), 'Run Date': run,
        'Verification Skipped': skip ? `Verified eligible ${str(skip.verified)}` : '', 'Next Check': skip ? str(skip.due) : '',
        'Re-verify Anyway': x.reverifyRequested ? 'Yes' : '',
        'Found Via Guarantor': flags.some((f) => f.code === 'GUARANTOR') ? 'Yes' : '',
      };
      for (const k of BENEFIT_KEYS) row[k] = str(best[k]);
      const out = {};
      for (const k of CLINIC_COLS) out[k] = row[k] == null ? '' : row[k];
      Object.defineProperty(out, 'key', { value: per, enumerable: false }); // non-column patient key (MRN)
      return out;
    });
  }

  /** One row per flag for patients in clinicRows(location, ...). */
  function clinicReviewRows(location, patients, results, flagsMap, runDate) {
    const run = fmtDate(dayOf(runDate));
    const fm = flagsMap instanceof Map ? flagsMap : new Map(Object.entries(flagsMap || {}));
    const out = [];
    for (const { p, per, nv } of clinicSet(location, patients, results, runDate)) {
      for (const f of fm.get(per) || []) {
        const a = f.appt || {};
        out.push({
          'Appt Date': a.date || nv.date, 'Severity': REVIEW_SEV[f.severity] || 'Low', 'Pill': f.label, 'Issue': f.code,
          'MRN': per, 'Last Name': str(p['Last Name']), 'First Name': str(p['First Name']),
          'Event': a.event || nv.event, 'Rendering': a.rendering || nv.provider, 'Location': nv.loc,
          'Detail': f.detail || '', 'Fix': FIX[f.code] || '', 'Link': f.link || '', 'Run Date': run,
        });
      }
    }
    return out;
  }

  /**
   * Plain-text clinic summary. LASFIR + MRN only (no names or DOB; Per Nbr when there is no MRN). Grouped by visit date.
   * rows = clinicRows output. opts: { runDate: Date, includeClear: false }.
   */
  function clinicEmail(location, rows, opts) {
    const o = opts || {};
    const loc = allLocations(location) ? 'All locations' : str(location);
    const groups = new Map();
    let listed = 0;
    for (const r of rows || []) {
      const items = [];
      const pills = str(r.Pills).split(';').map((s) => s.trim()).filter((s) => s && s !== PILL.KAISER_MM.label);
      if (pills.length) items.push(pills.join(', '));
      if (r['OOP Status'] === 'Met') items.push('OOP max met');
      else if (r['OOP Status'] === 'Approaching') items.push(`OOP max approaching${str(r['OOP Remaining']) ? ` (${str(r['OOP Remaining'])} left)` : ''}`);
      if (r['Deductible Status'] === 'Met' && r['OOP Status'] !== 'Met') items.push('Deductible met');
      else if (r['Deductible Status'] === 'Approaching' && r['OOP Status'] !== 'Met') items.push(`Deductible approaching${str(r['Deductible Remaining']) ? ` (${str(r['Deductible Remaining'])} left)` : ''}`);
      if (!items.length && !o.includeClear) continue;
      const date = str(r['Next Appt']) || 'No date';
      if (!groups.has(date)) groups.set(date, []);
      const who = `${lasfir(r['Last Name'], r['First Name'])} (${C.idLabel(r)})`;
      groups.get(date).push(`- ${str(r['Next Time']) ? `${str(r['Next Time'])} ` : ''}${who}: ${items.length ? items.join('; ') : 'No action'}`);
      listed++;
    }
    const head = [`Sohar pre-check: ${loc}`, `Run ${fmtDate(dayOf(o.runDate))}. ${listed} of ${(rows || []).length} patients need attention.`];
    if (!listed) return head.concat(['', 'Nothing to flag.']).join('\n');
    const dates = [...groups.keys()].sort((a, b) => (parseDate(a) || 0) - (parseDate(b) || 0));
    const body = [];
    for (const d of dates) body.push('', d, ...groups.get(d));
    body.push('', `Re-verify in Sohar: ${REVERIFY_URL}`);
    return head.concat(body).join('\n');
  }

  // ---------- Payer evidence (R10) ----------
  const PAYER_EVIDENCE_COLS = ['Key', 'MRN', 'NextGen Payer', 'Fin Class', 'Member ID Prefix', 'Location', 'Guar Zip3', 'State',
    'Confidence', 'Sohar Payer', 'Sohar Plan', 'Sohar Status', 'First Seen', 'Last Seen'];
  const prefix3 = (id) => str(id).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 3);
  const evidenceKey = (per, payer, pre) => `${str(per)}|${low(payer)}|${pre}`;

  /**
   * Upsert one row per patient-insurance observation. Keeps First Seen and earlier Sohar fields when the new ones are blank.
   * perToKey (optional, core.perKeyMap): rows saved before v0.12 (Per Nbr) are re-keyed to the patient's MRN.
   */
  function payerEvidence(patients, results, previousRows, today, perToKey) {
    const stamp = localDate(today);
    const byPer = groupResults(results);
    const map = new Map();
    for (const r0 of previousRows || []) {
      if (!r0 || !str(r0.Key)) continue;
      const r = Object.assign({}, r0);
      if (!str(r.MRN) && str(r['Per Nbr'])) {
        // The old Key is "<Per Nbr>|payer|prefix": only its first part changes.
        const per = str(r['Per Nbr']);
        r.MRN = C.migrateKey(C.perKey(per), perToKey);
        r.Key = str(r.Key).startsWith(`${per}|`) ? `${r.MRN}${str(r.Key).slice(per.length)}` : evidenceKey(r.MRN, r['NextGen Payer'], str(r['Member ID Prefix']));
        delete r['Per Nbr'];
      }
      const k = str(r.Key);
      if (!map.has(k) || str(map.get(k)['Last Seen']) < str(r['Last Seen'])) map.set(k, r);
    }
    for (const p of patients || []) {
      if (!p) continue;
      const x = xOf(p);
      const per = keyOf(p);
      if (!per) continue;
      const obs = (Array.isArray(x.verifyRows) && x.verifyRows.length ? x.verifyRows.filter(Boolean) : [{ nextgenPayer: p['Primary Payer'], memberId: p['Member ID'] }])
        .filter((v) => str(v.nextgenPayer) && !isSelfPay(v.nextgenPayer));
      const rows = (byPer.get(per) || []).filter(Boolean);
      const firstAppt = upcoming(p, new Date(0))[0];
      for (const v of obs) {
        const payer = str(v.nextgenPayer);
        const pre = prefix3(v.memberId);
        const key = evidenceKey(per, payer, pre);
        const lstEnc = /lst|last|enc/i.test(str(v.source));
        const match = rows.filter((r) => low(r['NextGen Payer']) === low(payer) || (pre && prefix3(r['NextGen Member ID']) === pre));
        const res = bestResult(match);
        const prev = map.get(key) || {};
        const keep = (k, val) => (str(val) ? str(val) : str(prev[k]));
        map.set(key, {
          'Key': key, 'MRN': per, 'NextGen Payer': payer,
          'Fin Class': keep('Fin Class', lstEnc ? x.finClassLstEnc : x.finClass), 'Member ID Prefix': pre,
          'Location': keep('Location', x.location || (firstAppt && firstAppt.loc)),
          'Guar Zip3': keep('Guar Zip3', str(x.guarZip || p.Zip).replace(/\D/g, '').slice(0, 3)),
          'State': keep('State', p.State),
          'Confidence': keep('Confidence', v.confidence || x.payerConfidence || p.Confidence),
          'Sohar Payer': keep('Sohar Payer', res && res['Sohar Payer']),
          'Sohar Plan': keep('Sohar Plan', res && res.Plan),
          'Sohar Status': keep('Sohar Status', res && res['Sohar Status']),
          'First Seen': str(prev['First Seen']) || stamp, 'Last Seen': stamp,
        });
      }
    }
    return [...map.values()]
      .map((r) => { const o = {}; for (const k of PAYER_EVIDENCE_COLS) o[k] = r[k] == null ? '' : r[k]; return o; })
      .sort((a, b) => low(a['NextGen Payer']).localeCompare(low(b['NextGen Payer'])) || a.Key.localeCompare(b.Key));
  }

  /** Pivot: [{nextgenPayer, prefix, location, count, soharPayers: string[] (most common first)}], biggest first. */
  function evidenceSummary(rows) {
    const g = new Map();
    for (const r of rows || []) {
      if (!r) continue;
      const k = `${low(r['NextGen Payer'])}|${str(r['Member ID Prefix'])}|${low(r.Location)}`;
      if (!g.has(k)) g.set(k, { nextgenPayer: str(r['NextGen Payer']), prefix: str(r['Member ID Prefix']), location: str(r.Location), count: 0, _sp: new Map() });
      const e = g.get(k);
      e.count++;
      const sp = str(r['Sohar Payer']);
      if (sp) e._sp.set(sp, (e._sp.get(sp) || 0) + 1);
    }
    return [...g.values()]
      .map(({ _sp, ...e }) => Object.assign(e, { soharPayers: [..._sp.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).map((x) => x[0]) }))
      .sort((a, b) => b.count - a.count || a.nextgenPayer.localeCompare(b.nextgenPayer) || a.prefix.localeCompare(b.prefix) || a.location.localeCompare(b.location));
  }

  return {
    REVERIFY_URL, PILLS, CLINIC_COLS, CLINIC_REVIEW_COLS, CLINIC_REQUEST_COLS, PAYER_EVIDENCE_COLS,
    isMediCal, isKaiser, isVA, isSFHP, isQMB, isCarelon, isUMR, isNP, isTMSEvent, isMedMgmtEvent, isVIP, isMedicare,
    patientFlags, allFlags, reviewRows, clinicRows, clinicReviewRows, clinicEmail, payerEvidence, evidenceSummary,
    bestResult, lasfir, collapseFlags,
  };
});
