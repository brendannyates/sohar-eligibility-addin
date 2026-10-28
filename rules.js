/* rules.js — pure clinic rules (R5-R12), no Office.js, no DOM. Loaded by index.html, dashboard.html and tests/.
   Needs core.js loaded first (R9 shares core.reverifyCheck with the build's re-verify skip), and payer-rules.js for the
   esketamine payer credential rule (R13; without it that rule is silent). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'), require('./payer-rules.js'));
  else root.SoharRules = factory(root.SoharCore, root.PayerRules);
})(typeof self !== 'undefined' ? self : this, function (C, PR) {
  'use strict';

  const REVERIFY_URL = 'https://app.soharhealth.com/create?type=verification';

  // Display order = flag sort order.
  const PILLS = [
    { code: 'MEDICAL', label: 'Medi-Cal', severity: 'high', color: 'red' },
    { code: 'MEDICAL_NEW', label: 'New Medi-Cal', severity: 'high', color: 'red' },
    { code: 'SFHP_PROVIDER', label: 'SFHP provider', severity: 'high', color: 'red' },
    { code: 'KAISER_MM_NOAUTH', label: 'Kaiser/VA MM no auth', severity: 'high', color: 'red' },
    { code: 'NP_TMS', label: 'NP-TMS', severity: 'high', color: 'red' },
    // R14 (v0.16): an intervention's authorization belongs to the insurance it was obtained under.
    { code: 'INS_CHANGED', label: 'Insurance changed?', severity: 'high', color: 'red' },
    // R13 esketamine payer credential rules (Payer Credential Rules sheet). Besse is a flag only: nothing is re-routed.
    { code: 'ESK_CRED', label: 'ESK credential', severity: 'high', color: 'red' },
    { code: 'BESSE_NO', label: 'Not Besse-approved', severity: 'medium', color: 'amber' },
    { code: 'REVERIFY', label: 'Re-verify', severity: 'medium', color: 'amber' },
    { code: 'QMB', label: 'QMB', severity: 'medium', color: 'amber' },
    { code: 'SECONDARY', label: 'Secondary may cover', severity: 'info', color: 'blue' },
    { code: 'VIP', label: 'VIP', severity: 'info', color: 'blue' },
    { code: 'DUAL', label: 'Two insurances sent', severity: 'info', color: 'blue' },
    // v0.12: Sohar found coverage when it searched with the guarantor as the person (discovery id <MRN>-G).
    { code: 'GUARANTOR', label: 'Via guarantor', severity: 'medium', color: 'blue' },
    { code: 'ESK_STATE', label: 'ESK payer state?', severity: 'info', color: 'grey' },
    { code: 'BESSE_PENDING', label: 'Besse pending', severity: 'info', color: 'blue' },
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
  /**
   * Local calendar day of a Date (or now) as a UTC-midnight Date, comparable with parseDate output. A Date already at UTC
   * midnight is a day value (parseDate, or dayOf itself) and is kept, as core.dayOf does: converting it again west of UTC
   * gave the day before, so yesterday's visit still counted as upcoming.
   */
  function dayOf(d) {
    const x = d instanceof Date && !isNaN(d) ? d : new Date();
    if (x.getUTCHours() === 0 && x.getUTCMinutes() === 0 && x.getUTCSeconds() === 0 && x.getUTCMilliseconds() === 0) return new Date(x.getTime());
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
  /** Esketamine: "Esketamine Treatment", "Spravato First Dose", "ESK Obs". Not IV ketamine. */
  const isESKEvent = (e) => /esketamine|spravato|\besk\b/i.test(str(e));
  /** IV ketamine: "IV Ketamine Follow Up", "IVK". Not esketamine. */
  const isIVKEvent = (e) => /\biv\s*ketamine\b|\bivk\b/i.test(str(e));
  /** Interventions that run on an authorization: TMS, esketamine, IV ketamine. Consults come before the auth. */
  const isInterventionEvent = (e) => (isTMSEvent(e) || isESKEvent(e) || isIVKEvent(e)) && !/consult/i.test(str(e));
  const NOT_MM =/\btms\b|\bmt\b|mag\s*v|theta|esketamine|spravato|ketamine|\bivk\b|\bivn\b|meeting|training|on\s*call|peer\s*to\s*peer|adhd|testing|psychotherapy/i;
  /** Medication management: "Med Management ...", plain "Telepsych", "Tel/IP New Patient", "Transition Telepsych Patient". */
  const isMedMgmtEvent = (e) => {
    const s = str(e);
    if (!s || NOT_MM.test(s)) return false;
    return /med(ication)?\s*(management|mgmt)|telepsych|new\s+patient/i.test(s);
  };

  // ---------- inputs ----------
  function xOf(p) { return (p && p.x && typeof p.x === 'object') ? p.x : {}; }
  const INACTIVE = /cancel|no\s*show|resched|deleted/i;

  /** Upcoming appointments (date >= day, a dayOf / parseDate day value), soonest first. Falls back to Next Appt / Next Event. */
  function upcoming(p, day) {
    const x = xOf(p);
    let list = Array.isArray(x.appointments) ? x.appointments.slice() : [];
    if (!list.length && p && str(p['Next Appt'])) {
      list = [{ date: p['Next Appt'], event: p['Next Event'], loc: x.location, rendering: x.rendering, authCd: x.authCd, authCdLstEnc: x.authCdLstEnc }];
    }
    return list
      .map((a) => Object.assign({}, a || {}, { _d: parseDate(a && a.date) }))
      .filter((a) => a._d && a._d >= day && !INACTIVE.test(str(a.status)))
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

    // R13 Esketamine payer credential rules (cfg.payerRules, read from the Payer Credential Rules sheet; none = silent).
    // The payer (primary, else Fin Class, as R7) is matched with a state hint from the patient, else the visit location.
    // Credential from the rendering name; the ESK step from the event. Besse is a flag only: routing does not change.
    const credRules = Array.isArray(c.payerRules) ? c.payerRules : [];
    const esk = credRules.length && PR ? up.filter((a) => isESKEvent(a.event)) : [];
    if (esk.length) {
      const locState = {};
      for (const [k, v] of Object.entries(c.locationState || {})) locState[low(k)] = str(v).toUpperCase();
      let first = null, ambiguous = null;
      for (const a of esk) {
        const appt = apptInfo(a, x);
        const loc = str(a.loc) || str(x.location);
        const hint = str(p.State) || locState[low(loc)] || loc;
        const payer = guar.find((t) => PR.match(credRules, t, hint));
        const m = payer ? PR.match(credRules, payer, hint) : null;
        if (!m) continue;
        if (!first) first = { payer, m, appt };
        // isNP first: credOf reads a "Do" surname as a DO ("Do, Minh NP" -> MD).
        const r = PR.check(m, isNP(appt.rendering) ? 'NP' : PR.credOf(appt.rendering), PR.step(a.event, /first/i.test(str(a.event))));
        if (r.ok === false) out.push(mk('ESK_CRED', `${r.why}: ${appt.event} on ${appt.date} with ${appt.rendering}`, { appt, group: r.why }));
        else if (r.ok === null && m.ambiguous.length && !ambiguous) ambiguous = { payer, m, appt };
      }
      if (first) {
        const b = PR.besse(first.m), when = `${first.appt.event} ${first.appt.date}`;
        // Besse unknown because the state rules disagree: the same "which state" review item.
        if (b === 'unknown' && first.m.ambiguous.length && !ambiguous) ambiguous = first;
        if (ambiguous) out.push(mk('ESK_STATE', `${ambiguous.payer} matches ${ambiguous.m.ambiguous.join(' and ')} on Payer Credential Rules: check which state the plan is from (${ambiguous.appt.event} ${ambiguous.appt.date} with ${ambiguous.appt.rendering || 'no rendering provider'})`, { appt: ambiguous.appt }));
        if (b === 'not approved') out.push(mk('BESSE_NO', `Payer isn't Besse-approved: ${first.m.payer} (${when})`, { appt: first.appt }));
        else if (b === 'pending') out.push(mk('BESSE_PENDING', `Besse approval pending for ${first.m.payer} (${when})`, { appt: first.appt }));
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

    // R14 Interventions (TMS, esketamine, IV ketamine; not consults): when Sohar disagrees with the chart (not eligible,
    // coverage ending before the visit, or a discovery that found different coverage), the authorization on file may be
    // for insurance the patient no longer has. Flag only.
    const iv = up.find((a) => isInterventionEvent(a.event));
    if (iv) {
      const why = [];
      const v = latest(rows.filter((r) => isVerification(r) && !isSuperseded(r)));
      if (v && /^complete\.ineligible/i.test(str(v['Sohar Status']))) why.push(`Sohar says not eligible${str(v['Sohar Payer']) ? ` with ${str(v['Sohar Payer'])}` : ''}`);
      const end = v && isEligible(v) ? parseDate(v['Coverage End']) : null;
      if (end && end < iv._d) why.push(`coverage ends ${fmtDate(end)}, before the visit`);
      const diff = latest(rows.filter((r) => !isSuperseded(r) && str(r['Coverage Differs']) === 'Yes'));
      if (diff) why.push(`Sohar found ${str(diff['Sohar Payer']) || 'other coverage'}; NextGen has ${str(diff['NextGen Payer']) || primary || 'another payer'}`);
      if (why.length) {
        const appt = apptInfo(iv, x);
        out.push(mk('INS_CHANGED', `${appt.event} ${appt.date}: ${why.join('; ')}. The authorization on file may be for the old insurance.`, { appt }));
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
    INS_CHANGED: 'Confirm the patient\'s current insurance. If it changed, update NextGen and get a new authorization from the new payer before the visit.',
    REVERIFY: `Resubmit in Sohar: ${REVERIFY_URL}`,
    QMB: 'Do not collect Medicare cost-sharing from the patient until billing confirms.',
    SECONDARY: 'Check the secondary before planning or collecting a copay, coinsurance or deductible.',
    VIP: 'MHS-sponsored treatment. Verify any other insurance; otherwise discovery only.',
    DUAL: 'Choose the insurance Sohar returned eligible in NextGen.',
    KAISER_MM: 'Confirm the authorization covers this visit.',
    GUARANTOR: 'Add the guarantor\'s policy in NextGen with the patient as a dependent, then send the patient for verification.',
    ESK_CRED: 'Reassign to a provider the payer allows, or confirm with billing.',
    ESK_STATE: 'Check which state the plan is from, then confirm on Payer Credential Rules that the payer allows this provider.',
    BESSE_NO: 'Confirm with billing before using Besse (buy-and-bill) stock for this patient.',
    BESSE_PENDING: 'Besse approval for this payer is pending. Confirm with billing before using Besse stock.',
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
    // One group per provider and reason (an NP may be blocked for the first treatment but not for observation).
    ESK_CRED: { byProvider: true, lead: (d, f) => str(f.group) || 'This payer does not allow this provider to bill esketamine' },
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
      const k = `${f.code}|${spec.byProvider ? low(a.rendering) : ''}|${str(f.group)}`;
      if (!groups.has(k)) { const g = { first: f, spec, list: [] }; groups.set(k, g); out.push(g); }
      groups.get(k).list.push(f);
    }
    return out.map((g) => {
      if (!g.list) return g;
      if (g.list.length === 1) return g.first;
      const f = g.first;
      const provs = [...new Set(g.list.map((x) => str(x.appt.rendering)).filter(Boolean))];
      const dates = g.list.map((x) => shortDate(x.appt.date)).join(', ');
      return Object.assign({}, f, { detail: `${g.spec.lead(str(f.detail), f)}: ${g.list.length} visits${provs.length === 1 ? ` with ${provs[0]}` : ''} (${dates})`, count: g.list.length });
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

  // ---------- Payer Credential Rules sheet (R13) ----------
  // The sheet is imported from the shared CSV (never committed) and staff may edit it in Excel.
  /** Table rows read from the sheet (objects by header) -> rules for cfg.payerRules. No sheet / no module = []. */
  function payerRulesFrom(tableRows) {
    const rows = (tableRows || []).filter(Boolean);
    if (!PR || !rows.length) return [];
    const head = Object.keys(rows[0]);
    const res = PR.read([head].concat(rows.map((r) => head.map((h) => (r[h] == null ? '' : r[h])))));
    return res.rules || [];
  }
  /** Rules -> sheet rows (objects keyed by PayerRules.HEADERS) for writeAddinTable. */
  function payerRuleRows(rules) {
    if (!PR) return [];
    const [head, ...rows] = PR.toRows(rules || []);
    return rows.map((v) => Object.fromEntries(head.map((h, i) => [h, v[i] == null ? '' : v[i]])));
  }

  // ---------- Clinic Eligibility file (v0.14): central Sohar -> the clinic's Clinic Template Tool (CTT) ----------
  // The file carries data only: the clinic's patients, Sohar results, build review, payer learnings and Done marks, minus
  // patient and guarantor names and DOBs (LASFIR added). CTT re-derives rows, flags, benefits and tasks with this same
  // rules.js and core.js (byte-identical copies) and sends the clinic's edits back on the Clinic Work sheet.
  const CTT_EXTENSION = { id: 'cf213823-6d83-43e7-b7c9-37666b4d7c3d', version: '1.0.0.1' }; // CTT's installed Id / version
  const SE_KIND = 'Clinic Eligibility', SE_FORMAT = '1';
  const SE_SHEETS = { readMe: 'Read Me', work: 'Clinic Work', info: 'SE Info', patients: 'SE Patients', results: 'SE Results', review: 'SE Review',
    payers: 'SE Payers', done: 'SE Done', rules: 'Payer Credential Rules' };
  const CLINIC_WORK_COLS = ['Kind', 'Key', 'MRN', 'LASFIR', 'Value', 'Note', 'By', 'When'];
  const WORK_KINDS = { reverify: 'Re-verify anyway', done: 'Task done', note: 'Note', payer: 'Payer note' };
  // Admin-side tables filled from returned files (newest per MRN or payer + Location wins).
  const CLINIC_NOTE_COLS = ['MRN', 'Location', 'Note', 'By', 'When'];
  const CLINIC_PAYER_NOTE_COLS = ['NextGen Payer', 'Location', 'Card Says', 'Note', 'By', 'When'];
  // Clinic Work rows already brought back (newest When per Location + Kind + Key), so a file sent again changes nothing.
  // Kept 120 days by the row's When; an older file sent again is applied again. Value: the Task done tick (Yes/No) last acted
  // on, so a later note-only edit (which carries the old tick along) never redoes a Done the admin has changed since.
  const CLINIC_APPLIED_COLS = ['Location', 'Kind', 'Key', 'When', 'Value'];
  const APPLIED_DAYS = 120;
  // Done set from a returned file. Marked, so a later clinic No clears only the clinic's own Done, never the admin's "Yes".
  const clinicDoneMark = (loc) => `Yes (clinic${str(loc) ? ` ${str(loc)}` : ''})`;
  const isClinicDone = (v) => /^yes \(clinic\b[^)]*\)$/i.test(str(v));
  // Tasks the clinic can do with the patient (the rest stay with the admin).
  const CLINIC_TASK_TYPES = ['Add insurance in NextGen', 'Update insurance in NextGen', 'Fix member ID in NextGen', 'Fix name in NextGen',
    'Update benefits in NextGen', 'Confirm self-pay or get card'];
  const CELL_MAX = 32000; // Excel holds 32,767 characters a cell, and xlsx.js cuts a longer one silently
  // Never in the file. Subscriber First / Last / DOB are the guarantor when "Use guarantor as subscriber" is on.
  const PHI_COLS = ['Last Name', 'First Name', 'DOB', 'Zip', 'Subscriber First', 'Subscriber Last', 'Subscriber DOB'];
  const SE_DEFAULTS = { sfhpProviders: [], reverifyDays: 7, reverifyNewMonth: false, approachingPct: 0.2, locationState: {}, defaultState: '' };
  const isOn = (v) => /^(on|yes|y|true|1|x)$/i.test(str(v));
  const workKind = (k) => Object.keys(WORK_KINDS).find((n) => low(WORK_KINDS[n]) === low(k)) || '';
  const workId = (kind, key) => `${kind}|${kind === 'payer' ? low(key) : str(key)}`;
  const whenAt = (v) => Date.parse(str(v)); // a Clinic Work When (ISO) as ms; NaN when blank or unreadable
  /** NextGen Tasks Clinic Note from a Task done row: "called patient, new card scanned (Front desk, San Rafael, 10/23)". */
  function doneNote(r, loc) {
    const t = new Date(whenAt(r.When));
    const md = isNaN(t) ? '' : `${t.getMonth() + 1}/${t.getDate()}`;
    const tail = [str(r.By), str(loc), md].filter(Boolean).join(', ');
    return `${str(r.Note)}${tail ? ` (${tail})` : ''}`;
  }
  // The master workbook's Clinic Work Inbox: the Power Automate flow appends each clinic file's new Clinic Work rows with
  // the file's Location, name and the time it copied them. Flow-owned: Sohar creates it once and reads it, never rewrites it.
  const INBOX_COLS = ['Location'].concat(CLINIC_WORK_COLS, ['File', 'Received']);

  /** Table rows as objects: objects pass through; an array grid (XL.readWorkbook) takes its first non-empty row as the header. */
  function tableObjects(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const first = list.find((r) => r != null);
    if (!Array.isArray(first)) return list.filter((r) => r && typeof r === 'object');
    const [head, ...body] = list.filter((r) => Array.isArray(r) && Array.from(r).some((v) => str(v) !== ''));
    const h = Array.from(head || [], str);
    return body.map((r) => Object.fromEntries(h.map((k, i) => [k, r[i] == null ? '' : r[i]]).filter(([k]) => k)));
  }

  // "Last, First" and "First Last" as whole words (and either in brackets, dropped whole) for one person; null when a part
  // is blank. Whole words: patient "Le, Ed" must not hit provider "Cole, Edward MD".
  const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  function nameRe(last, first) {
    const l = reEsc(str(last)), f = reEsc(str(first));
    if (!l || !f) return null;
    const both = `(?<!\\p{L})(?:${l},\\s*${f}|${f}\\s+${l})(?!\\p{L})`;
    return new RegExp(`\\s*\\(\\s*${both}\\s*\\)|${both}`, 'giu');
  }
  // Only free text can quote a name. Structured fields (provider, location, payer, member ID, x) feed the derivation as is.
  const SCRUB_COLS = ['Action', 'Detail', 'Fix', 'Notes', 'Route Reason', 'Flags'];
  // Core's guarantor Action quotes "(Last, First)"; either part can be blank (a one-word Guar Name), so it goes whatever it holds.
  const GUAR_QUOTE = /(Found through the guarantor)\s*\([^()]*\)/g;
  /** The free-text fields of r with the names in res ([[regex, replacement]]) and core's guarantor quote taken out. */
  const scrubText = (s, res) => res.reduce((t, [re, word]) => t.replace(re, (m) => (m.trim()[0] === '(' ? '' : word)), s.replace(GUAR_QUOTE, '$1'));
  // A typed note can name the patient by one name alone or quote the DOB, so the SE Done Clinic Note gets more than the
  // full-name scrub: each name, and each part of one ("Smith-Jones", "De la Cruz"), as a whole word, and the DOB written
  // as M/D/YY(YY) with / - or ., as YYYY-MM-DD, or with the month's name. Over-scrubbing a word ("Will") beats a leak.
  function wordRes(last, first, word) {
    const parts = [last, first].map(str).flatMap((n) => [n].concat(n.split(/[\s-]+/))).filter((n) => n.replace(/[^\p{L}]/gu, '').length >= 2);
    return [...new Set(parts.map(low))].sort((a, b) => b.length - a.length)
      .map((n) => [new RegExp(`(?<!\\p{L})${reEsc(n)}(?!\\p{L})`, 'giu'), word]);
  }
  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  function dobRes(dob) {
    const d = parseDate(dob);
    if (!d) return [];
    const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
    const n0 = (v) => `0?${v}`, yr = `(?:${y}|${String(y).slice(2)})`;
    const mon = `${MONTHS[m - 1].slice(0, 3)}${MONTHS[m - 1].slice(3) ? `(?:${MONTHS[m - 1].slice(3)})?` : ''}\\.?`;
    const forms = [`${n0(m)}([/.-])${n0(day)}\\1${yr}`, `${y}-0?${m}-0?${day}`, `${mon}\\s+${n0(day)}(?:st|nd|rd|th)?,?\\s+${y}`, `${n0(day)}\\s+${mon},?\\s+${y}`];
    return [[new RegExp(`(?<![\\d/.-])(?:${forms.join('|')})(?![\\d/-])`, 'gi'), '[date removed]']];
  }
  function scrub(r, res) {
    const o = Object.assign({}, r);
    for (const k of SCRUB_COLS) if (typeof o[k] === 'string') o[k] = scrubText(o[k], res);
    return o;
  }
  function stripRecord(r, res, tag) {
    const o = {};
    for (const k of Object.keys(r || {})) if (k !== 'x' && !PHI_COLS.includes(k)) o[k] = r[k];
    o.LASFIR = tag;
    return scrub(o, res);
  }

  /** JSON for one Record cell. Too long: past appointments cut to the last 5 (future ones kept); still too long: an error. */
  function recordCell(rec, loc, runDay) {
    let s = JSON.stringify(rec);
    const appts = rec.x && Array.isArray(rec.x.appointments) ? rec.x.appointments : null;
    if (s.length >= CELL_MAX && appts) {
      const day = (a) => parseDate(a && a.date);
      const past = appts.filter((a) => day(a) && day(a) < runDay).sort((a, b) => day(a) - day(b));
      const keep = new Set(past.slice(-5));
      s = JSON.stringify(Object.assign({}, rec, { x: Object.assign({}, rec.x, { appointments: appts.filter((a) => !past.includes(a) || keep.has(a)) }) }));
    }
    if (s.length >= CELL_MAX) {
      throw new Error(`Clinic file for ${loc}: ${C.idLabel(rec) || 'a record'} has too much data for one Excel cell (${s.length.toLocaleString('en-US')} characters; the limit is ${CELL_MAX.toLocaleString('en-US')}), even with past appointments cut to the last 5. The file was not made.`);
    }
    return s;
  }

  /** Rules / benefits settings that differ from the defaults (Config Re-verify days, SFHP providers, ...). */
  function seSettings(cfg) {
    const c = cfg || {};
    const v = {
      sfhpProviders: c.sfhpProviders || [], reverifyDays: c.reverifyDays != null ? +c.reverifyDays : 7,
      reverifyNewMonth: !!(c.reverifyNewMonth != null ? c.reverifyNewMonth : c.reverify && c.reverify.newMonth),
      approachingPct: c.approachingPct != null ? +c.approachingPct : 0.2, locationState: c.locationState || {}, defaultState: str(c.defaultState),
    };
    const out = {};
    for (const k of Object.keys(SE_DEFAULTS)) if (JSON.stringify(v[k]) !== JSON.stringify(SE_DEFAULTS[k])) out[k] = v[k];
    return out;
  }

  /**
   * The Clinic Eligibility workbook for one location, as an XL.buildWorkbook spec plus {meta, name, counts}.
   * src: {patients (in-scope, after the skip window), results (merged Sohar Results), review (build review), learningRows,
   * metricRows, tasks (NextGen Tasks rows with Done), cfg, payerRuleRows, runAt, version}. Clinic = clinicSet membership
   * (next visit at loc on runAt); unmatched and no-identifier results stay with the admin.
   */
  function clinicPack(loc, src) {
    const s = src || {};
    const location = str(loc);
    if (allLocations(location)) throw new Error('Choose one location: a Clinic Template Tool file is for one clinic.');
    const runAt = s.runAt instanceof Date && !isNaN(s.runAt) ? s.runAt : new Date();
    const runDay = dayOf(runAt), runDate = fmtDate(runDay);
    const set = clinicSet(location, s.patients, s.results, runAt);
    const keys = new Set(set.map((e) => e.per));
    const patients = [], results = [], carried = new Set();
    const scrubs = new Map();
    for (const { p, per, rows } of set) {
      const x = xOf(p), g = x.guarantor;
      const tag = lasfir(p['Last Name'], p['First Name']);
      const res = [[nameRe(p['Last Name'], p['First Name']), tag], [g ? nameRe(g.lastName, g.firstName) : null, 'the guarantor']].filter((e) => e[0]);
      const noteRes = res.concat(wordRes(p['Last Name'], p['First Name'], tag), g ? wordRes(g.lastName, g.firstName, 'the guarantor') : [],
        dobRes(p.DOB), g ? dobRes(g.dob) : []);
      scrubs.set(per, { res, tag, noteRes });
      const rec = stripRecord(p, res, tag);
      const xs = Object.assign({}, x);
      delete xs.guarantor; delete xs.guarZip;
      rec.x = xs;
      patients.push({ Key: per, Record: recordCell(rec, location, runDay) });
      for (const r of rows) {
        const rr = stripRecord(r, res, tag);
        rr.key = per; // the patient key travels on the record (results carry it as a non-column)
        results.push({ Key: per, Record: recordCell(rr, location, runDay) });
      }
      for (const name of [p['Primary Payer'], p['Secondary Payer'] || x.secPayer]) if (str(name)) carried.add(low(name));
    }
    const review = (s.review || []).filter((r) => r && keys.has(str(r.MRN))).map((r) => {
      const { res, tag } = scrubs.get(str(r.MRN));
      const o = {};
      for (const k of C.REVIEW_COLS) o[k] = r[k] == null ? '' : r[k];
      // The build's name items quote the raw report name.
      if (/^name /i.test(str(o.Issue))) o.Detail = 'Not shown in the clinic file (patient name).';
      o.LASFIR = tag;
      return { Key: str(r.MRN), Record: recordCell(scrub(o, res), location, runDay) };
    });
    const payers = new Map();
    const addPayer = (r, part) => {
      const k = low(r && r['NextGen Payer']);
      if (!k || !carried.has(k)) return;
      if (!payers.has(k)) payers.set(k, { 'NextGen Payer': str(r['NextGen Payer']), learning: null, metrics: null });
      payers.get(k)[part] = r;
    };
    for (const r of s.learningRows || []) addPayer(r, 'learning');
    for (const r of s.metricRows || []) addPayer(r, 'metrics');
    // Done, and the Clinic Note the admin holds (the clinic's done note as brought back, or one typed on NextGen Tasks), so the
    // clinic sees its note arrived. A note typed by hand can quote the patient, so names (whole or one alone) and DOBs go.
    const done = (s.tasks || []).filter(Boolean).map((t) => [t, scrubs.get(keyOf(t)) || scrubs.get(str(t['Task Key']).split('|')[0])])
      .filter(([t, sc]) => sc && (str(t.Done) || str(t['Clinic Note'])))
      .map(([t, sc]) => ({ 'Task Key': str(t['Task Key']), 'Done': str(t.Done), 'Clinic Note': scrubText(str(t['Clinic Note']), sc.noteRes) }));
    const ruleRows = (s.payerRuleRows || []).filter(Boolean);
    const settings = seSettings(s.cfg);
    const meta = { Kind: SE_KIND, Location: location, 'Run Date': runDate, 'Run At': runAt.toISOString(), 'Sohar Version': str(s.version), Format: SE_FORMAT,
      Settings: JSON.stringify(settings) };
    const readMe = [
      `Clinic Eligibility file for ${location}, made ${runDate} by the central Sohar Eligibility admin.`,
      'It lists the patients whose next visit is at this clinic: their Sohar insurance results, benefits and what to do before the visit.',
      'To work the list, open Home > Clinic Template Tool. The pane shows Eligibility, Benefits and Payers.',
      'Patients show as LASFIR (first three letters of the last name, then of the first name) and MRN. The file has no patient names or dates of birth.',
      'Tick tasks done (with a short note if it helps), ask for a re-verify, or add a note in the pane. Your changes are saved on the Clinic Work sheet. Do not edit the hidden sheets.',
      'This file has PHI (MRNs, insurance and appointments). Keep it in your clinic\'s folder in the Clinic Eligibility library, MHS OneDrive or Teams only.',
      'Save the file in your clinic\'s folder in the Clinic Eligibility library. Excel on the web saves your changes on its own and the admin picks them up from there: you do not need to send the file back.',
      'If the file is not in the library, or the admin asks for it, save it and send it back to the admin through OneDrive or Teams. The admin brings your changes into Sohar.',
    ];
    const sheets = [
      { name: SE_SHEETS.readMe, table: 'tblSERead', columns: ['Read Me'], rows: readMe.map((t) => ({ 'Read Me': t })), colWidths: [110] },
      { name: SE_SHEETS.work, table: 'tblClinicWork', columns: CLINIC_WORK_COLS, rows: [], textColumns: CLINIC_WORK_COLS },
      { name: SE_SHEETS.info, table: 'tblSEInfo', columns: ['Setting', 'Value'], rows: Object.entries(meta).map(([k, v]) => ({ Setting: k, Value: v })), hidden: true },
      { name: SE_SHEETS.patients, table: 'tblSEPatients', columns: ['Key', 'Record'], rows: patients, hidden: true },
      { name: SE_SHEETS.results, table: 'tblSEResults', columns: ['Key', 'Record'], rows: results, hidden: true },
      { name: SE_SHEETS.review, table: 'tblSEReview', columns: ['Key', 'Record'], rows: review, hidden: true },
      { name: SE_SHEETS.payers, table: 'tblSEPayers', columns: ['Key', 'Record'],
        rows: [...payers].map(([k, v]) => ({ Key: k, Record: recordCell(v, location, runDay) })), hidden: true },
      { name: SE_SHEETS.done, table: 'tblSEDone', columns: ['Task Key', 'Done', 'Clinic Note'], rows: done, hidden: true },
    ];
    if (ruleRows.length) sheets.push({ name: SE_SHEETS.rules, table: 'tblPayerCredRules', columns: PR ? PR.HEADERS : Object.keys(ruleRows[0]), rows: ruleRows, hidden: true });
    return {
      meta, sheets, webExtension: Object.assign({}, CTT_EXTENSION), now: runAt,
      name: `Clinic Eligibility - ${location.replace(/[\\/:*?"<>|]/g, '-')} ${localDate(runAt)}.xlsx`,
      counts: { patients: patients.length, results: results.length, review: review.length, payers: payers.size, done: done.length, rules: ruleRows.length },
    };
  }

  /**
   * A Clinic Eligibility file's tables ({sheet name: rows}, rows as objects or an XL.readWorkbook grid) -> {meta, settings,
   * patients, results, review, payers, done, payerRules}. Optional sheets may be missing. Not a Clinic Eligibility file, or
   * a format this code does not read: a readable error.
   */
  function clinicUnpack(tables) {
    const t = tables || {};
    const find = (name) => { const k = Object.keys(t).find((n) => low(n) === low(name)); return k ? tableObjects(t[k]) : []; };
    const info = find(SE_SHEETS.info);
    if (!info.length) throw new Error('This is not a Clinic Eligibility file: it has no SE Info sheet. Ask the admin for a new file.');
    const meta = {};
    for (const r of info) if (str(r.Setting)) meta[str(r.Setting)] = str(r.Value);
    if (meta.Kind !== SE_KIND) throw new Error(`This file is ${meta.Kind ? `a "${meta.Kind}" file` : 'not marked'}, not a Clinic Eligibility file. Ask the admin for a new file.`);
    if (meta.Format !== SE_FORMAT) throw new Error(`This Clinic Eligibility file is format ${meta.Format || '(blank)'}; this add-in reads format ${SE_FORMAT}. Close and reopen the pane to update it, or ask the admin for a new file.`);
    const json = (sheet, text, i) => {
      try { return JSON.parse(text); } catch (e) { throw new Error(`The ${sheet} sheet, row ${i + 2}, cannot be read (${e.message}). Ask the admin for a new file.`); }
    };
    const records = (sheet) => find(sheet).map((r, i) => (str(r.Record) ? json(sheet, str(r.Record), i) : null)).filter((r) => r && typeof r === 'object');
    const blank = Object.fromEntries(PHI_COLS.map((c) => [c, '']));
    const settings = Object.assign(JSON.parse(JSON.stringify(SE_DEFAULTS)), meta.Settings ? json(SE_SHEETS.info, meta.Settings, info.findIndex((r) => str(r.Setting) === 'Settings')) : {});
    delete meta.Settings;
    return {
      meta, settings,
      // The stripped columns come back blank, so the rules see the same shape as admin rows.
      patients: records(SE_SHEETS.patients).map((p) => Object.assign({}, blank, p)),
      results: records(SE_SHEETS.results).map((r) => Object.assign({}, blank, r)),
      review: records(SE_SHEETS.review),
      payers: records(SE_SHEETS.payers),
      // Clinic Note: files made before it was added have no such column (still Format 1), so it reads blank.
      done: find(SE_SHEETS.done).map((r) => ({ 'Task Key': str(r['Task Key']), 'Done': str(r.Done), 'Clinic Note': str(r['Clinic Note']) })).filter((r) => r['Task Key']),
      payerRules: payerRulesFrom(find(SE_SHEETS.rules)),
    };
  }

  /** Clinic Work rows -> Map(kind|key -> row); the last row for each Kind + Key wins. Unknown kinds and blank keys are dropped. */
  function lastWork(work) {
    const m = new Map();
    for (const r of tableObjects(work)) {
      const kind = workKind(r.Kind), key = str(r.Key);
      if (!kind || !key) continue;
      const id = workId(kind, key);
      m.delete(id);
      m.set(id, Object.assign({ kind }, ...CLINIC_WORK_COLS.map((c) => ({ [c]: str(r[c]) })), { Kind: WORK_KINDS[kind] }));
    }
    return m;
  }

  /**
   * The clinic's view of an unpacked file at `now`: {rows (clinicRows + LASFIR), review (flag rows, then build review rows,
   * CLINIC_REVIEW_COLS + LASFIR + Source), flags (Map), benefits ({rows, pending}, benefitsView for the clinic's patients),
   * tasks (buildTasks with Done from SE Done then Clinic Work; clinic: true on patient-facing types; doneBy; clinicNote, the
   * Note on the task's Clinic Work "Task done" row; sentNote and Clinic Note, the admin's Clinic Note from SE Done), payers, notes
   * (Map MRN -> Clinic Work note), counts}. Clinic Work overrides Re-verify anyway, task Done and notes.
   */
  function clinicDerive(pack, work, now) {
    const pk = pack || {};
    const today = now instanceof Date && !isNaN(now) ? now : new Date();
    const st = Object.assign({}, SE_DEFAULTS, pk.settings || {});
    const loc = str(pk.meta && pk.meta.Location);
    const w = lastWork(work);
    const get = (kind, key) => w.get(workId(kind, key));
    const patients = (pk.patients || []).map((p) => {
      const r = get('reverify', keyOf(p));
      return r ? Object.assign({}, p, { x: Object.assign({}, xOf(p), { reverifyRequested: isOn(r.Value) }) }) : p;
    });
    const results = pk.results || [];
    const tags = new Map(patients.map((p) => [keyOf(p), str(p.LASFIR)]));
    const cfg = { today, sfhpProviders: st.sfhpProviders, reverifyDays: st.reverifyDays, reverifyNewMonth: !!st.reverifyNewMonth,
      payerRules: pk.payerRules || [], locationState: st.locationState };
    const all = allFlags(patients, results, cfg);
    const rows = clinicRows(loc, patients, results, all, today);
    const byKey = new Map();
    for (const r of rows) { r.LASFIR = tags.get(r.key) || ''; byKey.set(r.key, r); }
    // A patient whose visits here have all passed since the export drops out of everything.
    const flags = new Map([...all].filter(([k]) => byKey.has(k)));
    const review = clinicReviewRows(loc, patients, results, flags, today).map((r) => Object.assign(r, { LASFIR: tags.get(r.MRN) || '', Source: 'Flag' }))
      .concat((pk.review || []).filter((r) => byKey.has(str(r.MRN))).map((r) => {
        const c = byKey.get(str(r.MRN));
        return { 'Appt Date': c['Next Appt'], 'Severity': str(r.Severity), 'Pill': '', 'Issue': str(r.Issue), 'MRN': str(r.MRN), 'Last Name': '', 'First Name': '',
          'Event': c['Next Event'], 'Rendering': c['Next Provider'], 'Location': c.Location, 'Detail': str(r.Detail), 'Fix': str(r.Fix), 'Link': '',
          'Run Date': c['Run Date'], LASFIR: c.LASFIR, Source: 'Build' };
      }));
    const pers = new Map(rows.map((r) => [r['Per Nbr'], r.LASFIR]));
    const view = C.benefitsView(results);
    const benefits = { rows: view.rows.filter((b) => pers.has(str(b['Per Nbr']))).map((b) => Object.assign(b, { LASFIR: pers.get(str(b['Per Nbr'])) })),
      pending: (view.pending || []).filter((r) => pers.has(str(r['Per Nbr']))) };
    const sent = new Map((pk.done || []).map((d) => [str(d['Task Key']), str(d.Done)]).filter((e) => e[0] && e[1]));
    const sentNotes = new Map((pk.done || []).map((d) => [str(d['Task Key']), str(d['Clinic Note'])]).filter((e) => e[0] && e[1]));
    // A Task Key starts with the patient key (the MRN column can differ for a shared MRN).
    const who = (t) => str(t['Task Key']).split('|')[0];
    const tasks = C.buildTasks(results).filter((t) => byKey.has(who(t))).map((t) => {
      const k = str(t['Task Key']), was = sent.get(k) || '', cw = get('done', k);
      // SE Done holds the admin's Done, or a clinic tick brought back earlier (clinicDoneMark), which the clinic can still undo.
      const admin = isClinicDone(was) ? '' : was;
      // A blank Value is a note saved with the box never ticked here: Done stays as the file has it.
      const clinicDone = !admin && (cw && cw.Value ? isOn(cw.Value) : !!was);
      // Clinic No only undoes the clinic's own tick; the admin's Done stays. clinicNote: the clinic's own note (Clinic Work).
      // sentNote: the Clinic Note the admin had when the file was made ('note (By, Location, M/D)'), whatever the clinic edits.
      const sentNote = sentNotes.get(k) || '';
      return Object.assign(t, { Done: admin || (clinicDone ? was || 'Yes' : ''), 'Clinic Note': sentNote, doneBy: admin ? 'admin' : clinicDone ? 'clinic' : '',
        clinic: CLINIC_TASK_TYPES.includes(t.Task), LASFIR: tags.get(who(t)) || '', clinicNote: cw ? cw.Note : '', sentNote });
    });
    const notes = new Map();
    for (const r of w.values()) if (r.kind === 'note' && r.Note && byKey.has(r.Key)) notes.set(r.Key, { Note: r.Note, By: r.By, When: r.When });
    const carried = new Map();
    for (const p of patients) {
      if (!byKey.has(keyOf(p))) continue;
      for (const name of [p['Primary Payer'], p['Secondary Payer'] || xOf(p).secPayer]) {
        const k = low(name);
        if (!k || isSelfPay(name)) continue;
        if (!carried.has(k)) carried.set(k, { payer: str(name), keys: new Set() });
        carried.get(k).keys.add(keyOf(p));
      }
    }
    const recs = new Map((pk.payers || []).map((r) => [low(r['NextGen Payer']), r]));
    const payers = [...carried].map(([k, c]) => {
      const rec = recs.get(k) || {}, n = get('payer', c.payer);
      return { key: k, payer: c.payer, patients: c.keys.size, learning: rec.learning || null, metrics: rec.metrics || null,
        note: n ? { cardSays: n.Value, note: n.Note, by: n.By, when: n.When } : null };
    }).sort((a, b) => b.patients - a.patients || a.payer.localeCompare(b.payer));
    const open = tasks.filter((t) => !t.Done);
    const counts = {
      patients: rows.length, flagged: rows.filter((r) => r.Pills).length, review: review.length, high: review.filter((r) => r.Severity === 'High').length,
      tasks: tasks.length, openTasks: open.length, clinicTasks: open.filter((t) => t.clinic).length,
      reverifyAnyway: rows.filter((r) => r['Re-verify Anyway']).length, notes: notes.size, taskNotes: tasks.filter((t) => t.clinicNote).length, sentNotes: tasks.filter((t) => t.sentNote).length, payerNotes: payers.filter((p) => p.note).length, benefits: benefits.rows.length,
    };
    return { rows, review, flags, benefits, tasks, payers, notes, counts };
  }

  /**
   * Clinic Work after a change (or a list of them): the row with the same Kind + Key is replaced and the new one goes last,
   * with When stamped (ISO; change.When or `when` or now) and By. change: {Kind, Key, MRN, LASFIR, Value, Note, By}.
   * A change that leaves out Value or Note keeps the earlier row's (ticking a task keeps its done note, saving the note keeps
   * the tick). Returns rows as objects by CLINIC_WORK_COLS (other rows kept as they are, in order).
   */
  function workRows(work, change, when) {
    let rows = tableObjects(work).map((r) => Object.fromEntries(CLINIC_WORK_COLS.map((c) => [c, str(r[c])])));
    const stamp = (when instanceof Date && !isNaN(when) ? when : new Date()).toISOString();
    for (const ch of (Array.isArray(change) ? change : [change]).filter(Boolean)) {
      const kind = workKind(ch.Kind), key = str(ch.Key);
      if (!kind) throw new Error(`Clinic Work: "${str(ch.Kind)}" is not one of ${Object.values(WORK_KINDS).join(', ')}.`);
      if (!key) throw new Error(`Clinic Work: a ${WORK_KINDS[kind]} row needs a Key.`);
      const id = workId(kind, key);
      const prev = rows.find((r) => workId(workKind(r.Kind), r.Key) === id) || {};
      rows = rows.filter((r) => workId(workKind(r.Kind), r.Key) !== id);
      const row = Object.fromEntries(CLINIC_WORK_COLS.map((c) => [c, str(ch[c])]));
      for (const c of ['Value', 'Note']) if (ch[c] === undefined && prev[c]) row[c] = prev[c];
      rows.push(Object.assign(row, { Kind: WORK_KINDS[kind], Key: key, When: str(ch.When) || stamp }));
    }
    return rows;
  }

  /**
   * Admin: a returned file's Clinic Work applied to the admin tables (pure; the caller writes them). ctx: {location,
   * requests (Re-verify Requests rows), tasks (NextGen Tasks rows, read fresh from the sheet), notes (Clinic Notes),
   * payerNotes (Clinic Payer Notes), applied (Clinic Work Applied rows), now}. Rows not newer than the last one brought back
   * for the same Location + Kind + Key, or older than APPLIED_DAYS, are left out (counts.old: the same file sent again).
   * Re-verify anyway: core.mergeClinicRequests, By "Clinic file <Location>". Task done Yes: Done "Yes (clinic <Location>)" on
   * an open task; No clears Done only when it is that clinic mark (an admin "Yes" or "x" stays); a blank Value leaves Done,
   * and so does a Value the ledger says was already acted on (a note edit carries the tick along; the admin may have
   * reopened the task since). Task done rows apply only to tasks whose Next Location is the file's (counts.otherClinic).
   * A Task done row's Note goes to doneNotes (for the task's Clinic Note) whatever the Value; a blank Note changes nothing.
   * Notes and payer notes: newest When per MRN (payer) + Location wins; a blank newest note removes the row. Returns
   * {requests, done: Map(Task Key -> mark|''), doneNotes: Map(Task Key -> 'note (By, Location, M/D)'), notes, payerNotes,
   * applied, counts}. counts.applied is every row newer than the ledger; clinicChanges(counts) is how many changed something
   * (counts.tasksChanged: Task done rows that changed Done or the note).
   */
  function mergeClinicWork(work, ctx) {
    const c = ctx || {};
    const loc = str(c.location);
    const now = c.now instanceof Date && !isNaN(c.now) ? c.now : new Date();
    const counts = { reverifyAdded: 0, reverifyRemoved: 0, doneSet: 0, doneCleared: 0, doneNotes: 0, tasksChanged: 0, notes: 0, payerNotes: 0, skipped: 0, otherClinic: 0, old: 0, applied: 0 };
    const seenId = (l, kind, key) => `${low(l)}|${workId(kind, key)}`;
    const tick = (v) => (isOn(v) ? 'Yes' : str(v) ? 'No' : '');
    const seen = new Map();
    for (const r of tableObjects(c.applied)) {
      const kind = workKind(r.Kind), key = str(r.Key);
      if (kind && key) seen.set(seenId(r.Location, kind, key), { Location: str(r.Location), Kind: WORK_KINDS[kind], Key: key, When: str(r.When), Value: tick(r.Value) });
    }
    const keepFrom = new Date(now.getTime() - APPLIED_DAYS * 864e5).toISOString();
    const w = [], acted = new Map(); // Task done row -> the tick acted on last time ('' when never)
    for (const r of lastWork(work).values()) {
      const id = seenId(loc, r.kind, r.Key), prev = seen.get(id);
      // Older than the ledger keeps: it was brought back before (or is too old to act on), so trimming never re-applies it.
      if (whenAt(r.When) < whenAt(keepFrom) || (prev && r.When <= prev.When)) { counts.old++; continue; }
      const was = prev ? prev.Value : '';
      seen.set(id, { Location: loc, Kind: r.Kind, Key: r.Key, When: r.When, Value: r.kind === 'done' ? tick(r.Value) || was : '' });
      if (r.kind === 'done') acted.set(r, was);
      w.push(r);
    }
    counts.applied = w.length;
    const applied = [...seen.values()].filter((r) => r.When >= keepFrom);
    const of = (kind) => w.filter((r) => r.kind === kind);
    const m = C.mergeClinicRequests(c.requests || [], of('reverify').map((r) => ({ 'MRN': r.Key, 'Re-verify Anyway': isOn(r.Value) ? 'On' : 'Off', 'Requested': r.When })),
      `Clinic file${loc ? ` ${loc}` : ''}`);
    counts.reverifyAdded = m.added; counts.reverifyRemoved = m.removed;
    const cur = new Map(tableObjects(c.tasks).map((t) => [str(t['Task Key']), t]));
    const done = new Map(), doneNotes = new Map();
    for (const r of of('done')) {
      if (!cur.has(r.Key)) { counts.skipped++; continue; } // the task is gone (fixed, or the visit passed)
      const t = cur.get(r.Key), d = str(t.Done);
      // Only the clinic of the task's next visit (the one its file was made for) can tick it; a row typed for another
      // clinic's patient is left out. A file for every location ('' or All) can tick any task.
      if (!allLocations(loc) && !sameLoc(t['Next Location'], loc)) { counts.otherClinic++; continue; }
      // A blank Value is a note only (saved with the box never ticked): Done stays. So is a tick already acted on: the
      // clinic only edited the note, and Done may have been changed on NextGen Tasks since.
      const v = tick(r.Value), fresh = v !== acted.get(r);
      if (fresh && v === 'Yes' && !d) { done.set(r.Key, clinicDoneMark(loc)); counts.doneSet++; }
      else if (fresh && v === 'No' && isClinicDone(d)) { done.set(r.Key, ''); counts.doneCleared++; }
      // A blank note never erases the one already on NextGen Tasks.
      if (r.Note) { doneNotes.set(r.Key, doneNote(r, loc)); counts.doneNotes++; }
      if (done.has(r.Key) || r.Note) counts.tasksChanged++; // a tick and its note are one change
    }
    const upsert = (prev, cols, list, idOf, toRow) => {
      const map = new Map();
      for (const r of tableObjects(prev)) { const o = Object.fromEntries(cols.map((k) => [k, str(r[k])])); map.set(idOf(o), o); }
      let n = 0;
      for (const r of list) {
        const o = toRow(r), id = idOf(o), old = map.get(id);
        if (old && str(old.When) > o.When) continue; // a newer note is already there
        if (str(o.Note) || str(o['Card Says'])) { map.set(id, o); n++; } else if (map.delete(id)) n++; // a blank with nothing to remove changes nothing
      }
      return { rows: [...map.values()], n };
    };
    const notes = upsert(c.notes, CLINIC_NOTE_COLS, of('note'), (o) => `${o.MRN}|${low(o.Location)}`,
      (r) => ({ 'MRN': r.Key, 'Location': loc, 'Note': r.Note || r.Value, 'By': r.By, 'When': r.When }));
    const payerNotes = upsert(c.payerNotes, CLINIC_PAYER_NOTE_COLS, of('payer'), (o) => `${low(o['NextGen Payer'])}|${low(o.Location)}`,
      (r) => ({ 'NextGen Payer': r.Key, 'Location': loc, 'Card Says': r.Value, 'Note': r.Note, 'By': r.By, 'When': r.When }));
    counts.notes = notes.n; counts.payerNotes = payerNotes.n;
    return { requests: m.rows, done, doneNotes, notes: notes.rows, payerNotes: payerNotes.rows, applied, counts };
  }
  /** How many Clinic Work rows changed something in mergeClinicWork's counts (a tick already acted on, or a blank note over none, did not). */
  const clinicChanges = (n) => ['reverifyAdded', 'reverifyRemoved', 'tasksChanged', 'notes', 'payerNotes']
    .reduce((k, c) => k + (Number(n && n[c]) || 0), 0);

  /**
   * Admin: the Clinic Work Inbox (tblClinicWorkInbox, appended by the Power Automate flow, INBOX_COLS) -> Map(Location ->
   * Clinic Work rows), each Location in the order first seen and its rows oldest When first (so the newest per Kind + Key wins
   * in mergeClinicWork). Rows with a blank Location, Kind, Key or When, an unknown Kind, or older than APPLIED_DAYS are left
   * out; map.counts = {rows, blank, unknown, old}. Apply with mergeClinicWork per Location, threading requests, tasks,
   * notes, payerNotes and applied: the ledger makes a re-read or a row delivered twice (inbox and file) apply once.
   */
  function inboxWork(rows, now) {
    const at = now instanceof Date && !isNaN(now) ? now : new Date();
    const keepFrom = at.getTime() - APPLIED_DAYS * 864e5;
    const out = new Map(), names = new Map();
    const counts = { rows: 0, blank: 0, unknown: 0, old: 0 };
    for (const r of tableObjects(rows)) {
      const loc = str(r.Location);
      if (!loc || !str(r.Kind) || !str(r.Key) || !str(r.When)) { counts.blank++; continue; }
      if (!workKind(r.Kind)) { counts.unknown++; continue; }
      if (whenAt(r.When) < keepFrom) { counts.old++; continue; }
      if (!names.has(low(loc))) { names.set(low(loc), loc); out.set(loc, []); }
      out.get(names.get(low(loc))).push(Object.fromEntries(CLINIC_WORK_COLS.map((c) => [c, str(r[c])])));
      counts.rows++;
    }
    for (const list of out.values()) list.sort((a, b) => (whenAt(a.When) || 0) - (whenAt(b.When) || 0));
    out.counts = counts;
    return out;
  }

  return {
    payerRuleRows,
    REVERIFY_URL, PILLS, CLINIC_COLS, CLINIC_REVIEW_COLS, CLINIC_REQUEST_COLS, PAYER_EVIDENCE_COLS,
    isMediCal, isKaiser, isVA, isSFHP, isQMB, isCarelon, isUMR, isNP, isTMSEvent, isESKEvent, isIVKEvent, isInterventionEvent, isMedMgmtEvent, isVIP, isMedicare,
    payerRulesFrom,
    patientFlags, allFlags, reviewRows, clinicRows, clinicReviewRows, clinicEmail, payerEvidence, evidenceSummary,
    bestResult, lasfir, collapseFlags,
    CTT_EXTENSION, SE_SHEETS, CLINIC_WORK_COLS, CLINIC_NOTE_COLS, CLINIC_PAYER_NOTE_COLS, CLINIC_APPLIED_COLS, CLINIC_TASK_TYPES,
    clinicPack, clinicUnpack, clinicDerive, workRows, mergeClinicWork, clinicChanges, tableObjects, INBOX_COLS, inboxWork, APPLIED_DAYS,
  };
});
