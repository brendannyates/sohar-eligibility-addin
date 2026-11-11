// Office Script "Clinic lookup" (Clinic Workbook Pilot, v0.29). Run by the flow "Clinic Workbook - Clinic lookup" against the
// master workbook, after the flow has turned the clinic key into a clinic name. No PHI or keys in this file.
// Rules (Brendan, 11/10; flagged for Compliance):
// - scheduled:  the MRN has an upcoming visit at this clinic on the ledger -> summary + the last 5 events (90 days). No reason.
// - seen-12mo:  a visit at this clinic in the last 12 months -> a reason is required -> summary, no history.
// - no-history: in the master but never seen here -> a reason and a note are required -> coverage, copay and card only.
// - not-found:  not in the master -> nothing to show; "Request eligibility check" (purpose request-check, with a reason)
//               adds a Needs Check row the next rotation picks up.
// Every call returns a log row for the audit workbook (who, clinic, MRN, tier, reason, outcome, which fields; no values).
//
// input (JSON text): {clinic, mrn, purpose, reason, note, staff, requestId, today ("YYYY-MM-DD", Pacific), now (ISO)}
// Returns JSON text: {status, body (what the add-in gets), log (LOOKUP_LOG_COLS object)}.

type Cell = string | number | boolean;
type Rec = { [k: string]: string };
interface Ask { clinic: string, mrn: string, purpose: string, reason: string, note: string, staff: string, requestId: string, today: string, now: string }
interface Answer { status: number, body: { [k: string]: Cell | Rec | Rec[] }, log: Rec, needs: Rec | null }

const SUMMARY = ['LASFIR', 'MRN', 'Coverage', 'Payer', 'Plan', 'Copay', 'Coinsurance', 'Deductible Remaining', 'Deductible Status',
  'OOP Remaining', 'OOP Status', 'Card On File', 'OK to Collect', 'Next Visit Here', 'As Of'];
const SMALL = ['LASFIR', 'MRN', 'Coverage', 'Payer', 'Copay', 'Coinsurance', 'Card On File', 'As Of'];

function str(v: Cell | undefined | null): string { return v === undefined || v === null ? '' : String(v).trim(); }
function mrnOf(v: Cell | undefined): string { return str(v).replace(/^0+(?=\d)/, ''); }
function low(v: Cell | undefined): string { return str(v).toLowerCase(); }
/** A ledger date (YYYY-MM-DD, M/D/YYYY or an Excel serial) as YYYY-MM-DD; '' when unreadable. */
function isoDay(v: Cell): string {
  if (typeof v === 'number' && v > 59) { const d = new Date(Math.round((v - 25569) * 864e5)); return d.toISOString().slice(0, 10); }
  const s = str(v);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  return '';
}
function addDays(iso: string, n: number): string { const d = new Date(`${iso}T12:00:00Z`); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function coverageWord(s: string): string {
  if (!s) return 'Not checked';
  if (/^complete\.eligible/i.test(s)) return 'Eligible';
  if (/^complete\.ineligible/i.test(s)) return 'Not eligible';
  if (/^error/i.test(s)) return 'Check failed';
  if (/pending/i.test(s)) return 'Pending';
  return s;
}
function objects(header: Cell[], body: Cell[][]): Rec[] {
  const h = header.map((x) => str(x));
  return body.filter((r) => r.some((c) => str(c) !== '')).map((r) => { const o: Rec = {}; h.forEach((k, i) => { if (k) o[k] = typeof r[i] === 'number' && /date|when|as of|checked/i.test(k) ? String(r[i]) : str(r[i]); }); return o; });
}

/** Pure: the whole decision from the tables (as objects) and the ask. */
function decide(ask: Ask, ledger: Rec[], history: Rec[], cards: Rec[], consents: Rec[]): Answer {
  const mrn = mrnOf(ask.mrn), clinic = str(ask.clinic), today = ask.today;
  const purpose = ['scheduled', 'unscheduled', 'request-check'].indexOf(ask.purpose) >= 0 ? ask.purpose : 'unscheduled';
  const reason = str(ask.reason).slice(0, 80), note = str(ask.note).slice(0, 300);
  const mine = ledger.filter((r) => mrnOf(r.MRN) === mrn && !/not on the latest report/i.test(str(r.Note)));
  const here = mine.filter((r) => low(r.Location) === low(clinic));
  const upcoming = here.filter((r) => isoDay(r['Appt Date']) >= today).sort((a, b) => (isoDay(a['Appt Date']) + str(a.Time) < isoDay(b['Appt Date']) + str(b.Time) ? -1 : 1));
  const past = here.filter((r) => { const d = isoDay(r['Appt Date']); return d && d < today && d >= addDays(today, -365); });
  const tier = upcoming.length ? 'scheduled' : past.length ? 'seen-12mo' : mine.length ? 'no-history' : 'not-found';
  let decision = 'ok';
  if (!/^\d{3,12}$/.test(mrn)) decision = 'bad-mrn';
  else if (tier !== 'scheduled' && !reason) decision = 'need-reason';
  else if (purpose === 'request-check') decision = tier === 'not-found' ? 'queue' : 'already-there';
  else if (tier === 'no-history' && !note) decision = 'need-note';
  else if (tier === 'not-found') decision = 'not-found';
  const newest = mine.slice().sort((a, b) => (str(a['Last Checked']) < str(b['Last Checked']) ? 1 : -1)).find((r) => str(r['Sohar Status'])) || mine[0] || {};
  const card = cards.find((r) => mrnOf(r.MRN) === mrn);
  const ok = consents.filter((r) => mrnOf(r.MRN) === mrn).sort((a, b) => (str(a.When) < str(b.When) ? 1 : -1))[0];
  const full: Rec = {
    'LASFIR': str(newest.LASFIR), 'MRN': mrn, 'Coverage': coverageWord(str(newest['Sohar Status'])), 'Payer': str(newest['Sohar Payer']), 'Plan': str(newest.Plan),
    'Copay': str(newest.Copay), 'Coinsurance': str(newest.Coinsurance), 'Deductible Remaining': str(newest['Deductible Remaining']), 'Deductible Status': str(newest['Deductible Status']),
    'OOP Remaining': str(newest['OOP Remaining']), 'OOP Status': str(newest['OOP Status']),
    'Card On File': str(card ? card['Card On File'] : newest['Card On File']), 'OK to Collect': ok ? str(ok['OK to Collect']) : str(newest['Verbal Consent']),
    'Next Visit Here': upcoming[0] ? `${isoDay(upcoming[0]['Appt Date'])} ${str(upcoming[0].Time)}`.trim() : '', 'As Of': str(newest['Last Checked']),
  };
  const fields = tier === 'scheduled' || tier === 'seen-12mo' ? SUMMARY : SMALL;
  const patient: Rec = {};
  for (const f of fields) patient[f] = full[f] || '';
  const from = addDays(today, -90);
  const hist = tier === 'scheduled' ? history.filter((r) => mrnOf(r.MRN) === mrn && (isoDay(r.When) || str(r.When).slice(0, 10)) >= from)
    .sort((a, b) => (str(a.When) < str(b.When) ? 1 : -1)).slice(0, 5)
    .map((r) => ({ When: isoDay(r.When) || str(r.When).slice(0, 10), Kind: str(r.Kind), Detail: str(r.Detail).slice(0, 120), Amount: str(r.Amount) })) : [];
  const msg: { [k: string]: string } = {
    'bad-mrn': 'Type the MRN (digits only).',
    'need-reason': 'This patient has no visit booked here. Pick why you need the details.',
    'need-note': 'This patient has no history at this clinic. Add a note saying why you need the details.',
    'not-found': 'This MRN is not in the eligibility master. You can ask for an eligibility check.',
    'queue': 'Added to the next eligibility check. The result comes with the next clinic file.',
    'already-there': 'This patient is already in the master, so no check was added.',
    'ok': '',
  };
  let body: { [k: string]: Cell | Rec | Rec[] };
  if (decision === 'ok') body = { ok: true, tier, patient, history: hist, message: msg.ok };
  else if (decision === 'not-found') body = { ok: true, tier, message: msg['not-found'] };
  else if (decision === 'queue') body = { ok: true, tier, queued: true, message: msg.queue };
  else if (decision === 'already-there') body = { ok: true, tier, message: msg['already-there'] };
  else body = { ok: false, tier: decision === 'bad-mrn' ? '' : tier, error: decision, needReason: decision === 'need-reason', needNote: decision === 'need-note', message: msg[decision] };
  const shown = decision === 'ok' ? fields.join(', ') + (hist.length ? `, history ${hist.length}` : '') : '';
  const log: Rec = { 'When': ask.now, 'Clinic': clinic, 'Staff': str(ask.staff).slice(0, 60), 'MRN': mrn, 'Tier': decision === 'bad-mrn' ? '' : tier,
    'Reason': [reason, note].filter(Boolean).join(' · '), 'Outcome': decision, 'Fields': shown, 'Request Id': str(ask.requestId).slice(0, 60) };
  const needs: Rec | null = decision === 'queue' ? { 'Row Key': str(ask.requestId) || `${mrn}|${ask.now}`, 'MRN': mrn, 'LASFIR': '', 'Location': clinic, 'Tier': tier,
    'Reason': reason, 'Note': note, 'Requested By': str(ask.staff).slice(0, 60), 'Requested At': ask.now, 'Status': 'Open', 'Done At': '', 'Result': '' } : null;
  return { status: 200, body, log, needs };
}

function main(workbook: ExcelScript.Workbook, input: string): string {
  let ask: Ask;
  try { ask = JSON.parse(input); } catch (e) { return JSON.stringify({ status: 400, body: { ok: false, error: 'bad-request', message: 'The lookup request was not readable.' }, log: null }); }
  const read = (name: string): Rec[] => {
    const t = workbook.getTable(name);
    if (!t) return [];
    return objects(t.getHeaderRowRange().getValues()[0], t.getRangeBetweenHeaderAndTotal().getValues() as Cell[][]);
  };
  const a = decide(ask, read('tblEligibilityChecks'), read('tblPatientHistory'), read('tblCardOnFile'), read('tblClinicConsent'));
  if (a.needs) {
    const t = workbook.getTable('tblNeedsCheck');
    if (t) {
      const h = t.getHeaderRowRange().getValues()[0].map((x) => str(x));
      for (const c of h) t.getColumn(c).getRangeBetweenHeaderAndTotal().setNumberFormatLocal('@');
      t.addRows(-1, [h.map((c) => a.needs ? a.needs[c] || '' : '')]);
    } else a.body = { ok: false, error: 'not-set-up', message: 'Eligibility checks cannot be requested yet. Ask the admin.' };
  }
  return JSON.stringify({ status: a.status, body: a.body, log: a.log });
}
