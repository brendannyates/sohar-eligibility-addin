import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const C = require('../core.js');
const F = require('../fees.js');
const REC = require('../record.js');

// v0.20: one record per patient, Sohar Flow results, report profiles, guarantor vs subscriber, and the West Coast
// calculator (parse + math). Synthetic data only: every name, MRN, payer and rate below is made up.

let refN = 0;
const row = (o) => ({ ...Object.fromEntries(C.RESULTS_COLS.map((c) => [c, ''])), 'Sohar Ref': `ref-${++refN}`, ...o });

// ---------- Sohar Flow results ----------
const flowCsv = (rows) => [REC.FLOW_COLS, ...rows.map((r) => REC.FLOW_COLS.map((c) => r[c] ?? ''))];
test('Flow files are recognised by their header and read by column name', () => {
  const shuffled = [...REC.FLOW_COLS].reverse();
  const rows = [['Flow Results'], shuffled, shuffled.map((c) => (c === 'MRN' ? '700001' : c === 'Key' ? 'k1' : c === 'Sohar Est Patient' ? '41.11' : ''))];
  assert.equal(REC.isFlowFile(rows), true);
  assert.equal(REC.isFlowFile([['id', 'status', 'planName']]), false);
  const [r] = REC.flowRecords(rows);
  assert.equal(r.MRN, '700001');
  assert.equal(r['Sohar Est Patient'], '41.11');
});

test('mergeFlow keeps one row per visit run; an older file never replaces a newer result', () => {
  const newer = { Key: 'k1', MRN: '700001', 'Run Status': 'COMPLETE', 'Last Updated': '2026-10-29T02:00:00Z', 'Sohar Est Patient': '41.11' };
  const older = { ...newer, 'Run Status': 'IN_PROGRESS', 'Last Updated': '2026-10-28T02:00:00Z', 'Sohar Est Patient': '' };
  const a = REC.mergeFlow([], [{ name: 'Flow Results (SANDBOX).csv', rows: flowCsv([newer, { Key: 'k2', MRN: '700002' }]) }], 'now');
  assert.equal(a.added, 2);
  const b = REC.mergeFlow(a.rows, [{ name: 'old.csv', rows: flowCsv([older]) }], 'later');
  assert.equal(b.replaced, 0);
  assert.equal(b.rows.find((r) => r.Key === 'k1')['Run Status'], 'COMPLETE');
  assert.equal(REC.isSandbox(b.rows[0]), true);
  const c = REC.mergeFlow(a.rows, [{ name: 'Flow Results.csv', rows: flowCsv([{ ...newer, 'Last Updated': '2026-10-30T01:00:00Z', 'Sohar Est Patient': '51.11' }]) }], 'x');
  assert.equal(c.replaced, 1);
  assert.equal(c.rows.find((r) => r.Key === 'k1')['Sohar Est Patient'], '51.11');
});

test('flowFor: upcoming visits soonest first; with none upcoming, the latest past one', () => {
  const f = [
    { MRN: '700001', 'Appt Date': '11/05/2026', Time: '1:30 PM' },
    { MRN: '700001', 'Appt Date': '11/05/2026', Time: '9:00 AM' },
    { MRN: '700001', 'Appt Date': '10/01/2026', Time: '9:00 AM' },
    { MRN: '700002', 'Appt Date': '10/01/2026', Time: '9:00 AM' },
    { MRN: '700002', 'Appt Date': '09/01/2026', Time: '9:00 AM' },
  ];
  const asOf = new Date('2026-10-30T00:00:00Z');
  assert.deepEqual(REC.flowFor(f, '700001', asOf).map((r) => r.Time), ['9:00 AM', '1:30 PM']);
  assert.deepEqual(REC.flowFor(f, '700002', asOf).map((r) => r['Appt Date']), ['10/01/2026']);
  assert.deepEqual(REC.flowFor(f, '799999', asOf), []);
  assert.equal(REC.flowAmount('41.1'), '$41.10');
  assert.equal(REC.flowAmount(''), '');
});

// ---------- one record per patient ----------
const results = [
  row({ MRN: '700001', 'Per Nbr': '9001', LASFIR: 'SMIJAN', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-28T10:00:00Z',
    'Sohar Payer': 'Acme Health', Deductible: '1000', 'Deductible Remaining': '400', 'OOP Max': '3000', 'OOP Remaining': '2500', 'Work Queue': 'Clear',
    'Upcoming Detail': '11/05/2026 9:00 AM Follow-up @ Harbor Point; 11/12/2026 9:00 AM Follow-up @ Harbor Point', 'Next Appt': '11/05/2026' }),
  row({ MRN: '700001', 'Per Nbr': '9001', LASFIR: 'SMIJAN', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Date': '2026-10-20T10:00:00Z', 'Sohar Payer': 'Acme Health', 'Work Queue': 'Action needed' }),
  row({ MRN: '700001', 'Per Nbr': '9001', LASFIR: 'SMIJAN', 'Sent As': 'Verification', 'Sohar Status': 'error.member.id', 'Sohar Date': '2026-10-29T10:00:00Z', Action: 'Superseded by a later result', 'Work Queue': 'Action needed' }),
  row({ MRN: '700002', 'Per Nbr': '9002', LASFIR: 'DOEALE', 'Sent As': 'Discovery', 'Sohar Status': 'No coverage found', 'Sohar Date': '2026-10-27T10:00:00Z', 'Work Queue': 'Action needed' }),
];

test('patientRecord: newest live result leads; the rest and older ledger checks form a newest-first timeline', () => {
  const outcomes = [
    { MRN: '700001', Kind: 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Date': '2026-10-20T11:00:00Z' }, // same day + status as a result: not repeated
    { MRN: '700001', Kind: 'Verification', 'Sohar Status': 'complete.ineligible', 'Sohar Date': '2026-09-15T10:00:00Z', 'Returned Payer': 'Old Plan' },
    { MRN: '700002', Kind: 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-09-01T10:00:00Z' },
  ];
  const rec = REC.patientRecord(results, '700001', { outcomes, flow: [{ MRN: '700001', 'Appt Date': '11/05/2026', Time: '9:00 AM' }], asOf: new Date('2026-10-30T00:00:00Z') });
  assert.equal(rec.lead['Sohar Status'], 'complete.eligible', 'a superseded result never leads');
  assert.equal(rec.results.length, 3);
  assert.deepEqual(rec.timeline.map((e) => `${e.source}:${e.status}`), ['result:error.member.id', 'result:Coverage found', 'ledger:complete.ineligible']);
  assert.equal(rec.timeline[2].payer, 'Old Plan');
  assert.deepEqual(rec.upcoming, ['11/05/2026 9:00 AM Follow-up @ Harbor Point', '11/12/2026 9:00 AM Follow-up @ Harbor Point']);
  assert.equal(rec.flow.length, 1);
  assert.ok(rec.financial, 'verified benefits give financial progress');
  assert.equal(rec.financial['Deductible Remaining'], '400');
  assert.equal(REC.patientRecord(results, '799999'), null);
  assert.equal(REC.patientRecord(results, '700002').financial, null, 'discovery alone has no benefits');
});

test('byPatient: one entry per MRN with every queue its results are in', () => {
  const g = REC.byPatient(results);
  assert.equal(g.length, 2);
  const a = g.find((x) => x.row.MRN === '700001');
  assert.equal(a.resultCount, 3);
  assert.equal(a.row['Sohar Status'], 'complete.eligible');
  assert.deepEqual(a.queues.sort(), ['Action needed', 'Clear']);
  assert.equal(a.rows[0]['Sohar Status'], 'error.member.id', 'rows newest first');
  const loose = REC.byPatient([row({ 'Sohar Ref': 'r1' }), row({ 'Sohar Ref': 'r2' })]);
  assert.equal(loose.length, 2, 'unmatched rows stay one per result');
});

// ---------- report profile: by header, never by position ----------
test('profiles reads gender, pronouns, previous last name and guarantor by header name', () => {
  const hdr = ['', 'Appointment Pre-Check Complete', 'Res Name', 'Pat Sts', 'Pat Name', 'Appt Date', 'Beg Tm', 'Event', 'Loc Name', 'Gender Identity', 'Current Gender',
    'Preferred Pronoun', 'MRN', 'Status', 'Birth Dt', 'Payer Name (Cht Guar)', 'Per Nbr', 'Prev Lst Name', 'Guar Name', 'Guar Zip'];
  const at = (o) => hdr.map((h) => o[h] ?? '');
  const rows = [['Mindful Health Solutions'], [], hdr,
    at({ 'Pat Name': 'Rivera, Sam', 'Appt Date': '11/05/2026', 'Per Nbr': '9101', MRN: '710001', 'Gender Identity': 'Non-binary', 'Current Gender': 'X', 'Preferred Pronoun': 'they/them', 'Guar Name': 'Rivera, Sam', 'Payer Name (Cht Guar)': 'Acme Health' }),
    at({ 'Pat Name': 'Nguyen, Lee', 'Appt Date': '11/06/2026', 'Per Nbr': '9102', MRN: '710002', 'Prev Lst Name': 'Tran', 'Guar Name': 'Tran, Lee', 'Payer Name (Cht Guar)': 'Acme Health' }),
    at({ 'Pat Name': 'Okafor, Ada', 'Appt Date': '11/06/2026', 'Per Nbr': '9103', MRN: '710003', 'Guar Name': 'Okafor, Ben', 'Payer Name (Cht Guar)': 'Acme Health' }),
  ];
  const det = C.detectHeader(rows);
  assert.ok(det, 'header found below the title rows');
  const p = REC.profiles(rows, det);
  assert.deepEqual([p.get('9101').genderIdentity, p.get('9101').currentGender, p.get('9101').pronouns], ['Non-binary', 'X', 'they/them']);
  assert.equal(p.get('9101').guarantorSelf, true);
  assert.equal(p.get('9102').guarantorSelf, true, 'guarantor under the previous last name is the patient');
  assert.equal(p.get('9103').guarantorSelf, false);
  assert.equal(REC.profiles(rows, det, new Set(['9103'])).size, 1, 'only the asked Per Nbrs');
});

test('guarantorCheck: NextGen guarantor vs the Sohar subscriber, LASFIR only', () => {
  const pt = { firstName: 'Ada', lastName: 'Okafor' };
  const parent = { firstName: 'Ben', lastName: 'Okafor' }, other = { firstName: 'Cy', lastName: 'Moss' };
  const prof = (g, self) => ({ patient: pt, guarantor: g, guarantorSelf: !!self, prevLastName: '' });
  assert.equal(REC.guarantorCheck(prof(parent), parent).status, 'Match');
  assert.equal(REC.guarantorCheck(prof(parent), pt).status, 'Subscriber is the patient, guarantor is not');
  const dep = REC.guarantorCheck(prof(pt, true), parent);
  assert.equal(dep.status, 'Guarantor is the patient, subscriber is not');
  assert.equal(dep.severity, 'medium', 'often fine (an adult on a partner plan); check the insured');
  assert.match(dep.text, /insured/);
  assert.equal(REC.guarantorCheck(prof(parent), other).severity, 'high');
  const jr = { firstName: 'Ada', lastName: 'Okafor', dob: '02/02/1960' };
  assert.equal(REC.guarantorCheck({ patient: { ...pt, dob: '03/03/2000' }, guarantor: { ...pt, dob: '03/03/2000' }, guarantorSelf: true }, jr).status, 'Guarantor is the patient, subscriber is not', 'same name, different DOB: another person');
  assert.match(dep.text, /OKABEN/);
  const diff = REC.guarantorCheck(prof(parent), other);
  assert.equal(diff.status, 'Different people');
  assert.doesNotMatch(diff.text, /Okafor|Moss|Ben|Cy\b/, 'no names, LASFIR only');
  assert.equal(REC.guarantorCheck(prof(parent), null).status, 'No subscriber from Sohar');
  assert.equal(REC.guarantorCheck({ patient: pt, guarantor: null }, parent).status, 'No guarantor on the report');
  assert.equal(REC.samePerson({ firstName: 'Katherine', lastName: 'Ng', dob: '01/02/1980' }, { firstName: 'Kat', lastName: 'Ng', dob: '01/02/1980' }), true);
  assert.equal(REC.samePerson({ firstName: 'Kat', lastName: 'Ng', dob: '01/02/1980' }, { firstName: 'Kat', lastName: 'Ng', dob: '03/04/1990' }), false, 'DOBs differ');
  assert.deepEqual(REC.subscriberOf({ subscriberfirstname: 'Ben', subscriberlastname: 'Okafor', subscriberdateofbirth: '' }), { firstName: 'Ben', lastName: 'Okafor', dob: '' });
  assert.equal(REC.subscriberOf({ subscriberfirstname: '', subscriberlastname: '' }), null);
});

test('no Flow or profile column carries a name or DOB', () => {
  for (const c of REC.FLOW_SHEET_COLS) assert.doesNotMatch(c, /(first|last)\s*name|^name$|dob|birth|subscriber/i, c);
});

// ---------- West Coast calculator ----------
const calcPage = (B) => `<!DOCTYPE html><html><head><script>\nconst BUILTIN = ${JSON.stringify(B)};\nfunction classifyProvider() {}\n</script></head></html>`;
const BUILTIN = {
  t: { loc: ['Therapy-CA', 'Bayview'], payer: ['Self Pay-CA', 'Cash', 'Acme Health', 'Zephyr'], prov: ['Therapy', 'Rate', 'MD', 'Nurse', 'MD/NP', 'PsyD'],
    cpt: ['90834', '99214', '90792'], desc: ['Psychotherapy', 'Established visit', 'Psych eval'], note: ['', 'eff 1/1/2026'] },
  r: [[0, 0, 0, 0, 0, 191.11, 0], [1, 1, 1, 1, 1, 151.11, 0], [1, 2, 2, 1, 1, 171.11, 1], [1, 2, 3, 1, 1, 151.11, 0], [1, 3, 4, 2, 2, 291.11, 0], [1, 2, 5, 0, 0, 91.11, 0]],
  locations: [{ name: 'Therapy-CA', info: 'Therapy rates' }, { name: 'Bayview', info: 'Locations: Harbor Point, Mill Street | Bayview Fee Schedule 2026 Locality 9 | made up' }],
  source: 'Synthetic-Fee-Schedule.xlsx',
};

test('parseCalculatorHtml: built-in rates become fee schedule rows (MD, NP, both, Cash); therapy is skipped', () => {
  const html = calcPage(BUILTIN);
  assert.equal(F.isCalculatorHtml(html), true);
  const p = F.parseCalculatorHtml(html);
  assert.equal(p.source, 'Synthetic-Fee-Schedule.xlsx');
  assert.deepEqual(p.tabs.map((t) => t.name), ['Bayview']);
  assert.equal(p.skipped[0].name, 'Therapy-CA');
  assert.equal(p.counts.therapy, 2);
  const t = p.tabs[0];
  assert.equal(t.locality, 'Locality 9');
  assert.deepEqual(t.locations, ['Harbor Point', 'Mill Street']);
  const idx = F.feeIndex(p.rows, p.tabs).get('Bayview');
  assert.equal(idx.payers.get('Acme Health').MD.get('99214'), 171.11);
  assert.equal(idx.payers.get('Acme Health').NP.get('99214'), 151.11);
  assert.equal(idx.payers.get('Zephyr').MD.get('90792'), 291.11, 'MD/NP column fills both');
  assert.equal(idx.payers.get('Zephyr').NP.get('90792'), 291.11);
  assert.equal(idx.cash.get('99214'), 151.11);
  assert.throws(() => F.parseCalculatorHtml('<html>no schedule</html>'), /no built-in fee schedule/);
  assert.throws(() => F.parseCalculatorHtml('const BUILTIN = {"t":{}};\n'), /not in the expected layout/);
});

test('mergeSchedules: saved rates are kept, calculator rates fill the gaps, differences are counted', () => {
  const calc = F.parseCalculatorHtml(calcPage(BUILTIN));
  const saved = [{ Tab: 'Bayview', Locality: 'Locality 9', Locations: '', Code: '99214', Description: '', Payer: 'Acme Health', Level: 'MD', Rate: 181.11, Note: '' }];
  const m = F.mergeSchedules(saved, [{ name: 'Bayview', locations: [] }], calc);
  assert.equal(m.conflicts, 1);
  assert.equal(m.added, calc.rows.length - 1);
  assert.equal(F.feeIndex(m.rows).get('Bayview').payers.get('Acme Health').MD.get('99214'), 181.11, 'saved rate wins');
  assert.equal(m.tabs.length, 1);
  const fresh = F.mergeSchedules([], [], calc);
  assert.equal(fresh.added, calc.rows.length);
  assert.equal(fresh.counts.tabs, 1);
});

test('calcEstimate is the calculator\'s math and agrees with the Benefits table math on the same inputs', () => {
  const lines = [{ code: '99214', units: 1, allowed: 171.11 }, { code: '90833', units: 1, allowed: 91.11 }];
  const ben = { deductibleRemaining: '100', coinsurance: '0.2', copay: '20', oopRemaining: '1000' };
  const a = F.estimateMath(lines, ben, F.CALC_MODE_TOGGLES.all);
  const b = F.calcEstimate(lines.map((l) => ({ allowed: l.allowed })), { planType: 'commercial', mode: 'all', deductible: 100, deductibleMet: 0, coins: 0.2, copay: 20, oopMax: 1000, oopMet: 0 });
  assert.equal(b.patient, a.patient);
  assert.equal(b.insurance, a.plan);
  const c = F.calcEstimate([{ allowed: 200 }], { planType: 'commercial', mode: 'ded_coins', deductible: 500, deductibleMet: 450, coins: 0.2, oopMax: 600, oopMet: 580 });
  assert.equal(c.patient, 20, 'capped at the OOP max remaining');
  assert.equal(c.oopReached, true);
  assert.equal(F.calcEstimate([{ allowed: 200 }], { planType: 'self_pay', mode: 'all', copay: 30 }).patient, 200);
  assert.equal(F.calcEstimate([{ allowed: 200 }], { planType: 'medicare', mode: 'ded_coins', deductible: 250, deductibleMet: 0, coins: 0.2, secCoins: true, secDed: true }).patient, 0, 'Medigap covers both');
  assert.equal(F.calcEstimate([{ allowed: 200, svcCoins: 0.5 }], { planType: 'commercial', mode: 'ded_coins', coins: 0.2 }).patient, 100, 'per-code coinsurance');
  assert.equal(F.calcModeOf(F.CALC_MODE_TOGGLES.copay_only), 'copay_only');
  assert.equal(F.calcModeOf({ copay: 'None' }), 'ded_coins');
  assert.equal(F.calcPlanType('Noridian JE'), 'medicare');
  assert.equal(F.calcPlanType('Acme Medicare Advantage'), 'medicare_adv');
  assert.equal(F.calcPlanType('Partnership Medi-Cal'), 'medicaid');
  assert.equal(F.calcPlanType('TriWest VA CCN'), 'va');
});

test('calcInputs pre-fills the calculator from Sohar; calcPatient prices MD and NP from the payer column', () => {
  const calc = F.parseCalculatorHtml(calcPage(BUILTIN));
  const idx = F.feeIndex(calc.rows, calc.tabs).get('Bayview');
  const p = idx.payers.get('Acme Health');
  const tables = { rates: { 'Bayview|Acme Health': { '99214': { MD: p.MD.get('99214'), NP: p.NP.get('99214') } } }, gross: { Bayview: {} }, cash: { Bayview: Object.fromEntries(idx.cash) }, desc: { Bayview: Object.fromEntries(idx.desc) } };
  const ctx = { tab: 'Bayview', ok: true, feePayer: 'Acme Health', rateKey: 'Bayview|Acme Health', payer: 'Acme Health', flags: [], defaults: { ...F.TOGGLE_DEFAULTS },
    benefits: { inn: { deductibleRemaining: '400', coinsurance: '20%', copay: '0', oopRemaining: '2500' }, oon: {} }, codes: [{ code: '99214', units: 1 }] };
  const inp = F.calcInputs(ctx, { Deductible: '1000', 'OOP Max': '3000', Plan: 'PPO' });
  assert.deepEqual([inp.planType, inp.mode, inp.deductible, inp.deductibleMet, inp.coins, inp.copay, inp.oopMax, inp.oopMet], ['commercial', 'all', 1000, 600, 20, 0, 3000, 500]);
  const e = F.calcPatient(ctx, tables, inp, ctx.codes);
  assert.equal(e.md.patient, 171.11, 'all of it toward the deductible');
  assert.equal(e.np.patient, 151.11);
  assert.ok(F.calcText(e.md, inp).some((t) => /Patient pays: \$171\.11/.test(t)));
  const self = F.calcPatient(ctx, tables, { ...inp, planType: 'self_pay' }, ctx.codes);
  assert.equal(self.md.patient, 151.11, 'self pay uses the Cash rate');
  assert.equal(self.same, true);
  assert.match(F.calcPatient(ctx, tables, inp, [{ code: '90999', units: 1 }]).md.none, /No MD rate/);
  assert.deepEqual(F.calcInputs({ ...ctx, benefits: { inn: {}, oon: {} } }, {}).missing, ['Deductible', 'OOP max', 'Coinsurance']);
});

// ---------- visit pre-check and last updated ----------
test('reportVisits reads Appointment Pre-Check Complete by header; precheckItems lets a tick win over the report', () => {
  const hdr = ['Per Nbr', 'Appt Date', 'Beg Tm', 'Event', 'Loc Name', 'Appointment Pre-Check Complete', 'MRN', 'Payer Name (Cht Guar)'];
  const rows = [hdr, ['9201', '10/01/2026', '9:00 AM', 'Old visit', 'Harbor Point', 'Yes', '720001', 'Acme'],
    ['9201', '11/06/2026', '1:30 PM', 'Follow-up', 'Harbor Point', '', '720001', 'Acme'],
    ['9201', '11/05/2026', '9:00 AM', 'Follow-up', 'Harbor Point', 'Yes', '720001', 'Acme'],
    ['9202', '11/05/2026', '10:00 AM', 'Intake', 'Mill Street', 'No', '720002', 'Acme']];
  const det = C.detectHeader(rows);
  const v = REC.reportVisits(rows, det, null, '10/30/2026');
  assert.deepEqual(v.get('9201').map((x) => `${x.date} ${x.time} ${x.precheck}`), ['11/05/2026 9:00 AM Yes', '11/06/2026 1:30 PM '], 'from the as-of day, soonest first');
  assert.equal(v.get('9202')[0].precheck, 'No');
  assert.equal(v.get('9202')[0].hasPrecheck, true);
  const noCol = REC.reportVisits([hdr.slice(0, 5), ['9203', '11/05/2026', '9:00 AM', 'Intake', 'Mill Street']], C.detectHeader([hdr.slice(0, 5), ['9203', '11/05/2026', '9:00 AM', 'Intake', 'Mill Street']]), null, null);
  assert.equal(noCol.get('9203')[0].hasPrecheck, false);
  const first = v.get('9201')[0];
  let items = REC.precheckItems(first, '720001', []);
  assert.deepEqual(items.map((i) => [i.id, i.done, i.source]), [['precheck', true, 'NextGen report'], ['nextgen', false, '']], 'two items: no cost-estimate item');
  const t = { key: '720001', lasfir: 'SYNPAT', date: first.date, time: first.time, by: 'Jordan Lee' };
  let saved = REC.setPrecheck([], { ...t, item: 'precheck', done: false, when: '10/30/2026 9:00 AM' });
  saved = REC.setPrecheck(saved, { ...t, item: 'nextgen', done: true, when: '10/30/2026 9:01 AM' });
  saved = REC.setPrecheck(saved, { ...t, item: 'nextgen', done: true, when: '10/30/2026 9:02 AM', by: 'Sam Ortiz' });
  assert.equal(saved.length, 2, 'one row per visit and item');
  assert.deepEqual([saved[0].LASFIR, saved[0].By], ['SYNPAT', 'Jordan Lee'], 'every row carries the LASFIR and the admin');
  items = REC.precheckItems(first, '720001', saved);
  assert.deepEqual(items.map((i) => [i.id, i.done, i.source]), [['precheck', false, 'Marked here'], ['nextgen', true, 'Marked here']]);
  assert.deepEqual([items[1].when, items[1].by], ['10/30/2026 9:02 AM', 'Sam Ortiz']);
  assert.throws(() => REC.setPrecheck([], { ...t, item: 'nope' }), /Unknown pre-check item/);
  assert.throws(() => REC.setPrecheck([], { ...t, by: '', item: 'precheck', done: true }), /Pick your name/);
  // Notes: one list per patient (not per visit), every note kept, newest first, with the admin; ticks are not notes.
  const pt = { key: '720001', lasfir: 'SYNPAT', by: 'Jordan Lee' };
  saved = REC.addPatientNote(saved, { ...pt, note: '  Called,   left message ', when: '10/30/2026 9:05 AM' });
  saved = REC.addPatientNote(saved, { ...pt, by: 'Sam Ortiz', note: 'Patient called back', when: '10/30/2026 9:30 AM' });
  // A note saved per visit by v0.20 still shows, with its visit.
  saved = saved.concat({ Key: '720001|11/05/2026|9:00 AM', MRN: '720001', LASFIR: 'SYNPAT', 'Appt Date': '11/05/2026', Time: '9:00 AM', Item: 'note', Done: '', Note: 'Old visit note', By: 'Alex Kim', When: '10/30/2026 8:00 AM' });
  assert.deepEqual(REC.patientNotes('720001', saved), [
    { note: 'Old visit note', by: 'Alex Kim', when: '10/30/2026 8:00 AM', visit: '11/05/2026 9:00 AM' },
    { note: 'Patient called back', by: 'Sam Ortiz', when: '10/30/2026 9:30 AM', visit: '' },
    { note: 'Called, left message', by: 'Jordan Lee', when: '10/30/2026 9:05 AM', visit: '' }]);
  const noteRow = saved.find((r) => r.Note === 'Patient called back');
  assert.deepEqual([noteRow.Key, noteRow.MRN, noteRow.LASFIR, noteRow['Appt Date'], noteRow.Time], ['720001', '720001', 'SYNPAT', '', ''], 'a patient note has no visit');
  assert.deepEqual(REC.patientNotes('799999', saved), []);
  assert.equal(REC.precheckItems(first, '720001', saved).length, 2, 'notes are not checklist items');
  assert.throws(() => REC.addPatientNote(saved, { ...pt, note: '   ' }), /empty/);
  assert.throws(() => REC.addPatientNote(saved, { ...pt, by: '', note: 'x' }), /Pick your name/);
  assert.equal(REC.addPatientNote([], { ...pt, note: 'x'.repeat(900) })[0].Note.length, REC.NOTE_MAX);
  assert.deepEqual(REC.PRECHECK_COLS, ['Key', 'MRN', 'LASFIR', 'Appt Date', 'Time', 'Item', 'Done', 'Note', 'By', 'When']);
  for (const c of REC.PRECHECK_COLS) assert.doesNotMatch(c, /first|last name|dob|birth/i);
  assert.deepEqual(REC.adminNames(' Jordan Lee; Sam Ortiz ;;Jordan Lee\nAlex Kim'), ['Jordan Lee', 'Sam Ortiz', 'Alex Kim']);
  assert.equal(REC.PRECHECK_CONFIG_ROWS[0].setting, 'Admin names');
  const old = REC.setPrecheck(saved, { ...t, date: '10/01/2026', time: '9:00 AM', item: 'precheck', done: true, when: '' });
  const kept = REC.prunePrecheck(old, '10/30/2026');
  assert.equal(kept.length, old.length - 1, 'ticks for past visits are dropped');
  assert.equal(kept.filter((r) => r.Item === 'note').length, 3, 'notes are the patient\'s: never pruned, even one saved against a past visit');
  assert.equal(REC.prunePrecheck(old.concat({ ...noteRow, 'Appt Date': '09/01/2026' }), '10/30/2026').filter((r) => r.Item === 'note').length, 4);
});

test('freshness: newest source time; fresh within a day, aging after, stale after 7 days', () => {
  const now = new Date('2026-10-30T18:00:00Z');
  const h = (n) => new Date(now.getTime() - n * 36e5).toISOString();
  assert.equal(REC.freshness({}, now).level, 'none');
  const f = REC.freshness({ report: h(2), sohar: h(5), flow: h(1) }, now);
  assert.equal(f.level, 'fresh');
  assert.equal(f.newest, Date.parse(h(1)));
  assert.equal(REC.freshness({ report: h(2), sohar: h(50) }, now).level, 'aging');
  assert.equal(REC.freshness({ report: h(2), sohar: h(24 * 8) }, now).level, 'stale');
  assert.equal(REC.freshness({ report: h(2) }, now).level, 'stale', 'no Sohar results at all');
  assert.equal(REC.freshness({ report: '2026-10-30 16:00 UTC', sohar: h(3) }, now).level, 'fresh', 'the Config UTC stamp is read');
});

test('guarantor text always names the patient by LASFIR', () => {
  const pt = { firstName: 'Ada', lastName: 'Okafor' };
  const c = REC.guarantorCheck({ patient: pt, guarantor: pt, guarantorSelf: true }, { firstName: 'Ben', lastName: 'Okafor' });
  assert.match(c.text, /OKAADA, the patient/);
  assert.match(REC.guarantorCheck({ patient: pt, guarantor: pt, guarantorSelf: true }, pt).text, /match \(OKAADA, the patient\)/);
  assert.match(REC.guarantorCheck({ patient: pt, guarantor: { firstName: 'Cy', lastName: 'Moss' } }, { firstName: 'Di', lastName: 'Park' }).text, /patient OKAADA/);
});

// ---------- Sohar Flow import (v0.21; v0.31 columns) ----------
test('Flow import: the header is the production Flow\'s CSV template (v27), exactly; no tin', () => {
  // docs/ARCHITECTURE.md section 5, pinned to the production template's input (template-v27.yaml): never rename or reorder.
  const template = 'id,patient.firstName,patient.lastName,patient.dateOfBirth,patient.state,placeOfServiceCode,specialtyCode,chartPayerId,chartMemberId,providerGroupName,estimate,rateTable';
  assert.equal(REC.FLOW_IMPORT_COLS.join(','), template);
  assert.equal(C.toCSV(REC.FLOW_IMPORT_COLS, []).split(/\r?\n/)[0], template);
  assert.match(REC.flowImportFileName(new Date(2026, 9, 30)), /^Sohar Ready Import - Flow 2026-10-30\.csv$/);
  assert.deepEqual(REC.FLOW_IMPORT_CONFIG_ROWS, [], 'the TIN and provider group rows are gone');
  assert.deepEqual(REC.PROVIDER_GROUPS, { CA: 'MINDFUL HEALTH SOLUTIONS A PROFESSIONAL MEDICAL CORPORATION', TX: 'Mindful Health Solutions TX PA', WA: 'Mindful Health Solutions Wa Pc' });
});

test('Flow import rows: chart policy only for verification patients; provider group by state; estimate and rateTable', () => {
  const p = (o) => ({ 'In Scope': 'Yes', 'First Name': 'Synth', 'Last Name': 'Patient', DOB: '01/02/1980', State: 'CA', POS: '10', Route: 'verification',
    'Sohar payerId': '60054', 'Member ID': 'SYN100', x: { verifyRows: [{ payerId: '60054', memberId: 'SYN100' }] }, ...o });
  const pts = [
    p({ MRN: '730001' }),
    p({ MRN: '730002', State: 'TX', x: { verifyRows: [{ payerId: '60054', memberId: 'SYN200', reverifySkip: true }, { payerId: '94036', memberId: 'SYN201' }] } }),
    p({ MRN: '730003', Route: 'discovery', State: 'wa', 'Sohar payerId': '99999', 'Member ID': 'GUESS1', x: {} }),
    p({ MRN: '730004', x: { verifyRows: [{ payerId: '60054', memberId: 'SYN400', reverifySkip: true }] } }),
    p({ MRN: '730005', DOB: '' }),
    p({ MRN: '730006', 'In Scope': 'No' }),
    p({ MRN: '730008', Route: 'discovery', x: { excluded: { line: 'TMS treatment', reason: 'Config Include in import: TMS treatment is Off' } } }),
    p({ MRN: '730009', State: 'GA' }),
  ];
  const rates = { 60054: [{ code: '90792', contractedRate: 1.11 }, { code: '99214', contractedRate: 2.11 }, { code: '99215', contractedRate: 3.11 }] };
  const out = REC.flowImportRows(pts, { specialty: 'PSY', estimateOf: (pt) => (pt.MRN === '730001' ? { due: true } : pt.MRN === '730003' ? 'yes' : false), rateTableOf: (pt) => (pt.MRN === '730001' ? rates : {}) });
  assert.deepEqual(out.rows.map((r) => r.id), ['730001', '730002', '730003', '730009']);
  assert.deepEqual(out.rows[0], { id: '730001', 'patient.firstName': 'Synth', 'patient.lastName': 'Patient', 'patient.dateOfBirth': '01/02/1980', 'patient.state': 'CA',
    placeOfServiceCode: '10', specialtyCode: 'PSY', chartPayerId: '60054', chartMemberId: 'SYN100',
    providerGroupName: 'MINDFUL HEALTH SOLUTIONS A PROFESSIONAL MEDICAL CORPORATION', estimate: 'yes', rateTable: JSON.stringify(rates) });
  assert.deepEqual(Object.keys(out.rows[0]), REC.FLOW_IMPORT_COLS);
  assert.deepEqual(JSON.parse(out.rows[0].rateTable)['60054'][1], { code: '99214', contractedRate: 2.11 }, 'a JSON string the Flow can $eval');
  assert.deepEqual([out.rows[1].chartPayerId, out.rows[1].chartMemberId, out.rows[1].providerGroupName, out.rows[1].estimate, out.rows[1].rateTable],
    ['94036', 'SYN201', 'Mindful Health Solutions TX PA', 'no', ''], 'the first insurance still due; no estimate, no rates');
  assert.deepEqual([out.rows[2].chartPayerId, out.rows[2].chartMemberId, out.rows[2].providerGroupName, out.rows[2].rateTable],
    ['', '', 'Mindful Health Solutions Wa Pc', ''], 'discovery patients: a guessed payer is never sent as the chart policy; an empty rate table is blank');
  assert.equal(out.rows[3].providerGroupName, '', 'no MHS group in that state');
  assert.equal(out.estimates, 2);
  assert.deepEqual(out.skipped, [{ key: '730004', why: 'already covered, or too early to re-verify' }, { key: '730005', why: 'no name, DOB or state' },
    { key: '730008', why: 'Config Include in import: TMS treatment is Off' }]);
  assert.equal(REC.flowImportRows([p({ MRN: '730007' })], {}).rows[0].specialtyCode, 'PSY');
  assert.equal(REC.flowImportRows([p({ MRN: '730007' })], {}).rows[0].estimate, 'no');
});

test('Flow Results v0.31 columns: Source, Policy Key, Reset Basis, Est 90792 / 99214 / 99215; old files still read', () => {
  for (const c of ['Source', 'Policy Key', 'Reset Basis', 'Est 90792', 'Est 99214', 'Est 99215']) assert.ok(REC.FLOW_COLS.includes(c), c);
  const old = REC.FLOW_COLS.filter((c) => !['Source', 'Policy Key', 'Reset Basis', 'Est 90792', 'Est 99214', 'Est 99215'].includes(c));
  const recs = REC.flowRecords([old, old.map((c) => (c === 'MRN' ? '700009' : c === 'Key' ? 'k9' : ''))]);
  assert.deepEqual([recs[0].Source, recs[0]['Est 99214'], recs[0].MRN], ['Flow', '', '700009']);
  assert.equal(REC.resetBasisFor([{ 'Policy Key': 'AC1#1234', 'Reset Basis': 'Unknown', 'Last Updated': '2026-01-02' }, { 'Policy Key': 'AC1#1234', 'Reset Basis': 'Plan', 'Last Updated': '2026-08-02' }], 'AC1#1234'), 'Plan');
});

// ---------- Verification History (v0.31) ----------
test('policyKey: payerId + # + the member ID\'s last 4; never the full member ID', () => {
  assert.equal(REC.policyKey('94036', 'syn-00012345'), '94036#2345');
  assert.equal(REC.policyKey('94036', ''), '94036#');
  assert.equal(REC.policyKey('', 'X1'), '');
  assert.deepEqual(REC.VERIFICATION_HISTORY_COLS.slice(-4), ['Policy Key', 'Projected Patient Responsibility', 'Estimate Note', 'Source']);
  assert.ok(!REC.VERIFICATION_HISTORY_COLS.some((c) => /member id|first name|last name|dob/i.test(c)), 'no member IDs or names on the sheet');
});

test('Verification History: append-only, a re-read updates fields only, old policies stay, sorted by Pulled At', () => {
  const r = (o) => row({ MRN: '750001', LASFIR: 'SYNPAT', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar payerId': '60054', 'Sohar Payer': 'Synth Mutual',
    'Sohar Member ID': 'SYN-0001', Match: 'MRN', 'Sohar Date': '2026-10-01T17:00:00Z', 'Next Location': 'Harbor Point', ...o });
  const a = r({ 'Sohar Ref': 'v-a', 'Deductible Remaining': '500' });
  const first = REC.historyFromResults([a, r({ 'Sohar Ref': 'v-n', Match: 'Name only. Confirm.' }), r({ 'Sohar Ref': '' })]);
  assert.equal(first.length, 1, 'name-only and ref-less rows are left out');
  assert.deepEqual([first[0]['Policy Key'], first[0].Source, first[0].Clinic, first[0]['Sohar ID'], first[0]['Pulled At']], ['60054#0001', 'Import', 'Harbor Point', 'v-a', '2026-10-01T17:00:00.000Z']);
  assert.deepEqual(Object.keys(first[0]), REC.VERIFICATION_HISTORY_COLS);
  assert.ok(!JSON.stringify(first).includes('SYN-0001'), 'the member ID never goes on the sheet');
  assert.equal(REC.historyFromResults([r({ 'Sohar Ref': 'v-d', Match: 'Member ID' })])[0].Source, 'Dashboard');
  // A discovery finds a new policy later: the old one stays.
  const b = r({ 'Sohar Ref': 'v-b', 'Sohar payerId': '94036', 'Sohar Payer': 'Other Synth', 'Sohar Member ID': 'NEW-9999', 'Sohar Date': '2026-11-02T17:00:00Z' });
  let m = REC.mergeHistory(first, REC.historyFromResults([b]));
  assert.deepEqual([m.added, m.updated, m.rows.length], [1, 0, 2]);
  // An earlier check read later sorts into place; re-reading v-a updates a field and keeps its Source.
  const c = r({ 'Sohar Ref': 'v-c', 'Sohar Date': '2026-09-01T17:00:00Z' });
  m = REC.mergeHistory(m.rows, [...REC.historyFromResults([c]), ...REC.historyFromResults([{ ...a, 'Deductible Remaining': '450' }], 'Dashboard')]);
  assert.deepEqual(m.rows.map((x) => x['Sohar ID']), ['v-c', 'v-a', 'v-b']);
  assert.deepEqual([m.added, m.updated], [1, 1]);
  const va = m.rows.find((x) => x['Sohar ID'] === 'v-a');
  assert.deepEqual([va['Deductible Remaining'], va.Source], ['450', 'Import']);
  assert.equal(REC.mergeHistory(m.rows, []).rows.length, 3, 'nothing is ever dropped');
  // Flow results become Flow rows.
  const flow = REC.historyFromFlow([{ MRN: '750001', LASFIR: 'SYNPAT', Location: 'Harbor Point', 'Estimate Policy': 'NextGen', 'NextGen Policy Status': 'complete.eligible',
    'NextGen Verification Link': 'https://app.soharhealth.com/verifications/v-f', 'Policy Key': '60054#0001', 'Sohar Est Patient': '45.11', 'Last Updated': '2026-11-05T17:00:00Z' }, { MRN: '750001' }]);
  assert.deepEqual([flow.length, flow[0].Source, flow[0]['Sohar ID'], flow[0]['Sohar payerId'], flow[0]['Projected Patient Responsibility']], [1, 'Flow', 'v-f', '60054', '45.11']);
});

test('policyTimeline: one timeline per policy, newest first, merged with the rotation\'s Patient History', () => {
  const h = (id, at, key, payer, src = 'Import') => ({ ...Object.fromEntries(REC.VERIFICATION_HISTORY_COLS.map((c) => [c, ''])), MRN: '750001', 'Sohar ID': id, 'Pulled At': at,
    'Policy Key': key, 'Sohar Payer': payer, 'Sohar Status': 'complete.eligible', Kind: 'Verification', Source: src });
  const vh = [h('v-1', '2026-09-01T17:00:00Z', '60054#0001', 'Synth Mutual'), h('v-2', '2026-10-01T17:00:00Z', '60054#0001', 'Synth Mutual'),
    h('v-3', '2026-11-02T17:00:00Z', '94036#9999', 'Other Synth', 'Flow'), { ...h('v-x', '2026-11-03', 'X#1', 'Else'), MRN: '750002' }];
  const ph = [{ 'Row Key': 'v-2', MRN: '0750001', When: '2026-10-01', Kind: 'Verification', Detail: 'complete.eligible · Synth Mutual', Source: 'Rotation', Ref: 'v-2' },
    { 'Row Key': 'v-4', MRN: '750001', When: '2026-10-15T17:00:00Z', Kind: 'Verification', Detail: 'complete.ineligible · Synth Mutual', Source: 'Rotation', Ref: 'v-4' },
    { 'Row Key': 'w-5', MRN: '750001', When: '2026-08-01T17:00:00Z', Kind: 'Discovery', Detail: 'No coverage found', Source: 'Webhook', Ref: 'w-5' }];
  const t = REC.policyTimeline(vh, ph, '0750001');
  assert.deepEqual(t.map((g) => g.policyKey), ['94036#9999', '60054#0001', '']);
  assert.deepEqual(t[1].rows.map((r) => r['Sohar ID']), ['v-4', 'v-2', 'v-1'], 'newest first; v-2 once; the rotation row joins by payer');
  assert.deepEqual([t[1].payerId, t[1].last4, t[1].payer], ['60054', '0001', 'Synth Mutual']);
  assert.equal(t[2].rows[0].Source, 'Webhook');
});


// ---------- patient view (v0.22): plans, milestones, what changed, projection, action items ----------
const ver = (o) => row({ 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', MRN: '740001', 'Per Nbr': '940001', ...o });
test('plansOf groups results by payer and member ID; the lead result\'s plan is current', () => {
  const now = ver({ 'Sohar Payer': 'Synth Mutual', 'Sohar Member ID': 'SYN-A1', 'Sohar Date': '2026-10-29T16:40:00Z' });
  const disc = row({ 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', MRN: '740001', 'Sohar Payer': 'Synth Mutual', 'Sohar Date': '2026-10-29T16:31:00Z' });
  const old = ver({ 'Sohar Payer': 'Old Synth HMO', 'Sohar Member ID': 'OLD-9', 'Sohar Status': 'error.member.inactive', 'Sohar Date': '2026-09-03T18:12:00Z', Action: 'Superseded by a later result' });
  const plans = REC.plansOf([old, disc, now], now);
  assert.equal(plans.length, 2);
  assert.equal(plans[0].current, true);
  assert.deepEqual(plans[0].results, [now, disc], 'a discovery with no member ID joins the same payer');
  assert.equal(plans[1].payer, 'Old Synth HMO');
  assert.equal(plans[1].superseded, true);
  const marks = REC.milestones([old, disc, now], plans);
  assert.deepEqual(marks.get(disc), ['New plan, first seen']);
  assert.deepEqual(REC.sinceLast(plans), { a: 'Status changed', sub: 'Status Coverage found → complete.eligible since Oct 29.' });
  assert.equal(REC.sinceLast([{ ...plans[0], results: [now] }, plans[1]]).a, 'New plan');
});

test('milestones mark the first result each benefit state was seen on', () => {
  const a = ver({ 'Sohar Date': '2026-09-22T21:10:00Z', 'Deductible Status': 'Met', 'OOP Status': 'Not met' });
  const b = ver({ 'Sohar Date': '2026-10-20T16:30:00Z', 'Deductible Status': 'Met', 'OOP Status': 'Approaching' });
  const c = ver({ 'Sohar Date': '2026-10-27T17:05:00Z', 'Deductible Status': 'Met', 'OOP Status': 'Approaching' });
  const m = REC.milestones([c, b, a], REC.plansOf([c, b, a], c));
  assert.deepEqual(m.get(a), ['Deductible met, first seen']);
  assert.deepEqual(m.get(b), ['Approaching OOP max, first seen']);
  assert.equal(m.get(c), undefined);
});

const fv = (date, est, o) => ({ Key: `k-${date}`, MRN: '740001', 'Appt Date': date, Time: '11:00 AM', Event: 'Med management', 'Run Status': 'COMPLETE', 'Sohar Est Patient': est, File: 'Flow Results.csv', ...o });
test('projection walks the booked visits and marks where the OOP max is met', () => {
  const fin = { Deductible: '1000', 'Deductible Remaining': '0', 'OOP Max': '3000', 'OOP Remaining': '290', Copay: '' };
  const fl = ['11/03/2026', '11/10/2026', '11/17/2026', '11/24/2026', '12/01/2026'].map((d) => fv(d, '96.40'));
  const p = REC.projection(fin, fl);
  assert.equal(p.oop.at, fl[3], 'met at visit 4 of 5');
  assert.deepEqual(p.rows.map((r) => r.applied), [96.4, 96.4, 96.4, 0.8, 0]);
  assert.deepEqual(p.rows[3].marks, ['OOP max met at this visit']);
  assert.equal(p.ded.at, null, 'already met');
  assert.equal(REC.leadLine({ lead: ver({}), proj: p }).text, 'Approaching OOP max: $290.00 left. Met at the Nov 24 visit if visits go as booked (visit 4 of 5).');
});

test('projection: copays skip the deductible, and sandbox or missing estimates are left out', () => {
  const fin = { Deductible: '2500', 'Deductible Remaining': '2200', 'OOP Max': '6000', 'OOP Remaining': '5700', Copay: '55' };
  const fl = [fv('11/02/2026', '55'), fv('11/09/2026', '210', { File: 'Flow Results (SANDBOX).csv' }), fv('11/16/2026', '', { Event: 'TMS mapping' })];
  const p = REC.projection(fin, fl);
  assert.equal(p.ded.applied, 0, 'a $55 copay goes to the OOP max only');
  assert.equal(p.oop.applied, 55);
  assert.deepEqual(p.missing, [fl[1], fl[2]]);
  assert.equal(REC.projection({ Deductible: '', 'OOP Max': '' }, fl), null);
  assert.equal(REC.projection(null, fl), null);
});

test('actionItems: most urgent first; leadLine names the worst problem', () => {
  const now = new Date('2026-10-30T16:15:00Z'); // 9:15 AM Pacific; fixed in UTC so the day counts match in any time zone
  const fresh = REC.freshness({ report: '2026-10-22T14:05:00Z', sohar: '2026-10-29T21:15:00Z', flow: '' }, now);
  const visit = (date, done) => ({ date, time: '1:00 PM', items: [{ id: 'precheck', done }, { id: 'nextgen', done: false }] });
  const lead = row({ 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Payer': 'Synth North', MRN: '740002' });
  const items = REC.actionItems({ lead, fresh, checks: [{ severity: 'high' }], visits: [visit('11/06/2026', false), visit('11/20/2026', false), visit('11/04/2026', true)],
    flow: [fv('11/20/2026', '', { Event: 'Esketamine session' })], tasks: [{ Task: 'Add insurance', 'Due By': '11/05/2026' }], now });
  assert.deepEqual(items.map((i) => [i.tone, i.tag]), [['red', 'Fix now'], ['red', 'Fix now'], ['amber', 'Before visit'], ['amber', 'Pre-check'], ['amber', 'NextGen task'], ['grey', 'Later'], ['grey', 'Verify']]);
  assert.equal(items[0].text, 'Upload a new NextGen report. It is 8 days old.');
  assert.equal(items[2].text, 'Run verification on Synth North.');
  assert.equal(items[6].text, 'Check deductible, OOP max and status for Nov 20 Esketamine session');
  assert.equal(REC.leadLine({ lead, checks: [{ severity: 'high' }], fresh, next: '11/06/2026' }).text, 'Guarantor and subscriber are different people. Fix the guarantor in NextGen before the Nov 6 visit.');
  assert.equal(REC.leadLine({ lead, fresh: REC.freshness({ report: now, sohar: now }, now), next: '11/06/2026' }).text, 'Discovery found Synth North. Run verification to get benefits before the Nov 6 visit.');
  assert.equal(REC.leadLine({ lead: ver({ 'Work Queue': 'Clear' }) }).tone, 'green');
});

test('projection: a visit\'s code rules decide whether it counts toward the deductible (v0.23)', () => {
  const fin = { Deductible: '1000', 'Deductible Remaining': '500', 'OOP Max': '3000', 'OOP Remaining': '2000', Copay: '' };
  const fl = [fv('11/03/2026', '100', { 'Est Codes': '99214' }), fv('11/10/2026', '100', { 'Est Codes': '90836' })];
  const p = REC.projection(fin, fl, { dedApplies: (f) => (f['Est Codes'] === '99214' ? false : null) });
  assert.deepEqual(p.rows.map((r) => [r.toward, r.ruled]), [[false, true], [true, false]]);
  assert.deepEqual([p.ded.applied, p.oop.applied], [100, 200], 'copays and waived visits still count toward the OOP max');
});
