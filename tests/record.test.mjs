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
