import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const R = createRequire(import.meta.url)('../rules.js');

// Synthetic data only.
const TODAY = new Date(2026, 9, 15); // Thu 10/15/2026 local
const CFG = { today: TODAY, sfhpProviders: [], reverifyDays: 7 };

function pat(o = {}, x = {}) {
  return Object.assign({ 'Per Nbr': '1001', MRN: `M${(o && o['Per Nbr']) || '1001'}`, 'Last Name': 'Testerson', 'First Name': 'Alexa', DOB: '01/02/1980', State: 'CA', 'In Scope': 'Yes', 'Primary Payer': 'Blue Shield CA', 'Secondary Payer': '' }, o, {
    x: Object.assign({ finClass: 'Blue Shield', finClassLstEnc: 'Blue Shield', appointments: [], verifyRows: [] }, x),
  });
}
const appt = (o) => Object.assign({ date: '10/20/2026', time: '09:00', event: 'Med Management Telepsych', loc: 'San Rafael', rendering: 'Doe, Jamie MD', status: '', authCd: '', authCdLstEnc: '' }, o);
const fresh = (o = {}) => Object.assign({ 'Per Nbr': '1001', MRN: `M${o['Per Nbr'] || '1001'}`, 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-13T17:00:00Z', 'Sohar Payer': 'Blue Shield of California', Plan: 'PPO' }, o);
const codes = (fl) => fl.map((f) => f.code);

test('PILLS: shape and required codes', () => {
  const want = ['MEDICAL', 'MEDICAL_NEW', 'SFHP_PROVIDER', 'KAISER_MM_NOAUTH', 'NP_TMS', 'VIP', 'QMB', 'SECONDARY', 'REVERIFY', 'DUAL', 'KAISER_MM'];
  for (const c of want) assert.ok(R.PILLS.find((p) => p.code === c), c);
  for (const p of R.PILLS) {
    assert.ok(['high', 'medium', 'info'].includes(p.severity));
    assert.ok(['red', 'amber', 'blue', 'grey'].includes(p.color));
    assert.ok(p.label);
  }
  assert.equal(R.REVERIFY_URL, 'https://app.soharhealth.com/create?type=verification');
});

test('classifiers: payers', () => {
  assert.ok(R.isKaiser('Kaiser Foundation Northern California'));
  assert.ok(R.isKaiser('Kaiser Medicare Northern CA'));
  assert.ok(!R.isKaiser('Anthem Blue Cross'));
  assert.ok(R.isVA('Veterans Affairs Nor Cal'));
  assert.ok(R.isVA('Veteran Affairs'));
  assert.ok(R.isVA('VA'));
  assert.ok(!R.isVA('Nevada Health'));
  assert.ok(R.isVA('TriWest')); // TriWest (VA community care) shares the VA bucket
  assert.ok(R.isSFHP('San Francisco Health Plan'));
  assert.ok(R.isSFHP('SFHP Medi-Cal'));
  assert.ok(R.isQMB('QMB Noridian Southern CA'));
  assert.ok(!R.isQMB('Noridian Northern CA'));
  assert.ok(R.isMediCal('Health Net Medi-Cal'));
  assert.ok(R.isMediCal('MEDI CAL'));
  assert.ok(R.isMediCal('California Medicaid'));
  assert.ok(!R.isMediCal('Medicare Advantage'));
  assert.ok(!R.isMediCal('Noridian Southern CA'));
  assert.ok(R.isCarelon('Carelon Behavioral Health'));
  assert.ok(R.isCarelon('Beacon Health Options'));
  assert.ok(R.isCarelon('Beacon'));
  assert.ok(!R.isCarelon('Cigna'));
  assert.ok(R.isUMR('UMR'));
  assert.ok(!R.isUMR('Humana Medicare'));
  assert.ok(R.isVIP('VIP'));
  assert.ok(R.isVIP('Partial VIP'));
  assert.ok(!R.isVIP('Vipassana Health'));
});

test('classifiers: rendering credentials', () => {
  assert.ok(R.isNP('Rivera, Sam PMHNP'));
  assert.ok(R.isNP('Rivera, Sam NP'));
  assert.ok(R.isNP('Rivera, Sam FNP'));
  assert.ok(R.isNP('Rivera, Sam APRN'));
  assert.ok(R.isNP('Rivera, Sam PMHNP-BC'));
  assert.ok(!R.isNP('Doe, Jamie MD'));
  assert.ok(!R.isNP('Doe, Jamie DO'));
  assert.ok(!R.isNP('Npolo, Chris MD'));
  assert.ok(!R.isNP(''));
});

test('classifiers: events', () => {
  for (const e of ['TMS Treatment Unilateral', 'Telepsych TMS F/U', 'TMS Theta Tx Bilateral', 'TMS Follow Up', 'MT Redo', 'MT MagV']) {
    assert.ok(R.isTMSEvent(e), e);
    assert.ok(!R.isMedMgmtEvent(e), e);
  }
  for (const e of ['Med Management Telepsych', 'Med Management In-Office', 'Telepsych', 'Tel New Patient - Mature', 'IP New Patient - Mature', 'Transition Telepsych Patient']) {
    assert.ok(R.isMedMgmtEvent(e), e);
    assert.ok(!R.isTMSEvent(e), e);
  }
  for (const e of ['Esketamine Follow Up', 'Esketamine Tx Start', 'IV Ketamine Follow Up', 'IV Ketamine Consult Tele', 'Meeting', 'Provider Training',
    'Daily On Call', 'Peer To Peer', 'ADHD Testing Review Tele', 'Psychotherapy Individual 60', '']) {
    assert.ok(!R.isMedMgmtEvent(e), e);
    assert.ok(!R.isTMSEvent(e), e);
  }
});

test('R5: VA med management telepsych without auth flags', () => {
  const p = pat({ 'Primary Payer': 'Veterans Affairs Nor Cal' }, { finClass: 'Veteran Affairs', finClassLstEnc: 'Veteran Affairs', appointments: [appt({ event: 'Telepsych' })] });
  const fl = R.patientFlags(p, [fresh()], CFG);
  const f = fl.find((x) => x.code === 'KAISER_MM_NOAUTH');
  assert.ok(f);
  assert.equal(f.severity, 'high');
  assert.equal(f.appt.date, '10/20/2026');
  assert.match(f.detail, /^Veterans Affairs Nor Cal med management/);
  assert.ok(codes(fl).includes('KAISER_MM'));
});

test('R5: Kaiser auth on Lst Enc suppresses; TMS and Esketamine never flag; past appts ignored', () => {
  const base = { 'Primary Payer': 'Kaiser Foundation Northern California' };
  const withAuth = pat(base, { finClass: 'Kaiser', appointments: [appt({ authCdLstEnc: 'A123' })] });
  let fl = R.patientFlags(withAuth, [fresh()], CFG);
  assert.ok(!codes(fl).includes('KAISER_MM_NOAUTH'));
  assert.ok(codes(fl).includes('KAISER_MM'));
  const guarAuth = pat(base, { finClass: 'Kaiser', appointments: [appt({ authCd: 'A9' })] });
  assert.ok(!codes(R.patientFlags(guarAuth, [fresh()], CFG)).includes('KAISER_MM_NOAUTH'));
  const tms = pat(base, { finClass: 'Kaiser', appointments: [appt({ event: 'TMS Treatment Unilateral' }), appt({ event: 'Esketamine Follow Up' }), appt({ event: 'IV Ketamine Follow Up' })] });
  fl = R.patientFlags(tms, [fresh()], CFG);
  assert.ok(!codes(fl).includes('KAISER_MM_NOAUTH'));
  assert.ok(!codes(fl).includes('KAISER_MM'));
  const past = pat(base, { finClass: 'Kaiser', appointments: [appt({ date: '10/01/2026' })] });
  assert.ok(!codes(R.patientFlags(past, [], CFG)).includes('KAISER_MM_NOAUTH'));
  const cancelled = pat(base, { finClass: 'Kaiser', appointments: [appt({ status: 'Cancelled' })] });
  assert.ok(!codes(R.patientFlags(cancelled, [fresh()], CFG)).includes('KAISER_MM_NOAUTH'));
});

test('R5: one flag per appointment; Fin Class alone triggers; falls back to x auth when appt has no auth keys', () => {
  const p = pat({ 'Primary Payer': 'Some Plan' }, { finClass: 'Kaiser', appointments: [appt({}), appt({ date: '10/27/2026' })] });
  assert.equal(R.patientFlags(p, [fresh()], CFG).filter((f) => f.code === 'KAISER_MM_NOAUTH').length, 2);
  const q = pat({ 'Primary Payer': 'Kaiser Foundation Northern California' }, { finClass: 'Kaiser', authCd: 'Z1', appointments: [{ date: '10/20/2026', event: 'Telepsych' }] });
  assert.ok(!codes(R.patientFlags(q, [fresh()], CFG)).includes('KAISER_MM_NOAUTH'));
});

test('R6: Medi-Cal not accepted; Kaiser and SFHP Medi-Cal accepted; QMB is not Medi-Cal', () => {
  const bad = pat({ 'Primary Payer': 'Health Net Medi-Cal' }, { finClass: 'Contracted Commercial', finClassLstEnc: 'Contracted Commercial', appointments: [appt({})] });
  let fl = R.patientFlags(bad, [fresh()], CFG);
  const m = fl.find((f) => f.code === 'MEDICAL');
  assert.ok(m);
  assert.equal(m.severity, 'high');
  assert.ok(codes(fl).includes('MEDICAL_NEW'));

  const kaiser = pat({ 'Primary Payer': 'Kaiser Medi-Cal' }, { finClass: 'Kaiser', finClassLstEnc: 'Kaiser', appointments: [appt({ event: 'TMS Follow Up' })] });
  assert.ok(!codes(R.patientFlags(kaiser, [fresh()], CFG)).includes('MEDICAL'));
  const kaiserPlan = pat({ 'Primary Payer': 'Kaiser Foundation Northern California' }, { finClass: 'Kaiser' });
  assert.ok(!codes(R.patientFlags(kaiserPlan, [fresh({ 'Sohar Payer': 'Kaiser Permanente', Plan: 'Medi-Cal Managed Care' })], CFG)).includes('MEDICAL'));

  const sfhp = pat({ 'Primary Payer': 'San Francisco Health Plan' }, { finClass: 'Contracted Commercial', appointments: [appt({})] });
  fl = R.patientFlags(sfhp, [fresh({ 'Sohar Payer': 'San Francisco Health Plan', Plan: 'Medi-Cal' })], CFG);
  assert.ok(!codes(fl).includes('MEDICAL'));
  assert.ok(codes(fl).includes('SFHP_PROVIDER'));

  const qmb = pat({ 'Primary Payer': 'QMB Noridian Southern CA' }, { finClass: 'Medicare', finClassLstEnc: 'Medicare', appointments: [appt({ event: 'TMS Follow Up' })] });
  fl = R.patientFlags(qmb, [fresh()], CFG);
  assert.ok(!codes(fl).includes('MEDICAL'));
  const q = fl.find((f) => f.code === 'QMB');
  assert.ok(q);
  assert.match(q.detail, /cannot be billed to the patient/);
});

test('R6: Medi-Cal from Sohar verification or discovery; always high; NEW unless last visit had non-accepted Medi-Cal', () => {
  const p = pat({}, { finClass: 'Blue Shield', finClassLstEnc: 'Blue Shield' });
  let fl = R.patientFlags(p, [fresh(), { 'Per Nbr': '1001', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Payer': 'Anthem Blue Cross Medi-Cal', 'Sohar Date': '2026-10-13' }], CFG);
  assert.ok(codes(fl).includes('MEDICAL'));
  assert.match(fl.find((f) => f.code === 'MEDICAL').detail, /Sohar discovery/);
  const old = pat({ 'Primary Payer': 'Health Net Medi-Cal' }, { finClass: 'Contracted Commercial', finClassLstEnc: 'Medi-Cal' });
  fl = R.patientFlags(old, [fresh()], CFG);
  assert.equal(fl.find((f) => f.code === 'MEDICAL').severity, 'high');
  assert.ok(!codes(fl).includes('MEDICAL_NEW'));
  // Kaiser Medi-Cal (accepted) last visit, now Health Net Medi-Cal (not accepted): newly on a plan MHS does not take.
  const moved = pat({ 'Primary Payer': 'Health Net Medi-Cal' }, { finClass: 'Contracted Commercial', finClassLstEnc: 'Kaiser Medi-Cal' });
  fl = R.patientFlags(moved, [fresh()], CFG);
  assert.equal(fl.find((f) => f.code === 'MEDICAL').severity, 'high');
  assert.match(fl.find((f) => f.code === 'MEDICAL_NEW').detail, /accepted Medi-Cal plan/);
});

test('R6: SFHP provider list', () => {
  const p = pat({ 'Primary Payer': 'San Francisco Health Plan' }, { finClass: 'Contracted Commercial', appointments: [appt({ rendering: 'Doe, Jamie MD' }), appt({ date: '10/22/2026', rendering: 'Rivera, Sam PMHNP' })] });
  let fl = R.patientFlags(p, [fresh()], Object.assign({}, CFG, { sfhpProviders: ['Doe'] }));
  const s = fl.filter((f) => f.code === 'SFHP_PROVIDER');
  assert.equal(s.length, 1);
  assert.equal(s[0].appt.rendering, 'Rivera, Sam PMHNP');
  fl = R.patientFlags(p, [fresh()], Object.assign({}, CFG, { sfhpProviders: ['Doe, Jamie', 'Rivera'] }));
  assert.ok(!codes(fl).includes('SFHP_PROVIDER'));
  fl = R.patientFlags(p, [fresh()], CFG);
  const e = fl.filter((f) => f.code === 'SFHP_PROVIDER');
  assert.equal(e.length, 1);
  assert.match(e[0].detail, /confirm provider is approved/);
});

test('R7: UMR/Carelon NP rendering for TMS', () => {
  const umr = pat({ 'Primary Payer': 'UMR' }, { finClass: 'Contracted Commercial', appointments: [appt({ event: 'TMS Treatment Unilateral', rendering: 'Rivera, Sam PMHNP' })] });
  const f = R.patientFlags(umr, [fresh()], CFG).find((x) => x.code === 'NP_TMS');
  assert.ok(f);
  assert.match(f.detail, /UMR/);
  const md = pat({ 'Primary Payer': 'UMR' }, { appointments: [appt({ event: 'TMS Treatment Unilateral', rendering: 'Doe, Jamie MD' })] });
  assert.ok(!codes(R.patientFlags(md, [fresh()], CFG)).includes('NP_TMS'));
  const esk = pat({ 'Primary Payer': 'Carelon Behavioral Health' }, { appointments: [appt({ event: 'Esketamine Follow Up', rendering: 'Rivera, Sam NP' })] });
  assert.ok(!codes(R.patientFlags(esk, [fresh()], CFG)).includes('NP_TMS'));
  const beacon = pat({ 'Primary Payer': 'Other Name' }, { finClass: 'Beacon', appointments: [appt({ event: 'MT MagV', rendering: 'Rivera, Sam FNP' })] });
  assert.match(R.patientFlags(beacon, [fresh()], CFG).find((x) => x.code === 'NP_TMS').detail, /Carelon/);
  const aetna = pat({ 'Primary Payer': 'Aetna' }, { finClass: 'Aetna', appointments: [appt({ event: 'TMS Follow Up', rendering: 'Rivera, Sam NP' })] });
  assert.ok(!codes(R.patientFlags(aetna, [fresh()], CFG)).includes('NP_TMS'));
});

test('R8: secondary may cover; stronger with Medicare', () => {
  const p = pat({ 'Secondary Payer': 'Aetna' });
  let f = R.patientFlags(p, [fresh({ Copay: '$40' })], CFG).find((x) => x.code === 'SECONDARY');
  assert.equal(f.severity, 'info');
  assert.match(f.detail, /Secondary Aetna may cover the deductible, copay or coinsurance/);
  assert.match(f.detail, /copay \$40/);
  const mc = pat({ 'Primary Payer': 'Noridian Southern CA', 'Secondary Payer': 'Blue Shield CA' }, { finClass: 'Medicare' });
  f = R.patientFlags(mc, [fresh()], CFG).find((x) => x.code === 'SECONDARY');
  assert.equal(f.severity, 'medium');
  assert.match(f.detail, /likely covers/);
  const qmbSec = pat({ 'Secondary Payer': 'QMB Noridian Northern CA' });
  const fl = R.patientFlags(qmbSec, [fresh()], CFG);
  assert.equal(fl.find((x) => x.code === 'SECONDARY').severity, 'medium');
  assert.ok(codes(fl).includes('QMB'));
  assert.ok(!codes(R.patientFlags(pat({ 'Secondary Payer': 'Self Pay' }), [fresh()], CFG)).includes('SECONDARY'));
  assert.ok(!codes(R.patientFlags(pat({}, { secPayer: '' }), [fresh()], CFG)).includes('SECONDARY'));
  assert.ok(codes(R.patientFlags(pat({}, { secPayer: 'Cigna' }), [fresh()], CFG)).includes('SECONDARY'));
});

test('R9: re-verify 8 and 7 days flag (approved 2026-10-20: due at Re-verify days), 6 does not, none flags, no upcoming visit no flag', () => {
  const p = pat({}, { appointments: [appt({ event: 'TMS Follow Up' })] });
  const at = (d) => [fresh({ 'Sohar Date': d })];
  let f = R.patientFlags(p, at('2026-10-07T20:00:00Z'), CFG).find((x) => x.code === 'REVERIFY');
  assert.ok(f);
  assert.equal(f.link, R.REVERIFY_URL);
  assert.match(f.detail, /8 days ago/);
  assert.match(R.patientFlags(p, at('2026-10-08'), CFG).find((x) => x.code === 'REVERIFY').detail, /7 days ago/);
  assert.ok(!codes(R.patientFlags(p, at('10/09/2026'), CFG)).includes('REVERIFY'));
  f = R.patientFlags(p, [], CFG).find((x) => x.code === 'REVERIFY');
  assert.match(f.detail, /No Sohar verification/);
  // discovery rows and "Not in Sohar file" don't count as a verification
  assert.ok(codes(R.patientFlags(p, [fresh({ 'Sent As': 'Discovery' }), fresh({ 'Sohar Status': 'Not in Sohar file' })], CFG)).includes('REVERIFY'));
  assert.ok(!codes(R.patientFlags(pat({}, { appointments: [] }), [], CFG)).includes('REVERIFY'));
  // fallback to Next Appt when no appointments array
  assert.ok(codes(R.patientFlags({ 'Per Nbr': '9', 'Next Appt': '10/30/2026' }, [], CFG)).includes('REVERIFY'));
});

test('R9 shares core.reverifyCheck: not eligible, new month, coverage end, per insurance sent', () => {
  const rv = (p, rows, cfg = CFG) => R.patientFlags(p, rows, cfg).find((x) => x.code === 'REVERIFY');
  const p = pat({}, { appointments: [appt({ date: '10/20/2026' })] });
  // A recent ineligible or pending result is never "current".
  assert.match(rv(p, [fresh({ 'Sohar Date': '2026-10-14', 'Sohar Status': 'complete.ineligible' })]).detail, /came back complete\.ineligible/);
  assert.ok(rv(p, [fresh({ 'Sohar Date': '2026-10-14', 'Sohar Status': 'pending' })]));
  // Coverage ending before the visit.
  assert.match(rv(p, [fresh({ 'Sohar Date': '2026-10-14', 'Coverage End': '10/18/2026' })]).detail, /coverage ends 10\/18\/2026/);
  // Next visit in a new month within 7 days: due; further out: not yet.
  const nov = pat({}, { appointments: [appt({ date: '11/02/2026' })] });
  const late = { ...CFG, today: new Date(2026, 9, 30) };
  assert.match(rv(nov, [fresh({ 'Sohar Date': '2026-10-28' })], late).detail, /new month/);
  const nov9 = pat({}, { appointments: [appt({ date: '11/09/2026' })] });
  assert.equal(rv(nov9, [fresh({ 'Sohar Date': '2026-10-28' })], late), undefined);
  // Re-verify in new month On: any month change re-verifies.
  assert.ok(rv(nov9, [fresh({ 'Sohar Date': '2026-10-28' })], { ...late, reverifyNewMonth: true }));
  // With verifyRows: only a result for the same payerId + member ID counts.
  const vr = { payerId: '94036', payerName: 'Blue Shield of California', memberId: 'BS-1', source: 'guarantor', nextgenPayer: 'Blue Shield CA' };
  const pv = pat({}, { appointments: [appt({})], verifyRows: [vr] });
  const same = fresh({ 'Sohar Date': '2026-10-14', 'Sohar payerId': '94036', 'Sohar Member ID': 'bs1' });
  assert.equal(rv(pv, [same]), undefined, 'same policy, eligible yesterday');
  assert.ok(rv(pv, [{ ...same, 'Sohar Member ID': 'BS-2' }]), 'different member ID');
  assert.ok(rv(pv, [{ ...same, 'Sohar payerId': '99999' }]), 'different payerId');
  assert.ok(rv(pv, [{ ...same, Match: 'Name only. Confirm.' }]), 'a name-only match never counts');
});

test('R4/R3: VIP and two insurances pills', () => {
  const vip = pat({ 'Primary Payer': 'VIP' }, { finClass: 'VIP', vip: true, verifyRows: [] });
  assert.match(R.patientFlags(vip, [fresh()], CFG).find((f) => f.code === 'VIP').detail, /discovery only/);
  const vip2 = pat({ 'Primary Payer': 'Partial VIP' }, { finClass: 'VIP', finClassLstEnc: 'Aetna', verifyRows: [{ payerId: '60054', payerName: 'Aetna', memberId: 'W1', source: 'lstEnc', nextgenPayer: 'Aetna' }] });
  assert.match(R.patientFlags(vip2, [fresh()], CFG).find((f) => f.code === 'VIP').detail, /Verifying other insurance: Aetna/);
  const dual = pat({}, { verifyRows: [{ payerId: '1', nextgenPayer: 'Odd Plan', source: 'guarantor' }, { payerId: '2', nextgenPayer: 'Aetna', source: 'lstEnc' }] });
  const f = R.patientFlags(dual, [fresh()], CFG).find((x) => x.code === 'DUAL');
  assert.match(f.detail, /Odd Plan \(guarantor\) and Aetna \(lstEnc\)/);
});

test('defensive: missing x, null patient, odd inputs', () => {
  assert.deepEqual(R.patientFlags(null, null, null), []);
  assert.deepEqual(R.patientFlags({ 'Per Nbr': '1' }, undefined, {}), []);
  const m = R.allFlags([{ 'Per Nbr': '1' }, null, {}], null, CFG);
  assert.equal(m.size, 1);
  assert.deepEqual(R.reviewRows(new Map(), []), []);
  assert.deepEqual(R.clinicRows('San Rafael', null, null, null, TODAY), []);
  assert.deepEqual(R.payerEvidence(null, null, null, TODAY), []);
  assert.deepEqual(R.evidenceSummary(null), []);
});

test('flags sort by pill order', () => {
  const p = pat({ 'Primary Payer': 'Health Net Medi-Cal', 'Secondary Payer': 'Aetna' }, { appointments: [appt({})] });
  const c = codes(R.patientFlags(p, [], CFG));
  assert.deepEqual(c, ['MEDICAL', 'MEDICAL_NEW', 'REVERIFY', 'SECONDARY']);
});

// ---------- allFlags / reviewRows ----------
function sample() {
  const a = pat({ 'Per Nbr': '1', 'Last Name': 'Alpha', 'First Name': 'Anna', 'Primary Payer': 'Kaiser Foundation Northern California', 'Secondary Payer': 'Aetna' },
    { finClass: 'Kaiser', appointments: [appt({ date: '10/21/2026', loc: 'San Rafael' })], location: 'San Rafael' });
  const b = pat({ 'Per Nbr': '2', 'Last Name': 'Bravo', 'First Name': 'Ben' }, { appointments: [appt({ date: '10/16/2026', time: '13:00', loc: 'San Rafael', event: 'TMS Follow Up' })] });
  const c = pat({ 'Per Nbr': '3', 'Last Name': 'Charlie', 'First Name': 'Cy' }, { appointments: [appt({ date: '10/18/2026', loc: 'Newport Beach' })] });
  const d = pat({ 'Per Nbr': '4', 'In Scope': 'No' }, { appointments: [appt({ loc: 'San Rafael' })] });
  const results = [
    fresh({ 'Per Nbr': '1', 'Sohar Date': '2026-10-01', 'OOP Status': 'Met', 'Copay': '$30', 'Choose in NextGen': 'Kaiser Foundation Northern California', Action: 'Eligible. No action.' }),
    fresh({ 'Per Nbr': '2', 'Sohar Date': '2026-10-14', 'Deductible Status': 'Approaching', 'Deductible Remaining': '$100', 'Sohar Link': 'https://example.invalid/v/2' }),
    fresh({ 'Per Nbr': '2', 'Sohar Date': '2026-10-12', 'Sohar Status': 'complete.ineligible' }),
  ];
  return { patients: [a, b, c, d], results };
}

test('allFlags + reviewRows', () => {
  const { patients, results } = sample();
  const m = R.allFlags(patients, results, CFG);
  assert.ok(codes(m.get('M1')).includes('KAISER_MM_NOAUTH'), 'flags are keyed by the patient key (MRN)');
  assert.ok(codes(m.get('M1')).includes('REVERIFY'));
  assert.deepEqual(m.get('M2'), []);
  const rv = R.reviewRows(m, patients);
  assert.deepEqual(Object.keys(rv[0]), ['Severity', 'Issue', 'MRN', 'Detail', 'Fix']);
  assert.deepEqual(Object.keys(rv[0]), createRequire(import.meta.url)('../core.js').REVIEW_COLS);
  assert.equal(rv[0].MRN, 'M1');
  assert.equal(rv[0].Severity, 'High');
  assert.ok(!rv.some((r) => r.Issue === 'Kaiser/VA MM'));
  assert.ok(rv.some((r) => r.Issue === 'Re-verify' && r.Fix.includes(R.REVERIFY_URL)));
  assert.ok(rv.some((r) => r.Issue === 'Secondary may cover' && r.Severity === 'Low'));
});

test('clinicRows: location filter, scope, order, benefits from latest eligible verification', () => {
  const { patients, results } = sample();
  const m = R.allFlags(patients, results, CFG);
  const rows = R.clinicRows('san rafael', patients, results, m, TODAY);
  assert.deepEqual(rows.map((r) => r['Per Nbr']), ['2', '1']);
  assert.deepEqual(Object.keys(rows[0]), R.CLINIC_COLS);
  const b = rows[0];
  assert.equal(b['Sohar Status'], 'complete.eligible');
  assert.equal(b['Deductible Status'], 'Approaching');
  assert.equal(b['Verified Date'], '10/14/2026');
  assert.equal(b['Days Since Verified'], 1);
  assert.equal(b['Re-verify'], '');
  assert.equal(b['Run Date'], '10/15/2026');
  const a = rows[1];
  assert.equal(a['Re-verify Link'], R.REVERIFY_URL);
  assert.match(a.Pills, /Kaiser\/VA MM no auth/);
  assert.match(a.Pills, /Re-verify/);
  assert.equal(a['OOP Status'], 'Met');
  assert.equal(a['Choose in NextGen'], 'Kaiser Foundation Northern California');
  assert.match(a.Notes, /Secondary Aetna/);
  assert.equal(R.clinicRows('All', patients, results, m, TODAY).length, 3);
  assert.equal(R.clinicRows('Castro Street', patients, results, m, TODAY).length, 0);
});

test('clinicReviewRows', () => {
  const { patients, results } = sample();
  const m = R.allFlags(patients, results, CFG);
  const rows = R.clinicReviewRows('San Rafael', patients, results, m, TODAY);
  assert.ok(rows.length >= 3);
  assert.deepEqual(Object.keys(rows[0]), R.CLINIC_REVIEW_COLS);
  assert.ok(rows.every((r) => r.MRN === 'M1'));
  assert.ok(rows.some((r) => r.Issue === 'KAISER_MM_NOAUTH' && r['Appt Date'] === '10/21/2026'));
});

test('clinicEmail: LASFIR + MRN only, grouped by date', () => {
  const { patients, results } = sample();
  const m = R.allFlags(patients, results, CFG);
  const rows = R.clinicRows('San Rafael', patients, results, m, TODAY);
  const txt = R.clinicEmail('San Rafael', rows, { runDate: TODAY });
  assert.match(txt, /Sohar pre-check: San Rafael/);
  assert.match(txt, /ALPANN \(MRN M1\)/);
  assert.match(txt, /BRABEN \(MRN M2\)/);
  assert.match(txt, /Deductible approaching \(\$100 left\)/);
  assert.match(txt, /OOP max met/);
  assert.ok(txt.indexOf('10/16/2026') < txt.indexOf('10/21/2026'));
  for (const bad of ['Alpha', 'Anna', 'Bravo', '01/02/1980']) assert.ok(!txt.includes(bad), bad);
  assert.ok(!/Kaiser\/VA MM[;,\n]|Kaiser\/VA MM$/.test(txt.replace(/Kaiser\/VA MM no auth/g, '')));
  const none = R.clinicEmail('Castro Street', [], { runDate: TODAY });
  assert.match(none, /Nothing to flag/);
});

test('payerEvidence: upsert keeps First Seen and earlier Sohar fields; summary pivot', () => {
  const p1 = pat({ 'Per Nbr': '7', 'Primary Payer': 'Odd Local Plan', 'Member ID': 'xq9-12345', Confidence: 'Auto-Low', State: 'CA' },
    { finClass: 'Contracted Commercial', finClassLstEnc: 'Aetna', guarZip: '94901-1234', location: 'San Rafael',
      verifyRows: [{ nextgenPayer: 'Odd Local Plan', memberId: 'xq9-12345', payerId: 'P1', source: 'guarantor' }, { nextgenPayer: 'Aetna', memberId: 'W77', payerId: '60054', source: 'lstEnc' }] });
  const res = [{ 'Per Nbr': '7', MRN: 'M7', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'NextGen Payer': 'Odd Local Plan', 'Sohar Payer': 'Local Admin Co', Plan: 'HMO', 'Sohar Date': '2026-10-14' }];
  const prev = [{ Key: '7|odd local plan|XQ9', 'Per Nbr': '7', 'NextGen Payer': 'Odd Local Plan', 'First Seen': '2026-09-01', 'Last Seen': '2026-09-01', 'Sohar Payer': 'Old Name' },
    { Key: '8|other|ABC', 'Per Nbr': '8', 'NextGen Payer': 'Other', 'Member ID Prefix': 'ABC', 'Location': 'San Rafael', 'First Seen': '2026-08-01', 'Last Seen': '2026-08-01', 'Sohar Payer': 'Local Admin Co' }];
  // prev was saved before v0.12 (Per Nbr keys): re-keyed to the MRN with the report's Per Nbr -> MRN pairs.
  const rows = R.payerEvidence([p1], res, prev, TODAY, new Map([['7', 'M7'], ['8', 'M8']]));
  assert.equal(rows.length, 3);
  assert.ok(rows.some((r) => r.Key === 'M8|other|ABC' && r.MRN === 'M8'), 'a patient not in this run is re-keyed too');
  assert.deepEqual(Object.keys(rows[0]), R.PAYER_EVIDENCE_COLS);
  const odd = rows.find((r) => r.Key === 'M7|odd local plan|XQ9');
  assert.equal(odd['First Seen'], '2026-09-01');
  assert.equal(odd['Last Seen'], '2026-10-15');
  assert.equal(odd['Sohar Payer'], 'Local Admin Co');
  assert.equal(odd['Guar Zip3'], '949');
  assert.equal(odd['Fin Class'], 'Contracted Commercial');
  assert.equal(odd.Confidence, 'Auto-Low');
  const aet = rows.find((r) => r['NextGen Payer'] === 'Aetna');
  assert.equal(aet['Fin Class'], 'Aetna');
  assert.equal(aet['Member ID Prefix'], 'W77');
  assert.equal(aet['First Seen'], '2026-10-15');
  assert.equal(aet['Sohar Payer'], '');
  // a later run with no Sohar result keeps the earlier Sohar Payer
  const again = R.payerEvidence([p1], [], rows, new Date(2026, 9, 20));
  assert.equal(again.find((r) => r.Key === odd.Key)['Sohar Payer'], 'Local Admin Co');
  assert.equal(again.find((r) => r.Key === odd.Key)['First Seen'], '2026-09-01');

  const sum = R.evidenceSummary([...rows, Object.assign({}, odd, { Key: 'M9|odd local plan|XQ9', MRN: 'M9', 'Sohar Payer': 'Local Admin Co' })]);
  assert.deepEqual(sum[0], { nextgenPayer: 'Odd Local Plan', prefix: 'XQ9', location: 'San Rafael', count: 2, soharPayers: ['Local Admin Co'] });
  assert.equal(sum.length, 3);
});

test('payerEvidence: falls back to Primary Payer + Member ID, skips Self Pay', () => {
  const p = pat({ 'Per Nbr': '5', 'Primary Payer': 'Cigna', 'Member ID': 'U123' }, { verifyRows: undefined });
  const s = pat({ 'Per Nbr': '6', 'Primary Payer': 'Self Pay' }, {});
  const rows = R.payerEvidence([p, s], [], [], TODAY);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].Key, 'M5|cigna|U12');
});

test('lasfir local copy', () => {
  assert.equal(R.lasfir("O'Neil-Ray", 'Jo Ann'), 'ONEJOA');
  assert.equal(R.lasfir('', ''), '');
});

test('clinicRows: same-day visits sort by clock time, not text (9:00 AM before 10:00 AM before 1:00 PM)', () => {
  const patients = [
    pat({ 'Per Nbr': '3001', 'Last Name': 'Aaa' }, { appointments: [appt({ time: '1:00 PM' })] }),
    pat({ 'Per Nbr': '3002', 'Last Name': 'Bbb' }, { appointments: [appt({ time: '10:00 AM' })] }),
    pat({ 'Per Nbr': '3003', 'Last Name': 'Ccc' }, { appointments: [appt({ time: '9:00 AM' })] }),
  ];
  const rows = R.clinicRows('All', patients, [], new Map(), TODAY);
  assert.deepEqual(rows.map((r) => r['Per Nbr']), ['3003', '3002', '3001']);
});

test('R8: a secondary only on the last encounter still flags, worded as last visit; QMB there counts', () => {
  const p = pat({ 'Primary Payer': 'Cigna' }, { finClass: 'Cigna', finClassLstEnc: 'Cigna', secPayerLstEnc: 'Aetna' });
  const f = R.patientFlags(p, [fresh()], CFG).find((x) => x.code === 'SECONDARY');
  assert.ok(f);
  assert.equal(f.severity, 'info');
  assert.match(f.detail, /^The secondary billed at the last visit \(Aetna\) may cover/);
  const mc = pat({ 'Primary Payer': 'Noridian Southern CA' }, { finClass: 'Medicare', secPayerLstEnc: 'QMB Noridian Southern CA' });
  const fl = R.patientFlags(mc, [fresh()], CFG);
  assert.equal(fl.find((x) => x.code === 'SECONDARY').severity, 'medium');
  assert.ok(codes(fl).includes('QMB'));
  // chart secondary wins over the last-encounter one
  const both = pat({ 'Secondary Payer': 'Cigna' }, { secPayerLstEnc: 'Aetna' });
  assert.match(R.patientFlags(both, [fresh()], CFG).find((x) => x.code === 'SECONDARY').detail, /^Secondary Cigna/);
});

test('bestResult: skips Superseded rows; same-day tie goes to the later timestamp', () => {
  const kp = fresh({ 'Sohar Date': '2026-10-14T01:00:00Z', 'Sohar Payer': 'Kaiser', Plan: 'KP PLAN', Action: 'Superseded: complete.eligible on Kaiser.' });
  const bs = fresh({ 'Sohar Date': '2026-10-14T02:00:00Z', Plan: 'SYNTH PPO', Action: 'Eligible on Blue Shield.' });
  assert.equal(R.bestResult([kp, bs]).Plan, 'SYNTH PPO');
  assert.equal(R.bestResult([bs, kp]).Plan, 'SYNTH PPO');
  const a = fresh({ 'Sohar Date': '2026-10-14T01:00:00Z', Plan: 'A' }), b = fresh({ 'Sohar Date': '2026-10-14T09:00:00Z', Plan: 'B' });
  assert.equal(R.bestResult([b, a]).Plan, 'B');
  assert.equal(R.bestResult([a, b]).Plan, 'B');
  const old = fresh({ 'Sohar Date': '2026-08-01T01:00:00Z', Plan: 'OLD PPO' });
  const now = fresh({ 'Sohar Date': '2026-10-14T01:00:00Z', 'Sohar Status': 'complete.ineligible', Plan: '' });
  assert.equal(R.bestResult([old, now]), now, 'an August eligible result does not beat a fresh ineligible one');
});

test('collapseFlags: repeated per-appointment flags become one line; review rows keep one per appointment', () => {
  const p = pat({ 'Primary Payer': 'UMR' }, { finClass: 'UMR', appointments: [
    appt({ date: '10/16/2026', event: 'TMS Treatment Unilateral', rendering: 'Lark, Jordan PMHNP' }),
    appt({ date: '10/17/2026', event: 'TMS Treatment Unilateral', rendering: 'Lark, Jordan PMHNP' }),
    appt({ date: '10/19/2026', event: 'TMS Treatment Unilateral', rendering: 'Moss, Kai NP' }),
  ] });
  const fl = R.patientFlags(p, [fresh()], CFG);
  assert.equal(fl.filter((f) => f.code === 'NP_TMS').length, 3);
  const c = R.collapseFlags(fl).filter((f) => f.code === 'NP_TMS');
  assert.equal(c.length, 2, 'one line per provider');
  assert.equal(c[0].detail, 'UMR does not allow an NP rendering provider for TMS: 2 visits with Lark, Jordan PMHNP (10/16, 10/17)');
  assert.match(c[1].detail, /on 10\/19\/2026 with Moss, Kai NP$/, 'a single visit keeps its own wording');
  const m = new Map([['M1001', fl]]);
  assert.equal(R.reviewRows(m, [p]).filter((r) => r.Issue === 'NP-TMS').length, 3);
  const row = R.clinicRows('All', [p], [fresh()], m, TODAY)[0];
  assert.equal(row['Review Items'].split(' | ').filter((t) => /NP rendering/.test(t)).length, 2);
});

test('payerEvidence: confidence comes from each verify row', () => {
  const p = pat({ 'Primary Payer': 'Zeta Odd Plan', Confidence: 'Auto-Low' }, { payerConfidence: 'Auto-Low', location: 'San Rafael', verifyRows: [
    { payerId: '1', nextgenPayer: 'Zeta Odd Plan', memberId: 'ZZ1', source: 'guarantor', confidence: 'Auto-Low' },
    { payerId: '2', nextgenPayer: 'Anthem Blue Cross', memberId: 'AN1', source: 'last encounter', confidence: 'Fin Class (report)' },
  ] });
  const ev = R.payerEvidence([p], [], [], TODAY);
  assert.equal(ev.find((r) => r['NextGen Payer'] === 'Anthem Blue Cross').Confidence, 'Fin Class (report)');
  assert.equal(ev.find((r) => r['NextGen Payer'] === 'Zeta Odd Plan').Confidence, 'Auto-Low');
});

test('R5: TriWest med management without auth flags, named by its own payer name', () => {
  const p = pat({ 'Primary Payer': 'TriWest' }, { finClass: 'Triwest', appointments: [appt({ event: 'Telepsych' })] });
  const f = R.patientFlags(p, [], CFG).find((x) => x.code === 'KAISER_MM_NOAUTH');
  assert.ok(f);
  assert.match(f.detail, /^TriWest med management/);
});

test('Dual eligible: Medicare with Medi-Cal secondary keeps the Medi-Cal flag and is treated like QMB', () => {
  const p = pat({ 'Primary Payer': 'Noridian Southern CA', 'Secondary Payer': 'Health Net Medi-Cal' }, { finClass: 'Medicare', finClassLstEnc: 'Medicare', appointments: [appt({ event: 'TMS Follow Up' })] });
  const fl = R.patientFlags(p, [], CFG);
  assert.ok(codes(fl).includes('MEDICAL'));
  const q = fl.find((x) => x.code === 'QMB');
  assert.ok(q);
  assert.match(q.detail, /treat like QMB/);
  // Commercial primary with Medi-Cal secondary: Medi-Cal flag, no QMB treatment
  const c = pat({ 'Primary Payer': 'Aetna', 'Secondary Payer': 'Health Net Medi-Cal' }, { finClass: 'Aetna', appointments: [appt({ event: 'TMS Follow Up' })] });
  assert.ok(!codes(R.patientFlags(c, [], CFG)).includes('QMB'));
});

test('R14: intervention visit + Sohar disagrees = Insurance changed? (flag only)', () => {
  const tms = pat({}, { appointments: [appt({ event: 'TMS Treatment Unilateral' })] });
  const flags = (p, rows) => R.patientFlags(p, rows, CFG).filter((f) => f.code === 'INS_CHANGED');
  assert.equal(flags(tms, [fresh()]).length, 0, 'eligible, no coverage change: no flag');
  const inel = flags(tms, [fresh({ 'Sohar Status': 'complete.ineligible' })]);
  assert.equal(inel.length, 1);
  assert.match(inel[0].detail, /^TMS Treatment Unilateral 10\/20\/2026: Sohar says not eligible with Blue Shield of California\. The authorization on file may be for the old insurance\.$/);
  assert.match(flags(tms, [fresh({ 'Coverage End': '10/18/2026' })])[0].detail, /coverage ends 10\/18\/2026, before the visit/);
  assert.equal(flags(tms, [fresh({ 'Coverage End': '12/31/2026' })]).length, 0);
  const moved = [fresh(), { 'Per Nbr': '1001', MRN: 'M1001', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Date': '2026-10-14T17:00:00Z', 'Sohar Payer': 'Acme Health', 'NextGen Payer': 'Blue Shield CA', 'Coverage Differs': 'Yes' }];
  assert.match(flags(tms, moved)[0].detail, /Sohar found Acme Health; NextGen has Blue Shield CA/);
  // ESK and IV ketamine count; consults and med management do not; an ineligible row superseded by a newer eligible one does not.
  assert.equal(flags(pat({}, { appointments: [appt({ event: 'Esketamine Follow Up' })] }), [fresh({ 'Sohar Status': 'complete.ineligible' })]).length, 1);
  assert.equal(flags(pat({}, { appointments: [appt({ event: 'IV Ketamine Follow Up' })] }), [fresh({ 'Sohar Status': 'complete.ineligible' })]).length, 1);
  assert.equal(flags(pat({}, { appointments: [appt({ event: 'TMS Consult In Person' })] }), [fresh({ 'Sohar Status': 'complete.ineligible' })]).length, 0);
  assert.equal(flags(pat({}, { appointments: [appt()] }), [fresh({ 'Sohar Status': 'complete.ineligible' })]).length, 0);
  assert.equal(flags(tms, [fresh({ 'Sohar Status': 'complete.ineligible', Action: 'Superseded by a newer check' })]).length, 0);
  assert.ok(R.isInterventionEvent('MT Redo') && R.isIVKEvent('IVK infusion') && !R.isIVKEvent('Esketamine'));
  assert.equal(R.PILLS.find((p) => p.code === 'INS_CHANGED').severity, 'high');
});
