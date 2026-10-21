import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const req = createRequire(import.meta.url);
const C = req('../core.js');
const R = req('../rules.js');

// v0.12: MRN as the patient key, guarantor resubmission, payer mapping discovery, payer tests, payer metrics.
// Approved by Brendan 2026-10-20. Synthetic data only.

const HDR = ['Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Md Rc', 'Birth Dt', 'Sts', 'Guar Name', 'Guar Birth Dt', 'Guar Zip',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Fin Class (Cht Guar)', 'Fin Class (Lst Enc)', 'Pol Nbr (Lst Enc)'];
const row = (o) => HDR.map((k) => (o[k] == null ? '' : o[k]));
const MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
  { nextgenPayerName: 'Synthetic Blue Card', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Auto-Low' },
];
const SOHAR = [
  { payerId: '94036', payerName: 'Blue Shield of California', aliasPayerNames: '' },
  { payerId: '47198', payerName: 'Blue Cross of California', aliasPayerNames: 'Anthem Blue Cross' },
  { payerId: '62308', payerName: 'Cigna', aliasPayerNames: '' },
  { payerId: 'BCBS1', payerName: 'Blue Cross Blue Shield Blue Card', aliasPayerNames: '' },
];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 20)), scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP };
const pt = (per, mrn, name, payer, pol, extra) => row({ 'Per Nbr': per, 'Md Rc': mrn, 'Pat Name': name, 'Birth Dt': '1/2/1990', 'Appt Dt': '10/27/2026', Event: 'Telepsych',
  'Loc Name': 'San Rafael', Sts: 'Expected', 'Payer Name (Cht Guar)': payer, 'Pol Nbr (Cht Guar)': pol, 'Guar Zip': '94901', ...(extra || {}) });
const IMPORT = () => [HDR,
  pt('400', 'M400', 'Tester, Kid', '', '', { 'Guar Name': 'Tester, Parent', 'Guar Birth Dt': '3/3/1965' }),
  pt('401', 'M401', 'Tester, Teen', 'Cigna', 'C401', { 'Guar Name': 'Tester, Mom', 'Guar Birth Dt': '' }),
  pt('402', 'M402', 'Tester, Self', 'Blue Shield CA', 'BS402', { 'Guar Name': 'Tester, Self' }),
  pt('403', 'M403', 'Tester, Card', 'Synthetic Blue Card', 'XYZ403'),
  pt('404', 'M404', 'Tester, Card Two', 'Synthetic Blue Card', 'XYZ404'),
  pt('405', '', 'Tester, Nomrn', 'Cigna', 'C405'),
  pt('406', 'M406', 'Tester, Twin A', 'Cigna', 'C406'),
  pt('407', 'M406', 'Tester, Twin B', 'Cigna', 'C407'),
];
const DISC_HDR = ['requestId', 'id', 'resolved', 'created', 'firstName', 'lastName', 'dateOfBirth', 'memberId', 'responsePayerId', 'responsePayerName'];
const disc = (o) => DISC_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const VER_HDR = ['requestId', 'id', 'created', 'resolved', 'memberId', 'payerName', 'status', 'statusReason'];
const ver = (o) => VER_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const files = (v, d) => [...(v ? [{ name: 'v.csv', rows: [VER_HDR, ...v.map(ver)] }] : []), ...(d ? [{ name: 'd.csv', rows: [DISC_HDR, ...d.map(disc)] }] : [])];

test('patient key: MRN; a blank or shared MRN falls back to PER-<Per Nbr> with a Review item; Sohar ids follow', () => {
  const b = C.buildPatients(IMPORT(), CFG);
  const by = Object.fromEntries(b.patients.map((p) => [p['Per Nbr'], p]));
  assert.equal(C.keyOf(by['400']), 'M400');
  assert.equal(C.keyOf(by['405']), 'PER-405');
  assert.equal(C.keyOf(by['406']), 'PER-406', 'two Per Nbrs with one MRN: neither uses it');
  assert.equal(C.keyOf(by['407']), 'PER-407');
  assert.ok(b.review.some((r) => r.Issue === 'MRN missing' && r.MRN === 'PER-405'));
  assert.equal(b.review.filter((r) => r.Issue === 'MRN shared by two patients').length, 2);
  assert.deepEqual(Object.keys(b.review[0]), C.REVIEW_COLS);
  assert.equal(b.perToKey.get('400'), 'M400');
  const ex = C.buildExports(b.patients, CFG);
  assert.ok(ex.verification.some((r) => r.id === 'M401'));
  assert.ok(ex.verification.some((r) => r.id === 'PER-405'));
  assert.ok(ex.discovery.some((r) => r.id === 'M400'));
  assert.equal(C.idLabel(by['400']), 'MRN M400');
  assert.equal(C.idLabel(by['405']), 'Per 405');
  assert.deepEqual(C.splitId('M400-G'), { base: 'M400', tag: 'G', test: 0 });
  assert.deepEqual(C.splitId('M400-T3'), { base: 'M400', tag: 'T', test: 3 });
  assert.deepEqual(C.splitId('12-345'), { base: '12-345', tag: '', test: 0 });
});

test('MRN ids since: a check created before it is read as a Per Nbr, after it as the MRN', () => {
  // An MRN that is also another patient's Per Nbr: the time decides.
  const rows = [HDR, pt('500', '600', 'Tester, Alpha', 'Cigna', 'C500'), pt('600', '700', 'Tester, Bravo', 'Cigna', 'C600')];
  const cfg = { ...CFG, mrnSince: '2026-10-20 12:00 UTC' };
  const v = [
    { requestId: 'a', id: '600', created: '2026-10-19T10:00:00Z', resolved: '2026-10-19T11:00:00Z', memberId: 'C600', payerName: 'Cigna', status: 'complete.eligible' },
    { requestId: 'b', id: '600', created: '2026-10-21T10:00:00Z', resolved: '2026-10-21T11:00:00Z', memberId: 'C500', payerName: 'Cigna', status: 'complete.eligible' },
  ];
  const res = C.mergeSoharResults(rows, files(v), cfg);
  const a = res.results.find((r) => r['Sohar Ref'] === 'a'), bb = res.results.find((r) => r['Sohar Ref'] === 'b');
  assert.deepEqual([a['Per Nbr'], a.Match], ['600', 'Per Nbr (old id)']);
  assert.deepEqual([bb['Per Nbr'], bb.Match, bb.key], ['500', 'MRN', '600']);
  // No setting yet: the key is tried first.
  assert.equal(C.mergeSoharResults(rows, files([v[0]]), CFG).results.find((r) => r['Sohar Ref'] === 'a')['Per Nbr'], '500');
  // A tagged id is always a key.
  assert.equal(C.migrateKey('PER-600', new Map([['600', '700']])), '700');
  assert.equal(C.migrateKey('700', new Map([['700', 'x']])), '700', 'only PER- keys convert');
});

test('one verification per insurance: a last encounter naming the same policy is not sent twice', () => {
  const rows = [HDR, pt('410', 'M410', 'Tester, Same', 'Synthetic Blue Card', 'BS410', { 'Fin Class (Cht Guar)': 'Odd', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': 'BS410' }),
    pt('411', 'M411', 'Tester, Fin', 'Blue Shield CA', 'BS411', { 'Fin Class (Cht Guar)': 'Blue Shield' })];
  const b = C.buildPatients(rows, CFG);
  const p = b.patients.find((x) => x['Per Nbr'] === '410');
  const keys = (p.x.verifyRows || []).map((v) => `${v.payerId}|${v.memberId}`);
  assert.equal(new Set(keys).size, keys.length, 'no duplicate payerId + member ID');
  const ex = C.buildExports(b.patients, CFG);
  assert.ok(ex.verification.filter((r) => r.id === 'M410').length <= 1);
});

test('guarantor: candidates, requests, the discovery row, the tie-back and the Review flag', () => {
  const b = C.buildPatients(IMPORT(), CFG);
  const kid = b.patients.find((p) => p['Per Nbr'] === '400');
  assert.deepEqual(kid.x.guarantor, { name: 'Tester, Parent', firstName: 'Parent', lastName: 'Tester', dob: '03/03/1965' });
  assert.equal(b.patients.find((p) => p['Per Nbr'] === '402').x.guarantor, null, 'the patient is their own guarantor');
  // 400: discovery found nothing; 401: verification ineligible; 403: nothing back yet.
  const first = C.mergeSoharResults(IMPORT(), files(
    [{ requestId: 'v1', id: 'M401', created: '2026-10-21', resolved: '2026-10-21T01:00:00Z', memberId: 'C401', payerName: 'Cigna', status: 'complete.ineligible' }],
    [{ requestId: 'd1', id: 'M400', created: '2026-10-21', resolved: '2026-10-21' }]), CFG);
  const cands = C.guarantorCandidates(b.patients, first.results, []);
  assert.deepEqual(cands.map((c) => [c.key, c.reason, c.hasDob]), [['M400', C.GUARANTOR_REASONS.none, true], ['M401', C.GUARANTOR_REASONS.failed, false]]);
  const reqs = C.addGuarantorRequests([], cands, '2026-10-21T12:00:00Z', 'Admin');
  assert.deepEqual(Object.keys(reqs[0]), C.GUARANTOR_REQUEST_COLS);
  assert.deepEqual(C.guarantorCandidates(b.patients, first.results, reqs), [], 'already requested');
  const active = C.pruneGuarantorRequests(reqs, first.results).active;
  const ex = C.buildExports(b.patients, { ...CFG, guarantorRequests: active });
  const g = ex.discovery.filter((r) => /-G$/.test(r.id));
  assert.deepEqual(g.map((r) => [r.id, r.firstName, r.lastName, r.dateOfBirth, r.zipCode]), [['M400-G', 'Parent', 'Tester', '03/03/1965', '94901']]);
  assert.ok(!ex.discovery.some((r) => r.id === 'M400'), 'the guarantor row replaces a discovery that found nothing (no repeat)');
  assert.deepEqual(ex.guarantorNoDob.map((p) => p['Per Nbr']), ['401'], 'no guarantor DOB in the report: waits');
  for (const r of ex.discovery) assert.deepEqual(Object.keys(r), C.DISCOVERY_COLS);

  // The guarantor search comes back with coverage: tied to the patient, flagged, and verified with the guarantor as subscriber.
  const second = C.mergeSoharResults(IMPORT(), files(null, [
    { requestId: 'd1', id: 'M400', created: '2026-10-21', resolved: '2026-10-21' },
    { requestId: 'g1', id: 'M400-G', created: '2026-10-22', resolved: '2026-10-22T09:00:00Z', memberId: 'BSG9', responsePayerId: '94036', responsePayerName: 'Blue Shield of California' }]), CFG);
  const gr = second.results.find((r) => r['Sohar Ref'] === 'g1');
  assert.deepEqual([gr['Per Nbr'], gr.Match, gr['Searched As']], ['400', 'MRN', 'Guarantor']);
  assert.match(gr.Action, /^Found through the guarantor \(Tester, Parent\): the patient is likely a dependent on this policy\./);
  const dv = second.discoveryVerification.find((r) => r.id === 'M400');
  assert.deepEqual([dv.payerId, dv.memberId, dv.subscriberFirstName, dv.subscriberLastName, dv.subscriberDateOfBirth], ['94036', 'BSG9', 'Parent', 'Tester', '03/03/1965']);
  assert.equal(C.pruneGuarantorRequests(reqs, second.results).active.has('M400'), false, 'the request is done once the guarantor search is back');
  assert.equal(C.pruneGuarantorRequests(reqs, second.results).active.has('M401'), true);
  const flags = R.allFlags(b.patients, second.results, { today: new Date(2026, 9, 22) });
  const f = (flags.get('M400') || []).find((x) => x.code === 'GUARANTOR');
  assert.ok(f, 'Via guarantor flag');
  assert.equal(f.label, 'Via guarantor');
  assert.ok(R.reviewRows(flags, b.patients).some((r) => r.Issue === 'Via guarantor' && r.MRN === 'M400' && r.Severity === 'Medium'), 'on the Review sheet');
  const cr = R.clinicRows('All', b.patients, second.results, flags, new Date(2026, 9, 22)).find((r) => r['Per Nbr'] === '400');
  assert.equal(cr['Found Via Guarantor'], 'Yes');
  const tasks = C.buildTasks(second.results).filter((t) => t['Per Nbr'] === '400');
  assert.equal(new Set(tasks.map((t) => t['Task Key'])).size, tasks.length, 'patient and guarantor tasks have their own keys');
  assert.ok(tasks.some((t) => /\|G$/.test(t['Task Key'])));
});

test('guarantor themes: guarantor only, patient only, both, neither; no patient identifiers on the sheet rows', () => {
  const r = (key, searched, status, payer, date) => ({ key, MRN: key, 'Per Nbr': '', 'Sent As': 'Discovery', 'Searched As': searched, 'Sohar Status': status, 'Sohar payerId': payer, 'Sohar Payer': payer, 'Sohar Date': date || '2026-10-22', 'NextGen Payer': 'Cigna', Match: 'MRN' });
  const results = [
    r('A', 'Patient', 'No coverage found', ''), r('A', 'Guarantor', 'Coverage found', 'BS'),
    r('B', 'Patient', 'Coverage found', 'CG'), r('B', 'Guarantor', 'No coverage found', ''),
    r('C', 'Patient', 'Coverage found', 'CG'), r('C', 'Guarantor', 'Coverage found', 'CG'),
    r('D', 'Patient', 'Coverage found', 'CG'), r('D', 'Guarantor', 'Coverage found', 'BS'),
    r('E', 'Patient', 'No coverage found', ''), r('E', 'Guarantor', 'No coverage found', ''),
    r('F', 'Patient', 'No coverage found', ''),
  ];
  const t = C.guarantorThemes(results);
  assert.deepEqual(t.counts, { 'Guarantor only': 1, 'Patient only': 1, 'Both, same payer': 1, 'Both, different payers': 1, Neither: 1 });
  assert.deepEqual(t.rows.map((x) => x.Theme), ['Guarantor only', 'Both, different payers', 'Both, same payer', 'Patient only', 'Neither']);
  for (const x of t.rows) assert.deepEqual(Object.keys(x), C.GUARANTOR_THEME_COLS);
});

test('payer mapping discovery: unconfident payer + member ID, member ID at the end, one discovery each; Off sends them as before', () => {
  const b = C.buildPatients(IMPORT(), CFG);
  const ex = C.buildExports(b.patients, CFG);
  assert.deepEqual(ex.mapping.map((r) => [r.id, r.memberId]), [['M403-M', 'XYZ403'], ['M404-M', 'XYZ404']]);
  assert.deepEqual(C.MAPPING_DISCOVERY_COLS, [...C.DISCOVERY_COLS, 'memberId']);
  assert.ok(!ex.discovery.some((r) => /^M40[34]/.test(r.id)));
  const off = C.buildExports(b.patients, { ...CFG, mappingDiscovery: false });
  assert.deepEqual([off.mapping.length, off.discovery.filter((r) => /^M40[34]$/.test(r.id)).length], [0, 2]);
  // Results come back tied to the patient.
  const res = C.mergeSoharResults(IMPORT(), files(null, [{ requestId: 'm1', id: 'M403-M', created: '2026-10-21', resolved: '2026-10-21', responsePayerId: 'BCBS1', responsePayerName: 'Blue Cross Blue Shield Blue Card' }]), CFG);
  const m = res.results.find((r) => r['Sohar Ref'] === 'm1');
  assert.deepEqual([m['Per Nbr'], m['Searched As']], ['403', 'Payer mapping']);
  assert.ok(!res.results.some((r) => r['Per Nbr'] === '403' && r['Sohar Status'] === 'Not in Sohar file'));
  const o = C.payerOutcomes(res.results, res.patients, 't').find((x) => x.MRN === 'M403');
  assert.deepEqual([o.Source, o['Member Match'], o.Signal], ['payer mapping', 'Yes', 'other payer'], 'member ID sent and none echoed: evidence for the mapping');
});

test('payer tests: similar payers x patients, ids <key>-T<n>, results kept apart, a payer eligible for every patient is suggested', () => {
  const b = C.buildPatients(IMPORT(), CFG);
  const ranked = C.rankPayers('Synthetic Blue Card', SOHAR, 3);
  assert.ok(ranked.length >= 2 && ranked.every((x) => x.payerId));
  assert.equal(C.testWorthy('Exact', '62308', 0), false);
  assert.equal(C.testWorthy('Exact', '62308', 1), true, 'a payer error on record makes a name worth testing');
  assert.equal(C.testWorthy('Auto-Medium', 'X', 0), true);
  const built = C.buildPayerTests(b.patients, { soharPayers: SOHAR, payerMap: MAP, previous: [], payers: 3, patients: 2, when: 't0' });
  const n = ranked.length;
  assert.equal(built.rows.length, 2 * n, 'two patients x the similar payers');
  assert.ok(built.rows.every((r) => /^M40[34]-T\d$/.test(r.id)));
  for (const r of built.rows) assert.deepEqual(Object.keys(r), C.VERIFICATION_COLS);
  assert.deepEqual(Object.keys(built.table[0]), C.PAYER_TEST_COLS);
  assert.ok(built.rows.every((r) => r.memberId === (r.id.startsWith('M403') ? 'XYZ403' : 'XYZ404')), 'the same policy each time');
  // Same build again: no new tests, the same rows (nothing came back yet).
  const again = C.buildPayerTests(b.patients, { soharPayers: SOHAR, payerMap: MAP, previous: built.table, payers: 3, patients: 2 });
  assert.equal(again.table.length, built.table.length);
  assert.deepEqual(again.rows.map((r) => r.id).sort(), built.rows.map((r) => r.id).sort());
  // Results: one Sohar payer eligible for both patients, the others not.
  const win = built.table.find((t) => t['Candidate payerId'] === 'BCBS1') ? 'BCBS1' : built.table[0]['Candidate payerId'];
  const v = built.table.map((t, i) => ({ requestId: `t${i}`, id: t['Test Id'], created: '2026-10-21', resolved: '2026-10-21T02:00:00Z', memberId: t.MRN === 'M403' ? 'XYZ403' : 'XYZ404',
    payerName: t['Candidate Payer'], status: t['Candidate payerId'] === win ? 'complete.eligible' : 'error.payer' }));
  const res = C.mergeSoharResults(IMPORT(), files(v), { ...CFG, payerTestRows: built.table, payerTests: C.payerTestConfig({}) });
  assert.equal(res.payerTests.length, built.table.length);
  assert.ok(!res.results.some((r) => /-T\d$/.test(String(r['Sohar Ref'])) || r['Searched As'] === 'Payer test'), 'never the patient\'s results');
  const rec = C.recordPayerTests(built.table, res.payerTests);
  assert.equal(rec.updated, built.table.length);
  const sum = C.payerTestSummary(rec.table, { patients: 2 });
  const s = sum.get('synthetic blue card');
  assert.deepEqual([s.suggestion.payerId, s.suggestion.n, s.suggestion.source], [win, 2, 'test']);
  // Done: a name with a winner gets no more tests.
  assert.equal(C.buildPayerTests(b.patients, { soharPayers: SOHAR, payerMap: MAP, previous: rec.table, payers: 3, patients: 2 }).rows.length, 0);
  // The rollup shows the suggestion (never applied).
  const rollup = C.withTestSuggestions(C.learnPayers([], MAP, C.LEARN_DEFAULTS), sum);
  const r = rollup.find((x) => x.key === 'synthetic blue card');
  assert.equal(r.suggestion.payerId, win);
  assert.match(C.payerLearningRows([r])[0].Suggestion, /2 of 2 payer test patients eligible/);
  assert.match(C.payerLearningRows([r])[0]['Payer tests'], /2 of 2 eligible/);
});

test('payer metrics: unmapped and failed appointments per payer, test success, recurring from history', () => {
  const b = C.buildPatients(IMPORT(), CFG);
  const res = C.mergeSoharResults(IMPORT(), files([{ requestId: 'v1', id: 'M401', created: '2026-10-21', resolved: '2026-10-21T01:00:00Z', memberId: 'C401', payerName: 'Cigna', status: 'complete.ineligible' }]), CFG);
  const m = C.payerMetrics(b.patients, res.results, [], []);
  const card = m.rows.find((r) => r['NextGen Payer'] === 'Synthetic Blue Card');
  assert.deepEqual([card['Unmapped patients'], card['Unmapped appts']], ['2', '2']);
  const cig = m.rows.find((r) => r['NextGen Payer'] === 'Cigna');
  assert.deepEqual([cig['Failed verification patients'], cig['Failed verification appts'], cig['Most common failure']], ['1', '1', 'complete.ineligible (1)']);
  for (const r of m.rows) assert.deepEqual(Object.keys(r), C.PAYER_METRIC_COLS);
  assert.equal(m.totals.unmappedAppts, 2);
  let hist = [];
  for (const d of ['2026-10-18', '2026-10-19', '2026-10-20']) hist = C.appendMetricHistory(hist, m.rows, d);
  hist = C.appendMetricHistory(hist, m.rows, '2026-10-20');
  assert.equal(hist.filter((h) => h.Date === '2026-10-20').length, m.rows.length, 'one row per payer and day');
  const later = C.payerMetrics(b.patients, res.results, [], hist);
  assert.equal(later.rows.find((r) => r['NextGen Payer'] === 'Synthetic Blue Card').Recurring, 'Yes');
  assert.equal(later.totals.recurring, later.rows.length);
});

test('clinic rows carry the skip window and the clinic\'s Re-verify anyway; requests are keyed by MRN', () => {
  const b = C.buildPatients(IMPORT(), CFG);
  const hist = C.mergeSoharResults(IMPORT(), files([{ requestId: 'v9', id: 'M402', created: '2026-10-19', resolved: '2026-10-19T19:00:00Z', memberId: 'BS402', payerName: 'Blue Shield of California', status: 'complete.eligible' }]), CFG);
  const out = C.applyReverifySkips(b.patients, hist.results, { today: new Date(2026, 9, 20, 9) });
  assert.deepEqual(out.skipped.map((s) => [s.key, s.mrn, s.per, s.location]), [['M402', 'M402', '402', 'San Rafael']]);
  const rows = R.clinicRows('All', out.patients, hist.results, new Map(), new Date(2026, 9, 20, 9));
  const r = rows.find((x) => x['Per Nbr'] === '402');
  assert.deepEqual([r['Verification Skipped'], r['Next Check'], r['Re-verify Anyway']], ['Verified eligible 10/19/2026', '10/26/2026', '']);
  assert.equal(C.keyOf(r), 'M402');
  const asked = C.applyReverifySkips(b.patients, hist.results, { today: new Date(2026, 9, 20, 9), requests: new Map([['M402', 't']]) });
  const r2 = R.clinicRows('All', asked.patients, hist.results, new Map(), new Date(2026, 9, 20, 9)).find((x) => x['Per Nbr'] === '402');
  assert.deepEqual([r2['Re-verify Anyway'], r2['Verification Skipped']], ['Yes', 'Verified eligible 10/19/2026']);
  assert.deepEqual(R.CLINIC_REQUEST_COLS, ['MRN', 'Last Name', 'First Name', 'Next Appt', 'Re-verify Anyway', 'Requested']);
  assert.ok(R.CLINIC_COLS.includes('Re-verify Anyway') && R.CLINIC_REVIEW_COLS.includes('MRN') && !R.CLINIC_REVIEW_COLS.includes('Per Nbr'));
});

test('config: payer mapping discovery On and payer tests Off by default; new file names', () => {
  assert.deepEqual(C.payerTestConfig({}), { on: false, mappingDiscovery: true, payers: 3, patients: 2 });
  assert.deepEqual(C.payerTestConfig({ 'Payer tests': 'On', 'Payer mapping discovery': 'Off', 'Payer test payers': '5', 'Payer test patients': 'x' }), { on: true, mappingDiscovery: false, payers: 5, patients: 2 });
  assert.match(C.fileName('mapping', new Date(2026, 9, 20)), /^Sohar Ready Import - Discovery \(Payer Mapping\) 2026-10-20\.csv$/);
  assert.match(C.fileName('payerTests', new Date(2026, 9, 20)), /^Sohar Ready Import - Verification \(Payer Tests\) 2026-10-20\.csv$/);
  assert.ok(C.PAYER_TEST_CONFIG_ROWS.every((r) => r.what.length > 20));
});
