import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const C = createRequire(import.meta.url)('../core.js');

// Re-verify skip window, Sohar ref memory, no-identifier results, error.payer.npi. Approved by Brendan 2026-10-20.
// Synthetic data only.

const HDR = ['Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Birth Dt', 'Sts', 'Md Rc', 'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)'];
const row = (o) => HDR.map((k) => (o[k] == null ? '' : o[k]));
const PAYER_MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 20)), scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: PAYER_MAP };
const pt = (per, name, appt, payer, pol) => row({ 'Per Nbr': per, 'Md Rc': `M${per}`, 'Pat Name': name, 'Birth Dt': '1/2/1980', 'Appt Dt': appt, 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': payer, 'Pol Nbr (Cht Guar)': pol });
const IMPORT = () => [HDR,
  pt('200', 'Tester, Alpha', '10/27/2026', 'Blue Shield CA', 'BS200'),
  pt('201', 'Tester, Bravo', '11/02/2026', 'Cigna', 'C201'),
  pt('202', 'Tester, Charlie', '10/28/2026', 'Cigna', 'C202'),
  pt('203', 'Tester, Delta', '10/26/2026', 'Blue Shield CA', 'BS203'),
  pt('205', 'Tester, Echo', '10/29/2026', 'Cigna', 'C205'),
  pt('206', 'Tester, Foxtrot', '10/30/2026', '', ''),
];
const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'memberId', 'payerName', 'status', 'statusReason', 'firstName', 'lastName', 'coverageEndDate'];
const ver = (o) => VER_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const DISC_HDR = ['requestId', 'id', 'url', 'resolved', 'firstName', 'lastName', 'dateOfBirth', 'memberId', 'responsePayerId', 'responsePayerName'];
const disc = (o) => DISC_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const vfile = (...recs) => ({ name: 'verification output.csv', rows: [VER_HDR, ...recs.map(ver)] });
const URL = (id) => `https://app.soharhealth.com/verifications/${id}`;
const TODAY = new Date(2026, 9, 20, 9, 0); // Tue 10/20/2026, 9:00 local

const HISTORY = () => vfile(
  { requestId: 'q-200', id: '200', url: URL('v-200'), resolved: '2026-10-18T19:00:00Z', memberId: 'BS200', payerName: 'Blue Shield of California', status: 'complete.eligible' },
  { requestId: 'q-201', id: '201', resolved: '2026-10-18T19:00:00Z', memberId: 'C201', payerName: 'Cigna', status: 'complete.eligible' },
  { requestId: 'q-202', id: '202', resolved: '2026-10-19T19:00:00Z', memberId: 'C202', payerName: 'Cigna', status: 'complete.ineligible' },
  { requestId: 'q-203', id: '203', resolved: '2026-10-19T19:00:00Z', memberId: 'BS999', payerName: 'Blue Shield of California', status: 'complete.eligible' },
  { requestId: 'q-205', id: '205', resolved: '2026-10-13T19:00:00Z', memberId: 'C205', payerName: 'Cigna', status: 'complete.eligible' },
);

// ---------- A. reverifyDue ----------
const vrow = (date, status = 'complete.eligible', end = '') => ({ 'Sent As': 'Verification', 'Sohar Status': status, 'Sohar Date': date, 'Coverage End': end });
const due = (rows, today, nextAppt, o = {}) => C.reverifyDue(rows, { today, nextAppt, days: 7, ...o });
const d = (y, m, dd) => new Date(y, m - 1, dd, 10, 0);

test('reverifyDue: window is calendar days, due at N (1/31 -> 2/7 with N=7)', () => {
  const r = [vrow('2027-01-31')];
  const at6 = due(r, d(2027, 2, 6), '02/20/2027');
  assert.equal(at6.due, false);
  assert.equal(at6.code, 'fresh');
  assert.equal(at6.dueDate, '02/07/2027');
  const at7 = due(r, d(2027, 2, 7), '02/20/2027');
  assert.equal(at7.due, true);
  assert.equal(at7.code, 'window');
  assert.equal(at7.daysSince, 7);
  assert.equal(due([vrow('2027-01-10')], d(2027, 1, 16), '01/20/2027').due, false, '6 days');
  assert.equal(due([vrow('2027-01-10')], d(2027, 1, 17), '01/20/2027').due, true, '7 days');
});

test('reverifyDue: new-month exception (default) and year change', () => {
  const r = [vrow('2027-01-31')];
  assert.equal(due(r, d(2027, 1, 31), '02/01/2027').code, 'new month', 'appt 2/1 seen on 1/31');
  assert.equal(due(r, d(2027, 2, 1), '02/01/2027').due, true, 'appt 2/1 seen on 2/1');
  const far = due(r, d(2027, 2, 1), '02/14/2027');
  assert.equal(far.due, false, 'appt 2/14 is skipped');
  assert.equal(far.dueDate, '02/07/2027');
  assert.equal(due(r, d(2027, 2, 6), '02/14/2027').due, false, 'still skipped on 2/6');
  assert.equal(due(r, d(2027, 2, 7), '02/14/2027').due, true, 'due 2/7');
  assert.equal(due([vrow('2026-12-31')], d(2026, 12, 31), '01/02/2027').code, 'new month', 'year change');
  // Same month: no exception.
  assert.equal(due([vrow('2027-01-10')], d(2027, 1, 12), '01/14/2027').due, false);
  // Next check comes sooner when the new-month visit enters the window before N days pass.
  assert.equal(due([vrow('2027-01-27')], d(2027, 1, 27), '02/09/2027').dueDate, '02/03/2027');
});

test('reverifyDue: Re-verify in new month On', () => {
  const o = { newMonth: true };
  assert.equal(due([vrow('2027-01-25')], d(2027, 2, 1), '02/05/2027', o).due, true, 'today in a new month');
  assert.equal(due([vrow('2027-01-25')], d(2027, 1, 26), '02/20/2027', o).due, true, 'appt in a new month, outside the window');
  const ok = due([vrow('2027-01-10')], d(2027, 1, 12), '01/14/2027', o);
  assert.equal(ok.due, false);
  assert.equal(ok.dueDate, '01/17/2027');
  assert.equal(due([vrow('2027-01-28')], d(2027, 1, 28), '01/30/2027', o).dueDate, '02/01/2027', 'next check on the 1st');
});

test('reverifyDue: never skip on failed, pending or missing results; coverage end before the visit', () => {
  for (const s of ['complete.ineligible', 'error.member.id', 'error', 'pending', '']) {
    const r = due([vrow('2027-01-10', s)], d(2027, 1, 11), '01/14/2027');
    assert.equal(r.due, true, s);
    assert.equal(r.code, 'not eligible', s);
  }
  assert.equal(due([vrow('2027-01-09'), vrow('2027-01-10', 'complete.ineligible')], d(2027, 1, 11), '01/14/2027').due, true, 'the latest decides');
  assert.equal(due([vrow('2027-01-10', 'complete.ineligible'), vrow('2027-01-11')], d(2027, 1, 12), '01/14/2027').due, false);
  assert.equal(due([], d(2027, 1, 11), '01/14/2027').code, 'none');
  assert.equal(due([{ ...vrow('2027-01-10'), 'Sohar Status': 'Not in Sohar file' }], d(2027, 1, 11), '01/14/2027').code, 'none');
  assert.equal(due([{ ...vrow('2027-01-10'), 'Sent As': 'Discovery' }], d(2027, 1, 11), '01/14/2027').code, 'none');
  const end = due([vrow('2027-01-10', 'complete.eligible', '01/12/2027')], d(2027, 1, 11), '01/14/2027');
  assert.equal(end.code, 'coverage end');
  assert.equal(due([vrow('2027-01-10', 'complete.eligible', '01/14/2027')], d(2027, 1, 11), '01/14/2027').due, false, 'ends on the visit day');
});

test('reverifyDue: the verification day is the local calendar day of a timestamp', () => {
  const ts = '2027-01-11T05:30:00Z'; // evening of 1/10 on the West Coast, 1/11 in UTC
  const t = new Date(Date.parse(ts));
  const local = `${String(t.getMonth() + 1).padStart(2, '0')}/${String(t.getDate()).padStart(2, '0')}/${t.getFullYear()}`;
  assert.equal(due([vrow(ts)], d(2027, 1, 12), '01/14/2027').verified, local);
  assert.equal(C.fmtDate(C.soharDay(ts)), local);
  assert.equal(C.fmtDate(C.soharDay('2027-01-11')), '01/11/2027', 'a date-only value stays that day');
  assert.equal(C.fmtDate(C.soharDay(46398.75)), '01/11/2027', 'an Excel serial keeps its day');
});

test('policyResults / reverifyCheck: same payerId + member ID only; name-only matches never count', () => {
  const v = { payerId: '94036', payerName: 'Blue Shield of California', memberId: '00BS-200' };
  const r = { 'Per Nbr': '200', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2027-01-10', 'Sohar payerId': '94036', 'Sohar Payer': 'Blue Shield of California', 'Sohar Member ID': 'bs200', Match: 'Per Nbr' };
  assert.equal(C.policyResults([r], v).length, 1, 'member ID ignores case, punctuation and leading zeros');
  assert.equal(C.policyResults([{ ...r, 'Sohar Member ID': '' }], v).length, 0, 'blank member ID never matches');
  assert.equal(C.policyResults([{ ...r, 'Sohar payerId': '62308' }], v).length, 0);
  assert.equal(C.policyResults([{ ...r, 'Sohar Payer': 'Anthem Blue Cross' }], v).length, 0, 'a different Sohar payer name');
  assert.equal(C.policyResults([{ ...r, Match: 'Name only. Confirm.' }], v).length, 0);
  assert.equal(C.policyResults([{ ...r, Match: 'Sohar ref' }], v).length, 1);
  assert.equal(C.policyResults([r], v, '201').length, 0, 'other patient');
  assert.equal(C.policyResults([r, { ...r, 'Sohar Member ID': 'X' }], null).length, 2, 'no verifyRows: any trusted verification');
  assert.equal(C.reverifyCheck([r], v, { today: d(2027, 1, 12), nextAppt: '01/14/2027' }).due, false);
});

test('reverifyConfig: defaults and parsing', () => {
  assert.deepEqual(C.reverifyConfig({}), { skip: true, days: 7, newMonth: false });
  assert.deepEqual(C.reverifyConfig({ 'Skip recent verifications': 'Off', 'Re-verify days': '10', 'Re-verify in new month': 'On' }), { skip: false, days: 10, newMonth: true });
  assert.equal(C.reverifyConfig({ 'Re-verify days': 'x' }).days, 7);
  assert.equal(C.reverifyConfig({ 'Re-verify days': 0 }).days, 0);
  assert.deepEqual(C.REVERIFY_CONFIG_ROWS.map((r) => r.setting), ['Skip recent verifications', 'Re-verify days', 'Re-verify in new month']);
  for (const r of C.REVERIFY_CONFIG_ROWS) assert.ok(r.what.length > 20);
});

// ---------- A. applying the skip to a build ----------
function skipBuild(opts = {}) {
  const built = C.buildPatients(IMPORT(), CFG);
  const res = C.mergeSoharResults(IMPORT(), [HISTORY()], CFG);
  const out = C.applyReverifySkips(built.patients, res.results, { ...C.REVERIFY_DEFAULTS, today: TODAY, ...opts });
  return { built, res, out, by: Object.fromEntries(out.patients.map((p) => [p['Per Nbr'], p])) };
}

test('applyReverifySkips: recent eligible same policy leaves the verification file; others stay', () => {
  const { built, out, by } = skipBuild();
  assert.deepEqual(out.skipped.map((s) => s.per).sort(), ['200', '201']);
  const p = by['200'];
  assert.equal(p.Route, 'verification', 'Route stays verification');
  assert.equal(p['Route Reason'], 'mapped payer + member ID; skipped: verified eligible 10/18/2026, next check 10/25/2026');
  assert.deepEqual(p.x.reverifySkip, { verified: '10/18/2026', due: '10/25/2026', reason: 'verified eligible 10/18/2026, next check 10/25/2026', partial: false });
  assert.deepEqual(Object.keys(p).filter((k) => k !== 'x'), C.PATIENT_COLS, 'Patients columns unchanged');
  assert.equal(by['201'].x.reverifySkip.due, '10/25/2026', 'next visit 11/02 is a new month but 13 days out');
  assert.equal(by['202'].x.reverifySkip, undefined, 'ineligible: re-verify');
  assert.equal(by['203'].x.reverifySkip, undefined, 'a different member ID on the chart: never skipped');
  assert.equal(by['205'].x.reverifySkip, undefined, 'verified 7 days ago: due');
  assert.equal(by['206'].Route, 'discovery', 'discovery routing unchanged');
  const ex = C.buildExports(out.patients, CFG);
  assert.deepEqual(ex.verification.map((r) => r.id).sort(), ['M202', 'M203', 'M205'], 'Sohar ids are MRNs');
  assert.deepEqual(out.skipped.map((x) => [x.key, x.mrn]).sort(), [['M200', 'M200'], ['M201', 'M201']]);
  assert.deepEqual(Object.keys(ex.verification[0]), C.VERIFICATION_COLS);
  assert.deepEqual(ex.discovery.map((r) => r.id), ['M206']);
  assert.equal(C.buildExports(built.patients, CFG).verification.length, 5, 'the input patients are not changed');
  assert.equal(built.patients.find((x) => x['Per Nbr'] === '200')['Route Reason'], 'mapped payer + member ID');
});

test('applyReverifySkips: Skip Off, a Re-verify anyway request, out of scope, and a new-month visit inside the window', () => {
  assert.equal(skipBuild({ skip: false }).out.skipped.length, 0);
  const asked = skipBuild({ requests: new Map([['M200', '2026-10-19T12:00:00Z']]) });
  assert.deepEqual(asked.out.skipped.map((s) => s.per), ['201']);
  assert.ok(C.buildExports(asked.out.patients, CFG).verification.some((r) => r.id === 'M200'));
  assert.deepEqual(asked.out.requested.map((x) => [x.key, x.requested]), [['M200', '2026-10-19T12:00:00Z']], 'sent anyway on request: listed apart');
  assert.equal(asked.by['200'].x.reverifyRequested, true);
  // 201 (verified 10/18, next visit 11/02): skipped through 10/24, due 10/25.
  assert.ok(skipBuild({ today: new Date(2026, 9, 24, 9) }).by['201'].x.reverifySkip);
  assert.equal(skipBuild({ today: new Date(2026, 9, 25, 9) }).by['201'].x.reverifySkip, undefined);
  // Out of scope patients are never touched.
  const built = C.buildPatients(IMPORT(), CFG);
  const res = C.mergeSoharResults(IMPORT(), [HISTORY()], CFG);
  const outScope = built.patients.map((p) => (p['Per Nbr'] === '200' ? { ...p, 'In Scope': 'No' } : p));
  assert.ok(!C.applyReverifySkips(outScope, res.results, { today: TODAY }).skipped.some((s) => s.per === '200'));
});

test('applyReverifySkips: two insurances, only the recently verified one is left out', () => {
  const vA = { payerId: '94036', payerName: 'Blue Shield of California', memberId: 'A1', source: 'guarantor', nextgenPayer: 'Odd Plan' };
  const vB = { payerId: '62308', payerName: 'Cigna', memberId: 'B1', source: 'last encounter', nextgenPayer: 'Cigna' };
  const p = { 'Per Nbr': '300', 'In Scope': 'Yes', Route: 'verification', 'Route Reason': 'two insurances', 'Next Appt': '10/27/2026', 'First Name': 'Golf', 'Last Name': 'Tester', x: { verifyRows: [vA, vB] } };
  const r = { 'Per Nbr': '300', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-19', 'Sohar payerId': '94036', 'Sohar Payer': 'Blue Shield of California', 'Sohar Member ID': 'A1', Match: 'Per Nbr' };
  const out = C.applyReverifySkips([p], [r], { today: TODAY });
  const q = out.patients[0];
  assert.equal(q.x.reverifySkip.partial, true);
  assert.match(q['Route Reason'], /skipped: verified eligible 10\/19\/2026, next check 10\/26\/2026 \(the other insurance is still sent\)$/);
  assert.ok(q.x.verifyRows[0].reverifySkip && !q.x.verifyRows[1].reverifySkip);
  const ex = C.buildExports(out.patients, {});
  assert.deepEqual(ex.verification.map((v) => v.payerId), ['62308']);
});

test('Re-verify Requests: kept until a newer verification exists; one per patient', () => {
  const results = [
    { MRN: 'M200', 'Per Nbr': '200', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-18T19:00:00Z', Match: 'MRN' },
    { MRN: 'M201', 'Per Nbr': '201', 'Sent As': 'Verification', 'Sohar Status': 'pending', 'Sohar Date': '2026-10-21T01:00:00Z', Match: 'MRN' },
    { MRN: 'M202', 'Per Nbr': '202', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-22', Match: 'Name only. Confirm.' },
  ];
  const req = [
    { MRN: 'M200', Requested: '2026-10-19T12:00:00.000Z', By: 'Clinic tab' }, { MRN: 'M201', Requested: '2026-10-20T12:00:00.000Z' },
    { MRN: 'M202', Requested: '2026-10-20T12:00:00.000Z' }, { MRN: 'M203', Requested: 'not a date' }, { MRN: 'M200', Requested: '2026-10-01' },
  ];
  const pr = C.pruneReverifyRequests(req, results);
  assert.deepEqual(pr.rows.map((r) => r.MRN), ['M200', 'M202', 'M203'], '201 has a newer (pending) verification; 202 only a name match');
  assert.equal(pr.dropped, 2);
  assert.equal(pr.active.get('M200'), '2026-10-19T12:00:00.000Z');
  assert.equal(pr.rows[0].By, 'Clinic tab');
  const added = C.addReverifyRequest(pr.rows, 'M202', '2026-10-21T08:00:00.000Z', 'Admin');
  assert.deepEqual(added.map((r) => [r.MRN, r.Requested, r.By]), [['M200', '2026-10-19T12:00:00.000Z', 'Clinic tab'], ['M203', 'not a date', ''], ['M202', '2026-10-21T08:00:00.000Z', 'Admin']]);
  assert.deepEqual(Object.keys(added[0]), C.REVERIFY_REQUEST_COLS);
  assert.deepEqual(C.addReverifyRequest([], ['A', 'B'], 't', 'Admin: all skipped').map((r) => r.MRN), ['A', 'B']);
  assert.deepEqual(C.removeReverifyRequest(added, ['M203', 'M200']).map((r) => r.MRN), ['M202']);
  // Saved before v0.12: a Per Nbr row is read as PER-<Per Nbr>, and converted to the MRN with the report's pairs.
  const old = C.pruneReverifyRequests([{ 'Per Nbr': '201', Requested: '2026-10-22T00:00:00Z' }], results, new Map([['201', 'M201']]));
  assert.deepEqual(old.rows.map((r) => r.MRN), ['M201']);
  assert.deepEqual(C.pruneReverifyRequests([{ 'Per Nbr': '201', Requested: '2026-10-22T00:00:00Z' }], results).rows.map((r) => r.MRN), ['PER-201']);
  // A clinic file brought back: ticks add, unticks remove.
  const m = C.mergeClinicRequests(added, [{ MRN: 'M205', 'Re-verify Anyway': 'On', Requested: 't9' }, { MRN: 'M203', 'Re-verify Anyway': 'Off' }, { MRN: 'M202', 'Re-verify Anyway': 'On' }], 'Clinic file San Rafael');
  assert.deepEqual([m.added, m.removed, m.rows.map((r) => [r.MRN, r.By])], [1, 1, [['M200', 'Clinic tab'], ['M202', 'Admin'], ['M205', 'Clinic file San Rafael']]]);
});

// ---------- B. Sohar ref memory ----------
const FIRST = () => vfile({ requestId: 'q-200', id: '200', url: URL('v-200'), resolved: '2026-10-18T19:00:00Z', memberId: 'BS200', payerName: 'Blue Shield of California', status: 'complete.eligible', firstName: 'Alpha', lastName: 'Tester' });
const BLANK = (o = {}) => vfile({ requestId: 'q-200', url: URL('v-200'), resolved: '2026-10-18T19:00:00Z', payerName: 'Blue Shield of California', status: 'complete.eligible', ...o });

test('refs: learned from a Per Nbr match (requestId and the id in the Sohar link), keyed to the original result', () => {
  const res = C.mergeSoharResults(IMPORT(), [FIRST()], CFG);
  const key = res.results.find((r) => r['Per Nbr'] === '200').resultKey;
  assert.deepEqual(res.refs, [
    { 'Sohar Ref': 'q-200', MRN: 'M200', Kind: 'verification', 'Result Key': key },
    { 'Sohar Ref': 'v-200', MRN: 'M200', Kind: 'verification', 'Result Key': key },
  ]);
  assert.deepEqual(C.refsOf('verification', { requestid: 'a', url: URL('b'), verificationid: 'c' }), ['a', 'b', 'c']);
  assert.deepEqual(C.refsOf('verification', { requestid: 'a', url: 'https://example.test/other/b', verificationid: 'a' }), ['a']);
  // Known refs are not learned again.
  assert.deepEqual(C.mergeSoharResults(IMPORT(), [FIRST()], { ...CFG, refs: res.refs }).refs, []);
});

test('refs: a re-export with blank identifiers matches by Sohar ref and never duplicates uploads or the ledger', () => {
  const first = C.mergeSoharResults(IMPORT(), [FIRST()], CFG);
  const refs = C.appendRefs([], first.refs, '10/20/2026 9:00 AM').rows;
  const saved = C.accumulateResults([], [FIRST()], '2026-10-20T16:00:00Z', refs).saved;
  const ledger = C.payerOutcomes(first.results, first.patients, 't1');
  const key = first.results.find((r) => r['Per Nbr'] === '200').resultKey;

  // Same workbook: replaces the saved record and keeps its identifiers.
  const again = C.accumulateResults(saved, [BLANK()], '2026-10-21T16:00:00Z', refs);
  assert.equal(again.added, 0);
  assert.equal(again.replaced, 1);
  assert.equal(again.saved.length, 1);
  const rec = JSON.parse(again.saved[0].Record);
  assert.equal(rec.id, '200');
  assert.equal(rec.memberid, 'BS200');
  assert.equal(rec.firstname, 'Alpha');

  // After "Clear saved Sohar results": refs survive, so the blank row still lands on the same patient and key.
  const cleared = C.accumulateResults([], [BLANK()], '2026-10-21T16:00:00Z', refs);
  assert.equal(cleared.saved[0].Key, key);
  const res = C.mergeSoharResults(IMPORT(), cleared.files, { ...CFG, refs });
  const r = res.results.find((x) => x['Sent As'] === 'Verification' && x['Sohar Status'] !== 'Not in Sohar file');
  assert.equal(r.Match, 'Sohar ref');
  assert.equal(r['Per Nbr'], '200');
  assert.equal(r.resultKey, key);
  assert.equal(r['Sohar payerId'], '94036');
  assert.equal(C.appendOutcomes(ledger, C.payerOutcomes(res.results, res.patients, 't2')).added, 0, 'no duplicate ledger row');
  assert.deepEqual(res.refs, [], 'nothing new to learn from a Sohar ref match');

  // The link id alone is enough.
  const linkOnly = C.mergeSoharResults(IMPORT(), [BLANK({ requestId: '' })], { ...CFG, refs });
  const lr = linkOnly.results.find((x) => x['Sohar Link'] === URL('v-200'));
  assert.equal(lr.Match, 'Sohar ref');
  assert.equal(lr.resultKey, key);
});

test('refs: never learned from a name match; a ref for someone not in the report is an ordinary No match', () => {
  const nameOnly = vfile({ requestId: 'q-n', resolved: '2026-10-18', payerName: 'Cigna', status: 'complete.eligible', firstName: 'Bravo', lastName: 'Tester' });
  const res = C.mergeSoharResults(IMPORT(), [nameOnly], CFG);
  assert.equal(res.results.find((r) => r['Sohar Ref'] === 'q-n').Match, 'Name only. Confirm.');
  assert.deepEqual(res.refs, []);
  const byMember = vfile({ requestId: 'q-m', resolved: '2026-10-18', memberId: 'C201', payerName: 'Cigna', status: 'complete.eligible' });
  assert.deepEqual(C.mergeSoharResults(IMPORT(), [byMember], CFG).refs.map((r) => [r['Sohar Ref'], r.MRN]), [['q-m', 'M201']], 'a Member ID match teaches');
  const refs = [{ 'Sohar Ref': 'q-x', 'Per Nbr': '999', Kind: 'verification', 'Result Key': 'v|999|cigna|C999', 'First Matched': 't' }];
  const gone = C.mergeSoharResults(IMPORT(), [vfile({ requestId: 'q-x', resolved: '2026-10-18', payerName: 'Cigna', status: 'complete.eligible' })], { ...CFG, refs });
  const r = gone.results.find((x) => x['Sohar Ref'] === 'q-x');
  assert.equal(r.Match, 'No match');
  assert.match(r.Action, /remembered for Per Nbr 999/);
  assert.ok(C.buildTasks(gone.results).some((t) => t.Task === 'Find patient for Sohar result' && t['Sohar Ref'] === 'q-x'));
});

test('refs: discovery re-export with a blank id keeps its key', () => {
  const dfile = (o) => ({ name: 'discovery output.csv', rows: [DISC_HDR, disc({ requestId: 'r-206', resolved: '2026-10-18', responsePayerId: '62308', responsePayerName: 'Cigna', ...o })] });
  const first = C.mergeSoharResults(IMPORT(), [dfile({ id: '206', firstName: 'Foxtrot', lastName: 'Tester', dateOfBirth: '01/02/1980' })], CFG);
  assert.equal(first.refs.length, 1);
  const res = C.mergeSoharResults(IMPORT(), [dfile({})], { ...CFG, refs: first.refs });
  const r = res.results.find((x) => x['Sent As'] === 'Discovery' && x['Sohar Ref'] === 'r-206');
  assert.equal(r.Match, 'Sohar ref');
  assert.equal(r.resultKey, 'd|206');
});

test('appendRefs: append-only, first row for a ref wins, First Matched stamped', () => {
  // The previous row was saved before v0.12 (Per Nbr): kept, as a PER- key, converted when the report pairs it with an MRN.
  const prev = [{ 'Sohar Ref': 'q-1', 'Per Nbr': '100', Kind: 'verification', 'Result Key': 'k1', 'First Matched': 'old' }];
  const out = C.appendRefs(prev, [
    { 'Sohar Ref': 'q-1', MRN: '999', Kind: 'verification', 'Result Key': 'kX' },
    { 'Sohar Ref': 'q-1', MRN: 'M101', Kind: 'discovery', 'Result Key': 'd|101' },
    { 'Sohar Ref': 'q-2', MRN: '', Kind: 'verification', 'Result Key': 'k2' },
  ], 'now');
  assert.equal(out.added, 1);
  assert.deepEqual(out.rows, [{ 'Sohar Ref': 'q-1', MRN: 'PER-100', Kind: 'verification', 'Result Key': 'k1', 'First Matched': 'old' }, { 'Sohar Ref': 'q-1', MRN: 'M101', Kind: 'discovery', 'Result Key': 'd|101', 'First Matched': 'now' }]);
  assert.deepEqual(Object.keys(out.rows[1]), C.SOHAR_REF_COLS);
  assert.equal(C.refIndex(out.rows).get('verification|q-1').per, 'PER-100');
  assert.equal(C.appendRefs(prev, [], 'now', new Map([['100', 'M100']])).rows[0].MRN, 'M100');
});

// ---------- C. no identifiers ----------
test('No identifiers: one Review row, no NextGen task per row; ordinary No match unchanged', () => {
  const v = vfile(
    { requestId: 'q-a', url: URL('v-a'), resolved: '2026-10-18', payerName: 'Cigna', status: 'complete.eligible' },
    { requestId: 'q-b', resolved: '2026-10-18', payerName: 'Cigna', status: 'error.member.id' },
    { requestId: 'q-c', id: '777', resolved: '2026-10-18', memberId: 'Z1', payerName: 'Aetna', status: 'error.member.id' },
  );
  const dfile = { name: 'discovery output.csv', rows: [DISC_HDR,
    disc({ requestId: 'r-a', resolved: '2026-10-18', responsePayerId: '62308', responsePayerName: 'Cigna' }),
    disc({ requestId: 'r-b', resolved: '2026-10-18', firstName: 'Nobody', lastName: 'Known', dateOfBirth: '01/01/1990' })] };
  const res = C.mergeSoharResults(IMPORT(), [v, dfile], CFG);
  const by = Object.fromEntries(res.results.filter((r) => r['Sohar Ref']).map((r) => [r['Sohar Ref'], r]));
  for (const ref of ['q-a', 'q-b', 'r-a']) {
    assert.equal(by[ref].Match, 'No identifiers', ref);
    assert.match(by[ref].Action, /^No patient identifiers in this Sohar row .*export anomaly or a check run outside the batch: open the Sohar link/);
    assert.equal(by[ref]['Work Queue'], 'Unmatched');
  }
  assert.equal(by['q-c'].Match, 'No match');
  assert.equal(by['r-b'].Match, 'No match', 'names present but not found');
  assert.equal(res.counts.noIdentifiers, 3);
  assert.equal(res.counts.unmatched, 2);
  const tasks = C.buildTasks(res.results);
  assert.deepEqual(tasks.filter((t) => t.Task === 'Find patient for Sohar result').map((t) => t['Sohar Ref']).sort(), ['q-c', 'r-b']);
  const rv = C.noIdentifierReviewRows(res.results);
  assert.equal(rv.length, 1);
  assert.deepEqual(Object.keys(rv[0]), C.REVIEW_COLS);
  assert.equal(rv[0].Severity, 'High');
  assert.equal(rv[0].Issue, 'Sohar results with no patient identifiers');
  assert.match(rv[0].Detail, /^3 Sohar results with no id/);
  assert.deepEqual(C.noIdentifierReviewRows([]), []);
  assert.equal(C.payerOutcomes(res.results, res.patients, 't').filter((o) => ['q-a', 'q-b', 'r-a'].some((k) => o.Key.includes(k))).length, 0, 'not on the ledger');
  assert.ok(!C.adminMessage(res.results).includes('undefined'));
});

// ---------- D. learning ----------
test('error.payer.npi (provider NPI setup) is neutral, never a payer error', () => {
  assert.equal(C.verificationSignal('error.payer.npi'), 'neutral');
  assert.equal(C.verificationSignal('error.payer.npi (Provider NPI not registered with payer)'), 'neutral');
  assert.equal(C.verificationSignal('error.payer'), 'payer error');
  assert.equal(C.verificationSignal('error.payer.id (Invalid payer ID)'), 'payer error');
  const o = (per, status) => ({ Key: per, 'Per Nbr': per, Kind: 'Verification', 'NextGen Payer': 'Synthetic Gold', Source: 'chart', 'Sent payerId': 'G1', 'Sohar Status': status, Signal: C.verificationSignal(status), 'Sohar Date': '2026-10-14' });
  const map = [{ nextgenPayerName: 'Synthetic Gold', payerId: 'G1', payerName: 'Gold', confidence: 'Auto-High' }];
  const r = C.learnPayers(['1', '2', '3'].map((p) => o(p, 'error.payer.npi')), map, {})[0];
  assert.equal(r.payerErrors, 0);
  assert.equal(r.decision, '', 'NPI errors never demote');
});

// ---------- review fixes ----------
test('skip window: a changed payerId with the same member ID is never skipped (the payerId sent is recorded at upload)', () => {
  const MAP = [
    { nextgenPayerName: 'Medi-Cal', payerId: 'MCD01', payerName: 'Medi-Cal', confidence: 'Exact' },
    { nextgenPayerName: 'Partnership HP', payerId: 'PHP01', payerName: '', confidence: 'Exact' },
  ];
  const cfg = { ...CFG, payerMap: MAP };
  const sameName = { ...CFG, payerMap: [MAP[0], { ...MAP[1], payerName: 'Medi-Cal' }] };
  const chart = (payer) => [HDR, pt('300', 'Tester, Hotel', '10/27/2026', payer, 'CIN123')];
  const sent = vfile({ requestId: 'q-300', id: '300', url: URL('v-300'), resolved: '2026-10-18T19:00:00Z', memberId: 'CIN123', payerName: 'Medi-Cal', status: 'complete.eligible' });
  // As the build does it: saved uploads matched again against today's chart, then the skip window.
  const skipped = (rows, saved, c) => {
    const m = C.mergeSoharResults(rows, C.accumulateResults(saved, [], '', []).files, c);
    const out = C.applyReverifySkips(C.buildPatients(rows, c).patients, m.results, { ...C.REVERIFY_DEFAULTS, today: TODAY });
    return { pers: out.skipped.map((s) => s.per), ver: C.buildExports(out.patients, c).verification.map((v) => `${v.id}|${v.payerId}`) };
  };

  // Uploaded on 10/18 while the chart said Medi-Cal: the saved record keeps the payerId sent.
  const up = C.combineResults(chart('Medi-Cal'), [], [sent], '2026-10-18T20:00:00Z', [], cfg);
  assert.equal(JSON.parse(up.saved[0].Record)._sentpayerid, 'MCD01');
  assert.equal(up.res.results.find((r) => r['Per Nbr'] === '300').sentPayerId, 'MCD01');
  assert.deepEqual(skipped(chart('Medi-Cal'), up.saved, cfg).pers, ['300'], 'same policy: skipped');
  // The chart now says Partnership HP with the same CIN (Sohar payer name blank on the Payer Map, or the same name).
  assert.deepEqual(skipped(chart('Partnership HP'), up.saved, cfg), { pers: [], ver: ['M300|PHP01'] });
  assert.deepEqual(skipped(chart('Partnership HP'), up.saved, sameName), { pers: [], ver: ['M300|PHP01'] }, 'two payerIds sharing a display name');
  const shown = C.mergeSoharResults(chart('Partnership HP'), C.accumulateResults(up.saved, [], '', []).files, cfg).results.find((r) => r['Per Nbr'] === '300');
  assert.equal(shown['Sohar payerId'], 'MCD01', 'the Results sheet shows the payerId sent');
  assert.equal(C.payerOutcomes([shown], C.buildPatients(chart('Partnership HP'), cfg).patients, 't')[0]['Sent payerId'], 'MCD01');

  // Saved before the payerId was recorded: only today's chart is known, so a blank payer name is due, not a match.
  const old = C.accumulateResults([], [sent], 't', []).saved;
  assert.equal(JSON.parse(old[0].Record)._sentpayerid, undefined);
  assert.deepEqual(skipped(chart('Partnership HP'), old, cfg).pers, []);
  assert.deepEqual(skipped(chart('Medi-Cal'), old, cfg).pers, ['300'], 'names agree: still skipped');

  // A re-upload of the same check keeps the recorded payerId; a later upload never stamps older saved records.
  assert.equal(JSON.parse(C.combineResults(chart('Partnership HP'), up.saved, [sent], 't2', [], cfg).saved[0].Record)._sentpayerid, 'MCD01');
  const other = vfile({ requestId: 'q-x', id: '999', resolved: '2026-10-19', memberId: 'Z9', payerName: 'Cigna', status: 'complete.eligible' });
  const later = C.combineResults(chart('Partnership HP'), old, [other], 't3', [], cfg);
  assert.equal(JSON.parse(later.saved.find((s) => /\|300\|/.test(s.Key)).Record)._sentpayerid, undefined);
  // Name-only matches are not stamped.
  const nameOnly = vfile({ requestId: 'q-n', resolved: '2026-10-18', payerName: 'Medi-Cal', memberId: 'OTHER', status: 'complete.eligible', firstName: 'Hotel', lastName: 'Tester' });
  assert.deepEqual(C.combineResults(chart('Medi-Cal'), [], [nameOnly], 't', [], cfg).res.sentPayerIds, []);
});

test('refs learned in the same pass: a blank re-export is neither "No identifiers" nor a duplicate (same upload and upgrade)', () => {
  // The first match alone shows the gap: the ref is learned after the blank record was keyed and matched.
  const naive = C.mergeSoharResults(IMPORT(), C.accumulateResults([], [FIRST(), BLANK()], 't', []).files, CFG);
  assert.equal(naive.counts.noIdentifiers, 1);
  assert.equal(naive.refsPending, 1);

  const check = (cr) => {
    assert.equal(cr.saved.length, 1, 'one saved record');
    assert.equal(cr.res.counts.noIdentifiers, 0);
    assert.equal(cr.res.refsPending, 0);
    assert.deepEqual(C.noIdentifierReviewRows(cr.res.results), []);
    const r = cr.res.results.filter((x) => x['Sohar Ref'] === 'q-200');
    assert.deepEqual(r.map((x) => [x.Match, x['Per Nbr']]), [['Per Nbr (old id)', '200']]);
    assert.deepEqual(cr.refs.map((x) => x['Sohar Ref']), ['q-200', 'v-200'], 'the refs are still saved');
    assert.equal(C.payerOutcomes(cr.res.results, cr.res.patients, 't').length, 1);
  };
  // (b) Original and blank re-export in the same upload.
  check(C.combineResults(IMPORT(), [], [FIRST(), BLANK()], 't', [], CFG));
  // (a) Upgrade: the original is already saved but the Sohar Refs sheet is empty.
  const saved = C.accumulateResults([], [FIRST()], 't0', []).saved;
  const up = C.combineResults(IMPORT(), saved, [BLANK()], 't1', [], CFG);
  check(up);
  assert.equal(up.acc.added, 0);
  assert.equal(up.acc.replaced, 1);
  // Nothing pending: one match only, refs as before.
  const plain = C.combineResults(IMPORT(), [], [FIRST()], 't', [], CFG);
  assert.equal(plain.res.refsPending, 0);
  assert.equal(plain.refs.length, 2);
});

test('payer outcomes ledger: a name-only verification adds no agree or payer-error signal', () => {
  const rec = { requestId: 'q-n', resolved: '2026-10-18', payerName: 'Cigna', status: 'error.payer.notfound', firstName: 'Bravo', lastName: 'Tester' };
  const res = C.mergeSoharResults(IMPORT(), [vfile(rec)], CFG);
  assert.equal(res.results.find((r) => r['Sohar Ref'] === 'q-n').Match, 'Name only. Confirm.');
  assert.deepEqual(C.payerOutcomes(res.results, res.patients, 't'), []);
  const ok = C.mergeSoharResults(IMPORT(), [vfile({ ...rec, id: '201' })], CFG);
  assert.deepEqual(C.payerOutcomes(ok.results, ok.patients, 't').map((o) => [o.MRN, o['Sent payerId'], o.Signal]), [['M201', '62308', 'payer error']], 'an id match still counts');
});
