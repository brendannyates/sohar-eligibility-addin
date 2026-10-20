import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const C = createRequire(import.meta.url)('../core.js');

// Synthetic data only.
const HDR = ['Pat Sts', 'Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Birth Dt', 'Rendering', 'Sts', 'Md Rc',
  'Pol Nbr (Lst Enc)', 'Payer Name (Cht Guar)', 'Sec Payer (Cht Guar)', 'Pol Nbr (Cht Guar)'];
function row(o) {
  const r = new Array(HDR.length).fill('');
  for (const [k, v] of Object.entries(o)) r[HDR.indexOf(k)] = v;
  return r;
}
const TITLE = [[], ['', 'Mindful Health Solutions'], ['', 'Appointments Listing'], []];
const PAYER_MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
  { nextgenPayerName: 'Optum Health Behavioral Solutions', payerId: '87726', payerName: 'UnitedHealthcare', confidence: 'Probable' },
  { nextgenPayerName: 'UHC Commercial', payerId: '87726', payerName: 'UnitedHealthcare', confidence: 'Confirmed' },
  { nextgenPayerName: 'Optum UC', payerId: '87726', payerName: 'UnitedHealthcare', confidence: 'Review' },
];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 13)), scope: 'upcoming', locationState: { 'San Rafael': 'CA', 'Walnut Creek': 'CA' }, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: PAYER_MAP };

const MERGE_ROWS = () => [...TITLE, HDR,
  row({ 'Per Nbr': '100', 'Md Rc': '900', 'Pat Name': 'Rivera, Ana', 'Birth Dt': '3/4/1980', 'Appt Dt': '11/2/2026', 'Beg Tm': 0.375, 'Event': 'TMS Treatment Unilateral', 'Loc Name': 'San Rafael', 'Rendering': 'Dr Test', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': '00912345' }),
  row({ 'Per Nbr': '100', 'Md Rc': '900', 'Pat Name': 'Rivera, Ana', 'Birth Dt': '3/4/1980', 'Appt Dt': '10/20/2026', 'Beg Tm': 0.5, 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': '00912345' }),
  row({ 'Per Nbr': '101', 'Pat Name': 'Okafor, Chidi', 'Birth Dt': '1/1/1990', 'Appt Dt': '10/21/2026', 'Event': 'TMS Treatment Unilateral', 'Loc Name': 'San Rafael', 'Sts': 'Expected' }),
  row({ 'Per Nbr': '102', 'Pat Name': 'Lindqvist, Sven', 'Birth Dt': '2/2/1975', 'Appt Dt': '10/22/2026', 'Event': 'Telepsych', 'Loc Name': 'Walnut Creek', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': 'BS9' }),
  row({ 'Per Nbr': '103', 'Pat Name': 'Haddad, Mira', 'Birth Dt': '5/5/1985', 'Appt Dt': '10/23/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected' }),
  row({ 'Per Nbr': '104', 'Pat Name': 'Mensah, Afia', 'Birth Dt': '7/7/1970', 'Appt Dt': '10/24/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C77' }),
  row({ 'Per Nbr': '105', 'Pat Name': 'Byrne, Niamh', 'Birth Dt': '8/8/1988', 'Appt Dt': '10/25/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C88' }),
];

// Column layout of Sohar's verification output (the batch export that echoes our id).
const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'flowId', 'flowRunId', 'batchId', 'memberId', 'payerName', 'status', 'statusReason',
  'subscriberFirstName', 'subscriberLastName', 'subscriberDateOfBirth', 'subscriberRelationshipCode', 'planName', 'networkType', 'planNumber', 'planNetworkId',
  'coverageStartDate', 'coverageEndDate', 'groupName', 'groupNumber', 'innCopay', 'innCoinsurance', 'innTotalDeductible', 'innRemainingDeductible',
  'innCoverageLevelCode', 'innOutOfPocket', 'innOutOfPocketRemaining', 'oopInnCoverageLevelCode',
  'innTierTwoCopay', 'innTierTwoCoinsurance', 'innTierTwoTotalDeductible', 'innTierTwoRemainingDeductible',
  'oonCopay', 'oonCoinsurance', 'oonTotalDeductible', 'oonRemainingDeductible', 'oonOutOfPocket', 'oonOutOfPocketRemaining'];
function ver(o) { return VER_HDR.map((k) => (o[k] == null ? '' : String(o[k]))); }
const DISC_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'flowId', 'flowRunId', 'batchId', 'firstName', 'lastName', 'dateOfBirth', 'state', 'zipCode',
  'requestPayerId', 'requestPayerName', 'memberId', 'responsePayerId', 'responsePayerName'];
function disc(o) { return DISC_HDR.map((k) => (o[k] == null ? '' : String(o[k]))); }

function merged() {
  const v = [VER_HDR,
    ver({ requestId: 'q-1', id: '100', url: 'https://example.test/q-1', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Blue Shield of California', status: 'complete.eligible',
      planName: 'SYNTH GOLD PPO', networkType: 'PR', coverageStartDate: '1/1/2026', innCopay: '40', innCoinsurance: '0.2', innTotalDeductible: '1500', innRemainingDeductible: '0',
      innCoverageLevelCode: 'IND', innOutOfPocket: '5000', innOutOfPocketRemaining: '800', oopInnCoverageLevelCode: 'IND',
      innTierTwoTotalDeductible: '9999', innTierTwoRemainingDeductible: '9999', oonTotalDeductible: '3000', oonRemainingDeductible: '3000', oonOutOfPocket: '10000', oonOutOfPocketRemaining: '10000' }),
    ver({ requestId: 'q-2', id: '102', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Blue Shield of California', status: 'complete.ineligible', statusReason: 'INACTIVE' }),
    ver({ requestId: 'q-3', id: '104', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Cigna', status: 'complete.eligible', innTotalDeductible: '0', innRemainingDeductible: '0', innOutOfPocket: '3000', innOutOfPocketRemaining: '0' }),
    ver({ requestId: 'q-4', id: '105', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Cigna', status: 'complete.eligible', coverageEndDate: '10/15/2026', innTotalDeductible: '500', innRemainingDeductible: '250' }),
    ver({ requestId: 'q-5', id: '777', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Aetna', status: 'error.member.id' }),
  ];
  const d = [DISC_HDR,
    disc({ requestId: 'r-1', id: '101', url: 'https://example.test/r-1', resolved: '2026-10-14', state: 'CA', responsePayerId: '62308', responsePayerName: 'Cigna' }),
    disc({ requestId: 'r-2', id: '103', resolved: '2026-10-14', state: 'CA', responsePayerId: '87726', responsePayerName: 'UnitedHealthcare' }),
    disc({ requestId: 'r-3', id: '999', resolved: '2026-10-14', firstName: 'Kenji', lastName: 'Nakamura', dateOfBirth: '06/06/1965', state: 'CA', responsePayerId: 'ZZZ01', responsePayerName: 'Synthetic Health' }),
  ];
  return C.mergeSoharResults(MERGE_ROWS(), [{ name: 'verification output.csv', rows: v }, { name: 'discovery output.csv', rows: d }, { name: 'junk.csv', rows: [['a'], ['b']] }], CFG);
}

test('parseCSV handles quotes, embedded commas and newlines, CRLF and BOM', () => {
  const rows = C.parseCSV('﻿a,b,c\r\n"x, y","he said ""hi""",\r\n"multi\nline",2,3\n');
  assert.deepEqual(rows, [['a', 'b', 'c'], ['x, y', 'he said "hi"', ''], ['multi\nline', '2', '3']]);
});

test('fmtTime formats Excel fractions and passes text through', () => {
  assert.equal(C.fmtTime(0.375), '9:00 AM');
  assert.equal(C.fmtTime(0.5625), '1:30 PM');
  assert.equal(C.fmtTime('10:15 AM'), '10:15 AM');
});

test('resultKind recognizes both Sohar verification layouts and the discovery export', () => {
  assert.equal(C.resultKind(['verificationId', 'status', 'memberId', 'firstName', 'lastName', 'payerName', 'created']), 'verification');
  assert.equal(C.resultKind(VER_HDR), 'verification');
  assert.equal(C.resultKind(DISC_HDR), 'discovery');
  assert.equal(C.resultKind(['Per Nbr', 'Appt Dt']), null);
});

test('verification output matches on id and carries in-network tier-one benefits', () => {
  const res = merged();
  assert.equal(res.error, undefined);
  assert.deepEqual(res.unknownFiles, ['junk.csv']);
  for (const r of res.results) assert.deepEqual(Object.keys(r).filter((k) => !['x', 'resultKey', 'chooseGuess'].includes(k)).sort(), [...C.RESULTS_COLS].sort());
  assert.ok(res.results.find((r) => r['Per Nbr'] === '100').x.appointments.length, 'non-column x carries the patient extras');
  const by = (per, sent) => res.results.find((r) => r['Per Nbr'] === per && r['Sent As'] === sent);

  const a = by('100', 'Verification');
  assert.equal(a.Match, 'Per Nbr');
  assert.equal(a.MRN, '900');
  assert.equal(a['Next Appt'], '10/20/2026', 'soonest upcoming appointment');
  assert.equal(a['Next Time'], '12:00 PM');
  assert.equal(a['Upcoming Appts'], 2);
  assert.equal(a['Upcoming Detail'], '10/20 Telepsych; 11/02 TMS Treatment Unilateral');
  assert.equal(a.Deductible, '1500', 'tier 2 and out-of-network ignored');
  assert.equal(a['Deductible Status'], 'Met');
  assert.equal(a['OOP Remaining'], '800');
  assert.equal(a['OOP Status'], 'Approaching', '800 of 5000 is 16%, under the 20% default');
  assert.equal(a.Coinsurance, '20%');
  assert.equal(a.Copay, '40');
  assert.equal(a.Plan, 'SYNTH GOLD PPO');
  assert.equal(a['Coverage Level'], 'IND');
  assert.equal(a['Benefit Flag'], 'Deductible met; OOP max approaching');
  assert.equal(a['Work Queue'], 'Clear');

  const b = by('102', 'Verification');
  assert.equal(b['Sohar Status'], 'complete.ineligible (INACTIVE)');
  assert.equal(b['Work Queue'], 'Action needed');

  const c = by('104', 'Verification');
  assert.equal(c['Deductible Status'], 'None');
  assert.equal(c['OOP Status'], 'Met');
  assert.equal(c['Benefit Flag'], 'OOP max met');

  const e = by('105', 'Verification');
  assert.match(e.Action, /^Coverage ends 10\/15\/2026, before the next visit/);
  assert.equal(e['Deductible Status'], 'Not met');

  assert.equal(res.results.find((r) => r['Sohar Ref'] === 'q-5')['Work Queue'], 'Unmatched');
  assert.equal(res.counts.withBenefits, 3);
  assert.equal(res.results[0]['Next Appt'], '10/20/2026', 'sorted by soonest appointment');
});

test('discovery says which NextGen insurance to choose from the Payer Map', () => {
  const res = merged();
  const by = (per) => res.results.find((r) => r['Per Nbr'] === per && r['Sent As'] === 'Discovery');

  const one = by('101');
  assert.equal(one['Choose in NextGen'], 'Cigna');
  assert.match(one.Action, /Coverage found\. Add it to NextGen, then verify\. Choose "Cigna" in NextGen\. Get the member ID from the Sohar link\./);

  const many = by('103');
  assert.equal(many['Choose in NextGen'], 'One of: Optum Health Behavioral Solutions | UHC Commercial', 'Review-confidence rows are not offered');
  assert.match(many.Action, /match the plan type on the card/);

  const none = res.results.find((r) => r['Sohar Ref'] === 'r-3');
  assert.equal(none['Choose in NextGen'], 'Not in Payer Map: Synthetic Health (ZZZ01)');
  assert.equal(none['Work Queue'], 'Unmatched');
});

test('mergeSoharResults still matches the older verification export by member ID', () => {
  const old = [['verificationId', 'status', 'memberId', 'firstName', 'lastName', 'payerName', 'created'], ['v-1', 'complete.eligible', '912345', 'Ana', 'Rivera', 'Blue Shield of California', '2026-10-14']];
  const r = C.mergeSoharResults(MERGE_ROWS(), [{ name: 'old.csv', rows: old }], CFG).results.find((x) => x['Sohar Ref'] === 'v-1');
  assert.equal(r.Match, 'Member ID', 'leading zeros ignored');
  assert.equal(r['Per Nbr'], '100');
});

test('mergeSoharResults needs the appointment report on the Import sheet', () => {
  assert.match(C.mergeSoharResults([], [], CFG).error, /Import sheet/);
});

test('accumulatorStatus: met, approaching, not met, none', () => {
  assert.equal(C.accumulatorStatus('1000', '0', '', 0.2), 'Met');
  assert.equal(C.accumulatorStatus('1000', '150', '', 0.2), 'Approaching');
  assert.equal(C.accumulatorStatus('1000', '150', '', 0.1), 'Not met');
  assert.equal(C.accumulatorStatus('0', '0', '', 0.2), 'None');
  assert.equal(C.accumulatorStatus('', '', 'true', 0.2), 'Met');
  assert.equal(C.accumulatorStatus('', '', '', 0.2), '');
});

test('benefitFields falls back to loose headers for other exports', () => {
  const b = C.benefitFields({ individualdeductible: '1500', deductibleremaining: '200', familydeductibleremaining: '0', oopmax: '$3,000', oopmetamount: '$3,000.00' });
  assert.equal(b['Deductible Status'], 'Approaching');
  assert.equal(b['OOP Status'], 'Met');
});

test('benefitsExport keeps verified patients only, one row each, in the TMS Tracker column order', () => {
  const ex = C.benefitsExport(merged().results);
  assert.deepEqual(ex.map((r) => r['Per Nbr']).sort(), ['100', '102', '104', '105']);
  assert.deepEqual(Object.keys(ex[0]), C.BENEFITS_EXPORT_COLS);
});

test('adminMessage: action list and benefits list, LASFIR not full names', () => {
  const res = merged();
  const msg = C.adminMessage(res.results, { asOf: res.asOf });
  assert.match(msg, /SAN RAFAEL|San Rafael \(\d+\)/);
  assert.match(msg, /OKACHI · Per 101 · Coverage found/);
  assert.doesNotMatch(msg, /Okafor|Chidi|Rivera/);
  const ben = C.adminMessage(res.results, { asOf: res.asOf, mode: 'benefits' });
  assert.match(ben, /10\/20 12:00 PM · RIVANA · Per 100 · Deductible met; OOP max approaching \(ded remaining 0, OOP remaining 800\)/);
  assert.match(ben, /MENAFI · Per 104 · OOP max met/);
  assert.doesNotMatch(ben, /LINSVE|BYRNIA/, 'not-met and ineligible patients are left out');
  assert.equal(C.lasfir("O'Neil", 'Al'), 'ONEAL');
});

test('buildTasks lists what to add or update in NextGen, soonest visit first', () => {
  const tasks = C.buildTasks(merged().results);
  for (const t of tasks) assert.deepEqual(Object.keys(t), C.TASK_COLS);
  const of = (per) => tasks.filter((t) => t['Per Nbr'] === per).map((t) => t.Task);
  assert.deepEqual(of('100'), ['Update benefits in NextGen']);
  assert.match(tasks.find((t) => t['Per Nbr'] === '100').Details, /Copay 40; Coinsurance 20%; Deductible 1500 \(0 left, Met\); OOP max 5000 \(800 left, Approaching\); Level IND; Plan SYNTH GOLD PPO/);
  assert.deepEqual(of('101'), ['Add insurance in NextGen'], 'discovery found coverage, NextGen has none');
  assert.match(tasks.find((t) => t['Per Nbr'] === '101').Details, /Choose Cigna\. Member ID: open the Sohar link\./);
  assert.deepEqual(of('102'), ['Update insurance in NextGen'], 'ineligible');
  assert.deepEqual(of('103'), ['Add insurance in NextGen']);
  assert.deepEqual(of('105'), ['Update insurance in NextGen', 'Update benefits in NextGen'], 'coverage ends before visit, plus benefits');
  const refs = tasks.filter((t) => !t['Per Nbr']).map((t) => t.Task).sort();
  assert.deepEqual(refs, ['Find patient for Sohar result', 'Find patient for Sohar result']);
  assert.equal(tasks[0]['Due By'], '10/20/2026');
  assert.equal(tasks[0]['Per Nbr'], '100', 'same day sorts by time: 12:00 PM before later visits');
});

test('keepDone carries Done marks across rebuilds by Task Key', () => {
  const tasks = C.buildTasks(merged().results);
  const prev = [{ 'Task Key': '100|Update benefits in NextGen', Done: 'Yes' }, { 'Task Key': 'gone|x', Done: 'Yes' }];
  const kept = C.keepDone(tasks, prev);
  assert.equal(kept.find((t) => t['Task Key'] === '100|Update benefits in NextGen').Done, 'Yes');
  assert.equal(kept.filter((t) => t.Done).length, 1);
});

test('chunkMessage and messageAssembler round-trip a large payload in pieces', () => {
  const big = { type: 'data', results: Array.from({ length: 500 }, (_, i) => ({ 'Per Nbr': String(i), Action: 'x'.repeat(100) })) };
  const pieces = C.chunkMessage(big, 5000);
  assert.ok(pieces.length > 5);
  const take = C.messageAssembler();
  let out = null;
  for (const p of pieces.reverse()) out = take(p) || out; // order does not matter
  assert.deepEqual(out, big);
  assert.deepEqual(take(JSON.stringify({ type: 'ready' })), { type: 'ready' });
});

test('taskMessage lists open tasks only, LASFIR not names', () => {
  const tasks = C.keepDone(C.buildTasks(merged().results), [{ 'Task Key': '100|Update benefits in NextGen', Done: 'Yes' }]);
  const msg = C.taskMessage(tasks);
  assert.match(msg, /OKACHI · Per 101 · Add insurance in NextGen: Choose Cigna\./);
  assert.doesNotMatch(msg, /RIVANA|Okafor|Chidi/, 'done task and full names left out');
});

// Discovery -> verification round trip. Synthetic data only.
function rawFiles() {
  const v = [VER_HDR, ver({ requestId: 'q-1', id: '100', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Blue Shield of California', status: 'complete.eligible' })];
  const d = [DISC_HDR,
    // Live discovery output carries the member ID; the 103 row leaves it out to cover the blank case.
    disc({ requestId: 'r-1', id: '101', url: 'https://example.test/r-1', resolved: '2026-10-14', firstName: 'Chidi', lastName: 'Okafor', dateOfBirth: '01/01/1990', state: 'CA', memberId: 'CG101', responsePayerId: '62308', responsePayerName: 'Cigna' }),
    disc({ requestId: 'r-2', id: '103', resolved: '2026-10-14', state: 'CA', memberId: 'UH103', responsePayerId: '87726', responsePayerName: 'UnitedHealthcare' }),
    disc({ requestId: 'r-4', id: '104', resolved: '2026-10-14', state: 'CA', memberId: 'C77', responsePayerId: '62308', responsePayerName: 'Cigna' }),
  ];
  return [{ name: 'verification output.csv', rows: v }, { name: 'discovery output.csv', rows: d }];
}

test('discovery coverage found becomes a verification import with the discovered payer', () => {
  const res = C.mergeSoharResults(MERGE_ROWS(), rawFiles(), CFG);
  const byId = Object.fromEntries(res.discoveryVerification.map((r) => [r.id, r]));
  assert.deepEqual(Object.keys(byId).sort(), ['101', '103', '104']);
  assert.equal(byId['101'].payerId, '62308');
  assert.equal(byId['101'].firstName, 'Chidi');
  assert.equal(byId['101'].dateOfBirth, '01/01/1990');
  assert.equal(byId['101'].memberId, 'CG101', 'member ID from the discovery output');
  assert.equal(byId['103'].memberId, 'UH103');
  assert.equal(byId['101'].placeOfServiceCode, '11');
  assert.equal(byId['103'].placeOfServiceCode, '10');
  assert.equal(byId['104'].memberId, 'C77');
  const csv = C.toCSV(C.VERIFICATION_COLS, res.discoveryVerification);
  assert.equal(csv.split('\r\n')[0], C.VERIFICATION_COLS.join(','));
  const d101 = res.results.find((r) => r['Per Nbr'] === '101' && r['Sent As'] === 'Discovery');
  assert.equal(d101['Discovery Follow-up'], 'Ready to verify');
  assert.match(d101.Action, /Copay, deductible and OOP need a Sohar verification/);
  assert.equal(d101.Copay, '', 'discovery carries no benefits');
});

test('a later verification upload is added to saved results and closes the discovery loop', () => {
  const first = C.accumulateResults([], rawFiles(), '2026-10-14');
  assert.equal(first.added, 4);
  const v2 = [VER_HDR, ver({ requestId: 'q-9', id: '101', resolved: '2026-10-15T01:00:00.000Z', memberId: 'CG101', payerName: 'Cigna', status: 'complete.eligible',
    planName: 'SYNTH OPEN ACCESS', innCopay: '30', innTotalDeductible: '1000', innRemainingDeductible: '100', innOutOfPocket: '4000', innOutOfPocketRemaining: '2500' }),
    // Same patient and payer but another member ID: a different policy, so it does not close 103's discovery.
    ver({ requestId: 'q-10', id: '103', resolved: '2026-10-15T01:00:00.000Z', memberId: 'UH-OTHER', payerName: 'UnitedHealthcare', status: 'complete.eligible' })];
  const second = C.accumulateResults(first.saved, [{ name: 'verification output 2.csv', rows: v2 }], '2026-10-15');
  assert.equal(second.added, 2);
  assert.equal(second.saved.length, 6);
  const again = C.accumulateResults(second.saved, [{ name: 'verification output 2.csv', rows: v2 }], '2026-10-15');
  assert.equal(again.replaced, 2, 'same file twice replaces, does not duplicate');
  assert.equal(again.saved.length, 6);

  const res = C.mergeSoharResults(MERGE_ROWS(), second.files, CFG);
  assert.ok(res.results.some((r) => r['Per Nbr'] === '100' && r['Sent As'] === 'Verification'), 'first upload still there');
  const d = res.results.find((r) => r['Per Nbr'] === '101' && r['Sent As'] === 'Discovery');
  const v = res.results.find((r) => r['Per Nbr'] === '101' && r['Sent As'] === 'Verification');
  assert.equal(d['Discovery Follow-up'], 'Verified: complete.eligible');
  assert.match(d.Action, /^Coverage found\. Add it to NextGen\. .*Verified eligible/);
  assert.equal(v['Discovery Follow-up'], 'From discovery');
  assert.match(v.Action, /^Eligible on Cigna, the plan discovery found\. Add it to NextGen/);
  assert.equal(v['Work Queue'], 'Action needed');
  assert.equal(v['Deductible Status'], 'Approaching');
  assert.deepEqual(res.discoveryVerification.map((r) => r.id).sort(), ['103', '104'], '101 no longer needs sending');

  assert.ok(C.benefitsExport(res.results).some((r) => r['Per Nbr'] === '101' && r.Plan === 'SYNTH OPEN ACCESS'));
  const tasks = C.buildTasks(res.results).filter((t) => t['Per Nbr'] === '101');
  assert.deepEqual(tasks.map((t) => t.Task).sort(), ['Add insurance in NextGen', 'Update benefits in NextGen']);
  assert.match(tasks.find((t) => t.Task === 'Add insurance in NextGen').Details, /verified: complete\.eligible; benefits are on the verification row/);
});

test('failed verifications go to a discovery import; eligible, pending and already-discovered patients do not', () => {
  const v = [VER_HDR,
    ver({ requestId: 'q-2', id: '102', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Blue Shield of California', status: 'complete.ineligible', statusReason: 'INACTIVE' }),
    ver({ requestId: 'q-4', id: '105', resolved: '2026-10-14T01:00:00.000Z', memberId: 'C88', payerName: 'Cigna', status: 'error.member.id' }),
    ver({ requestId: 'q-3', id: '104', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Cigna', status: 'pending' }),
    ver({ requestId: 'q-1', id: '100', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Blue Shield of California', status: 'complete.eligible' }),
    ver({ requestId: 'q-7', id: '777', resolved: '2026-10-14T01:00:00.000Z', payerName: 'Aetna', status: 'error.member.id' })];
  const res = C.mergeSoharResults(MERGE_ROWS(), [{ name: 'verification output.csv', rows: v }], CFG);
  const byId = Object.fromEntries(res.verificationDiscovery.map((r) => [r.id, r]));
  assert.deepEqual(Object.keys(byId).sort(), ['102', '105'], 'eligible 100, pending 104 and unmatched 777 stay out');
  assert.equal(byId['102'].payerId, '', 'ineligible: no payer hint, the patient may have changed plans');
  assert.equal(byId['105'].payerId, '62308', 'error: the payer is kept as a hint');
  assert.equal(byId['102'].lastName, 'Lindqvist');
  assert.equal(C.toCSV(C.DISCOVERY_COLS, res.verificationDiscovery).split('\r\n')[0], C.DISCOVERY_COLS.join(','));
  const r102 = res.results.find((r) => r['Per Nbr'] === '102');
  assert.equal(r102['Discovery Follow-up'], 'Sent to discovery');
  assert.match(r102.Action, /discovery import \(from verification\)/);
  assert.match(C.buildTasks(res.results).find((t) => t['Per Nbr'] === '102').Details, /on the discovery import/);
  assert.match(C.fileName('verifyFailed', new Date(2026, 9, 15)), /^Sohar Ready Import - Discovery \(from Verification\) 2026-10-15\.csv$/);

  // The discovery result comes back with the same payer: the patient is not sent again, and the old
  // failed verification does not count as the follow-up, so the plan goes out for a fresh verification.
  const d = [DISC_HDR, disc({ requestId: 'r-9', id: '105', resolved: '2026-10-15T02:00:00.000Z', state: 'CA', memberId: 'C88-01', responsePayerId: '62308', responsePayerName: 'Cigna' })];
  const res2 = C.mergeSoharResults(MERGE_ROWS(), [{ name: 'verification output.csv', rows: v }, { name: 'discovery output.csv', rows: d }], CFG);
  assert.deepEqual(res2.verificationDiscovery.map((r) => r.id), ['102']);
  const dv = res2.discoveryVerification.find((r) => r.id === '105');
  assert.ok(dv, '105 is on the discovery verification import');
  assert.equal(dv.memberId, 'C88-01');
});

test('two insurances: one eligible keeps the patient off the discovery import', () => {
  const rows = [...TITLE, HDR, row({ 'Per Nbr': '300', 'Pat Name': 'Tanaka, Yui', 'Birth Dt': '4/4/1984', 'Appt Dt': '10/22/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C30', 'Pol Nbr (Lst Enc)': 'BS30' })];
  const v = [VER_HDR,
    ver({ requestId: 'a', id: '300', resolved: '2026-10-14T01:00:00.000Z', memberId: 'C30', payerName: 'Cigna', status: 'complete.ineligible' }),
    ver({ requestId: 'b', id: '300', resolved: '2026-10-14T01:00:00.000Z', memberId: 'BS30', payerName: 'Blue Shield of California', status: 'complete.eligible' })];
  const res = C.mergeSoharResults(rows, [{ name: 'v.csv', rows: v }], CFG);
  assert.deepEqual(res.verificationDiscovery, []);
});

test('discovery from verification ignores out-of-scope patients and replaced policies', () => {
  const rows = [...MERGE_ROWS(), row({ 'Per Nbr': '106', 'Pat Name': 'Novak, Ivo', 'Birth Dt': '9/9/1960', 'Appt Dt': '9/1/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Completed', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C106' })];
  const v = [VER_HDR,
    ver({ requestId: 'a', id: '106', resolved: '2026-10-14T01:00:00.000Z', memberId: 'C106', payerName: 'Cigna', status: 'complete.ineligible' }),
    // 104's chart now has Cigna C77; an old result for another Cigna policy is history.
    ver({ requestId: 'b', id: '104', resolved: '2026-08-01T01:00:00.000Z', memberId: 'OLD104', payerName: 'Cigna', status: 'complete.ineligible' })];
  const res = C.mergeSoharResults(rows, [{ name: 'v.csv', rows: v }], CFG);
  assert.deepEqual(res.verificationDiscovery, []);
});

test('a discovery older than the failure does not block the follow-up', () => {
  const v = [VER_HDR, ver({ requestId: 'a', id: '104', resolved: '2026-10-14T01:00:00.000Z', memberId: 'C77', payerName: 'Cigna', status: 'complete.ineligible' })];
  const d = [DISC_HDR, disc({ requestId: 'r', id: '104', resolved: '2026-06-01T01:00:00.000Z', state: 'CA' })];
  const res = C.mergeSoharResults(MERGE_ROWS(), [{ name: 'v.csv', rows: v }, { name: 'd.csv', rows: d }], CFG);
  assert.deepEqual(res.verificationDiscovery.map((r) => r.id), ['104']);
});

test('a member ID Sohar rejected is not reused on the verification from discovery', () => {
  const v = [VER_HDR, ver({ requestId: 'a', id: '105', resolved: '2026-10-14T01:00:00.000Z', memberId: 'C88', payerName: 'Cigna', status: 'error.member.id' })];
  const d = [DISC_HDR, disc({ requestId: 'r', id: '105', resolved: '2026-10-15T01:00:00.000Z', state: 'CA', responsePayerId: '62308', responsePayerName: 'Cigna' })];
  const res = C.mergeSoharResults(MERGE_ROWS(), [{ name: 'v.csv', rows: v }, { name: 'd.csv', rows: d }], CFG);
  assert.equal(res.discoveryVerification.find((r) => r.id === '105').memberId, '', 'blank, so the pane asks for the right ID');
});

test('two policies with the same payer are both saved', () => {
  const v = [VER_HDR,
    ver({ requestId: 'a', id: '102', resolved: '2026-10-14T01:00:00.000Z', memberId: 'BS1', payerName: 'Blue Shield of California', status: 'complete.eligible' }),
    ver({ requestId: 'b', id: '102', resolved: '2026-10-14T01:00:00.000Z', memberId: 'BS2', payerName: 'Blue Shield of California', status: 'complete.ineligible' })];
  const acc = C.accumulateResults([], [{ name: 'v.csv', rows: v }], '2026-10-14');
  assert.equal(acc.added, 2);
  // Records saved under the old key (no member ID) are re-keyed, so the next upload replaces instead of duplicating.
  const old = acc.saved.map((s) => ({ ...s, Key: s.Key.replace(/\|[^|]*$/, '') }));
  assert.equal(C.accumulateResults(old, [{ name: 'v.csv', rows: v }], '2026-10-15').saved.length, 2);
});

test('soharTime reads ISO text and Excel serials; timeMins orders 9 AM before 10 AM', () => {
  assert.equal(C.soharTime('2026-10-14T01:00:00.000Z'), Date.UTC(2026, 9, 14, 1));
  assert.equal(new Date(C.soharTime(46309.5)).toISOString(), '2026-10-14T12:00:00.000Z');
  assert.equal(new Date(C.soharTime('46309.5')).toISOString(), '2026-10-14T12:00:00.000Z');
  assert.ok(C.timeMins('9:00 AM') < C.timeMins('10:00 AM'));
  assert.equal(C.timeMins(0.375), 540);
  assert.equal(C.timeMins('13:30'), 810);
});

test('same-day visits: the earliest sets POS and next visit', () => {
  const rows = [...TITLE, HDR,
    row({ 'Per Nbr': '400', 'Pat Name': 'Park, Jae', 'Birth Dt': '1/2/1981', 'Appt Dt': '10/20/2026', 'Beg Tm': '10:00 AM', 'Event': 'TMS Treatment Unilateral', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C400' }),
    row({ 'Per Nbr': '400', 'Pat Name': 'Park, Jae', 'Birth Dt': '1/2/1981', 'Appt Dt': '10/20/2026', 'Beg Tm': '9:00 AM', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C400' })];
  const p = C.buildPatients(rows, CFG).patients[0];
  assert.equal(p['Next Event'], 'Telepsych');
  assert.equal(p.POS, '10');
});

test('as-of date is the day picked, in any time zone', () => {
  const d = new Date(Date.UTC(2026, 9, 15));
  assert.equal(C.fmtDate(C.parseDate(d)), '10/15/2026');
  const rows = [...TITLE, HDR, row({ 'Per Nbr': '500', 'Pat Name': 'Diaz, Luz', 'Birth Dt': '1/1/1970', 'Appt Dt': '10/14/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C500' })];
  const built = C.buildPatients(rows, { ...CFG, asOf: d });
  assert.equal(C.fmtDate(built.asOf), '10/15/2026');
  assert.notEqual(built.patients[0]['In Scope'], 'Yes', 'a 10/14 visit is before an as-of of 10/15');
});

test('keepMemberIds keeps IDs typed on the Discovery Verify sheet', () => {
  const rows = [{ id: '101', payerId: '62308', memberId: '' }, { id: '103', payerId: '87726', memberId: '' }];
  const kept = C.keepMemberIds(rows, [{ id: '101', payerId: '62308', memberId: 'TYPED1' }, { id: '103', payerId: '99999', memberId: 'OLDPAYER' }]);
  assert.equal(kept[0].memberId, 'TYPED1');
  assert.equal(kept[1].memberId, '', 'a different payer does not carry the ID over');
});

test('trackerExport: one final row per verified patient, pending discovery left out, service from events', () => {
  const first = C.accumulateResults([], rawFiles(), '2026-10-14');
  const v2 = [VER_HDR, ver({ requestId: 'q-9', id: '101', resolved: '2026-10-15T01:00:00.000Z', memberId: 'CG101', payerName: 'Cigna', status: 'complete.eligible', innTotalDeductible: '1000', innRemainingDeductible: '100' })];
  const res = C.mergeSoharResults(MERGE_ROWS(), C.accumulateResults(first.saved, [{ name: 'v2.csv', rows: v2 }], '2026-10-15').files, CFG);
  const out = C.trackerExport(res.results, '10/15/2026 9:00 AM');
  const by = Object.fromEntries(out.rows.map((r) => [r['Per Nbr'], r]));
  assert.deepEqual(Object.keys(by).sort(), ['100', '101']);
  assert.equal(by['101'].Source, 'Discovery, then verification');
  assert.equal(by['100'].Source, 'Verification');
  assert.equal(by['101'].Service, 'TMS');
  assert.equal(by['100'].Service, 'TMS');
  assert.equal(by['101'].Exported, '10/15/2026 9:00 AM');
  assert.deepEqual(out.pending.map((r) => r['Per Nbr']).sort(), ['103', '104']);
  const header = C.toCSV(C.TRACKER_EXPORT_COLS, out.rows).split('\r\n')[0].split(',');
  assert.deepEqual(header.slice(0, C.BENEFITS_EXPORT_COLS.length), C.BENEFITS_EXPORT_COLS, 'benefits columns first, so the TMS import still reads it');
  assert.equal(C.serviceLine('10/20 Esketamine Follow Up; 10/21 TMS Treatment Unilateral'), 'TMS; ESK');
  assert.equal(C.serviceLine('Telepsych'), '');
});
