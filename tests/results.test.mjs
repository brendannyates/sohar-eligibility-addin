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
  for (const r of res.results) assert.deepEqual(Object.keys(r).sort(), [...C.RESULTS_COLS].sort());
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
