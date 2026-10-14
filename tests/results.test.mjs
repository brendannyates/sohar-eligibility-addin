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
const PAYER_MAP = [{ nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' }];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 13)), scope: 'upcoming', locationState: { 'San Rafael': 'CA', 'Walnut Creek': 'CA' }, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: PAYER_MAP };

const MERGE_ROWS = () => [...TITLE, HDR,
  row({ 'Per Nbr': '100', 'Pat Name': 'Rivera, Ana', 'Birth Dt': '3/4/1980', 'Appt Dt': '11/2/2026', 'Beg Tm': 0.375, 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Rendering': 'Dr Test', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': '00912345' }),
  row({ 'Per Nbr': '100', 'Pat Name': 'Rivera, Ana', 'Birth Dt': '3/4/1980', 'Appt Dt': '10/20/2026', 'Beg Tm': 0.5, 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': '00912345' }),
  row({ 'Per Nbr': '101', 'Pat Name': 'Okafor, Chidi', 'Birth Dt': '1/1/1990', 'Appt Dt': '10/21/2026', 'Event': 'TMS Treatment Unilateral', 'Loc Name': 'San Rafael', 'Sts': 'Expected' }),
  row({ 'Per Nbr': '102', 'Pat Name': 'Lindqvist, Sven', 'Birth Dt': '2/2/1975', 'Appt Dt': '10/22/2026', 'Event': 'Telepsych', 'Loc Name': 'Walnut Creek', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': 'BS9' }),
  row({ 'Per Nbr': '103', 'Pat Name': 'Haddad, Mira', 'Birth Dt': '5/5/1985', 'Appt Dt': '10/23/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected' }),
];

test('parseCSV handles quotes, embedded commas and newlines, CRLF and BOM', () => {
  const rows = C.parseCSV('﻿a,b,c\r\n"x, y","he said ""hi""",\r\n"multi\nline",2,3\n');
  assert.deepEqual(rows, [['a', 'b', 'c'], ['x, y', 'he said "hi"', ''], ['multi\nline', '2', '3']]);
});

test('fmtTime formats Excel fractions and passes text through', () => {
  assert.equal(C.fmtTime(0.375), '9:00 AM');
  assert.equal(C.fmtTime(0.5625), '1:30 PM');
  assert.equal(C.fmtTime('10:15 AM'), '10:15 AM');
});

test('resultKind recognizes the Sohar verification and discovery exports', () => {
  assert.equal(C.resultKind(['verificationId', 'status', 'memberId', 'firstName', 'lastName', 'payerName', 'created']), 'verification');
  assert.equal(C.resultKind(['requestId', 'id', 'url', 'created', 'resolved', 'responsePayerId', 'responsePayerName']), 'discovery');
  assert.equal(C.resultKind(['Per Nbr', 'Appt Dt']), null);
});

test('mergeSoharResults joins verification by member ID and discovery by Per Nbr to the next appointment', () => {
  const ver = [['verificationId', 'status', 'memberId', 'firstName', 'lastName', 'payerName', 'created', 'deductibleRemaining', 'outOfPocketMax', 'outOfPocketRemaining', 'familyDeductibleRemaining'],
    ['v-1', 'complete.ineligible', '912345', 'Ana', 'Rivera', 'Blue Shield of California', '2026-10-14', '0', '5000', '1200.50', '900'],
    ['v-2', 'complete.eligible', 'ZZ1', 'Nobody', 'Synthetic', 'Cigna', '2026-10-14', '', '', '', '']];
  const disc = [['requestId', 'id', 'url', 'created', 'resolved', 'flowId', 'flowRunId', 'batchId', 'firstName', 'lastName', 'dateOfBirth', 'state', 'zipCode', 'requestPayerId', 'requestPayerName', 'memberId', 'responsePayerId', 'responsePayerName'],
    ['r-1', '101', 'https://example.test/r-1', '2026-10-14', '2026-10-14', '', '', '', 'Chidi', 'Okafor', '01/01/1990', 'CA', '', '', '', 'M55', '62308', 'Cigna'],
    ['r-2', '999', '', '2026-10-14', '2026-10-14', '', '', '', 'Mira', 'Haddad', '05/05/1985', 'CA', '', '', '', '', '', '']];
  const res = C.mergeSoharResults(MERGE_ROWS(), [{ name: 'v.csv', rows: ver }, { name: 'd.csv', rows: disc }, { name: 'junk.csv', rows: [['a'], ['b']] }], CFG);
  assert.equal(res.error, undefined);
  assert.deepEqual(res.unknownFiles, ['junk.csv']);
  for (const r of res.results) assert.deepEqual(Object.keys(r).sort(), [...C.RESULTS_COLS].sort());
  const by = (per, sent) => res.results.find((r) => r['Per Nbr'] === per && r['Sent As'] === sent);

  const a = by('100', 'Verification');
  assert.equal(a.Match, 'Member ID', 'leading zeros ignored');
  assert.equal(a['Next Appt'], '10/20/2026', 'soonest upcoming appointment');
  assert.equal(a['Next Time'], '12:00 PM');
  assert.equal(a['Upcoming Appts'], 2);
  assert.equal(a['Work Queue'], 'Action needed');
  assert.equal(a['Deductible Met'], 'Yes');
  assert.equal(a['Deductible Remaining'], '0', 'family column ignored');
  assert.equal(a['OOP Met'], 'No');

  const b = by('101', 'Discovery');
  assert.equal(b.Match, 'Per Nbr');
  assert.equal(b['Sohar Status'], 'Coverage found');
  assert.match(b.Action, /Add it to NextGen/);
  assert.equal(b['Sohar Link'], 'https://example.test/r-1');

  const d = by('103', 'Discovery');
  assert.equal(d.Match, 'Name + DOB', 'falls back when id does not match');
  assert.equal(d['Sohar Status'], 'No coverage found');

  const miss = res.results.find((r) => r['Last Name'] === 'Synthetic');
  assert.equal(miss['Work Queue'], 'Unmatched');
  assert.equal(miss['Per Nbr'], '');

  assert.equal(by('102', 'Verification')['Sohar Status'], 'Not in Sohar file');
  assert.equal(res.counts.withBenefits, 1);
  assert.equal(res.results[0]['Next Appt'], '10/20/2026', 'sorted by soonest appointment');
});

test('mergeSoharResults needs the appointment report on the Import sheet', () => {
  assert.match(C.mergeSoharResults([], [], CFG).error, /Import sheet/);
});

test('benefitFields reads met flags and met-vs-total amounts', () => {
  assert.equal(C.benefitFields({ deductiblemet: 'true' })['Deductible Met'], 'Yes');
  assert.equal(C.benefitFields({ oopmax: '$3,000', oopmetamount: '$3,000.00' })['OOP Met'], 'Yes');
  assert.equal(C.benefitFields({ individualdeductible: '1500', deductibleremaining: '200' })['Deductible Met'], 'No');
  assert.equal(C.benefitFields({ status: 'x' })['Deductible Met'], '');
});

test('adminMessage groups by location and uses LASFIR, not full names', () => {
  const rows = [
    { 'Work Queue': 'Action needed', 'Next Location': 'San Rafael', 'Next Appt': '10/20/2026', 'Next Time': '9:00 AM', 'Last Name': 'Smith', 'First Name': 'Jane', 'Per Nbr': '5', Action: 'Ineligible.', 'Deductible Met': 'Yes', 'OOP Met': '' },
    { 'Work Queue': 'Clear', 'Next Location': 'San Rafael', 'Last Name': 'Ok', 'First Name': 'Fine', 'Per Nbr': '6', Action: 'Eligible. No action.' },
  ];
  const msg = C.adminMessage(rows, { asOf: new Date(Date.UTC(2026, 9, 13)) });
  assert.match(msg, /San Rafael \(1\)/);
  assert.match(msg, /10\/20 9:00 AM · SMIJAN · Per 5 · Ineligible\. \(deductible met\)/);
  assert.doesNotMatch(msg, /Smith|Jane|Per 6/);
  assert.equal(C.lasfir("O'Neil", 'Al'), 'ONEAL');
});
