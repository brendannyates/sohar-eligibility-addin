import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const C = createRequire(import.meta.url)('../core.js');

const HDR = ['Appointment Pre-Check Complete', 'Res Name', 'Pat Sts', 'Portal Enrollment', 'Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Details',
  'Loc Name', 'Gender Identity', 'Current Gender', 'Preferred Pronoun', 'End Tm', 'Dur', 'Per Nbr', 'Birth Dt', 'Rendering', 'Sts', 'Rschd',
  'Rschd Desc', 'Md Rc', 'Guar Name', 'Pol Nbr (Lst Enc)', 'Payer Name (Cht Guar)', 'Sec Payer (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Cht Bal',
  'Pat Slf Bal', 'Crt Dt', 'Created By', 'Mod Dt', 'Modified By', 'No Show/Late Cancel Charged?'];

function row(o) {
  const r = new Array(HDR.length).fill('');
  for (const [k, v] of Object.entries(o)) r[HDR.indexOf(k)] = v;
  return r;
}
const TITLE = [[], ['', '', '', '', '', 'Mindful Health Solutions'], [], ['', '', '', '', '', 'Appointments Listing'], ['', '', '', '', '', 'From 6/1/2026 to 12/31/2026'], ['10/12/26 8:00 PM - Report Server'], []];
const PAYER_MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Optum Health Behavioral Solutions', payerId: '87726', payerName: 'UnitedHealthcare', confidence: 'Probable' },
  { nextgenPayerName: 'Self Pay', payerId: '', payerName: '', confidence: 'Not a payer' },
  { nextgenPayerName: 'Janssen Carepath Savings Program', payerId: '', payerName: '', confidence: 'Not a payer' },
  { nextgenPayerName: 'Optum UC', payerId: '', payerName: '', confidence: 'Review' },
];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 13)), scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, teleKeywords: ['Tele'], posTele: '10', posOffice: '11', specialty: 'PSY', payerMap: PAYER_MAP };

test('detectHeader finds the NextGen header row below the title block', () => {
  const rows = [...TITLE, HDR, row({ 'Per Nbr': '1', 'Appt Dt': '6/1/2026' })];
  const d = C.detectHeader(rows);
  assert.equal(d.headerRow, TITLE.length);
  assert.equal(d.colMap.payer, HDR.indexOf('Payer Name (Cht Guar)'));
  assert.equal(d.colMap.polNbr, HDR.indexOf('Pol Nbr (Cht Guar)'));
  assert.equal(d.colMap.state, undefined, 'no patient state column in this layout');
  assert.deepEqual(d.missing, []);
});

test('detectHeader returns null when Per Nbr is absent', () => {
  assert.equal(C.detectHeader([['Pat Name', 'Appt Dt'], ['x', '1/1/2026']]), null);
});

test('parseName handles NextGen "Last, First Middle" and suffixes', () => {
  assert.deepEqual(C.parseName('Rivera, Ana Maria'), { firstName: 'Ana', lastName: 'Rivera', flag: '' });
  assert.deepEqual(C.parseName('Okafor Jr, Chidi'), { firstName: 'Chidi', lastName: 'Okafor Jr', flag: '' });
  assert.deepEqual(C.parseName('Smith, John Jr'), { firstName: 'John', lastName: 'Smith Jr', flag: '' });
  assert.equal(C.parseName('Ana Maria Rivera').flag.length > 0, true);
  assert.equal(C.parseName('').flag, 'Missing name');
});

test('parseDate accepts text, ISO and Excel serials; fmtDate is MM/DD/YYYY', () => {
  assert.equal(C.fmtDate(C.parseDate('6/1/2026')), '06/01/2026');
  assert.equal(C.fmtDate(C.parseDate('2026-06-01 00:00:00')), '06/01/2026');
  assert.equal(C.fmtDate(C.parseDate(46174)), '06/01/2026');
  assert.equal(C.parseDate(''), null);
});

test('buildPatients dedupes per patient, routes, and derives POS/state/scope', () => {
  const rows = [...TITLE, HDR,
    // patient A: Blue Shield, telepsych upcoming -> verification, POS 10
    row({ 'Per Nbr': '100', 'Md Rc': '900', 'Pat Name': 'Rivera, Ana', 'Birth Dt': '3/4/1980', 'Appt Dt': '6/1/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Kept', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': 'XEA123', 'Pol Nbr (Lst Enc)': 'XEA123' }),
    row({ 'Per Nbr': '100', 'Md Rc': '900', 'Pat Name': 'Rivera, Ana', 'Birth Dt': '3/4/1980', 'Appt Dt': '11/2/2026', 'Event': 'Med Management Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': 'XEA123', 'Pol Nbr (Lst Enc)': 'XEA123' }),
    // patient B: no payer, in-office upcoming -> discovery, POS 11
    row({ 'Per Nbr': '101', 'Pat Name': 'Okafor, Chidi', 'Birth Dt': '1/1/1990', 'Appt Dt': '10/20/2026', 'Event': 'TMS Treatment Unilateral', 'Loc Name': 'San Rafael', 'Sts': 'Expected' }),
    // patient C: Self Pay, only a cancelled future appt -> discovery, out of scope
    row({ 'Per Nbr': '102', 'Pat Name': 'Lindqvist, Sven', 'Birth Dt': '2/2/1975', 'Appt Dt': '10/20/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Cancelled', 'Payer Name (Cht Guar)': 'Self Pay' }),
    // patient D: unmapped payer -> discovery + High review
    row({ 'Per Nbr': '103', 'Pat Name': 'Haddad, Mira', 'Birth Dt': '5/5/1985', 'Appt Dt': '12/1/2026', 'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Pacific Harbor Health Plan', 'Pol Nbr (Cht Guar)': 'PH1' }),
    // patient E: Probable payer, member id mismatch -> verification + Low review
    row({ 'Per Nbr': '104', 'Pat Name': 'Nakamura, Kenji', 'Birth Dt': '6/6/1965', 'Appt Dt': '12/1/2026', 'Event': 'Esketamine Follow Up', 'Loc Name': 'San Rafael', 'Sts': 'Expected', 'Payer Name (Cht Guar)': 'Optum Health Behavioral Solutions', 'Pol Nbr (Cht Guar)': 'U1', 'Pol Nbr (Lst Enc)': 'U2' }),
    ['Total Appointments (6)'],
  ];
  const res = C.buildPatients(rows, CFG);
  assert.equal(res.error, undefined);
  assert.equal(res.patients.length, 5);
  const by = Object.fromEntries(res.patients.map((p) => [p['Per Nbr'], p]));
  assert.equal(by['100'].Route, 'verification');
  assert.equal(by['100'].POS, '10');
  assert.equal(by['100'].State, 'CA');
  assert.equal(by['100']['Appt Count'], 2);
  assert.equal(by['100']['Next Appt'], '11/02/2026');
  assert.equal(by['100'].DOB, '03/04/1980');
  assert.equal(by['101'].Route, 'discovery');
  assert.equal(by['101'].POS, '11');
  assert.equal(by['102'].Route, 'discovery');
  assert.equal(by['102']['In Scope'], 'No', 'a cancelled future appointment is not upcoming');
  assert.equal(by['103'].Route, 'discovery');
  assert.deepEqual(res.unmappedPayers, { 'Pacific Harbor Health Plan': 1 });
  assert.equal(by['104'].Route, 'verification');
  assert.equal(by['104']['Member ID Mismatch'], 'Yes');
  assert.ok(res.review.some((r) => r.Issue === 'Unmapped payer' && r.Severity === 'High'));
  assert.ok(res.review.some((r) => r.Issue === 'Member ID mismatch' && r.MRN === 'PER-104'), 'review rows name the patient key: PER-<Per Nbr> without an MRN');
  assert.ok(res.review.some((r) => r.Issue === 'MRN missing' && r.MRN === 'PER-104'));
  assert.ok(!res.review.some((r) => r.Issue === 'MRN missing' && r.MRN === '900'));

  const ex = C.buildExports(res.patients, CFG);
  assert.equal(ex.verification.length, 2);
  // 103 (unmapped payer, member ID on file) goes on the payer mapping discovery with its member ID instead.
  assert.equal(ex.discovery.length, 1);
  assert.deepEqual(ex.mapping.map((r) => [r.id, r.memberId]), [['PER-103-M', 'PH1']]);
  assert.deepEqual(Object.keys(ex.mapping[0]), C.MAPPING_DISCOVERY_COLS);
  assert.equal(C.buildExports(res.patients, { ...CFG, mappingDiscovery: false }).discovery.length, 2, 'Payer mapping discovery Off: back on the discovery file');
  assert.deepEqual(Object.keys(ex.verification[0]), C.VERIFICATION_COLS);
  assert.deepEqual(Object.keys(ex.discovery[0]), C.DISCOVERY_COLS);
  assert.equal(ex.verification.find((r) => r.id === '900').memberId, 'XEA123');
  assert.equal(ex.verification[0].specialtyCode, 'PSY');
});

test('scope "all" includes patients without upcoming appointments', () => {
  const rows = [...TITLE, HDR, row({ 'Per Nbr': '1', 'Pat Name': 'A, B', 'Birth Dt': '1/1/1990', 'Appt Dt': '6/1/2026', 'Sts': 'Kept', 'Loc Name': 'San Rafael' })];
  assert.equal(C.buildPatients(rows, CFG).patients[0]['In Scope'], 'No');
  assert.equal(C.buildPatients(rows, { ...CFG, scope: 'all' }).patients[0]['In Scope'], 'Yes');
});

test('missing name and DOB columns still build, but flag High issues', () => {
  const H = HDR.filter((h) => h !== 'Pat Name' && h !== 'Birth Dt');
  const r = new Array(H.length).fill(''); r[H.indexOf('Per Nbr')] = '7'; r[H.indexOf('Appt Dt')] = '11/1/2026'; r[H.indexOf('Sts')] = 'Expected';
  r[H.indexOf('Payer Name (Cht Guar)')] = 'Blue Shield CA'; r[H.indexOf('Pol Nbr (Cht Guar)')] = 'X';
  const res = C.buildPatients([...TITLE, H, r], CFG);
  assert.equal(res.patients[0].Route, 'verification');
  assert.ok(res.review.some((x) => x.Issue === 'DOB missing'));
  assert.ok(res.review.some((x) => x.Issue === 'Name missing or incomplete'));
});

test('toCSV quotes commas and quotes, keeps template column order', () => {
  const csv = C.toCSV(['a', 'b'], [{ a: 'x,y', b: 'he said "hi"' }, { a: '', b: 7 }]);
  assert.equal(csv, 'a,b\r\n"x,y","he said ""hi"""\r\n,7\r\n');
});

test('suggestPayers ranks the obvious match first', () => {
  const sohar = [{ payerId: '94036', payerName: 'Blue Shield of California' }, { payerId: '47198', payerName: 'Blue Cross of California' }, { payerId: '62308', payerName: 'Cigna' }, { payerId: 'C1SCA', payerName: 'Blue Shield California Promise Health Plan (Medi-Cal)' }];
  assert.equal(C.suggestPayers('Blue Shield CA Blue Card', sohar, 3)[0].payerId, '94036');
  assert.equal(C.suggestPayers('Cigna Generic', sohar, 3)[0].payerId, '62308');
  assert.deepEqual(C.suggestPayers('', sohar, 3), []);
});
