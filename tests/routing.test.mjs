import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const C = createRequire(import.meta.url)('../core.js');

// Synthetic data only. Header list is the current NextGen Appointments Listing layout.
const HDR = ['Appointment Pre-Check Complete', 'Res Name', 'Pat Sts', 'Portal Enrollment', 'Pat Name', 'Appt Date', 'Beg Tm', 'Event', 'Details',
  'Loc Name', 'Gender Identity', 'Current Gender', 'Preferred Pronoun', 'MRN', 'Status', 'Dur', 'Birth Dt', 'Rendering', 'End Tm', 'Per Nbr',
  'Nxt Appt Dt', 'Nxt Beg Tm', 'Nxt Event', 'Rschd', 'Rschd Desc', 'Guar Zip', 'Chkin Dt', 'Chkout Dt', 'Sec Payer (Lst Enc)', 'Fin Class (Lst Enc)',
  'Pol Nbr (Lst Enc)', 'Auth Cd (Lst Enc)', 'Payer Name (Cht Guar)', 'Sec Payer (Cht Guar)', 'Fin Class (Cht Guar)', 'Pol Nbr (Cht Guar)',
  'Auth Cd (Cht Guar)', 'Other ID Nbr', 'Crt Dt', 'Created By', 'Mod Dt', 'Modified By', 'Measures Needed?', 'No Show/Late Cancel Charged?'];
const H = (k) => HDR.indexOf(k);
function row(o) {
  const r = new Array(HDR.length).fill('');
  for (const [k, v] of Object.entries(o)) { assert.ok(H(k) >= 0, k); r[H(k)] = v; }
  return r;
}
const TITLE = [[], ['', '', '', '', '', 'Mindful Health Solutions'], ['', '', '', '', '', 'Appointments Listing'], []];
const SOHAR = [
  { payerId: '94036', payerName: 'Blue Shield of California', aliasPayerNames: '' },
  { payerId: '47198', payerName: 'Blue Cross of California', aliasPayerNames: 'Anthem Blue Cross; Anthem CA' },
  { payerId: '62308', payerName: 'Cigna', aliasPayerNames: '' },
  { payerId: '60054', payerName: 'Aetna', aliasPayerNames: '' },
  { payerId: '37237', payerName: 'UMR', aliasPayerNames: '' },
  { payerId: '94135', payerName: 'Kaiser Foundation Health Plan of Northern California', aliasPayerNames: '' },
  { payerId: '94134', payerName: 'Kaiser Foundation Health Plan of Southern California', aliasPayerNames: '' },
  { payerId: '99726', payerName: 'TriWest Healthcare Alliance', aliasPayerNames: '' },
  { payerId: '01111', payerName: 'Medicare of California', aliasPayerNames: '' },
  { payerId: '87726', payerName: 'UnitedHealthcare', aliasPayerNames: '' },
];
const MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
  { nextgenPayerName: 'Self Pay', payerId: '', payerName: '', confidence: 'Not a payer' },
  { nextgenPayerName: 'Synthetic Review Plan', payerId: '60054', payerName: 'Aetna', confidence: 'Review' },
];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 13)), scope: 'upcoming', locationState: { 'San Rafael': 'CA', 'Newport Beach': 'CA' }, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP, soharPayers: SOHAR };
const base = (per, extra) => row({ 'Per Nbr': per, 'MRN': `M${per}`, 'Pat Name': `Test${per}, Pat`, 'Birth Dt': '1/2/1980', 'Appt Date': '10/20/2026', 'Beg Tm': 0.375,
  'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Status': 'Scheduled', 'Rendering': 'Doe, Jan MD', 'Res Name': 'Doe, Jan MD', 'Guar Zip': '94901', ...extra });

test('new NextGen header list maps to the right fields; Nxt columns are not the appointment', () => {
  const d = C.detectHeader([...TITLE, HDR, base('1', {})]);
  const cm = d.colMap;
  assert.equal(cm.apptDt, H('Appt Date'));
  assert.equal(cm.begTm, H('Beg Tm'));
  assert.equal(cm.event, H('Event'));
  assert.equal(cm.sts, H('Status'));
  assert.equal(cm.patSts, H('Pat Sts'));
  assert.equal(cm.mrn, H('MRN'));
  assert.equal(cm.zip, H('Guar Zip'));
  assert.equal(cm.payer, H('Payer Name (Cht Guar)'));
  assert.equal(cm.secPayer, H('Sec Payer (Cht Guar)'));
  assert.equal(cm.secPayerLstEnc, H('Sec Payer (Lst Enc)'));
  assert.equal(cm.polNbr, H('Pol Nbr (Cht Guar)'));
  assert.equal(cm.polNbrLstEnc, H('Pol Nbr (Lst Enc)'));
  assert.equal(cm.finClass, H('Fin Class (Cht Guar)'));
  assert.equal(cm.finClassLstEnc, H('Fin Class (Lst Enc)'));
  assert.equal(cm.authCd, H('Auth Cd (Cht Guar)'));
  assert.equal(cm.authCdLstEnc, H('Auth Cd (Lst Enc)'));
  assert.equal(cm.rendering, H('Rendering'));
  assert.equal(cm.resName, H('Res Name'));
  assert.deepEqual(d.missing, []);
  assert.equal(d.footerRow, -1);
  assert.deepEqual(d.warnings, []);
});

test('R1: data stops at "Total Appointments"; leftover old-layout rows are counted and ignored', () => {
  // Old, longer layout left below the footer: staff names sit where the payer column now is.
  const old = (per) => { const r = new Array(HDR.length).fill(''); r[H('Per Nbr')] = per; r[H('Appt Date')] = '9/1/2026'; r[H('Payer Name (Cht Guar)')] = 'Staffer, Old NP'; return r; };
  const rows = [...TITLE, HDR,
    base('10', { 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': 'BS10' }),
    base('11', { 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'CG11' }),
    ['Total Appointments (2)'],
    old('900'), old('901'), [], old('902')];
  const d = C.detectHeader(rows);
  assert.equal(d.footerRow, TITLE.length + 3);
  assert.equal(d.dataEnd, d.footerRow);
  assert.equal(d.leftoverRows, 3);
  assert.match(d.warnings[0], /^3 leftover rows below the report total were ignored — upload the report instead of pasting/);
  const res = C.buildPatients(rows, CFG);
  assert.deepEqual(res.patients.map((p) => p['Per Nbr']).sort(), ['10', '11']);
  assert.equal(res.dataRows, 2);
  assert.ok(!res.autoMapped.some((a) => /Staffer/.test(a.nextgenPayerName)), 'leftover staff names never reach payer mapping');
  assert.equal(res.detection.leftoverRows, 3);
});

test('R1: provider names in the payer column raise a shifted-columns warning', () => {
  const rows = [...TITLE, HDR,
    base('20', { 'Payer Name (Cht Guar)': 'Doe, Jan MD' }),
    base('21', { 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C21' }), ['Total Appointments (2)']];
  const d = C.detectHeader(rows);
  assert.deepEqual(d.shiftedColumns, ['Doe, Jan MD']);
  assert.match(d.warnings.join(' '), /columns look shifted/);
  assert.equal(d.leftoverRows, 0);
});

test('R2: autoMatchPayer confidences', () => {
  const m = (n) => C.autoMatchPayer(n, SOHAR);
  assert.deepEqual([m('Blue Shield CA').confidence, m('Blue Shield CA').payerId], ['Auto-High', '94036'], 'ca -> california, "of" dropped');
  assert.equal(m('Cigna').confidence, 'Auto-High');
  assert.deepEqual([m('Anthem Blue Cross').confidence, m('Anthem Blue Cross').payerId], ['Auto-High', '47198'], 'alias match');
  assert.match(m('Anthem Blue Cross').evidence, /alias/);

  const k = m('Kaiser Foundation Northern California');
  assert.deepEqual([k.confidence, k.payerId], ['Auto-Medium', '94135'], 'clear winner over Southern California');
  assert.deepEqual([m('TriWest').confidence, m('TriWest').payerId], ['Auto-Medium', '99726']);
  assert.deepEqual([m('Blue Shield CA Blue Card').confidence, m('Blue Shield CA Blue Card').payerId], ['Auto-Medium', '94036']);

  const km = m('Kaiser Medicare Northern CA');
  assert.deepEqual([km.confidence, km.payerId], ['Auto-Low', '94135'], 'weak match keeps the best guess');
  const nor = m('Noridian Northern CA');
  assert.deepEqual([nor.confidence, nor.payerId], ['Auto-Low', ''], 'only generic words in common');
  assert.match(nor.evidence, /distinctive/);
  assert.equal(C.autoMatchPayer('Blue Shield CA', []).confidence, 'Auto-Low');

  // Two different payerIds with the same normalized name: not High.
  const dup = [...SOHAR, { payerId: 'X1', payerName: 'Cigna', aliasPayerNames: '' }];
  assert.equal(C.autoMatchPayer('Cigna', dup).confidence, 'Auto-Low');
});

test('R2: buildPatients auto-maps new names, routes High/Medium to verification, manual rows win', () => {
  const rows = [...TITLE, HDR,
    base('30', { 'Payer Name (Cht Guar)': 'Anthem Blue Cross', 'Pol Nbr (Cht Guar)': 'AB30' }),
    base('31', { 'Payer Name (Cht Guar)': 'TriWest', 'Pol Nbr (Cht Guar)': 'TW31', 'Sec Payer (Cht Guar)': 'Aetna' }),
    base('32', { 'Payer Name (Cht Guar)': 'Noridian Northern CA', 'Pol Nbr (Cht Guar)': 'NN32' }),
    base('33', { 'Payer Name (Cht Guar)': 'Synthetic Review Plan', 'Pol Nbr (Cht Guar)': 'SR33' }),
    base('34', { 'Payer Name (Cht Guar)': 'Self Pay' }),
    ['Total Appointments (5)']];
  const res = C.buildPatients(rows, CFG);
  const by = Object.fromEntries(res.patients.map((p) => [p['Per Nbr'], p]));
  const auto = Object.fromEntries(res.autoMapped.map((a) => [a.nextgenPayerName, a]));
  assert.deepEqual(Object.keys(auto).sort(), ['Aetna', 'Anthem Blue Cross', 'Noridian Northern CA', 'TriWest'], 'map rows and Self Pay are not auto-matched');
  assert.equal(auto['Anthem Blue Cross'].confidence, 'Auto-High');
  assert.match(auto['TriWest'].notes, /^Auto-matched by the add-in: Matched 1\/1 words/);
  assert.equal(auto['Noridian Northern CA'].confidence, 'Auto-Low');
  for (const a of res.autoMapped) for (const k of ['nextgenPayerName', 'payerId', 'payerName', 'confidence', 'notes']) assert.ok(k in a);

  assert.equal(by['30'].Route, 'verification');
  assert.equal(by['30'].Confidence, 'Auto-High');
  assert.equal(by['31'].Route, 'verification');
  assert.equal(by['31']['Secondary Sohar payerId'], '60054');
  assert.equal(by['32'].Route, 'discovery');
  assert.equal(by['32']['Sohar payerId'], '', 'a low-confidence guess is not sent as the discovery hint');
  assert.ok(res.review.some((r) => r['Per Nbr'] === '32' && r.Issue === 'Payer needs manual classification'));
  assert.equal(by['33'].Confidence, 'Review', 'manual confidence on an existing row wins');
  assert.equal(by['33'].Route, 'discovery');
  assert.equal(by['34'].Route, 'discovery', 'Self Pay unchanged');
  for (const p of res.patients) {
    assert.deepEqual(Object.keys(p).filter((k) => k !== 'x'), C.PATIENT_COLS);
    assert.equal(typeof p.x, 'object');
  }
  // Without a Sohar payer list the old behavior holds.
  const old = C.buildPatients(rows, { ...CFG, soharPayers: undefined });
  assert.deepEqual(old.autoMapped, []);
  assert.equal(old.patients.find((p) => p['Per Nbr'] === '30').Route, 'discovery');
});

test('patient x: appointments, fin class, auth codes, next rendering', () => {
  const rows = [...TITLE, HDR,
    base('40', { 'Appt Date': '11/3/2026', 'Beg Tm': 0.5, 'Event': 'TMS Treatment Unilateral', 'Loc Name': 'Newport Beach', 'Rendering': 'Roe, Sam PMHNP', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C40', 'Fin Class (Cht Guar)': 'Cigna', 'Fin Class (Lst Enc)': 'Cigna', 'Auth Cd (Cht Guar)': 'A1' }),
    base('40', { 'Appt Date': '10/21/2026', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C40', 'Fin Class (Cht Guar)': 'Cigna' }),
    base('40', { 'Appt Date': '10/22/2026', 'Status': 'Cancelled', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C40' }),
    base('40', { 'Appt Date': '9/1/2026', 'Status': 'Completed', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C40' }),
    ['Total Appointments (4)']];
  const p = C.buildPatients(rows, CFG).patients[0];
  assert.equal(p.x.finClass, 'Cigna');
  assert.equal(p.x.authCd, 'A1');
  assert.equal(p.x.rendering, 'Doe, Jan MD', 'next appointment (10/21) rendering');
  assert.equal(p.x.location, 'San Rafael');
  assert.equal(p.x.guarZip, '94901');
  assert.equal(p.x.vip, false);
  assert.equal(p.x.payerConfidence, 'Exact');
  assert.deepEqual(p.x.appointments.map((a) => a.date), ['10/21/2026', '11/03/2026'], 'cancelled and past left out, soonest first');
  assert.deepEqual(p.x.appointments[1], { date: '11/03/2026', time: '12:00 PM', event: 'TMS Treatment Unilateral', loc: 'Newport Beach', rendering: 'Roe, Sam PMHNP', status: 'Scheduled', authCd: 'A1', authCdLstEnc: '' });
  assert.deepEqual(p.x.verifyRows, [{ payerId: '62308', payerName: 'Cigna', memberId: 'C40', source: 'guarantor', nextgenPayer: 'Cigna', confidence: 'Exact' }]);
});

// Shared fixture for R3 / R4: Blue Shield patients teach the report that Fin Class "Blue Shield" means Blue Shield CA.
const BS = (per) => base(per, { 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': `BS${per}`, 'Fin Class (Cht Guar)': 'Blue Shield', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': `BS${per}` });

test('finClassPayerIndex: most common confidently mapped payer per Fin Class; overrides win', () => {
  const entries = [
    { finClass: 'Blue Shield', payer: 'Blue Shield CA' }, { finClass: 'Blue Shield', payer: 'Blue Shield CA' },
    { finClass: 'Blue Shield', payer: 'BS Unknown Variant' }, { finClass: 'Blue Shield', payer: 'BS Unknown Variant' }, { finClass: 'Blue Shield', payer: 'BS Unknown Variant' },
    { finClass: 'Medicare', payer: 'Medi Auto' }, { finClass: 'Self Pay', payer: 'Self Pay' },
  ];
  const map = [...MAP, { nextgenPayerName: 'Medi Auto', payerId: '01111', payerName: 'Medicare of California', confidence: 'Auto-Medium' }];
  const idx = C.finClassPayerIndex(entries, map);
  const bs = idx.get('blue shield');
  assert.deepEqual([bs.nextgenPayer, bs.payerId, bs.count, bs.total], ['Blue Shield CA', '94036', 2, 5], 'unmapped names do not count');
  assert.equal(idx.has('medicare'), false, 'Auto-Medium is not confident enough');
  assert.equal(idx.has('self pay'), false);
  const ov = C.finClassPayerIndex(entries, map, { overrides: { Medicare: '01111' } });
  assert.equal(ov.get('medicare').payerId, '01111');
  assert.equal(ov.get('medicare').source, 'config');
});

test('R3: low-confidence chart payer + different last encounter sends both insurances', () => {
  const rows = [...TITLE, HDR, BS('50'), BS('51'),
    // Chart payer is a weak auto-match (best guess Kaiser NorCal); last encounter was Blue Shield with another policy.
    base('52', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM52', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': 'BSL52' }),
    // Same, but the last encounter's Fin Class is not learnable from the report: discovery as before.
    base('53', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM53', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Beacon', 'Pol Nbr (Lst Enc)': 'BE53' }),
    // Low confidence but the last encounter looks the same: discovery.
    base('54', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM54', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Medicare Advantage', 'Pol Nbr (Lst Enc)': 'KM54' }),
    ['Total Appointments (5)']];
  const res = C.buildPatients(rows, CFG);
  const by = Object.fromEntries(res.patients.map((p) => [p['Per Nbr'], p]));
  const p = by['52'];
  assert.equal(p.Route, 'verification');
  assert.match(p['Route Reason'], /^two insurances: chart Kaiser Medicare Northern CA \(Auto-Low\) and last encounter Blue Shield/);
  assert.match(p.Flags, /Two insurances sent/);
  assert.deepEqual(p.x.verifyRows, [
    { payerId: '94135', payerName: 'Kaiser Foundation Health Plan of Northern California', memberId: 'KM52', source: 'guarantor', nextgenPayer: 'Kaiser Medicare Northern CA', confidence: 'Auto-Low' },
    { payerId: '94036', payerName: 'Blue Shield of California', memberId: 'BSL52', source: 'last encounter', nextgenPayer: 'Blue Shield CA', confidence: 'Fin Class (report)' },
  ]);
  assert.equal(by['53'].Route, 'discovery');
  assert.equal(by['54'].Route, 'discovery');
  assert.ok(res.review.some((r) => r['Per Nbr'] === '52' && r.Issue === 'Two insurances sent to verification'));

  const ex = C.buildExports(res.patients, CFG);
  const v52 = ex.verification.filter((r) => r.id === '52');
  assert.equal(v52.length, 2, 'one verification row per insurance, same id');
  assert.deepEqual(v52.map((r) => [r.payerId, r.memberId]), [['94135', 'KM52'], ['94036', 'BSL52']]);
  for (const r of ex.verification) assert.deepEqual(Object.keys(r), C.VERIFICATION_COLS);
  assert.ok(ex.discovery.some((r) => r.id === '53'));
});

test('R4: VIP verifies other insurance when there is one, else discovery only', () => {
  const rows = [...TITLE, HDR, BS('60'), BS('61'),
    base('62', { 'Payer Name (Cht Guar)': 'VIP', 'Fin Class (Cht Guar)': 'VIP', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': 'BSL62' }),
    base('63', { 'Payer Name (Cht Guar)': 'VIP', 'Fin Class (Cht Guar)': 'Partial VIP', 'Fin Class (Lst Enc)': 'Partial VIP' }),
    base('64', { 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Fin Class (Cht Guar)': 'Partial VIP', 'Sec Payer (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'BS64' }),
    // VIP chart entry with a secondary: the report has no secondary policy number, so discovery with the secondary as the hint.
    base('65', { 'Payer Name (Cht Guar)': 'VIP', 'Fin Class (Cht Guar)': 'VIP', 'Sec Payer (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'VIP65' }),
    ['Total Appointments (6)']];
  const res = C.buildPatients(rows, CFG);
  const by = Object.fromEntries(res.patients.map((p) => [p['Per Nbr'], p]));
  assert.ok(!res.autoMapped.some((a) => a.nextgenPayerName === 'VIP'), 'VIP is not auto-matched');

  assert.equal(by['62'].Route, 'verification');
  assert.equal(by['62'].x.vip, true);
  assert.match(by['62'].Flags, /VIP/);
  assert.deepEqual(by['62'].x.verifyRows.map((v) => [v.source, v.payerId, v.memberId]), [['last encounter', '94036', 'BSL62']]);
  assert.match(by['62']['Route Reason'], /^VIP \(MHS-sponsored\); verifying other insurance: last encounter Blue Shield CA/);

  assert.equal(by['63'].Route, 'discovery');
  assert.equal(by['63'].x.vip, true);
  assert.equal(by['63']['Sohar payerId'], '');
  assert.match(by['63']['Route Reason'], /no other insurance on file: discovery only/);

  assert.equal(by['64'].Route, 'verification', 'partial VIP in Fin Class: the real chart payer is the other insurance');
  assert.deepEqual(by['64'].x.verifyRows.map((v) => [v.source, v.payerId, v.memberId, v.nextgenPayer]), [['guarantor', '94036', 'BS64', 'Blue Shield CA']]);
  assert.match(by['64'].Flags, /VIP/);

  assert.equal(by['65'].Route, 'discovery', 'never send the VIP policy number to the secondary payer');
  assert.equal(by['65']['Sohar payerId'], '62308');
  assert.deepEqual(by['65'].x.verifyRows, []);
  assert.match(by['65']['Route Reason'], /secondary Cigna has no policy number in the report: discovery/);
});

// Sohar verification output layout (subset). Synthetic.
const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'memberId', 'payerName', 'status', 'statusReason', 'planName', 'innCopay', 'innTotalDeductible', 'innRemainingDeductible'];
const ver = (o) => VER_HDR.map((k) => (o[k] == null ? '' : String(o[k])));

test('mergeSoharResults keeps both verification rows and marks the better one', () => {
  const rows = [...TITLE, HDR, BS('50'), BS('51'),
    base('52', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM52', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': 'BSL52' }),
    base('55', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM55', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': 'BSL55' }),
    ['Total Appointments (4)']];
  const cfg = { ...CFG, payerMap: [...MAP, ...C.buildPatients(rows, CFG).autoMapped] };
  const v = [VER_HDR,
    ver({ requestId: 'q-1', id: '52', resolved: '2026-10-14T01:00:00Z', memberId: 'KM52', payerName: 'Kaiser Foundation Health Plan of Northern California', status: 'complete.ineligible', statusReason: 'INACTIVE' }),
    ver({ requestId: 'q-2', id: '52', resolved: '2026-10-14T01:00:00Z', memberId: 'BSL52', payerName: 'Blue Shield of California', status: 'complete.eligible', planName: 'SYNTH PPO', innCopay: '25', innTotalDeductible: '500', innRemainingDeductible: '100' }),
    ver({ requestId: 'q-3', id: '55', resolved: '2026-10-14T01:00:00Z', memberId: 'KM55', payerName: 'Kaiser Foundation Health Plan of Northern California', status: 'error.member.id' }),
    ver({ requestId: 'q-4', id: '55', resolved: '2026-10-14T01:00:00Z', memberId: 'BSL55', payerName: 'Blue Shield of California', status: 'complete.ineligible' }),
  ];
  const res = C.mergeSoharResults(rows, [{ name: 'ver.csv', rows: v }], cfg);
  assert.equal(res.error, undefined);
  const r52 = res.results.filter((r) => r['Per Nbr'] === '52');
  assert.equal(r52.length, 2, 'both rows kept');
  const good = r52.find((r) => r['Sohar Ref'] === 'q-2'), bad = r52.find((r) => r['Sohar Ref'] === 'q-1');
  assert.match(good.Action, /^Eligible on Blue Shield of California, the last encounter insurance\. Two insurances were checked; make this the chart insurance in NextGen \(choose "Blue Shield CA"\)\./);
  assert.equal(good['Choose in NextGen'], 'Blue Shield CA');
  assert.equal(good['NextGen Payer'], 'Blue Shield CA');
  assert.equal(good['NextGen Member ID'], 'BSL52');
  assert.equal(good['Sohar payerId'], '94036');
  assert.equal(good['Work Queue'], 'Action needed');
  assert.match(bad.Action, /^Superseded: complete\.ineligible \(INACTIVE\) on Kaiser/);
  assert.equal(bad['Work Queue'], 'Clear');
  assert.equal(bad['NextGen Payer'], 'Kaiser Medicare Northern CA');
  assert.equal(good.x.verifyRows.length, 2, 'result rows carry the patient x');
  assert.equal(res.counts.twoInsurances, 2);

  const r55 = res.results.filter((r) => r['Per Nbr'] === '55');
  for (const r of r55) assert.match(r.Action, /Two insurances were checked; neither came back eligible\. This patient is on the discovery import \(from verification\)\.$/);
  assert.deepEqual(res.verificationDiscovery.map((r) => r.id), ['55'], 'both insurances failed: one discovery row');
  assert.ok(!res.verificationDiscovery.some((r) => r.id === '52'), 'one eligible insurance keeps 52 out');

  const tr = C.trackerExport(res.results, '');
  assert.equal(tr.rows.find((r) => r['Per Nbr'] === '52')['Sohar Ref'], 'q-2', 'tracker export uses the eligible insurance');
  const tasks = C.buildTasks(res.results).filter((t) => t['Per Nbr'] === '52').map((t) => t.Task).sort();
  assert.deepEqual(tasks, ['Update benefits in NextGen', 'Update insurance in NextGen'], 'no task for the superseded row');
  for (const r of res.results) assert.deepEqual(Object.keys(r).filter((k) => !['x', 'resultKey', 'chooseGuess'].includes(k)).sort(), [...C.RESULTS_COLS].sort());
});

test('two insurances: an old eligible result on another policy does not override a fresh ineligible one', () => {
  const rows = [...TITLE, HDR, BS('50'), BS('51'), ['Total Appointments (2)']];
  const v = [VER_HDR,
    ver({ requestId: 'q-3', id: '50', resolved: '2026-08-01T01:00:00Z', memberId: 'OLD50', payerName: 'Blue Shield of California', status: 'complete.eligible', planName: 'OLD PPO' }),
    ver({ requestId: 'q-4', id: '50', resolved: '2026-10-14T01:00:00Z', memberId: 'BS50', payerName: 'Blue Shield of California', status: 'complete.ineligible' }),
  ];
  const res = C.mergeSoharResults(rows, [{ name: 'ver.csv', rows: v }], CFG);
  const r50 = res.results.filter((r) => r['Per Nbr'] === '50');
  assert.equal(res.counts.twoInsurances, 0, 'August is history, not a second insurance');
  assert.ok(r50.every((r) => !/^Superseded|Two insurances/.test(r.Action)));
  assert.notEqual(r50.find((r) => r['Sohar Ref'] === 'q-4')['Work Queue'], 'Clear');
  assert.equal(C.trackerExport(res.results, '').rows.find((r) => r['Per Nbr'] === '50')['Sohar Ref'], 'q-4');
});

test('autoMatchPayer: an unmatched brand word that names another payer is never Auto-Medium', () => {
  const S = [...SOHAR, { payerId: 'WH1', payerName: 'Western Health Advantage', aliasPayerNames: '' }, { payerId: 'OB1', payerName: 'Optum Behavioral Health', aliasPayerNames: '' }];
  const m = C.autoMatchPayer('Western Health Advantage Optum', S);
  assert.equal(m.payerId, 'WH1');
  assert.equal(m.confidence, 'Auto-Low');
  assert.match(m.evidence, /"optum" is in another payer's name/);
  assert.equal(C.autoMatchPayer('Western Health Advantage Plan', S).confidence, 'Auto-High');
});

test('trackerExport: Secondary Payer and Cost-share Flags trail the old columns', () => {
  assert.deepEqual(C.TRACKER_EXPORT_COLS.slice(-5), ['Service', 'Source', 'Exported', 'Secondary Payer', 'Cost-share Flags']);
  const rows = [...TITLE, HDR, BS('50'), ['Total Appointments (1)']];
  const v = [VER_HDR, ver({ requestId: 'q-1', id: '50', resolved: '2026-10-14T01:00:00Z', memberId: 'BS50', payerName: 'Blue Shield of California', status: 'complete.eligible' })];
  const res = C.mergeSoharResults(rows, [{ name: 'ver.csv', rows: v }], CFG);
  const withPills = res.results.map((r) => ({ ...r, Pills: 'Re-verify; QMB; Secondary may cover' }));
  const out = C.trackerExport(withPills, '', new Map([['50', 'QMB Noridian Southern CA']])).rows[0];
  assert.equal(out['Secondary Payer'], 'QMB Noridian Southern CA');
  assert.equal(out['Cost-share Flags'], 'QMB; Secondary may cover');
  assert.equal(C.trackerExport(res.results, '').rows[0]['Secondary Payer'], '');
});
