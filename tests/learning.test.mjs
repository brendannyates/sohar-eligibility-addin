import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
const C = createRequire(import.meta.url)('../core.js');

// Payer learning (v0.10). Synthetic data only.
const HDR = ['Res Name', 'Pat Sts', 'Pat Name', 'Appt Date', 'Beg Tm', 'Event', 'Loc Name', 'MRN', 'Status', 'Birth Dt', 'Rendering', 'Per Nbr', 'Guar Zip',
  'Sec Payer (Lst Enc)', 'Fin Class (Lst Enc)', 'Pol Nbr (Lst Enc)', 'Payer Name (Cht Guar)', 'Sec Payer (Cht Guar)', 'Fin Class (Cht Guar)', 'Pol Nbr (Cht Guar)'];
function row(o) {
  const r = new Array(HDR.length).fill('');
  for (const [k, v] of Object.entries(o)) { assert.ok(HDR.includes(k), k); r[HDR.indexOf(k)] = v; }
  return r;
}
const TITLE = [[], ['', '', 'Mindful Health Solutions'], []];
const base = (per, extra) => row({ 'Per Nbr': per, 'MRN': `M${per}`, 'Pat Name': `Test${per}, Pat`, 'Birth Dt': '1/2/1980', 'Appt Date': '10/20/2026', 'Beg Tm': 0.375,
  'Event': 'Telepsych', 'Loc Name': 'San Rafael', 'Status': 'Scheduled', 'Rendering': 'Doe, Jan MD', 'Guar Zip': '94901', ...extra });
const SOHAR = [
  { payerId: '94036', payerName: 'Blue Shield of California', aliasPayerNames: '' },
  { payerId: '62308', payerName: 'Cigna', aliasPayerNames: '' },
  { payerId: '60054', payerName: 'Aetna', aliasPayerNames: '' },
  { payerId: '94135', payerName: 'Kaiser Foundation Health Plan of Northern California', aliasPayerNames: '' },
  { payerId: '01111', payerName: 'Medicare of California', aliasPayerNames: '' },
  { payerId: 'WH1', payerName: 'Western Health Advantage', aliasPayerNames: '' },
];
const MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
  { nextgenPayerName: 'Synthetic Review Plan', payerId: '60054', payerName: 'Aetna', confidence: 'Review' },
];
const CFG = { asOf: new Date(Date.UTC(2026, 9, 13)), scope: 'upcoming', locationState: { 'San Rafael': 'CA' }, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP, soharPayers: SOHAR };
const BS = (per) => base(per, { 'Payer Name (Cht Guar)': 'Blue Shield CA', 'Pol Nbr (Cht Guar)': `BS${per}`, 'Fin Class (Cht Guar)': 'Blue Shield', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': `BS${per}` });

const VER_HDR = ['requestId', 'id', 'resolved', 'memberId', 'payerName', 'status', 'statusReason'];
const ver = (o) => VER_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const DISC_HDR = ['requestId', 'id', 'resolved', 'firstName', 'lastName', 'dateOfBirth', 'memberId', 'responsePayerId', 'responsePayerName'];
const disc = (o) => DISC_HDR.map((k) => (o[k] == null ? '' : String(o[k])));

// One ledger row. Kind V or D.
let seq = 0;
const out = (o) => ({
  Key: `k${++seq}`, 'Per Nbr': o.per, Kind: o.kind === 'D' ? 'Discovery' : 'Verification', 'NextGen Payer': o.name, 'Fin Class': o.fc || '',
  Source: o.src || (o.kind === 'D' ? 'discovery' : 'chart'), 'Sent payerId': o.sent || '', 'Sent Confidence': '', 'Sohar Status': o.status || '',
  'Returned payerId': o.ret || '', 'Returned Payer': o.retName || '', Signal: o.sig, 'Sohar Date': o.date || '2026-10-01T10:00:00Z', Recorded: 'test',
});
const L = C.LEARN_DEFAULTS;

test('learningConfig: defaults, Off, and shares as 50%, 0.5 or 50', () => {
  assert.deepEqual(C.learningConfig({}), L);
  assert.equal(C.learningConfig({ Learning: 'Off' }).on, false);
  assert.equal(C.learningConfig({ Learning: 'on' }).on, true);
  for (const v of ['50%', 0.5, '0.5', 50, '50']) assert.equal(C.learningConfig({ 'Learn: demote discovery share': v }).demoteDiscoveryShare, 0.5, String(v));
  assert.equal(C.learningConfig({ 'Learn: suggest share': '' }).suggestShare, 0.8);
  assert.equal(C.learningConfig({ 'Learn: promote patients': '5' }).promotePatients, 5);
  assert.equal(C.learningConfig({ 'Learn: promote patients': 'x' }).promotePatients, 3);
  assert.equal(C.learningConfig({ 'Learn: promote patients': 0 }).promotePatients, 3, 'at least one patient');
  assert.deepEqual(C.LEARNING_CONFIG_ROWS.map((r) => r.setting), ['Learning', 'Learn: promote patients', 'Learn: demote payer errors',
    'Learn: demote discovery share', 'Learn: demote min patients', 'Learn: suggest patients', 'Learn: suggest share']);
  assert.ok(C.LEARNING_CONFIG_ROWS.every((r) => r.what.length > 20));
});

test('Learned verifies and is confident; labels: automatic vs human; who needs classification', () => {
  assert.ok(C.VERIFIABLE.has('learned') && C.CONFIDENT.has('learned'));
  const idx = C.payerMapIndex([{ nextgenPayerName: 'Plan L', payerId: '9', payerName: 'Nine', confidence: 'Learned' }]);
  assert.equal(C.lookupPayer('Plan L', idx).status, 'mapped');
  for (const c of ['Auto-High', 'auto-medium', 'Auto-Low', 'Learned']) assert.equal(C.isAutoLabel(c), true, c);
  for (const c of ['Exact', 'Probable', 'Confirmed', 'Not a payer', 'Review', 'Checked by Ana', '']) assert.equal(C.isAutoLabel(c), false, c);
  assert.equal(C.needsClassification('Review', '1'), true);
  assert.equal(C.needsClassification('Auto-Low', '1'), true);
  assert.equal(C.needsClassification('Checked by Ana', '1'), true, 'unrecognized label');
  assert.equal(C.needsClassification('', '1'), true, 'blank label');
  assert.equal(C.needsClassification('Exact', ''), true, 'verifiable label without a payerId');
  assert.equal(C.needsClassification('Exact', '1'), false);
  assert.equal(C.needsClassification('Learned', '1'), false);
  assert.equal(C.needsClassification('Not a payer', ''), false);
});

test('autoMatchPayer: confirmed Payer Map names are aliases of their payerId (exact and fuzzy)', () => {
  const map = [...MAP, { nextgenPayerName: 'Synthetic Gold Network', payerId: 'WH1', payerName: 'Western Health Advantage', confidence: 'Confirmed' },
    { nextgenPayerName: 'Synthetic Silver Network', payerId: '60054', payerName: 'Aetna', confidence: 'Auto-Medium' }];
  const confirmed = C.confirmedNames(map);
  assert.deepEqual(confirmed.map((c) => c.name), ['Blue Shield CA', 'Cigna', 'Synthetic Gold Network'], 'Exact/Probable/Confirmed/Learned only');
  const ex = C.autoMatchPayer('SYNTHETIC GOLD NETWORK, INC.', SOHAR, { confirmed });
  assert.deepEqual([ex.payerId, ex.confidence, ex.evidence], ['WH1', 'Auto-High', 'Same as confirmed name "Synthetic Gold Network"']);
  assert.equal(C.autoMatchPayer('Synthetic Gold Network', SOHAR).confidence, 'Auto-Low', 'without the map the name is unknown');
  const fz = C.autoMatchPayer('Synthetic Gold Network PPO', SOHAR, { confirmed });
  assert.equal(fz.payerId, 'WH1', 'confirmed name is a fuzzy candidate');
  assert.match(fz.evidence, /in confirmed "Synthetic Gold Network"/);
  // A Sohar payerId missing from the Sohar list still works through the confirmed name.
  const extra = C.autoMatchPayer('Local Synthetic Trust', SOHAR, { confirmed: [{ name: 'Local Synthetic Trust', payerId: 'ZZ9', payerName: 'Synthetic Trust' }] });
  assert.deepEqual([extra.payerId, extra.confidence], ['ZZ9', 'Auto-High']);
});

test('buildPatients: confirmed names teach auto-matching of new names', () => {
  const map = [...MAP, { nextgenPayerName: 'Synthetic Gold Network', payerId: 'WH1', payerName: 'Western Health Advantage', confidence: 'Confirmed' }];
  const rows = [...TITLE, HDR, base('1', { 'Payer Name (Cht Guar)': 'Synthetic Gold Network Inc', 'Pol Nbr (Cht Guar)': 'G1' }), ['Total Appointments (1)']];
  const res = C.buildPatients(rows, { ...CFG, payerMap: map });
  assert.deepEqual(res.autoMapped.map((a) => [a.nextgenPayerName, a.payerId, a.confidence]), [['Synthetic Gold Network Inc', 'WH1', 'Auto-High']]);
  assert.equal(res.patients[0].Route, 'verification');
});

test('buildPatients: Review and unrecognized labels get the High manual-classification item; unmapped fix says step 1; Learned verifies', () => {
  const map = [...MAP, { nextgenPayerName: 'Odd Label Plan', payerId: '62308', payerName: 'Cigna', confidence: 'Checked by Ana' },
    { nextgenPayerName: 'Learned Plan', payerId: '62308', payerName: 'Cigna', confidence: 'Learned' }];
  const rows = [...TITLE, HDR,
    base('1', { 'Payer Name (Cht Guar)': 'Synthetic Review Plan', 'Pol Nbr (Cht Guar)': 'R1' }),
    base('2', { 'Payer Name (Cht Guar)': 'Odd Label Plan', 'Pol Nbr (Cht Guar)': 'O2' }),
    base('3', { 'Payer Name (Cht Guar)': 'Totally Unknown Synthetic Payer', 'Pol Nbr (Cht Guar)': 'U3' }),
    base('4', { 'Payer Name (Cht Guar)': 'Learned Plan', 'Pol Nbr (Cht Guar)': 'L4' }),
    ['Total Appointments (4)']];
  const res = C.buildPatients(rows, { ...CFG, payerMap: map, soharPayers: undefined });
  const issues = (per) => res.review.filter((r) => r['Per Nbr'] === per).map((r) => r.Issue);
  assert.ok(issues('1').includes('Payer needs manual classification'));
  assert.match(res.review.find((r) => r['Per Nbr'] === '1' && r.Issue === 'Payer needs manual classification').Detail, /\(confidence: Review\)/);
  assert.ok(issues('2').includes('Payer needs manual classification'));
  assert.equal(res.patients.find((p) => p['Per Nbr'] === '1').Route, 'discovery', 'routing unchanged');
  const un = res.review.find((r) => r['Per Nbr'] === '3' && r.Issue === 'Unmapped payer');
  assert.match(un.Fix, /^Map it in step 1/);
  const p4 = res.patients.find((p) => p['Per Nbr'] === '4');
  assert.equal(p4.Route, 'verification');
  assert.match(p4['Route Reason'], /^mapped payer \+ member ID/);
});

test('learned Fin Class: drives the last-encounter payer as "Fin Class (learned)"; Config overrides still win', () => {
  // Chart payer unmapped; the last encounter's Fin Class "Synthetic Fin" is not learnable from this report.
  const rows = [...TITLE, HDR,
    base('9', { 'Payer Name (Cht Guar)': 'Totally Unknown Synthetic Payer', 'Pol Nbr (Cht Guar)': 'U9', 'Fin Class (Cht Guar)': 'Other', 'Fin Class (Lst Enc)': 'Synthetic Fin', 'Pol Nbr (Lst Enc)': 'SF9' }),
    ['Total Appointments (1)']];
  const learned = [{ finClass: 'Synthetic Fin', payerId: '60054', payerName: 'Aetna', nextgenPayer: 'Aetna PPO', patients: 3 }];
  const res = C.buildPatients(rows, { ...CFG, soharPayers: undefined, learnedFinClasses: learned });
  const p = res.patients[0];
  assert.equal(p.Route, 'verification');
  assert.deepEqual(p.x.verifyRows, [{ payerId: '60054', payerName: 'Aetna', memberId: 'SF9', source: 'last encounter', nextgenPayer: 'Aetna PPO', confidence: 'Fin Class (learned)' }]);
  const ov = C.buildPatients(rows, { ...CFG, soharPayers: undefined, learnedFinClasses: learned, finClassPayers: { 'Synthetic Fin': '62308' } });
  assert.deepEqual([ov.patients[0].x.verifyRows[0].payerId, ov.patients[0].x.verifyRows[0].confidence], ['62308', 'Fin Class (config)']);
  assert.equal(C.buildPatients(rows, { ...CFG, soharPayers: undefined }).patients[0].Route, 'discovery', 'nothing learned: discovery as before');
});

test('learnFinClasses: one payerId agreeing on enough patients with no payer errors; ambiguous Fin Classes are never learned', () => {
  const v = (per, fc, sent, sig) => out({ per, kind: 'V', name: 'Proxy Payer', fc, src: 'last encounter', sent, sig, retName: `Payer ${sent}` });
  const rows = [
    v('1', 'Fin A', 'P1', 'agree'), v('2', 'Fin A', 'P1', 'agree'), v('3', 'Fin A', 'P1', 'agree'),
    v('1', 'Fin B', 'P1', 'agree'), v('2', 'Fin B', 'P1', 'agree'), v('3', 'Fin B', 'P1', 'agree'), v('4', 'Fin B', 'P1', 'payer error'),
    v('1', 'Fin C', 'P1', 'agree'), v('2', 'Fin C', 'P1', 'agree'), v('3', 'Fin C', 'P1', 'agree'), v('4', 'Fin C', 'P2', 'agree'),
    v('1', 'Fin D', 'P1', 'agree'), v('1', 'Fin D', 'P1', 'agree'), v('2', 'Fin D', 'P1', 'agree'),
    out({ per: '9', kind: 'V', name: 'Chart Payer', fc: 'Fin E', src: 'chart', sent: 'P1', sig: 'agree' }),
  ];
  const got = C.learnFinClasses(rows, L);
  assert.deepEqual(got, [{ finClass: 'Fin A', payerId: 'P1', payerName: 'Payer P1', nextgenPayer: 'Proxy Payer', patients: 3 }]);
});

// ---------- the ledger ----------

const LEDGER_ROWS = () => [...TITLE, HDR, BS('50'), BS('51'),
  // Two insurances: Auto-Low chart guess (Kaiser NorCal) + last encounter Blue Shield (learned from the report).
  base('52', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM52', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Blue Shield', 'Pol Nbr (Lst Enc)': 'BSL52' }),
  base('53', { 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'C53' }),
  // Discovery: Auto-Low guess Kaiser NorCal (no guess sent), same last encounter.
  base('70', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM70', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Medicare Advantage', 'Pol Nbr (Lst Enc)': 'KM70' }),
  base('71', { 'Payer Name (Cht Guar)': 'Kaiser Medicare Northern CA', 'Pol Nbr (Cht Guar)': 'KM71', 'Fin Class (Cht Guar)': 'Medicare Advantage', 'Fin Class (Lst Enc)': 'Medicare Advantage', 'Pol Nbr (Lst Enc)': 'KM71' }),
  base('72', {}), base('73', {}),
  ['Total Appointments (8)']];
function ledgerMerge(extraVer, extraDisc) {
  const rows = LEDGER_ROWS();
  const cfg = { ...CFG, payerMap: [...MAP, ...C.buildPatients(rows, CFG).autoMapped] };
  const v = [VER_HDR,
    ver({ requestId: 'q-1', id: '52', resolved: '2026-10-14T01:00:00Z', memberId: 'KM52', payerName: 'Kaiser Foundation Health Plan of Northern California', status: 'complete.ineligible' }),
    ver({ requestId: 'q-2', id: '52', resolved: '2026-10-14T01:00:00Z', memberId: 'BSL52', payerName: 'Blue Shield of California', status: 'complete.eligible' }),
    ver({ requestId: 'q-3', id: '50', resolved: '2026-10-14T01:00:00Z', memberId: 'BS50', payerName: 'Blue Shield of California', status: 'error.payer', statusReason: 'Invalid payer ID' }),
    ver({ requestId: 'q-4', id: '51', resolved: '2026-10-14T01:00:00Z', memberId: 'BS51', payerName: 'Blue Shield of California', status: 'error.member.id' }),
    ver({ requestId: 'q-5', id: '53', resolved: '2026-10-14T01:00:00Z', memberId: 'C53', payerName: 'Cigna', status: 'error', statusReason: 'Payer not responding' }),
    ...(extraVer || []),
  ];
  const d = [DISC_HDR,
    disc({ requestId: 'r-1', id: '70', resolved: '2026-10-14', responsePayerId: '94135', responsePayerName: 'Kaiser Foundation Health Plan of Northern California' }),
    disc({ requestId: 'r-2', id: '71', resolved: '2026-10-14', responsePayerId: '62308', responsePayerName: 'Cigna' }),
    disc({ requestId: 'r-3', id: '72', resolved: '2026-10-14', responsePayerId: 'WH1', responsePayerName: 'Western Health Advantage' }),
    disc({ requestId: 'r-4', id: '73', resolved: '2026-10-14' }),
    ...(extraDisc || []),
  ];
  return { res: C.mergeSoharResults(rows, [{ name: 'v.csv', rows: v }, { name: 'd.csv', rows: d }], cfg), cfg };
}

test('payerOutcomes: one row per Sohar outcome with the right signal, source and payer', () => {
  const { res } = ledgerMerge();
  assert.equal(res.error, undefined);
  const rows = C.payerOutcomes(res.results, res.patients, '10/20/2026 9:00 AM');
  for (const r of rows) assert.deepEqual(Object.keys(r), C.PAYER_OUTCOME_COLS);
  const by = (per, kind, extra) => rows.filter((r) => r['Per Nbr'] === per && r.Kind === kind && (!extra || extra(r)));
  const [chart52] = by('52', 'Verification', (r) => r.Source === 'chart');
  assert.deepEqual([chart52['NextGen Payer'], chart52['Sent payerId'], chart52['Sent Confidence'], chart52.Signal, chart52['Fin Class']],
    ['Kaiser Medicare Northern CA', '94135', 'Auto-Low', 'ineligible', 'Medicare Advantage']);
  const [le52] = by('52', 'Verification', (r) => r.Source === 'last encounter');
  assert.deepEqual([le52['NextGen Payer'], le52['Sent payerId'], le52['Sent Confidence'], le52.Signal, le52['Fin Class'], le52['Returned Payer']],
    ['Blue Shield CA', '94036', 'Fin Class (report)', 'agree', 'Blue Shield', 'Blue Shield of California'], 'proxy name and the Lst Enc Fin Class');
  assert.equal(by('50', 'Verification')[0].Signal, 'payer error');
  assert.equal(by('51', 'Verification')[0].Signal, 'member id error');
  assert.equal(by('53', 'Verification')[0].Signal, 'neutral', 'a payer that is not responding says nothing about the mapping');
  const d70 = by('70', 'Discovery')[0], d71 = by('71', 'Discovery')[0], d72 = by('72', 'Discovery')[0], d73 = by('73', 'Discovery')[0];
  assert.deepEqual([d70['Sent payerId'], d70['Returned payerId'], d70.Signal], ['94135', '94135', 'agree'], 'Auto-Low: compared with the guess');
  assert.deepEqual([d71['Returned payerId'], d71.Signal], ['62308', 'other payer']);
  assert.deepEqual([d72['NextGen Payer'], d72['Sent payerId'], d72.Signal], ['', '', 'found']);
  assert.equal(d73.Signal, 'neutral');
  assert.ok(rows.every((r) => r.Recorded === '10/20/2026 9:00 AM'));
  assert.ok(!rows.some((r) => r['Sohar Status'] === 'Not in Sohar file'));
  assert.match(d70.Key, /^d\|70\|2026-10-14$/);
});

test('appendOutcomes: re-uploading the same file adds nothing; a later re-check is a new outcome; old rows are kept as they were', () => {
  const { res } = ledgerMerge();
  const first = C.appendOutcomes([], C.payerOutcomes(res.results, res.patients, 'first'));
  assert.equal(first.added, first.rows.length);
  const again = C.appendOutcomes(first.rows, C.payerOutcomes(res.results, res.patients, 'second'));
  assert.equal(again.added, 0);
  assert.ok(again.rows.every((r) => r.Recorded === 'first'), 'the first recording wins');
  const later = ledgerMerge(null, [disc({ requestId: 'r-9', id: '70', resolved: '2026-10-19', responsePayerId: '94135', responsePayerName: 'Kaiser Foundation Health Plan of Northern California' })]).res;
  const third = C.appendOutcomes(again.rows, C.payerOutcomes(later.results, later.patients, 'third'));
  assert.equal(third.added, 1);
  assert.deepEqual(third.rows.slice(0, again.rows.length), again.rows);
  assert.equal(third.rows.at(-1)['Sohar Date'], '2026-10-19');
});

test('merge fixes: Coverage Differs uses the guess; guessed payers name the NextGen insurance and ask to confirm the mapping', () => {
  const { res } = ledgerMerge();
  const d = (per) => res.results.find((r) => r['Per Nbr'] === per && r['Sent As'] === 'Discovery');
  assert.equal(d('70')['Coverage Differs'], 'No', 'Sohar found the guessed payer');
  assert.equal(d('71')['Coverage Differs'], 'Yes');
  assert.equal(d('70')['Choose in NextGen'], 'Kaiser Medicare Northern CA');
  assert.match(d('70').Action, /^Coverage found matches NextGen\. Verify\. The Payer Map only guesses this NextGen insurance; confirm the payer mapping\./);
  assert.equal(d('72')['Choose in NextGen'], 'Not in Payer Map: Western Health Advantage (WH1)', 'no Payer Map row of any kind names it');
  const tasks = C.buildTasks(res.results);
  const t70 = tasks.filter((t) => t['Per Nbr'] === '70').map((t) => t.Task);
  assert.ok(t70.includes('Confirm payer mapping'));
  assert.ok(!t70.includes('Add payer to Payer Map'));
  assert.ok(tasks.some((t) => t['Per Nbr'] === '72' && t.Task === 'Add payer to Payer Map'));
  // The Action text alone (a re-read of the Sohar Results sheet has no non-column properties) still gives the task.
  const plain = res.results.map(({ chooseGuess, x, resultKey, ...r }) => r);
  assert.ok(C.buildTasks(plain).some((t) => t['Per Nbr'] === '70' && t.Task === 'Confirm payer mapping'));
  assert.ok(C.TASK_TYPES.includes('Confirm payer mapping'));
  // Review rows are guesses too: discovery finding Aetna (60054) names "Synthetic Review Plan".
  const r2 = ledgerMerge(null, [disc({ requestId: 'r-5', id: '73', resolved: '2026-10-15', responsePayerId: '60054', responsePayerName: 'Aetna' })]).res;
  const a73 = r2.results.find((r) => r['Per Nbr'] === '73' && r['Sohar payerId'] === '60054');
  assert.equal(a73['Choose in NextGen'], 'Synthetic Review Plan');
  assert.ok(C.buildTasks(r2.results).some((t) => t['Per Nbr'] === '73' && t.Task === 'Confirm payer mapping'));
});

// ---------- the rollup ----------

test('learnPayers: distinct patients, latest outcome per patient, score, and the five decisions', () => {
  const map = [
    { nextgenPayerName: 'Promote Me', payerId: 'P1', payerName: 'One', confidence: 'Auto-Medium' },
    { nextgenPayerName: 'Not Yet', payerId: 'P1', payerName: 'One', confidence: 'Auto-Low' },
    { nextgenPayerName: 'Demote Errors', payerId: 'P2', payerName: 'Two', confidence: 'Learned' },
    { nextgenPayerName: 'Demote Share', payerId: 'P3', payerName: 'Three', confidence: 'Auto-High' },
    { nextgenPayerName: 'Human Plan', payerId: 'P4', payerName: 'Four', confidence: 'Exact' },
    { nextgenPayerName: 'Suggest Me', payerId: 'P5', payerName: 'Five', confidence: 'Auto-Low' },
    { nextgenPayerName: 'Review Plan', payerId: '', payerName: '', confidence: 'Review' },
  ];
  const rows = [
    // 3 distinct patients agree (patient 1 twice, counted once).
    out({ per: '1', name: 'Promote Me', sent: 'P1', sig: 'agree' }), out({ per: '1', name: 'Promote Me', sent: 'P1', sig: 'agree' }),
    out({ per: '2', name: 'PROMOTE ME', sent: 'P1', sig: 'agree' }), out({ per: '3', kind: 'D', name: 'Promote Me', sent: 'P1', ret: 'P1', sig: 'agree' }),
    // A later neutral result (member ID error, nothing found) does not erase the agreement.
    out({ per: '2', name: 'Promote Me', sent: 'P1', sig: 'member id error', date: '2026-10-05' }), out({ per: '3', kind: 'D', name: 'Promote Me', sent: 'P1', sig: 'neutral', date: '2026-10-05' }),
    // Two agree and one conflict: never promoted. Patient 9's later agree replaces their earlier payer error.
    out({ per: '7', name: 'Not Yet', sent: 'P1', sig: 'agree' }), out({ per: '8', name: 'Not Yet', sent: 'P1', sig: 'agree' }),
    out({ per: '9', name: 'Not Yet', sent: 'P1', sig: 'payer error', date: '2026-09-01' }), out({ per: '9', name: 'Not Yet', sent: 'P1', sig: 'agree', date: '2026-09-20' }),
    out({ per: '10', kind: 'D', name: 'Not Yet', sent: 'P1', ret: 'Q9', retName: 'Nine', sig: 'other payer' }),
    // Payer errors on two patients.
    out({ per: '1', name: 'Demote Errors', sent: 'P2', sig: 'payer error' }), out({ per: '2', name: 'Demote Errors', sent: 'P2', sig: 'payer error' }),
    out({ per: '3', name: 'Demote Errors', sent: 'P2', sig: 'agree' }),
    // Discovery: 2 of 3 found another payer.
    ...['1', '2'].map((per) => out({ per, kind: 'D', name: 'Demote Share', sent: 'P3', ret: 'Q3', retName: 'Other Three', sig: 'other payer' })),
    out({ per: '3', kind: 'D', name: 'Demote Share', sent: 'P3', ret: 'P3', sig: 'agree' }),
    // A human label with the same problem is flagged, never changed.
    out({ per: '1', name: 'Human Plan', sent: 'P4', sig: 'payer error' }), out({ per: '2', name: 'Human Plan', sent: 'P4', sig: 'payer error' }),
    // Discovery keeps returning Q5 for an Auto-Low name.
    ...['1', '2', '3', '4'].map((per) => out({ per, kind: 'D', name: 'Suggest Me', sent: 'P5', ret: 'Q5', retName: 'Q Five', sig: 'other payer' })),
    out({ per: '5', kind: 'D', name: 'Suggest Me', sent: 'P5', ret: 'P5', sig: 'agree' }),
    // A Review row with no payerId and an unmapped name: "found" outcomes feed suggestions.
    ...['1', '2', '3'].map((per) => out({ per, kind: 'D', name: 'Review Plan', ret: 'R1', retName: 'Arr One', sig: 'found' })),
    ...['1', '2', '3'].map((per) => out({ per, kind: 'D', name: 'Brand New Name', ret: 'N1', retName: 'En One', sig: 'found' })),
    out({ per: '4', kind: 'D', name: 'Brand New Name', ret: '', sig: 'neutral' }),
    // VIP and Self Pay never learn.
    out({ per: '1', name: 'VIP', sent: 'P1', sig: 'agree' }), out({ per: '1', name: 'Self Pay', sent: 'P1', sig: 'agree' }),
  ];
  const got = new Map(C.learnPayers(rows, map, L).map((r) => [r.nextgenPayer, r]));
  const pm = got.get('Promote Me');
  assert.deepEqual([pm.agree, pm.conflicts, pm.n, pm.score, pm.decision, pm.note], [3, 0, 3, 80, 'promote', 'Learned from Sohar: 3 patients agreed']);
  const ny = got.get('Not Yet');
  assert.deepEqual([ny.agree, ny.payerErrors, ny.otherPayer, ny.conflicts, ny.decision], [3, 0, 1, 1, ''], 'latest outcome per patient');
  assert.equal(ny.score, Math.round(100 * 4 / 6));
  const de = got.get('Demote Errors');
  assert.deepEqual([de.payerErrors, de.decision], [2, 'demote']);
  assert.match(de.note, /^Demoted by Sohar learning \(was Learned\): 2 patients with a Sohar payer error/);
  const ds = got.get('Demote Share');
  assert.deepEqual([ds.discovery, ds.otherPayer, ds.agree, ds.decision, ds.top.payerId, ds.top.n], [3, 2, 1, 'demote', 'Q3', 2]);
  const hp = got.get('Human Plan');
  assert.equal(hp.decision, 'flag');
  assert.match(hp.note, /^Sohar disagrees with Payer Map: 2 patients with a Sohar payer error/);
  const sm = got.get('Suggest Me');
  assert.deepEqual([sm.suggestion && sm.suggestion.payerId, sm.suggestion && sm.suggestion.n, sm.decision], ['Q5', 4, ''], '4 of 5 = 80%');
  assert.equal(got.get('Review Plan').suggestion.payerId, 'R1');
  const bn = got.get('Brand New Name');
  assert.deepEqual([bn.mapped, bn.discovery, bn.suggestion.payerId], [false, 3, 'N1'], 'nothing-found results do not count as discovery checked');
  assert.ok(!got.has('VIP') && !got.has('Self Pay'));
  assert.equal(C.learningEvidence(ny), 'Sohar: 3 agreed, 1 other payer, score 67');
  assert.equal(C.learningEvidence(de), 'Sohar: 1 agreed, 2 payer errors, score 40');
  assert.equal(C.learningEvidence({ n: 0 }), '');
  assert.equal(C.learningEvidence({ n: 2, agree: 0, conflicts: 0, discovery: 0 }), 'Sohar: 2 patients checked, nothing that confirms or contradicts the payer yet');
  // Thresholds come from Config.
  const strict = new Map(C.learnPayers(rows, map, { ...L, promotePatients: 4, suggestShare: 0.9 }).map((r) => [r.nextgenPayer, r]));
  assert.equal(strict.get('Promote Me').decision, '');
  assert.equal(strict.get('Suggest Me').suggestion, null);
});

test('planLearning: promote and demote automatic labels, re-score Auto-Low upward, never touch human labels; Off changes nothing', () => {
  const map = [
    { nextgenPayerName: 'Promote Me', payerId: 'P1', payerName: 'One', confidence: 'Auto-Medium', notes: 'Auto-matched by the add-in: x' },
    { nextgenPayerName: 'Demote Errors', payerId: 'P2', payerName: 'Two', confidence: 'Auto-High' },
    { nextgenPayerName: 'Human Plan', payerId: 'P4', payerName: 'Four', confidence: 'Confirmed' },
    // Re-score candidates: the confirmed name "Synthetic Gold Network" now makes this an exact match.
    { nextgenPayerName: 'Synthetic Gold Network', payerId: 'WH1', payerName: 'Western Health Advantage', confidence: 'Confirmed' },
    { nextgenPayerName: 'Synthetic Gold Network Inc', payerId: '', payerName: '', confidence: 'Auto-Low' },
    // Would re-score to Auto-High (exact name), but learning demoted it earlier: left alone.
    { nextgenPayerName: 'Blue Shield CA Plan', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Auto-Low', notes: 'Demoted by Sohar learning (was Auto-High): ... | old' },
    { nextgenPayerName: 'Blue Shield California', payerId: '', payerName: '', confidence: 'Auto-Low' },
    { nextgenPayerName: 'Synthetic Gold Network LLC', payerId: '', payerName: '', confidence: 'Review' },
  ];
  const rows = [
    ...['1', '2', '3'].map((per) => out({ per, name: 'Promote Me', sent: 'P1', sig: 'agree' })),
    ...['1', '2'].map((per) => out({ per, name: 'Demote Errors', sent: 'P2', sig: 'payer error' })),
    ...['1', '2'].map((per) => out({ per, name: 'Human Plan', sent: 'P4', sig: 'payer error' })),
    // Blue Shield California has a Sohar conflict: re-score must leave it alone.
    out({ per: '5', kind: 'D', name: 'Blue Shield California', ret: '62308', retName: 'Cigna', sig: 'found' }),
  ];
  // The conflict above only counts against a payerId; give the row its old guess so the conflict is real.
  map[6].payerId = '94036';
  const plan = C.planLearning(rows, map, SOHAR, L);
  const ch = new Map(plan.changes.map((c) => [c.entry.nextgenPayerName, c]));
  assert.deepEqual([...ch.keys()].sort(), ['Demote Errors', 'Promote Me', 'Synthetic Gold Network Inc']);
  assert.deepEqual([ch.get('Promote Me').entry.confidence, ch.get('Promote Me').by, ch.get('Promote Me').old.confidence], ['Learned', 'Learning', 'Auto-Medium']);
  assert.deepEqual([ch.get('Demote Errors').entry.confidence, ch.get('Demote Errors').entry.payerId], ['Auto-Low', 'P2'], 'demote keeps the payerId');
  const rs = ch.get('Synthetic Gold Network Inc');
  assert.deepEqual([rs.by, rs.entry.payerId, rs.entry.confidence, rs.evidence], ['Re-score', 'WH1', 'Auto-High', 'Same as confirmed name "Synthetic Gold Network"']);
  assert.deepEqual(plan.counts, { promoted: 1, demoted: 1, rescored: 1, suggestions: 0, flagged: 1 });
  const after = new Map(plan.rollup.map((r) => [r.nextgenPayer, r]));
  assert.equal(after.get('Promote Me').confidence, 'Learned', 'rollup reflects the changes');
  assert.equal(after.get('Human Plan').confidence, 'Confirmed');
  assert.equal(after.get('Demote Errors').decision, '', 'a demoted row is Auto-Low now: no promotion while conflicts stay');
  assert.match(plan.payerMap.find((r) => r.nextgenPayerName === 'Promote Me').notes, /^Learned from Sohar: 3 patients agreed \| Auto-matched/);
  // A demoted row is never re-scored back up, even by the next plan.
  const next = C.planLearning(rows, plan.payerMap, SOHAR, L);
  assert.ok(!next.changes.some((c) => c.entry.nextgenPayerName === 'Demote Errors'));
  assert.ok(!next.changes.length, 'stable: nothing flips back');
  // Learning Off: no changes, but flags still show.
  const off = C.planLearning(rows, map, SOHAR, { ...L, on: false });
  assert.deepEqual(off.changes, []);
  assert.deepEqual(off.finClasses, []);
  assert.equal(off.counts.flagged, 1);
  const review = C.learningReviewRows(off.rollup);
  assert.deepEqual(review.map((r) => [r.Severity, r.Issue, r['Per Nbr']]), [['High', 'Sohar disagrees with Payer Map', '']]);
  assert.match(review[0].Detail, /^Human Plan -> Four \(P4, Confirmed\): 2 patients with a Sohar payer error/);
});

test('payerLearningRows: one row per payer, no patient identifiers', () => {
  const rows = [...['1', '2', '3'].map((per) => out({ per, kind: 'D', name: 'Suggest Me', sent: 'P5', ret: 'Q5', retName: 'Q Five', sig: 'other payer' }))];
  const map = [{ nextgenPayerName: 'Suggest Me', payerId: 'P5', payerName: 'Five', confidence: 'Auto-Low' }, { nextgenPayerName: 'Quiet', payerId: 'P6', payerName: 'Six', confidence: 'Exact' }];
  const sheet = C.payerLearningRows(C.learnPayers(rows, map, L));
  assert.deepEqual(sheet.map((r) => Object.keys(r)), [C.PAYER_LEARNING_COLS, C.PAYER_LEARNING_COLS]);
  assert.ok(!C.PAYER_LEARNING_COLS.includes('Per Nbr'));
  const s = sheet[0];
  assert.deepEqual([s['NextGen Payer'], s['Evidence score'], s['Discovery checked'], s['Discovery other payer'], s['Most common other payer'], s.Suggestion],
    ['Suggest Me', '20', '3', '3', 'Q Five (Q5): 3 of 3', 'Use Q Five (Q5): 3 of 3 discovery patients']);
  assert.match(s['Last outcome'], /^10\/01\/2026 Discovery: other payer$/);
  assert.deepEqual([sheet[1]['NextGen Payer'], sheet[1]['Evidence score'], sheet[1]['Last outcome']], ['Quiet', '', '']);
});

// ---------- the change log ----------

test('mapChangeLog and sheetEditLog: baseline, add-in writes, sheet edits, no-ops skipped', () => {
  const map = [{ nextgenPayerName: 'A Plan', payerId: '1', confidence: 'Exact' }, { nextgenPayerName: 'B Plan', payerId: '2', confidence: 'Auto-Low' }];
  const base = C.sheetEditLog(map, [], '10/20/2026 9:00 AM');
  assert.deepEqual(base.map((r) => [r['NextGen Payer'], r.By, r['New payerId'], r['New Confidence'], r['Old payerId']]), [['A Plan', 'Baseline', '1', 'Exact', ''], ['B Plan', 'Baseline', '2', 'Auto-Low', '']]);
  for (const r of base) assert.deepEqual(Object.keys(r), C.PAYER_MAP_LOG_COLS);
  const writes = C.mapChangeLog(map, [
    { entry: { nextgenPayerName: 'B Plan', payerId: '2', confidence: 'Learned' }, by: 'Learning', score: 80, evidence: 'Learned from Sohar: 3 patients agreed' },
    { entry: { nextgenPayerName: 'A Plan', payerId: '1', confidence: 'exact' }, by: 'Pane' },
    { entry: { nextgenPayerName: 'C Plan', payerId: '3', confidence: 'Auto-Medium' }, by: 'Auto-match', score: 0.81, evidence: 'Matched; 4 patients in this report' },
  ], '10/20/2026 9:01 AM');
  assert.deepEqual(writes.map((r) => [r['NextGen Payer'], r['Old Confidence'], r['New Confidence'], r.By, r.Score]),
    [['B Plan', 'Auto-Low', 'Learned', 'Learning', '80'], ['C Plan', '', 'Auto-Medium', 'Auto-match', '0.81']], 'unchanged write not logged');
  const log = [...base, ...writes];
  const now = [{ nextgenPayerName: 'A Plan', payerId: '9', confidence: 'Exact' }, { nextgenPayerName: 'B Plan', payerId: 2, confidence: 'learned' },
    { nextgenPayerName: 'C Plan', payerId: '3', confidence: 'Auto-Medium' }, { nextgenPayerName: 'D Plan', payerId: '4', confidence: 'Probable' }];
  const edits = C.sheetEditLog(now, log, '10/21/2026 8:00 AM');
  assert.deepEqual(edits.map((r) => [r['NextGen Payer'], r['Old payerId'], r['New payerId'], r.By, r.Evidence]),
    [['A Plan', '1', '9', 'Sheet edit', 'Changed on the Payer Map sheet'], ['D Plan', '', '4', 'Sheet edit', 'Added on the Payer Map sheet']]);
  assert.ok(!C.PAYER_MAP_LOG_COLS.includes('Per Nbr'));
});

test('mergeNotes: newest first, earlier kept, capped', () => {
  assert.equal(C.mergeNotes('new', 'old'), 'new | old');
  assert.equal(C.mergeNotes('new', ''), 'new');
  assert.equal(C.mergeNotes('', 'old'), 'old');
  assert.equal(C.mergeNotes('same', 'same | older'), 'same | older', 'repeating the latest note adds nothing');
  assert.equal(C.mergeNotes('n', 'x'.repeat(600)).length, 500);
});

// ---------- review fixes ----------

test('verificationSignal: outages and timeouts are neutral; a wrong or unknown payer is a payer error', () => {
  for (const s of ['error.payer.unavailable', 'error.other (payer timeout)', 'error (Payer not responding)', 'error.payer (Payer system down, retry later)', 'error.payer.timeout'])
    assert.equal(C.verificationSignal(s), 'neutral', s);
  for (const s of ['error.payer', 'error.payer (Invalid payer ID)', 'error (Payer not found)', 'error.other (Unknown payer)'])
    assert.equal(C.verificationSignal(s), 'payer error', s);
  assert.equal(C.verificationSignal('error.other (something else)'), 'neutral');
  assert.equal(C.verificationSignal('error.member.name (payer not found)'), 'neutral', 'member errors never count against the payer');
  assert.equal(C.verificationSignal('error.member.id'), 'member id error');
  assert.equal(C.verificationSignal('complete.eligible'), 'agree');
  assert.equal(C.verificationSignal('complete.ineligible'), 'ineligible');
});

test('learnPayers: last-encounter payer errors test the Fin Class guess, not the proxy name', () => {
  const le = (per, sig) => out({ per, kind: 'V', name: 'Partnership Synthetic', fc: 'Medi-Cal', src: 'last encounter', sent: 'P1', sig });
  const rows = [le('1', 'payer error'), le('2', 'payer error')];
  for (const confidence of ['Auto-High', 'Confirmed']) {
    const map = [{ nextgenPayerName: 'Partnership Synthetic', payerId: 'P1', payerName: 'One', confidence }];
    const r = C.learnPayers(rows, map, L)[0];
    assert.deepEqual([r.payerErrors, r.conflicts, r.decision], [0, 0, ''], confidence);
    assert.deepEqual(C.planLearning(rows, map, SOHAR, L).changes, [], `${confidence}: no demote`);
  }
  // The same outcomes still keep the Fin Class from being learned.
  assert.deepEqual(C.learnFinClasses([...rows, le('3', 'agree'), le('4', 'agree'), le('5', 'agree')], L), []);
  // Chart verifications of the name still count.
  const chart = ['1', '2'].map((per) => out({ per, kind: 'V', name: 'Partnership Synthetic', sent: 'P1', sig: 'payer error' }));
  assert.equal(C.learnPayers(chart, [{ nextgenPayerName: 'Partnership Synthetic', payerId: 'P1', payerName: 'One', confidence: 'Auto-High' }], L)[0].decision, 'demote');
});

test('planLearning: re-score never upgrades an Auto-Low row into a payerId Sohar contradicts (no flip-flop)', () => {
  const map = [
    { nextgenPayerName: 'Synthetic Gold Network', payerId: 'WH1', payerName: 'Western Health Advantage', confidence: 'Confirmed' },
    { nextgenPayerName: 'synthetic gold network.', payerId: '', payerName: '', confidence: 'Auto-Low' },
  ];
  const rows = ['1', '2', '3', '4'].map((per) => out({ per, kind: 'D', name: 'synthetic gold network.', ret: 'R9', retName: 'Arr Nine', sig: 'found' }));
  const plan = C.planLearning(rows, map, SOHAR, L);
  assert.deepEqual(plan.changes, [], 'discovery points to R9: no upgrade to WH1');
  const r = plan.rollup.find((x) => x.key === 'synthetic gold network.');
  assert.equal(r.suggestion.payerId, 'R9', 'the suggestion stays');
  assert.deepEqual(C.planLearning(rows, plan.payerMap, SOHAR, L).changes, [], 'stable on the next build');
  // One payer error against the candidate also blocks it (verification of WH1 under this name).
  const perr = [out({ per: '7', kind: 'V', name: 'synthetic gold network.', sent: 'WH1', sig: 'payer error' })];
  assert.deepEqual(C.planLearning(perr, map, SOHAR, L).changes, []);
  // Evidence that agrees with the candidate does not block it.
  const agree = [out({ per: '8', kind: 'D', name: 'synthetic gold network.', ret: 'WH1', retName: 'Western Health Advantage', sig: 'found' })];
  const ok = C.planLearning(agree, map, SOHAR, L);
  assert.deepEqual(ok.changes.map((c) => [c.by, c.entry.payerId, c.entry.confidence]), [['Re-score', 'WH1', 'Auto-High']]);
});

test('merge with a learned Fin Class records the outcome against that Fin Class and payerId', () => {
  const rows = [...TITLE, HDR,
    base('9', { 'Payer Name (Cht Guar)': 'Totally Unknown Synthetic Payer', 'Pol Nbr (Cht Guar)': 'U9', 'Fin Class (Cht Guar)': 'Other', 'Fin Class (Lst Enc)': 'Synthetic Fin', 'Pol Nbr (Lst Enc)': 'SF9' }),
    ['Total Appointments (1)']];
  const learned = [{ finClass: 'Synthetic Fin', payerId: '60054', payerName: 'Aetna', nextgenPayer: 'Aetna PPO', patients: 3 }];
  const files = [{ name: 'v.csv', rows: [VER_HDR, ver({ requestId: 'q-9', id: '9', resolved: '2026-10-14T01:00:00Z', memberId: 'SF9', payerName: 'Aetna', status: 'error.payer', statusReason: 'Invalid payer ID' })] }];
  const cfg = { ...CFG, soharPayers: undefined };
  const pick = (res) => { const o = C.payerOutcomes(res.results, res.patients, 't')[0]; return [o.Source, o['Sent payerId'], o['Fin Class'], o.Signal]; };
  assert.deepEqual(pick(C.mergeSoharResults(rows, files, { ...cfg, learnedFinClasses: learned })), ['last encounter', '60054', 'Synthetic Fin', 'payer error']);
  assert.deepEqual(pick(C.mergeSoharResults(rows, files, cfg)), ['chart', '', 'Other', 'payer error'], 'without them the outcome is lost to learning');
  // The pane passes them: every merge call in app.js must carry learnedFinClasses (directly or through the build's cfgFor()).
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  const calls = app.split('\n').filter((l) => l.includes('C.mergeSoharResults(') || l.includes('C.combineResults('));
  assert.ok(calls.length >= 2, 'processResults (combineResults) and the build\'s re-verify check');
  for (const l of calls) assert.match(l, /learnedFinClasses:|cfgFor\(\)/);
  assert.match(app, /const cfgFor = \(\) => \(\{[^}]*learnedFinClasses: state\.finClasses/);
});

test('duplicate Payer Map names: no learning, no suggestion, a Medium review row', () => {
  const map = [{ nextgenPayerName: 'Aetna PPO', payerId: '60054', payerName: 'Aetna', confidence: 'Auto-Medium' },
    { nextgenPayerName: 'aetna ppo', payerId: '60054', payerName: 'Aetna', confidence: 'Auto-Medium' }];
  const rows = ['1', '2', '3'].map((per) => out({ per, name: 'Aetna PPO', sent: '60054', sig: 'agree' }));
  const plan = C.planLearning(rows, map, SOHAR, L);
  assert.deepEqual(plan.changes, []);
  assert.equal(plan.counts.promoted, 0);
  const r = plan.rollup.find((x) => x.key === 'aetna ppo');
  assert.deepEqual([r.duplicate, r.duplicates, r.decision, r.suggestion], [true, 2, '', null]);
  assert.deepEqual(C.learningReviewRows(plan.rollup).map((x) => [x.Severity, x.Issue, x['Per Nbr']]), [['Medium', 'Payer Map name listed twice', '']]);
  assert.equal(C.planLearning(rows, [map[0]], SOHAR, L).counts.promoted, 1, 'one row: learned as usual');
});

test('FLAG: "Looks right" marks the disagreement reviewed until more conflicts arrive; the review is logged', () => {
  const human = (notes) => [{ nextgenPayerName: 'Blue Shield X', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Confirmed', notes }];
  const rows = ['1', '2', '3'].map((per) => out({ per, kind: 'D', name: 'Blue Shield X', ret: 'Q1', retName: 'Q One', sig: 'other payer' }));
  assert.equal(C.learnPayers(rows, human(''), L)[0].decision, 'flag');
  const ack = C.flagAckNote(3, '10/20/2026 9:00 AM');
  assert.equal(ack, 'Sohar disagreement reviewed 10/20/2026 9:00 AM, conflicts=3');
  const reviewed = C.learnPayers(rows, human(`${ack}; Checked in the add-in (Confirmed) | older note`), L)[0];
  assert.equal(reviewed.decision, '');
  assert.deepEqual(C.learningReviewRows([reviewed]), []);
  const more = [...rows, out({ per: '4', kind: 'D', name: 'Blue Shield X', ret: 'Q1', retName: 'Q One', sig: 'other payer' })];
  assert.equal(C.learnPayers(more, human(`${ack} | x`), L)[0].decision, 'flag', 'a new conflict brings the flag back');
  assert.equal(C.learnPayers(rows, human(`${C.flagAckNote(1, 'd')} | ${C.flagAckNote(9, 'older')}`), L)[0].decision, 'flag', 'the latest review counts');
  // The review is on the Payer Map Log even though label and payerId are unchanged.
  const log = C.mapChangeLog(human(''), [{ entry: { nextgenPayerName: 'Blue Shield X', payerId: '94036', confidence: 'Confirmed' }, by: 'Pane', evidence: ack, always: true }], 'now');
  assert.deepEqual(log.map((x) => [x.By, x['Old Confidence'], x['New Confidence'], x.Evidence]), [['Pane', 'Confirmed', 'Confirmed', ack]]);
});

test('pane state: the outcomes ledger and learned Fin Classes start unread (null), so the first build reads the ledger', () => {
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  const init = /const state = \{[\s\S]*?\};/.exec(app)[0];
  assert.match(init, /\boutcomes: null\b/);
  assert.match(init, /\bfinClasses: null\b/);
  assert.match(app, /if \(!state\.outcomes\) state\.outcomes = await X\.readPayerOutcomes\(\)/);
  assert.match(app, /if \(state\.finClasses\) return state\.finClasses;/);
});
