import test from 'node:test';
import assert from 'node:assert/strict';
import { isDeepStrictEqual, inspect } from 'node:util';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { createMock } from './helpers/excel-mock.mjs';
const C = createRequire(import.meta.url)('../core.js');
const REC = createRequire(import.meta.url)('../record.js');

// Runs the real excel.js against tests/helpers/excel-mock.mjs. Synthetic data only.
// The fail-safe sweeps fail a save at every op / sync / Excel close and check nothing is ever lost.
const TODO = undefined;

const uploads = (n, tag) => Array.from({ length: n }, (_, i) => ({
  Kind: i % 3 ? 'verification' : 'discovery', Key: `${tag}-${i}`, File: `${tag}.csv`, Uploaded: '2026-10-01T12:00:00.000Z',
  Record: JSON.stringify({ id: `${tag}-${i}`, memberId: `00${i}`, payerId: '01111', note: 'x'.repeat(40) }),
}));
const tasks = (n, tag) => Array.from({ length: n }, (_, i) => Object.fromEntries(C.TASK_COLS.map((c) =>
  [c, c === 'Done' ? (i % 2 ? 'Yes' : '') : c === 'MRN' ? `000${i}` : `${tag} ${c} ${i}`])));
const expected = (cols, rows) => rows.map((r) => Object.fromEntries(cols.map((c) => [c, r[c] == null ? '' : r[c]])));
// Payer learning tables (v0.10). All text.
const outcomes = (n, tag) => Array.from({ length: n }, (_, i) => Object.fromEntries(C.PAYER_OUTCOME_COLS.map((c) =>
  [c, c === 'MRN' ? `00${i}` : c === 'Sent payerId' ? '01111' : c === 'Key' ? `${tag}|${i}` : `${tag} ${c} ${i}`])));
const mapLog = (n, tag) => Array.from({ length: n }, (_, i) => Object.fromEntries(C.PAYER_MAP_LOG_COLS.map((c) =>
  [c, c === 'New payerId' ? '01111' : c === 'Score' ? '0.81' : `${tag} ${c} ${i}`])));
const learning = (n, tag) => Array.from({ length: n }, (_, i) => Object.fromEntries(C.PAYER_LEARNING_COLS.map((c) =>
  [c, c === 'Sohar payerId' ? '01111' : c === 'Evidence score' ? '83' : `${tag} ${c} ${i}`])));
// Sohar ref memory and Re-verify anyway requests. All text (Per Nbr keeps leading zeros).
const refs = (n, tag) => Array.from({ length: n }, (_, i) => ({ 'Sohar Ref': `${tag}-q-${i}`, 'MRN': `00${i}`, Kind: i % 2 ? 'discovery' : 'verification',
  'Result Key': `v|00${i}|synthetic payer|0${i}`, 'First Matched': '10/20/2026 9:00 AM' }));
const requests = (n, tag) => Array.from({ length: n }, (_, i) => ({ 'MRN': `0${i}`, Requested: `2026-10-20T1${i % 10}:00:00.000Z`, By: 'Clinic tab' }));
// Payer Credential Rules (R13): made-up payers, as rules.payerRuleRows writes them.
const PR = createRequire(import.meta.url)('../payer-rules.js');
const RL = createRequire(import.meta.url)('../rules.js');
const credRules = (n, tag) => Array.from({ length: n }, (_, i) => ({ Payer: `${tag} Payer ${Math.floor(i / 2)}`, State: i % 4 < 2 ? '' : 'WA', Credential: i % 2 ? 'MD' : 'NP',
  'Besse Approved': ['Yes', 'No', ''][i % 3], 'NP Fully Approved': i % 2 ? 'N/A' : 'No', 'Esketamine First Treatment': 'No', 'Esketamine Start': 'Yes', 'Esketamine Observation': 'Yes',
  'Match Names': `${tag} payer ${Math.floor(i / 2)}`, Notes: i % 5 ? '' : 'synthetic note' }));
// v0.14 Clinic Notes / Clinic Payer Notes from returned Clinic Eligibility files. Made-up MRNs and payers.
const clinicNotes = (n, tag) => Array.from({ length: n }, (_, i) => ({ MRN: `00900${i}`, Location: `${tag} Clinic`, Note: `${tag} note ${i}`, By: 'Front desk', When: `2026-10-22T1${i % 10}:00:00.000Z` }));
const clinicPayerNotes = (n, tag) => Array.from({ length: n }, (_, i) => ({ 'NextGen Payer': `${tag} Payer ${i}`, Location: `${tag} Clinic`, 'Card Says': `${tag} PPO ${i}`,
  Note: i % 3 ? '' : 'new card', By: 'Front desk', When: `2026-10-22T1${i % 10}:00:00.000Z` }));
// v0.20 Sohar Flow results: made-up runs, LASFIR-style tags, no names.
const flowRows = (n, tag) => Array.from({ length: n }, (_, i) => Object.fromEntries(REC.FLOW_SHEET_COLS.map((c) => [c, c === 'Key' ? `${tag}-run-${i}` : c === 'MRN' ? `00800${i}` : `${tag} ${c} ${i}`])));
const clinicApplied = (n, tag) => Array.from({ length: n }, (_, i) => ({ Location: `${tag} Clinic`, Kind: 'Task done', Key: `00900${i}|Fix member ID in NextGen`, When: `2026-10-22T1${i % 10}:00:00.000Z` }));

// ---------- happy path ----------

test('writeTasks creates the sheet and table; a rewrite replaces every row', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readTasks(), []);
  await X.writeTasks(C.TASK_COLS, tasks(7, 'old'));
  assert.deepEqual(m.sheetNames(), ['Import', 'NextGen Tasks']);
  assert.deepEqual(m.tableInfo('tblNextGenTasks'), { sheet: 'NextGen Tasks', address: 'A1:P8', rows: 7, cols: 16 }); // v0.14: Clinic Note last; v0.19: LASFIR for the two name columns
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(7, 'old')));
  assert.equal((await X.readTasks())[0].MRN, '0000', 'text format keeps leading zeros');
  assert.equal(m.frozenRows('NextGen Tasks'), 1);
  assert.equal(m.activeSheet(), 'Import', 'writeTasks does not activate');
  await X.writeTasks(C.TASK_COLS, tasks(3, 'new'));
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(3, 'new')));
  assert.equal(m.usedAddress('NextGen Tasks'), 'A1:P4', 'no leftover rows from the longer old table');
  assert.deepEqual(m.tableNames(), ['tblNextGenTasks']);
});

test('writeUploads keeps Sohar Uploads hidden and round-trips the JSON records', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readUploads(), []);
  await X.writeUploads(C.UPLOAD_COLS, uploads(5, 'a'));
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'VeryHidden');
  assert.deepEqual(await X.readUploads(), expected(C.UPLOAD_COLS, uploads(5, 'a')));
  await X.writeUploads(C.UPLOAD_COLS, uploads(2, 'b'));
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'VeryHidden');
  assert.deepEqual(await X.readUploads(), expected(C.UPLOAD_COLS, uploads(2, 'b')));
  await X.writeUploads(C.UPLOAD_COLS, []);
  assert.deepEqual(await X.readUploads(), []);
  assert.equal(m.tableInfo('tblSoharUploads').rows, 1, 'an empty table keeps one blank body row');
});

test('writeIrUploads / readIrUploads round-trip records on a hidden sheet', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readIrUploads(), []);
  const recs = [{ id: 'IR-1', text: 'Insurance card update' }, { id: 'IR-2', text: 'Payer changed' }];
  await X.writeIrUploads(recs);
  assert.equal(m.sheetVisibility('IR Uploads'), 'VeryHidden');
  assert.deepEqual(await X.readIrUploads(), recs);
});

test('writeBenefits keeps % columns numeric with 0% format and data bars', async () => {
  const m = createMock(), X = m.SoharExcel;
  const cols = C.BENEFIT_SHEET_COLS.concat('Pills');
  const rows = [{ 'Per Nbr': '100', 'OOP % Met': 0.5, 'Deductible % Met': 1 }, { 'Per Nbr': '101', 'OOP % Met': '', 'Deductible % Met': 0.25 }];
  await X.writeBenefits(cols, rows, C.BENEFIT_PCT_COLS);
  const got = m.tableRows('tblBenefits');
  assert.equal(got[0]['OOP % Met'], 0.5);
  assert.equal(got[1]['Deductible % Met'], 0.25);
  assert.equal(got[0]['Per Nbr'], '100');
  const j = cols.indexOf('OOP % Met');
  const fmts = m.numberFormats('Benefits', 'A1:AB3');
  assert.equal(fmts[0][j], '@', 'header is text');
  assert.equal(fmts[1][j], '0%');
  assert.equal(fmts[1][0], '@');
  const cfs = m.conditionalFormats('Benefits');
  assert.equal(cfs.length, 2);
  assert.ok(cfs.every((cf) => cf.type === 'DataBar' && cf.rules.lowerBoundRule.formula === '0' && cf.rules.upperBoundRule.formula === '1'));
  // A rebuild drops the old bars instead of stacking them.
  await X.writeBenefits(cols, rows, C.BENEFIT_PCT_COLS);
  assert.equal(m.conditionalFormats('Benefits').length, 2);
});

test('writeExport and writeResults activate their sheet', async () => {
  const m = createMock(), X = m.SoharExcel;
  await X.writeExport(['id', 'name'], [{ id: '1', name: 'Test' }]);
  assert.equal(m.activeSheet(), 'Export');
  await X.writeResults(['Per Nbr'], [{ 'Per Nbr': '0042' }]);
  assert.equal(m.activeSheet(), 'Sohar Results');
  assert.deepEqual(await X.readTable('tblSoharResults'), [{ 'Per Nbr': '0042' }]);
});

test('readDiscoveryVerify and readVerifyFailed return [] when missing and the rows when present', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readDiscoveryVerify(), []);
  assert.deepEqual(await X.readVerifyFailed(), []);
  const dv = [{ firstName: 'Ana', lastName: 'Test', memberId: '00912345', payerId: '01111', id: '100' }];
  await X.writeDiscoveryVerify(C.VERIFICATION_COLS, dv);
  assert.deepEqual(await X.readDiscoveryVerify(), expected(C.VERIFICATION_COLS, dv));
  m.seedTable('Discovery from Verify', 'tblVerifyFailed', C.DISCOVERY_COLS, [{ firstName: 'Sven', zipCode: '02134', id: '102' }]);
  assert.deepEqual(await X.readVerifyFailed(), expected(C.DISCOVERY_COLS, [{ firstName: 'Sven', zipCode: '02134', id: '102' }]));
});

test('readTable reads more than 500 rows in chunks, skips blank rows, and throws on a missing template table', async () => {
  const m = createMock(), X = m.SoharExcel;
  const rows = Array.from({ length: 1234 }, (_, i) => ({ A: `a${i}`, B: i % 5 ? `b${i}` : '' }));
  m.seedTable('Big', 'tblBig', ['A', 'B'], [...rows.slice(0, 600), { A: '', B: '' }, ...rows.slice(600)]);
  m.inject();
  assert.deepEqual(await X.readTable('tblBig'), expected(['A', 'B'], rows));
  const loads = m.valueLoads();
  assert.ok(loads.length >= 4 && Math.max(...loads) <= 500, `chunked loads: ${loads}`);
  await assert.rejects(X.readTable('tblMissing'), /tblMissing is missing/);
  assert.deepEqual(await X.readTable('tblMissing', true), []);
});

test('a stray or renamed table on an add-in sheet does not block the rebuild', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('NextGen Tasks', 'Table7', ['x', 'y'], [{ x: '1', y: '2' }]);
  await X.writeTasks(C.TASK_COLS, tasks(2, 'n'));
  assert.deepEqual(m.tableNames(), ['tblNextGenTasks']);
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(2, 'n')));
});

test('setTaskDone and tableRowCount work on the written table', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.equal(await X.tableRowCount('tblNextGenTasks'), 0);
  await X.writeTasks(C.TASK_COLS, tasks(3, 't'));
  assert.equal(await X.tableRowCount('tblNextGenTasks'), 3);
  await X.setTaskDone('t Task Key 0', true);
  assert.equal((await X.readTasks())[0].Done, 'Yes');
});

test('replaceTable works with Table.resize (1.13) and with the rows.add fallback', async () => {
  for (const ver of ['1.17', '1.12']) {
    const m = createMock({ sets: { ExcelApi: ver } }), X = m.SoharExcel;
    m.seedTable('Patients', 'tblPatients', ['Per Nbr', 'Zip'], [{ 'Per Nbr': '1', Zip: '02134' }, { 'Per Nbr': '2', Zip: '02135' }, { 'Per Nbr': '3', Zip: '0' }]);
    const rows = [{ 'Per Nbr': '0010', Zip: '01234' }, { 'Per Nbr': '0011', Zip: '' }];
    await X.replaceTable('tblPatients', ['Per Nbr', 'Zip'], rows);
    assert.deepEqual(await X.readTable('tblPatients'), expected(['Per Nbr', 'Zip'], rows), `ExcelApi ${ver}`);
    assert.equal(m.tableInfo('tblPatients').rows, 2, `ExcelApi ${ver}`);
    await assert.rejects(X.replaceTable('tblPatients', ['Per Nbr'], rows), /expects 1/);
  }
});

test('upsertPayerMapMany keeps payerIds as text, including rows added below', async () => {
  const m = createMock(), X = m.SoharExcel;
  const cols = ['NextGen Payer Name', 'Sohar payerId', 'Sohar payerName', 'Confidence', 'Action', 'Notes', 'Last reviewed'];
  m.seedTable('Payer Map', 'tblPayerMap', cols, [], { text: false });
  await X.upsertPayerMapMany([
    { nextgenPayerName: 'Plan A', payerId: '01111', payerName: 'A', confidence: 'Exact' },
    { nextgenPayerName: 'Plan B', payerId: '02222', payerName: 'B', confidence: 'Review' },
    { nextgenPayerName: 'Plan C', payerId: '03333', payerName: 'C', confidence: 'Exact' },
  ]);
  const map = await X.readPayerMap();
  assert.deepEqual(map.map((r) => r.payerId), ['01111', '02222', '03333']);
});

test('Payer Outcomes (hidden), Payer Map Log and Payer Learning round-trip as text; the log appends', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readPayerOutcomes(), []);
  assert.deepEqual(await X.readPayerMapLog(), []);
  await X.writePayerOutcomes(C.PAYER_OUTCOME_COLS, outcomes(4, 'a'));
  assert.equal(m.sheetVisibility('Payer Outcomes'), 'Hidden');
  assert.equal(m.autofits('Payer Outcomes'), 0, 'no autofit on the hidden ledger');
  assert.deepEqual(await X.readPayerOutcomes(), expected(C.PAYER_OUTCOME_COLS, outcomes(4, 'a')));
  assert.equal((await X.readPayerOutcomes())[0].MRN, '000', 'leading zeros kept');
  // Append-only through core.appendOutcomes: the old rows plus the new ones.
  const more = C.appendOutcomes(await X.readPayerOutcomes(), [...outcomes(2, 'a'), ...outcomes(3, 'b')]);
  assert.equal(more.added, 3);
  await X.writePayerOutcomes(C.PAYER_OUTCOME_COLS, more.rows);
  assert.deepEqual(await X.readPayerOutcomes(), expected(C.PAYER_OUTCOME_COLS, [...outcomes(4, 'a'), ...outcomes(3, 'b')]));
  assert.equal(m.sheetVisibility('Payer Outcomes'), 'Hidden');

  assert.equal(await X.appendPayerMapLog(C.PAYER_MAP_LOG_COLS, mapLog(2, 'first')), 2);
  assert.equal(await X.appendPayerMapLog(C.PAYER_MAP_LOG_COLS, []), 0);
  await X.appendPayerMapLog(C.PAYER_MAP_LOG_COLS, mapLog(1, 'second'));
  assert.deepEqual(await X.readPayerMapLog(), expected(C.PAYER_MAP_LOG_COLS, [...mapLog(2, 'first'), ...mapLog(1, 'second')]));
  assert.equal(m.sheetVisibility('Payer Map Log'), 'Visible');
  assert.ok(m.autofits('Payer Map Log') > 0, 'visible sheets are still autofit');
  assert.equal((await X.readPayerMapLog())[0]['New payerId'], '01111');

  await X.writePayerLearning(C.PAYER_LEARNING_COLS, learning(3, 'x'));
  assert.deepEqual(m.tableRows('tblPayerLearning'), expected(C.PAYER_LEARNING_COLS, learning(3, 'x')));
  await X.writePayerLearning(C.PAYER_LEARNING_COLS, learning(1, 'y'));
  assert.deepEqual(m.tableRows('tblPayerLearning'), expected(C.PAYER_LEARNING_COLS, learning(1, 'y')), 'rewritten each refresh');
  assert.equal(m.sheetVisibility('Payer Learning'), 'Visible');
  assert.equal(m.activeSheet(), 'Import', 'none of them activates');
  // Clearing the saved Sohar results leaves the ledger alone.
  await X.writeUploads(C.UPLOAD_COLS, []);
  assert.equal((await X.readPayerOutcomes()).length, 7);
});

test('Sohar Refs and Re-verify Requests (hidden) round-trip as text; refs survive clearing saved results', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readSoharRefs(), []);
  assert.deepEqual(await X.readReverifyRequests(), []);
  await X.writeSoharRefs(C.SOHAR_REF_COLS, refs(3, 'a'));
  assert.equal(m.sheetVisibility('Sohar Refs'), 'Hidden');
  assert.deepEqual(await X.readSoharRefs(), expected(C.SOHAR_REF_COLS, refs(3, 'a')));
  assert.equal((await X.readSoharRefs())[0].MRN, '000', 'leading zeros kept');
  // Append-only through core.appendRefs.
  const more = C.appendRefs(await X.readSoharRefs(), [...refs(1, 'a'), ...refs(2, 'b')], 'now');
  assert.equal(more.added, 2);
  await X.writeSoharRefs(C.SOHAR_REF_COLS, more.rows);
  assert.deepEqual(await X.readSoharRefs(), expected(C.SOHAR_REF_COLS, [...refs(3, 'a'), ...refs(2, 'b')]));
  await X.writeUploads(C.UPLOAD_COLS, []);
  assert.equal((await X.readSoharRefs()).length, 5, '"Clear saved Sohar results" keeps the ref memory');

  await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, C.addReverifyRequest([], '0042', '2026-10-20T16:00:00.000Z', 'Clinic tab'));
  assert.equal(m.sheetVisibility('Re-verify Requests'), 'Hidden');
  assert.deepEqual(await X.readReverifyRequests(), [{ 'MRN': '0042', Requested: '2026-10-20T16:00:00.000Z', By: 'Clinic tab' }]);
  const pr = C.pruneReverifyRequests(await X.readReverifyRequests(), [{ 'MRN': '0042', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-21' }]);
  await X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, pr.rows);
  assert.deepEqual(await X.readReverifyRequests(), [], 'dropped once a newer verification exists');
  assert.equal(m.activeSheet(), 'Import', 'neither activates');
});

test('readConfig: re-verify settings with defaults; ensureConfigRows adds the re-verify rows', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Config', 'tblConfig', ['Setting', 'Value', 'What it does'], [{ Setting: 'Re-verify days', Value: '10', 'What it does': 'mine' }]);
  m.seedTable('Location', 'tblLocationState', ['Location', 'State'], [{ Location: 'San Rafael', State: 'CA' }]);
  let cfg = await X.readConfig();
  assert.deepEqual(cfg.reverify, { skip: true, days: 10, newMonth: false });
  assert.equal(cfg.reverifyDays, 10, 'one setting shared by the skip rule and the clinic Re-verify flag');
  assert.deepEqual(await X.ensureConfigRows(C.REVERIFY_CONFIG_ROWS), ['Skip recent verifications', 'Re-verify in new month']);
  assert.equal(m.tableRows('tblConfig').find((r) => r.Setting === 'Re-verify days').Value, '10', 'existing row unchanged');
  await X.setConfigValue('Re-verify in new month', 'On');
  cfg = await X.readConfig();
  assert.equal(cfg.reverify.newMonth, true);
});

test('ensureConfigRows adds only missing settings, with the description, and never changes existing ones', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Config', 'tblConfig', ['Setting', 'Value', 'What it does'], [{ Setting: 'Scope', Value: 'All patients', 'What it does': 'x' },
    { Setting: 'Learn: promote patients', Value: '5', 'What it does': 'mine' }]);
  const added = await X.ensureConfigRows(C.LEARNING_CONFIG_ROWS);
  assert.deepEqual(added, C.LEARNING_CONFIG_ROWS.map((r) => r.setting).filter((s) => s !== 'Learn: promote patients'));
  const rows = m.tableRows('tblConfig');
  assert.deepEqual(rows.find((r) => r.Setting === 'Learn: promote patients'), { Setting: 'Learn: promote patients', Value: '5', 'What it does': 'mine' });
  assert.equal(rows.find((r) => r.Setting === 'Learning')['What it does'], C.LEARNING_CONFIG_ROWS[0].what);
  assert.deepEqual(await X.ensureConfigRows(C.LEARNING_CONFIG_ROWS), [], 'second run adds nothing');
  m.seedTable('Location', 'tblLocationState', ['Location', 'State'], [{ Location: 'San Rafael', State: 'CA' }]);
  const cfg = await X.readConfig();
  assert.equal(cfg.learning.on, true);
  assert.equal(cfg.learning.promotePatients, 5);
  assert.equal(cfg.learning.demoteDiscoveryShare, 0.5);
  assert.equal(cfg.learning.suggestShare, 0.8);
});

test('upsertPayerMapMany keeps earlier notes, stamps Last reviewed only for a person, and writes it as text', async () => {
  const m = createMock(), X = m.SoharExcel;
  const cols = ['NextGen Payer Name', 'Sohar payerId', 'Sohar payerName', 'Confidence', 'Action', 'Notes', 'Last reviewed', 'Owner'];
  m.seedTable('Payer Map', 'tblPayerMap', cols, [
    { 'NextGen Payer Name': 'Plan A', 'Sohar payerId': '01111', 'Sohar payerName': 'A', Confidence: 'Auto-Medium', Notes: 'Auto-matched by the add-in: x', 'Last reviewed': 46300, Owner: 'Ana' },
    { 'NextGen Payer Name': 'Plan B', 'Sohar payerId': '02222', 'Sohar payerName': 'B', Confidence: 'Auto-Low', Notes: 'old', 'Last reviewed': '09/01/2026' },
  ], { text: false });
  await X.upsertPayerMapMany([
    { nextgenPayerName: 'Plan A', payerId: '01111', payerName: 'A', confidence: 'Learned', notes: 'Learned from Sohar: 3 patients agreed' },
    { nextgenPayerName: 'Plan C', payerId: '03333', payerName: 'C', confidence: 'Auto-Low', notes: 'Auto-matched by the add-in: y' },
  ]);
  await X.upsertPayerMap({ nextgenPayerName: 'plan b', payerId: '02222', payerName: 'B', confidence: 'Confirmed', notes: 'Confirmed in the add-in' }, { human: true });
  const rows = m.tableRows('tblPayerMap');
  const a = rows.find((r) => r['NextGen Payer Name'] === 'Plan A'), b = rows.find((r) => r['NextGen Payer Name'] === 'plan b'), c = rows.find((r) => r['NextGen Payer Name'] === 'Plan C');
  assert.equal(a.Notes, 'Learned from Sohar: 3 patients agreed | Auto-matched by the add-in: x');
  assert.equal(a['Last reviewed'], C.fmtDate(C.excelSerialToDate(46300)), 'automatic write keeps the date, as text');
  assert.equal(a.Owner, 'Ana', 'other columns kept');
  assert.equal(a.Action, 'verification', 'Learned verifies');
  assert.equal(b['Last reviewed'], C.fmtDate(new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()))), 'a person\'s write stamps today');
  assert.equal(b.Notes, 'Confirmed in the add-in | old');
  assert.equal(c['Last reviewed'], '', 'new automatic rows leave it blank');
  assert.equal(c['Sohar payerId'], '03333');
  assert.equal(c.Action, 'discovery');
  const fm = m.numberFormats('Payer Map', 'G2:G4').map((r) => r[0]);
  assert.deepEqual(fm, ['@', '@', '@'], 'Last reviewed is always text');
});

test('writeImportRows / readImportRows round-trip text', async () => {
  const m = createMock(), X = m.SoharExcel;
  await X.writeImportRows([['Per Nbr', 'Zip'], ['001', '02134'], ['002']]);
  assert.deepEqual(await X.readImportRows(), [['Per Nbr', 'Zip'], ['001', '02134'], ['002', '']]);
});

// ---------- mock fidelity ----------

test('mock: deferred loads, dimension checks, overlap and name rules behave like Excel', async () => {
  const m = createMock();
  await m.Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem('Import');
    const rg = ws.getRange('A1:B2');
    assert.throws(() => rg.values, (e) => e.code === 'PropertyNotLoaded');
    const nul = ctx.workbook.tables.getItemOrNullObject('nope');
    assert.throws(() => nul.isNullObject, (e) => e.code === 'PropertyNotLoaded');
    assert.throws(() => ws.tables.items, (e) => e.code === 'PropertyNotLoaded');
    nul.load('isNullObject');
    await ctx.sync();
    assert.equal(nul.isNullObject, true);
    nul.delete();
    await assert.rejects(ctx.sync(), (e) => e.code === 'ItemNotFound');
    rg.values = [['a', 'b']];
    await assert.rejects(ctx.sync(), (e) => e.code === 'InvalidArgument');
    rg.numberFormat = [['@', 'General'], ['General', 'General']];
    rg.values = [['01', '01'], ['true', 'x']];
    rg.load('values');
    await ctx.sync();
    assert.deepEqual(rg.values, [['01', 1], [true, 'x']], 'General cells parse numbers and booleans; @ cells keep text');
    const t = ws.tables.add(rg, true);
    t.name = 'tblA';
    await ctx.sync();
    ws.tables.add(ws.getRange('B2:C3'), true);
    await assert.rejects(ctx.sync(), (e) => e.code === 'InvalidOperation');
    const t2 = ctx.workbook.worksheets.add('Other').tables.add('Other!A1:A2', true);
    t2.name = 'TBLA';
    await assert.rejects(ctx.sync(), /already exists/);
    t2.name = 'AB12';
    await assert.rejects(ctx.sync(), /table name/);
    ctx.workbook.worksheets.add('x'.repeat(32));
    await assert.rejects(ctx.sync(), /sheet name/);
  });
  assert.deepEqual(m.sheetNames(), ['Import', 'Other']);
  assert.deepEqual(m.tableNames(), ['tblA', 'Table1']);
});

test('mock: failAtOp keeps earlier ops (no rollback); killAfterSync stops the batch and closes Excel', async () => {
  const m = createMock();
  m.inject({ failAtOp: 2 });
  await assert.rejects(m.Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem('Import');
    ws.getRange('A1').values = [['one']];
    ws.getRange('A2').values = [['two']];
    ws.getRange('A3').values = [['three']];
    await ctx.sync();
  }), /failAtOp/);
  assert.deepEqual(m.values('Import', 'A1:A3'), [['one'], [''], ['']]);
  m.inject({ killAfterSync: 1 });
  let after = false;
  await assert.rejects(m.Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem('Import').getRange('B1').values = [['kept']];
    await ctx.sync();
    after = true;
  }), /closed/);
  assert.equal(after, false);
  assert.equal(m.values('Import', 'B1')[0][0], 'kept');
  await assert.rejects(m.SoharExcel.readTasks(), /closed/);
  m.inject();
  assert.deepEqual(await m.SoharExcel.readTasks(), []);
});

test('mock: APIs above the supported set fail at sync; payload cap rejects big requests', async () => {
  const m = createMock({ sets: { ExcelApi: '1.8' } });
  assert.equal(m.Office.context.requirements.isSetSupported('ExcelApi', '1.9'), false);
  await assert.rejects(m.Excel.run({ delayForCellEdit: true }, async (ctx) => {
    ctx.workbook.worksheets.getItem('Import').getRange('A1').copyFrom('Import!B1', 'All');
    await ctx.sync();
  }), (e) => e.code === 'ApiNotFound');
  m.inject({ payloadLimit: 1000 });
  await assert.rejects(m.SoharExcel.writeTasks(['a'], Array.from({ length: 200 }, (_, i) => ({ a: `row ${i}` }))), /payload/);
});

// ---------- fail-safe sweeps ----------

const CASES = {
  uploads: { write: (X, rows) => X.writeUploads(C.UPLOAD_COLS, rows), read: (X) => X.readUploads(), cols: C.UPLOAD_COLS, make: uploads, sheet: 'Sohar Uploads', hidden: true },
  tasks: { write: (X, rows) => X.writeTasks(C.TASK_COLS, rows), read: (X) => X.readTasks(), cols: C.TASK_COLS, make: tasks, sheet: 'NextGen Tasks', hidden: false },
  outcomes: { write: (X, rows) => X.writePayerOutcomes(C.PAYER_OUTCOME_COLS, rows), read: (X) => X.readPayerOutcomes(), cols: C.PAYER_OUTCOME_COLS, make: outcomes, sheet: 'Payer Outcomes', hidden: true },
  mapLog: { write: (X, rows) => X.writePayerMapLog(C.PAYER_MAP_LOG_COLS, rows), read: (X) => X.readPayerMapLog(), cols: C.PAYER_MAP_LOG_COLS, make: mapLog, sheet: 'Payer Map Log', hidden: false },
  learning: { write: (X, rows) => X.writePayerLearning(C.PAYER_LEARNING_COLS, rows), read: (X) => X.readTable('tblPayerLearning', true), cols: C.PAYER_LEARNING_COLS, make: learning, sheet: 'Payer Learning', hidden: false },
  refs: { write: (X, rows) => X.writeSoharRefs(C.SOHAR_REF_COLS, rows), read: (X) => X.readSoharRefs(), cols: C.SOHAR_REF_COLS, make: refs, sheet: 'Sohar Refs', hidden: true },
  requests: { write: (X, rows) => X.writeReverifyRequests(C.REVERIFY_REQUEST_COLS, rows), read: (X) => X.readReverifyRequests(), cols: C.REVERIFY_REQUEST_COLS, make: requests, sheet: 'Re-verify Requests', hidden: true },
  credRules: { write: (X, rows) => X.writePayerCredRules(PR.HEADERS, rows), read: (X) => X.readPayerCredRules(), cols: PR.HEADERS, make: credRules, sheet: 'Payer Credential Rules', hidden: false },
  clinicNotes: { write: (X, rows) => X.writeClinicNotes(RL.CLINIC_NOTE_COLS, rows), read: (X) => X.readClinicNotes(), cols: RL.CLINIC_NOTE_COLS, make: clinicNotes, sheet: 'Clinic Notes', hidden: false },
  clinicPayerNotes: { write: (X, rows) => X.writeClinicPayerNotes(RL.CLINIC_PAYER_NOTE_COLS, rows), read: (X) => X.readClinicPayerNotes(), cols: RL.CLINIC_PAYER_NOTE_COLS, make: clinicPayerNotes, sheet: 'Clinic Payer Notes', hidden: false },
  flow: { write: (X, rows) => X.writeFlowResults(REC.FLOW_SHEET_COLS, rows), read: (X) => X.readFlowResults(), cols: REC.FLOW_SHEET_COLS, make: flowRows, sheet: 'Flow Results', hidden: true },
  clinicApplied: { write: (X, rows) => X.writeClinicWorkApplied(RL.CLINIC_APPLIED_COLS, rows), read: (X) => X.readClinicWorkApplied(), cols: RL.CLINIC_APPLIED_COLS, make: clinicApplied, sheet: 'Clinic Work Applied', hidden: true },
};
const KNOBS = { op: 'failAtOp', sync: 'failAtSync', kill: 'killAfterSync' };
const short = (v) => (v instanceof Error ? `Error: ${v.message}` : Array.isArray(v) ? `${v.length} rows` : inspect(v));

/**
 * For every op (or sync) k of a successful rewrite: fail there, then read back with injection off.
 * The reader must give exactly OLD or exactly NEW (with no old table: [] or NEW), never partial or empty-over-OLD.
 * A writer that resolves must have saved NEW, and a later clean write must still work.
 */
async function sweep(kase, mode, nOld, nNew, viaRecover) {
  const { write, read, cols, make, sheet, hidden } = CASES[kase];
  const OLD = nOld == null ? null : make(nOld, 'old'), NEW = make(nNew, 'new'), NEXT = make(4, 'next');
  const setup = async () => { const m = createMock(); if (OLD) await write(m.SoharExcel, OLD); return m; };
  const m0 = await setup();
  m0.inject();
  await write(m0.SoharExcel, NEW);
  const total = mode === 'op' ? m0.opCount() : m0.syncCount();
  await write(m0.SoharExcel, NEXT);
  const visibleAfter = m0.visibleSheets();
  assert.ok(total > 0);
  const bad = [];
  for (let k = 1; k <= total; k++) {
    const m = await setup(), X = m.SoharExcel;
    m.inject({ [KNOBS[mode]]: k });
    let resolved = true;
    try { await write(X, NEW); } catch (e) { resolved = false; }
    m.inject();
    // viaRecover: the pane is reopened (recoverAll) before anything reads.
    if (viaRecover) { try { await X.recoverAll(); } catch (e) { bad.push(`${KNOBS[mode]}=${k}: recoverAll threw ${e.message}`); } }
    let got; try { got = await read(X); } catch (e) { got = e; }
    const isNew = isDeepStrictEqual(got, expected(cols, NEW));
    const isOld = OLD ? isDeepStrictEqual(got, expected(cols, OLD)) : isDeepStrictEqual(got, []);
    if (!(isNew || (!resolved && isOld))) bad.push(`${KNOBS[mode]}=${k}: writer ${resolved ? 'resolved' : 'rejected'}, reader got ${short(got)}`);
    try {
      await write(X, NEXT);
      if (!isDeepStrictEqual(await read(X), expected(cols, NEXT))) bad.push(`${KNOBS[mode]}=${k}: the next clean write did not save`);
      if (hidden && m.sheetVisibility(sheet) === 'Visible') bad.push(`${KNOBS[mode]}=${k}: ${sheet} visible after the next clean write`);
      if (!isDeepStrictEqual(m.visibleSheets(), visibleAfter)) bad.push(`${KNOBS[mode]}=${k}: visible sheets ${m.visibleSheets()} after the next clean write`);
    } catch (e) { bad.push(`${KNOBS[mode]}=${k}: the next clean write threw ${e.message}`); }
  }
  assert.deepEqual(bad, [], `${kase} ${mode} sweep over ${total} points (OLD ${nOld ?? 'none'}, NEW ${nNew})`);
}

// NEW > 1000 rows so the write spans several chunks (a failure between chunks would leave a partial table).
test('fail-safe: Sohar Uploads (hidden) survives a failure at every op', { todo: TODO }, () => sweep('uploads', 'op', 4, 1100));
test('fail-safe: Sohar Uploads (hidden) survives a failed sync at every point', { todo: TODO }, () => sweep('uploads', 'sync', 4, 1100));
test('fail-safe: Sohar Uploads (hidden) survives Excel closing after every sync', { todo: TODO }, () => sweep('uploads', 'kill', 4, 1100));
test('fail-safe: NextGen Tasks survives a failure at every op', { todo: TODO }, () => sweep('tasks', 'op', 1100, 3));
test('fail-safe: NextGen Tasks survives a failed sync at every point', { todo: TODO }, () => sweep('tasks', 'sync', 1100, 3));
test('fail-safe: NextGen Tasks survives Excel closing after every sync', { todo: TODO }, () => sweep('tasks', 'kill', 1100, 3));
test('fail-safe: reopening the pane after any failure recovers or keeps the old rows (uploads, tasks)', async () => {
  await sweep('uploads', 'op', 4, 1100, true);
  await sweep('uploads', 'kill', 4, 1100, true);
  await sweep('tasks', 'op', 1100, 3, true);
  await sweep('tasks', 'kill', 1100, 3, true);
});
// Payer learning tables (v0.10): the ledger grows past one chunk; the log and the rollup are rewritten from fewer rows.
// Sohar Refs grows past one chunk (append-only); Re-verify Requests shrinks when requests are dropped.
for (const [kase, nOld, nNew] of [['outcomes', 4, 1100], ['mapLog', 1100, 1103], ['learning', 1100, 3], ['refs', 4, 1100], ['requests', 1100, 2], ['credRules', 6, 40],
  ['clinicNotes', 4, 1100], ['clinicPayerNotes', 30, 2], ['clinicApplied', 4, 1100]]) {
  test(`fail-safe: ${CASES[kase].sheet} survives a failure at every op, a failed sync, and Excel closing`, async () => {
    await sweep(kase, 'op', nOld, nNew);
    await sweep(kase, 'sync', nOld, nNew);
    await sweep(kase, 'kill', nOld, nNew);
  });
  test(`fail-safe: ${CASES[kase].sheet} after any failure, reopened or first-ever, recovers or keeps the old rows`, async () => {
    await sweep(kase, 'op', nOld, nNew, true);
    await sweep(kase, 'kill', nOld, nNew, true);
    await sweep(kase, 'op', null, nNew);
    await sweep(kase, 'kill', null, nNew);
  });
}

// First-ever write: nothing to lose, but a half-made sheet or table must not block the next save.
test('fail-safe: a failed first-ever write never blocks the next save (uploads)', async () => {
  await sweep('uploads', 'op', null, 1100);
  await sweep('uploads', 'kill', null, 1100);
});
test('fail-safe: a failed first-ever write never blocks the next save (tasks)', async () => {
  await sweep('tasks', 'op', null, 1100);
  await sweep('tasks', 'kill', null, 1100);
});

// ---------- interrupted saves and leftovers ----------

/** The state a save leaves when Excel closes mid-swap: old sheet gone, complete new rows as tblX_stg on the staging sheet. */
function cutOffMidSwap(m, table, cols, rows) {
  m.seedTable(`~saving ${table}`, `${table}_stg`, cols, rows, { hidden: true });
}

test('a save cut off mid-swap days ago: the next read puts the new rows in place and recoverAll reports it', async () => {
  const m = createMock(), X = m.SoharExcel;
  cutOffMidSwap(m, 'tblSoharUploads', C.UPLOAD_COLS, uploads(3, 'cut'));
  assert.deepEqual(await X.readUploads(), expected(C.UPLOAD_COLS, uploads(3, 'cut')), 'never [] while a complete copy exists');
  assert.deepEqual(m.tableNames(), ['tblSoharUploads']);
  assert.equal(m.tableInfo('tblSoharUploads').sheet, 'Sohar Uploads');
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'VeryHidden');
  assert.ok(!m.sheetNames().some((n) => n.startsWith('~saving')));
  assert.deepEqual(await X.recoverAll(), ['Sohar Uploads'], 'the pane can say what was put back');
  assert.deepEqual(await X.recoverAll(), [], 'reported once');
});

test('recoverAll finishes a cut-off swap, replaces a half-deleted old sheet, and sweeps unfinished staging', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.addSheet('NextGen Tasks', [['leftover']]); // old sheet whose table is already gone
  cutOffMidSwap(m, 'tblNextGenTasks', C.TASK_COLS, tasks(2, 'cut'));
  m.seedTable('Sohar Results', 'tblSoharResults', ['Per Nbr'], [{ 'Per Nbr': '1' }]);
  m.addSheet('~saving tblSoharResults', [['partial']], { hidden: true }); // unfinished: never trusted
  assert.deepEqual(await X.recoverAll(), ['NextGen Tasks']);
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(2, 'cut')));
  assert.equal(m.sheetVisibility('NextGen Tasks'), 'Visible');
  assert.deepEqual(await X.readTable('tblSoharResults'), [{ 'Per Nbr': '1' }], 'intact table untouched');
  assert.ok(!m.sheetNames().some((n) => n.startsWith('~saving')), m.sheetNames().join());
});

test('if the rename cannot be finished, reads still return the staged rows and a save refuses instead of overwriting', async () => {
  const m = createMock(), X = m.SoharExcel;
  cutOffMidSwap(m, 'tblNextGenTasks', C.TASK_COLS, tasks(2, 'cut'));
  m.inject({ failAtOp: 1 }); // the first change of the rescue fails
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(2, 'cut')));
  m.inject({ failAtOp: 1 });
  await assert.rejects(X.writeTasks(C.TASK_COLS, tasks(1, 'new')), /could not be finished. Nothing was overwritten/);
  m.inject();
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(2, 'cut')), 'staged rows survived the refused save');
  await X.writeTasks(C.TASK_COLS, tasks(1, 'new'));
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(1, 'new')));
});

test('setTaskDone and tableRowCount see a cut-off save', async () => {
  const m = createMock(), X = m.SoharExcel;
  cutOffMidSwap(m, 'tblNextGenTasks', C.TASK_COLS, tasks(3, 'cut'));
  assert.equal(await X.tableRowCount('tblNextGenTasks'), 3);
  await X.setTaskDone('cut Task Key 0', true);
  assert.equal((await X.readTasks())[0].Done, 'Yes');
});

test('a save keeps the sheet where it was in the tab order and its visibility', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('NextGen Tasks', 'tblNextGenTasks', C.TASK_COLS, tasks(2, 'old'));
  m.addSheet('Export', [['x']]);
  m.seedTable('Sohar Uploads', 'tblSoharUploads', C.UPLOAD_COLS, uploads(1, 'old'), { hidden: true });
  m.addSheet('Config', [['y']]);
  const order = m.sheetNames();
  await X.writeTasks(C.TASK_COLS, tasks(3, 'new'));
  await X.writeUploads(C.UPLOAD_COLS, uploads(2, 'new'));
  assert.deepEqual(m.sheetNames(), order);
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'VeryHidden');
  assert.equal(m.sheetVisibility('NextGen Tasks'), 'Visible');
});

test('the old table on a renamed sheet is replaced, not duplicated', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('My Tasks', 'tblNextGenTasks', C.TASK_COLS, tasks(2, 'old'));
  await X.writeTasks(C.TASK_COLS, tasks(1, 'new'));
  assert.deepEqual(m.tableNames(), ['tblNextGenTasks']);
  assert.equal(m.tableInfo('tblNextGenTasks').sheet, 'NextGen Tasks');
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(1, 'new')));
});

test('big Sohar Uploads saves stay under the web payload cap (chunks capped by size)', async () => {
  const m = createMock(), X = m.SoharExcel;
  const big = uploads(3000, 'big').map((u) => ({ ...u, Record: JSON.stringify({ id: u.Key, blob: 'y'.repeat(1500) }) }));
  m.inject({ payloadLimit: 5e6 });
  await X.writeUploads(C.UPLOAD_COLS, big);
  assert.equal((await X.readUploads()).length, 3000);
});

test('Benefits data bars and Export activation still apply after the swap', async () => {
  const m = createMock(), X = m.SoharExcel;
  const cols = C.BENEFIT_SHEET_COLS;
  m.seedTable('Benefits', 'tblBenefits', cols, [{ 'Per Nbr': 'old' }]);
  await X.writeBenefits(cols, [{ 'Per Nbr': '1', 'OOP % Met': 0.4 }], C.BENEFIT_PCT_COLS);
  assert.equal(m.conditionalFormats('Benefits').length, C.BENEFIT_PCT_COLS.length);
  await X.writeExport(['a'], [{ a: '1' }]);
  assert.equal(m.activeSheet(), 'Export');
});

// ---------- review fixes ----------

test('another Excel window sweeping this save mid-way never produces a false success or lost rows', async () => {
  // Pane A saves while pane B (a co-author, or the same file open on the web) opens and runs recoverAll.
  const bad = [];
  for (let delay = 0; delay < 60; delay++) {
    const m = createMock(), X = m.SoharExcel;
    await X.writeUploads(C.UPLOAD_COLS, uploads(3, 'old'));
    m.inject();
    const a = X.writeUploads(C.UPLOAD_COLS, uploads(5, 'new')).then(() => true, () => false);
    let t = 0;
    const b = (async () => { while (t++ < delay) await Promise.resolve(); return X.recoverAll(); })().catch(() => []);
    const [ok] = await Promise.all([a, b]);
    const got = await X.readUploads();
    const isNew = isDeepStrictEqual(got, expected(C.UPLOAD_COLS, uploads(5, 'new')));
    const isOld = isDeepStrictEqual(got, expected(C.UPLOAD_COLS, uploads(3, 'old')));
    if (ok && !isNew) bad.push(`delay ${delay}: save reported success but read ${got.length} rows`);
    if (!isNew && !isOld) bad.push(`delay ${delay}: read ${short(got)}`);
  }
  assert.deepEqual(bad, []);
});

test('a structure-protected workbook gets a clear message and keeps its data', async () => {
  const m = createMock(), X = m.SoharExcel;
  await X.writeTasks(C.TASK_COLS, tasks(2, 'old'));
  m.protectStructure(true);
  await assert.rejects(X.writeTasks(C.TASK_COLS, tasks(1, 'new')), /structure is protected; unprotect it \(Review > Protect Workbook\)/);
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(2, 'old')));
  assert.ok(!m.sheetNames().some((n) => n.startsWith('~saving')));
  m.protectStructure(false);
  await X.writeTasks(C.TASK_COLS, tasks(1, 'new'));
  assert.deepEqual(await X.readTasks(), expected(C.TASK_COLS, tasks(1, 'new')));
});

test('saving the sheet the user is looking at keeps them on it', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('NextGen Tasks', 'tblNextGenTasks', C.TASK_COLS, tasks(2, 'old'));
  await X.activate('NextGen Tasks');
  assert.equal(m.activeSheet(), 'NextGen Tasks');
  await X.writeTasks(C.TASK_COLS, tasks(1, 'new'));
  assert.equal(m.activeSheet(), 'NextGen Tasks');
  await X.activate('Import');
  await X.writeTasks(C.TASK_COLS, tasks(1, 'new2'));
  assert.equal(m.activeSheet(), 'Import', 'not pulled onto a sheet they were not viewing');
});

test('Sohar Uploads is hidden again on save even if someone unhid it', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Sohar Uploads', 'tblSoharUploads', C.UPLOAD_COLS, uploads(1, 'old')); // visible
  await X.writeUploads(C.UPLOAD_COLS, uploads(2, 'new'));
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'VeryHidden');
});

// The swap paths for an old table on a renamed sheet, and a renamed old table with a clashing target sheet.
async function sweepSetup(seed, mode) {
  const OLD = tasks(3, 'old'), NEW = tasks(1100, 'new');
  const fresh = () => { const m = createMock(); seed(m, OLD); return m; };
  const m0 = fresh(); m0.inject(); await m0.SoharExcel.writeTasks(C.TASK_COLS, NEW);
  const total = mode === 'op' ? m0.opCount() : m0.syncCount();
  const bad = [];
  for (let k = 1; k <= total; k++) {
    const m = fresh(), X = m.SoharExcel;
    m.inject({ [KNOBS[mode]]: k });
    let resolved = true; try { await X.writeTasks(C.TASK_COLS, NEW); } catch (e) { resolved = false; }
    m.inject();
    let got; try { got = await X.readTasks(); } catch (e) { got = e; }
    const isNew = isDeepStrictEqual(got, expected(C.TASK_COLS, NEW)), isOld = isDeepStrictEqual(got, expected(C.TASK_COLS, OLD));
    if (!(isNew || (!resolved && isOld))) bad.push(`${KNOBS[mode]}=${k}: writer ${resolved ? 'resolved' : 'rejected'}, reader got ${short(got)}`);
    try { await X.writeTasks(C.TASK_COLS, tasks(2, 'next')); if (m.tableNames().length !== 1) bad.push(`${KNOBS[mode]}=${k}: tables ${m.tableNames()}`); }
    catch (e) { bad.push(`${KNOBS[mode]}=${k}: next write threw ${e.message}`); }
  }
  return bad;
}
test('fail-safe: an old table on a renamed sheet survives a failure at every op and Excel closing at every sync', async () => {
  const seed = (m, OLD) => m.seedTable('My Tasks', 'tblNextGenTasks', C.TASK_COLS, OLD);
  assert.deepEqual([...await sweepSetup(seed, 'op'), ...await sweepSetup(seed, 'kill')], []);
});
test('fail-safe: a renamed old table plus a clashing NextGen Tasks sheet survives every failure point', async () => {
  const seed = (m, OLD) => { m.seedTable('My Tasks', 'tblNextGenTasks', C.TASK_COLS, OLD); m.addSheet('NextGen Tasks', [['stray']]); };
  assert.deepEqual([...await sweepSetup(seed, 'op'), ...await sweepSetup(seed, 'kill')], []);
});

test('fail-safe: a report upload cut off at any point leaves Import old, empty or complete, never partial', async () => {
  const report = (n, tag) => [['Per Nbr', 'Pat Name'], ...Array.from({ length: n }, (_, i) => [`00${i}`, `${tag} ${i}`])];
  const OLD = report(5, 'old'), NEW = report(2500, 'new');
  const m0 = createMock(); await m0.SoharExcel.writeImportRows(OLD); m0.inject(); await m0.SoharExcel.writeImportRows(NEW);
  const total = m0.opCount();
  const bad = [];
  for (const mode of ['op', 'kill']) {
    for (let k = 1; k <= (mode === 'op' ? total : m0.syncCount()); k++) {
      const m = createMock(), X = m.SoharExcel;
      await X.writeImportRows(OLD);
      m.inject({ [KNOBS[mode]]: k });
      try { await X.writeImportRows(NEW); } catch (e) { /* expected */ }
      m.inject();
      const got = await X.readImportRows();
      const ok = isDeepStrictEqual(got, OLD) || isDeepStrictEqual(got, NEW) || got.length === 0;
      if (!ok) bad.push(`${KNOBS[mode]}=${k}: Import has ${got.length} rows`);
      await X.recoverAll();
      if (m.sheetNames().includes('~saving Import')) bad.push(`${KNOBS[mode]}=${k}: staging sheet left after recoverAll`);
    }
  }
  assert.deepEqual(bad, []);
  const m = createMock(); await m.SoharExcel.writeImportRows(NEW);
  assert.deepEqual(await m.SoharExcel.readImportRows(), NEW, 'leading zeros kept through the staged copy');
});

test('a template table that fails mid-write says to refresh', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Patients', 'tblPatients', ['A', 'B'], [{ A: '1', B: '2' }]);
  m.inject({ failAtOp: 3 });
  await assert.rejects(X.replaceTable('tblPatients', ['A', 'B'], [{ A: 'x', B: 'y' }]), /may be incomplete; click refresh/);
});

test('two Excel windows saving the same table at once: always one complete version, never a false success', async () => {
  const bad = [];
  for (let delay = 0; delay < 80; delay++) {
    const m = createMock(), X = m.SoharExcel;
    await X.writeUploads(C.UPLOAD_COLS, uploads(3, 'old'));
    m.inject();
    const a = X.writeUploads(C.UPLOAD_COLS, uploads(4, 'a')).then(() => true, () => false);
    let t = 0;
    const b = (async () => { while (t++ < delay) await Promise.resolve(); return X.writeUploads(C.UPLOAD_COLS, uploads(6, 'b')); })().then(() => true, () => false);
    const [okA, okB] = await Promise.all([a, b]);
    const got = await X.readUploads();
    const is = (n, tag) => isDeepStrictEqual(got, expected(C.UPLOAD_COLS, uploads(n, tag)));
    if (!(is(3, 'old') || is(4, 'a') || is(6, 'b'))) bad.push(`delay ${delay}: read ${short(got)}`);
    if ((okA || okB) && is(3, 'old')) bad.push(`delay ${delay}: a save reported success but the old rows are still there`);
    if (okA && !okB && !is(4, 'a')) bad.push(`delay ${delay}: only A succeeded but A's rows are not there`);
    if (okB && !okA && !is(6, 'b')) bad.push(`delay ${delay}: only B succeeded but B's rows are not there`);
    try { await X.writeUploads(C.UPLOAD_COLS, uploads(2, 'next')); } catch (e) { bad.push(`delay ${delay}: next save threw ${e.message}`); }
  }
  assert.deepEqual(bad, []);
});

test('replaceTable renames an older template header in place (v0.12: Review "Per Nbr" is "MRN")', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Review', 'tblReview', ['Severity', 'Issue', 'Per Nbr', 'Detail', 'Fix'], [{ Severity: 'Low', Issue: 'x', 'Per Nbr': '1', Detail: '', Fix: '' }]);
  await X.replaceTable('tblReview', C.REVIEW_COLS, [{ Severity: 'High', Issue: 'MRN missing', MRN: 'PER-7', Detail: 'd', Fix: 'f' }]);
  assert.deepEqual(await X.readTable('tblReview'), [{ Severity: 'High', Issue: 'MRN missing', MRN: 'PER-7', Detail: 'd', Fix: 'f' }]);
  await assert.rejects(X.replaceTable('tblReview', ['Severity', 'Issue'], []), /expects 2/, 'a different column count still stops');
});

test('v0.12 hidden tables: Guarantor Requests and Payer Tests round-trip hidden; Payer Mapping and metrics are visible', async () => {
  const m = createMock(), X = m.SoharExcel;
  await X.writeGuarantorRequests(C.GUARANTOR_REQUEST_COLS, [{ MRN: '0042', Requested: 't', Reason: C.GUARANTOR_REASONS.none, By: 'Admin' }]);
  assert.equal(m.sheetVisibility('Guarantor Requests'), 'Hidden');
  assert.deepEqual(await X.readGuarantorRequests(), [{ MRN: '0042', Requested: 't', Reason: C.GUARANTOR_REASONS.none, By: 'Admin' }]);
  await X.writePayerTests(C.PAYER_TEST_COLS, [Object.fromEntries(C.PAYER_TEST_COLS.map((c) => [c, c === 'MRN' ? '007' : 'x']))]);
  assert.equal(m.sheetVisibility('Payer Tests'), 'Hidden');
  assert.equal((await X.readPayerTests())[0].MRN, '007', 'leading zeros kept');
  await X.writePayerMapping(C.MAPPING_DISCOVERY_COLS, [{ firstName: 'A', lastName: 'B', dateOfBirth: '01/01/1990', state: 'CA', zipCode: '01234', payerId: '', id: '007-M', memberId: '0099' }]);
  assert.equal(m.sheetVisibility('Payer Mapping'), 'Visible');
  assert.equal((await X.readPayerMapping())[0].zipCode, '01234');
  await X.writePayerMetrics(C.PAYER_METRIC_COLS, []);
  await X.writePayerMetricHistory(C.PAYER_METRIC_HISTORY_COLS, [{ Date: '2026-10-20', 'NextGen Payer': 'P', 'Unmapped appts': '3' }]);
  assert.equal((await X.readPayerMetricHistory())[0]['Unmapped appts'], '3');
  assert.equal(m.activeSheet(), 'Import', 'none of them activates');
});

test('Payer Credential Rules: missing sheet reads [] (rule silent); import writes a visible sheet that reads back as the same rules', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readPayerCredRules(), []);
  assert.deepEqual(RL.payerRulesFrom(await X.readPayerCredRules()), []);
  const rules = PR.read([PR.HEADERS,
    ['Zenith CA', 'CA', 'NP', 'No', 'No', 'No', 'No', 'No', 'zenith', ''],
    ['Zenith CA', 'CA', 'MD', 'No', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith', ''],
    ['Orbit Medicare', '', 'NP', '', 'Yes', 'Yes', 'Yes', 'Yes', 'orbit medicare', 'Besse pending'],
  ]).rules;
  await X.writePayerCredRules(PR.HEADERS, RL.payerRuleRows(rules));
  assert.equal(m.sheetVisibility('Payer Credential Rules'), 'Visible');
  assert.equal(m.activeSheet(), 'Import', 'the import does not activate');
  const sig = (rs) => rs.map((r) => [r.payer, r.state, r.cred, r.besse, r.first, r.start, r.obs, r.match.join(), r.notes]);
  assert.deepEqual(sig(RL.payerRulesFrom(await X.readPayerCredRules())), sig(rules));
  await X.writePayerCredRules(PR.HEADERS, RL.payerRuleRows(rules.slice(0, 1)));
  assert.equal((await X.readPayerCredRules()).length, 1, 'a re-import replaces every row');
});

test('v0.14 Clinic Notes and Clinic Payer Notes: missing reads []; mergeClinicWork rows save visible, read back as text and merge again', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readClinicNotes(), []);
  assert.deepEqual(await X.readClinicPayerNotes(), []);
  const W = (Kind, Key, Value, Note, When) => ({ Kind, Key, MRN: Kind === 'Note' ? Key : '', LASFIR: '', Value, Note, By: 'Front desk', When });
  // A fixed now: mergeClinicWork skips rows older than APPLIED_DAYS, so new Date() would fail this test from February 2027.
  const ctx = (notes, payerNotes) => ({ location: 'Synthetic Bay', requests: [], tasks: [], sentDone: [], notes, payerNotes, now: new Date('2026-10-23T12:00:00Z') });
  const first = RL.mergeClinicWork([W('Note', '0090001', '', 'Card on Monday', '2026-10-22T17:00:00.000Z'),
    W('Payer note', 'Zenith CA', 'Zenith CA PPO 0012', 'new group', '2026-10-22T17:01:00.000Z')], ctx([], []));
  await X.writeClinicNotes(RL.CLINIC_NOTE_COLS, first.notes);
  await X.writeClinicPayerNotes(RL.CLINIC_PAYER_NOTE_COLS, first.payerNotes);
  assert.equal(m.sheetVisibility('Clinic Notes'), 'Visible');
  assert.equal(m.sheetVisibility('Clinic Payer Notes'), 'Visible');
  assert.equal(m.activeSheet(), 'Import', 'neither activates');
  assert.deepEqual(await X.readClinicNotes(), first.notes);
  assert.equal((await X.readClinicNotes())[0].MRN, '0090001', 'leading zeros kept');
  assert.deepEqual(await X.readClinicPayerNotes(), first.payerNotes);
  // A later file: an older note loses, a newer blank one clears the payer note.
  const second = RL.mergeClinicWork([W('Note', '0090001', '', 'stale', '2026-10-21T09:00:00.000Z'), W('Payer note', 'zenith ca', '', '', '2026-10-23T09:00:00.000Z')],
    ctx(await X.readClinicNotes(), await X.readClinicPayerNotes()));
  await X.writeClinicNotes(RL.CLINIC_NOTE_COLS, second.notes);
  await X.writeClinicPayerNotes(RL.CLINIC_PAYER_NOTE_COLS, second.payerNotes);
  assert.deepEqual((await X.readClinicNotes()).map((r) => r.Note), ['Card on Monday']);
  assert.deepEqual(await X.readClinicPayerNotes(), []);
});

// ---------- v0.14 Clinic Work Inbox: flow-owned, created once, only read ----------
// The Power Automate flow appends rows (modelled here with rows.add, as the Excel connector's "Add a row into a table" does).
const inboxRows = (n, tag) => Array.from({ length: n }, (_, i) => ({ Location: `${tag} Clinic`, Kind: i % 2 ? 'Note' : 'Task done', Key: i % 2 ? `00900${i}` : `00900${i}|Fix member ID in NextGen`,
  MRN: `00900${i}`, LASFIR: 'SYNPAT', Value: i % 2 ? '' : 'Yes', Note: `${tag} note ${i}`, By: 'Front desk', When: `2026-10-22T1${i % 10}:00:00.000Z`,
  File: `Clinic Eligibility - ${tag} Clinic 2026-10-22.xlsx`, Received: '2026-10-22T20:00:00.000Z' }));
async function flowAppends(m, rows) {
  await m.Excel.run(async (ctx) => { ctx.workbook.tables.getItem('tblClinicWorkInbox').rows.add(null, rows.map((r) => RL.INBOX_COLS.map((c) => r[c]))); await ctx.sync(); });
}
const inboxSnapshot = (m) => ({ info: m.tableInfo('tblClinicWorkInbox'), rows: m.tableRows('tblClinicWorkInbox', true), vis: m.sheetVisibility('Clinic Work Inbox'),
  fmts: m.numberFormats('Clinic Work Inbox', m.tableInfo('tblClinicWorkInbox').address) });

test('Clinic Work Inbox: missing reads null; set up once makes a visible header-only text table; the flow\'s rows read back as text', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.equal(await X.readClinicInbox(), null, 'missing: the pane offers to set it up');
  assert.equal(await X.createClinicInbox(RL.INBOX_COLS), true);
  assert.deepEqual(m.tableInfo('tblClinicWorkInbox'), { sheet: 'Clinic Work Inbox', address: 'A1:K2', rows: 1, cols: 11 });
  assert.equal(m.sheetVisibility('Clinic Work Inbox'), 'Visible');
  assert.deepEqual(m.values('Clinic Work Inbox', 'A1:K1'), [RL.INBOX_COLS]);
  assert.deepEqual(m.numberFormats('Clinic Work Inbox', 'A1:K2'), [RL.INBOX_COLS.map(() => '@'), RL.INBOX_COLS.map(() => '@')]);
  assert.deepEqual(m.sheetNames(), ['Import', 'Clinic Work Inbox'], 'no staging sheet left');
  assert.equal(m.activeSheet(), 'Import', 'set up does not move you');
  assert.deepEqual(await X.readClinicInbox(), [], 'header only reads as no rows');
  await flowAppends(m, inboxRows(3, 'Synthetic'));
  const got = await X.readClinicInbox();
  assert.deepEqual(got, expected(RL.INBOX_COLS, inboxRows(3, 'Synthetic')));
  assert.equal(got[0].MRN, '009000', 'appended rows take the text format: leading zeros kept');
  assert.equal(got[0].When, '2026-10-22T10:00:00.000Z', 'When stays ISO text, not a date serial');
  assert.equal(RL.inboxWork(got, new Date('2026-10-23T12:00:00Z')).get('Synthetic Clinic').length, 3, 'rules.inboxWork reads the rows as they come back');
  const before = inboxSnapshot(m);
  assert.equal(await X.createClinicInbox(RL.INBOX_COLS), false, 'already there');
  assert.deepEqual(inboxSnapshot(m), before, 'a second set up leaves the table exactly as it is');
});

test('Clinic Work Inbox: set up refuses a sheet of yours with the same name and changes nothing', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.addSheet('Clinic Work Inbox', [['my notes']]);
  await assert.rejects(X.createClinicInbox(RL.INBOX_COLS), /already has a sheet named "Clinic Work Inbox".*Nothing was changed/);
  assert.deepEqual(m.sheetNames(), ['Import', 'Clinic Work Inbox']);
  assert.deepEqual(m.values('Clinic Work Inbox', 'A1:A1'), [['my notes']]);
  assert.deepEqual(m.tableNames(), []);
});

test('fail-safe: Clinic Work Inbox set up cut off at any op, sync or Excel close: no half-made table, nothing else changed, the next try finishes', async () => {
  const seed = async (m) => { await m.SoharExcel.writeTasks(C.TASK_COLS, tasks(3, 'old')); };
  const m0 = createMock(); await seed(m0);
  const tasksBefore = await m0.SoharExcel.readTasks(), visibleBefore = m0.visibleSheets();
  const bad = [];
  for (const mode of ['op', 'sync', 'kill']) {
    const probe = createMock(); await seed(probe); probe.inject();
    await probe.SoharExcel.createClinicInbox(RL.INBOX_COLS);
    const total = mode === 'op' ? probe.opCount() : probe.syncCount();
    assert.ok(total > 2, `${mode}: ${total} points`);
    for (let k = 1; k <= total; k++) {
      const m = createMock(), X = m.SoharExcel; await seed(m);
      m.inject({ [KNOBS[mode]]: k });
      let resolved = true; try { await X.createClinicInbox(RL.INBOX_COLS); } catch (e) { resolved = false; }
      m.inject();
      const tag = `${KNOBS[mode]}=${k}`;
      if (m.tableNames().includes('tblClinicWorkInbox')) {
        // The real name only ever sits on a complete header-only table.
        const t = m.tableInfo('tblClinicWorkInbox');
        if (t.cols !== 11 || !isDeepStrictEqual(m.values(t.sheet, 'A1:K1'), [RL.INBOX_COLS])) bad.push(`${tag}: half-made inbox ${inspect(t)}`);
      } else if (resolved) bad.push(`${tag}: resolved without a table`);
      if (!isDeepStrictEqual(await X.readTasks(), tasksBefore)) bad.push(`${tag}: NextGen Tasks changed`);
      const extra = m.visibleSheets().filter((s) => !visibleBefore.includes(s) && s !== 'Clinic Work Inbox' && !s.startsWith('~saving'));
      if (extra.length) bad.push(`${tag}: new visible sheets ${extra}`);
      try {
        await X.createClinicInbox(RL.INBOX_COLS);
        if (!isDeepStrictEqual(m.tableInfo('tblClinicWorkInbox'), { sheet: 'Clinic Work Inbox', address: 'A1:K2', rows: 1, cols: 11 })) bad.push(`${tag}: next try left ${inspect(m.tableInfo('tblClinicWorkInbox'))}`);
        if (m.sheetVisibility('Clinic Work Inbox') !== 'Visible') bad.push(`${tag}: inbox hidden after the next try`);
        if (m.sheetNames().some((s) => s.startsWith('~saving'))) bad.push(`${tag}: staging sheet left after the next try: ${m.sheetNames()}`);
        if (m.tableNames().some((t) => /_stg$/.test(t))) bad.push(`${tag}: staging table left: ${m.tableNames()}`);
        await flowAppends(m, inboxRows(2, 'After'));
        if ((await X.readClinicInbox()).length !== 2) bad.push(`${tag}: the flow's rows did not read back`);
      } catch (e) { bad.push(`${tag}: the next try threw ${e.message}`); }
    }
  }
  assert.deepEqual(bad, []);
});

test('Clinic Work Inbox: no write path touches it (every mutating export run against a workbook holding it), and it is not OWNED', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Config', 'tblConfig', ['Setting', 'Value', 'What it does'], [{ Setting: 'Scope', Value: 'Upcoming', 'What it does': '' }]);
  m.seedTable('Location', 'tblLocationState', ['Location', 'State'], [{ Location: 'Synthetic Clinic', State: 'CA' }]);
  m.seedTable('Patients', 'tblPatients', ['Per Nbr', 'Zip'], [{ 'Per Nbr': '1', Zip: '02134' }]);
  m.seedTable('Payer Map', 'tblPayerMap', ['NextGen Payer Name', 'Sohar payerId', 'Sohar payerName', 'Confidence', 'Action', 'Notes', 'Last reviewed'], [], { text: false });
  m.registerXlsx('b64-report', [{ name: 'Report', rows: [['Per Nbr', 'Zip'], ['001', '02134']] }]);
  await X.createClinicInbox(RL.INBOX_COLS);
  await flowAppends(m, inboxRows(4, 'Synthetic'));
  const before = inboxSnapshot(m);
  const two = (cols) => [Object.fromEntries(cols.map((c) => [c, `${c} 1`])), Object.fromEntries(cols.map((c) => [c, `${c} 2`]))];
  const calls = {
    upsertPayerMap: () => X.upsertPayerMap({ nextgenPayerName: 'Plan Z', payerId: '01111', payerName: 'Z', confidence: 'Exact' }),
    upsertPayerMapMany: () => X.upsertPayerMapMany([{ nextgenPayerName: 'Plan Y', payerId: '02222', payerName: 'Y', confidence: 'Exact' }]),
    replaceTable: () => X.replaceTable('tblPatients', ['Per Nbr', 'Zip'], [{ 'Per Nbr': '0010', Zip: '01234' }]),
    ensureConfigRows: () => X.ensureConfigRows(C.REVERIFY_CONFIG_ROWS),
    setConfigValue: () => X.setConfigValue('Last updated', 'now'),
    appendPayerMapLog: () => X.appendPayerMapLog(C.PAYER_MAP_LOG_COLS, mapLog(2, 'a')),
    writeFee: () => X.writeFee('locations', ['Location'], [{ Location: 'Synthetic Clinic' }]),
    writeTasks: () => X.writeTasks(C.TASK_COLS, tasks(3, 't')),
    setTaskDone: () => X.setTaskDone('t Task Key 0', true),
    writeImportRows: () => X.writeImportRows([['Per Nbr'], ['001']]),
    importXlsx: () => X.importXlsx('b64-report'),
    readXlsxRows: () => X.readXlsxRows('b64-report'),
    recoverAll: () => X.recoverAll(),
    activate: () => X.activate('Clinic Work Inbox'),
    createClinicInbox: () => X.createClinicInbox(RL.INBOX_COLS),
    lockNameSheets: () => X.lockNameSheets(),
  };
  // Every other write* export: a plain (cols, rows) table save.
  for (const k of Object.keys(X)) if (/^write/.test(k) && !calls[k]) calls[k] = () => X[k](['Key', 'Record'], two(['Key', 'Record']));
  // A new mutating export must be listed here (or be a write*) so this proof keeps covering every write path.
  const mutating = Object.keys(X).filter((k) => typeof X[k] === 'function' && /^(write|upsert|replace|ensure|set|append|import|recover|activate|create|insert|delete|clear|add|lock)/.test(k));
  assert.deepEqual(mutating.filter((k) => !calls[k]), [], 'every mutating export is exercised');
  assert.ok(mutating.length >= 35, `${mutating.length} write paths`);
  const bad = [];
  for (const [k, f] of Object.entries(calls)) {
    try { await f(); } catch (e) { bad.push(`${k} threw ${e.message}`); continue; }
    if (!isDeepStrictEqual(inboxSnapshot(m), before)) { bad.push(`${k} changed the inbox`); break; }
  }
  assert.deepEqual(bad, []);
  assert.deepEqual(await X.readClinicInbox(), expected(RL.INBOX_COLS, inboxRows(4, 'Synthetic')));
  // And by construction: the inbox is outside the stage-and-swap set, and only its two functions name it.
  const src = readFileSync(new URL('../excel.js', import.meta.url), 'utf8');
  const owned = src.slice(src.indexOf('const OWNED = {'), src.indexOf('};', src.indexOf('const OWNED = {')));
  assert.doesNotMatch(owned, /Inbox/);
  assert.deepEqual(src.match(/tblClinicWorkInbox/g), ['tblClinicWorkInbox'], 'the table name appears once, in INBOX');
  const uses = [...src.matchAll(/INBOX\.(table|sheet)/g)].map((x) => src.lastIndexOf('async function', x.index)).map((i) => src.slice(i, src.indexOf('(', i)));
  assert.deepEqual([...new Set(uses)], ['async function readClinicInbox', 'async function createClinicInbox']);
});
