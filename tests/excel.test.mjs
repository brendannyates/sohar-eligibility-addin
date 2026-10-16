import test from 'node:test';
import assert from 'node:assert/strict';
import { isDeepStrictEqual, inspect } from 'node:util';
import { createRequire } from 'node:module';
import { createMock } from './helpers/excel-mock.mjs';
const C = createRequire(import.meta.url)('../core.js');

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

// ---------- happy path ----------

test('writeTasks creates the sheet and table; a rewrite replaces every row', async () => {
  const m = createMock(), X = m.SoharExcel;
  assert.deepEqual(await X.readTasks(), []);
  await X.writeTasks(C.TASK_COLS, tasks(7, 'old'));
  assert.deepEqual(m.sheetNames(), ['Import', 'NextGen Tasks']);
  assert.deepEqual(m.tableInfo('tblNextGenTasks'), { sheet: 'NextGen Tasks', address: 'A1:P8', rows: 7, cols: 16 });
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
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'Hidden');
  assert.deepEqual(await X.readUploads(), expected(C.UPLOAD_COLS, uploads(5, 'a')));
  await X.writeUploads(C.UPLOAD_COLS, uploads(2, 'b'));
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'Hidden');
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
  assert.equal(m.sheetVisibility('IR Uploads'), 'Hidden');
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
      if (hidden && m.sheetVisibility(sheet) !== 'Hidden') bad.push(`${KNOBS[mode]}=${k}: ${sheet} visible after the next clean write`);
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
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'Hidden');
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
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'Hidden');
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
  assert.equal(m.sheetVisibility('Sohar Uploads'), 'Hidden');
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
