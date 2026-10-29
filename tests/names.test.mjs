// v0.19 (Brendan, 10/28): no sheet shows a patient's or subscriber's name or DOB once the report is in. LASFIR is the
// lookup, MRN stays. Names live in memory (the Sohar files need them) and on the very hidden Import sheet.
// Synthetic data only.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createMock } from './helpers/excel-mock.mjs';
const C = createRequire(import.meta.url)('../core.js');

const NAME_COLS = ['Last Name', 'First Name', 'DOB', 'Subscriber First', 'Subscriber Last', 'Subscriber DOB',
  'firstName', 'lastName', 'dateOfBirth', 'subscriberFirstName', 'subscriberLastName', 'subscriberDateOfBirth'];
const SHEET_LISTS = ['PATIENT_COLS', 'RESULTS_COLS', 'TASK_COLS', 'BENEFITS_EXPORT_COLS', 'TRACKER_EXPORT_COLS', 'IR_COLS',
  'BENEFIT_SHEET_COLS', 'VERIFICATION_SHEET_COLS', 'DISCOVERY_SHEET_COLS', 'MAPPING_SHEET_COLS'];

test('no sheet or tracker column list has a name or DOB column; each has MRN or the Sohar id, and LASFIR', () => {
  for (const k of SHEET_LISTS) {
    const cols = C[k];
    assert.ok(Array.isArray(cols), k);
    assert.deepEqual(cols.filter((c) => NAME_COLS.includes(c)), [], `${k} has no name columns`);
    assert.ok(cols.includes('LASFIR'), `${k} has LASFIR`);
    assert.ok(cols.includes('MRN') || cols.includes('id'), `${k} keeps the MRN (or the Sohar id, which is the MRN key)`);
  }
});

test('the template sheets keep their column count, so the Dashboard formulas and replaceTable rename in place', () => {
  assert.equal(C.PATIENT_COLS.length, 29);
  assert.equal(C.VERIFICATION_SHEET_COLS.length, C.VERIFICATION_COLS.length);
  assert.equal(C.DISCOVERY_SHEET_COLS.length, C.DISCOVERY_COLS.length);
  // The Dashboard counts blanks in Patients C (no name at all), D (name incomplete) and E (no DOB).
  assert.deepEqual(C.PATIENT_COLS.slice(2, 5), ['LASFIR', 'Name on File', 'DOB on File']);
  // Sohar's own templates are untouched: the files still carry names.
  assert.deepEqual(C.VERIFICATION_COLS.slice(0, 3), ['firstName', 'lastName', 'dateOfBirth']);
});

test('tagOf and cellOf: LASFIR and "on file" marks from either row shape, a sheet value wins', () => {
  const p = { 'Last Name': 'Zzalder', 'First Name': 'Bennet', DOB: '02/03/1971', 'Subscriber Last': 'Zzalder', 'Subscriber First': 'Corin', 'Subscriber DOB': '' };
  assert.equal(C.tagOf(p), 'ZZABEN');
  assert.equal(C.cellOf(p, 'LASFIR'), 'ZZABEN');
  assert.equal(C.cellOf(p, 'Name on File'), 'Yes');
  assert.equal(C.cellOf(p, 'DOB on File'), 'Yes');
  assert.equal(C.cellOf(p, 'Subscriber LASFIR'), 'ZZACOR');
  assert.equal(C.cellOf(p, 'Subscriber DOB on File'), '');
  const sohar = { firstName: 'Bennet', lastName: "O'Zzal-der", dateOfBirth: '' };
  assert.equal(C.tagOf(sohar), 'OZZBEN', 'punctuation dropped');
  assert.equal(C.cellOf(sohar, 'DOB on File'), '');
  assert.equal(C.cellOf({ 'Last Name': 'Zzonly' }, 'Name on File'), '', 'first name missing: not on file');
  assert.equal(C.tagOf({ LASFIR: 'ZZXYYY', 'Last Name': 'Other' }), 'ZZXYYY', 'a row read back from a sheet keeps its LASFIR');
  assert.equal(C.cellOf({ LASFIR: '', 'Last Name': 'Zzalder', 'First Name': 'Bennet' }, 'LASFIR'), 'ZZABEN', 'blank LASFIR is worked out');
  assert.equal(C.cellOf({ Plan: 'PPO' }, 'Plan'), 'PPO');
  assert.equal(C.tagOf(null), '');
});

test('toCSV of sheet columns has LASFIR and no names; the Sohar file columns still carry them', () => {
  const r = { 'Per Nbr': '5001', MRN: '0900101', 'Last Name': 'Zzalder', 'First Name': 'Bennet', DOB: '02/03/1971', firstName: 'Bennet', lastName: 'Zzalder', dateOfBirth: '02/03/1971', id: '0900101', state: 'CA' };
  const sheet = C.toCSV(C.PATIENT_COLS, [r]) + C.toCSV(C.DISCOVERY_SHEET_COLS, [r]) + C.toCSV(C.BENEFITS_EXPORT_COLS, [r]);
  assert.doesNotMatch(sheet, /Zzalder|Bennet|1971/);
  assert.match(sheet, /ZZABEN/);
  assert.match(sheet, /0900101/, 'MRN kept');
  assert.match(C.toCSV(C.DISCOVERY_COLS, [r]), /^firstName,lastName,dateOfBirth.*\r\nBennet,Zzalder,02\/03\/1971,CA,/);
});

test('benefitsExport and trackerExport rows have LASFIR and MRN, no names (TMS and ESK match on MRN)', () => {
  const res = { 'Per Nbr': '5001', MRN: '0900101', 'Last Name': 'Zzalder', 'First Name': 'Bennet', DOB: '02/03/1971', 'Sent As': 'Verification',
    'Sohar Status': 'complete.eligible', 'Sohar Date': '2026-10-14T01:00:00.000Z', 'Sohar Ref': 'ref-1', 'Next Event': 'TMS Treatment', 'Deductible Remaining': '100', 'OOP Remaining': '2500' };
  const [b] = C.benefitsExport([res]);
  assert.equal(b.LASFIR, 'ZZABEN');
  assert.equal(b.MRN, '0900101');
  assert.ok(!('Last Name' in b) && !('DOB' in b));
  const out = C.trackerExport([res], '10/28/2026 9:00 AM');
  for (const row of out.rows) assert.doesNotMatch(JSON.stringify(row), /Zzalder|Bennet|1971/);
});

test('fillNames puts names back on Sohar file rows read from a sheet; typed member IDs stay; namesMissing guards the file', () => {
  const mem = [
    { firstName: 'Bennet', lastName: 'Zzalder', dateOfBirth: '02/03/1971', id: 'M1', payerId: '62308', memberId: '' },
    { firstName: 'Corin', lastName: 'Zzbrook', dateOfBirth: '04/05/1966', id: 'M2', payerId: '60054', memberId: '', subscriberFirstName: 'Dara', subscriberLastName: 'Zzbrook', subscriberDateOfBirth: '01/01/1960' },
  ];
  const sheet = [
    { LASFIR: 'ZZABEN', 'Name on File': 'Yes', 'DOB on File': 'Yes', id: 'M1', payerId: '62308', memberId: 'TYPED1' },
    { LASFIR: 'ZZBCOR', id: 'M2', payerId: '60054', memberId: '' },
  ];
  const out = C.fillNames(sheet, mem);
  assert.equal(out[0].firstName, 'Bennet');
  assert.equal(out[0].memberId, 'TYPED1', 'a member ID typed on the sheet goes out');
  assert.equal(out[1].subscriberLastName, 'Zzbrook');
  assert.doesNotThrow(() => C.namesMissing(out));
  assert.match(C.toCSV(C.VERIFICATION_COLS, out), /Bennet,Zzalder,02\/03\/1971/);
  // A sheet row the build no longer has (e.g. a stale sheet): no name, so the file is refused, not sent half-empty.
  const stale = C.fillNames([{ id: 'GONE', payerId: '1' }], mem);
  assert.throws(() => C.namesMissing(stale), /1 row has no name on file.*refresh/);
  assert.deepEqual(C.fillNames(sheet, undefined).map((r) => r.firstName), ['', '']);
});

test('lockNameSheets: Import, Sohar Uploads and IR Uploads go very hidden, Export goes, the user lands on Patients', async () => {
  const m = createMock({ sheets: ['Import', 'Patients'] }), X = m.SoharExcel;
  m.addSheet('Import', [['Per Nbr', 'Pat Name'], ['1', 'Zzalder, Bennet']]);
  m.seedTable('Sohar Uploads', 'tblSoharUploads', C.UPLOAD_COLS, [], { hidden: true });
  m.seedTable('IR Uploads', 'tblIrUploads', ['Record'], []);
  m.seedTable('Export', 'tblExport', ['firstName'], [{ firstName: 'Bennet' }]);
  await X.activate('Import');
  assert.equal(await X.lockNameSheets(), '');
  for (const s of ['Import', 'Sohar Uploads', 'IR Uploads']) assert.equal(m.sheetVisibility(s), 'VeryHidden', s);
  assert.ok(!m.sheetNames().includes('Export'), 'the Export sheet (names, one download) is removed');
  assert.equal(m.activeSheet(), 'Patients');
  assert.deepEqual(m.values('Import', 'A2:B2').flat().map(String), ['1', 'Zzalder, Bennet'], 'hides only, the report is still there to rebuild from');
  assert.deepEqual(m.visibleSheets(), ['Patients']);
  assert.equal(await X.lockNameSheets(), '', 'second run: nothing to do');
  assert.deepEqual((await X.readImportRows()).flat().map(String), ['Per Nbr', 'Pat Name', '1', 'Zzalder, Bennet'], 'the add-in still reads a very hidden Import');
});

test('lockNameSheets on a structure-protected workbook changes nothing and says why', async () => {
  const m = createMock({ sheets: ['Import', 'Patients'] }), X = m.SoharExcel;
  m.protectStructure(true);
  assert.match(await X.lockNameSheets(), /structure is protected/);
  assert.equal(m.sheetVisibility('Import'), 'Visible');
});

test('an older workbook: Patients names and DOBs are renamed in place, and no name reaches a cell', async () => {
  const OLD = ['Per Nbr', 'MRN', 'Last Name', 'First Name', 'DOB', ...C.PATIENT_COLS.slice(5, 25), 'Subscriber First', 'Subscriber Last', 'Subscriber DOB', 'Flags'];
  assert.equal(OLD.length, C.PATIENT_COLS.length);
  const m = createMock({ sheets: ['Import', 'Patients'] }), X = m.SoharExcel;
  m.seedTable('Patients', 'tblPatients', OLD, [Object.fromEntries(OLD.map((c) => [c, c === 'Last Name' ? 'Zzold' : c === 'First Name' ? 'Olda' : c === 'DOB' ? '01/01/1950' : '']))]);
  const p = { 'Per Nbr': '5001', MRN: '0900101', 'Last Name': 'Zzalder', 'First Name': 'Bennet', DOB: '02/03/1971', 'Subscriber First': 'Dara', 'Subscriber Last': 'Zzalder', 'Subscriber DOB': '', State: 'CA', x: {} };
  await X.replaceTable('tblPatients', C.PATIENT_COLS, [p]);
  const cells = JSON.stringify(m.tableRows('tblPatients', true));
  assert.doesNotMatch(cells, /Zzalder|Bennet|1971|Zzold|Olda|1950|Dara/);
  const [row] = await X.readTable('tblPatients');
  assert.equal(row.LASFIR, 'ZZABEN');
  assert.equal(row['Name on File'], 'Yes');
  assert.equal(row['DOB on File'], 'Yes');
  assert.equal(row['Subscriber LASFIR'], 'ZZADAR');
  assert.equal(row.MRN, '0900101');
});

test('Sohar Results and NextGen Tasks sheets get LASFIR and MRN, no names', async () => {
  const m = createMock(), X = m.SoharExcel;
  const r = Object.assign(Object.fromEntries(C.RESULTS_COLS.map((c) => [c, ''])), { 'Per Nbr': '5001', MRN: '0900101', 'Last Name': 'Zzalder', 'First Name': 'Bennet', DOB: '02/03/1971' });
  await X.writeResults(C.RESULTS_COLS, [r]);
  const t = Object.assign(Object.fromEntries(C.TASK_COLS.map((c) => [c, ''])), { MRN: '0900101', 'Last Name': 'Zzalder', 'First Name': 'Bennet', 'Task Key': 'k' });
  await X.writeTasks(C.TASK_COLS, [t]);
  const cells = JSON.stringify([m.tableRows('tblSoharResults', true), m.tableRows('tblNextGenTasks', true)]);
  assert.doesNotMatch(cells, /Zzalder|Bennet|1971/);
  assert.equal((await X.readTable('tblSoharResults'))[0].LASFIR, 'ZZABEN');
  assert.equal((await X.readTasks())[0].LASFIR, 'ZZABEN');
});
