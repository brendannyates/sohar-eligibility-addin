import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createMock } from './helpers/excel-mock.mjs';
const require = createRequire(import.meta.url);
const K = require('../payer-knowledge.js');

// v0.17 built-in payer knowledge. The bar and merge tests use made-up payers; the last tests check the generated
// block itself (payer names and Sohar payer IDs only, no patient data).
const pm = (name, id, conf, notes = '') => ({ 'NextGen Payer Name': name, 'Sohar payerId': id, 'Sohar payerName': `${name} (Sohar)`, Confidence: conf, Notes: notes });
const lrn = (name, agreed, errors) => ({ 'NextGen Payer': name, 'Patients agreed': agreed, 'Payer errors': errors });

test('the bar: Exact / Confirmed / Learned / Not a payer as they are; Auto-High and Probable only with Sohar agreement', () => {
  const sel = K.select({
    payerMap: [pm('Plan Exact', '01111', 'Exact', 'note with a person in it'), pm('Plan Confirmed', '02222', 'Confirmed'), pm('Plan Learned', '03333', 'Learned'),
      pm('Plan None', '', 'Not a payer'), pm('Plan Exact No Id', '', 'Exact'), pm('Plan Review', '04444', 'Review'), pm('Plan Low', '05555', 'Auto-Low'),
      pm('Plan Medium', '06666', 'Auto-Medium'), pm('Plan High Ok', '07777', 'Auto-High'), pm('Plan High Few', '08888', 'Auto-High'),
      pm('Plan Probable Err', '09999', 'Probable'), pm('Plan Probable Ok', '01212', 'Probable'), pm('Plan Probable None', '01313', 'Probable'), pm('plan exact', '09090', 'Exact')],
    payerLearning: [lrn('Plan High Ok', 3, 0), lrn('Plan High Few', 2, 0), lrn('Plan Probable Err', 9, 1), lrn('Plan Probable Ok', '4', '0')],
  });
  assert.deepEqual(sel.payerMap.map((r) => r[0]), ['Plan Exact', 'Plan Confirmed', 'Plan Learned', 'Plan None', 'Plan High Ok', 'Plan Probable Ok']);
  assert.deepEqual(sel.payerMap[0], ['Plan Exact', '01111', 'Plan Exact (Sohar)', 'Exact'], 'four fields only: notes are never copied');
});

test('Sohar payers dedupe by payerId; fee locations need a tab; High fee matches split into built in vs review', () => {
  const sel = K.select({
    soharPayers: [{ payerId: '01111', payerName: 'A' }, { payerId: '01111', payerName: 'A again' }, { payerId: '', payerName: 'none' }, { payerId: '02222', payerName: 'B', payerGroupId: 'G1' }],
    feeLocations: [{ 'NextGen Location': 'Clinic One', 'Fee Tab': 'Tab1' }, { 'NextGen Location': 'Clinic Tele', 'Fee Tab': '' }],
    feePayerMap: [{ Payer: 'Plan A', 'Fee Tab': 'Tab1', 'Fee Payer': 'Plan-A', Confidence: 'High' }, { Payer: 'Plan B HMO', 'Fee Tab': 'Tab1', 'Fee Payer': 'Plan B', Confidence: 'High', Note: 'token match' },
      { Payer: 'Plan C', 'Fee Tab': 'Tab1', 'Fee Payer': 'Plan C', Confidence: 'Low' }, { Payer: 'Plan A', 'Fee Tab': 'Tab1', 'Fee Payer': 'Other', Confidence: 'High' }],
  });
  assert.deepEqual(sel.soharPayers.map((r) => r[0]), ['01111', '02222']);
  assert.equal(sel.soharPayers[1][4], 'G1');
  assert.deepEqual(sel.feeLocations, [['Clinic One', 'Tab1']]);
  assert.deepEqual(sel.feePayerMap, [['Plan A', 'Tab1', 'Plan-A']], 'name = column (punctuation ignored), first row per payer and tab');
  assert.deepEqual(sel.review.map((r) => r.Payer), ['Plan B HMO']);
});

test('merge: sheet rows win and come first; built-in rows fill in, marked builtIn; an empty sheet gets every built-in row', () => {
  const built = K.payerMap();
  assert.ok(built.length > 0 && built.every((r) => r.builtIn === true));
  const first = built[0];
  const sheet = [{ nextgenPayerName: first.nextgenPayerName.toUpperCase(), payerId: '00000', payerName: 'Override', confidence: 'Review', notes: '' },
    { nextgenPayerName: 'Synthetic Local Plan', payerId: '01111', payerName: 'Local', confidence: 'Exact', notes: '' }];
  const merged = K.mergePayerMap(sheet);
  assert.deepEqual(merged.slice(0, 2), sheet);
  assert.equal(merged.filter((r) => r.nextgenPayerName.toLowerCase() === first.nextgenPayerName.toLowerCase()).length, 1, 'override replaces, never duplicates');
  assert.equal(merged.length, built.length + 1);
  assert.equal(K.mergePayerMap([]).length, built.length);
  assert.equal(K.mergeSoharPayers([{ payerId: K.soharPayers()[0].payerId, payerName: 'Sheet name', aliasPayerNames: '' }]).filter((p) => p.payerName === 'Sheet name').length, 1);
  const fpm = K.mergeFeePayerMap([{ Payer: '', 'Sohar payerId': '01111', 'Fee Tab': 'Tab1', 'Fee Payer': 'X', Source: 'manual' }]);
  assert.equal(fpm[0]['Sohar payerId'], '01111', 'payerId-only manual rows are kept');
});

test('excel: built-ins fill a workbook with no Payer Map or Sohar Payers; the first save makes the Payer Map; templateCheck allows both missing', async () => {
  const m = createMock({ knowledge: true }), X = m.SoharExcel;
  const map = await X.readPayerMap();
  assert.equal(map.length, K.payerMap().length);
  assert.ok(map.every((r) => r.builtIn));
  assert.equal((await X.readSoharPayers()).length, K.soharPayers().length);
  assert.ok(!(await X.templateCheck()).includes('tblPayerMap'));
  assert.ok(!(await X.templateCheck()).includes('tblSoharPayers'));
  await X.upsertPayerMapMany([{ nextgenPayerName: 'Synthetic New Plan', payerId: '01111', payerName: 'New', confidence: 'Auto-Low' }]);
  assert.equal(await X.tableExists('tblPayerMap'), true);
  const after = await X.readPayerMap();
  assert.equal(after[0].nextgenPayerName, 'Synthetic New Plan');
  assert.ok(!after[0].builtIn);
  assert.equal(after.length, K.payerMap().length + 1);
  // Changing a built-in payer saves it on the sheet as an override.
  const b = K.payerMap()[0];
  await X.upsertPayerMapMany([{ nextgenPayerName: b.nextgenPayerName, payerId: b.payerId, payerName: b.payerName, confidence: 'Review' }]);
  const over = (await X.readPayerMap()).filter((r) => r.nextgenPayerName === b.nextgenPayerName);
  assert.equal(over.length, 1);
  assert.equal(over[0].confidence, 'Review');
  assert.ok(!over[0].builtIn);
});

test('excel: without payer knowledge loaded, the sheets read exactly as before', async () => {
  const m = createMock(), X = m.SoharExcel;
  m.seedTable('Payer Map', 'tblPayerMap', ['NextGen Payer Name', 'Sohar payerId', 'Sohar payerName', 'Confidence', 'Action', 'Notes', 'Last reviewed'],
    [{ 'NextGen Payer Name': 'Plan A', 'Sohar payerId': '01111', 'Sohar payerName': 'A', Confidence: 'Exact' }]);
  assert.deepEqual((await X.readPayerMap()).map((r) => r.nextgenPayerName), ['Plan A']);
});

test('generated block: counts match, no duplicate keys, every mapped payerId is on the Sohar list, nothing person-shaped', () => {
  const map = K.payerMap(), sohar = K.soharPayers();
  assert.deepEqual(K.META.counts, { payerMap: map.length, soharPayers: sohar.length, feeLocations: K.feeLocations().length, feePayerMap: K.feePayerMap().length });
  assert.equal(new Set(map.map((r) => r.nextgenPayerName.toLowerCase())).size, map.length);
  assert.equal(new Set(sohar.map((r) => r.payerId)).size, sohar.length);
  const ids = new Set(sohar.map((r) => r.payerId));
  assert.deepEqual(map.filter((r) => r.payerId && !ids.has(r.payerId)).map((r) => r.nextgenPayerName), []);
  for (const r of map) assert.ok(K.BAR.keep.includes(r.confidence.toLowerCase()) || K.BAR.needEvidence.includes(r.confidence.toLowerCase()), r.confidence);
  // No dates of birth, MRN-style keys or member IDs: every value is a payer name, an ID, a label or a fee tab.
  const text = JSON.stringify([map, K.feeLocations(), K.feePayerMap()].map((l) => l.map((r) => { const { notes, Note, ...rest } = r; return rest; })));
  assert.doesNotMatch(text, /\b\d{1,2}\/\d{1,2}\/(19|20)\d{2}\b/);
  assert.doesNotMatch(text, /PER-\d|\bMRN\b|memberId/i);
});
