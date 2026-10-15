import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const C = createRequire(import.meta.url)('../core.js');

// Synthetic data only.
const mk = (o) => ({ ...Object.fromEntries(C.RESULTS_COLS.map((c) => [c, ''])), 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', ...o });

test('fileName: Sohar ready import names carry the local date', () => {
  const d = new Date(2026, 9, 15, 22, 30); // 10:30 PM local stays the 15th
  assert.equal(C.fileName('verification', d), 'Sohar Ready Import - Verification 2026-10-15.csv');
  assert.equal(C.fileName('discovery', d), 'Sohar Ready Import - Discovery 2026-10-15.csv');
  assert.equal(C.fileName('discoveryVerify', d), 'Sohar Ready Import - Verification (from Discovery) 2026-10-15.csv');
});

test('benefitsView: OOP met first, then approaching, then by OOP % met; progress amounts', () => {
  const rows = [
    mk({ 'Per Nbr': '1', 'Last Name': 'Low', 'Sohar Ref': 'a', 'OOP Max': '5000', 'OOP Remaining': '4000', 'OOP Status': 'Not met', Deductible: '1000', 'Deductible Remaining': '250', 'Deductible Status': 'Not met', 'Next Event': 'TMS Treatment' }),
    mk({ 'Per Nbr': '2', 'Last Name': 'Near', 'Sohar Ref': 'b', 'OOP Max': '3000', 'OOP Remaining': '300', 'OOP Status': 'Approaching' }),
    mk({ 'Per Nbr': '3', 'Last Name': 'Done', 'Sohar Ref': 'c', 'OOP Max': '2000', 'OOP Remaining': '0', 'OOP Status': 'Met' }),
    mk({ 'Per Nbr': '4', 'Last Name': 'Mid', 'Sohar Ref': 'd', 'OOP Max': '4000', 'OOP Remaining': '1600', 'OOP Status': 'Not met' }),
    mk({ 'Per Nbr': '5', 'Sent As': 'Discovery', 'Sohar Ref': 'e', 'Discovery Follow-up': 'Ready to verify' }),
  ];
  const v = C.benefitsView(rows);
  assert.deepEqual(v.rows.map((r) => r['Per Nbr']), ['3', '2', '4', '1']);
  assert.equal(v.rows[0].Flag, 'OOP max met');
  assert.equal(v.rows[1].Flag, 'OOP max approaching');
  assert.equal(v.rows[1]['OOP Met'], '2700');
  assert.equal(v.rows[1]['OOP % Met'], 0.9);
  const low = v.rows[3];
  assert.equal(low['OOP % Met'], 0.2);
  assert.equal(low['Deductible Met'], '750');
  assert.equal(low['Deductible % Met'], 0.75);
  assert.equal(low.Service, 'TMS');
  assert.equal(v.pending.length, 1);
  for (const r of v.rows) assert.deepEqual(Object.keys(r), C.BENEFIT_SHEET_COLS);

  const s = C.benefitsSummary(v);
  assert.equal(s.verified, 4);
  assert.equal(s.oopMet, 1);
  assert.equal(s.oopApproaching, 1);
  assert.deepEqual(s.watch.map((r) => r['Per Nbr']), ['3', '2']);
  assert.equal(s.pending, 1);
});

test('benefitsView: $0 or missing totals give blank progress, not a divide by zero', () => {
  const v = C.benefitsView([mk({ 'Per Nbr': '9', 'Sohar Ref': 'z', Deductible: '0', 'Deductible Remaining': '0', 'Deductible Status': 'None' })]);
  assert.equal(v.rows[0]['Deductible % Met'], '');
  assert.equal(v.rows[0]['OOP % Met'], '');
  assert.equal(C.benefitsSummary(v).withAmounts, 1);
});
