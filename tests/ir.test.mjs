import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const C = createRequire(import.meta.url)('../core.js');

// Synthetic data only. Header layout follows the IR Portal Filtered tab (data starting in column C).
const IR = [
  ['', '', 'Incident ID', 'Last Name', 'First Name', 'DOB', 'Description', 'Facility', 'MRN', 'DOS', 'Provider', 'Due Date'],
  ['', '', 'I-1', 'Okafor', 'Chidi', '1/1/1990', 'ER/HIM/NB/Missing Insurance', 'San Rafael', '0901', '9/30/2026', 'DOE JANE', '10/20/2026'],
  ['', '', 'I-2', 'Rivera', 'Ana', '3/4/1980', 'ER/HIM/NB/Missing Demo/Verify Insurance - Checked NG, insurance on file inactive', 'San Rafael', '900', '9/15/2026', 'DOE JANE', '10/18/2026'],
  ['', '', 'I-3', 'Rivera', 'Ana', '3/4/1980', 'ER/HIM/NB/Missing Provider Signature', 'San Rafael', '900', '9/15/2026', 'DOE JANE', '10/18/2026'],
  ['', '', 'I-4', 'Haddad', 'Mira', '5/5/1985', 'ER/HIM/NB/Authorization Missing - Auth on file is expired', 'Walnut Creek', '', '9/1/2026', 'ROE JO', '10/25/2026'],
  ['', '', 'I-5', 'Petrov', 'Ilya', '2/2/1970', 'ER/HIM/NB/Provider Out of Network', 'Walnut Creek', '777', '9/2/2026', 'ROE JO', ''],
];
const base = (o) => Object.fromEntries(C.RESULTS_COLS.map((c) => [c, o[c] ?? '']));
const RESULTS = [
  base({ 'Per Nbr': '101', MRN: '901', 'Last Name': 'Okafor', 'First Name': 'Chidi', DOB: '01/01/1990', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'Sohar Payer': 'Cigna', 'Choose in NextGen': 'Cigna', 'Next Appt': '10/21/2026', 'Next Location': 'San Rafael', 'Sohar Date': '2026-10-14', Match: 'Per Nbr' }),
  base({ 'Per Nbr': '100', MRN: '900', 'Last Name': 'Rivera', 'First Name': 'Ana', DOB: '03/04/1980', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Payer': 'Blue Shield of California', Plan: 'SYNTH PPO', 'Coverage Start': '10/1/2026', 'Benefit Flag': 'Deductible met', 'Sohar Date': '2026-10-14', Match: 'Per Nbr' }),
  base({ 'Per Nbr': '103', MRN: '', 'Last Name': 'Haddad', 'First Name': 'Mira', DOB: '05/05/1985', 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar Payer': 'Aetna', Plan: 'SYNTH HMO', 'Sohar Date': '2026-10-14', Match: 'Per Nbr' }),
];
const TASKS = [{ 'Per Nbr': '101', Task: 'Add insurance in NextGen', Done: '' }, { 'Per Nbr': '100', Task: 'Update benefits in NextGen', Done: 'Yes' }];

test('irCategory keeps insurance-related descriptions only', () => {
  assert.equal(C.irCategory('ER/HIM/NB/Missing Insurance-Please provide Medicare policy ID'), 'Missing insurance');
  assert.equal(C.irCategory('ER/HIM/NB/Missing Demo/Verify Insurance– "Cigna" Denied claim for member not found'), 'Verify insurance');
  assert.equal(C.irCategory('ER/HIM/NB/Provider Out of Network'), 'Out of network');
  assert.equal(C.irCategory('ER/HIM/NB/Authorization Missing- Auth on file is not valid'), 'Authorization');
  assert.equal(C.irCategory('ER/HIM/NB/Missing Provider Signature'), '');
});

test('parseIrPortal finds headers that start in column C', () => {
  const { records } = C.parseIrPortal(IR);
  assert.equal(records.length, 5);
  assert.equal(records[0].mrn, '0901');
  assert.match(C.parseIrPortal([['a', 'b'], [1, 2]]).error, /MRN and Description/);
});

test('irCrossReference joins IR items to Sohar results and suggests the fix', () => {
  const rows = C.irCrossReference(C.parseIrPortal(IR).records, RESULTS, TASKS);
  for (const r of rows) assert.deepEqual(Object.keys(r), C.IR_COLS);
  assert.equal(rows.length, 3, 'the signature item and the item with no Sohar result are left out');
  const by = (id) => rows.find((r) => r['Incident ID'] === id);

  const a = by('I-1');
  assert.equal(a.Match, 'MRN', 'leading zeros ignored');
  assert.match(a['Suggested Action'], /^Sohar found Cigna\. Add it in NextGen \(Cigna\)/);
  assert.equal(a['Open NextGen Tasks'], 'Add insurance in NextGen');

  const b = by('I-2');
  assert.match(b['Suggested Action'], /active now\. Check it covered the DOS.*Coverage started 10\/01\/2026, after the 09\/15\/2026 DOS/);
  assert.equal(b['Open NextGen Tasks'], '', 'done tasks are not listed');

  const c = by('I-4');
  assert.equal(c.Match, 'Name + DOB', 'no MRN on the IR row');
  assert.match(c['Suggested Action'], /does not check authorizations/);

  assert.equal(by('I-5'), undefined, 'no Sohar result for this MRN');
  const notReturned = C.irCrossReference(C.parseIrPortal(IR).records,
    [...RESULTS, { ...RESULTS[0], 'Per Nbr': '107', MRN: '777', 'Sohar Status': 'Not in Sohar file', 'Sent As': 'Verification' }], TASKS);
  assert.equal(notReturned.find((r) => r['Incident ID'] === 'I-5'), undefined, '"Not in Sohar file" counts as no result');

  assert.equal(rows[0]['Incident ID'], 'I-2', 'soonest due date first');
});

test('irMessage uses LASFIR and the last four MRN digits only', () => {
  const msg = C.irMessage(C.irCrossReference(C.parseIrPortal(IR).records, RESULTS, TASKS));
  assert.match(msg, /OKACHI · MRN \.\.\.0901 · DOS 09\/30\/2026 · Missing insurance: Sohar found Cigna/);
  assert.doesNotMatch(msg, /Okafor|Chidi/);
});
