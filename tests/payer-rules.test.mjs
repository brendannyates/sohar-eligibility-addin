// node --test tests/payer-rules.test.mjs   (made-up payers only; the real list lives in the workbook, never in the repo)
// Shared with the Clinic Template Tool and ESK Tracker repos (there: payer-rules.test.js, CommonJS); only the imports differ.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const PR = createRequire(import.meta.url)('../payer-rules.js');

const AOA = [
  ['Payer', 'State', 'Credential', 'Besse Approved', 'NP Fully Approved', 'Esketamine First Treatment', 'Esketamine Start', 'Esketamine Observation', 'Match Names', 'Notes'],
  ['Acme Health', '', 'NP', 'Y', 'Yes', 'Yes', 'Yes', 'Yes', 'acme; acme health', ''],
  ['Acme Health', '', 'MD', 'Y', 'N/A', 'Yes', 'Yes', 'Yes', 'acme; acme health', ''],
  ['Zenith CA', 'CA', 'NP', 'N', 'No', 'No', 'No', 'No', 'zenith', ''],
  ['Zenith CA', 'CA', 'MD', 'N', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith', ''],
  ['Zenith WA', '', 'NP', 'N', 'Yes', 'Yes', 'Yes', 'Yes', 'zenith', ''],
  ['Zenith WA', '', 'MD', 'N', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith', ''],
  ['Orbit (Fed)', '', 'NP', 'Y', 'No', 'No', 'No', 'Yes', 'orbit; orbit federal', 'NP may do observation only'],
  ['Orbit (Fed)', '', 'MD', 'Y', 'N/A', 'Yes', 'Yes', 'Yes', 'orbit; orbit federal', ''],
  ['Orbit Medicare', '', 'NP', '', 'Yes', 'Yes', 'Yes', 'Yes', 'orbit medicare', 'Besse pending'],
  ['Nowhere EAP', '', 'MD', '', 'No', 'No', 'No', 'No', '', ''],
];
const { rules } = PR.read(AOA);

test('reads the sheet, states from the name when the State column is blank', () => {
  assert.equal(rules.length, 10);
  assert.equal(rules.find(r => r.payer === 'Zenith WA').state, 'WA');
  assert.deepEqual(rules.find(r => r.payer === 'Nowhere EAP').match, ['nowhere eap']);   // no Match Names: the payer name
  assert.ok(PR.read([['Name', 'Thing']]).error);
});

test('longest matching name wins; Medicare-style variants beat the base plan', () => {
  assert.equal(PR.match(rules, 'ORBIT MEDICARE ADVANTAGE').payer, 'Orbit Medicare');
  assert.equal(PR.match(rules, 'Orbit Federal Services').payer, 'Orbit (Fed)');
  assert.equal(PR.match(rules, 'Acme-Health PPO').payer, 'Acme Health');
  assert.equal(PR.match(rules, 'Unrelated Plan'), null);
  assert.equal(PR.match(rules, ''), null);
});

test('state rules: the patient’s state decides, else the answer must agree', () => {
  assert.equal(PR.match(rules, 'Zenith Blue', 'Castro - CA').payer, 'Zenith CA');
  assert.equal(PR.match(rules, 'Zenith Blue', 'WA').payer, 'Zenith WA');
  assert.equal(PR.match(rules, 'Zenith Blue of Washington').payer, 'Zenith WA');
  const m = PR.match(rules, 'Zenith Blue');
  assert.deepEqual(m.ambiguous, ['Zenith CA', 'Zenith WA']);
  assert.equal(PR.check(m, 'NP', 'obs').ok, null);           // CA says no, WA says yes
  assert.equal(PR.check(m, 'MD', 'obs').ok, true);           // both say yes
  assert.equal(PR.besse(m), 'not approved');                 // both say N
});

test('credential × ESK step', () => {
  const orbit = PR.match(rules, 'Orbit Federal');
  assert.equal(PR.check(orbit, 'Lark PMHNP, Jordan', 'obs').ok, true);
  const r = PR.check(orbit, 'NP', 'first'); assert.equal(r.ok, false); assert.match(r.why, /NP may do observation only/);
  assert.equal(PR.check(orbit, 'Reyes MD, Sam', 'start').ok, true);
  assert.equal(PR.check(PR.match(rules, 'Orbit Medicare'), 'MD', 'obs').ok, null);   // no MD row
  assert.equal(PR.check(PR.match(rules, 'Nowhere EAP'), 'MD', 'obs').ok, false);
  assert.equal(PR.check(null, 'NP', 'obs').ok, null);
  assert.equal(PR.credOf('Do, Minh NP'), 'NP'); assert.equal(PR.credOf('Nguyen, Do PMHNP'), 'NP'); assert.equal(PR.credOf('Do MD, Minh'), 'MD');
  assert.equal(PR.step('Esketamine Start'), 'start');
  assert.equal(PR.step('Esketamine Follow Up', true), 'first');
  assert.equal(PR.step('Esketamine Follow Up', false), 'obs');
});

test('Besse approval, with blank = pending', () => {
  assert.equal(PR.besse(PR.match(rules, 'Acme')), 'approved');
  assert.equal(PR.besse(PR.match(rules, 'Orbit Medicare')), 'pending');
  assert.equal(PR.besse(PR.match(rules, 'Zenith', 'CA')), 'not approved');
  assert.equal(PR.besse(null), 'unknown');
});

test('rows written back read the same', () => {
  const again = PR.read(PR.toRows(rules)).rules;
  assert.deepEqual(again.map(r => [r.payer, r.cred, r.besse, r.first, r.start, r.obs, r.match.join()]), rules.map(r => [r.payer, r.cred, r.besse, r.first, r.start, r.obs, r.match.join()]));
});
