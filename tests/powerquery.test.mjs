// v0.30: the Power Query files in docs/powerquery are generated from rules.js (make-pq.mjs) and must be current: each
// names its output table, and the column lists match what the add-in reads. Power Query itself can't run here, so the
// M is checked for shape only; refresh it once in Excel (docs/powerquery/README.md).
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
const R = createRequire(import.meta.url)('../rules.js');
const read = (f) => readFileSync(new URL(`../docs/powerquery/${f}`, import.meta.url), 'utf8');

test('each query file names its table, and the shared column lists match rules.js', () => {
  const want = { 'qNGAppts.pq': 'tblNGAppts', 'qLedger.pq': 'tblEligibilityChecks', 'qHistory.pq': 'tblPatientHistory', 'qCards.pq': 'tblCardOnFile',
    'qCopays.pq': 'tblCopayActivity', 'qOpenTickets.pq': 'tblOpenTickets', 'qBalances.pq': 'tblPatientBalances' };
  const cols = { 'qLedger.pq': R.LEDGER_COLS, 'qHistory.pq': R.PATIENT_HISTORY_COLS, 'qCards.pq': R.CARD_ON_FILE_COLS, 'qCopays.pq': R.COPAY_ACTIVITY_COLS };
  for (const [f, table] of Object.entries(want)) {
    const s = read(f);
    assert.ok(s.split('\n').some((l) => l.startsWith(`// Output: ${table} `)), `${f} -> ${table}`);
    assert.match(s, /^let\r?\n/m, `${f} is a let expression`);
    assert.equal((s.match(/\(/g) || []).length, (s.match(/\)/g) || []).length, `${f}: parentheses balance`);
    assert.equal((s.match(/\{/g) || []).length, (s.match(/\}/g) || []).length, `${f}: braces balance`);
    if (cols[f]) assert.deepEqual(s.match(/^\/\/ Columns: (.*)$/m)[1].trim().split(' | '), cols[f], `${f} columns`);
    assert.doesNotMatch(s, /sharepoint\.com\/personal|@mymhs|byates/i, `${f}: no personal paths`);
  }
  assert.match(read('pPaths.pq'), /<your-tenant>/, 'the paths are placeholders in the repo');
});

test('the files are current with make-pq.mjs (run node docs/powerquery/make-pq.mjs after changing rules.js columns)', async () => {
  const before = Object.fromEntries(['qLedger.pq', 'qHistory.pq', 'qCards.pq', 'qCopays.pq'].map((f) => [f, read(f)]));
  await import(`../docs/powerquery/make-pq.mjs?t=${Date.now()}`);
  for (const [f, text] of Object.entries(before)) assert.equal(read(f), text, `${f} was out of date`);
});
