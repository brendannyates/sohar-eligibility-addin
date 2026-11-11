// v0.29 flows: the Office Scripts the Power Automate flows run (docs/flows/scripts/*.ts) and the rotation's client
// (docs/flows/rotation-push.mjs). The scripts are TypeScript; node strips the types and runs them against a small fake
// ExcelScript workbook. Synthetic data only: made-up MRNs 9001xx, made-up staff, San Rafael and Castro.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire, stripTypeScriptTypes } from 'node:module';
import vm from 'node:vm';
import * as RP from '../docs/flows/rotation-push.mjs';
const R = createRequire(import.meta.url)('../rules.js');

function load(file, names) {
  const js = stripTypeScriptTypes(readFileSync(new URL(`../docs/flows/scripts/${file}`, import.meta.url), 'utf8'));
  const ctx = vm.createContext({ console, JSON, Date, Math, Object, Array, String, Number, isNaN });
  vm.runInContext(`${js}\n;this.__x = { ${names.join(', ')} };`, ctx);
  return ctx.__x;
}
const RT = load('rotation-to-master.ts', ['planUpsert', 'openNeeds', 'main']);
const LK = load('clinic-lookup.ts', ['decide', 'main']);
const SM = load('lookup-summary.ts', ['summarize', 'main']);
const plain = (v) => JSON.parse(JSON.stringify(v));

// A fake ExcelScript workbook: tables as {header, body}; writes are recorded.
function book(tables) {
  const writes = [];
  const tbl = (name) => {
    const t = tables[name];
    if (!t) return undefined;
    const body = () => ({ getValues: () => t.body.map((r) => r.slice()), getRow: (i) => ({ setNumberFormats() {}, setValues: (v) => { writes.push(['set', name, i]); t.body[i] = v[0].slice(); } }) });
    return {
      getHeaderRowRange: () => ({ getValues: () => [t.header.slice()] }),
      getRangeBetweenHeaderAndTotal: body,
      getColumn: () => ({ getRangeBetweenHeaderAndTotal: () => ({ setNumberFormatLocal() {} }) }),
      addRows: (i, rows) => { writes.push(['add', name, rows.length]); for (const r of rows) t.body.push(r.slice()); },
    };
  };
  return { wb: { getTable: tbl }, writes, tables };
}
const blankRow = (cols) => cols.map(() => '');

test('rotation script: upsert updates by key (blank cells keep what is there), fills the starter row, adds the rest; one row at a time', () => {
  const t = { tblEligibilityChecks: { header: R.LEDGER_COLS.slice(), body: [blankRow(R.LEDGER_COLS)] } };
  const { wb, writes } = book(t);
  const rows = RP.ledgerRows([
    { MRN: '0900101', 'Appt Date': '11/11/2026', Time: '9:00 AM', Location: 'San Rafael', 'Sohar Status': 'complete.eligible', Copay: '30', 'Not A Column': 'x' },
    { MRN: '900102', 'Appt Date': '11/12/2026', Time: '1:00 PM', Location: 'Castro', Copay: '20' },
  ], R.LEDGER_COLS);
  assert.equal(rows[0]['Row Key'], '0900101|2026-11-11|9:00 AM');
  assert.equal(rows[0]['Appt Date'], '2026-11-11', 'the master keeps ISO days');
  assert.ok(!('Not A Column' in rows[0]));
  let out = JSON.parse(RT.main(wb, JSON.stringify({ op: 'upsert', table: 'ledger', rows })));
  assert.deepEqual(out, { ok: true, updated: 1, added: 1, skipped: 0 });
  assert.equal(t.tblEligibilityChecks.body.length, 2);
  // the next run: a new copay, a blank Sohar Status (keeps the old one), and a row with no key
  out = JSON.parse(RT.main(wb, JSON.stringify({ op: 'upsert', table: 'ledger', rows: [{ 'Row Key': '0900101|2026-11-11|9:00 AM', Copay: '35', 'Sohar Status': '' }, { Copay: '1' }] })));
  assert.deepEqual(out, { ok: true, updated: 1, added: 0, skipped: 1 });
  const row = Object.fromEntries(R.LEDGER_COLS.map((c, i) => [c, t.tblEligibilityChecks.body[0][i]]));
  assert.equal(row.Copay, '35');
  assert.equal(row['Sohar Status'], 'complete.eligible');
  assert.ok(writes.every((w) => w[0] === 'set' || w[0] === 'add'), 'never a whole-table rewrite');
  assert.match(RT.main(wb, JSON.stringify({ op: 'upsert', table: 'nope', rows: [] })), /Unknown op or table/);
  assert.match(RT.main(book({}).wb, JSON.stringify({ op: 'upsert', table: 'ledger', rows: [] })), /Set up the rotation tables/);
  // card on file keys on the MRN, leading zeros or not
  const c = { tblCardOnFile: { header: R.CARD_ON_FILE_COLS.slice(), body: [['0900101', 'LARWEN', 'No', 'Luma checklist', '2026-11-01']] } };
  JSON.parse(RT.main(book(c).wb, JSON.stringify({ op: 'upsert', table: 'cards', rows: [{ MRN: '900101', 'Card On File': 'Yes', 'As Of': '2026-11-09' }] })));
  assert.deepEqual(c.tblCardOnFile.body, [['0900101', 'LARWEN', 'Yes', 'Luma checklist', '2026-11-09']]);
});

test('rotation script: needs-open lists open requests; needs-done marks them and never adds a row', () => {
  const H = R.NEEDS_CHECK_COLS;
  const row = (k, st) => H.map((h) => ({ 'Row Key': k, MRN: '900109', Location: 'Castro', Status: st }[h] || ''));
  const t = { tblNeedsCheck: { header: H.slice(), body: [row('r1', 'Open'), row('r2', 'Done'), row('r3', '')] } };
  const { wb } = book(t);
  assert.deepEqual(plain(JSON.parse(RT.main(wb, JSON.stringify({ op: 'needs-open' }))).rows.map((r) => r['Row Key'])), ['r1', 'r3']);
  const out = JSON.parse(RT.main(wb, JSON.stringify({ op: 'needs-done', rows: [{ 'Row Key': 'r1', Status: 'Done', 'Done At': '2026-11-11T15:00:00Z', Result: 'Discovery sent', MRN: '999' }, { 'Row Key': 'zz', Status: 'Done' }] })));
  assert.deepEqual(out, { ok: true, updated: 1, added: 0, skipped: 1 });
  const r1 = Object.fromEntries(H.map((h, i) => [h, t.tblNeedsCheck.body[0][i]]));
  assert.deepEqual([r1.Status, r1.Result, r1.MRN], ['Done', 'Discovery sent', '900109'], 'only the four needs-done columns change');
  assert.equal(t.tblNeedsCheck.body.length, 3);
});

const L = (o) => Object.fromEntries(R.LEDGER_COLS.map((c) => [c, o[c] || '']));
const LEDGER = [
  L({ MRN: '900101', LASFIR: 'LARWEN', 'Appt Date': '2026-11-11', Time: '9:00 AM', Location: 'San Rafael', 'Sohar Status': 'complete.eligible', 'Sohar Payer': 'Cigna', Copay: '30', 'Deductible Remaining': '100', 'Last Checked': '2026-11-10T15:00:00Z' }),
  L({ MRN: '900102', LASFIR: 'QUIAMB', 'Appt Date': '2026-09-01', Time: '1:00 PM', Location: 'Castro', 'Sohar Status': 'complete.eligible', 'Sohar Payer': 'Cigna', Copay: '20', 'Last Checked': '2026-09-01T15:00:00Z' }),
  L({ MRN: '900103', LASFIR: 'TUMPER', 'Appt Date': '2026-11-20', Time: '11:00 AM', Location: 'Castro', 'Sohar Status': 'complete.ineligible', 'Sohar Payer': 'Blue Shield', 'Last Checked': '2026-11-10T15:00:00Z' }),
];
const HIST = [1, 2, 3, 4, 5, 6].map((i) => ({ MRN: '900101', When: `2026-11-0${i}T16:00:00Z`, Kind: 'Copay', Detail: 'Luma charge', Amount: '30' }));
const ask = (o) => Object.assign({ clinic: 'San Rafael', mrn: '', purpose: 'unscheduled', reason: '', note: '', staff: 'Desk A', requestId: 'req-1', today: '2026-11-10', now: '2026-11-10T18:00:00Z' }, o);

test('lookup script: the four tiers, reasons and notes, minimum fields, and a log row for every call', () => {
  const D = (o) => plain(LK.decide(ask(o), LEDGER, HIST, [{ MRN: '900101', 'Card On File': 'Yes' }], [{ MRN: '900101', 'OK to Collect': 'Yes', When: '2026-11-09T17:00:00Z' }]));
  // scheduled here: no reason needed; summary and the last 5 events
  let a = D({ mrn: '900101', purpose: 'scheduled' });
  assert.equal(a.body.tier, 'scheduled');
  assert.equal(a.body.ok, true);
  assert.equal(a.body.patient.Coverage, 'Eligible');
  assert.equal(a.body.patient['Card On File'], 'Yes');
  assert.equal(a.body.patient['OK to Collect'], 'Yes');
  assert.equal(a.body.patient['Next Visit Here'], '2026-11-11 9:00 AM');
  assert.equal(a.body.history.length, 5);
  assert.equal(a.log.Outcome, 'ok');
  assert.match(a.log.Fields, /history 5$/);
  // seen at the clinic in the last 12 months: a reason is needed, then summary without history
  assert.equal(D({ clinic: 'Castro', mrn: '900102' }).body.error, 'need-reason');
  a = D({ clinic: 'Castro', mrn: '900102', reason: 'Patient called this clinic' });
  assert.deepEqual([a.body.tier, a.body.ok, a.body.history.length], ['seen-12mo', true, 0]);
  // in the master, never seen here: reason and note, then coverage, copay and card only
  a = D({ mrn: '900103', reason: 'Walk-in or same-day add' });
  assert.deepEqual([a.body.tier, a.body.error, a.body.needNote], ['no-history', 'need-note', true]);
  a = D({ mrn: '900103', reason: 'Walk-in or same-day add', note: 'moving to San Rafael' });
  assert.deepEqual(Object.keys(a.body.patient), ['LASFIR', 'MRN', 'Coverage', 'Payer', 'Copay', 'Coinsurance', 'Card On File', 'As Of']);
  assert.ok(!('Deductible Remaining' in a.body.patient), 'minimum fields');
  assert.equal(a.log.Reason, 'Walk-in or same-day add · moving to San Rafael');
  // "scheduled" from a clinic the patient is not booked at still needs a reason
  assert.equal(D({ clinic: 'Castro', mrn: '900101', purpose: 'scheduled' }).body.error, 'need-reason');
  // not in the master: nothing shown; a request adds a Needs Check row
  a = D({ mrn: '900199', reason: 'Walk-in or same-day add' });
  assert.deepEqual([a.body.tier, a.body.ok, a.body.patient], ['not-found', true, undefined]);
  a = D({ mrn: '900199', reason: 'Walk-in or same-day add', purpose: 'request-check' });
  assert.equal(a.body.queued, true);
  assert.deepEqual([a.needs.MRN, a.needs.Location, a.needs.Status, a.needs['Requested By']], ['900199', 'San Rafael', 'Open', 'Desk A']);
  assert.equal(D({ mrn: '900101', reason: 'x', purpose: 'request-check' }).body.message, 'This patient is already in the master, so no check was added.');
  assert.equal(D({ mrn: 'LARWEN' }).body.error, 'bad-mrn');
  assert.deepEqual(Object.keys(a.log), R.LOOKUP_LOG_COLS, 'the log row has the audit table\'s columns');
  for (const k of Object.keys(a.log)) assert.ok(!/copay|deductible|payer/i.test(k));
});

test('lookup script main: reads the master tables, appends the Needs Check row, nothing else written', () => {
  const t = { tblEligibilityChecks: { header: R.LEDGER_COLS.slice(), body: LEDGER.map((r) => R.LEDGER_COLS.map((c) => r[c])) },
    tblNeedsCheck: { header: R.NEEDS_CHECK_COLS.slice(), body: [blankRow(R.NEEDS_CHECK_COLS)] } };
  const { wb, writes } = book(t);
  let out = JSON.parse(LK.main(wb, JSON.stringify(ask({ mrn: '900101', purpose: 'scheduled' }))));
  assert.equal(out.body.tier, 'scheduled');
  assert.deepEqual(writes, [], 'a lookup writes nothing');
  out = JSON.parse(LK.main(wb, JSON.stringify(ask({ mrn: '900188', purpose: 'request-check', reason: 'Walk-in or same-day add' }))));
  assert.equal(out.body.queued, true);
  assert.deepEqual(plain(writes), [['add', 'tblNeedsCheck', 1]]);
  const q = Object.fromEntries(R.NEEDS_CHECK_COLS.map((c, i) => [c, t.tblNeedsCheck.body[1][i]]));
  assert.deepEqual([q['Row Key'], q.MRN, q.Status], ['req-1', '900188', 'Open']);
  assert.equal(JSON.parse(LK.main(wb, 'not json')).status, 400);
});

test('lookup summary: Sunday to Saturday on the Pacific calendar, counts only, heavy unscheduled lookups flagged', () => {
  const rows = [];
  for (let i = 0; i < 11; i++) rows.push({ When: '2026-11-10T20:00:00Z', Clinic: 'Castro', Staff: 'Desk B', MRN: `9001${i}`, Tier: 'seen-12mo', Outcome: 'ok', Reason: 'Patient called this clinic' });
  rows.push({ When: '2026-11-09T01:00:00Z', Clinic: 'San Rafael', Staff: 'Desk A', MRN: '900101', Tier: 'scheduled', Outcome: 'ok' }); // Sat 11/8 Pacific: last week
  rows.push({ When: '2026-11-09T17:00:00Z', Clinic: 'San Rafael', Staff: 'Desk A', MRN: '900101', Tier: 'scheduled', Outcome: 'ok' });
  rows.push({ When: '2026-11-10T17:00:00Z', Clinic: '', Staff: 'x', MRN: '', Tier: '', Outcome: 'bad-key' });
  const s = plain(SM.summarize(rows, '2026-11-09', -8));
  assert.equal(s.total, 13);
  assert.deepEqual(s.flagged, ['Castro · Desk B']);
  assert.match(s.subject, /week of 2026-11-09: 13, 1 to look at/);
  assert.ok(!/9001\d/.test(s.html), 'no MRNs in the email');
  assert.ok(!s.html.includes('Patient called'), 'no reasons in the email');
});

test('rotation-push: batches, the key from the environment, readable errors', async () => {
  assert.equal(RP.batches(Array.from({ length: 450 }, (_, i) => i)).map((b) => b.length).join(','), '200,200,50');
  await assert.rejects(RP.push('ledger', [{}], {}), /Set CLINIC_ROTATION_URL/);
  const sent = [];
  const fetchFn = async (u, o) => { sent.push(JSON.parse(o.body)); return { ok: true, status: 200, json: async () => ({ ok: true, updated: 1, added: 2, skipped: 0, rows: [{ 'Row Key': 'r1' }] }) }; };
  const env = { CLINIC_ROTATION_URL: 'https://flow.example.test/x', CLINIC_ROTATION_KEY: 'synthetic-key' };
  assert.deepEqual(await RP.push('history', Array.from({ length: 250 }, (_, i) => ({ 'Row Key': `h${i}` })), env, fetchFn), { updated: 2, added: 4, skipped: 0 });
  assert.deepEqual(sent.map((b) => [b.op, b.table, b.rows.length, b.key]), [['upsert', 'history', 200, 'synthetic-key'], ['upsert', 'history', 50, 'synthetic-key']]);
  assert.deepEqual(await RP.needsOpen(env, fetchFn), [{ 'Row Key': 'r1' }]);
  await RP.needsDone([{ 'Row Key': 'r1', Result: 'Verification sent' }], env, fetchFn);
  assert.equal(sent.pop().rows[0].Status, 'Done');
  await assert.rejects(RP.push('ledger', [{ MRN: '1', 'Appt Date': '2026-11-11' }], env, async () => ({ ok: false, status: 401, json: async () => ({ ok: false, error: 'bad-key' }) })), /401 bad-key/);
});

test('the build sheet carries the scripts as they are (run node docs/flows/make-build-sheet.mjs after changing one)', () => {
  const html = readFileSync(new URL('../docs/flows/Clinic Workbook flows - Build Sheet.html', import.meta.url), 'utf8');
  const DATA = JSON.parse(html.match(/const DATA = (\{.*\});\r?\n/)[1]);
  const files = { rotation: 'scripts/rotation-to-master.ts', lookup: 'scripts/clinic-lookup.ts', summary: 'scripts/lookup-summary.ts', push: 'rotation-push.mjs' };
  for (const [k, f] of Object.entries(files)) assert.equal(DATA.CODE[k], readFileSync(new URL(`../docs/flows/${f}`, import.meta.url), 'utf8'), `${f} is on the page as it is`);
  assert.deepEqual(DATA.COLS.log, R.LOOKUP_LOG_COLS);
  assert.ok(!/sig=|CLINIC_ROTATION_KEY = '[^R]/.test(html), 'no flow signature or key on the page');
});
