import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const R = require('../rules.js');
const XL = require('../xlsx.js');

// v0.14 admin side of the Clinic Eligibility file: what bringBackClinicFiles reads from a file the clinic saved in the
// Clinic Template Tool (Clinic Work + hidden SE Info / SE Done, as XL.readWorkbook gives them), and the pane wiring.
// Synthetic data only (made-up MRNs 9000xx and payers).

const LOC = 'Synthetic Bay';
const RUN = new Date(Date.UTC(2026, 9, 22, 16));
const src = (s) => readFileSync(new URL(`../${s}`, import.meta.url), 'utf8');

/** The clinic's saved file: the admin's pack with Clinic Work filled in and the admin's Done marks on SE Done. */
async function returnedFile(work, sentDone) {
  const pack = R.clinicPack(LOC, { patients: [], results: [], runAt: RUN, version: '0.14.0' });
  pack.sheets = pack.sheets.map((s) => (s.name === R.SE_SHEETS.work ? { ...s, rows: work } : s.name === R.SE_SHEETS.done ? { ...s, rows: sentDone } : s));
  return XL.readWorkbook(XL.buildWorkbook(pack));
}
const sheetOf = (wb, name) => wb.sheets.find((x) => x.name.trim().toLowerCase() === name.toLowerCase());

test('a returned Clinic Eligibility file: Clinic Work and SE Info read by sheet name; the grids merge like objects', async () => {
  const at = (h) => new Date(Date.UTC(2026, 9, 23, h));
  const work = R.workRows([], [
    { Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'On', By: 'Front desk' },
    { Kind: 'Task done', Key: '900001|Fix member ID in NextGen', MRN: '900001', LASFIR: 'QUITES', Value: 'Yes', By: 'Front desk' },
    { Kind: 'Task done', Key: '900002|Update insurance in NextGen', MRN: '900002', LASFIR: 'BRABRA', Value: 'No', By: 'Front desk' },
    { Kind: 'Note', Key: '900002', MRN: '900002', LASFIR: 'BRABRA', Note: 'Brings the new card Monday', By: 'Front desk' },
    { Kind: 'Payer note', Key: 'Zenith CA', Value: 'Zenith CA PPO', Note: 'group 0012', By: 'Front desk' },
  ], at(9));
  const wb = await returnedFile(work, [{ 'Task Key': '900002|Update insurance in NextGen', Done: 'Yes' }]);
  assert.ok(!sheetOf(wb, 'Clinic Requests') && !sheetOf(wb, 'Clinic Info'), 'not mistaken for an old Sohar clinic file');
  const info = {};
  for (const r of R.tableObjects(sheetOf(wb, R.SE_SHEETS.info).rows)) info[String(r.Setting ?? '').trim()] = String(r.Value ?? '').trim();
  assert.equal(info.Kind, 'Clinic Eligibility');
  assert.equal(info.Location, LOC);
  const ctx = { location: info.Location, requests: [], notes: [], payerNotes: [], applied: [], now: at(10),
    tasks: [{ 'Task Key': '900001|Fix member ID in NextGen', Done: '', 'Next Location': LOC }, { 'Task Key': '900002|Update insurance in NextGen', Done: 'Yes', 'Next Location': LOC }] };
  const m = R.mergeClinicWork(sheetOf(wb, R.SE_SHEETS.work).rows, ctx);
  assert.deepEqual(m.requests, [{ MRN: '900001', Requested: at(9).toISOString(), By: `Clinic file ${LOC}` }]);
  assert.deepEqual([...m.done], [['900001|Fix member ID in NextGen', `Yes (clinic ${LOC})`]], 'the admin\'s own Done is not undone by the clinic');
  assert.deepEqual(m.notes, [{ MRN: '900002', Location: LOC, Note: 'Brings the new card Monday', By: 'Front desk', When: at(9).toISOString() }]);
  assert.deepEqual(m.payerNotes, [{ 'NextGen Payer': 'Zenith CA', Location: LOC, 'Card Says': 'Zenith CA PPO', Note: 'group 0012', By: 'Front desk', When: at(9).toISOString() }]);
  assert.deepEqual(m.counts, { reverifyAdded: 1, reverifyRemoved: 0, doneSet: 1, doneCleared: 0, doneNotes: 0, tasksChanged: 1, notes: 1, payerNotes: 1, skipped: 0, otherClinic: 0, old: 0, applied: 5 });
  assert.deepEqual(R.mergeClinicWork(work, ctx), m, 'grid and object rows agree');
  // The same saved file brought back a second time changes nothing.
  const twice = R.mergeClinicWork(sheetOf(wb, R.SE_SHEETS.work).rows, { ...ctx, requests: m.requests, notes: m.notes, payerNotes: m.payerNotes, applied: m.applied });
  assert.deepEqual([twice.counts.old, twice.counts.applied, twice.done.size, twice.counts.reverifyAdded], [5, 0, 0, 0]);
  // A blank Clinic Work sheet (the file sent back untouched) changes nothing.
  const none = R.mergeClinicWork(sheetOf(await returnedFile([], []), R.SE_SHEETS.work).rows, ctx);
  assert.deepEqual(none.counts, { reverifyAdded: 0, reverifyRemoved: 0, doneSet: 0, doneCleared: 0, doneNotes: 0, tasksChanged: 0, notes: 0, payerNotes: 0, skipped: 0, otherClinic: 0, old: 0, applied: 0 });
});

test('pane wiring: Clinic Template Tool file in the Admin tab and full screen, both note tables, the demo adapter', () => {
  const app = src('app.js'), dash = src('dashboard.js'), html = src('dashboard.html'), excel = src('excel.js');
  assert.match(app, /data-ctt="\$\{esc\(l\)\}">Clinic Template Tool file<\/button>/);
  assert.match(app, /m\.type === 'ctt-file'/);
  assert.match(dash, /toPane\(\{ type: 'ctt-file', location: \$\('cLoc'\)\.value \}\)/);
  assert.match(html, /id="btnCttFile"/);
  assert.match(app, /XL\.buildWorkbook\(pack\)/);
  for (const [t, sheet] of [['tblClinicNotes', 'Clinic Notes'], ['tblClinicPayerNotes', 'Clinic Payer Notes']]) {
    assert.match(excel, new RegExp(`${t}: \\{ sheet: '${sheet}' \\}`), `${t} is OWNED and visible`);
    assert.match(excel, new RegExp(`writeAddinTable\\('${sheet}', '${t}'`), `${t} is written through writeAddinTable`);
  }
  assert.match(excel, /tblClinicWorkApplied: \{ sheet: 'Clinic Work Applied', hidden: true \}/, 'the brought-back ledger is OWNED and hidden');
  assert.match(excel, /writeAddinTable\('Clinic Work Applied', 'tblClinicWorkApplied'/);
  for (const fn of ['readClinicNotes', 'writeClinicNotes', 'readClinicPayerNotes', 'writeClinicPayerNotes', 'readClinicWorkApplied', 'writeClinicWorkApplied']) {
    assert.match(app, new RegExp(`async ${fn}\\(`), `demo adapter has ${fn}`);
  }
  // Bring-back and the clinic file take Done from the NextGen Tasks sheet as it is now, not the copy from the upload.
  const bring = app.slice(app.indexOf('async function bringBackClinicFiles'), app.indexOf('function bringBackSummary'));
  assert.match(bring, /tasks = await liveTasks\(\)/);
  assert.doesNotMatch(bring, /currentTasks\(\)/);
  assert.match(app.slice(app.indexOf('async function cttWorkbook')), /^[\s\S]*?tasks: await liveTasks\(\)/);
  // The clinic files table keeps its buttons on their own row (fits a 350 px pane).
  assert.match(app, /<tr><td class="acts" colspan="3"><button class="small" data-cfile=/);
  assert.doesNotMatch(src('style.css'), /clinicfiles td\.acts \{[^}]*nowrap/);
});

test('release: VERSION, package.json and every ?v= agree', () => {
  const v = src('app.js').match(/const VERSION = '([\d.]+)';/)[1];
  assert.equal(v, '0.18.0');
  assert.equal(JSON.parse(src('package.json')).version, v);
  for (const f of ['index.html', 'dashboard.html', 'save.html']) {
    const found = [...src(f).matchAll(/\?v=([\d.]+)/g)].map((x) => x[1]);
    assert.ok(found.length, `${f} has ?v=`);
    assert.deepEqual([...new Set(found)], [v], `${f} ?v= all ${v}`);
  }
});
