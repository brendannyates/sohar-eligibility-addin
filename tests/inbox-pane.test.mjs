import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const C = require('../core.js');
const R = require('../rules.js');

// v0.14 round 2, admin pane: the clinic inbox line and done notes. The merge-and-save code in app.js is shared by
// "Bring back clinic files" and "Apply clinic changes"; it is run here as app.js has it, against an in-memory adapter.
// Synthetic data only: made-up MRNs 9008xx, made-up clinics and staff.

const src = (s) => readFileSync(new URL(`../${s}`, import.meta.url), 'utf8');
const app = src('app.js');
const cut = (from, to) => { const i = app.indexOf(from), j = app.indexOf(to, i); assert.ok(i >= 0 && j > i, `app.js has ${from}`); return app.slice(i, j); };

/** clinicMerge / mergeClinicRows / saveClinicMerge / bringBackSummary / inboxPlan from app.js, with the pane's helpers stubbed. */
function paneCode(X, state) {
  const body = cut('  function clinicMerge(', '  /** "2 clinic files read.') + cut('  function bringBackSummary(', '  // ---------- the clinic inbox')
    + cut('  async function inboxPlan(', '  async function loadInbox(');
  const calls = [];
  const liveTasks = async () => state.tasks;
  const make = new Function('R', 'C', 'X', 'state', 'liveTasks', 'build', 'renderClinicNotes', 'loadInbox',
    `${body}\nreturn { clinicMerge, clinicNews, mergeClinicRows, saveClinicMerge, bringBackSummary, inboxPlan, fileQuietRows };`);
  return { calls, ...make(R, C, X, state, liveTasks, async () => calls.push('build'), () => calls.push('render'), async () => calls.push('loadInbox')) };
}

function memoryAdapter(inbox) {
  const db = { requests: [], tasks: null, notes: [], payerNotes: [], applied: [] }, writes = [];
  const w = (k) => async (cols, rows) => { writes.push(k); db[k] = rows; return rows.length; };
  return {
    db, writes,
    async readClinicInbox() { return inbox; },
    async readReverifyRequests() { return db.requests; }, writeReverifyRequests: w('requests'),
    async readClinicNotes() { return db.notes; }, writeClinicNotes: w('notes'),
    async readClinicPayerNotes() { return db.payerNotes; }, writeClinicPayerNotes: w('payerNotes'),
    async readClinicWorkApplied() { return db.applied; }, writeClinicWorkApplied: w('applied'),
    writeTasks: w('tasks'),
  };
}

const ago = (h) => new Date(Date.now() - h * 36e5).toISOString();
const task = (key, o) => ({ ...Object.fromEntries(C.TASK_COLS.map((c) => [c, ''])), 'Task Key': key, ...(o || {}) });
const row = (Location, Kind, Key, Value, Note, When) => ({ Location, Kind, Key, MRN: Key.split('|')[0], LASFIR: 'SYNPAT', Value, Note, By: 'Front desk', When,
  File: `Clinic Eligibility - ${Location} 2026-10-23.xlsx`, Received: When });

test('inbox: the dry run counts new changes; Apply writes them with the bring-back writes (ledger last), then nothing is waiting', async () => {
  const A = '900801|Fix member ID in NextGen', B = '900802|Update benefits in NextGen', D = '900803|Fix name in NextGen';
  const inbox = [
    row('Synthetic Bay', 'Task done', A, 'Yes', 'Called the patient, new card scanned', ago(5)),
    row('synthetic bay', 'Task done', B, '', 'Copay on the card is $40', ago(4)), // a note only: Done stays as the admin has it
    row('Synthetic Hill', 'Task done', D, 'Yes', '', ago(3)),
    row('Synthetic Hill', 'Note', '900804', '', 'Bringing the new card on Monday', ago(2)),
    row('Synthetic Hill', 'Re-verify anyway', '900805', 'On', '', ago(1)),
    row('', 'Note', '900806', '', 'no location', ago(1)),
    row('Synthetic Hill', 'Mystery', '900807', '', '', ago(1)),
  ];
  const bay = { 'Next Location': 'Synthetic Bay' };
  const state = { tasks: [task(A, bay), task(B, { Done: 'Yes', ...bay }), task(D, { 'Clinic Note': 'Typed on the sheet', 'Next Location': 'Synthetic Hill' })] };
  const X = memoryAdapter(inbox), P = paneCode(X, state);
  const plan = await P.inboxPlan();
  assert.equal(plan.s.n.applied, 5, '5 new changes waiting');
  assert.equal(P.clinicNews(plan.s.n), 5, 'a tick and its note count once');
  assert.equal(plan.s.n.locations, 2, 'two clinics (location case does not split one)');
  assert.equal(plan.skipped, 2, 'blank location and unknown kind');
  assert.deepEqual(X.writes, [], 'a dry run saves nothing');
  await P.saveClinicMerge(plan.s);
  assert.deepEqual(X.writes, ['requests', 'tasks', 'notes', 'applied'], 'same writes as bring-back, the brought-back ledger last');
  const t = new Map(X.db.tasks.map((x) => [x['Task Key'], x]));
  assert.equal(t.get(A).Done, 'Yes (clinic Synthetic Bay)');
  assert.match(t.get(A)['Clinic Note'], /^Called the patient, new card scanned \(Front desk, Synthetic Bay, \d+\/\d+\)$/);
  assert.equal(t.get(B).Done, 'Yes', 'a note-only row leaves Done');
  assert.match(t.get(B)['Clinic Note'], /^Copay on the card is \$40 \(Front desk, Synthetic Bay,/);
  assert.equal(t.get(D).Done, 'Yes (clinic Synthetic Hill)');
  assert.equal(t.get(D)['Clinic Note'], 'Typed on the sheet', 'a blank note never erases one');
  assert.deepEqual(X.db.tasks.map((x) => Object.keys(x)), X.db.tasks.map(() => C.TASK_COLS), 'NextGen Tasks keeps its columns');
  assert.deepEqual(X.db.notes.map((x) => x.Note), ['Bringing the new card on Monday']);
  assert.deepEqual(X.db.requests.map((x) => x.MRN), ['900805']);
  assert.deepEqual(P.calls, ['build', 'render', 'loadInbox']);
  const summary = P.bringBackSummary({ ...plan.s.n, old: 0 }, 'Clinic inbox: 5 changes from 2 clinics applied.');
  assert.match(summary, /^Clinic inbox: 5 changes from 2 clinics applied\. 1 re-verify request added, 0 removed\./);
  assert.match(summary, /2 tasks marked done on NextGen Tasks\. 2 done notes saved in the Clinic Note column on NextGen Tasks\. 1 patient note saved on Clinic Notes\./);
  // Read again (the flow has not run since; or it appended one row twice): nothing new.
  inbox.push({ ...inbox[0] });
  const again = await P.inboxPlan();
  assert.equal(again.s.n.applied, 0);
  // The same row brought back in a returned file applies once too: the file path shares the ledger.
  const s = P.clinicMerge(X.db.requests);
  await P.mergeClinicRows(s, [inbox[3]].map(({ Location, File, Received, ...w }) => w), 'Synthetic Hill');
  assert.deepEqual([s.n.applied, s.n.old], [0, 1]);
  assert.equal(P.bringBackSummary({ ...s.n, read: 1 }), '1 clinic file read. 1 change was brought back before or is over 120 days old, so it was skipped (your later edits stay).');
});

test('inbox: a missing table is reported as missing (set-up button), a When that became a date serial still reads', async () => {
  const P = paneCode(memoryAdapter(null), { tasks: [] });
  assert.deepEqual(await P.inboxPlan(), { missing: true });
  const A = '900811|Fix member ID in NextGen';
  const serial = Date.now() / 864e5 + 25569 - 1 / 24; // an hour ago, as Excel would store it
  const X = memoryAdapter([row('Synthetic Bay', 'Task done', A, 'Yes', '', serial)]);
  const plan = await paneCode(X, { tasks: [task(A, { 'Next Location': 'Synthetic Bay' })] }).inboxPlan();
  assert.equal(plan.s.n.applied, 1);
});

test('inbox: a note edit after the admin reopened a clinic-ticked task leaves it open; rows that change nothing are filed, not waiting', async () => {
  const A = '900821|Add insurance in NextGen', B = '900822|Fix member ID in NextGen';
  const bay = { 'Next Location': 'Synthetic Bay' };
  const inbox = [row('Synthetic Bay', 'Task done', A, 'Yes', '', ago(5))];
  const state = { tasks: [task(A, bay), task(B, bay)] };
  const X = memoryAdapter(inbox), P = paneCode(X, state);
  let plan = await P.inboxPlan();
  await P.saveClinicMerge(plan.s);
  assert.equal(X.db.tasks.find((t) => t['Task Key'] === A).Done, 'Yes (clinic Synthetic Bay)');
  // The admin reopens A. The clinic then saves a done note (CTT carries the old Yes along, with a new When).
  state.tasks = X.db.tasks.map((t) => (t['Task Key'] === A ? { ...t, Done: '' } : t));
  inbox.push(row('Synthetic Bay', 'Task done', A, 'Yes', 'Card was for the old plan', ago(1)));
  plan = await P.inboxPlan();
  assert.deepEqual([plan.s.n.applied, plan.s.n.doneSet, plan.s.n.doneNotes, P.clinicNews(plan.s.n)], [1, 0, 1, 1], 'the note is news, the old tick is not');
  X.writes.length = 0;
  await P.saveClinicMerge(plan.s);
  const a = X.db.tasks.find((t) => t['Task Key'] === A);
  assert.equal(a.Done, '', 'still open, as the admin left it');
  assert.match(a['Clinic Note'], /^Card was for the old plan \(Front desk, Synthetic Bay,/);
  assert.equal(P.bringBackSummary({ ...plan.s.n, old: 0 }, 'Clinic inbox: 1 change from 1 clinic applied.'),
    'Clinic inbox: 1 change from 1 clinic applied. 1 done note saved in the Clinic Note column on NextGen Tasks.');
  // A blank note saved on B (never ticked, no note before) changes nothing: not waiting, filed in the ledger only.
  state.tasks = X.db.tasks;
  inbox.push(row('Synthetic Bay', 'Task done', B, '', '', ago(0.5)), row('Synthetic Bay', 'Note', '900822', '', '', ago(0.5)));
  plan = await P.inboxPlan();
  assert.deepEqual([plan.s.n.applied, P.clinicNews(plan.s.n)], [2, 0]);
  X.writes.length = 0;
  assert.equal(await P.fileQuietRows(plan), true);
  assert.deepEqual(X.writes, ['applied'], 'only the ledger');
  assert.deepEqual([(await P.inboxPlan()).s.n.applied, await P.fileQuietRows(await P.inboxPlan())], [0, false]);
});

test('inbox: a tick sent for a task whose next visit is at another clinic is skipped and said so', async () => {
  const A = '900831|Fix member ID in NextGen';
  const X = memoryAdapter([row('Synthetic Bay', 'Task done', A, 'Yes', 'done', ago(1))]);
  const P = paneCode(X, { tasks: [task(A, { 'Next Location': 'Synthetic Hill' })] });
  const plan = await P.inboxPlan();
  assert.deepEqual([plan.s.n.otherClinic, plan.s.n.doneSet, plan.s.n.doneNotes, P.clinicNews(plan.s.n)], [1, 0, 0, 1]);
  assert.match(P.bringBackSummary({ ...plan.s.n, old: 0 }, 'Clinic inbox: 1 change from 1 clinic applied.'),
    /1 task was ticked in a clinic file for a patient whose next visit is at another clinic, so it was skipped\./);
});

test('pane: Check again re-reads the inbox, focus lands on the result, an open that stopped still shows the inbox', () => {
  const html = src('index.html');
  assert.match(html, /<div id="clinicFilesNote" class="result" role="status" aria-live="polite" tabindex="-1" hidden><\/div>/);
  const load = cut('  async function loadInbox(', '  async function applyInbox(');
  assert.equal((load.match(/\$\{check\}/g) || []).length, 3, 'Check again on the error, not-set-up and normal lines');
  assert.match(load, /id="btnInboxCheck">Check again<\/button>/);
  assert.match(load, /const k = clinicNews\(p\.s\.n\)/);
  assert.match(app, /b\.id === 'btnInboxCheck'\) guard\(loadInbox, 'clinicFilesNote'\)\.then\(refocus\('btnInboxCheck'\)\)/);
  assert.match(app, /b\.id === 'btnInboxApply'\) guard\(applyInbox, 'clinicFilesNote'\)\.then\(refocus\('clinicFilesNote'\)\)/);
  assert.match(app, /b\.id === 'btnInboxSetup'\) guard\(setupInbox, 'clinicFilesNote'\)\.then\(refocus\('clinicFilesNote'\)\)/);
  const auto = cut('  async function autoRefresh(', '  /** Scan the Import sheet');
  assert.match(auto, /catch \(e\) \{[\s\S]*await loadInbox\(\);[\s\S]*\}/);
  assert.match(cut('  async function applyInbox(', '  async function setupInbox('), /'Clinic inbox: nothing to change\.'/);
});

test('task CSVs guard typed text against Excel formulas; the Sohar and tracker files are left exactly as they are', () => {
  const cols = ['Task', 'Clinic Note'];
  const rows = [{ Task: 'Fix member ID in NextGen', 'Clinic Note': '=HYPERLINK("http://example.invalid","open")' },
    { Task: '-left voicemail', 'Clinic Note': '+1 card scanned' }, { Task: '@sum', 'Clinic Note': '\tx' }, { Task: 'a-b', 'Clinic Note': '' }];
  assert.equal(C.toCSV(cols, rows, { safe: true }),
    'Task,Clinic Note\r\nFix member ID in NextGen,"\'=HYPERLINK(""http://example.invalid"",""open"")"\r\n\'-left voicemail,\'+1 card scanned\r\n\'@sum,\'\tx\r\na-b,\r\n');
  assert.equal(C.toCSV(cols, rows.slice(1, 2)), 'Task,Clinic Note\r\n-left voicemail,+1 card scanned\r\n', 'off by default');
  assert.match(app, /saveCSV\(C\.fileName\('tasks'\), C\.TASK_COLS, await currentTasks\(\), 'resultsResult', \{ safe: true \}\)/);
  const dash = src('dashboard.js');
  assert.match(dash, /data = tasks\.filter\(open\); safe = true; \}/);
  assert.match(dash, /data = tasks; safe = true; \}/);
  assert.match(dash, /C\.toCSV\(cols, data, \{ safe \}\)/);
  assert.equal((app.match(/\{ safe: true \}/g) || []).length, 1, 'no other pane download is guarded');
});

test('pane wiring: inbox line, set up, apply, done notes in the views and downloads, demo inbox', () => {
  const html = src('index.html'), dash = src('dashboard.js'), dhtml = src('dashboard.html');
  assert.match(html, /<div id="inboxBox" class="result"><\/div>/);
  const load = cut('  async function loadInbox(', '  async function applyInbox(');
  assert.match(load, /new clinic change', 'new clinic changes'\)\} waiting/);
  assert.match(load, /id="btnInboxApply">Apply clinic changes<\/button>/);
  assert.match(load, /id="btnInboxSetup">Set up the clinic inbox<\/button>/);
  const apply = cut('  async function applyInbox(', '  async function setupInbox(');
  assert.match(apply, /await inboxPlan\(\)/);
  assert.match(apply, /await saveClinicMerge\(p\.s\)/);
  assert.match(apply, /bringBackSummary\(/);
  const setup = cut('  async function setupInbox(', '  // ---------- clinic rules');
  assert.match(setup, /X\.createClinicInbox\(R\.INBOX_COLS\)/);
  assert.match(setup, /the sheet \$\{t\.sheet\} and the table \$\{t\.table\}/);
  const bring = cut('  async function bringBackClinicFiles(', '  function clinicMerge(');
  assert.match(bring, /await mergeClinicRows\(s, work\.rows, info\.Location\)/);
  assert.match(bring, /await saveClinicMerge\(s\)/);
  assert.doesNotMatch(bring, /R\.mergeClinicWork|X\.write/, 'bring-back has no merge or save code of its own');
  // liveTasks refreshes the Clinic Note from the sheet as well as Done.
  assert.match(cut('  async function liveTasks(', '  async function copyAdminMessage('), /'Clinic Note': cell\(s, 'Clinic Note'\)/);
  // Done notes show where Done shows (full screen task table and patient card) and in the task downloads.
  assert.match(dhtml, /<th>Details<\/th><th>Clinic note<\/th>/);
  assert.match(dash, /cell\(t\['Clinic Note'\], 'mid cnote'\)/);
  assert.match(dash, /<span class="cnote">Clinic: \$\{esc\(t\['Clinic Note'\]\)\}<\/span>/);
  assert.match(app, /saveCSV\(C\.fileName\('tasks'\), C\.TASK_COLS,/);
  assert.match(dash, /cols = C\.TASK_COLS; data = tasks\.filter\(open\)/);
  assert.equal(C.TASK_COLS[C.TASK_COLS.length - 1], 'Clinic Note', 'downloads end with Clinic Note');
  // The demo adapter has an inbox with two made-up rows, and can start without one (?noinbox).
  const demo = cut('    const inboxFirstRead = () => {', '    };');
  assert.equal((demo.match(/Kind: '/g) || []).length, 2);
  assert.match(app, /async readClinicInbox\(\) \{/);
  assert.match(app, /async createClinicInbox\(\) \{/);
  assert.match(app, /noinbox/);
});
