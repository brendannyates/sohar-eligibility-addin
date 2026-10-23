import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const C = require('../core.js');
const R = require('../rules.js');

// v0.14 round 2: done notes on NextGen Tasks (Clinic Note) and the central Clinic Work Inbox the Power Automate flow fills.
// Synthetic data only: made-up MRNs 9007xx, made-up clinic names and staff.
const NOW = new Date(Date.UTC(2026, 9, 23, 12));
// 09:00 UTC is the same calendar day from Los Angeles to Auckland, so the M/D in a done note does not depend on the zone.
const W = (Kind, Key, Value, When, o) => ({ Kind, Key, MRN: '', LASFIR: '', Value, Note: '', By: 'Front desk', When, ...(o || {}) });
const ctxOf = (o) => ({ location: 'San Rafael', requests: [], notes: [], payerNotes: [], applied: [], now: NOW, ...(o || {}) });

test('TASK_COLS ends with Clinic Note; keepDone carries it by Task Key (and the pre-v0.12 key) on its own, like Done', () => {
  assert.equal(C.TASK_COLS[C.TASK_COLS.length - 1], 'Clinic Note');
  assert.deepEqual(C.TASK_COLS.slice(0, -1).slice(-1), ['Task Key'], 'added at the end only');
  const t = (key, o) => ({ ...Object.fromEntries(C.TASK_COLS.map((c) => [c, ''])), MRN: '900701', Task: 'Fix member ID in NextGen', 'Task Key': key, ...(o || {}) });
  const fresh = [t('900701|Fix member ID in NextGen'), t('900701|Update benefits in NextGen', { Task: 'Update benefits in NextGen' }),
    t('900701|Fix name in NextGen', { Task: 'Fix name in NextGen' })];
  const prev = [
    { 'Task Key': '900701|Fix member ID in NextGen', Done: 'Yes (clinic San Rafael)', 'Clinic Note': 'new card scanned (Front desk, San Rafael, 10/23)' },
    { 'Task Key': '900701|Update benefits in NextGen', Done: '', 'Clinic Note': 'copay on the card is $40 (Front desk, San Rafael, 10/23)' },
    // Saved before v0.12: Per Nbr Task Key, the MRN column gives today's key.
    { 'Task Key': '701|Fix name in NextGen', MRN: '900701', 'Per Nbr': '701', Task: 'Fix name in NextGen', Done: 'x', 'Clinic Note': 'spelling fixed' },
  ];
  const kept = C.keepDone(fresh, prev);
  assert.deepEqual(kept.map((x) => [x.Done, x['Clinic Note']]), [
    ['Yes (clinic San Rafael)', 'new card scanned (Front desk, San Rafael, 10/23)'],
    ['', 'copay on the card is $40 (Front desk, San Rafael, 10/23)'],
    ['x', 'spelling fixed'],
  ]);
  assert.deepEqual(kept.map((x) => Object.keys(x)), fresh.map((x) => Object.keys(x)), 'same columns, same order');
  // Older rows without the column (saved before Clinic Note) keep working.
  assert.deepEqual(C.keepDone(fresh, [{ 'Task Key': '900701|Fix member ID in NextGen', Done: 'Yes' }]).map((x) => [x.Done, x['Clinic Note']]),
    [['Yes', ''], ['', ''], ['', '']]);
});

test('workRows: ticking a task keeps its done note, saving the note keeps the tick, an empty note clears it', () => {
  const key = '900701|Fix member ID in NextGen', at = (m) => new Date(Date.UTC(2026, 9, 23, 9, m));
  let work = R.workRows([], { Kind: 'Task done', Key: key, MRN: '900701', Value: 'Yes', By: 'Front desk' }, at(0));
  work = R.workRows(work, { Kind: 'Task done', Key: key, MRN: '900701', Note: 'called patient, new card scanned', By: 'Front desk' }, at(1));
  assert.deepEqual(work, [{ Kind: 'Task done', Key: key, MRN: '900701', LASFIR: '', Value: 'Yes', Note: 'called patient, new card scanned', By: 'Front desk', When: at(1).toISOString() }]);
  work = R.workRows(work, { Kind: 'Task done', Key: key, MRN: '900701', Value: 'No', By: 'Front desk' }, at(2));
  assert.deepEqual([work[0].Value, work[0].Note], ['No', 'called patient, new card scanned']);
  work = R.workRows(work, { Kind: 'Task done', Key: key, MRN: '900701', Note: '', By: 'Front desk' }, at(3));
  assert.deepEqual([work[0].Value, work[0].Note], ['No', '']);
  // A note saved with the box never ticked: no Value.
  const noteOnly = R.workRows([], { Kind: 'Task done', Key: key, Note: 'left a voicemail', By: 'Front desk' }, at(4));
  assert.deepEqual([noteOnly[0].Value, noteOnly[0].Note], ['', 'left a voicemail']);
});

test('mergeClinicWork: done notes go to doneNotes as "note (By, Location, M/D)"; a No still writes one, a blank one never erases', () => {
  const A = '900701|Fix member ID in NextGen', B = '900702|Update insurance in NextGen', Cc = '900703|Add insurance in NextGen', D = '900704|Fix name in NextGen';
  const SR = { 'Next Location': 'San Rafael' };
  const tasks = [{ 'Task Key': A, Done: '', ...SR }, { 'Task Key': B, Done: 'Yes (clinic San Rafael)', ...SR },
    { 'Task Key': Cc, Done: 'Yes (clinic San Rafael)', 'Clinic Note': 'earlier note', ...SR }, { 'Task Key': D, Done: 'x', ...SR }];
  const work = [
    W('Task done', A, 'Yes', '2026-10-22T09:00:00Z', { Note: 'called patient, new card scanned' }),
    W('Task done', B, 'No', '2026-10-22T09:01:00Z', { Note: 'card was for the old plan', By: 'Robin' }),
    W('Task done', Cc, 'Yes', '2026-10-22T09:02:00Z'),
    W('Task done', D, '', '2026-10-22T09:03:00Z', { Note: 'left a voicemail' }),
    W('Task done', 'GONE|Resend to Sohar', 'Yes', '2026-10-22T09:04:00Z', { Note: 'n/a' }),
  ];
  const m = R.mergeClinicWork(work, ctxOf({ tasks }));
  assert.deepEqual([...m.doneNotes], [
    [A, 'called patient, new card scanned (Front desk, San Rafael, 10/22)'],
    [B, 'card was for the old plan (Robin, San Rafael, 10/22)'],
    [D, 'left a voicemail (Front desk, San Rafael, 10/22)'],
  ], 'C has no new note (its earlier one stays), the gone task is skipped');
  assert.deepEqual([...m.done], [[A, 'Yes (clinic San Rafael)'], [B, '']], 'a blank Value (note only) leaves D\'s Done');
  assert.deepEqual([m.counts.doneNotes, m.counts.doneSet, m.counts.doneCleared, m.counts.skipped], [3, 1, 1, 1]);
  assert.equal(R.mergeClinicWork([W('Task done', A, 'Yes', '2026-10-22T09:00:00Z', { By: '' })], ctxOf({ tasks, location: '' })).doneNotes.size, 0, 'no note, nothing');
});

test('mergeClinicWork: rows older than APPLIED_DAYS are skipped as old and never re-applied once the ledger is trimmed', () => {
  const A = '900701|Fix member ID in NextGen';
  const tasks = [{ 'Task Key': A, Done: '', 'Next Location': 'San Rafael' }];
  const old = new Date(NOW.getTime() - (R.APPLIED_DAYS + 1) * 864e5).toISOString();
  const m = R.mergeClinicWork([W('Task done', A, 'Yes', old, { Note: 'long ago' }), W('Re-verify anyway', '900701', 'On', old)], ctxOf({ tasks }));
  assert.deepEqual([m.counts.old, m.counts.applied, m.done.size, m.doneNotes.size, m.requests.length, m.applied.length], [2, 0, 0, 0, 0, 0]);
  const recent = new Date(NOW.getTime() - (R.APPLIED_DAYS - 1) * 864e5).toISOString();
  assert.equal(R.mergeClinicWork([W('Task done', A, 'Yes', recent)], ctxOf({ tasks })).counts.doneSet, 1);
});

test('inboxWork: rows by Location (first seen, any case), oldest When first, in Clinic Work shape; blank, unknown and old rows counted', () => {
  assert.deepEqual(R.INBOX_COLS, ['Location', 'Kind', 'Key', 'MRN', 'LASFIR', 'Value', 'Note', 'By', 'When', 'File', 'Received']);
  const I = (Location, Kind, Key, Value, When, o) => ({ Location, ...W(Kind, Key, Value, When, o), File: `Clinic Eligibility - ${Location} 2026-10-22.xlsx`, Received: '2026-10-23T10:00:00Z' });
  const old = new Date(NOW.getTime() - (R.APPLIED_DAYS + 1) * 864e5).toISOString();
  const rows = [
    I('San Rafael', 'Task done', '900701|Fix member ID in NextGen', 'No', '2026-10-22T09:05:00Z'),
    I('Walnut Creek', 'Note', '900711', '', '2026-10-22T09:00:00Z', { Note: 'Brings the card Monday' }),
    I('san rafael', 'Task done', '900701|Fix member ID in NextGen', 'Yes', '2026-10-22T09:00:00Z', { Note: 'new card' }),
    I('', 'Note', '900712', '', '2026-10-22T09:00:00Z'),
    I('San Rafael', '', '900701', 'On', '2026-10-22T09:00:00Z'),
    I('San Rafael', 'Re-verify anyway', '', 'On', '2026-10-22T09:00:00Z'),
    I('San Rafael', 'Re-verify anyway', '900701', 'On', ''),
    I('San Rafael', 'Delete patient', '900701', 'Yes', '2026-10-22T09:00:00Z'),
    I('Walnut Creek', 'Re-verify anyway', '900711', 'On', old),
  ];
  const m = R.inboxWork(rows, NOW);
  assert.deepEqual([...m.keys()], ['San Rafael', 'Walnut Creek']);
  assert.deepEqual(m.get('San Rafael').map((r) => [r.Value, r.When]), [['Yes', '2026-10-22T09:00:00Z'], ['No', '2026-10-22T09:05:00Z']], 'the later edit last');
  assert.deepEqual(Object.keys(m.get('Walnut Creek')[0]), R.CLINIC_WORK_COLS);
  assert.deepEqual(m.counts, { rows: 3, blank: 4, unknown: 1, old: 1 });
  // The table as a grid (header first) reads the same.
  const grid = [R.INBOX_COLS, ...rows.map((r) => R.INBOX_COLS.map((c) => r[c]))];
  assert.deepEqual([...R.inboxWork(grid, NOW)], [...m]);
  assert.deepEqual(R.inboxWork(grid, NOW).counts, m.counts);
  assert.deepEqual([...R.inboxWork([], NOW)], []);
});

test('inbox apply: mergeClinicWork per Location, threaded; a re-read, or the same row also brought back in the file, applies once', () => {
  const A = '900701|Fix member ID in NextGen', K = '900711|Update benefits in NextGen';
  const rows = [
    { Location: 'San Rafael', ...W('Task done', A, 'Yes', '2026-10-22T09:00:00Z', { Note: 'new card scanned' }), File: 'f1.xlsx', Received: '2026-10-23T10:00:00Z' },
    { Location: 'San Rafael', ...W('Re-verify anyway', '900702', 'On', '2026-10-22T09:01:00Z'), File: 'f1.xlsx', Received: '2026-10-23T10:00:00Z' },
    { Location: 'Walnut Creek', ...W('Task done', K, 'Yes', '2026-10-22T09:02:00Z'), File: 'f2.xlsx', Received: '2026-10-23T10:01:00Z' },
    { Location: 'Walnut Creek', ...W('Payer note', 'Zenith CA', 'Zenith CA PPO', '2026-10-22T09:03:00Z', { Note: 'group 0012' }), File: 'f2.xlsx', Received: '2026-10-23T10:01:00Z' },
  ];
  const apply = (state, inbox) => {
    let s = { ...state }, tasks = state.tasks;
    const total = { doneSet: 0, doneNotes: 0, reverifyAdded: 0, payerNotes: 0, old: 0, applied: 0 };
    for (const [location, work] of inbox) {
      const m = R.mergeClinicWork(work, { location, requests: s.requests, tasks, notes: s.notes, payerNotes: s.payerNotes, applied: s.applied, now: NOW });
      tasks = tasks.map((t) => {
        const k = t['Task Key'];
        return m.done.has(k) || m.doneNotes.has(k) ? { ...t, ...(m.done.has(k) ? { Done: m.done.get(k) } : {}), ...(m.doneNotes.has(k) ? { 'Clinic Note': m.doneNotes.get(k) } : {}) } : t;
      });
      s = { requests: m.requests, notes: m.notes, payerNotes: m.payerNotes, applied: m.applied, tasks };
      for (const k of Object.keys(total)) total[k] += m.counts[k];
    }
    return { s, total };
  };
  const start = { requests: [], notes: [], payerNotes: [], applied: [], tasks: [{ 'Task Key': A, Done: '', 'Clinic Note': '', 'Next Location': 'San Rafael' },
    { 'Task Key': K, Done: '', 'Clinic Note': '', 'Next Location': 'Walnut Creek' }] };
  const first = apply(start, R.inboxWork(rows, NOW));
  assert.deepEqual(first.s.tasks, [{ 'Task Key': A, Done: 'Yes (clinic San Rafael)', 'Clinic Note': 'new card scanned (Front desk, San Rafael, 10/22)', 'Next Location': 'San Rafael' },
    { 'Task Key': K, Done: 'Yes (clinic Walnut Creek)', 'Clinic Note': '', 'Next Location': 'Walnut Creek' }]);
  assert.deepEqual(first.s.requests.map((r) => [r.MRN, r.By]), [['900702', 'Clinic file San Rafael']]);
  assert.deepEqual(first.s.payerNotes.map((r) => [r['NextGen Payer'], r.Location]), [['Zenith CA', 'Walnut Creek']]);
  assert.deepEqual(first.total, { doneSet: 2, doneNotes: 1, reverifyAdded: 1, payerNotes: 1, old: 0, applied: 4 });
  // The admin clears A's Done by hand; reading the inbox again (and the flow appending a duplicate) changes nothing.
  const cleared = { ...first.s, tasks: first.s.tasks.map((t) => (t['Task Key'] === A ? { ...t, Done: '' } : t)) };
  const again = apply(cleared, R.inboxWork(rows.concat([rows[0]]), NOW));
  assert.deepEqual(again.s.tasks, cleared.tasks);
  assert.deepEqual(again.total, { doneSet: 0, doneNotes: 0, reverifyAdded: 0, payerNotes: 0, old: 4, applied: 0 });
  // The clinic also sent the file back: its Clinic Work rows are the ones already applied from the inbox.
  const file = R.mergeClinicWork(R.inboxWork(rows, NOW).get('San Rafael'), { location: 'San Rafael', ...cleared, now: NOW });
  assert.deepEqual([file.counts.old, file.counts.applied, file.done.size, file.doneNotes.size], [2, 0, 0, 0]);
});
