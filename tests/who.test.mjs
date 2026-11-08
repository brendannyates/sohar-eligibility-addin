import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const W = require('../who.js');
const src = (f) => readFileSync(new URL(`../${f}`, import.meta.url), 'utf8');

// v0.28: who is using the add-in. Made-up staff names only.

test('the saved list: cleaned, one row per name (any case), unreadable settings read as empty', () => {
  assert.deepEqual(W.usersFrom([{ name: '  Sam   O. ', added: '11/7/2026' }, { name: 'sam o.' }, 'Ana R.', null, { name: '' }]),
    [{ name: 'Sam O.', added: '11/7/2026' }, { name: 'Ana R.', added: '' }]);
  assert.deepEqual(W.usersFrom('[{"name":"Jo K."}]'), [{ name: 'Jo K.', added: '' }]);
  assert.deepEqual(W.usersFrom('not json'), []);
  assert.deepEqual(W.usersFrom({ name: 'x' }), []);
  assert.equal(W.cleanName('x'.repeat(40)).length, 30);
});

test('adding a name never duplicates it and keeps the list spelling', () => {
  const a = W.addUser([{ name: 'Sam O.' }], 'ana r.', '11/7/2026');
  assert.deepEqual(a.users.map((u) => u.name), ['Sam O.', 'ana r.']);
  assert.equal(a.name, 'ana r.');
  const b = W.addUser(a.users, 'SAM O.', '');
  assert.equal(b.users.length, 2);
  assert.equal(b.name, 'Sam O.');
  assert.deepEqual(W.addUser([], '   ', '').users, []);
});

test('the screen offers the workbook list, then Config "Admin names" not already on it', () => {
  const L = W.listFor([{ name: 'Sam O.' }], ['sam o.', 'Jo K.']);
  assert.deepEqual(L.map((u) => [u.name, !!u.config]), [['Sam O.', false], ['Jo K.', true]]);
});

test('order: names used on this computer first, most recent first, then A-Z', () => {
  const L = W.pickOrder([{ name: 'Zed' }, { name: 'ana' }, { name: 'Bo' }, { name: 'Cy' }], ['Cy', 'Zed']);
  assert.deepEqual(L.map((u) => [u.name, u.here]), [['Cy', true], ['Zed', true], ['ana', false], ['Bo', false]]);
});

test("today's pick lasts the day and only while the name is still offered", () => {
  const rec = W.whoPick(W.whoRec(''), 'Sam O.', '2026-11-07');
  assert.deepEqual(rec, { day: '2026-11-07', name: 'Sam O.', used: ['Sam O.'] });
  const list = [{ name: 'Sam O.' }];
  assert.equal(W.pickedToday(rec, '2026-11-07', list), 'Sam O.');
  assert.equal(W.pickedToday(rec, '2026-11-08', list), '', 'a new day asks again');
  assert.equal(W.pickedToday(rec, '2026-11-07', [{ name: 'Ana R.' }]), '', 'removed from the list: asked again');
  const next = W.whoPick(rec, 'Ana R.', '2026-11-08');
  assert.deepEqual(next.used, ['Ana R.', 'Sam O.']);
  assert.deepEqual(W.whoRec('{bad'), { day: '', name: '', used: [] });
  assert.equal(W.dayIso(new Date(2026, 0, 5)), '2026-01-05');
});

test('the workbook key: same for the same file (query and case ignored), per workbook in localStorage', () => {
  const k = W.wbKeyOf('https://example.org/Book.xlsx?web=1');
  assert.equal(k, W.wbKeyOf('HTTPS://example.org/book.xlsx'));
  assert.notEqual(k, W.wbKeyOf('https://example.org/Other.xlsx'));
  assert.equal(W.wbKeyOf(''), '');
  assert.equal(W.whoKey(k), `sohar.who:${k}`);
  assert.equal(W.whoKey(''), '');
});

test('wiring: names in document settings (no sheet), the loading and name screens, the bell in both windows', () => {
  const app = src('app.js'), dash = src('dashboard.js'), excel = src('excel.js');
  assert.match(excel, /Office\.context\.document\.settings\.get\(k\)/);
  assert.match(excel, /s\.saveAsync\(/);
  assert.match(app, /X\.setDoc\(W\.USERS_SETTING, r\.users\)/);
  assert.doesNotMatch(app + excel, /writeAddinTable\([^)]*[Uu]sers/, 'no sheet for the people');
  assert.match(app, /<h1>Enter your name<\/h1>/);
  assert.match(app, /Pick yours below, or add it if you’re new\./);
  assert.match(app, /I’m not on the list/);
  assert.doesNotMatch(app + dash, /at the desk/i);
  assert.match(app, /who: \{ name: who\.name, users: whoList\(\), here: whoRead\(\)\.used \}/);
  assert.match(app, /m\.type === 'who-pick'/);
  assert.match(dash, /toPane\(\{ type: 'who-pick', name: adminName, typed: !!typed \}\)/);
  // Dev Updates: seen ids relayed both ways over the dialog messages
  assert.match(app, /DevUpdates\.init\(\{ app: 'sohar'/);
  assert.match(dash, /DevUpdates\.init\(\{ app: 'sohar', send: inOffice \? \(m\) => toPane\(m\) : null \}\)/);
  assert.match(app, /if \(window\.DevUpdates && DevUpdates\.receive\(m\)\) return;/);
  assert.match(dash, /if \(window\.DevUpdates && DevUpdates\.receive\(m\)\) return;/);
  assert.match(app, /DevUpdates\.setSend\(\(m\) => sendToDialog\(m\)\)/);
  assert.match(app, /DevUpdates\.dialogUrl\(/);
  for (const f of ['index.html', 'dashboard.html']) {
    const h = src(f);
    assert.match(h, /<script src="who\.js\?v=[\d.]+"><\/script>\s*<script src="dev-updates\.js\?v=[\d.]+"><\/script>/, `${f} loads who.js and dev-updates.js`);
    assert.match(h, /href="dev-updates\.css\?v=[\d.]+"/);
    assert.match(h, /id="boot" class="boot"/);
    assert.match(h, /id="whoScreen" class="boot who" hidden/);
    // toolbar: icon tools with hover labels; no slot for the bell, which floats
    assert.doesNotMatch(h, /duSlot|du-slot/, `${f}: no bell slot in the toolbar`);
    const t0 = h.indexOf('class="title-actions tools"'), tools = h.slice(t0, h.indexOf('</span>', t0));
    const btns = tools.match(/<button[^>]*>/g) || [];
    assert.ok(btns.length >= 1);
    for (const b of btns) { assert.match(b, /class="ibtn"/); assert.match(b, /data-tip="[^"]+"/); assert.doesNotMatch(b, /title=/); }
    // the only fetch is the add-in's own dev-updates.json (same origin)
    assert.match(h, /connect-src 'self'/);
  }
  assert.match(src('save.html'), /connect-src 'none'/);
  assert.equal((src('dev-updates.js').match(/fetch\(/g) || []).length, 1);
  for (const f of ['app.js', 'dashboard.js', 'excel.js', 'who.js', 'core.js', 'rules.js', 'xlsx.js', 'record.js', 'audit.js', 'fees.js']) {
    assert.doesNotMatch(src(f), /\bfetch\(|XMLHttpRequest|sendBeacon|WebSocket/, `${f} makes no network call`);
  }
});

test('excel.js keeps the list in document settings and adds no sheet', async () => {
  const { createMock } = await import('./helpers/excel-mock.mjs');
  const m = createMock(), X = m.SoharExcel;
  const sheets = () => m.wb.sheets ? [...m.wb.sheets].length : 0;
  const before = sheets();
  assert.equal(X.getDoc(W.USERS_SETTING), null);
  await X.setDoc(W.USERS_SETTING, W.addUser([], 'Sam O.', '').users);
  await X.refreshDoc();
  assert.deepEqual(W.usersFrom(X.getDoc(W.USERS_SETTING)).map((u) => u.name), ['Sam O.']);
  await X.setDoc(W.USERS_SETTING, null);
  assert.equal(X.getDoc(W.USERS_SETTING), null);
  assert.equal(X.docUrl(), '');
  assert.equal(sheets(), before);
});

test('dev-updates.json: newest first, each notice has an id, a date and plain bullets', () => {
  const d = JSON.parse(src('dev-updates.json'));
  assert.ok(d.updates.length >= 1);
  const ids = new Set();
  for (const u of d.updates) {
    assert.match(u.id, /^\d{4}-\d{2}-\d{2}-[a-z0-9-]+$/);
    assert.match(u.date, /^\d{4}-\d{2}-\d{2}$/);
    assert.ok(Array.isArray(u.bullets) && u.bullets.length);
    assert.ok(!ids.has(u.id)); ids.add(u.id);
  }
  const dates = d.updates.map((u) => u.date);
  assert.deepEqual(dates, [...dates].sort().reverse());
});
