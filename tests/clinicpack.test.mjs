import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
const require = createRequire(import.meta.url);
const C = require('../core.js');
const R = require('../rules.js');
const PR = require('../payer-rules.js');
const XL = require('../xlsx.js');

// v0.14 Clinic Eligibility file: central Sohar packs one clinic's data for the Clinic Template Tool, which unpacks and
// derives the same rows, flags, benefits and tasks with this rules.js / core.js. Synthetic data only (made-up names,
// MRNs 9000xx, payer rules for made-up plans on real-looking names are not the real Payer Credential Rules table).

const HDR = ['Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Md Rc', 'Birth Dt', 'Sts', 'Guar Name', 'Guar Birth Dt', 'Guar Zip',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Sec Payer (Cht Guar)', 'Rendering'];
const row = (o) => HDR.map((k) => (o[k] == null ? '' : o[k]));
const MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
];
const LOCSTATE = { 'San Rafael': 'CA', 'Walnut Creek': 'CA' };
const CFG = { asOf: new Date(Date.UTC(2026, 9, 22)), scope: 'upcoming', locationState: LOCSTATE, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP };
const MD = 'Reyes, Sam MD', NP = 'Lark, Jordan PMHNP';
const ap = (per, mrn, name, dob, date, time, event, loc, payer, pol, extra) => row({ 'Per Nbr': per, 'Md Rc': mrn, 'Pat Name': name, 'Birth Dt': dob, 'Appt Dt': date,
  'Beg Tm': time, Event: event, 'Loc Name': loc, Sts: 'Expected', 'Payer Name (Cht Guar)': payer, 'Pol Nbr (Cht Guar)': pol, 'Guar Zip': '94901', Rendering: MD, ...(extra || {}) });
const IMPORT = () => [HDR,
  ap('501', '900001', 'Quillfeather, Testa', '4/5/1981', '10/27/2026', '9:00 AM', 'Telepsych', 'San Rafael', 'Cigna', 'C501'),
  ap('501', '900001', 'Quillfeather, Testa', '4/5/1981', '10/29/2026', '10:00 AM', 'Esketamine Treatment', 'San Rafael', 'Cigna', 'C501', { Rendering: NP }),
  ap('502', '900002', 'Brambleton, Bravoline', '6/7/1975', '10/26/2026', '10:30 AM', 'Telepsych', 'San Rafael', 'Blue Shield CA', 'BS502', { 'Sec Payer (Cht Guar)': 'Cigna' }),
  ap('503', '900003', 'Kiddington, Juniper', '8/9/2012', '10/28/2026', '1:00 PM', 'Telepsych', 'San Rafael', '', '', { 'Guar Name': 'Parentisson, Guardia', 'Guar Birth Dt': '3/3/1970' }),
  ap('504', '900004', 'Distantly, Otherine', '2/2/1960', '10/27/2026', '9:00 AM', 'Telepsych', 'Walnut Creek', 'Cigna', 'C504'),
  ap('505', '900005', 'Onenamesson', '1/1/1966', '10/30/2026', '2:00 PM', 'Telepsych', 'San Rafael', 'Cigna', 'C505'),
];
// Strings that must never reach the clinic file (patient names, DOBs, the guarantor, the other clinic, the unmatched row).
const NEVER = ['Quillfeather', 'Testa', 'Brambleton', 'Bravoline', 'Kiddington', 'Juniper', 'Parentisson', 'Guardia', 'Onenamesson',
  'Distantly', 'Otherine', 'Nomatchly', 'Persona', '04/05/1981', '4/5/1981', '06/07/1975', '08/09/2012', '03/03/1970', '01/01/1966', '94901', '900004', '999999'];

const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'memberId', 'payerName', 'status', 'statusReason', 'firstName', 'lastName',
  'planName', 'innCopay', 'innTotalDeductible', 'innRemainingDeductible', 'innOutOfPocket', 'innOutOfPocketRemaining'];
const ver = (o) => VER_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const DISC_HDR = ['requestId', 'id', 'created', 'resolved', 'firstName', 'lastName', 'dateOfBirth', 'memberId', 'responsePayerId', 'responsePayerName'];
const disc = (o) => DISC_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const URL = (id) => `https://app.soharhealth.com/verifications/${id}`;
const FILES = () => [
  { name: 'verification output.csv', rows: [VER_HDR, ...[
    { requestId: 'v-501', id: '900001', url: URL('v-501'), created: '2026-10-21', resolved: '2026-10-21T17:00:00Z', memberId: 'C501', payerName: 'Cigna', status: 'complete.eligible',
      firstName: 'Testa', lastName: 'Quillfeather', planName: 'SYNTH OPEN ACCESS', innCopay: '30', innTotalDeductible: '1000', innRemainingDeductible: '100', innOutOfPocket: '4000', innOutOfPocketRemaining: '2500' },
    { requestId: 'v-502', id: '900002', url: URL('v-502'), created: '2026-10-21', resolved: '2026-10-21T18:00:00Z', memberId: 'BS502', payerName: 'Blue Shield of California', status: 'error.member.id' },
    { requestId: 'v-504', id: '900004', url: URL('v-504'), created: '2026-10-21', resolved: '2026-10-21T18:00:00Z', memberId: 'C504', payerName: 'Cigna', status: 'complete.eligible', innCopay: '20' },
    { requestId: 'v-999', id: '999999', created: '2026-10-21', resolved: '2026-10-21T18:00:00Z', memberId: 'Z999', payerName: 'Cigna', status: 'complete.eligible', firstName: 'Persona', lastName: 'Nomatchly' },
  ].map(ver)] },
  { name: 'discovery output.csv', rows: [DISC_HDR, ...[
    { requestId: 'd-503', id: '900003', created: '2026-10-21', resolved: '2026-10-21T19:00:00Z' },
    { requestId: 'g-503', id: '900003-G', created: '2026-10-21', resolved: '2026-10-21T20:00:00Z', memberId: 'BSG503', responsePayerId: '94036', responsePayerName: 'Blue Shield of California' },
  ].map(disc)] },
];
// Made-up credential rules: Cigna lets no NP bill esketamine and is not Besse-approved (so 900001's ESK visit is flagged).
const RULES = PR.read([PR.HEADERS,
  ['Cigna', '', 'NP', 'N', 'No', 'No', 'No', 'No', 'cigna', ''],
  ['Cigna', '', 'MD', 'N', 'N/A', 'Yes', 'Yes', 'Yes', 'cigna', ''],
]).rules;
const TODAY = new Date(2026, 9, 22, 9, 0); // Thu 10/22/2026, 9:00 local
const LOC = 'San Rafael';
const blankRow = (cols, o) => Object.assign(Object.fromEntries(cols.map((c) => [c, ''])), o);

/** The admin side, as app.js builds it: patients after the skip window, merged results, build review, tasks with Done. */
function admin() {
  const built = C.buildPatients(IMPORT(), CFG);
  const merged = C.mergeSoharResults(IMPORT(), FILES(), CFG);
  const patients = C.applyReverifySkips(built.patients, merged.results, { ...C.REVERIFY_DEFAULTS, today: TODAY }).patients.filter((p) => p['In Scope'] === 'Yes');
  const fresh = C.buildTasks(merged.results);
  const fix = fresh.find((t) => t.MRN === '900002' && t.Task === 'Fix member ID in NextGen');
  const tasks = C.keepDone(fresh, [{ 'Task Key': fix['Task Key'], Done: 'x' }]);
  const src = {
    patients, results: merged.results, review: built.review, tasks,
    learningRows: ['Cigna', 'Blue Shield CA', 'Aetna Synthetic'].map((n) => blankRow(C.PAYER_LEARNING_COLS, { 'NextGen Payer': n, 'Patients agreed': '4' })),
    metricRows: ['Cigna', 'Aetna Synthetic'].map((n) => blankRow(C.PAYER_METRIC_COLS, { 'NextGen Payer': n, 'Unmapped patients': '0' })),
    cfg: { sfhpProviders: [], reverifyDays: 7, reverify: { newMonth: false }, approachingPct: 0.2, locationState: LOCSTATE, defaultState: '' },
    payerRuleRows: R.payerRuleRows(RULES), runAt: TODAY, version: '0.14.0',
  };
  return { built, merged, src, fixKey: fix['Task Key'] };
}
/** The admin's own derivation for the clinic, at `now`. */
function adminView(src, now) {
  const cfg = { today: now, sfhpProviders: [], reverifyDays: 7, reverifyNewMonth: false, payerRules: R.payerRulesFrom(src.payerRuleRows), locationState: LOCSTATE };
  const flags = R.allFlags(src.patients, src.results, cfg);
  const rows = R.clinicRows(LOC, src.patients, src.results, flags, now);
  const keys = new Set(rows.map((r) => r.key)), pers = new Set(rows.map((r) => r['Per Nbr']));
  return {
    rows, keys, flags: new Map([...flags].filter(([k]) => keys.has(k))),
    review: R.clinicReviewRows(LOC, src.patients, src.results, flags, now),
    benefits: C.benefitsView(src.results).rows.filter((b) => pers.has(b['Per Nbr'])),
    tasks: src.tasks.filter((t) => keys.has(t['Task Key'].split('|')[0])),
  };
}
async function roundTrip(pack) {
  const bytes = XL.buildWorkbook(pack);
  const wb = await XL.readWorkbook(bytes);
  return { bytes, wb, tables: Object.fromEntries(wb.sheets.map((s) => [s.name, s.rows])) };
}
// Compare without the stripped fields and the clinic-only extras; the guarantor's name is taken out of core's Action text.
const STRIPPED = ['Last Name', 'First Name', 'DOB'];
const EXTRA = ['LASFIR', 'Source', 'doneBy', 'clinic', 'clinicNote', 'sentNote'];
const norm = (list) => list.map((r) => Object.fromEntries(Object.keys(r).filter((k) => !EXTRA.includes(k))
  .map((k) => [k, STRIPPED.includes(k) ? '' : typeof r[k] === 'string' ? r[k].replace(' (Parentisson, Guardia)', '') : r[k]])));
const byTaskKey = (a, b) => a['Task Key'].localeCompare(b['Task Key']);
const stringsOf = (tables) => JSON.stringify(tables);

test('clinicPack: only the clinic\'s patients, sheets in order, hidden SE tables, CTT webExtension, no names or DOBs', async () => {
  const { src } = admin();
  const pack = R.clinicPack(LOC, src);
  assert.equal(pack.name, `Clinic Eligibility - San Rafael ${C.localDate(TODAY)}.xlsx`);
  assert.deepEqual(pack.webExtension, { id: 'cf213823-6d83-43e7-b7c9-37666b4d7c3d', version: '1.0.0.1' });
  const { bytes, wb, tables } = await roundTrip(pack);
  assert.deepEqual(wb.sheets.map((s) => [s.name, s.hidden]), [['Read Me', false], ['Clinic Work', false], ['SE Info', true], ['SE Patients', true],
    ['SE Results', true], ['SE Review', true], ['SE Payers', true], ['SE Done', true], ['Payer Credential Rules', true]]);
  assert.deepEqual(Object.values(R.SE_SHEETS), wb.sheets.map((s) => s.name));
  assert.deepEqual(tables['Clinic Work'][0], R.CLINIC_WORK_COLS, 'Clinic Work: header only');
  assert.equal(tables['Clinic Work'].length, 1);
  assert.deepEqual(pack.sheets.find((s) => s.name === 'Clinic Work').textColumns, R.CLINIC_WORK_COLS, 'typed keys keep leading zeros');
  const raw = Buffer.from(bytes).toString('latin1');
  assert.ok(raw.includes('id="cf213823-6d83-43e7-b7c9-37666b4d7c3d" version="1.0.0.1"'), 'CTT add-in reference');
  assert.ok(!raw.includes('tblClinicMeta'), 'never Sohar\'s own clinic-file meta table');
  for (const t of ['tblSERead', 'tblClinicWork', 'tblSEInfo', 'tblSEPatients', 'tblSEResults', 'tblSEReview', 'tblSEPayers', 'tblSEDone', 'tblPayerCredRules']) assert.ok(raw.includes(`name="${t}"`), t);

  const keysOf = (sheet) => tables[sheet].slice(1).map((r) => r[0]);
  assert.deepEqual([...new Set(keysOf('SE Patients'))].sort(), ['900001', '900002', '900003', '900005']);
  assert.ok(keysOf('SE Results').every((k) => ['900001', '900002', '900003', '900005'].includes(k)), 'results: the clinic\'s patients only (no other clinic, no unmatched)');
  const text = stringsOf(tables);
  for (const s of NEVER) assert.ok(!text.toLowerCase().includes(s.toLowerCase()), `"${s}" is not in the file`);
  const keys = new Set();
  const walk = (v) => { if (Array.isArray(v)) v.forEach(walk); else if (v && typeof v === 'object') for (const [k, e] of Object.entries(v)) { keys.add(k); walk(e); } };
  for (const sheet of ['SE Patients', 'SE Results', 'SE Review', 'SE Payers']) tables[sheet].slice(1).forEach((r) => walk(JSON.parse(r[1])));
  for (const k of ['Last Name', 'First Name', 'DOB', 'Zip', 'Subscriber First', 'Subscriber Last', 'Subscriber DOB', 'guarantor', 'guarZip']) assert.ok(!keys.has(k), `no ${k} key`);
  assert.ok(keys.has('LASFIR') && keys.has('appointments') && keys.has('verifyRows') && keys.has('resultKey') && keys.has('key'),'what the derivation needs is there');
  const pts = tables['SE Patients'].slice(1).map((r) => JSON.parse(r[1]));
  assert.deepEqual(pts.map((p) => [p.MRN, p.LASFIR]).sort(), [['900001', 'QUITES'], ['900002', 'BRABRA'], ['900003', 'KIDJUN'], ['900005', 'ONE']]);
  assert.ok(tables['SE Results'].slice(1).every((r) => /^[A-Z]{3,6}$/.test(JSON.parse(r[1]).LASFIR)), 'LASFIR on every result');
  const g = tables['SE Results'].slice(1).map((r) => JSON.parse(r[1])).find((r) => r['Sohar Ref'] === 'g-503');
  assert.match(g.Action, /^Found through the guarantor: the patient is likely a dependent/, 'the guarantor\'s name is taken out of Action');
  assert.equal(g.key, '900003');
  const rev = tables['SE Review'].slice(1).map((r) => JSON.parse(r[1]));
  const nameItem = rev.find((r) => r.MRN === '900005' && /^Name /.test(r.Issue));
  assert.equal(nameItem.Detail, 'Not shown in the clinic file (patient name).');
  assert.deepEqual(tables['SE Payers'].slice(1).map((r) => r[0]).sort(), ['blue shield ca', 'cigna'], 'payers the clinic carries only');
  const cigna = JSON.parse(tables['SE Payers'].slice(1).find((r) => r[0] === 'cigna')[1]);
  assert.equal(cigna.learning['Patients agreed'], '4');
  assert.equal(cigna.metrics['Unmapped patients'], '0');
  assert.equal(JSON.parse(tables['SE Payers'].slice(1).find((r) => r[0] === 'blue shield ca')[1]).metrics, null);
  assert.deepEqual(tables['SE Done'].slice(1), [[admin().fixKey, 'x']]);
  const info = Object.fromEntries(tables['SE Info'].slice(1));
  assert.deepEqual([info.Kind, info.Location, info['Run Date'], info['Sohar Version'], info.Format], ['Clinic Eligibility', 'San Rafael', '10/22/2026', '0.14.0', '1']);
  assert.equal(info['Run At'], TODAY.toISOString());
  assert.deepEqual(JSON.parse(info.Settings), { locationState: LOCSTATE }, 'only settings that differ from the defaults');
  assert.ok(tables['Read Me'].slice(1).some((r) => /Home > Clinic Template Tool/.test(r[0])));
  assert.ok(tables['Read Me'].slice(1).some((r) => /OneDrive or Teams/.test(r[0])));
  assert.ok(tables['Read Me'].slice(1).some((r) => /your clinic's folder in the Clinic Eligibility library/.test(r[0]) && /do not need to send the file back/.test(r[0])));
  assert.ok(tables['Read Me'].slice(1).some((r) => /^If the file is not in the library.*send it back/.test(r[0])), 'sending it back stays as the fallback');
  assert.throws(() => R.clinicPack('All', src), /one location/);
  // No payer rules: no rules sheet.
  const { wb: wb2 } = await roundTrip(R.clinicPack(LOC, { ...src, payerRuleRows: [] }));
  assert.ok(!wb2.sheets.some((s) => s.name === 'Payer Credential Rules'));
});

test('clinicDerive: rows, flags, review, benefits and tasks equal the admin derivation for the same patients, minus names and DOB', async () => {
  const { src } = admin();
  const { tables } = await roundTrip(R.clinicPack(LOC, src));
  const se = R.clinicUnpack(tables);
  assert.equal(se.meta.Location, 'San Rafael');
  assert.equal(se.settings.reverifyDays, 7);
  assert.deepEqual(se.settings.locationState, LOCSTATE);
  assert.equal(se.payerRules.length, RULES.length, 'payer rules come back from the hidden sheet');
  for (const now of [TODAY, new Date(2026, 9, 27, 12, 0)]) {
    const a = adminView(src, now);
    const d = R.clinicDerive(se, [], now);
    if (now === TODAY) assert.ok(a.rows.length === 4 && a.tasks.length === 5 && a.benefits.length === 2 && a.review.length >= 5, 'the comparison is not vacuous');    assert.deepEqual(d.rows.map((r) => r.key), a.rows.map((r) => r.key), 'same patients, same order');
    assert.deepEqual(norm(d.rows), norm(a.rows));
    assert.deepEqual(d.rows.map((r) => r.LASFIR), a.rows.map((r) => C.lasfir(r['Last Name'], r['First Name'])));
    assert.ok(d.rows.every((r) => r['Last Name'] === '' && r['First Name'] === '' && r.DOB === ''));
    assert.deepEqual(d.flags, a.flags);
    assert.deepEqual(norm(d.review.filter((r) => r.Source === 'Flag')), norm(a.review));
    assert.deepEqual(norm(d.benefits.rows), norm(a.benefits));
    assert.deepEqual(norm(d.tasks).sort(byTaskKey), norm(a.tasks).sort(byTaskKey));
  }
  const d = R.clinicDerive(se, [], TODAY);
  const codes = (k) => (d.flags.get(k) || []).map((f) => f.code);
  assert.ok(codes('900001').includes('ESK_CRED') && codes('900001').includes('BESSE_NO'), 'ESK flags from the packed rules');
  assert.ok(codes('900002').includes('SECONDARY'));
  assert.ok(codes('900003').includes('GUARANTOR'));
  const r1 = d.rows.find((r) => r.key === '900001');
  assert.equal(r1['Verification Skipped'], 'Verified eligible 10/21/2026', 'skip window travels on x');
  assert.equal(r1['Deductible Status'], 'Approaching');
  assert.ok(d.review.some((r) => r.Source === 'Build' && r.MRN === '900005' && r.LASFIR === 'ONE' && /^Name /.test(r.Issue)), 'build review rows for clinic MRNs');
  const t = (mrn, task) => d.tasks.find((x) => x.MRN === mrn && x.Task === task);
  assert.equal(t('900002', 'Fix member ID in NextGen').Done, 'x', 'Done from SE Done');
  assert.equal(t('900002', 'Fix member ID in NextGen').doneBy, 'admin');
  assert.equal(t('900001', 'Update benefits in NextGen').clinic, true);
  assert.equal(t('900005', 'Resend to Sohar').clinic, false);
  assert.ok(d.tasks.every((x) => /^[A-Z]{3,6}$/.test(x.LASFIR)));
  assert.deepEqual(d.payers.map((p) => [p.payer, p.patients, !!p.learning]), [['Cigna', 3, true], ['Blue Shield CA', 1, true]]);
  assert.equal(d.counts.patients, 4);
  assert.equal(d.counts.openTasks, d.tasks.filter((x) => !x.Done).length);
  // Object rows (as CTT reads tables) unpack the same as the xlsx grid.
  const objs = Object.fromEntries(Object.entries(tables).map(([k, v]) => [k, R.tableObjects(v)]));
  assert.deepEqual(R.clinicUnpack(objs), se);
});

test('clinicDerive: Clinic Work overrides Re-verify anyway, task Done and notes; workRows upserts by Kind + Key', async () => {
  const { src, fixKey } = admin();
  const se = R.clinicUnpack((await roundTrip(R.clinicPack(LOC, src))).tables);
  const benefitsKey = R.clinicDerive(se, [], TODAY).tasks.find((x) => x.MRN === '900001' && x.Task === 'Update benefits in NextGen')['Task Key'];
  const t1 = new Date(Date.UTC(2026, 9, 22, 17)), t2 = new Date(Date.UTC(2026, 9, 22, 18));
  let work = R.workRows([], { Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'On', By: 'Front desk' }, t1);
  assert.deepEqual(work, [{ Kind: 'Re-verify anyway', Key: '900001', MRN: '900001', LASFIR: 'QUITES', Value: 'On', Note: '', By: 'Front desk', When: t1.toISOString() }]);
  assert.equal(R.clinicDerive(se, work, TODAY).rows.find((r) => r.key === '900001')['Re-verify Anyway'], 'Yes');
  work = R.workRows(work, [
    { Kind: 're-verify ANYWAY', Key: '900001', Value: 'Off', By: 'Front desk' },
    { Kind: 'Task done', Key: benefitsKey, MRN: '900001', Value: 'Yes', By: 'Front desk' },
    { Kind: 'Task done', Key: fixKey, MRN: '900002', Value: 'No', By: 'Front desk' },
    { Kind: 'Note', Key: '900002', MRN: '900002', Note: 'Card on file is the new one', By: 'Front desk' },
    { Kind: 'Payer note', Key: 'Cigna', Value: 'Cigna Open Access Plus', Note: 'New cards since 10/1', By: 'Front desk' },
  ], t2);
  assert.equal(work.filter((r) => r.Kind === 'Re-verify anyway').length, 1, 'upsert: one row per Kind + Key');
  assert.deepEqual(work.map((r) => r.Kind), ['Re-verify anyway', 'Task done', 'Task done', 'Note', 'Payer note']);
  assert.ok(work.every((r) => r.When === t2.toISOString()));
  assert.deepEqual(Object.keys(work[0]), R.CLINIC_WORK_COLS);
  assert.throws(() => R.workRows(work, { Kind: 'Delete patient', Key: '900001' }), /not one of/);
  const d = R.clinicDerive(se, work, TODAY);
  assert.equal(d.rows.find((r) => r.key === '900001')['Re-verify Anyway'], '', 'Off wins');
  const ben = d.tasks.find((x) => x['Task Key'] === benefitsKey);
  assert.deepEqual([ben.Done, ben.doneBy], ['Yes', 'clinic']);
  const fix = d.tasks.find((x) => x['Task Key'] === fixKey);
  assert.deepEqual([fix.Done, fix.doneBy], ['x', 'admin'], 'clinic No does not undo the admin\'s Done');
  assert.deepEqual(d.notes.get('900002'), { Note: 'Card on file is the new one', By: 'Front desk', When: t2.toISOString() });
  const cigna = d.payers.find((p) => p.key === 'cigna');
  assert.deepEqual(cigna.note, { cardSays: 'Cigna Open Access Plus', note: 'New cards since 10/1', by: 'Front desk', when: t2.toISOString() });
  assert.equal(d.counts.notes, 1);
  assert.equal(d.counts.payerNotes, 1);
  // Last row per Kind + Key wins when a file holds duplicates (an older writer, or a hand edit).
  const dup = work.concat([{ Kind: 'Task done', Key: benefitsKey, Value: 'No', When: t2.toISOString() }]);
  assert.equal(R.clinicDerive(se, dup, TODAY).tasks.find((x) => x['Task Key'] === benefitsKey).Done, '');
});

test('clinicDerive: each task carries its Clinic Work done note; a note saved unticked leaves Done as the file has it', async () => {
  const { src, fixKey } = admin();
  const se = R.clinicUnpack((await roundTrip(R.clinicPack(LOC, src))).tables);
  const ben = R.clinicDerive(se, [], TODAY).tasks.find((x) => x.MRN === '900001' && x.Task === 'Update benefits in NextGen')['Task Key'];
  assert.ok(R.clinicDerive(se, [], TODAY).tasks.every((t) => t.clinicNote === ''));
  const at = new Date(Date.UTC(2026, 9, 22, 17));
  let work = R.workRows([], [
    { Kind: 'Task done', Key: ben, MRN: '900001', Value: 'Yes', Note: 'called patient, new card scanned', By: 'Front desk' },
    { Kind: 'Task done', Key: fixKey, MRN: '900002', Note: 'asked for a photo of the card', By: 'Front desk' },
  ], at);
  let d = R.clinicDerive(se, work, TODAY);
  const t = (k) => d.tasks.find((x) => x['Task Key'] === k);
  assert.deepEqual([t(ben).Done, t(ben).doneBy, t(ben).clinicNote], ['Yes', 'clinic', 'called patient, new card scanned']);
  assert.deepEqual([t(fixKey).Done, t(fixKey).doneBy, t(fixKey).clinicNote], ['x', 'admin', 'asked for a photo of the card']);
  assert.equal(d.counts.taskNotes, 2);
  // Unticking keeps the note.
  work = R.workRows(work, { Kind: 'Task done', Key: ben, MRN: '900001', Value: 'No', By: 'Front desk' }, at);
  d = R.clinicDerive(se, work, TODAY);
  assert.deepEqual([t(ben).Done, t(ben).clinicNote], ['', 'called patient, new card scanned']);
});

test('clinicPack: SE Done carries the admin\'s Clinic Note back to the clinic (scrubbed); clinicDerive shows it as sentNote', async () => {
  const { src, fixKey } = admin();
  const ben = src.tasks.find((x) => x.MRN === '900001' && x.Task === 'Update benefits in NextGen')['Task Key'];
  const notes = { [fixKey]: 'asked for a photo of the card (Front desk, San Rafael, 10/21)',
    [ben]: 'called Testa Quillfeather, new card scanned (Front desk, San Rafael, 10/21)' };
  const tasks = src.tasks.map((t) => (notes[t['Task Key']] ? { ...t, 'Clinic Note': notes[t['Task Key']] } : t))
    .concat([{ ...src.tasks[0], 'Task Key': '900004|Update benefits in NextGen', MRN: '900004', Done: '', 'Clinic Note': 'Distantly called back' }]);
  const pack = R.clinicPack(LOC, { ...src, tasks });
  const { tables } = await roundTrip(pack);
  assert.deepEqual(tables['SE Done'][0], ['Task Key', 'Done', 'Clinic Note']);
  const rows = R.tableObjects(tables['SE Done']).sort(byTaskKey);
  assert.deepEqual(rows.map((r) => [r['Task Key'], r.Done, r['Clinic Note']]), [
    [ben, '', 'called QUITES, new card scanned (Front desk, San Rafael, 10/21)'],
    [fixKey, 'x', 'asked for a photo of the card (Front desk, San Rafael, 10/21)'],
  ], 'a note alone is enough; the other clinic\'s task stays out; the patient\'s name is scrubbed');
  assert.equal(pack.counts.done, 2);
  const text = stringsOf(tables);
  for (const s of NEVER) assert.ok(!text.toLowerCase().includes(s.toLowerCase()), `"${s}" is not in the file`);
  const se = R.clinicUnpack(tables);
  assert.deepEqual(se.done.find((d) => d['Task Key'] === ben), { 'Task Key': ben, 'Done': '', 'Clinic Note': rows[0]['Clinic Note'] });
  let d = R.clinicDerive(se, [], TODAY);
  const t = (k) => d.tasks.find((x) => x['Task Key'] === k);
  assert.deepEqual([t(ben).Done, t(ben).doneBy, t(ben).sentNote, t(ben).clinicNote], ['', '', rows[0]['Clinic Note'], ''], 'a note alone does not tick the task');
  assert.deepEqual([t(fixKey).Done, t(fixKey).sentNote], ['x', notes[fixKey]]);
  assert.equal(t(fixKey)['Clinic Note'], notes[fixKey], 'the Clinic Note column as the admin has it');
  assert.equal(d.tasks.filter((x) => x.sentNote).length, 2);
  assert.equal(d.counts.sentNotes, 2);
  // The clinic's new note sits beside the admin's copy; neither replaces the other.
  const work = R.workRows([], { Kind: 'Task done', Key: ben, MRN: '900001', Note: 'card photo emailed too', By: 'Front desk' }, new Date(Date.UTC(2026, 9, 22, 17)));
  d = R.clinicDerive(se, work, TODAY);
  assert.deepEqual([t(ben).sentNote, t(ben).clinicNote], [rows[0]['Clinic Note'], 'card photo emailed too']);
  // A file made before SE Done had Clinic Note (two columns, still Format 1) reads with blank notes.
  const old = { ...tables, 'SE Done': [['Task Key', 'Done'], [fixKey, 'x']] };
  const se0 = R.clinicUnpack(old);
  assert.deepEqual(se0.done, [{ 'Task Key': fixKey, 'Done': 'x', 'Clinic Note': '' }]);
  const d0 = R.clinicDerive(se0, [], TODAY);
  assert.ok(d0.tasks.every((x) => x.sentNote === '' && x['Clinic Note'] === ''));
  assert.equal(d0.tasks.find((x) => x['Task Key'] === fixKey).Done, 'x');
});

test('clinicPack: a Clinic Note typed by hand loses a first or last name alone and the patient\'s or guarantor\'s DOB', async () => {
  const { src } = admin();
  const ben = src.tasks.find((x) => x.MRN === '900001' && x.Task === 'Update benefits in NextGen')['Task Key'];
  const kid = src.tasks.find((x) => x.MRN === '900003')['Task Key'];
  const notes = {
    [ben]: 'Spoke w/ Testa, DOB 4/5/1981 confirmed; QUILLFEATHER ok; born 04-05-81, Apr 5, 1981, 5 April 1981, 1981-04-05; visit 4/6/1981 stays',
    [kid]: 'Guardia called (DOB 03/03/1970), Parentisson to bring the card; Juniper seen 3/3/1971',
  };
  const tasks = src.tasks.map((t) => (notes[t['Task Key']] ? { ...t, 'Clinic Note': notes[t['Task Key']] } : t));
  const { tables } = await roundTrip(R.clinicPack(LOC, { ...src, tasks }));
  const got = Object.fromEntries(R.tableObjects(tables['SE Done']).map((r) => [r['Task Key'], r['Clinic Note']]));
  assert.equal(got[ben], 'Spoke w/ QUITES, DOB [date removed] confirmed; QUITES ok; born [date removed], [date removed], [date removed], [date removed]; visit 4/6/1981 stays');
  assert.equal(got[kid], 'the guarantor called (DOB [date removed]), the guarantor to bring the card; KIDJUN seen 3/3/1971');
  const text = stringsOf(tables);
  for (const s of NEVER.concat(['04-05-81', 'Apr 5', '5 April', '1981-04-05'])) assert.ok(!text.toLowerCase().includes(s.toLowerCase()), `"${s}" is not in the file`);
  const d = R.clinicDerive(R.clinicUnpack(tables), [], TODAY);
  assert.equal(d.tasks.find((x) => x['Task Key'] === ben).sentNote, got[ben], 'CTT shows the scrubbed note');
  // Whole words only: a longer word holding the name stays.
  const t2 = src.tasks.map((t) => (t['Task Key'] === ben ? { ...t, 'Clinic Note': 'Testament on file' } : t));
  const two = await roundTrip(R.clinicPack(LOC, { ...src, tasks: t2 }));
  assert.equal(R.tableObjects(two.tables['SE Done']).find((r) => r['Task Key'] === ben)['Clinic Note'], 'Testament on file');
});

test('app.js cttWorkbook: SE Done takes Done and Clinic Note from the NextGen Tasks sheet as it is now, not the upload\'s copy', async () => {
  const { src, fixKey } = admin();
  const app = readFileSync(new globalThis.URL('../app.js', import.meta.url), 'utf8'); // URL here is the Sohar link helper
  const cut = (from, to) => { const i = app.indexOf(from), j = app.indexOf(to, i); assert.ok(i >= 0 && j > i, `app.js has ${from}`); return app.slice(i, j); };
  const body = cut('  async function cttWorkbook(', '  async function saveCttFile(') + cut('  async function currentTasks(', '  async function copyAdminMessage(');
  const ben = src.tasks.find((x) => x.MRN === '900001' && x.Task === 'Update benefits in NextGen')['Task Key'];
  // The upload's copy has no notes; since then a bring-back wrote one and the admin typed one on the sheet.
  const sheet = src.tasks.map((t) => (t['Task Key'] === ben ? { ...t, 'Clinic Note': 'called patient (Front desk, San Rafael, 10/22)' }
    : t['Task Key'] === fixKey ? { ...t, Done: 'x', 'Clinic Note': 'billing has the new card' } : t));
  const state = { tasks: src.tasks.map((t) => ({ ...t, 'Clinic Note': '' })), demo: false, results: src.results, review: src.review, learning: [], metrics: null,
    cfg: src.cfg, payerRules: R.payerRulesFrom(src.payerRuleRows) };
  const X = { async readTasks() { return sheet; } };
  // cttWorkbook runs at new Date(): pin it to the fixture's day, or the 10/26 visit drops out once that date passes.
  class FixedDate extends Date { constructor(...a) { super(...(a.length ? a : [TODAY.getTime()])); } static now() { return TODAY.getTime(); } }
  const make = new Function('R', 'C', 'XL', 'X', 'state', 'inScopePatients', 'VERSION', 'Date', `${body}\nreturn { cttWorkbook };`);
  const { cttWorkbook } = make(R, C, XL, X, state, () => src.patients, '0.14.0', FixedDate);
  const f = await cttWorkbook(LOC);
  const wb = await XL.readWorkbook(f.bytes);
  const done = R.tableObjects(wb.sheets.find((s) => s.name === 'SE Done').rows);
  assert.deepEqual(done.map((r) => [r['Task Key'], r.Done, r['Clinic Note']]).sort((a, b) => a[0].localeCompare(b[0])),
    [[ben, '', 'called patient (Front desk, San Rafael, 10/22)'], [fixKey, 'x', 'billing has the new card']]);
});

test('mergeClinicWork: requests, Done (No only clears clinic-set Done), notes and payer notes; last row per Kind + Key wins', () => {
  const W = (Kind, Key, Value, When, o) => ({ Kind, Key, MRN: '', LASFIR: '', Value, Note: '', By: 'Front desk', When, ...(o || {}) });
  const work = [
    W('Re-verify anyway', '900001', 'On', '2026-10-22T17:00:00Z'),
    W('Re-verify anyway', '900001', 'Off', '2026-10-22T17:01:00Z'),
    W('Re-verify anyway', '900001', 'On', '2026-10-22T17:02:00Z'),
    W('Re-verify anyway', '900005', 'Off', '2026-10-22T17:03:00Z'),
    W('Task done', 'A|Update benefits in NextGen', 'Yes', '2026-10-22T17:04:00Z'),
    W('Task done', 'B|Fix name in NextGen', 'Yes', '2026-10-22T17:05:00Z'),
    W('Task done', 'B|Fix name in NextGen', 'No', '2026-10-22T17:06:00Z'),
    W('Task done', 'C|Fix member ID in NextGen', 'No', '2026-10-22T17:07:00Z'),
    W('Task done', 'D|Add insurance in NextGen', 'No', '2026-10-22T17:08:00Z'),
    W('Task done', 'E|Fix member ID in NextGen', 'Yes', '2026-10-22T17:08:10Z'),
    W('Task done', 'E|Fix member ID in NextGen', 'No', '2026-10-22T17:08:20Z'),
    W('Task done', 'GONE|Resend to Sohar', 'Yes', '2026-10-22T17:09:00Z'),
    W('Note', '900002', '', '2026-10-22T17:10:00Z', { Note: 'Older than the admin copy' }),
    W('Note', '900003', '', '2026-10-22T17:11:00Z', { Note: 'Mom brings the card Monday' }),
    W('Payer note', 'Cigna', 'Cigna Open Access Plus', '2026-10-22T17:12:00Z', { Note: 'New cards' }),
    W('Unknown kind', 'x', 'Yes', '2026-10-22T17:13:00Z'),
  ];
  const ctx = {
    location: 'San Rafael',
    requests: [{ MRN: '900005', Requested: '2026-10-21T10:00:00Z', By: 'Clinic tab' }],
    // B: the clinic's Done from an earlier file. C, D: the admin's own Done. E: the admin marked it Done (pane, "Yes") after the export.
    tasks: [{ 'Task Key': 'A|Update benefits in NextGen', Done: '' }, { 'Task Key': 'B|Fix name in NextGen', Done: 'Yes (clinic San Rafael)' },
      { 'Task Key': 'C|Fix member ID in NextGen', Done: 'x' }, { 'Task Key': 'D|Add insurance in NextGen', Done: 'Yes' },
      { 'Task Key': 'E|Fix member ID in NextGen', Done: 'Yes' }].map((t) => ({ ...t, 'Next Location': 'san rafael' })), // any case
    notes: [{ MRN: '900002', Location: 'San Rafael', Note: 'Admin copy', By: 'Clinic file San Rafael', When: '2026-10-23T08:00:00Z' }],
    payerNotes: [], now: TODAY,
  };
  const m = R.mergeClinicWork(work, ctx);
  assert.deepEqual(m.requests, [{ MRN: '900001', Requested: '2026-10-22T17:02:00Z', By: 'Clinic file San Rafael' }], '900001 added (last On), 900005 removed');
  assert.deepEqual([...m.done], [['A|Update benefits in NextGen', 'Yes (clinic San Rafael)'], ['B|Fix name in NextGen', '']], 'C, D and E (the admin\'s Done) stay');
  assert.deepEqual(m.notes, [
    { MRN: '900002', Location: 'San Rafael', Note: 'Admin copy', By: 'Clinic file San Rafael', When: '2026-10-23T08:00:00Z' },
    { MRN: '900003', Location: 'San Rafael', Note: 'Mom brings the card Monday', By: 'Front desk', When: '2026-10-22T17:11:00Z' },
  ], 'newest per MRN + Location wins');
  assert.deepEqual(m.payerNotes, [{ 'NextGen Payer': 'Cigna', Location: 'San Rafael', 'Card Says': 'Cigna Open Access Plus', Note: 'New cards', By: 'Front desk', When: '2026-10-22T17:12:00Z' }]);
  assert.deepEqual(Object.keys(m.notes[0]), R.CLINIC_NOTE_COLS);
  assert.deepEqual(Object.keys(m.payerNotes[0]), R.CLINIC_PAYER_NOTE_COLS);
  assert.deepEqual(m.counts, { reverifyAdded: 1, reverifyRemoved: 1, doneSet: 1, doneCleared: 1, doneNotes: 0, tasksChanged: 2, notes: 1, payerNotes: 1, skipped: 1, otherClinic: 0, old: 0, applied: 11 });
  assert.deepEqual(m.applied.find((r) => r.Kind === 'Payer note'), { Location: 'San Rafael', Kind: 'Payer note', Key: 'Cigna', When: '2026-10-22T17:12:00Z', Value: '' });
  assert.deepEqual(m.applied.filter((r) => r.Kind === 'Task done').map((r) => r.Value), ['Yes', 'No', 'No', 'No', 'No', 'Yes'], 'the tick acted on');
  assert.deepEqual(Object.keys(m.applied[0]), R.CLINIC_APPLIED_COLS);
  // The sheet as XL.readWorkbook gives it (a grid, header first) merges the same; a newer blank note removes the row.
  const grid = [R.CLINIC_WORK_COLS, ...work.map((r) => R.CLINIC_WORK_COLS.map((c) => r[c]))];
  assert.deepEqual(R.mergeClinicWork(grid, ctx), m);
  const cleared = R.mergeClinicWork([W('Note', '900002', '', '2026-10-24T08:00:00Z')], ctx);
  assert.deepEqual(cleared.notes, []);
});

test('mergeClinicWork: the same file sent again changes nothing; only newer rows apply', () => {
  const W = (Kind, Key, Value, When) => ({ Kind, Key, MRN: '', LASFIR: '', Value, Note: '', By: 'Front desk', When });
  const KEY = '900001|Fix member ID in NextGen';
  const fileA = [W('Task done', KEY, 'Yes', '2026-10-20T17:00:00Z'), W('Re-verify anyway', '900001', 'On', '2026-10-20T17:01:00Z')];
  const ctx = { location: 'San Rafael', requests: [], tasks: [{ 'Task Key': KEY, Done: '', 'Next Location': 'San Rafael' }], notes: [], payerNotes: [], applied: [], now: TODAY };
  const first = R.mergeClinicWork(fileA, ctx);
  assert.deepEqual([...first.done], [[KEY, 'Yes (clinic San Rafael)']]);
  assert.equal(first.requests.length, 1);
  // The admin finds the ID still wrong and clears Done, and cancels the request. The clinic sends file A again with a new note.
  const again = R.mergeClinicWork(fileA.concat([{ ...W('Note', '900001', '', '2026-10-21T09:00:00Z'), Note: 'Card Monday' }]),
    { ...ctx, applied: first.applied });
  assert.deepEqual([...again.done], [], 'the old tick is not applied again');
  assert.deepEqual(again.requests, [], 'the cancelled request is not added again');
  assert.deepEqual(again.notes.map((r) => r.Note), ['Card Monday'], 'the new row applies');
  assert.deepEqual([again.counts.old, again.counts.applied, again.counts.doneSet, again.counts.reverifyAdded], [2, 1, 0, 0]);
  // A newer row that still says Yes (the clinic only edited the done note) leaves the admin's reopened task open; the note applies.
  const noteEdit = R.mergeClinicWork([{ ...W('Task done', KEY, 'Yes', '2026-10-22T09:00:00Z'), Note: 'new card scanned' }], { ...ctx, applied: again.applied });
  assert.deepEqual([[...noteEdit.done], noteEdit.doneNotes.size, noteEdit.counts.applied], [[], 1, 1]);
  // The clinic unticks and ticks again: each new tick applies; another location's ledger does not block this one.
  const off = R.mergeClinicWork([W('Task done', KEY, 'No', '2026-10-22T09:30:00Z')], { ...ctx, applied: noteEdit.applied });
  const later = R.mergeClinicWork([W('Task done', KEY, 'Yes', '2026-10-22T10:00:00Z')], { ...ctx, applied: off.applied });
  assert.deepEqual([...later.done], [[KEY, 'Yes (clinic San Rafael)']]);
  const elsewhere = R.mergeClinicWork(fileA, { ...ctx, location: 'Walnut Creek', applied: first.applied });
  assert.equal(elsewhere.counts.old, 0);
  // Ledger rows older than 120 days drop out.
  const aged = R.mergeClinicWork([], { ...ctx, applied: first.applied, now: new Date(2027, 2, 1) });
  assert.deepEqual(aged.applied, []);
});

test('clinicDerive: a clinic Done brought back earlier comes in SE Done as the clinic\'s, so the clinic can still undo it', async () => {
  const { src } = admin();
  const fresh = R.clinicDerive(R.clinicUnpack((await roundTrip(R.clinicPack(LOC, src))).tables), [], TODAY);
  const key = fresh.tasks.find((x) => x.MRN === '900001' && x.Task === 'Update benefits in NextGen')['Task Key'];
  const tasks = src.tasks.map((t) => (t['Task Key'] === key ? { ...t, Done: 'Yes (clinic San Rafael)' } : t));
  const se = R.clinicUnpack((await roundTrip(R.clinicPack(LOC, { ...src, tasks }))).tables);
  const t = (work) => R.clinicDerive(se, work, TODAY).tasks.find((x) => x['Task Key'] === key);
  assert.deepEqual([t([]).Done, t([]).doneBy], ['Yes (clinic San Rafael)', 'clinic']);
  const off = R.workRows([], { Kind: 'Task done', Key: key, MRN: '900001', Value: 'No', By: 'Front desk' }, new Date(Date.UTC(2026, 9, 23)));
  assert.deepEqual([t(off).Done, t(off).doneBy], ['', '']);
  // Back at the admin: that No clears the clinic's mark.
  const m = R.mergeClinicWork(off, { location: LOC, requests: [], tasks, notes: [], payerNotes: [], now: TODAY });
  assert.deepEqual([...m.done], [[key, '']]);
});

test('clinicPack: the name scrub takes whole words in free text only, so provider, location and payer stay as the admin has them', async () => {
  // Synthetic: patient "Le, Ed" whose provider "Cole, Edward MD" holds the name as a substring; a one-word guarantor name.
  const imp = [HDR,
    ap('701', '900701', 'Le, Ed', '5/5/1985', '10/27/2026', '9:00 AM', 'Telepsych', LOC, 'SFHP Medi-Cal', 'S701', { Rendering: 'Cole, Edward MD' }),
    ap('702', '900702', 'Kidsworth, Robin', '2/2/2014', '10/28/2026', '9:00 AM', 'Telepsych', LOC, '', '', { 'Guar Name': 'Parentisson', 'Guar Birth Dt': '3/3/1970' }),
  ];
  const files = [
    { name: 'discovery output.csv', rows: [DISC_HDR, disc({ requestId: 'd-702', id: '900702', created: '2026-10-21', resolved: '2026-10-21T19:00:00Z' }),
      disc({ requestId: 'g-702', id: '900702-G', created: '2026-10-21', resolved: '2026-10-21T20:00:00Z', memberId: 'BSG702', responsePayerId: '94036', responsePayerName: 'Blue Shield of California' })] },
  ];
  const built = C.buildPatients(imp, CFG);
  assert.equal(built.patients.find((p) => p.MRN === '900702').x.guarantor.firstName, '', 'the fixture has a one-word guarantor');
  const merged = C.mergeSoharResults(imp, files, CFG);
  const g = merged.results.find((r) => r['Sohar Ref'] === 'g-702');
  assert.match(g.Action, /^Found through the guarantor \(Parentisson, \)/, 'core quotes the one-word guarantor');
  const patients = C.applyReverifySkips(built.patients, merged.results, { ...C.REVERIFY_DEFAULTS, today: TODAY }).patients.filter((p) => p['In Scope'] === 'Yes');
  const src = { patients, results: merged.results, review: built.review, tasks: C.buildTasks(merged.results), cfg: { sfhpProviders: ['Cole'] }, runAt: TODAY };
  const { tables } = await roundTrip(R.clinicPack(LOC, src));
  const text = stringsOf(tables);
  assert.ok(!text.includes('Parentisson'), 'no one-word guarantor name');
  assert.ok(text.includes('Cole, Edward MD') && !text.includes('CoLEED'), 'the provider is not rewritten');
  const se = R.clinicUnpack(tables);
  assert.equal(se.results.find((r) => r['Sohar Ref'] === 'g-702').Action.startsWith('Found through the guarantor: '), true);
  const d = R.clinicDerive(se, [], TODAY);
  const cfg = { today: TODAY, sfhpProviders: ['Cole'], reverifyDays: 7, reverifyNewMonth: false, payerRules: [], locationState: {} };
  const flags = R.allFlags(patients, merged.results, cfg);
  const rows = R.clinicRows(LOC, patients, merged.results, flags, TODAY);
  const sig = (rs) => rs.map((r) => [r.key, r.Pills, r['Next Provider'], r.Location]);
  assert.equal(rows.length, 2);
  assert.ok(!/SFHP/.test(rows.find((r) => r.key === '900701').Pills), 'the admin sees Cole as an SFHP provider');
  assert.deepEqual(sig(d.rows), sig(rows), 'pills, provider and location equal the admin\'s');
  assert.deepEqual(d.flags, new Map([...flags].filter(([k]) => rows.some((r) => r.key === k))));
});

test('clinicPack: an oversized record has its past appointments trimmed to the last 5, then errors when still too long', async () => {
  const pad = 'Synthetic long event description '.repeat(5);
  const appt = (y, m, d) => ({ date: `${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}/${y}`, time: '09:00 AM', event: `TMS Treatment ${pad}`, loc: LOC, rendering: MD, status: 'Completed', authCd: '', authCdLstEnc: '' });
  const past = Array.from({ length: 200 }, (_, i) => appt(2025, 1 + (i % 12), 1 + (i % 28)));
  const future = [appt(2026, 10, 27), appt(2026, 10, 28)];
  const p = (appointments) => ({ 'Per Nbr': '601', MRN: '900601', 'Last Name': 'Bigchartson', 'First Name': 'Testa', DOB: '01/01/1990', 'In Scope': 'Yes', 'Primary Payer': 'Cigna',
    x: { key: '900601', location: LOC, appointments } });
  assert.ok(JSON.stringify(p(past.concat(future))).length > 32000, 'the fixture is oversized');
  const pack = R.clinicPack(LOC, { patients: [p(past.concat(future))], results: [], runAt: TODAY });
  const se = R.clinicUnpack((await roundTrip(pack)).tables);
  const kept = se.patients[0].x.appointments;
  assert.equal(kept.length, 7, 'future 2 + last 5 past');
  assert.deepEqual(kept.slice(-2), future);
  const pastKept = kept.filter((a) => a.date.endsWith('2025')).map((a) => C.parseDate(a.date).getTime());
  const last5 = past.map((a) => C.parseDate(a.date).getTime()).sort((a, b) => a - b).slice(-5);
  assert.deepEqual(pastKept.sort((a, b) => a - b), last5);
  assert.equal(R.clinicDerive(se, [], TODAY).rows.length, 1);
  const many = Array.from({ length: 200 }, (_, i) => appt(2027, 1 + (i % 12), 1 + (i % 28)));
  assert.throws(() => R.clinicPack(LOC, { patients: [p(many)], results: [], runAt: TODAY }), /San Rafael: MRN 900601 has too much data for one Excel cell .*limit is 32,000/);
});

test('clinicUnpack: wrong or missing SE Info is a readable error; optional sheets may be missing', () => {
  const info = (o) => [['Setting', 'Value'], ...Object.entries({ Kind: 'Clinic Eligibility', Location: LOC, 'Run Date': '10/22/2026', Format: '1', ...o })];
  assert.throws(() => R.clinicUnpack({ Clinic: [['MRN']] }), /not a Clinic Eligibility file: it has no SE Info sheet/);
  assert.throws(() => R.clinicUnpack({ 'SE Info': info({ Kind: 'Sohar Clinic' }) }), /"Sohar Clinic" file, not a Clinic Eligibility file/);
  assert.throws(() => R.clinicUnpack({ 'SE Info': info({ Format: '2' }) }), /format 2; this add-in reads format 1/);
  assert.throws(() => R.clinicUnpack({ 'SE Info': info(), 'SE Patients': [['Key', 'Record'], ['900001', '{not json']] }), /SE Patients sheet, row 2, cannot be read/);
  const se = R.clinicUnpack({ 'se info': info() });
  assert.deepEqual([se.patients, se.results, se.review, se.payers, se.done, se.payerRules], [[], [], [], [], [], []]);
  assert.deepEqual(se.settings, { sfhpProviders: [], reverifyDays: 7, reverifyNewMonth: false, approachingPct: 0.2, locationState: {}, defaultState: '' });
  assert.deepEqual(R.clinicDerive(se, [], TODAY).counts.patients, 0);
});
