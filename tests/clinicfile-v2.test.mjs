import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
const require = createRequire(import.meta.url);
const C = require('../core.js');
const R = require('../rules.js');
const XL = require('../xlsx.js');
const F = require('../fees.js');

// v0.29 Clinic Eligibility file v2: the San Rafael and Castro files replace the template for front-desk work. Guest
// patients, appointments, ledger estimates, bounded history, card on file, OK to collect, desk collections, fee tables,
// the Clinic Day sheet, very hidden tabs and the structure lock. Synthetic data only: made-up names, MRNs 9001xx.

const HDR = ['Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Loc Name', 'Per Nbr', 'Md Rc', 'Birth Dt', 'Sts', 'Guar Name', 'Guar Birth Dt', 'Guar Zip',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Sec Payer (Cht Guar)', 'Rendering'];
const row = (o) => HDR.map((k) => (o[k] == null ? '' : o[k]));
const MAP = [
  { nextgenPayerName: 'Blue Shield CA', payerId: '94036', payerName: 'Blue Shield of California', confidence: 'Exact' },
  { nextgenPayerName: 'Cigna', payerId: '62308', payerName: 'Cigna', confidence: 'Exact' },
];
const LOCSTATE = { 'San Rafael': 'CA', 'Castro': 'CA', 'Walnut Creek': 'CA' };
const CFG = { asOf: new Date(Date.UTC(2026, 10, 10)), scope: 'upcoming', locationState: LOCSTATE, teleKeywords: ['Tele'], specialty: 'PSY', payerMap: MAP };
const MD = 'Reyes, Sam MD';
const ap = (per, mrn, name, date, time, event, loc, payer, pol, extra) => row({ 'Per Nbr': per, 'Md Rc': mrn, 'Pat Name': name, 'Birth Dt': '3/4/1980', 'Appt Dt': date,
  'Beg Tm': time, Event: event, 'Loc Name': loc, Sts: 'Expected', 'Payer Name (Cht Guar)': payer, 'Pol Nbr (Cht Guar)': pol, 'Guar Zip': '94114', Rendering: MD, ...(extra || {}) });
const IMPORT = () => [HDR,
  ap('601', '900101', 'Larkspur, Wendeline', '11/11/2026', '9:00 AM', 'Telepsych', 'San Rafael', 'Cigna', 'C601'),
  ap('601', '900101', 'Larkspur, Wendeline', '11/13/2026', '10:00 AM', 'Med Management', 'San Rafael', 'Cigna', 'C601'),
  ap('602', '900102', 'Quincefield, Ambrose', '10/20/2026', '1:00 PM', 'Med Management', 'Castro', 'Cigna', 'C602'),
  ap('602', '900102', 'Quincefield, Ambrose', '11/12/2026', '1:00 PM', 'Med Management', 'Castro', 'Cigna', 'C602'),
  ap('602', '900102', 'Quincefield, Ambrose', '11/17/2026', '2:30 PM', 'Med Management', 'San Rafael', 'Cigna', 'C602'),
  ap('603', '900103', 'Tumbledown, Perpetua', '11/12/2026', '11:00 AM', 'Med Management', 'San Rafael', 'Blue Shield CA', 'BS603'),
  ap('604', '900104', 'Farawayson, Dorothea', '11/12/2026', '9:00 AM', 'Med Management', 'Walnut Creek', 'Cigna', 'C604'),
  ap('605', '900105', 'Cancellio, Rupert', '11/12/2026', '9:00 AM', 'Med Management', 'San Rafael', 'Cigna', 'C605', { Sts: 'Cancelled' }),
  ap('605', '900105', 'Cancellio, Rupert', '11/16/2026', '9:00 AM', 'Med Management', 'Castro', 'Cigna', 'C605'),
];
const NEVER = ['Larkspur', 'Wendeline', 'Quincefield', 'Ambrose', 'Tumbledown', 'Perpetua', 'Farawayson', 'Dorothea', 'Cancellio', 'Rupert',
  '03/04/1980', '3/4/1980', '94114', '900104', '128.21', 'Allowed'];

const VER_HDR = ['requestId', 'id', 'url', 'created', 'resolved', 'memberId', 'payerName', 'status', 'statusReason', 'firstName', 'lastName',
  'planName', 'innCopay', 'innTotalDeductible', 'innRemainingDeductible', 'innOutOfPocket', 'innOutOfPocketRemaining'];
const ver = (o) => VER_HDR.map((k) => (o[k] == null ? '' : String(o[k])));
const URL_ = (id) => `https://app.soharhealth.com/verifications/${id}`;
const FILES = () => [{ name: 'verification output.csv', rows: [VER_HDR, ...[
  { requestId: 'v-601', id: '900101', url: URL_('v-601'), created: '2026-11-09', resolved: '2026-11-09T17:00:00Z', memberId: 'C601', payerName: 'Cigna', status: 'complete.eligible',
    firstName: 'Wendeline', lastName: 'Larkspur', planName: 'SYNTH OPEN ACCESS', innCopay: '30', innTotalDeductible: '1000', innRemainingDeductible: '100', innOutOfPocket: '4000', innOutOfPocketRemaining: '2500' },
  { requestId: 'v-602', id: '900102', url: URL_('v-602'), created: '2026-11-09', resolved: '2026-11-09T18:00:00Z', memberId: 'C602', payerName: 'Cigna', status: 'complete.eligible', innCopay: '20' },
  { requestId: 'v-603', id: '900103', url: URL_('v-603'), created: '2026-11-09', resolved: '2026-11-09T18:00:00Z', memberId: 'BS603', payerName: 'Blue Shield of California', status: 'complete.ineligible' },
].map(ver)] }];
const TODAY = new Date(2026, 10, 10, 9, 0); // Tue 11/10/2026, 9:00 local
const L = (o) => Object.assign(Object.fromEntries(R.LEDGER_COLS.map((c) => [c, ''])), o);
const iso = (d) => new Date(d).toISOString();

function source(extra) {
  const built = C.buildPatients(IMPORT(), CFG);
  const merged = C.mergeSoharResults(IMPORT(), FILES(), CFG);
  const patients = C.applyReverifySkips(built.patients, merged.results, { ...C.REVERIFY_DEFAULTS, today: TODAY }).patients.filter((p) => p['In Scope'] === 'Yes');
  return Object.assign({
    patients, results: merged.results, review: built.review, tasks: [], learningRows: [], metricRows: [],
    cfg: { sfhpProviders: [], reverifyDays: 7, reverify: { newMonth: false }, approachingPct: 0.2, locationState: LOCSTATE, defaultState: '' },
    payerRuleRows: [], runAt: TODAY, version: '0.29.0',
    ledger: [
      L({ 'Row Key': '900101|11/11/2026|9:00 AM', MRN: '900101', LASFIR: 'LARWEN', 'Appt Date': '11/11/2026', Time: '9:00 AM', Location: 'San Rafael', Event: 'Telepsych',
        'Projected Patient Responsibility': '45', 'Est Patient': '44.10', Allowed: '128.21', 'Fee Tab': 'Marin', 'Fee Payer': 'CIGNA' }),
      L({ 'Row Key': '900101|11/01/2026|9:00 AM', MRN: '900101', 'Appt Date': '11/01/2026', Time: '9:00 AM', Location: 'San Rafael', 'Est Patient': '10' }),
      L({ 'Row Key': '900102|11/17/2026|2:30 PM', MRN: '900102', 'Appt Date': '11/17/2026', Time: '2:30 PM', Location: 'San Rafael', 'Est Patient': '30' }),
      L({ 'Row Key': '900102|11/12/2026|1:00 PM', MRN: '900102', 'Appt Date': '11/12/2026', Time: '1:00 PM', Location: 'Castro', 'Est Patient': '35' }),
      L({ 'Row Key': '900102|10/20/2026|1:00 PM', MRN: '900102', 'Appt Date': '10/20/2026', Time: '1:00 PM', Location: 'Castro', Event: 'Med Management' }),
      L({ 'Row Key': '900104|11/12/2026|9:00 AM', MRN: '900104', 'Appt Date': '11/12/2026', Time: '9:00 AM', Location: 'Walnut Creek', 'Est Patient': '99' }),
    ],
    history: [
      { 'Row Key': 'card-1', MRN: '900101', LASFIR: 'LARWEN', When: iso('2026-11-09T19:00:00Z'), Kind: 'Card on file', Location: 'San Rafael', Detail: 'Card saved in LumaPay for Larkspur', Source: 'Luma', Ref: 'card-1' },
      { 'Row Key': 'old-1', MRN: '900101', When: iso('2026-07-01T16:00:00Z'), Kind: 'Card on file', Detail: 'Older than 90 days', Source: 'Luma', Ref: 'old-1' },
    ],
    copayActivity: [1, 2, 3, 4, 5].map((i) => ({ 'Row Key': `luma-${i}`, MRN: '0900101', Date: `10/0${i}/2026`, Location: 'San Rafael', Kind: 'Luma charge', Amount: '30', Method: 'Card', Source: 'Luma', Ref: `luma-${i}` })),
    cards: [{ MRN: '900101', 'Card On File': 'Yes', Source: 'Luma checklist', 'As Of': '11/09/2026' }, { MRN: '900103', 'Card On File': 'No', Source: 'Luma checklist', 'As Of': '11/09/2026' }],
    consents: [
      { MRN: '900101', Location: 'Castro', 'OK to Collect': 'No', By: 'Desk B', When: iso('2026-11-01T17:00:00Z') },
      { MRN: '900101', Location: 'San Rafael', 'OK to Collect': 'Yes', By: 'Desk A', When: iso('2026-11-09T17:00:00Z') },
    ],
    desk: [],
    fees: {
      columns: F.FEE_COLS,
      rows: [{ Tab: 'Marin', Locality: 'Locality 52', Locations: 'San Rafael', Code: '99214', Description: 'Est pt', Payer: 'Synthetic Payer', Level: 'MD', Rate: 101.5, Note: '' }],
      rules: [{ Kind: 'payerMap', Record: Array.from({ length: 400 }, (_, i) => ({ Payer: `Synthetic Payer ${i}`, 'Fee Tab': 'Marin', 'Fee Payer': 'Synthetic Payer', Note: 'x'.repeat(80) })) },
        { Kind: 'config', Record: { teleKeywords: ['Tele'] } }],
    },
  }, extra || {});
}
const tablesOf = async (pack) => {
  const wb = await XL.readWorkbook(XL.buildWorkbook(pack));
  return { wb, tables: Object.fromEntries(wb.sheets.map((s) => [s.name, s.rows])) };
};
const objs = (tables, name) => R.tableObjects(tables[name]);

test('v2 clinic set: a guest patient comes along with only this clinic\'s visits; cancelled and elsewhere-only patients do not', async () => {
  const sr = R.clinicPack('San Rafael', source());
  const { tables } = await tablesOf(sr);
  assert.deepEqual([...new Set(objs(tables, 'SE Patients').map((r) => r.Key))].sort(), ['900101', '900102', '900103']);
  const guest = JSON.parse(objs(tables, 'SE Patients').find((r) => r.Key === '900102').Record);
  assert.deepEqual(guest.x.appointments.map((a) => [a.date, a.loc]), [['11/17/2026', 'San Rafael']], 'the Castro visits stay with the admin');
  assert.deepEqual(objs(tables, 'SE Appointments').map((a) => [a.Key, a.Date, a.Time, a.Location]),
    [['900101', '11/11/2026', '9:00 AM', 'San Rafael'], ['900101', '11/13/2026', '10:00 AM', 'San Rafael'], ['900103', '11/12/2026', '11:00 AM', 'San Rafael'], ['900102', '11/17/2026', '2:30 PM', 'San Rafael']],
    'patients soonest first, each with its visits here');
  const ca = R.clinicPack('Castro', source());
  const castro = (await tablesOf(ca)).tables;
  assert.deepEqual([...new Set(objs(castro, 'SE Patients').map((r) => r.Key))].sort(), ['900102', '900105'], 'Castro: its own patient and the one whose San Rafael visit was cancelled');
  // The default (Sohar's own clinic list) is unchanged: next visit at the location only.
  assert.deepEqual(R.clinicRows('San Rafael', sr.sheets && source().patients, source().results, new Map(), TODAY).map((r) => r.MRN).sort(), ['900101', '900103']);
});

test('v2 tables: ledger estimates for this clinic\'s upcoming visits (no Allowed), card on file, newest OK to collect, history bounded and scrubbed', async () => {
  const { tables, wb } = await tablesOf(R.clinicPack('San Rafael', source()));
  const est = objs(tables, 'SE Estimates');
  assert.deepEqual(est.map((e) => [e.Key, e['Appt Date'], e['Projected Patient Responsibility'] || e['Est Patient']]), [['900101', '11/11/2026', '45'], ['900102', '11/17/2026', '30']]);
  assert.ok(!R.SE_EST_COLS.includes('Allowed'));
  const copay = Object.fromEntries(objs(tables, 'SE Copay').map((c) => [c.Key, c]));
  assert.equal(copay['900101']['Card On File'], 'Yes');
  assert.equal(copay['900101']['OK to Collect'], 'Yes', 'the newest answer wins, whichever clinic took it');
  assert.equal(copay['900101']['Consent By'], 'Desk A');
  assert.equal(copay['900103']['Card On File'], 'No');
  assert.equal(copay['900102']['Card On File'], '');
  const h = objs(tables, 'SE History');
  const mine = h.filter((r) => r.Key === '900101');
  assert.equal(mine.length, 5, 'at most 5 events per patient');
  assert.ok(!mine.some((r) => /Older than 90 days/.test(r.Detail)), 'nothing older than 90 days');
  assert.equal(mine[0].Kind, 'Card on file', 'newest first');
  assert.equal(mine[0].Detail, 'Card saved in LumaPay for LARWEN', 'free text scrubbed of the name');
  assert.ok(mine.some((r) => r.Kind === 'OK to collect' && /Yes \(Desk A\)/.test(r.Detail)));
  assert.ok(h.some((r) => r.Key === '900102' && r.Kind === 'Visit' && r.Location === 'Castro' && r.When === '10/20/2026'), 'a guest\'s past visit elsewhere is history');
  assert.ok(h.some((r) => r.Key === '900102' && r.Kind === 'Verification'), 'Sohar results are history');
  for (const s of wb.sheets.filter((x) => x.name !== 'Clinic Day' && x.name !== 'Clinic Work' && x.name !== 'Read Me')) assert.equal(s.veryHidden, true, `${s.name} is very hidden`);
  const text = JSON.stringify(tables).toLowerCase();
  for (const s of NEVER) assert.ok(!text.includes(s.toLowerCase()), `"${s}" is not in the file`);
});

test('v2 fee tables: this clinic\'s rows, and a list too long for one cell split over rows and joined back', async () => {
  const pack = R.clinicPack('San Rafael', source());
  const rules = pack.sheets.find((s) => s.name === 'SE Fee Rules').rows;
  assert.ok(rules.filter((r) => r.Kind === 'payerMap').length >= 2, 'split');
  assert.ok(rules.every((r) => r.Record.length < 32000));
  const { tables } = await tablesOf(pack);
  const u = R.clinicUnpack(tables);
  const ft = R.feeTables(u.feeRules);
  assert.equal(ft.payerMap.length, 400);
  assert.equal(ft.payerMap[399].Payer, 'Synthetic Payer 399');
  assert.deepEqual(ft.config, { teleKeywords: ['Tele'] });
  assert.equal(u.fees.length, 1);
  assert.equal(String(u.fees[0].Rate), '101.5');
});

test('Clinic Day: one row per visit here, no MRN, coverage, copay, estimate, card, OK to collect from Clinic Work, a pending desk collection', async () => {
  const { tables } = await tablesOf(R.clinicPack('San Rafael', source()));
  const pack = R.clinicUnpack(tables);
  // The file as exported already carries Clinic Day.
  const first = objs(tables, 'Clinic Day');
  assert.deepEqual(first.map((r) => [r.Date, r.LASFIR]), [['11/11/2026', 'LARWEN'], ['11/12/2026', 'TUMPER'], ['11/13/2026', 'LARWEN'], ['11/17/2026', 'QUIAMB']]);
  const work = R.workRows([], [
    { Kind: 'OK to collect', Key: '900103', MRN: '900103', LASFIR: 'TUMPER', Value: 'Yes', By: 'Desk C' },
    { Kind: 'Collected', Key: '900101|2026-11-11T20:00:00.000Z', MRN: '900101', LASFIR: 'LARWEN', Value: '30', Note: 'Card · copay at check-in', By: 'Desk C' },
  ], new Date('2026-11-11T20:00:00Z'));
  const d = R.clinicDerive(pack, work, TODAY);
  const day = R.clinicDay(d, { priceOf: (a) => (a.Key === '900101' && a.Date === '11/13/2026' ? { amount: 41.2 } : {}) });
  assert.deepEqual(Object.keys(day[0]), R.CLINIC_DAY_COLS);
  assert.ok(!R.CLINIC_DAY_COLS.includes('MRN'));
  assert.ok(!JSON.stringify(day).includes('9001'), 'no MRN anywhere on Clinic Day');
  const [a, b, c, g] = day;
  assert.equal(a.Coverage, 'Eligible · Cigna');
  assert.equal(a.Copay, '$30');
  assert.equal(a['Ded left'], '$100');
  assert.equal(a['Est. owes'], '$45', 'the ledger\'s projected amount');
  assert.equal(a['Card on file'], 'Yes (11/9)');
  assert.match(a['OK to collect'], /^Yes · Desk A /);
  assert.equal(a['Copay activity'], '$30 desk, pending', 'collected today, not yet applied by the admin');
  assert.equal(b.Coverage, 'Not eligible · Blue Shield of California');
  assert.match(b['OK to collect'], /^Yes · Desk C/, 'Clinic Work wins over the file');
  assert.equal(b['Card on file'], 'No');
  assert.equal(c['Est. owes'], '$41.20 (est.)', 'CTT prices a visit the rotation has not');
  assert.match(c['Copay activity'], /^Last \$30 on 11\/1[01]$/);
  assert.equal(g['Est. owes'], '$30');
  assert.equal(d.collected.length, 1);
  assert.equal(d.counts.consent, 2);
  // An older file (no v2 sheets) still gives one row per patient.
  const old = R.clinicDay(R.clinicDerive({ ...pack, appts: [], estimates: [], history: [], copay: [] }, [], TODAY));
  assert.deepEqual(old.map((r) => r.LASFIR).sort(), ['LARWEN', 'QUIAMB', 'TUMPER']);
});

test('mergeClinicWork: OK to collect (newest per MRN + clinic, blank removes) and Collected (one row per entry, blank withdraws, bad amounts skipped)', () => {
  const W = (Kind, Key, Value, When, o) => ({ Kind, Key, MRN: Key.split('|')[0], LASFIR: 'SYNPAT', Value, Note: '', By: 'Desk A', When, ...(o || {}) });
  const work = [
    W('OK to collect', '900101', 'Yes', '2026-11-10T17:00:00Z'),
    W('OK to collect', '900103', 'No', '2026-11-10T17:01:00Z'),
    W('Collected', '900101|2026-11-10T17:02:00Z', '30', '2026-11-10T17:02:00Z', { Note: 'Card · copay' }),
    W('Collected', '900103|2026-11-10T17:03:00Z', 'thirty', '2026-11-10T17:03:00Z'),
    W('Collected', '900102|2026-11-10T17:04:00Z', '', '2026-11-10T17:04:00Z'),
  ];
  const now = new Date('2026-11-10T18:00:00Z');
  const ctx = { location: 'San Rafael', requests: [], tasks: [], notes: [], payerNotes: [], applied: [], now,
    consents: [{ MRN: '900101', Location: 'San Rafael', 'OK to Collect': 'No', By: 'Desk B', When: '2026-11-09T17:00:00Z' }],
    desk: [{ Ref: 'San Rafael|900102|2026-11-10T17:04:00Z', MRN: '900102', Location: 'San Rafael', Amount: '20.00', When: '2026-11-10T16:00:00Z' }] };
  const m = R.mergeClinicWork(work, ctx);
  assert.deepEqual(m.consents.map((c) => [c.MRN, c['OK to Collect'], c.By]).sort(), [['900101', 'Yes', 'Desk A'], ['900103', 'No', 'Desk A']]);
  assert.deepEqual(m.desk.map((d) => [d.Ref, d.Amount, d.Method, d.Note]), [['San Rafael|900101|2026-11-10T17:02:00Z', '30.00', 'Card', 'copay']], 'the blank withdrew 900102\'s');
  assert.equal(m.counts.consents, 2);
  assert.equal(m.counts.collected, 2);
  assert.equal(m.counts.skipped, 1, 'a typed word is not an amount');
  assert.equal(R.clinicChanges(m.counts), 4);
  const again = R.mergeClinicWork(work, { ...ctx, consents: m.consents, desk: m.desk, applied: m.applied });
  assert.equal(R.clinicChanges(again.counts), 0, 'the same rows again change nothing');
  // The inbox carries the new kinds through.
  const inbox = R.inboxWork(work.map((r) => ({ ...r, Location: 'San Rafael', File: 'f', Received: r.When })), now);
  assert.equal(inbox.get('San Rafael').length, 5);
  assert.equal(inbox.counts.unknown, 0);
});

test('structure lock: very hidden tabs, workbookProtection with SHA-512 hash and salt only; the hash matches Excel\'s algorithm', async () => {
  const salt = new Uint8Array(16).map((_, i) => i + 1);
  const pr = await XL.protectionHash('synthetic-only', { salt, spinCount: 1000 });
  // Reference: H0 = SHA-512(salt + UTF-16LE password); Hn = SHA-512(Hn-1 + LE32 n).
  let h = createHash('sha512').update(Buffer.concat([Buffer.from(salt), Buffer.from('synthetic-only', 'utf16le')])).digest();
  for (let i = 0; i < 1000; i++) { const it = Buffer.alloc(4); it.writeUInt32LE(i); h = createHash('sha512').update(Buffer.concat([h, it])).digest(); }
  assert.deepEqual(pr, { algorithmName: 'SHA-512', hashValue: h.toString('base64'), saltValue: Buffer.from(salt).toString('base64'), spinCount: 1000 });
  await assert.rejects(XL.protectionHash(''), /Type a password/);
  const pack = R.clinicPack('San Rafael', source({ protection: pr }));
  const parts = XL.buildParts(pack).files;
  const wbXml = parts.find((f) => f.name === 'xl/workbook.xml').data;
  assert.match(wbXml, /<workbookPr\/><workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="[^"]+" workbookSaltValue="[^"]+" workbookSpinCount="1000" lockStructure="1"\/><bookViews>/);
  assert.ok(!parts.some((f) => String(f.data).includes('synthetic-only')), 'the password is nowhere in the file');
  assert.match(wbXml, /<sheet name="SE Patients" sheetId="\d+" state="veryHidden"/);
  assert.match(wbXml, /<sheet name="Clinic Day" sheetId="1" r:id="rId1"\/>/);
  const noLock = XL.buildParts(R.clinicPack('San Rafael', source())).files.find((f) => f.name === 'xl/workbook.xml').data;
  assert.doesNotMatch(noLock, /workbookProtection/);
});

test('calculatorWins: the calculator\'s rate replaces the saved one and is listed; saved-only rows stay; new rates are added', () => {
  const saved = [
    { Tab: 'Marin', Payer: 'Synthetic A', Level: 'MD', Code: '99214', Rate: 100 },
    { Tab: 'Marin', Payer: 'Synthetic A', Level: 'NP', Code: '99214', Rate: 90 },
    { Tab: 'Marin', Payer: 'Synthetic A', Level: 'Gross', Code: '99214', Rate: 300 },
    { Tab: 'Marin', Payer: 'Synthetic A', Level: 'MD', Code: '99417', Rate: 50 },
  ];
  const calc = { rows: [
    { Tab: 'Marin', Payer: 'Synthetic A', Level: 'MD', Code: '99214', Rate: 104.25 },
    { Tab: 'Marin', Payer: 'Synthetic A', Level: 'NP', Code: '99214', Rate: 90 },
    { Tab: 'Marin', Payer: 'Synthetic B', Level: 'MD', Code: '99214', Rate: 88 },
  ], tabs: [], skipped: [] };
  const m = F.calculatorWins(saved, [], calc);
  assert.deepEqual(m.changed, [{ Tab: 'Marin', Payer: 'Synthetic A', Level: 'MD', Code: '99214', 'Old Rate': 100, 'New Rate': 104.25 }]);
  assert.equal(m.conflicts, 1);
  assert.equal(m.added, 1);
  assert.deepEqual(m.rows.map((r) => [r.Payer, r.Level, r.Code, r.Rate]), [['Synthetic A', 'MD', '99214', 104.25], ['Synthetic A', 'NP', '99214', 90],
    ['Synthetic A', 'Gross', '99214', 300], ['Synthetic A', 'MD', '99417', 50], ['Synthetic B', 'MD', '99214', 88]]);
  assert.deepEqual(F.FEE_SCHEDULE_LOG_COLS, ['When', 'Tab', 'Payer', 'Level', 'Code', 'Old Rate', 'New Rate', 'By', 'Source']);
});
