// v0.30: the master as the central workbook. Billing tickets, balances and financial notes (record.js), the Power Query
// freshness, and the clinic workbook as Start Here + the appointment list (rules.clinicStartPack). Synthetic data only:
// made-up names, MRNs 9002xx, made-up tickets.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createMock } from './helpers/excel-mock.mjs';
const require = createRequire(import.meta.url);
const REC = require('../record.js');
const R = require('../rules.js');
const XL = require('../xlsx.js');

const D = (y, m, d) => new Date(Date.UTC(y, m - 1, d));
const TODAY = new Date(2026, 10, 10, 9, 0); // Tue 11/10/2026
const T = (o) => Object.assign({ 'Ticket Id': '', Submitted: '', Clinic: 'San Rafael', LASFIR: 'TESPAT', MRN: '', Status: 'MHS - Submitted to Arietis', 'Task Type': 'Refund',
  Amount: '', 'Confirmation Due': '', 'Resolution Due': '', 'Escalation Flag': '', 'Last Reply At': '' }, o);
const TICKETS = [
  T({ 'Ticket Id': '101', MRN: 900201, Submitted: '11/2/2026', 'Confirmation Due': '11/4/2026', 'Resolution Due': '11/9/2026' }),          // past due
  T({ 'Ticket Id': '102', MRN: '0900201', Submitted: '11/6/2026', 'Confirmation Due': '11/10/2026', 'Resolution Due': '11/11/2026' }),   // due soon
  T({ 'Ticket Id': '103', MRN: '900202', Clinic: 'Castro', Submitted: '11/3/2026', 'Confirmation Due': '11/5/2026', 'Resolution Due': '11/17/2026' }), // waiting on Arietis
  T({ 'Ticket Id': '104', MRN: '900202', Clinic: 'Castro', Submitted: '11/9/2026', 'Confirmation Due': '11/11/2026', 'Resolution Due': '11/16/2026', 'Last Reply At': '11/9/2026' }),
  T({ 'Ticket Id': '105', MRN: '', Clinic: '', Submitted: 46310 }), // an Excel serial date, no MRN
];

test('business days skip weekends; tickets match the MRN with or without leading zeros or a numeric cell, oldest first, with SLA', () => {
  assert.equal(REC.businessDays(D(2026, 11, 6), D(2026, 11, 10)), 2, 'Fri to Tue');
  assert.equal(REC.businessDays(D(2026, 11, 10), D(2026, 11, 6)), -2);
  const t = REC.openTicketsOf(TICKETS, '900201', TODAY);
  assert.deepEqual(t.map((x) => [x.id, x.sla, x.age]), [['101', 'Past due', 6], ['102', 'Due soon', 2]]);
  const c = REC.openTicketsOf(TICKETS, '900202', TODAY);
  assert.deepEqual(c.map((x) => [x.id, x.sla]), [['103', 'Waiting on Arietis'], ['104', 'On time']]);
  assert.equal(c[0].resolutionDue, 'Nov 17');
  assert.deepEqual(REC.openTicketsOf(TICKETS, '', TODAY), [], 'no MRN matches nothing');
  assert.equal(REC.ticketList(TICKETS, TODAY).length, 5);
  const n = REC.ticketCounts(TICKETS, TODAY);
  assert.equal(n.total, 5);
  assert.equal(n.pastDue, 1);
  assert.equal(n.noMrn, 1);
  assert.deepEqual(n.byClinic, { 'San Rafael': 2, Castro: 2, '(no clinic)': 1 });
  assert.deepEqual(n.byAge, { '0-2': 2, '3-5': 1, '6-10': 1, '11+': 1 });
});

test('balance by MRN, else Per Nbr; financial notes newest first, cards scrubbed, a name and an MRN required', () => {
  const B = [{ MRN: '900201', 'Per Nbr': '', Balance: '$1,234.50', 'As Of': '11/9/2026' }, { MRN: '', 'Per Nbr': '702', Balance: '40', 'As Of': '' }];
  assert.deepEqual(REC.balanceOf(B, '0900201'), { balance: 1234.5, asOf: 'Nov 9' });
  assert.deepEqual(REC.balanceOf(B, '900299', '702'), { balance: 40, asOf: '' });
  assert.equal(REC.balanceOf(B, '900298', '703'), null);
  let rows = REC.addFinancialNote([], { mrn: '900201', lasfir: 'TESPAT', note: 'Card  4111 1111 1111 1111 declined, will call', category: 'Copay', by: 'Desk A', when: '11/10/2026 9:00 AM' });
  rows = REC.addFinancialNote(rows, { mrn: '900201', lasfir: 'TESPAT', note: 'Payment plan agreed', category: 'Nope', ticket: '101', by: 'Desk A', when: '11/10/2026 9:05 AM' });
  assert.deepEqual(Object.keys(rows[0]), REC.FINANCIAL_NOTE_COLS);
  assert.equal(rows[0].Note, 'Card [card number removed] declined, will call');
  assert.equal(rows[1].Category, 'Other', 'an unknown category becomes Other');
  const list = REC.financialNotesOf(rows, 900201);
  assert.deepEqual(list.map((x) => [x.note.slice(0, 7), x.ticket]), [['Payment', '101'], ['Card [c', '']]);
  assert.throws(() => REC.addFinancialNote([], { mrn: '900201', note: 'x' }), /Pick your name/);
  assert.throws(() => REC.addFinancialNote([], { mrn: '', note: 'x', by: 'Desk A' }), /no MRN/);
  assert.throws(() => REC.addFinancialNote([], { mrn: '900201', note: '   ', by: 'Desk A' }), /empty/);
  assert.equal(REC.addFinancialNote([], { mrn: '900201', note: 'x'.repeat(900), by: 'Desk A' })[0].Note.length, REC.NOTE_MAX);
});

test('Power Query freshness: up to date within a day, aging to 3 days, stale after or on any query error', () => {
  const at = (h) => new Date(TODAY.getTime() - h * 36e5).toISOString();
  assert.equal(REC.dataFreshness(null, TODAY).level, 'none');
  assert.equal(REC.dataFreshness({ at: at(5), errors: [], count: 4 }, TODAY).level, 'fresh');
  assert.deepEqual(REC.dataFreshness({ at: at(50), errors: [], count: 4 }, TODAY), { level: 'aging', text: '2 days old. Press Data > Refresh All.' });
  assert.equal(REC.dataFreshness({ at: at(100), errors: [], count: 4 }, TODAY).level, 'stale');
  assert.match(REC.dataFreshness({ at: at(1), errors: ['qOpenTickets: file not found'], count: 4 }, TODAY).text, /qOpenTickets: file not found/);
  assert.equal(REC.dataFreshness({ at: '', errors: [], count: 2 }, TODAY).level, 'stale', 'never refreshed');
});

// The NextGen Appointments Listing as it arrives: title rows, the header on row 7, a footer (the apptlist11.9.xlsx shape).
const HEAD = ['Appointment Pre-Check Complete', 'Res Name', 'Pat Sts', 'Pat Name', 'Appt Date', 'Beg Tm', 'Event', 'Details', 'Loc Name', 'Gender Identity',
  'MRN', 'Status', 'Dur', 'Birth Dt', 'Rendering', 'Payer Name (Lst Enc)', 'Per Nbr', 'Prev Lst Name', 'Guar Name', 'Guar Zip', 'Pol Nbr (Lst Enc)', 'Auth Cd (Lst Enc)',
  'Payer Name (Cht Guar)', 'Pol Nbr (Cht Guar)', 'Other ID Nbr', 'Crt Dt', 'Created By', 'Mod Dt', 'Modified By'];
const ROW = (o) => HEAD.map((h) => (o[h] == null ? '' : o[h]));
const GRID = () => [
  ['Mindful Health Solutions'], ['Appointments Listing'], [''], ['Run 11/10/2026 6:58 AM'], [''], [''],
  HEAD,
  ROW({ 'Res Name': 'Reyes, Sam MD', 'Pat Name': 'Featherstone, Ophelia Q', 'Appt Date': '11/12/2026', 'Beg Tm': '9:00 AM', Event: 'Med Management', Details: 'Ophelia asked about her copay',
    'Loc Name': 'San Rafael', MRN: '900201', Status: 'Expected', Dur: '30', 'Birth Dt': '2/3/1971', 'Per Nbr': '701', 'Prev Lst Name': 'Oldname', 'Guar Name': 'Featherstone, Grant',
    'Guar Zip': '94901', 'Pol Nbr (Lst Enc)': 'ZXQ123456', 'Auth Cd (Lst Enc)': 'AUTH77', 'Payer Name (Cht Guar)': 'Cigna', 'Pol Nbr (Cht Guar)': 'ZXQ123456', 'Other ID Nbr': 'OID99',
    'Crt Dt': '10/1/2026', 'Created By': 'Staffer, Pat', 'Mod Dt': '10/2/2026', 'Modified By': 'Staffer, Pat' }),
  ROW({ 'Res Name': 'Lark, Jordan PMHNP', 'Pat Name': 'Quintero, Bram', 'Appt Date': '11/13/2026', 'Beg Tm': '10:00 AM', Event: 'Telepsych', 'Loc Name': 'Castro', MRN: '900202',
    Status: 'Expected', 'Birth Dt': '5/6/1980', 'Per Nbr': '702', 'Payer Name (Cht Guar)': 'Aetna' }),
  ROW({ 'Res Name': 'Reyes, Sam MD', 'Pat Name': 'Featherstone, Ophelia Q', 'Appt Date': '10/30/2026', 'Beg Tm': '9:00 AM', Event: 'Med Management', 'Loc Name': 'San Rafael',
    MRN: '900201', Status: 'Kept', 'Birth Dt': '2/3/1971', 'Per Nbr': '701', 'Payer Name (Cht Guar)': 'Cigna' }),
  ['Total Appointments: 3'],
];
const NEVER = ['Featherstone', 'Ophelia', 'Quintero', 'Bram', 'Grant', 'Oldname', 'Staffer', '2/3/1971', '94901', 'ZXQ123456', 'AUTH77', 'OID99', 'copay', 'Castro', '900202'];

test('clinic appointment slice: this clinic only, LASFIR in place of names, no DOB, guarantor, policy, auth, details or staff', () => {
  const a = R.clinicApptRows('San Rafael', { importGrid: GRID() });
  assert.equal(a.columns[0], 'LASFIR');
  assert.deepEqual(a.rows.map((r) => [r.LASFIR, r['Appt Date'], r.Status]), [['FEAOPH', '11/12/2026', 'Expected'], ['FEAOPH', '10/30/2026', 'Kept']]);
  for (const c of a.columns) assert.ok(c === 'LASFIR' || R.CLINIC_APPT_KEEP.includes(c), c);
  for (const c of ['Pat Name', 'Birth Dt', 'Guar Name', 'Guar Zip', 'Pol Nbr (Cht Guar)', 'Auth Cd (Lst Enc)', 'Details', 'Created By', 'Prev Lst Name', 'Other ID Nbr', 'Gender Identity'])
    assert.ok(!a.columns.includes(c), `no ${c}`);
  // The NG Appointments query rows (objects, header already found) give the same slice.
  const objs = GRID().slice(7, 10).map((r) => Object.fromEntries(HEAD.map((h, i) => [h, r[i]])));
  assert.deepEqual(R.clinicApptRows('San Rafael', { ngAppts: objs }).rows, a.rows);
  assert.equal(R.lasfirOfName('Featherstone, Ophelia Q'), 'FEAOPH');
});

test('clinicStartPack: Start Here first, the Appointment List, CTT opens with it, no hidden sheets, no lock, nothing identifying', async () => {
  const pack = R.clinicPack('San Rafael', { eligibility: false, importGrid: GRID(), runAt: TODAY, version: '0.30.0', contact: 'Central eligibility team',
    links: [{ label: 'Billing ticket form', url: 'https://forms.example.test/x' }, { label: 'Not https', url: 'http://x.test' }] });
  assert.deepEqual(pack.sheets.map((s) => [s.name, !!s.hidden]), [['Start Here', false], ['Appointment List', false]]);
  assert.equal(pack.protection, null, 'CTT adds template and report tabs, so the structure is not locked');
  assert.deepEqual(pack.counts, { appointments: 2, links: 1, patients: 1 });
  assert.equal(pack.webExtension.id, R.CTT_EXTENSION.id);
  const wb = await XL.readWorkbook(XL.buildWorkbook(pack));
  const text = JSON.stringify(wb.sheets.map((s) => s.rows));
  for (const s of NEVER) assert.ok(!text.toLowerCase().includes(s.toLowerCase()), `"${s}" is not in the file`);
  assert.ok(text.includes('Central eligibility team') && text.includes('https://forms.example.test/x') && !text.includes('http://x.test'));
  assert.ok(!wb.sheets.some((s) => /^SE /.test(s.name)), 'no eligibility tables for now');
  // The v0.29 eligibility file is still there when asked for.
  assert.equal(R.clinicPack('San Rafael', { patients: [], results: [], runAt: TODAY }).sheets[0].name, 'Clinic Day');
});

test('excel: Financial Notes is an OWNED hidden table; Config reads the Start Here links (https only) and contact', async () => {
  const m = createMock(), X = m.SoharExcel;
  await X.writeFinancialNotes(REC.FINANCIAL_NOTE_COLS, [{ MRN: '0900201', LASFIR: 'TESPAT', Note: 'Payment plan', Category: 'Balance', 'Ticket Id': '', By: 'Desk A', When: '11/10/2026' }]);
  assert.equal(m.sheetVisibility('Financial Notes'), 'Hidden');
  assert.equal((await X.readFinancialNotes())[0].MRN, '0900201', 'MRN keeps its leading zero');
  m.seedTable('Config', 'tblConfig', ['Setting', 'Value', 'What it does'], [
    { Setting: 'Start Here links', Value: 'Billing ticket form | https://forms.example.test/x; Bad | http://x.test; Luma|https://luma.example.test', 'What it does': '' },
    { Setting: 'Start Here contact', Value: 'Central eligibility team', 'What it does': '' }]);
  m.seedTable('Location', 'tblLocationState', ['Location', 'State'], []);
  const cfg = await X.readConfig();
  assert.deepEqual(cfg.startLinks, [{ label: 'Billing ticket form', url: 'https://forms.example.test/x' }, { label: 'Luma', url: 'https://luma.example.test' }]);
  assert.equal(cfg.startContact, 'Central eligibility team');
});
