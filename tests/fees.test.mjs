import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';
import { createRequire } from 'node:module';
import { createMock } from './helpers/excel-mock.mjs';
const require = createRequire(import.meta.url);
const C = require('../core.js');
const F = require('../fees.js');
const X = require('../xlsx.js');

// Fee estimates (v0.11). Synthetic data only: every payer column, location and rate below is invented. The layout
// mimics a real fee schedule: Locations row, title with locality, CPT Code header at different offsets, payer names
// over MD / Nurse (NP) columns (merged or not), NA / #REF! / 0.00 cells, a combined code, a note row, hidden tabs.

// ---------- synthetic fixture ----------
const BAYVIEW = {
  name: 'Bayview', hidden: false,
  rows: [
    ['Mindful Health Solutions', '', '', '', '', '', '', 'Locations: Harbor Point, Mill Street and Telehealth'],
    ['Bayview Fee Schedule 2026 Locality 9'],
    [],
    ['Applies to the made-up Bay county'],
    ['', 'CPT Code', 'Description', 'Gross Charge', '', 'Cash', '', 'Acme Health', '', 'Blue Shield', '', 'BS 65', '', 'Imperial', 'Med Adv', 'Kaiser', 'TriWest/VA', '', 'Carelon', 'Carelon', '#REF!', ''],
    ['', '', '', '', '', 'Rate', '', 'MD', 'Nurse', 'MD', 'Nurse', 'MD', 'NP', 'MD', 'NP', 'MD', 'MD', 'Nurse', 'MD', 'NP', '#REF!', 'MD'],
    ['', 90792, 'Psych eval with medical services', 591.11, '', 351.11, '', 291.11, 251.11, '$1,191.11', '1,071.11', 0, '0.00', 'NA', 'N/A', '#REF!', 191.11, 171.11, 151.11, 91.11, 'NA', 71.11],
    ['', '90868', 'TMS subsequent delivery', 551.11, '', 291.11, '', 251.11, 191.11, 271.11, 251.11, '85%', 'see contract', 251.111177, 191.11, 171.11, '#VALUE!', '#N/A', 151.11, 91.11, 'x', 51.11],
    ['', '99417/G2212', 'Prolonged service', 91.11, '', 71.11, '', 71.11, 51.11, 51.11, 51.11, 51.11, 51.11, 71.11, 51.11, 71.11, 51.11, 51.11, 71.11, 51.11],
    ['', '99214', 'Established visit', 251.11, '', 151.11, '', 171.11, 151.11, 151.11, 91.11, 91.11, 71.11, 151.11, 91.11, 151.11, 91.11, 71.11, 91.11, 71.11],
    [],
    ['', 'Mail paper claims, not portal (made-up note)'],
    ['', '90837', 'After the note: not read', 151.11, '', 91.11, '', 71.11, 51.11],
  ],
  merges: [{ r1: 4, c1: 7, r2: 4, c2: 8 }, { r1: 4, c1: 11, r2: 4, c2: 12 }, { r1: 4, c1: 16, r2: 4, c2: 17 }],
};
// No Locations row, code in column A, header lower down: matched by tab name only.
const RIDGE = {
  name: 'Ridge', hidden: false,
  rows: [
    ['Mindful Health Solutions'],
    ['Ridge State Fee Schedule 2026'],
    [],
    [],
    [],
    ['CPT Code', 'Description', 'Gross Charge', '', 'Cash', '', 'Acme Health', '', 'Ridge Medicare Adv', ''],
    ['', '', '', '', 'Rate', '', 'MD', 'Nurse', 'MD', 'NP'],
    ['90868', 'TMS subsequent delivery', 451.11, '', 271.11, '', 251.11, 191.11, 191.11, 171.11],
    ['99417G2212', 'Prolonged service typed without a slash', 91.11, '', 71.11, '', 91.11, 51.11, 71.11, 51.11],
  ],
  merges: [],
};
const HIDDEN = { name: 'Summary', hidden: true, rows: [['', 'CPT Code', 'Acme Health'], ['', '', 'MD'], ['', '90868', 1]], merges: [] };
const THERAPY = { name: 'Therapy-Made Up', hidden: false, rows: [['', 'CPT Code', 'Description', 'Gross Charge', 'Acme Health', ''], ['', '', '', '', 'PsyD', 'Masters'], ['', '90837', 'x', 1, 2, 3]], merges: [] };
const WB = { sheets: [HIDDEN, BAYVIEW, RIDGE, THERAPY] };
const rateOf = (rows, tab, code, payer, level) => { const r = rows.find((x) => x.Tab === tab && x.Code === code && x.Payer === payer && x.Level === level); return r ? r : null; };

test('config: Fee estimates is Off unless the row says On; toggle defaults from Config', () => {
  assert.equal(F.feeConfig({}).on, false);
  assert.equal(F.feeConfig({ 'Fee estimates': 'Off' }).on, false);
  assert.equal(F.feeConfig({ 'Fee estimates': 'On' }).on, true);
  assert.equal(F.feeConfig({ 'Fee estimates': 'yes' }).on, true);
  assert.deepEqual(F.feeConfig({}).defaults, { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' });
  assert.deepEqual(F.feeConfig({ 'Estimate deductible': 'waived', 'Estimate coinsurance': 'before', 'Estimate copay': 'None' }).defaults, { deductible: 'Waived', coinsurance: 'Before deductible', copay: 'None' });
  assert.deepEqual(F.FEE_CONFIG_ROWS.map((r) => [r.setting, r.value]), [['Fee estimates', 'Off'], ['Estimate deductible', 'Applies'], ['Estimate coinsurance', 'After deductible'], ['Estimate copay', 'Added'],
    ['BlueCard host CA', ''], ['BlueCard host TX', ''], ['Medicare without Medicare column', 'Flag'], ['Medicare payer column exceptions', ''],
    ['IVK medication infusions 1-6', '400'], ['IVK medication infusions 7+', '450']]);
  assert.deepEqual(F.feeConfig({}).ivkMed, { first: 400, later: 450 });
  assert.deepEqual(F.feeConfig({ 'IVK medication infusions 1-6': '$410', 'IVK medication infusions 7+': '460' }).ivkMed, { first: 410, later: 460 });
});

test('parse: visible tabs only; tabs with no MD / NP payer columns are skipped', () => {
  const p = F.parseFeeSchedule(WB);
  assert.deepEqual(p.tabs.map((t) => t.name), ['Bayview', 'Ridge']);
  assert.deepEqual(p.skipped, [{ name: 'Summary', reason: 'hidden tab' }, { name: 'Therapy-Made Up', reason: 'no payer columns labeled MD or NP' }]);
  assert.deepEqual(Object.keys(p.rows[0]), F.FEE_COLS);
});

test('parse: Locations row, locality, payers carried over merged and MD / Nurse columns, joined names', () => {
  const t = F.parseFeeSchedule(WB).tabs.find((x) => x.name === 'Bayview');
  assert.deepEqual(t.locations, ['Harbor Point', 'Mill Street', 'Telehealth']);
  assert.equal(t.locality, 'Locality 9');
  assert.equal(t.title, 'Bayview Fee Schedule 2026 Locality 9');
  assert.deepEqual(t.payers, ['Acme Health', 'Blue Shield', 'BS 65', 'Imperial Med Adv', 'Kaiser', 'TriWest/VA', 'Carelon']);
  assert.equal(t.unlabeled, 1, 'an MD column with no payer name is not guessed');
  assert.deepEqual(t.notes, ['Mail paper claims, not portal (made-up note)']);
  assert.equal(t.afterStop, 1, 'code rows after the note are not read');
  assert.equal(t.codes, 5, '90792, 90868, 99417, G2212, 99214');
  const r = F.parseFeeSchedule(WB).tabs.find((x) => x.name === 'Ridge');
  assert.deepEqual(r.locations, []);
  assert.equal(r.locality, '');
  assert.deepEqual(r.payers, ['Acme Health', 'Ridge Medicare Adv']);
});

test('parse: rates, levels, and cells that are not rates (0, 0.00, NA, N/A, #REF!, %, text) keep their text as a note', () => {
  const rows = F.parseFeeSchedule(WB).rows;
  const get = (code, payer, level) => rateOf(rows, 'Bayview', code, payer, level);
  assert.equal(get('90792', 'Gross Charge', 'Gross').Rate, 591.11);
  assert.equal(get('90792', 'Cash', 'Cash').Rate, 351.11);
  assert.equal(get('90792', 'Acme Health', 'MD').Rate, 291.11);
  assert.equal(get('90792', 'Acme Health', 'NP').Rate, 251.11, 'Nurse is the NP level');
  assert.equal(get('90792', 'Blue Shield', 'MD').Rate, 1191.11, '$ and commas stripped');
  assert.equal(get('90792', 'Blue Shield', 'NP').Rate, 1071.11, 'blank header next to an MD column carries the payer');
  assert.deepEqual([get('90792', 'BS 65', 'MD').Rate, get('90792', 'BS 65', 'MD').Note], ['', '0']);
  assert.deepEqual([get('90792', 'BS 65', 'NP').Rate, get('90792', 'BS 65', 'NP').Note], ['', '0.00']);
  assert.deepEqual([get('90792', 'Imperial Med Adv', 'MD').Note, get('90792', 'Imperial Med Adv', 'NP').Note], ['NA', 'N/A']);
  assert.equal(get('90792', 'Kaiser', 'MD').Note, '#REF!');
  assert.equal(get('90792', 'Kaiser', 'NP'), null, 'a single-column payer has no NP rate');
  assert.equal(get('90792', 'TriWest/VA', 'NP').Rate, 171.11, 'merged header covers the Nurse column');
  assert.equal(get('90792', 'Carelon', 'NP').Rate, 91.11, 'a name repeated over its NP column is the same payer');
  assert.deepEqual([get('90868', 'BS 65', 'MD').Note, get('90868', 'BS 65', 'NP').Note], ['85%', 'see contract']);
  assert.equal(get('90868', 'Imperial Med Adv', 'MD').Rate, 251.1112);
  assert.deepEqual([get('90868', 'TriWest/VA', 'MD').Note, get('90868', 'TriWest/VA', 'NP').Note], ['#VALUE!', '#N/A']);
  assert.equal(get('99417', 'Acme Health', 'MD').Rate, 71.11, 'combined code: both codes get the row');
  assert.equal(get('G2212', 'Acme Health', 'MD').Rate, 71.11);
  assert.equal(rows.some((r) => r.Code === '90837'), false);
  assert.equal(rateOf(rows, 'Ridge', 'G2212', 'Acme Health', 'MD').Rate, 91.11, 'two codes typed together are split');
  assert.equal(rateOf(rows, 'Ridge', '90868', 'Ridge Medicare Adv', 'NP').Rate, 171.11);
  assert.equal(rows.filter((r) => r.Tab === 'Bayview')[0].Locations, 'Harbor Point; Mill Street; Telehealth');
  const p = F.parseFeeSchedule(WB);
  assert.equal(p.counts.unparsed, 9, '0, 0.00, NA, N/A, #REF! and 85%, text, #VALUE!, #N/A; cells under no payer are not counted');
  assert.equal(p.counts.tabs, 2);
});

test('parse helpers: codes, rates, Locations lines', () => {
  assert.deepEqual(F.parseCodes('99417/G2212'), ['99417', 'G2212']);
  assert.deepEqual(F.parseCodes(90868), ['90868']);
  assert.deepEqual(F.parseCodes(' g2083 '), ['G2083']);
  assert.deepEqual(F.parseCodes('IVKET'), ['IVKET'], 'an in-house code (one upper-case token)');
  assert.deepEqual(F.parseCodes('Mail paper claims, not portal'), []);
  assert.equal(F.parseCodes(''), null);
  assert.deepEqual(F.parseRate('$1,191.1'), { rate: 1191.1, note: '' });
  assert.equal(F.parseRate('#REF!').kind, 'error');
  assert.equal(F.parseRate('N/A').kind, 'na');
  assert.equal(F.parseRate('75%').kind, 'percent');
  assert.equal(F.parseRate('391..11').kind, 'text', 'a mistyped number is not a rate');
  assert.equal(F.parseRate(0).kind, 'zero');
  assert.deepEqual(F.parseLocations('Location: Lakeside. Hill Top'), ['Lakeside', 'Hill Top']);
  assert.deepEqual(F.parseLocations('Locations: North Bend, East Fork and Old Mill,'), ['North Bend', 'East Fork', 'Old Mill']);
  assert.equal(F.parseLocations('Bayview Fee Schedule'), null);
});

// ---------- xlsx reader ----------
function zipDeflated(files) {
  const enc = new TextEncoder(), locals = [], cds = [];
  let off = 0;
  for (const [name, text] of Object.entries(files)) {
    const data = enc.encode(text), comp = deflateRawSync(data), nm = enc.encode(name), crc = X.crc32(data);
    const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(8, 8); lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nm.length, 26);
    const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(8, 10); ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(nm.length, 28); ch.writeUInt32LE(off, 42);
    locals.push(lh, nm, comp); cds.push(ch, nm);
    off += 30 + nm.length + comp.length;
  }
  const cd = Buffer.concat(cds), eo = Buffer.alloc(22);
  eo.writeUInt32LE(0x06054b50, 0); eo.writeUInt16LE(cds.length / 2, 8); eo.writeUInt16LE(cds.length / 2, 10); eo.writeUInt32LE(cd.length, 12); eo.writeUInt32LE(off, 16);
  return new Uint8Array(Buffer.concat([...locals, cd, eo]));
}
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function sheetXml(rows, merges) {
  const body = rows.map((row, r) => `<row r="${r + 1}">${(row || []).map((v, c) => {
    if (v === '' || v == null) return '';
    const ref = `${X.colLetter(c + 1)}${r + 1}`;
    if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`;
    if (/^#/.test(v)) return `<c r="${ref}" t="e"><v>${xmlEsc(v)}</v></c>`;
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
  }).join('')}</row>`).join('');
  const mc = (merges || []).map((m) => `<mergeCell ref="${X.colLetter(m.c1 + 1)}${m.r1 + 1}:${X.colLetter(m.c2 + 1)}${m.r2 + 1}"/>`).join('');
  return `<?xml version="1.0"?><worksheet xmlns="x"><sheetData>${body}</sheetData>${mc ? `<mergeCells>${mc}</mergeCells>` : ''}</worksheet>`;
}
function workbookFiles(sheets, extra) {
  const files = {
    'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns:r="r"><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ''} r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships>${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="ws" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}</Relationships>`,
  };
  sheets.forEach((s, i) => { files[`xl/worksheets/sheet${i + 1}.xml`] = s.xml || sheetXml(s.rows, s.merges); });
  return { ...files, ...(extra || {}) };
}

test('readWorkbook: deflated entries, shared strings (rich text, phonetic, entities), cell types, hidden state, merges', async () => {
  const ss = '<sst><si><t>CPT Code</t></si><si><r><t>Acme </t></r><r><t xml:space="preserve">&amp; Co</t></r><rPh><t>ignored</t></rPh></si><si><t>Line_x000D_Two</t></si></sst>';
  const sheet1 = '<worksheet><sheetData><row r="2"><c r="B2" t="s"><v>0</v></c><c r="C2" t="s"><v>1</v></c><c r="D2" t="s"><v>2</v></c></row>'
    + '<row r="3"><c r="B3"><v>90868</v></c><c r="C3" t="str"><f>A1</f><v>formula text</v></c><c r="D3" t="e"><v>#REF!</v></c><c r="E3" t="b"><v>1</v></c><c r="F3" t="inlineStr"><is><t>inline &lt;x&gt;</t></is></c><c r="G3" s="1"/><c r="H3"><v>71.11</v></c></row>'
    + '</sheetData><mergeCells count="1"><mergeCell ref="C2:D2"/></mergeCells></worksheet>';
  const bytes = zipDeflated(workbookFiles([{ name: 'Fee & Tab', xml: sheet1 }, { name: 'Secret', hidden: true, rows: [['x']] }], { 'xl/sharedStrings.xml': ss }));
  const wb = await X.readWorkbook(bytes);
  assert.deepEqual(wb.sheets.map((s) => [s.name, s.hidden]), [['Fee & Tab', false], ['Secret', true]]);
  const r = wb.sheets[0].rows;
  assert.equal(r[0], undefined, 'row 1 is empty');
  assert.deepEqual(r[1].slice(1), ['CPT Code', 'Acme & Co', 'Line\rTwo']);
  assert.deepEqual(r[2].slice(1, 6), [90868, 'formula text', '#REF!', true, 'inline <x>']);
  assert.equal(r[2][6], undefined, 'a styled empty cell has no value');
  assert.equal(r[2][7], 71.11);
  assert.deepEqual(wb.sheets[0].merges, [{ r1: 1, c1: 2, r2: 1, c2: 3 }]);
  assert.deepEqual(wb.sheets[1].rows, [['x']]);
});

test('readWorkbook reads the add-in\'s own STORE workbooks too', async () => {
  const bytes = X.buildWorkbook({ sheets: [{ name: 'Clinic', table: 'tblClinic', columns: ['Per Nbr', 'Amount'], rows: [{ 'Per Nbr': '000123', Amount: 91.11 }] }, { name: 'Info', table: 'tblInfo', columns: ['Kind'], rows: [{ Kind: 'x' }], hidden: true }] });
  const wb = await X.readWorkbook(bytes);
  assert.deepEqual(wb.sheets.map((s) => [s.name, s.hidden]), [['Clinic', false], ['Info', true]]);
  assert.deepEqual(wb.sheets[0].rows, [['Per Nbr', 'Amount'], ['000123', 91.11]]);
});

test('readWorkbook + parseFeeSchedule on a deflated synthetic fee schedule gives the same rows as the object fixture', async () => {
  const wb = await X.readWorkbook(zipDeflated(workbookFiles(WB.sheets)));
  const fromFile = F.parseFeeSchedule(wb), fromObj = F.parseFeeSchedule(WB);
  assert.deepEqual(fromFile.tabs.map((t) => t.name), ['Bayview', 'Ridge']);
  assert.deepEqual(fromFile.rows, fromObj.rows);
  await assert.rejects(X.readWorkbook(new TextEncoder().encode('not a zip at all')), /not an \.xlsx/);
});

// ---------- matching ----------
const IDX = () => F.feeIndex(F.parseFeeSchedule(WB).rows);
const TABS = () => [...IDX().values()];

test('location -> fee tab: Locations list or tab name, whole words either way, ties and telehealth never guessed', () => {
  const t = TABS();
  assert.equal(F.matchLocationTab('Harbor Point', t).tab, 'Bayview');
  assert.equal(F.matchLocationTab('MHS - harbor point clinic', t).tab, 'Bayview', 'case and punctuation ignored');
  assert.equal(F.matchLocationTab('Mill', t).tab, 'Bayview', 'a short NextGen name inside a listed location');
  assert.equal(F.matchLocationTab('Ridge Downtown', t).tab, 'Ridge', 'no Locations row: tab name');
  assert.equal(F.matchLocationTab('Telehealth', t, ['Tele']).tab, '', 'telehealth in a Locations list is ignored');
  assert.match(F.matchLocationTab('Faraway', t).why, /^No fee tab for location Faraway/);
  const twin = [{ name: 'North', locations: ['Twin Oaks'] }, { name: 'South', locations: ['Twin Oaks'] }];
  const m = F.matchLocationTab('Twin Oaks', twin);
  assert.equal(m.tab, '');
  assert.match(m.why, /more than one fee tab \(North, South\)/);
});

test('Fee Locations: auto rows, manual edits win and persist across a new schedule', () => {
  const t = TABS();
  const first = F.feeLocationRows([], ['Harbor Point', 'Faraway', 'Telehealth'], t, ['Tele']);
  assert.deepEqual(first.map((r) => [r['NextGen Location'], r['Fee Tab'], r.Source]), [['Faraway', '', 'auto'], ['Harbor Point', 'Bayview', 'auto'], ['Telehealth', '', 'auto']]);
  assert.match(first[2].Note, /^Telehealth/);
  // A person types a tab for Faraway on the sheet (Source still says auto): it becomes manual and is kept.
  const edited = first.map((r) => (r['NextGen Location'] === 'Faraway' ? { ...r, 'Fee Tab': 'Ridge' } : r));
  const again = F.feeLocationRows(edited, ['Harbor Point', 'Faraway'], [], ['Tele']); // a new schedule with no tabs
  const by = Object.fromEntries(again.map((r) => [r['NextGen Location'], r]));
  assert.deepEqual([by.Faraway['Fee Tab'], by.Faraway.Source], ['Ridge', 'manual']);
  assert.deepEqual([by['Harbor Point']['Fee Tab'], by['Harbor Point'].Source], ['', 'auto'], 'auto rows follow the new schedule');
  const third = F.feeLocationRows(again, ['Faraway'], t, ['Tele']);
  assert.equal(third.find((r) => r['NextGen Location'] === 'Faraway')['Fee Tab'], 'Ridge', 'still manual');
});

test('Provider Home Clinic: most frequent in-person location; ties and telehealth-only are not guessed; manual wins', () => {
  const appts = [
    { rendering: 'Doe, Pat MD', loc: 'Mill Street', event: 'TMS' }, { rendering: 'Doe, Pat MD', loc: 'Mill Street', event: 'TMS' },
    { rendering: 'Doe, Pat MD', loc: 'Harbor Point', event: 'TMS' }, { rendering: 'Doe, Pat MD', loc: 'Harbor Point', event: 'Telepsych' },
    { rendering: 'Doe, Pat MD', loc: 'Telehealth', event: 'Follow Up' }, { rendering: 'Doe, Pat MD', loc: 'Telehealth', event: 'Follow Up' },
    { rendering: 'Roe, Sam NP', loc: 'Mill Street', event: 'Visit' }, { rendering: 'Roe, Sam NP', loc: 'Harbor Point', event: 'Visit' },
    { rendering: 'Poe, Lee NP', loc: 'Telehealth', event: 'Telepsych' },
  ];
  const rows = F.providerHomeRows([], appts, ['Tele']);
  const by = Object.fromEntries(rows.map((r) => [r.Provider, r]));
  assert.equal(by['Doe, Pat MD']['Home Location'], 'Mill Street');
  assert.equal(by['Roe, Sam NP']['Home Location'], '');
  assert.match(by['Roe, Sam NP'].Note, /^Tie: Harbor Point \/ Mill Street/);
  assert.equal(by['Poe, Lee NP']['Home Location'], '');
  const manual = rows.map((r) => (r.Provider === 'Roe, Sam NP' ? { ...r, 'Home Location': 'Harbor Point' } : r));
  const again = Object.fromEntries(F.providerHomeRows(manual, appts, ['Tele']).map((r) => [r.Provider, r]));
  assert.deepEqual([again['Roe, Sam NP']['Home Location'], again['Roe, Sam NP'].Source], ['Harbor Point', 'manual']);
});

test('payer -> fee column: name words, "/" sides, Medicare only to Medicare columns, ties and generic-only matches', () => {
  const cols = ['Acme Health', 'Blue Shield', 'BS 65', 'Imperial Med Adv', 'Kaiser', 'TriWest/VA', 'Carelon/Beacon', 'Beacon/Carelon', 'Health Alliance'];
  const m = (p) => { const r = F.matchFeePayer(p, cols); return [r.feePayer, r.confidence]; };
  assert.deepEqual(m('Blue Shield of California'), ['Blue Shield', 'High']);
  assert.deepEqual(m('Blue Shield 65 Plus'), ['BS 65', 'High'], 'the more specific column wins');
  assert.deepEqual(m('Imperial Health Plan Medicare Advantage'), ['Imperial Med Adv', 'High']);
  assert.deepEqual(m('Acme Health Medicare Advantage'), ['', 'Low'], 'a Medicare plan never lands on a commercial column');
  assert.deepEqual(m('TriWest Healthcare Alliance'), ['TriWest/VA', 'High']);
  assert.deepEqual(m('Kaiser Foundation Health Plan'), ['Kaiser', 'High']);
  assert.deepEqual(m('Carelon Behavioral Health'), ['', 'Low'], 'two columns match equally: no guess');
  assert.deepEqual(m('Health Alliance Plan'), ['Health Alliance', 'Medium'], 'generic words only');
  assert.deepEqual(m('Unknown Synthetic Payer'), ['', 'Low']);
  assert.deepEqual(m(''), ['', 'Low']);
});

test('Fee Payer Map: auto per payer and tab, manual edits win (also with a blank tab), every change logged once', () => {
  const idx = IDX();
  const pairs = [{ payer: 'Blue Shield of California', tab: 'Bayview' }, { payer: 'Carelon Behavioral Health', tab: 'Bayview' }];
  const rows = F.feePayerMapRows([], pairs, idx);
  assert.deepEqual(rows.map((r) => [r.Payer, r['Fee Tab'], r['Fee Payer'], r.Confidence, r.Source]),
    [['Blue Shield of California', 'Bayview', 'Blue Shield', 'High', 'auto'], ['Carelon Behavioral Health', 'Bayview', 'Carelon', 'High', 'auto']]);
  const log1 = F.feePayerLog(rows, [], 't1');
  assert.deepEqual(log1.map((l) => [l.Payer, l['New Fee Payer'], l.By]), [['Blue Shield of California', 'Blue Shield', 'Auto-match'], ['Carelon Behavioral Health', 'Carelon', 'Auto-match']]);
  assert.deepEqual(F.feePayerLog(rows, log1, 't2'), [], 'nothing changed, nothing logged');
  // A person changes Blue Shield to BS 65 on the sheet, and adds a payer row with a blank tab (every tab).
  const edited = rows.map((r) => (r.Payer.startsWith('Blue') ? { ...r, 'Fee Payer': 'BS 65' } : r)).concat([{ Payer: 'Ridge Senior Plan', 'Fee Tab': '', 'Fee Payer': 'Ridge Medicare Adv', Confidence: '', Source: 'manual', 'Auto Match': '', Note: '' }]);
  const again = F.feePayerMapRows(edited, pairs, idx);
  const bs = again.find((r) => r.Payer.startsWith('Blue'));
  assert.deepEqual([bs['Fee Payer'], bs.Source, bs.Confidence], ['BS 65', 'manual', 'Manual']);
  const log2 = F.feePayerLog(again, log1, 't3');
  assert.deepEqual(log2.map((l) => [l.Payer, l['Old Fee Payer'], l['New Fee Payer'], l.By]), [['Blue Shield of California', 'Blue Shield', 'BS 65', 'Sheet edit'], ['Ridge Senior Plan', '', 'Ridge Medicare Adv', 'Sheet edit']]);
  assert.equal(F.lookupFeePayer(again, 'ridge senior plan', 'Ridge')['Fee Payer'], 'Ridge Medicare Adv', 'blank tab applies to every tab');
});

test('estimate codes: seeds marked default, review; longest keyword wins; units', () => {
  assert.ok(F.ESTIMATE_CODE_SEEDS.every((r) => r.Source === 'default, review'));
  const c = (ev) => { const r = F.codesForEvent(ev, F.ESTIMATE_CODE_SEEDS); return r.none ? 'none' : r.codes.map((x) => `${x.code}x${x.units}`).join(' '); };
  // Every NextGen event name in use (10/2026, event text only) gets codes, or is labeled "none".
  const expect = {
    'TMS Treatment Unilateral': 'none', 'TMS Treatment Bilateral': 'none', 'TMS Theta Tx Bilateral': 'none', 'TMS Theta Tx Unilateral': 'none',
    'MT Redo': 'none', 'MT MagV': 'none', 'MT BW': 'none', 'MT BW H7': 'none', 'MT Hold - MagVenture': 'none', 'TMS Motor Threshold': 'none', 'TMS Mapping Session': 'none',
    'Esketamine Follow Up': 'none', 'Esketamine Follow Up Telepsych': 'none', 'Esketamine Tx Start': 'none', 'Hold-Esketamine Tx Start': 'none', 'Esketamine Daily Tx': 'none', 'Spravato Tx': 'none',
    'Tel New Patient - Mature': '90792x1', 'IP New Patient - Mature': '90792x1', 'TMS ConsultTelepsych': '90792x1', 'TMS Consult In Person': '90792x1',
    'TMS Consult+Med Mgt In Person': '90792x1', 'TMS Consult+Med Mgt Telepsych': '90792x1', 'TMS Consult Telepsych Internal': '90792x1', 'TMS Consult In Person Internal': '90792x1',
    'Esk ConsultTelepsych': '90792x1', 'Esketamine Consult In Person': '90792x1', 'ESK Consult+Med Mgt Telepsych': '90792x1', 'IV Ketamine Consult Tele': '90792x1',
    'Tel MindfulOne Consult': '90792x1', 'IP MindfulOne Consult': '90792x1',
    'IV Ketamine Treatment Start': '99215x1', 'IV Ketamine Follow Up': '99214x1',
    'Telepsych': '99214x1', 'Med Management Telepsych': '99214x1', 'Med Management In-Office': '99214x1', 'Telepsych TMS F/U': '99214x1', 'TMS Follow Up': '99214x1',
    'Kaiser TMS F/U Telepsych': '99214x1', 'Kaiser TMS F/U': '99214x1', 'Transition Telepsych Patient': '99214x1', 'Transition Appointment': '99214x1',
    'Telepsych Only': '99214x1', 'Telepsych Adolescent F/u': '99214x1', 'Tel MindfulConnect': '99214x1', 'ADHD Testing Review Tele': '99214x1', 'ADHD Testing Review In Person': '99214x1',
    'ADHD Testing 30': '96136x1', 'Psychotherapy Individual 60': '90837x1', 'Therapy 60 min': '90837x1',
    MMI: '99214x1', MMT: '99214x1', TFU: '99214x1', TTF: '99214x1', Tel: '99214x1',
  };
  for (const [ev, want] of Object.entries(expect)) assert.equal(c(ev), want, ev);
  assert.equal(c('Peer To Peer'), '');
  assert.match(F.codesForEvent('TMS Treatment Unilateral', F.ESTIMATE_CODE_SEEDS).note, /^TMS \(intervention\): no estimate; verify deductible, OOP max and insurance status$/);
  assert.match(F.codesForEvent('Esketamine Tx Start', F.ESTIMATE_CODE_SEEDS).note, /^Esketamine \(intervention\)/);
  assert.deepEqual(F.parseCodeList('None'), { codes: [], bad: [], none: true });
  assert.deepEqual(F.parseCodeList('90868 x2; 90869, g2083 X1'), { codes: [{ code: '90868', units: 2 }, { code: '90869', units: 1 }, { code: 'G2083', units: 1 }], bad: [] });
  assert.deepEqual(F.parseCodeList('90868 x2; nonsense here').bad, ['nonsense here']);
});

// ---------- estimate math ----------
const L = (allowed, units) => [{ code: '90868', units: units || 1, allowed }];
const B = (o) => ({ deductibleRemaining: '91.11', coinsurance: '20%', copay: '71.11', oopRemaining: '1091.11', ...o });

test('math: deductible, then coinsurance after deductible, then copay; plan pays the rest', () => {
  const e = F.estimateMath(L(451.11), B(), { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' });
  // 91.11 deductible + 20% of (451.11 - 91.11) = 72 + 71.11 copay = 234.22; plan 451.11 - 234.22 = 216.89
  assert.deepEqual([e.allowed, e.deductible, e.coinsurance, e.copay, e.patient, e.plan, e.deductibleAfter], [451.11, 91.11, 72, 71.11, 234.22, 216.89, 0]);
  assert.deepEqual(e.missing, []);
});

test('math: each toggle', () => {
  const run = (t) => F.estimateMath(L(451.11), B(), { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added', ...t });
  const w = run({ deductible: 'Waived' });
  assert.deepEqual([w.deductible, w.coinsurance, w.copay, w.patient, w.deductibleAfter], [0, 90.22, 71.11, 161.33, 91.11]);
  const b = run({ coinsurance: 'Before deductible' });
  assert.deepEqual([b.deductible, b.coinsurance, b.patient, b.deductibleAfter], [0, 90.22, 161.33, 0.89], 'coinsurance on the full allowed lowers the deductible left');
  const n = run({ coinsurance: 'None' });
  assert.deepEqual([n.deductible, n.coinsurance, n.patient], [91.11, 0, 162.22]);
  const c = run({ copay: 'None' });
  assert.deepEqual([c.copay, c.patient], [0, 163.11]);
});

test('math: caps at OOP remaining and at the allowed amount, never below 0, cents, coinsurance as 0.2 / 20 / 20%', () => {
  const T = { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' };
  const oop = F.estimateMath(L(451.11), B({ oopRemaining: '171.11' }), T);
  assert.deepEqual([oop.subtotal, oop.patient, oop.cap, oop.plan], [234.22, 171.11, 'oop', 280]);
  const met = F.estimateMath(L(451.11), B({ oopRemaining: '0' }), T);
  assert.equal(met.patient, 0);
  const big = F.estimateMath(L(51.11), B({ deductibleRemaining: '0', copay: '91.11' }), T);
  assert.deepEqual([big.patient, big.cap, big.plan], [51.11, 'allowed', 0]);
  const neg = F.estimateMath(L(171.11), B({ deductibleRemaining: '-51.11', coinsurance: '0', copay: '0' }), T);
  assert.equal(neg.patient, 0);
  for (const coins of ['0.2', '20', '20%', 0.2]) assert.equal(F.estimateMath(L(171.11), B({ deductibleRemaining: '0', copay: '' , coinsurance: coins }), { ...T, copay: 'None' }).patient, 34.22, `coinsurance ${coins}`);
  const odd = F.estimateMath(L(251.1133), B({ deductibleRemaining: '0', coinsurance: '15%', copay: '0' }), T);
  assert.deepEqual([odd.allowed, odd.coinsurance, odd.patient], [251.11, 37.67, 37.67], '15% of 251.11 = 37.6665 rounds up');
});

test('math: several codes are summed before cost sharing; a code with no rate is left out and named', () => {
  const lines = [{ code: '90868', units: 2, allowed: 151.11 }, { code: '90869', units: 1, allowed: 51.11 }, { code: '99999', units: 1, allowed: null }];
  const e = F.estimateMath(lines, B({ deductibleRemaining: '0' }), { deductible: 'Applies', coinsurance: 'After deductible', copay: 'None' });
  assert.deepEqual([e.allowed, e.coinsurance, e.patient], [353.33, 70.67, 70.67]);
  assert.deepEqual(e.missing, ['No rate for 99999']);
});

test('math: missing benefits are named and the estimate uses what is known', () => {
  const e = F.estimateMath(L(451.11), { coinsurance: '20%' }, { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' });
  assert.deepEqual(e.missing, ['Deductible remaining', 'Copay', 'OOP max remaining (no cap applied)']);
  assert.equal(e.patient, 90.22);
  const text = F.estimateText(e);
  assert.ok(text.some((t) => /^Missing: Deductible remaining; Copay; OOP max remaining/.test(t)));
  assert.ok(text.some((t) => t === 'Patient pays: $90.22 · Plan pays: $360.89'));
  // Nothing at all from Sohar: no patient amount rather than a misleading $0.
  const none = F.estimateMath(L(451.11), {}, { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' });
  assert.deepEqual([none.allowed, none.patient, none.plan, none.noBenefits], [451.11, null, null, true]);
  assert.match(F.estimateText(none).join(' '), /No deductible, coinsurance or copay from Sohar/);
});

// ---------- per patient ----------
const mk = (o) => ({ ...Object.fromEntries(C.RESULTS_COLS.map((c) => [c, ''])), 'Sent As': 'Verification', 'Sohar Status': 'complete.eligible', ...o });
const RESULTS = [
  mk({ 'Per Nbr': '1', 'Last Name': 'Alpha', 'Next Appt': '11/02/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Harbor Point Clinic', 'Next Provider': 'Doe, Pat MD',
    'Sohar Payer': 'Blue Shield of California', 'NextGen Payer': 'BS CA', Copay: '0', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '3091.11' }),
  mk({ 'Per Nbr': '2', 'Last Name': 'Beta', 'Next Appt': '11/03/2026', 'Next Event': 'Telepsych Follow Up', 'Next Location': 'Telehealth', 'Next Provider': 'Doe, Pat MD',
    'Sohar Payer': 'Acme Health Group', Copay: '51.11', Coinsurance: '0', 'Deductible Remaining': '0', 'OOP Remaining': '591.11' }),
  mk({ 'Per Nbr': '3', 'Last Name': 'Gamma', 'Next Appt': '11/04/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Faraway', 'Sohar Payer': 'Blue Shield of California' }),
  Object.assign(mk({ 'Per Nbr': '4', 'Last Name': 'Delta', 'Next Appt': '11/05/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Mill Street', 'Sohar Payer': 'Unknown Synthetic Payer', Network: 'OON' }),
    { oon: { 'Deductible Remaining': '91.11', Coinsurance: '40%', Copay: '', 'OOP Remaining': '', outOfNetwork: true } }),
  mk({ 'Per Nbr': '5', 'Last Name': 'Epsilon', 'Next Appt': '11/06/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Ridge', 'Sohar Payer': 'Acme Health Medicare Advantage' }),
  mk({ 'Per Nbr': '6', 'Last Name': 'Zeta', 'Next Appt': '11/07/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Harbor Point', 'Sohar Status': 'complete.ineligible', 'Sohar Payer': 'Old Payer', 'NextGen Payer': 'Kaiser Synthetic' }),
  mk({ 'Per Nbr': '7', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found' }),
];
const APPTS = [{ rendering: 'Doe, Pat MD', loc: 'Mill Street', event: 'TMS' }, { rendering: 'Doe, Pat MD', loc: 'Mill Street', event: 'TMS' }, { rendering: 'Doe, Pat MD', loc: 'Telehealth', event: 'Telepsych' }];
// The fixtures estimate "TMS Treatment" as a stand-in visit: an Estimate Codes sheet staff edited (the v1 rows), which
// the seed upgrade leaves alone.
const OLD_CODES = F.ESTIMATE_CODE_SEEDS_V1.map((r) => ({ ...r, Source: 'manual' }));
const prep = (o) => F.prepareEstimates({ results: RESULTS, scheduleRows: F.parseFeeSchedule(WB).rows, appts: APPTS, teleKeywords: ['Tele'], cfgDefaults: F.TOGGLE_DEFAULTS, when: 't', codes: OLD_CODES, ...o });

test('prepareEstimates: tab by location, telehealth by home clinic, payer column, codes; flags instead of guesses', () => {
  const p = prep();
  assert.deepEqual(Object.keys(p.contexts).sort(), ['1', '2', '3', '4', '5', '6'], 'verified patients only');
  assert.equal(p.seeded, false, 'edited Estimate Codes are kept');
  assert.equal(p.codeRows, OLD_CODES);
  const c = p.contexts;
  assert.deepEqual([c[1].tab, c[1].feePayer, c[1].payerSource, c[1].codes[0].code], ['Bayview', 'Blue Shield', 'Sohar', '90868']);
  assert.deepEqual([c[2].telehealth, c[2].home, c[2].tab, c[2].feePayer, c[2].codes[0].code], [true, 'Mill Street', 'Bayview', 'Acme Health', '99214']);
  assert.equal(c[3].ok, false);
  assert.ok(c[3].flags.includes('No fee tab for location Faraway'));
  assert.equal(c[4].oon, true);
  assert.equal(c[4].ok, true, 'out of network needs no payer column');
  assert.equal(c[5].ok, false);
  assert.ok(c[5].flags.some((f) => /^Map payer to a fee column: Acme Health Medicare Advantage on Ridge/.test(f)));
  assert.deepEqual([c[6].payer, c[6].payerSource, c[6].feePayer], ['Kaiser Synthetic', 'NextGen', 'Kaiser'], 'not eligible: the NextGen payer');
  assert.deepEqual(p.counts, { patients: 6, ready: 4, flagged: 2, labeled: 0, tabs: 2 });
  assert.ok(p.locationRows.some((r) => r['NextGen Location'] === 'Faraway' && r['Fee Tab'] === ''));
  assert.ok(p.homeRows.some((r) => r.Provider === 'Doe, Pat MD' && r['Home Location'] === 'Mill Street'));
  assert.equal(p.logRows.length, p.payerRows.length);
  assert.deepEqual(Object.keys(p.tables.rates).sort(), ['Bayview|Acme Health', 'Bayview|Blue Shield', 'Bayview|Kaiser'], 'only the rate tables in use');
});

test('estimatePatient: MD and NP, copay default only when Sohar returned one, OON uses Gross Charge, single-column payer has no NP', () => {
  const p = prep();
  const e1 = F.estimatePatient(p.contexts[1], p.tables);
  assert.equal(p.contexts[1].defaults.copay, 'None', 'Sohar copay $0: copay not added by default');
  assert.deepEqual([e1.md.allowed, e1.md.patient, e1.np.allowed, e1.np.patient], [271.11, 54.22, 251.11, 50.22]);
  assert.match(e1.basis, /^Bayview \(Locality 9\) · Blue Shield$/);
  const e2 = F.estimatePatient(p.contexts[2], p.tables);
  assert.equal(p.contexts[2].defaults.copay, 'Added');
  assert.deepEqual([e2.md.patient, e2.np.patient], [51.11, 51.11]);
  const e4 = F.estimatePatient(p.contexts[4], p.tables);
  assert.deepEqual([e4.md.allowed, e4.md.deductible, e4.md.coinsurance, e4.md.patient], [551.11, 91.11, 184, 275.11], 'Gross Charge, OON deductible 91.11, 40% of the rest');
  assert.ok(e4.md.missing.includes('OOP max remaining (no cap applied)'));
  const e6 = F.estimatePatient(p.contexts[6], p.tables);
  assert.equal(e6.md.allowed, 171.11);
  assert.equal(e6.np.patient, null);
  assert.match(e6.np.none, /^No NP rate for Kaiser on Bayview/);
  const e3 = F.estimatePatient(p.contexts[3], p.tables);
  assert.deepEqual([e3.md, e3.np], [null, null]);
  // Staff edits in the panel: other codes and units, toggles.
  const ed = F.estimatePatient(p.contexts[1], p.tables, { codes: [{ code: '90868', units: 1 }, { code: '99417', units: 2 }], toggles: { coinsurance: 'None', copay: 'None' } });
  assert.deepEqual([ed.md.allowed, ed.md.patient], [373.33, 0], '271.11 + 2 x 51.11');
});

test('defaults remembered per fee payer win over Config; manual mapping edits feed the estimate', () => {
  const defaults = F.rememberDefaults([], 'Blue Shield', { deductible: 'Waived', coinsurance: 'Before deductible', copay: 'Added' }, 'now');
  assert.deepEqual(defaults, [{ 'Fee Payer': 'Blue Shield', Deductible: 'Waived', Coinsurance: 'Before deductible', Copay: 'Added', Updated: 'now' }]);
  const p = prep({ defaults });
  assert.deepEqual(p.contexts[1].defaults, { deductible: 'Waived', coinsurance: 'Before deductible', copay: 'Added' }, 'remembered: copay stays Added');
  const locations = F.feeLocationRows([], ['Faraway'], [], ['Tele']).map((r) => ({ ...r, 'Fee Tab': 'Ridge' }));
  const payerMap = [{ Payer: 'Blue Shield of California', 'Fee Tab': 'Ridge', 'Fee Payer': 'Acme Health', Source: 'manual', 'Auto Match': '' }];
  const q = prep({ locations, payerMap });
  assert.deepEqual([q.contexts[3].tab, q.contexts[3].feePayer, q.contexts[3].ok], ['Ridge', 'Acme Health', true]);
  const bad = prep({ locations: locations.map((r) => ({ ...r, 'Fee Tab': 'Nope' })) });
  assert.ok(bad.contexts[3].flags.some((f) => /Fee tab "Nope" .* is not in the uploaded fee schedule/.test(f)));
  const codes = [{ 'Event keyword': 'TMS', Codes: '90868 x1; 99417 x1', Source: '', Note: '' }];
  const r = prep({ codes });
  assert.equal(r.seeded, false);
  assert.deepEqual(r.contexts[1].codes, [{ code: '90868', units: 1 }, { code: '99417', units: 1 }]);
});

test('fee status: file, date, counts; warns after 90 days', () => {
  const parsed = F.parseFeeSchedule(WB);
  const meta = F.feeMetaRows('Synthetic Fees.xlsx', '2026-06-01T12:00:00.000Z', parsed);
  const now = new Date('2026-07-01T12:00:00Z');
  const s = F.feeStatus(meta, parsed.rows.length, now);
  assert.equal(s.text, 'Fee schedule: Synthetic Fees.xlsx, uploaded 6/1/2026, 2 tabs, 5 codes');
  assert.equal(s.stale, false);
  const old = F.feeStatus(meta, parsed.rows.length, new Date('2026-10-20T12:00:00Z'));
  assert.equal(old.stale, true);
  assert.match(old.warn, /141 days old/);
  assert.equal(F.feeStatus([], 0).loaded, false);
  assert.deepEqual(F.readMeta(meta).tabSummary.map((t) => t.name), ['Bayview', 'Ridge']);
});

test('demo schedule is synthetic and parses', () => {
  const s = F.demoSchedule();
  assert.ok(s.counts.tabs >= 4);
  assert.ok(s.skipped.some((x) => x.reason === 'hidden tab'));
  assert.match(JSON.stringify(F.DEMO_WORKBOOK), /made-up/);
});

test('core: out-of-network benefits ride on a non-column property', () => {
  const o = C.oonBenefitFields({ ooncopay: '51.11', ooncoinsurance: '0.4', oontotaldeductible: '3051.11', oonremainingdeductible: '2551.11', oonoutofpocket: '9051.11', oonoutofpocketremaining: '8051.11', networktype: 'OON' });
  assert.deepEqual(o, { Copay: '51.11', Coinsurance: '40%', Deductible: '3051.11', 'Deductible Remaining': '2551.11', 'OOP Max': '9051.11', 'OOP Remaining': '8051.11', outOfNetwork: true });
  assert.equal(C.oonBenefitFields({ networktype: 'PPO' }).outOfNetwork, false);
  assert.equal(C.oonBenefitFields({ innetwork: 'false' }).outOfNetwork, true);
});

// ---------- sheets ----------
test('fee sheets: the schedule and its info are hidden, mapping sheets visible; every save replaces the sheet', async () => {
  const m = createMock(), XL = m.SoharExcel;
  assert.deepEqual(await XL.readFee('schedule'), []);
  const rows = F.parseFeeSchedule(WB).rows.slice(0, 5).map((r) => ({ ...r, Rate: String(r.Rate) }));
  await XL.writeFee('schedule', F.FEE_COLS, rows);
  await XL.writeFee('meta', F.FEE_META_COLS, [{ Key: 'File', Value: 'x.xlsx' }]);
  await XL.writeFee('locations', F.FEE_LOCATION_COLS, [{ 'NextGen Location': 'Harbor Point', 'Fee Tab': 'Bayview', Source: 'auto', 'Auto Match': 'Bayview', Note: '' }]);
  assert.equal(m.sheetVisibility('Fee Schedule'), 'Hidden');
  assert.equal(m.sheetVisibility('Fee Schedule Info'), 'Hidden');
  assert.equal(m.sheetVisibility('Fee Locations'), 'Visible');
  assert.equal((await XL.readFee('schedule')).length, 5);
  await XL.writeFee('schedule', F.FEE_COLS, rows.slice(0, 2));
  assert.equal((await XL.readFee('schedule')).length, 2, 're-upload replaces it completely');
  assert.deepEqual(Object.keys(XL.FEE_TABLES).sort(), ['codes', 'defaults', 'homes', 'locations', 'meta', 'payerLog', 'payerMap', 'schedule']);
  await assert.rejects(XL.readFee('nope'), /Unknown fee table/);
});

test('Off means v0.10 behavior in the pane: Benefits columns, payload, Config rows and the fee section are gated', () => {
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  assert.match(app, /const feesOn = \(\) => !!\(F && state\.cfg && state\.cfg\.fees && state\.cfg\.fees\.on\);/);
  assert.match(app, /const benefitCols = \(\) => C\.BENEFIT_SHEET_COLS\.concat\('Pills', \.\.\.\(feesOn\(\) \? \['Est MD', 'Est NP'\] : \[\]\)\);/);
  assert.match(app, /Pills: pillsFor\(C\.keyOf\(r\)\), \.\.\.\(feesOn\(\) \? estCols\(r\['Per Nbr'\]\) : \{\}\)/);
  assert.match(app, /if \(feesOn\(\)\) base\.fees = feePayload\(\);/);
  assert.match(app, /if \(feesOn\(\)\) await refreshFees\(results\)/);
  assert.match(app, /feesOn\(\) \? F\.FEE_CONFIG_ROWS : F\.FEE_CONFIG_ROWS\.slice\(0, 1\)/);
  // Every X.readFee / X.writeFee call sits in a function that only runs when fees are On.
  const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
  assert.match(html, /<details class="more" id="feeSection" hidden>/);
  for (const f of ['index.html', 'dashboard.html']) {
    const h = readFileSync(new URL(`../${f}`, import.meta.url), 'utf8');
    assert.match(h, /<script src="core\.js\?v=[\d.]+"><\/script>\s*<script src="rules\.js\?v=[\d.]+"><\/script>\s*<script src="fees\.js\?v=[\d.]+"><\/script>/, `${f} loads fees.js after core.js`);
  }
});

// ---------- review fixes (synthetic data only) ----------
// A tab with one "MD/NP" column, the same payer header twice, and an NP column whose payer header is blank.
const LAKESHORE = {
  name: 'Lakeshore', hidden: false,
  rows: [
    ['', 'Locations: Lake Town'],
    ['Lakeshore Fee Schedule 2026 Locality 3'],
    ['', 'CPT Code', 'Description', 'Gross Charge', 'Cash', 'Acme Health', '', 'Beacon/Carelon', '', 'Beacon/Carelon', '', 'Kaiser', '', 'Imperial', '', ''],
    ['', '', '', '', 'Rate', 'MD', 'Nurse', 'MD', 'Nurse', 'MD', 'NP', 'MD/NP', '', 'MD', '', 'NP'],
    ['', '90868', 'TMS subsequent delivery', 451.11, 291.11, 251.11, 191.11, 191.11, 171.11, 171.11, 151.11, 271.11, '', 191.11, '', 151.11],
    ['', '90869', 'Motor threshold', 391.11, 251.11, 'NA', 'NA', 171.11, 151.11, 171.11, 151.11, 'NA', '', 'NA', '', 151.11],
  ],
  merges: [],
};
const LWB = { sheets: [LAKESHORE] };

test('fix: an "MD/NP" column is one rate for both levels, so NP is estimated too', () => {
  const p = F.parseFeeSchedule(LWB);
  const t = p.tabs[0];
  assert.deepEqual(t.payers, ['Acme Health', 'Beacon/Carelon', 'Beacon/Carelon (2)', 'Kaiser', 'Imperial']);
  assert.equal(rateOf(p.rows, 'Lakeshore', '90868', 'Kaiser', 'MD').Rate, 271.11);
  assert.equal(rateOf(p.rows, 'Lakeshore', '90868', 'Kaiser', 'NP').Rate, 271.11);
  assert.equal(t.unparsed.na, 4, 'an NA in an MD/NP column counts once');
  const res = [mk({ 'Per Nbr': '1', 'Next Appt': '11/02/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Lake Town', 'Sohar Payer': 'Kaiser Foundation Health Plan',
    Copay: '0', Coinsurance: '10%', 'Deductible Remaining': '0', 'OOP Remaining': '3091.11' })];
  const q = F.prepareEstimates({ results: res, scheduleRows: p.rows, appts: [], teleKeywords: ['Tele'], cfgDefaults: F.TOGGLE_DEFAULTS, when: 't', codes: OLD_CODES });
  const e = F.estimatePatient(q.contexts[1], q.tables);
  assert.equal(q.contexts[1].feePayer, 'Kaiser');
  assert.deepEqual([e.md.allowed, e.np.allowed, e.np.none], [271.11, 271.11, undefined]);
  // Two codes, one with no rate: the total is partial and says so.
  const two = F.estimatePatient(q.contexts[1], q.tables, { codes: [{ code: '90868', units: 1 }, { code: '90869', units: 1 }] });
  assert.equal(two.md.partial, true);
  assert.equal(F.partialNote(two.md), 'partial: no rate for 90869');
});

test('fix: a rate column with no payer name is reported in the tab summary, the upload note and the status', () => {
  const p = F.parseFeeSchedule(LWB);
  assert.deepEqual(p.tabs[0].unlabeledCols, [{ col: 'P', level: 'NP', why: 'no payer name' }]);
  assert.deepEqual(F.parseFeeSchedule(WB).tabs[0].unlabeledCols, [{ col: 'V', level: 'MD', why: 'no payer name' }]);
  const w = F.parseWarnings(p.tabs);
  assert.deepEqual(w, ['Lakeshore: 1 rate column skipped (P NP, no payer name)', '4 rate cells not read as rates (0, NA, errors, percentages or text)']);
  const meta = F.feeMetaRows('Synthetic.xlsx', '2026-10-01T12:00:00.000Z', p);
  assert.deepEqual(F.readMeta(meta).tabSummary[0].unlabeledCols, p.tabs[0].unlabeledCols);
  assert.deepEqual(F.feeStatus(meta, p.rows.length, new Date('2026-10-20T12:00:00Z')).issues, w);
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  assert.match(app, /const issues = F\.parseWarnings\(parsed\.tabs\);/);
  assert.match(app, /for \(const i of st\.issues \|\| \[\]\)/);
});

test('fix: a second column with the same payer header ties with the first instead of losing to it', () => {
  const payers = F.parseFeeSchedule(LWB).tabs[0].payers;
  const r = F.matchFeePayer('Carelon Behavioral Health', payers);
  assert.deepEqual([r.feePayer, r.confidence], ['', 'Low']);
  assert.match(r.note, /matches Beacon\/Carelon and Beacon\/Carelon \(2\); pick one/);
  assert.deepEqual(F.matchFeePayer('Acme Health Plan', payers).feePayer, 'Acme Health');
});

test('fix: location matching never matches inside a word', () => {
  const tabs = [{ name: 'Marin', locations: ['San Rafael'] }, { name: 'Orange', locations: [] }, { name: 'ElDorado', locations: [] }, { name: 'San Joaquin', locations: [] }];
  const m = (l) => F.matchLocationTab(l, tabs).tab;
  assert.equal(m('San Marino'), '');
  assert.equal(m('San Marino Clinic'), '');
  assert.equal(m('Marina del Rey'), '');
  assert.equal(m('Orangevale'), '');
  assert.equal(m('San Rafael Clinic'), 'Marin');
  assert.equal(m('Marin Clinic'), 'Marin');
  assert.equal(m('Orange County'), 'Orange');
  assert.equal(m('El Dorado Hills'), 'ElDorado', 'words run together still match');
  assert.equal(m('SanJoaquin Clinic'), 'San Joaquin');
});

test('fix: OOP max already met with no other benefits from Sohar is $0, not "no estimate"', () => {
  const e = F.estimateMath(L(151.11), { oopRemaining: '0' });
  assert.deepEqual([e.patient, e.plan, e.cap, e.noBenefits], [0, 151.11, 'oop', false]);
  assert.ok(F.estimateText(e).includes('Patient pays: $0.00 · Plan pays: $151.11'));
  assert.equal(F.estimateMath(L(151.11), { oopRemaining: '591.11' }).noBenefits, true, 'OOP left and nothing else: still no estimate');
});

test('fix: a code with no rate makes the total partial; the pane card does not call it a benefits gap', () => {
  const lines = [{ code: '90867', units: 1, allowed: 171.11 }, { code: '90869', units: 1, allowed: null }];
  const e = F.estimateMath(lines, B({ deductibleRemaining: '0', copay: '0' }), { deductible: 'Applies', coinsurance: 'After deductible', copay: 'Added' });
  assert.deepEqual([e.patient, e.partial, e.noRate], [34.22, true, ['90869']]);
  assert.equal(F.partialNote(e), 'partial: no rate for 90869');
  assert.equal(F.benefitsMissing(e), false);
  assert.ok(F.estimateText(e).some((t) => /^Partial: no rate for 90869/.test(t)));
  assert.equal(F.partialNote(F.estimateMath(L(151.11), B())), '', 'every code has a rate');
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  assert.match(app, /const estAmount = \(e\) => \(e && e\.patient != null \? `\$\{e\.patient\.toFixed\(2\)\}\$\{F\.partialNote\(e\)/);
  assert.match(app, /F\.benefitsMissing\(e\.md\) \|\| F\.benefitsMissing\(e\.np\)/);
  assert.match(readFileSync(new URL('../dashboard.js', import.meta.url), 'utf8'), /F\.partialNote\(est\)/);
});

test('fix: changing one toggle remembers only that toggle; a $0-copay patient never saves Copay None for the fee payer', () => {
  const d1 = F.rememberDefaults([], 'Blue Shield', { deductible: 'Waived' }, 't1');
  assert.deepEqual(d1, [{ 'Fee Payer': 'Blue Shield', Deductible: 'Waived', Coinsurance: '', Copay: '', Updated: 't1' }]);
  const d2 = F.rememberDefaults(d1, 'Blue Shield', { coinsurance: 'None' }, 't2');
  assert.deepEqual(d2[0], { 'Fee Payer': 'Blue Shield', Deductible: 'Waived', Coinsurance: 'None', Copay: '', Updated: 't2' }, 'merged into the row');
  const res = RESULTS.concat([mk({ 'Per Nbr': '8', 'Next Appt': '11/09/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Harbor Point', 'Sohar Payer': 'Blue Shield of California',
    Copay: '71.11', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '3091.11' })]);
  const p = prep({ results: res, defaults: d1 });
  assert.deepEqual(p.contexts[1].defaults, { deductible: 'Waived', coinsurance: 'After deductible', copay: 'None' }, '$0 copay: not added for this patient');
  assert.deepEqual(p.contexts[8].defaults, { deductible: 'Waived', coinsurance: 'After deductible', copay: 'Added' }, 'another patient on the fee payer keeps the $71.11 copay');
  assert.equal(F.estimatePatient(p.contexts[8], p.tables).md.copay, 71.11);
  // Staff pick copay None on purpose: that is remembered for the fee payer.
  const d3 = F.rememberDefaults(d1, 'Blue Shield', { copay: 'None' }, 't3');
  assert.equal(prep({ results: res, defaults: d3 }).contexts[8].defaults.copay, 'None');
  const dash = readFileSync(new URL('../dashboard.js', import.meta.url), 'utf8');
  assert.match(dash, /toggles: \{ \[s\.dataset\.etoggle\]: s\.value \}/, 'the full-screen panel sends only the toggle changed');
});

test('fix: the panel says toggles are not remembered when no fee payer is mapped (out of network)', () => {
  const dash = readFileSync(new URL('../dashboard.js', import.meta.url), 'utf8');
  assert.match(dash, /no fee payer column is mapped for this patient, so they are not remembered/);
  assert.equal(prep().contexts[4].feePayer, '', 'the OON synthetic patient has no fee payer');
});

test('fix: a manual Fee Payer Map row with a blank Fee Tab applies on every tab, through prepareEstimates', () => {
  const payerMap = [{ Payer: 'Blue Shield of California', 'Fee Tab': '', 'Fee Payer': 'Acme Health', Confidence: '', Source: 'manual', 'Auto Match': '', Note: '' }];
  const p = prep({ payerMap });
  assert.equal(p.contexts[1].feePayer, 'Acme Health');
  assert.ok(!p.payerRows.some((r) => r.Payer === 'Blue Shield of California' && r['Fee Tab']), 'no auto row competes with the every-tab row');
  // An auto row already on the sheet from before the manual row was added also loses.
  const auto = { Payer: 'Blue Shield of California', 'Fee Tab': 'Bayview', 'Fee Payer': 'Blue Shield', Confidence: 'High', Source: 'auto', 'Auto Match': 'Blue Shield', Note: '' };
  assert.equal(F.lookupFeePayer([auto, ...payerMap], 'Blue Shield of California', 'Bayview')['Fee Payer'], 'Acme Health');
  assert.equal(prep({ payerMap: [auto, ...payerMap] }).contexts[1].feePayer, 'Acme Health');
  // A manual row for this exact tab still beats the every-tab row.
  const exact = { ...auto, 'Fee Payer': 'Kaiser', Source: 'manual' };
  assert.equal(F.lookupFeePayer([exact, ...payerMap], 'Blue Shield of California', 'Bayview')['Fee Payer'], 'Kaiser');
});

test('fix: a home clinic typed on Provider Home Clinic gets a fee tab even when no visit in the report is there', () => {
  const res = RESULTS.concat([mk({ 'Per Nbr': '9', 'Next Appt': '11/10/2026', 'Next Event': 'Telepsych Follow Up', 'Next Location': 'Telehealth', 'Next Provider': 'Roe, Sam NP',
    'Sohar Payer': 'Blue Shield of California', Copay: '0', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '3091.11' })]);
  const appts = APPTS.concat([{ rendering: 'Roe, Sam NP', loc: 'Telehealth', event: 'Telepsych' }]);
  const homes = [{ Provider: 'Roe, Sam NP', 'Home Location': 'Harbor Point', Source: 'manual', 'Auto Match': '', Note: '' }];
  const p = prep({ results: res, appts, homes });
  const c = p.contexts[9];
  assert.deepEqual([c.telehealth, c.home, c.tab, c.feePayer], [true, 'Harbor Point', 'Bayview', 'Blue Shield']);
  assert.ok(!c.flags.some((f) => /No fee tab/.test(f)));
  assert.ok(p.locationRows.some((r) => r['NextGen Location'] === 'Harbor Point' && r['Fee Tab'] === 'Bayview'), 'the home clinic is on Fee Locations, where staff can fix it');
});

test('fix: a refresh re-reads the editable fee sheets, so a sheet edit made after the first read is not overwritten', async () => {
  const m = createMock(), XL = m.SoharExcel;
  const t = TABS();
  await XL.writeFee('locations', F.FEE_LOCATION_COLS, F.feeLocationRows([], ['Harbor Point', 'Faraway'], t, ['Tele']));
  const cached = await XL.readFee('locations'); // the pane's copy from its first read
  // Staff type a fee tab for Faraway on the sheet after that read.
  await XL.writeFee('locations', F.FEE_LOCATION_COLS, cached.map((r) => (r['NextGen Location'] === 'Faraway' ? { ...r, 'Fee Tab': 'Ridge' } : r)));
  // A refresh that adds a new location writes the sheet back: from a fresh read the edit survives.
  const fresh = await XL.readFee('locations');
  await XL.writeFee('locations', F.FEE_LOCATION_COLS, F.feeLocationRows(fresh, ['Mill Street'], t, ['Tele']));
  const by = Object.fromEntries((await XL.readFee('locations')).map((r) => [r['NextGen Location'], r]));
  assert.deepEqual([by.Faraway['Fee Tab'], by.Faraway.Source], ['Ridge', 'manual']);
  assert.equal(by['Mill Street']['Fee Tab'], 'Bayview');
  assert.equal(F.feeLocationRows(cached, ['Mill Street'], t, ['Tele']).find((r) => r['NextGen Location'] === 'Faraway')['Fee Tab'], '', 'the stale copy would have lost it');
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  assert.match(app, /const FEE_EDITABLE = \['locations', 'homes', 'payerMap', 'payerLog', 'codes', 'defaults'\];/);
  assert.match(app, /if \(state\.fees && state\.fees\.loaded\) \{\s*for \(const k of FEE_EDITABLE\) state\.fees\[k\] = await X\.readFee\(k\);/);
});

test('fix: schedule and info saved from different uploads are flagged, and the old tab summary is not used', () => {
  const a = F.parseFeeSchedule(WB), b = F.parseFeeSchedule(LWB);
  const oldMeta = F.feeMetaRows('Old.xlsx', '2026-10-01T12:00:00.000Z', a);
  assert.equal(F.metaMatches(oldMeta, a.rows.length), true);
  assert.equal(F.metaMatches(oldMeta, b.rows.length), false, 'new rows, old info');
  const s = F.feeStatus(oldMeta, b.rows.length, new Date('2026-10-02T12:00:00Z'));
  assert.equal(s.mismatch, true);
  assert.match(s.warn, /does not match the saved fee schedule rows.*Upload the fee schedule again/);
  assert.equal(F.metaMatches(oldMeta.filter((r) => r.Key !== 'Rows'), 1), true, 'info saved before the row count existed is trusted');
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  assert.match(app, /tabSummary: F\.metaMatches\(f\.meta, f\.schedule\.length\) \? F\.readMeta\(f\.meta\)\.tabSummary : \[\]/);
  assert.match(app, /the fee schedule rates were replaced, but its info \(file name, upload date, tab summary\) was not saved/);
});

// v0.15: regional payer matching. Invented tabs, payers and rates; regions come from Config Location to State.
test('v0.15: payerState, tabRegions and productOf', () => {
  assert.equal(F.payerState('Blue Cross Blue Shield of Texas'), 'TX');
  assert.equal(F.payerState('Blue Shield of California'), 'CA');
  assert.equal(F.payerState('Acme Health of West Virginia'), 'WV', 'longest state name wins');
  assert.equal(F.payerState('Acme Health'), '');
  const locs = [{ 'NextGen Location': 'Harbor Point', 'Fee Tab': 'Bayview' }, { 'NextGen Location': 'Mill Street', 'Fee Tab': 'Bayview' },
    { 'NextGen Location': 'Lone Ridge', 'Fee Tab': 'Ridge' }, { 'NextGen Location': 'Faraway', 'Fee Tab': '' }];
  const rg = F.tabRegions(locs, { 'harbor point': 'ca', 'Lone Ridge': 'TX' });
  assert.deepEqual(rg.get('Bayview'), { states: ['CA'], missing: ['Mill Street'] });
  assert.deepEqual(rg.get('Ridge'), { states: ['TX'], missing: [] });
  assert.equal(rg.has(''), false);
  assert.equal(F.productOf('Acme BS MHSA'), 'mhsa');
  assert.equal(F.productOf('Acme Medicare Advantage'), 'medicare');
  assert.equal(F.productOf('Acme Health PPO'), '');
});

test('v0.15: matchFeePayerSohar blocks another state, keeps the product line, falls back to Sohar and group names', () => {
  const cols = ['Blue Shield', 'BS MHSA', 'Acme Health', 'Acme Medicare'];
  const ca = { states: ['CA'], missing: [] };
  const tx = F.matchFeePayerSohar({ payer: 'Acme Health of Texas', soharName: 'Acme Health of Texas' }, cols, ca);
  assert.equal(tx.feePayer, '');
  assert.match(tx.note, /TX plan; this tab is CA/);
  assert.equal(F.matchFeePayerSohar({ payer: 'Acme BS MHSA', soharName: 'Blue Shield of California', nextgen: 'Acme BS MHSA' }, cols, ca).feePayer, 'BS MHSA');
  // NextGen name unknown to the fee schedule: Sohar's group payer name finds the column.
  const g = F.matchFeePayerSohar({ payer: 'ZZ Unknown Plan', soharName: 'ZZ Unknown Plan', groupName: 'Acme Health' }, cols, ca);
  assert.equal(g.feePayer, 'Acme Health');
  assert.match(g.note, /matched on "Acme Health"/);
  // A Medicare plan never takes a commercial column, whatever name matched.
  assert.equal(F.matchFeePayerSohar({ payer: 'Acme Health', plan: 'Acme Medicare Advantage HMO' }, cols, ca).feePayer, 'Acme Medicare');
  // Different coverage than NextGen: the NextGen name is not used.
  assert.equal(F.matchFeePayerSohar({ payer: 'ZZ Other', nextgen: 'Acme Health', coverageDiffers: true }, cols, ca).feePayer, '');
  // A tab with no Config region: a match is only Medium, with the fix in the note.
  const none = F.matchFeePayerSohar({ payer: 'Acme Health' }, cols, { states: [], missing: ['Mill Street'] });
  assert.deepEqual([none.feePayer, none.confidence], ['Acme Health', 'Medium']);
  assert.match(none.note, /Add Mill Street to Config Location to State/);
});

test('v0.15: prepareEstimates keys product lines by the NextGen name, fills Sohar payerId and Region, honors a Sohar payerId manual row', () => {
  const res = [
    mk({ 'Per Nbr': '21', 'Next Appt': '11/02/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Harbor Point Clinic', 'Next Provider': 'Doe, Pat MD',
      'Sohar Payer': 'Blue Shield of California', 'Sohar payerId': 'SBX1', 'NextGen Payer': 'BS 65 Senior', Copay: '0', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '100' }),
    mk({ 'Per Nbr': '22', 'Next Appt': '11/02/2026', 'Next Event': 'TMS Treatment', 'Next Location': 'Harbor Point Clinic', 'Next Provider': 'Doe, Pat MD',
      'Sohar Payer': 'Blue Cross Blue Shield of Texas', 'Sohar payerId': 'SBX2', 'NextGen Payer': 'BCBS TX', Copay: '0', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '100' }),
  ];
  const locationState = { 'Harbor Point Clinic': 'CA' };
  const q = F.prepareEstimates({ results: res, scheduleRows: F.parseFeeSchedule(WB).rows, appts: [], teleKeywords: ['Tele'], cfgDefaults: F.TOGGLE_DEFAULTS, when: 't', codes: OLD_CODES, locationState });
  const c21 = q.contexts[21], c22 = q.contexts[22];
  assert.equal(c21.payerSource, 'NextGen (product line)');
  assert.equal(c21.payer, 'BS 65 Senior');
  assert.equal(c21.region, 'CA');
  assert.equal(c22.ok, false, 'a Texas Blue plan never gets a California rate by name');
  assert.ok(c22.flags.some((f) => /set Config "BlueCard host CA"/.test(f)), 'an out-of-state Blue plan waits for Config BlueCard host');
  const row22 = q.payerRows.find((r) => r.Payer === 'Blue Cross Blue Shield of Texas');
  assert.deepEqual([row22['Sohar payerId'], row22.Region, row22['Fee Payer']], ['SBX2', 'CA', '']);
  assert.deepEqual(Object.keys(row22), F.FEE_PAYER_MAP_COLS, 'rows carry exactly the sheet columns');
  // Staff map the Sohar payer itself (blank Payer): it applies whatever NextGen calls the plan.
  const manual = { Payer: '', 'Sohar payerId': 'SBX2', 'Fee Tab': '', Region: '', 'Fee Payer': 'Kaiser', Confidence: '', Source: 'manual', 'Auto Match': '', Note: '' };
  const q2 = F.prepareEstimates({ results: res, scheduleRows: F.parseFeeSchedule(WB).rows, appts: [], teleKeywords: ['Tele'], cfgDefaults: F.TOGGLE_DEFAULTS, when: 't', codes: OLD_CODES, locationState, payerMap: [manual] });
  assert.equal(q2.contexts[22].feePayer, 'Kaiser');
  assert.ok(q2.payerRows.some((r) => r.Payer === '' && r['Sohar payerId'] === 'SBX2'), 'the manual Sohar payerId row is kept');
});

test('v0.15: Config decides out-of-state Blue plans (BlueCard host) and Medicare plans with no Medicare column', () => {
  const cols = ['Acme Blue', 'Blue Shield', 'Kaiser', 'Acme Health'];
  const ca = { states: ['CA'], missing: [] };
  const il = { payer: 'Blue Cross Blue Shield of Illinois', soharName: 'Blue Cross Blue Shield of Illinois' };
  assert.match(F.matchFeePayerSohar(il, cols, ca, F.feeConfig({})).note, /set Config "BlueCard host CA"/);
  const host = F.feeConfig({ 'BlueCard host CA': 'acme blue' });
  assert.deepEqual(host.blueCardHosts, { CA: 'acme blue' });
  const hit = F.matchFeePayerSohar(il, cols, ca, host);
  assert.deepEqual([hit.feePayer, hit.confidence], ['Acme Blue', 'High']);
  assert.match(F.matchFeePayerSohar(il, cols, ca, F.feeConfig({ 'BlueCard host CA': 'Nope' })).note, /not a column on this tab/);
  // Not a Blue plan: another state is still blocked.
  assert.equal(F.matchFeePayerSohar({ payer: 'Acme Health of Texas', soharName: 'Acme Health of Texas' }, cols, ca, host).feePayer, '');

  const senior = { payer: 'Kaiser', plan: 'Kaiser Senior Advantage HMO' };
  assert.equal(F.matchFeePayerSohar(senior, cols, ca, F.feeConfig({})).feePayer, '', 'default: flagged');
  assert.match(F.matchFeePayerSohar(senior, cols, ca, F.feeConfig({})).note, /only Kaiser/);
  const exc = F.matchFeePayerSohar(senior, cols, ca, F.feeConfig({ 'Medicare payer column exceptions': 'Kaiser; Other' }));
  assert.equal(exc.feePayer, 'Kaiser');
  assert.match(exc.note, /Medicare payer column exceptions/);
  assert.equal(F.matchFeePayerSohar({ payer: 'Acme Health', plan: 'Medicare Advantage' }, cols, ca, F.feeConfig({ 'Medicare payer column exceptions': 'Kaiser' })).feePayer, '', 'exceptions are per column');
  assert.equal(F.matchFeePayerSohar({ payer: 'Acme Health', plan: 'Medicare Advantage' }, cols, ca, F.feeConfig({ 'Medicare without Medicare column': 'Use payer column' })).feePayer, 'Acme Health');
  for (const s of ['BlueCard host CA', 'BlueCard host TX', 'Medicare without Medicare column', 'Medicare payer column exceptions']) assert.ok(F.FEE_CONFIG_ROWS.some((r) => r.setting === s), s);
});

// ---------- v0.16: Estimate Codes by event, interventions, IV ketamine ----------
test('v0.16 seeds: untouched v1 Estimate Codes are upgraded to the current seeds; edited ones are kept', () => {
  const up = prep({ codes: F.ESTIMATE_CODE_SEEDS_V1.map((r) => ({ ...r })) });
  assert.equal(up.seeded, true);
  assert.equal(up.codeRows.length, F.ESTIMATE_CODE_SEEDS.length);
  assert.equal(prep({ codes: F.ESTIMATE_CODE_SEEDS_V1.map((r, i) => (i ? { ...r } : { ...r, Codes: '90867 x2' })) }).seeded, false);
  assert.equal(prep({ codes: [] }).seeded, true);
});

test('v0.16 interventions: labeled "No estimate", not flagged, no Fee Payer Map pair', () => {
  const p = prep({ codes: undefined });   // current seeds: every fixture "TMS Treatment" is an intervention
  assert.deepEqual(p.counts, { patients: 6, ready: 1, flagged: 0, labeled: 5, tabs: 2 });
  const e = F.estimatePatient(p.contexts[1], p.tables);
  assert.deepEqual([e.md, e.np, e.flags], [null, null, ['No estimate: TMS (intervention): no estimate; verify deductible, OOP max and insurance status']]);
  assert.ok(!p.payerRows.some((r) => r.Payer === 'Kaiser Synthetic'), 'interventions add no Fee Payer Map rows');
  assert.equal(F.estimatePatient(p.contexts[2], p.tables).md.patient, 51.11, 'a 30-minute visit is still estimated');
});

const IVK_RESULTS = [
  mk({ 'Per Nbr': '21', 'Next Appt': '11/10/2026', 'Next Event': 'IV Ketamine Follow Up', 'Next Location': 'Harbor Point Clinic', 'Next Provider': 'Doe, Pat MD',
    'Sohar Payer': 'Acme Health Group', Copay: '0', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '5000' }),
  mk({ 'Per Nbr': '22', 'Next Appt': '11/10/2026', 'Next Event': 'IV Ketamine Follow Up', 'Next Location': 'Harbor Point Clinic', 'Sohar Payer': 'Acme Health Medicare Advantage' }),
  mk({ 'Per Nbr': '23', 'Sent As': 'Discovery', 'Sohar Status': 'No coverage found', 'NextGen Payer': 'Self Pay', 'Next Appt': '11/10/2026', 'Next Event': 'IV Ketamine Follow Up', 'Next Location': 'Harbor Point Clinic' }),
  mk({ 'Per Nbr': '24', 'Next Appt': '11/10/2026', 'Next Event': 'IV Ketamine Treatment Start', 'Next Location': 'Harbor Point Clinic', 'Sohar Payer': 'Acme Health Group', Coinsurance: '20%', 'Deductible Remaining': '0' }),
  mk({ 'Per Nbr': '25', 'Next Appt': '11/10/2026', 'Next Event': 'IV Ketamine Follow Up', 'Next Location': 'Harbor Point Clinic', 'Sohar Payer': 'Acme Health Group', Copay: '0', Coinsurance: '20%', 'Deductible Remaining': '0', 'OOP Remaining': '5000' }),
  mk({ 'Per Nbr': '26', 'Sent As': 'Discovery', 'Sohar Status': 'Coverage found', 'NextGen Payer': 'Self Pay', 'Sohar Payer': 'Acme Health Group', 'Next Appt': '11/10/2026', 'Next Event': 'Telepsych' }),
];
const IVK_APPTS = [
  { per: '21', date: '11/01/2026', event: 'IV Ketamine Treatment Start', status: 'Completed' }, { per: '21', date: '11/03/2026', event: 'IV Ketamine Follow Up', status: 'Completed' },
  { per: '21', date: '11/05/2026', event: 'IV Ketamine Follow Up', status: 'Cancelled' }, { per: '21', date: '11/07/2026', event: 'IV Ketamine Follow Up', status: 'Completed' },
  { per: '21', date: '11/12/2026', event: 'IV Ketamine Follow Up', status: '' },
  ...[1, 2, 3, 4, 5, 6, 7].map((d) => ({ per: '23', date: `10/0${d}/2026`, event: 'IV Ketamine Follow Up', status: 'Completed' })),
];
const ivkPrep = () => F.prepareEstimates({ results: IVK_RESULTS, scheduleRows: F.parseFeeSchedule(WB).rows, appts: IVK_APPTS, teleKeywords: ['Tele'],
  cfgDefaults: F.TOGGLE_DEFAULTS, when: 't', feeRules: F.feeConfig({}) });

test('v0.16 IVK: infusion number from the report; 1-6 = $400, 7+ = $450; unknown = confirm', () => {
  assert.deepEqual(F.ivkInfusion('21', '11/10/2026', 'IV Ketamine Follow Up', IVK_APPTS), { n: 4, exact: true }, 'start + 2 kept visits before; cancelled and later ones skipped');
  assert.deepEqual(F.ivkInfusion('23', '11/10/2026', 'IV Ketamine Follow Up', IVK_APPTS), { n: 8, exact: false, atLeast: true });
  assert.deepEqual(F.ivkInfusion('25', '11/10/2026', 'IV Ketamine Follow Up', IVK_APPTS), { n: null, exact: false });
  assert.deepEqual(F.ivkInfusion('9', '11/10/2026', 'IV Ketamine Treatment Start', []), { n: 1, exact: true });
  assert.equal(F.ivkMedication({ n: 6, exact: true }).amount, 400);
  assert.equal(F.ivkMedication({ n: 7, exact: true }).amount, 450);
  assert.match(F.ivkMedication({ n: null }).text, /\$400\.00 for infusions 1-6 or \$450\.00 from 7; confirm the infusion number/);
  assert.equal(F.isIvkInfusion('IV Ketamine Consult Tele'), false);
});

test('v0.16 IVK: commercial = visit through benefits + medication; government = no estimate; self pay = Cash + medication', () => {
  const p = ivkPrep(), c = p.contexts;
  assert.deepEqual(Object.keys(c).sort(), ['21', '22', '23', '24', '25'], 'self-pay IVK joins the verified patients; a self-pay non-IVK patient does not');
  const e21 = F.estimatePatient(c[21], p.tables);
  assert.deepEqual([e21.md.visitPatient, e21.md.patient, e21.np.patient, e21.md.medication.amount], [34.22, 434.22, 430.22, 400], '20% of 171.11 / 151.11, plus $400');
  const t21 = F.estimateText(e21.md);
  assert.ok(t21.includes('Visit: $34.22'), t21.join(' | '));
  assert.ok(t21.includes('IV ketamine medication (self pay), infusion 4: $400.00'), t21.join(' | '));
  assert.ok(t21.includes('Patient pays: $434.22 · Plan pays: $136.89'), t21.join(' | '));
  assert.match(c[22].noEstimate, /^IV ketamine on Acme Health Medicare Advantage \(not self pay or commercial\): no estimate; verify deductible, OOP max and insurance status$/);
  assert.equal(F.estimatePatient(c[22], p.tables).md, null);
  assert.deepEqual([c[23].selfPay, c[23].payer, c[23].ok], [true, 'Self pay', true]);
  const e23 = F.estimatePatient(c[23], p.tables);
  assert.deepEqual([e23.md.allowed, e23.md.patient, e23.np.patient, e23.basis], [151.11, 601.11, 601.11, 'Bayview · Self pay (Cash)'], 'Cash 99214 + $450 (infusion 8 or later)');
  assert.ok(F.estimateText(e23.md).includes('IV ketamine medication (self pay), infusion 8 or later: $450.00'));
  assert.deepEqual([c[24].codes[0].code, c[24].medication.amount], ['99215', 400], 'Treatment Start: 99215, infusion 1');
  const e25 = F.estimatePatient(c[25], p.tables);
  assert.deepEqual([e25.md.patient, e25.md.medication.amount], [34.22, null], 'unknown infusion number: the medication is not in the total');
  assert.equal(p.counts.labeled, 1);
  assert.ok(!p.payerRows.some((r) => /self pay/i.test(r.Payer)), 'self pay needs no fee payer column');
});
