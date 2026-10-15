import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const X = createRequire(import.meta.url)('../xlsx.js');

const dec = new TextDecoder();
const part = (files, name) => {
  const f = files.find(x => x.name === name);
  assert.ok(f, 'missing part ' + name);
  return dec.decode(f.data);
};

// Synthetic data only.
const SPEC = {
  now: new Date(2026, 9, 15, 9, 30, 0),
  sheets: [
    { name: 'Clinic', table: 'tblClinic', columns: ['Per Nbr', 'Last Name', 'Coinsurance', 'Re-verify', 'Notes', 'Notes'],
      pctColumns: ['Coinsurance'], linkColumns: ['Re-verify'], linkText: { 'Re-verify': 'Re-verify in Sohar' },
      rows: [
        { 'Per Nbr': '000123', 'Last Name': 'TESTER & <Co>', Coinsurance: 0.2, 'Re-verify': 'https://app.soharhealth.com/create?type=verification&x=1', Notes: 'a\u0001b' },
        { 'Per Nbr': 45, 'Last Name': 'SAMPLE', Coinsurance: '', 'Re-verify': 'javascript:alert(1)', Notes: 'quote "q"' },
      ] },
    { name: 'Review', table: 'tblClinicReview', columns: ['Severity', 'Issue'], rows: [] },
    { name: 'Clinic Info', table: 'tblClinicMeta', columns: ['Kind', 'Location'], rows: [{ Kind: 'Sohar Clinic', Location: 'Testville' }], hidden: true },
  ],
};

test('crc32 of known strings', () => {
  assert.equal(X.crc32('123456789'), 0xCBF43926);
  assert.equal(X.crc32(''), 0);
  assert.equal(X.crc32('The quick brown fox jumps over the lazy dog'), 0x414FA339);
});

test('zip structure: local headers, central directory, EOCD, CRCs', () => {
  const z = X.zip([{ name: 'a.txt', data: 'hello' }, { name: 'dir/b.txt', data: '123456789' }], new Date(2026, 0, 2, 3, 4, 6));
  const v = new DataView(z.buffer);
  assert.equal(v.getUint32(0, true), 0x04034b50);
  assert.equal(v.getUint16(8, true), 0, 'STORE method');
  const e = z.length - 22;
  assert.equal(v.getUint32(e, true), 0x06054b50);
  assert.equal(v.getUint16(e + 8, true), 2);
  assert.equal(v.getUint16(e + 10, true), 2);
  const cdSize = v.getUint32(e + 12, true), cdOff = v.getUint32(e + 16, true);
  assert.equal(cdOff + cdSize, e);
  const back = X.unzipStored(z);
  assert.deepEqual(back.map(f => f.name), ['a.txt', 'dir/b.txt']);
  assert.equal(dec.decode(back[1].data), '123456789');
  assert.equal(back[1].crc, 0xCBF43926);
  assert.equal(back[0].crc, X.crc32('hello'));
});

test('xml escaping strips control characters', () => {
  assert.equal(X.escapeXml('a & b < c > d "e" \'f\''), 'a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;');
  assert.equal(X.escapeXml('x\u0000\u0007\u000by\tz\n'), 'xy\tz\n');
  assert.equal(X.escapeXml(null), '');
  assert.equal(X.escapeXml(12), '12');
  assert.equal(X.escapeXml('ok\uD800'), 'ok', 'lone surrogate removed');
});

test('column letters beyond Z', () => {
  assert.equal(X.colLetter(1), 'A');
  assert.equal(X.colLetter(26), 'Z');
  assert.equal(X.colLetter(27), 'AA');
  assert.equal(X.colLetter(52), 'AZ');
  assert.equal(X.colLetter(53), 'BA');
  assert.equal(X.colLetter(702), 'ZZ');
  assert.equal(X.colLetter(703), 'AAA');
  assert.equal(X.colLetter(16384), 'XFD');
  assert.throws(() => X.colLetter(0));
  assert.equal(X.cellRef(28, 5), 'AB5');
});

test('table ref math', () => {
  assert.equal(X.tableRef(6, 10), 'A1:F11');
  assert.equal(X.tableRef(6, 0), 'A1:F2', 'empty table keeps one data row');
  assert.equal(X.tableRef(30, 3), 'A1:AD4');
});

test('unique table column names', () => {
  assert.deepEqual(X.uniqueNames(['Notes', 'notes', '', 'Notes', null, 'A\nB']), ['Notes', 'notes 2', 'Column3', 'Notes 3', 'Column5', 'A B']);
});

test('sheet and table names are sanitized and unique', () => {
  const used = new Set();
  assert.equal(X.sanitizeSheetName('A/B:C', used), 'A B C');
  assert.equal(X.sanitizeSheetName('a b c', used), 'a b c (2)');
  assert.equal(X.sanitizeSheetName('x'.repeat(40)).length, 31);
  const t = new Set();
  assert.equal(X.sanitizeTableName('tbl Clinic', t), 'tbl_Clinic');
  assert.equal(X.sanitizeTableName('tbl_Clinic', t), 'tbl_Clinic_2');
  assert.equal(X.sanitizeTableName('A1'), 'tbl_A1');
  assert.equal(X.sanitizeTableName('9lives'), 'tbl_9lives');
});

test('buildWorkbook produces the expected parts', () => {
  const u8 = X.buildWorkbook(SPEC);
  const files = X.unzipStored(u8);
  const names = files.map(f => f.name);
  assert.equal(names[0], '[Content_Types].xml');
  for (const n of ['_rels/.rels', 'docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
    'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml', 'xl/tables/table1.xml', 'xl/tables/table3.xml', 'xl/worksheets/_rels/sheet1.xml.rels']) {
    assert.ok(names.includes(n), n);
  }
  for (const f of files) assert.equal(f.crc, X.crc32(f.data), 'crc ' + f.name);
  assert.ok(!names.some(n => n.includes('webextension')));

  const wb = part(files, 'xl/workbook.xml');
  assert.match(wb, /<sheet name="Clinic Info" sheetId="3" state="hidden" r:id="rId3"\/>/);
  assert.match(wb, /activeTab="0"/);

  const t1 = part(files, 'xl/tables/table1.xml');
  assert.match(t1, /name="tblClinic" displayName="tblClinic" ref="A1:F3"/);
  assert.match(t1, /<autoFilter ref="A1:F3"\/>/);
  assert.match(t1, /<tableColumn id="6" name="Notes 2"\/>/);
  assert.match(t1, /TableStyleMedium2/);
  assert.match(part(files, 'xl/tables/table2.xml'), /ref="A1:B2"/);
  assert.match(part(files, 'xl/tables/table3.xml'), /id="3" name="tblClinicMeta"/);

  const s1 = part(files, 'xl/worksheets/sheet1.xml');
  assert.match(s1, /state="frozen"/);
  assert.match(s1, /<c r="A2" s="2" t="inlineStr"><is><t xml:space="preserve">000123<\/t>/, 'ID kept as text');
  assert.match(s1, /<c r="A3" s="2" t="inlineStr"><is><t xml:space="preserve">45<\/t>/, 'numeric ID written as text');
  assert.match(s1, /TESTER &amp; &lt;Co&gt;/);
  assert.match(s1, /<c r="C2" s="3"><v>0.2<\/v><\/c>/, 'percent format');
  assert.match(s1, /<c r="D2" s="4" t="inlineStr"><is><t xml:space="preserve">Re-verify in Sohar/);
  assert.match(s1, /<hyperlink ref="D2" r:id="rId2"\/>/);
  assert.ok(!/ref="D3"/.test(s1.split('<hyperlinks>')[1] || ''), 'unsafe URL not linked');
  assert.match(s1, /javascript:alert\(1\)/, 'unsafe URL kept as plain text');
  assert.match(s1, /<t xml:space="preserve">ab<\/t>/, 'control char stripped');
  assert.match(s1, /<tablePart r:id="rId1"\/>/);
  // element order: sheetData < hyperlinks < pageMargins < tableParts
  const order = ['<sheetViews', '<cols', '<sheetData', '<hyperlinks', '<pageMargins', '<tableParts'].map(t => s1.indexOf(t));
  assert.deepEqual(order, [...order].sort((a, b) => a - b));

  const r1 = part(files, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.match(r1, /Target="https:\/\/app\.soharhealth\.com\/create\?type=verification&amp;x=1" TargetMode="External"/);
  assert.match(r1, /Target="\.\.\/tables\/table1\.xml"/);
});

test('webExtension adds task pane parts and relationships', () => {
  const u8 = X.buildWorkbook(Object.assign({}, SPEC, {
    webExtension: { id: '9786ae8a-f258-4f2c-9627-4c6f0f25db7e', version: '0.1.0.0', store: 'EXCatalog', storeType: 'EXCatalog' },
  }));
  const files = X.unzipStored(u8);
  const we = part(files, 'xl/webextensions/webextension1.xml');
  assert.match(we, /<we:reference id="9786ae8a-f258-4f2c-9627-4c6f0f25db7e" version="0.1.0.0" store="EXCatalog" storeType="EXCatalog"\/>/);
  assert.match(we, /name="Office.AutoShowTaskpaneWithDocument" value="true"/);
  assert.match(we, /id="\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-8[0-9A-F]{3}-[0-9A-F]{12}\}"/);
  assert.match(part(files, 'xl/webextensions/taskpanes.xml'), /r:id="rId1"/);
  assert.match(part(files, 'xl/webextensions/_rels/taskpanes.xml.rels'), /Target="webextension1.xml"/);
  assert.match(part(files, '_rels/.rels'), /webextensiontaskpanes" Target="xl\/webextensions\/taskpanes.xml"/);
  const ct = part(files, '[Content_Types].xml');
  assert.match(ct, /application\/vnd.ms-office.webextensiontaskpanes\+xml/);
  assert.match(ct, /application\/vnd.ms-office.webextension\+xml/);
});

test('buildWorkbook rejects empty or all-hidden specs; base64 round-trips', () => {
  assert.throws(() => X.buildWorkbook({ sheets: [] }));
  assert.throws(() => X.buildWorkbook({ sheets: [{ name: 'a', columns: ['x'], rows: [], hidden: true }] }));
  const u8 = X.buildWorkbook({ sheets: [{ name: 'a', columns: ['x'], rows: [{ x: 1 }] }] });
  assert.deepEqual(new Uint8Array(Buffer.from(X.toBase64(u8), 'base64')), u8);
});
