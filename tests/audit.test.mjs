import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const XL = require('../xlsx.js');
const AU = require('../audit.js');

// v0.23: the account audit (InstaMed + Arietis + NextGen kept visits) and the legacy .xls reader. Synthetic data only:
// every name, MRN, card number, code amount and date below is made up.

const TODAY = Date.UTC(2026, 10, 2);
const WHO = { lasfir: 'TESPAT', mrn: '700001' };
const demo = () => AU.demoRows('Tester', 'Pat', '700001', TODAY);
function demoAudit(extra) {
  const d = demo();
  const im = AU.parseInstamed(d.instamed), ar = AU.parseArietis(d.arietis), ap = AU.parseAppointments(d.appointments);
  const id = AU.identify('appointments', ap, WHO);
  return AU.runAudit({ who: WHO, instamed: im, arietis: ar, appointments: { ...ap, visits: id.visits }, today: TODAY, ...(extra || {}) });
}

// ---------- a tiny BIFF8 .xls writer, for the reader test ----------
function rec(type, bytes) { const b = new Uint8Array(4 + bytes.length); const v = new DataView(b.buffer); v.setUint16(0, type, true); v.setUint16(2, bytes.length, true); b.set(bytes, 4); return b; }
function le(...parts) { // [kind, value] -> bytes
  const out = [];
  for (const [k, v] of parts) {
    const b = new Uint8Array(k === 'u8' ? 1 : k === 'u16' ? 2 : k === 'u32' ? 4 : k === 'f64' ? 8 : v.length);
    const dv = new DataView(b.buffer);
    if (k === 'u8') dv.setUint8(0, v); else if (k === 'u16') dv.setUint16(0, v, true); else if (k === 'u32') dv.setUint32(0, v, true); else if (k === 'f64') dv.setFloat64(0, v, true); else b.set(v);
    out.push(b);
  }
  const n = out.reduce((a, b) => a + b.length, 0), r = new Uint8Array(n); let o = 0; for (const b of out) { r.set(b, o); o += b.length; } return r;
}
const latin = (s) => Uint8Array.from([...s].map((c) => c.charCodeAt(0)));
const utf16 = (s) => { const b = new Uint8Array(s.length * 2); [...s].forEach((c, i) => { b[i * 2] = c.charCodeAt(0) & 255; b[i * 2 + 1] = c.charCodeAt(0) >> 8; }); return b; };
const cat = (bs) => le(...bs.map((b) => ['raw', b]));
function buildXls() {
  // shared strings: 0 'Patient Account #' split across SST and CONTINUE, 1 'Amount', 2 a UTF-16 string
  const s0 = 'Patient Account #';
  const sstHead = le(['u32', 3], ['u32', 3], ['u16', s0.length], ['u8', 0], ['raw', latin(s0.slice(0, 7))]);
  const cont = le(['u8', 0], ['raw', latin(s0.slice(7))], ['u16', 6], ['u8', 0], ['raw', latin('Amount')], ['u16', 5], ['u8', 1], ['raw', utf16('Ünïcø')]);
  const bof = (dt) => rec(0x0809, le(['u16', 0x0600], ['u16', dt], ['u16', 0], ['u16', 0], ['u32', 0], ['u32', 0]));
  const eof = rec(0x000A, new Uint8Array(0));
  const sheet = (name, hidden, pos) => rec(0x0085, le(['u32', pos], ['u8', hidden ? 1 : 0], ['u8', 0], ['u8', name.length], ['u8', 0], ['raw', latin(name)]));
  const cell = (r, c) => [['u16', r], ['u16', c], ['u16', 15]];
  const s1 = cat([bof(0x10),
    rec(0x00FD, le(...cell(0, 0), ['u32', 0])), rec(0x00FD, le(...cell(0, 1), ['u32', 1])),
    rec(0x0203, le(...cell(1, 1), ['f64', 44.7])), rec(0x027E, le(...cell(1, 0), ['u32', (46189 << 2) | 2])),
    rec(0x00BD, le(['u16', 2], ['u16', 0], ['u16', 15], ['u32', (1234 << 2) | 3], ['u16', 15], ['u32', (-5 << 2) | 2], ['u16', 1])), // 12.34, -5
    rec(0x0006, le(...cell(3, 0), ['raw', Uint8Array.from([0, 0, 0, 0, 0, 0, 0xFF, 0xFF])], ['u16', 0], ['u32', 0], ['u16', 0])), rec(0x0207, le(['u16', 3], ['u8', 0], ['raw', latin('abc')])),
    rec(0x00FD, le(...cell(4, 2), ['u32', 2])), eof]);
  const s2 = cat([bof(0x10), rec(0x0203, le(...cell(0, 0), ['f64', 1])), eof]);
  const pad = rec(0x00FF, new Uint8Array(4200)); // EXTSST-type filler so the stream is past the mini-stream cutoff
  const globalsLen = (pos1, pos2) => cat([bof(5), sheet('Payment Transactions', false, pos1), sheet('Hidden', true, pos2), rec(0x00FC, sstHead), rec(0x003C, cont), pad, eof]).length;
  const g0 = globalsLen(0, 0);
  const globals = cat([bof(5), sheet('Payment Transactions', false, g0), sheet('Hidden', true, g0 + s1.length), rec(0x00FC, sstHead), rec(0x003C, cont), pad, eof]);
  const wb = cat([globals, s1, s2]);
  // Compound File: header, FAT sector 0, directory sector 1, stream from sector 2
  const nSec = Math.ceil(wb.length / 512);
  const total = 512 + 512 * (2 + nSec), out = new Uint8Array(total), dv = new DataView(out.buffer);
  out.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
  dv.setUint16(0x18, 0x3E, true); dv.setUint16(0x1A, 3, true); dv.setUint16(0x1C, 0xFFFE, true); dv.setUint16(0x1E, 9, true); dv.setUint16(0x20, 6, true);
  dv.setUint32(0x2C, 1, true); dv.setUint32(0x30, 1, true); dv.setUint32(0x38, 4096, true); dv.setUint32(0x3C, 0xFFFFFFFE, true); dv.setUint32(0x44, 0xFFFFFFFE, true);
  dv.setUint32(0x4C, 0, true); for (let i = 1; i < 109; i++) dv.setUint32(0x4C + i * 4, 0xFFFFFFFF, true);
  const fat = 512;
  for (let i = 0; i < 128; i++) dv.setUint32(fat + i * 4, 0xFFFFFFFF, true);
  dv.setUint32(fat, 0xFFFFFFFD, true); dv.setUint32(fat + 4, 0xFFFFFFFE, true);
  for (let s = 0; s < nSec; s++) dv.setUint32(fat + (2 + s) * 4, s === nSec - 1 ? 0xFFFFFFFE : 3 + s, true);
  const dir = 1024;
  const entry = (i, nm, type, start, size) => { const o = dir + i * 128; [...nm].forEach((c, j) => dv.setUint16(o + j * 2, c.charCodeAt(0), true)); dv.setUint16(o + 64, (nm.length + 1) * 2, true); out[o + 66] = type; dv.setUint32(o + 116, start, true); dv.setUint32(o + 120, size, true); };
  entry(0, 'Root Entry', 5, 0xFFFFFFFE, 0); entry(1, 'Workbook', 2, 2, wb.length);
  out.set(wb, 1536);
  return out;
}

test('readXls reads a BIFF8 workbook: shared strings across CONTINUE, RK, MULRK, numbers, formula text, hidden sheets', async () => {
  const bytes = buildXls();
  assert.equal(XL.isCfb(bytes), true);
  const wb = await XL.readAnyWorkbook(bytes);
  assert.deepEqual(wb.sheets.map((s) => [s.name, s.hidden]), [['Payment Transactions', false], ['Hidden', true]]);
  const r = wb.sheets[0].rows;
  assert.deepEqual(r[0], ['Patient Account #', 'Amount']);
  assert.deepEqual(r[1], [46189, 44.7]);
  assert.deepEqual(r[2], [12.34, -5]);
  assert.equal(r[3][0], 'abc');
  assert.deepEqual(r[4], ['', '', 'Ünïcø']);
  assert.deepEqual(wb.sheets[1].rows, [[1]]);
});
test('readAnyWorkbook says what to do with a web page saved as .xls and with other files', async () => {
  await assert.rejects(XL.readAnyWorkbook(new TextEncoder().encode('<html><table></table></html>')), /web page/);
  await assert.rejects(XL.readAnyWorkbook(new TextEncoder().encode('hello')), /not an Excel workbook/);
  assert.throws(() => XL.readXls(new Uint8Array(600)), /not an Excel 97-2003/);
});
test('readAnyWorkbook still reads .xlsx', async () => {
  const x = XL.buildWorkbook({ sheets: [{ name: 'S', columns: ['A', 'B'], rows: [{ A: 'x', B: 2 }] }] });
  const wb = await XL.readAnyWorkbook(x);
  assert.equal(wb.sheets[0].rows[1][0], 'x');
});

// ---------- reading the three files ----------
test('each file is recognised by its header; the Arietis charges carry their transactions', () => {
  const d = demo();
  assert.equal(AU.detectKind(d.instamed), 'instamed');
  assert.equal(AU.detectKind(d.arietis), 'arietis');
  assert.equal(AU.detectKind(d.appointments), 'appointments');
  assert.equal(AU.detectKind([['Appt Date', 'Event']]), '');
  const ar = AU.parseArietis(d.arietis);
  assert.equal(ar.charges.length, 13);
  const first = ar.charges[0];
  assert.equal(first.code, '90868');
  assert.equal(first.insPaid, 246.34); assert.equal(first.coins, 105.57); assert.equal(first.contractual, 2099.09); assert.equal(first.patPaid, 105.57);
  assert.equal(first.processed, true); assert.equal(first.calcBalance, 0);
  const em = ar.charges.find((c) => c.code === '99214');
  assert.deepEqual(em.mods, ['95']);
  assert.equal(AU.chargeCode('CLAIM-1E'), null);
  assert.deepEqual(AU.chargeCode('90833 - 95'), { code: '90833', mods: ['95'] });
});
test('InstaMed: voided transactions are listed but not counted; the report total is read', () => {
  const im = AU.parseInstamed(demo().instamed);
  assert.equal(im.txns.length, 8);
  assert.equal(im.txns.filter((t) => !t.counted).length, 1);
  assert.equal(im.reportTotal, 758.44);
  assert.equal(im.total, 745.93);
  assert.deepEqual(im.lasfirs, ['TESPAT']);
});
test('identify: another patient\'s file is not used; a whole-clinic NextGen report keeps only this patient', () => {
  const d = demo();
  assert.equal(AU.identify('arietis', AU.parseArietis(d.arietis), WHO).ok, true);
  const other = AU.identify('arietis', AU.parseArietis(d.arietis), { lasfir: 'OTHPER', mrn: '1' });
  assert.equal(other.ok, false); assert.match(other.text, /for TESPAT, not OTHPER/);
  const rows = [...d.appointments, ['Someone, Else', '700002', 46300, '9:00 A', 'TMS Treatment Unilateral', 'Kept', 'Demo Clinic', 'X']];
  const id = AU.identify('appointments', AU.parseAppointments(rows), WHO);
  assert.equal(id.ok, true);
  assert.ok(id.visits.every((v) => v.mrn === '700001'));
});

// ---------- codes ----------
test('codes billed against the visit\'s event', () => {
  const v = (event) => ({ event, loc: 'Demo Clinic' });
  const ch = (code, mods) => ({ code, mods: mods || [] });
  assert.deepEqual(AU.codeChecks([v('TMS Treatment Unilateral')], [ch('90868')]), []);
  assert.equal(AU.codeChecks([v('TMS Treatment Unilateral')], [ch('99214')])[0].sev, 'high');
  const lvl = AU.codeChecks([v('Med Management')], [ch('99215')]);
  assert.equal(lvl[0].sev, 'medium'); assert.match(lvl[0].text, /different E\/M established level/);
  const add = AU.codeChecks([v('Med Management')], [ch('99214'), ch('90833')]);
  assert.equal(add.length, 1); assert.equal(add[0].sev, 'low'); assert.match(add[0].text, /add-on 90833/);
  assert.match(AU.codeChecks([v('Telepsych TMS F/U')], [ch('99214')])[0].text, /without modifier 95/);
  assert.deepEqual(AU.codeChecks([v('Telepsych TMS F/U')], [ch('99214', ['95'])]), []);
  assert.deepEqual(AU.codeChecks([v('Esketamine Treatment')], [ch('G2082')]), []);
  // 90867 is the MT visit, 90869 the MTR (MT redo) (Brendan, 11/2).
  assert.deepEqual(AU.expectedFor('TMS Motor Threshold').lines, [['90867']]);
  assert.deepEqual(AU.expectedFor('MT').lines, [['90867']]);
  for (const ev of ['MTR', 'MT-Redo', 'MT Redo', 'Motor Threshold Redo', 'Motor threshold re-determination']) assert.deepEqual(AU.expectedFor(ev).lines, [['90869']], ev);
  assert.match(AU.codeChecks([v('MT-Redo')], [ch('90867')])[0].text, /expected 90869, billed 90867 \(different TMS motor threshold level\)/);
});

// ---------- payments ----------
test('InstaMed to Arietis: batches, a payment split over two charges, one never posted', () => {
  const a = demoAudit();
  const p = a.payments;
  assert.equal(p.matches.length, 6);
  assert.ok(p.matches.some((m) => m.amount === 211.14 && m.lines.length === 2));
  assert.deepEqual(p.unposted.map((t) => t.amount), [71.51]);
  const f = a.anomalies.find((x) => x.source === 'InstaMed' && x.sev === 'high');
  assert.match(f.text, /\$71\.51 .*not posted in Arietis/);
});
test('identical payments on one card: the most recent is the one left unposted (Arietis posts late), and it says so', () => {
  const day = (n) => Date.UTC(2026, 9, n);
  const im = { txns: [1, 2, 3].map((n, i) => ({ day: day(n), last4: '4242', amount: 50, counted: true, row: i })), start: day(1), end: day(3) };
  const line = (n, b) => ({ type: 'patPmt', posted: day(n), amount: -50, ref: '4242', refType: 'IM', batch: b, row: n });
  const ar = { charges: [{ dos: day(1), code: '90868', lines: [line(3, '1'), line(4, '2')] }], asOf: day(4) };
  const p = AU.matchPayments(im, ar);
  assert.deepEqual(p.unposted.map((t) => t.day), [day(3)]);
  assert.equal(p.unposted[0].ambiguous, true);
});
test('a payment and its reversal cancel', () => {
  const day = Date.UTC(2026, 5, 1);
  const ar = { charges: [{ dos: day, code: '99214', lines: [
    { type: 'patPmt', posted: day, amount: -35, ref: '9999', refType: 'CC', batch: '4', row: 1 }, { type: 'patRev', posted: day, amount: 35, ref: '9999', refType: 'CC', batch: '4', row: 2 }] }] };
  const p = AU.matchPayments(null, ar);
  assert.equal(p.reversed.length, 2); assert.equal(p.other.length, 0); assert.equal(p.postedTotal, 0);
});

// ---------- benefits and the re-quote ----------
test('benefits learned from processed claims: TMS coinsurance, a visit copay once per day', () => {
  const a = demoAudit();
  const L = a.learned.classes;
  assert.equal(L.TMS.rule, 'coins'); assert.equal(L.TMS.coins, 0.3);
  assert.equal(L.Visit.rule, 'copay'); assert.equal(L.Visit.copay, 41);
  assert.equal(L.TMS.allowed['90868'], 351.91);
});
test('requote: deductible first, then coinsurance; copay once a day; capped at the OOP max', () => {
  const cls = { TMS: { rule: 'coins', coins: 0.2 }, Visit: { rule: 'copay', copay: 30 } };
  const q = AU.requote([
    { day: 1, lines: [{ code: '90868', allowed: 300 }] },
    { day: 2, lines: [{ code: '99214', allowed: 100 }, { code: '90833', allowed: 80 }] },
    { day: 3, lines: [{ code: '90868', allowed: 300 }] },
  ], { deductible: 200, dedMet: 0, oopMax: 500, oopMet: 250, cls });
  // day 1: $200 deductible + 20% of the other $100; day 2: one $30 copay for both codes; day 3: $60 coinsurance, but
  // only $0 of the OOP max is left
  assert.deepEqual(q.map((x) => x.patient), [220, 30, 0]);
  assert.equal(q[0].ded, 200); assert.equal(q[1].copay, 30); assert.equal(q[2].capped, true); assert.equal(q[2].oopLeft, 0);
});
test('plan year from Sohar\'s coverage start, else January 1', () => {
  const p = AU.planYearOf(Date.UTC(2026, 2, 10), '07/01/2025');
  assert.equal(AU.iso(p.start), '2025-07-01'); assert.equal(AU.iso(p.end), '2026-06-30');
  assert.equal(AU.iso(AU.planYearOf(Date.UTC(2026, 2, 10)).start), '2026-01-01');
});

// ---------- the audit ----------
test('runAudit: the true balance adds up from its parts and the anomalies name each system', () => {
  const a = demoAudit({ booked: [{ date: '11/03/2026', event: 'TMS Treatment Unilateral' }] });
  const b = a.balance;
  assert.equal(b.true, Math.round((b.processed + b.pendingPatient - b.unposted) * 100) / 100);
  assert.equal(b.arietis, Math.round((b.processed + b.pendingBilled) * 100) / 100);
  assert.equal(b.pendingCount, 4);
  assert.equal(b.booked, 105.57); // 30% of 351.91 for the booked visit
  const t = (re) => a.anomalies.find((x) => re.test(x.text));
  assert.equal(t(/doesn't add up/).sev, 'high');
  assert.equal(t(/marked "No Show"/).source, 'NextGen');
  assert.equal(t(/not billed yet/).sev, 'low');
  assert.equal(t(/add-on 90833/).source, 'Codes');
});
test('runAudit: Sohar\'s coinsurance and the calculator are cross-checked against the claims', () => {
  const a = demoAudit({ sohar: { 'OOP Max': '3000', 'OOP Remaining': '2900', Coinsurance: '0.2', Copay: '41', 'Sohar Date': '09/01/2026' } });
  assert.ok(a.anomalies.some((x) => x.source === 'Sohar' && /30% coinsurance; Sohar says 20%/.test(x.text)));
  assert.ok(a.calc.days > 0);
  assert.ok(a.anomalies.some((x) => x.source === 'Calculator'));
  // met = Sohar's figure plus what posted after its check, when that is more than our claims alone
  assert.ok(a.inputs.oopMet >= 100);
});
test('confirming a day\'s codes turns its code findings into confirmed ones', () => {
  const a0 = demoAudit();
  const day = a0.anomalies.find((x) => x.source === 'Codes').day;
  const a = demoAudit({ confirms: { [AU.iso(day)]: true } });
  const f = a.anomalies.find((x) => x.source === 'Codes' && x.day === day);
  assert.equal(f.sev, 'info'); assert.equal(f.confirmed, true);
});
test('a typed quote that differs from the processed patient share is flagged', () => {
  const a0 = demoAudit();
  const d = a0.days.find((x) => x.processed && x.resp > 0);
  const a = demoAudit({ quotes: { [d.iso]: '10' } });
  assert.ok(a.anomalies.some((x) => x.source === 'Quote' && x.day === d.day));
});
test('the copied summary has LASFIR and MRN but no names or card numbers', () => {
  const a = demoAudit();
  const s = AU.summaryText(a, WHO, ['InstaMed payments', 'Arietis account detail']);
  assert.match(s, /TESPAT \(MRN 700001\)/);
  assert.doesNotMatch(s, /tester|\bpat\b/i);
  assert.doesNotMatch(s, /4242/);
});

// ---------- saved audits and the export (v0.23.1) ----------
const NAME_COLS = ['Last Name', 'First Name', 'DOB', 'Patient', 'Pat Name', 'Patient Last Name', 'Patient First Name', 'Card Holder Name', 'Card # - Last 4 Digits'];
test('saved audit columns: LASFIR and MRN, no name, DOB or card columns', () => {
  for (const cols of [AU.AUDIT_COLS, AU.AUDIT_LINE_COLS]) {
    assert.deepEqual(cols.filter((c) => NAME_COLS.includes(c)), []);
    assert.ok(cols.includes('LASFIR') && cols.includes('MRN') && cols[0] === 'Key');
  }
});
test('auditRecord: one summary row and its lines, with no names or card numbers anywhere', () => {
  const a = demoAudit({ booked: [{ date: '11/03/2026', event: 'TMS Treatment Unilateral' }] });
  const rec = AU.auditRecord(a, { ...WHO, key: '700001' }, { by: 'Synthetic Admin', when: '11/2/2026 1:30 PM', sources: ['InstaMed payments', 'Arietis account detail'] });
  assert.equal(rec.summary.Key, '700001'); assert.equal(rec.summary['True Balance'], a.balance.true); assert.equal(rec.summary.High, a.counts.high);
  const parts = new Set(rec.lines.map((l) => l.Part));
  assert.deepEqual([...parts], ['Finding', 'Visit', 'Re-quote', 'Payment']);
  const all = JSON.stringify(rec);
  assert.doesNotMatch(all, /tester|4242/i);
  assert.ok(rec.lines.every((l) => AU.AUDIT_LINE_COLS.every((c) => c in l)));
  assert.ok(rec.lines.some((l) => l.Part === 'Payment' && l.Status === 'Not posted' && l.Paid === 71.51));
});
test('saving again replaces only that patient\'s audit', () => {
  const a = demoAudit();
  const r1 = AU.auditRecord(a, { ...WHO, key: '700001' }, { when: 'first' });
  const other = AU.auditRecord(a, { lasfir: 'OTHPER', mrn: '700002', key: '700002' }, { when: 'other' });
  let saved = AU.mergeAudits(AU.mergeAudits({ summary: [], lines: [] }, r1), other);
  const r2 = AU.auditRecord(a, { ...WHO, key: '700001' }, { when: 'second' });
  saved = AU.mergeAudits(saved, r2);
  assert.equal(saved.summary.length, 2);
  assert.equal(AU.savedFor(saved, '700001').Saved, 'second');
  assert.equal(saved.lines.filter((l) => l.Key === '700001').length, r2.lines.length);
  assert.ok(saved.lines.filter((l) => l.Key === '700001').every((l) => l.Saved === 'second'));
  assert.equal(saved.lines.filter((l) => l.Key === '700002').length, other.lines.length);
});
test('the export compiles saved audits into Summary, Findings, Visits, Re-quote and Payments, money as numbers', async () => {
  const a = demoAudit();
  const rec = AU.auditRecord(a, { ...WHO, key: '700001' }, { when: 'w' });
  // as read back from the hidden sheet: every value is text
  const asText = (r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, v === '' ? '' : String(v)]));
  const saved = { summary: [asText(rec.summary)], lines: rec.lines.map(asText) };
  const spec = AU.exportSpec(saved);
  assert.deepEqual(spec.sheets.map((s) => s.name), ['Summary', 'Findings', 'Visits', 'Re-quote', 'Payments']);
  assert.equal(spec.sheets[0].rows[0]['True Balance'], a.balance.true);
  assert.equal(typeof spec.sheets[2].rows[0].Balance, 'number');
  assert.ok(!spec.sheets[0].columns.includes('Key'));
  assert.equal(AU.exportSpec(saved, ['nobody']).sheets[0].rows.length, 0);
  const wb = await XL.readWorkbook(XL.buildWorkbook(spec));
  assert.equal(wb.sheets.length, 5);
  assert.equal(wb.sheets[0].rows[1][wb.sheets[0].rows[0].indexOf('LASFIR')], 'TESPAT');
  assert.match(AU.exportName('TESPAT', new Date(2026, 10, 2)), /^Account Audit - TESPAT 2026-11-02\.xlsx$/);
  assert.match(AU.exportName('', new Date(2026, 10, 2)), /^Account Audits 2026-11-02\.xlsx$/);
});
