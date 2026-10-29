// Rebuilds the data block in payer-knowledge.js from the master Sohar workbook.
//   node scripts/build-payer-knowledge.mjs "<path to Sohar_Eligibility_Template.xlsx>" [--review "<review.csv>"]
// Reads Payer Map, Payer Learning, Sohar Payers, Fee Locations and Fee Payer Map; keeps what clears the bar in
// payer-knowledge.js (select). The workbook stays where it is: never copy it into this repo. --review writes the fuzzy
// High Fee Payer Map matches for a person to check (keep that file outside the repo too).
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const require = createRequire(import.meta.url);
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const XL = require(path.join(ROOT, 'xlsx.js'));
const K = require(path.join(ROOT, 'payer-knowledge.js'));

const args = process.argv.slice(2);
const opt = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : '');
const reviewAt = opt('--review'), approvedAt = opt('--approved');
const file = args.find((a, i) => !a.startsWith('--') && !['--review', '--approved'].includes(args[i - 1]));
if (!file) { console.error('Usage: node scripts/build-payer-knowledge.mjs <workbook.xlsx> [--review out.csv] [--approved reviewed.csv]'); process.exit(1); }
if ([file, reviewAt, approvedAt].some((f) => f && path.resolve(f).startsWith(ROOT + path.sep))) {
  console.error('Keep the workbook and the review lists outside the repo.'); process.exit(1);
}

const str = (v) => (v == null ? '' : String(v).trim());
// A CSV saved by Excel (BOM, quoted fields) -> rows as objects keyed by its header.
function readCsv(f) {
  const t = readFileSync(f, 'utf8').replace(/^﻿/, ''), out = []; let row = [], cur = '', q = false;
  for (let i = 0; i < t.length; i++) { const ch = t[i];
    if (q) { if (ch === '"' && t[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
    else if (ch === '"') q = true; else if (ch === ',') { row.push(cur); cur = ''; } else if (ch === '\n') { row.push(cur.replace(/\r$/, '')); out.push(row); row = []; cur = ''; } else cur += ch; }
  if (cur || row.length) { row.push(cur); out.push(row); }
  const [h, ...body] = out.filter((r) => r.some((c) => str(c)));
  return body.map((r) => Object.fromEntries(h.map((k, i) => [str(k), r[i] ?? ''])));
}
const wb = await XL.readWorkbook(new Uint8Array(readFileSync(file)));
const table = (name) => {
  const s = wb.sheets.find((x) => x.name === name);
  if (!s || !s.rows || !s.rows.length) return [];
  const h = Array.from(s.rows[0], (c) => str(c));
  return s.rows.slice(1).filter((r) => (r || []).some((c) => str(c))).map((r) => Object.fromEntries(h.map((k, i) => [k, r[i] == null ? '' : r[i]])));
};
const sel = K.select({ payerMap: table('Payer Map'), payerLearning: table('Payer Learning'), soharPayers: table('Sohar Payers'),
  feeLocations: table('Fee Locations'), feePayerMap: table('Fee Payer Map') });

// --approved: the review list after a person filled "OK to build in?" (Yes / No / No, <flag> / a column name).
let feeFlags = [];
if (approvedAt) {
  const cols = new Map();
  for (const r of table('Fee Schedule')) { const t = str(r.Tab), p = str(r.Payer); if (t && p) { if (!cols.has(t)) cols.set(t, new Set()); cols.get(t).add(p); } }
  const ap = K.approved(readCsv(approvedAt), (tab) => [...(cols.get(tab) || [])]);
  const have = new Set(sel.feePayerMap.map(([p, t]) => `${p.toLowerCase()}|${t.toLowerCase()}`));
  for (const row of ap.feePayerMap) { const k = `${row[0].toLowerCase()}|${row[1].toLowerCase()}`; if (!have.has(k)) { have.add(k); sel.feePayerMap.push(row); } }
  feeFlags = ap.feeFlags;
  console.log(`Approved list: ${ap.feePayerMap.length} built in, ${ap.feeFlags.length} flags, ${ap.skipped.length} skipped`);
  for (const s of ap.skipped) console.log(`  skipped ${s.payer} on ${s.tab}: "${s.answer}"`);
}

const rows = (list) => (list.length ? `[\n${list.map((r) => `    ${JSON.stringify(r)},`).join('\n')}\n  ]` : '[]');
const counts = { payerMap: sel.payerMap.length, soharPayers: sel.soharPayers.length, feeLocations: sel.feeLocations.length, feePayerMap: sel.feePayerMap.length, feeFlags: feeFlags.length };
const block = [
  '  // BEGIN GENERATED: scripts/build-payer-knowledge.mjs. Do not edit by hand.',
  `  const META = ${JSON.stringify({ generated: new Date().toISOString().slice(0, 10), counts })};`,
  '  // [NextGen Payer Name, Sohar payerId, Sohar payerName, Confidence]',
  `  const PAYER_MAP = ${rows(sel.payerMap)};`,
  '  // [payerId, payerName, aliasPayerIds, aliasPayerNames, payerGroupId]',
  `  const SOHAR_PAYERS = ${rows(sel.soharPayers)};`,
  '  // [NextGen Location, Fee Tab]',
  `  const FEE_LOCATIONS = ${rows(sel.feeLocations)};`,
  '  // [Payer, Fee Tab, Fee Payer]',
  `  const FEE_PAYER_MAP = ${rows(sel.feePayerMap)};`,
  '  // [Payer, Fee Tab, Flag]',
  `  const FEE_FLAGS = ${rows(feeFlags)};`,
  '  // END GENERATED',
].join('\n');

const target = path.join(ROOT, 'payer-knowledge.js');
const src = readFileSync(target, 'utf8');
const a = src.indexOf('  // BEGIN GENERATED'), b = src.indexOf('  // END GENERATED');
if (a < 0 || b < 0) { console.error('payer-knowledge.js has no GENERATED markers.'); process.exit(1); }
writeFileSync(target, src.slice(0, a) + block + src.slice(b + '  // END GENERATED'.length));
console.log('payer-knowledge.js:', JSON.stringify(counts));

if (reviewAt) {
  const q = (v) => (/[",\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
  const cols = ['Payer', 'Fee Tab', 'Fee Payer', 'Auto Match', 'Note', 'OK to build in? (Yes / No / correct column)'];
  writeFileSync(reviewAt, '﻿' + [cols, ...sel.review.map((r) => cols.map((c) => str(r[c])))].map((r) => r.map(q).join(',')).join('\r\n') + '\r\n');
  console.log(`Review list: ${sel.review.length} fuzzy High matches -> ${reviewAt}`);
}
