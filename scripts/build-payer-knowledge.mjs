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
const file = args.find((a) => !a.startsWith('--'));
const reviewAt = args.includes('--review') ? args[args.indexOf('--review') + 1] : '';
if (!file) { console.error('Usage: node scripts/build-payer-knowledge.mjs <workbook.xlsx> [--review out.csv]'); process.exit(1); }
if (path.resolve(file).startsWith(ROOT + path.sep) || (reviewAt && path.resolve(reviewAt).startsWith(ROOT + path.sep))) {
  console.error('Keep the workbook and the review list outside the repo.'); process.exit(1);
}

const str = (v) => (v == null ? '' : String(v).trim());
const wb = await XL.readWorkbook(new Uint8Array(readFileSync(file)));
const table = (name) => {
  const s = wb.sheets.find((x) => x.name === name);
  if (!s || !s.rows || !s.rows.length) return [];
  const h = Array.from(s.rows[0], (c) => str(c));
  return s.rows.slice(1).filter((r) => (r || []).some((c) => str(c))).map((r) => Object.fromEntries(h.map((k, i) => [k, r[i] == null ? '' : r[i]])));
};
const sel = K.select({ payerMap: table('Payer Map'), payerLearning: table('Payer Learning'), soharPayers: table('Sohar Payers'),
  feeLocations: table('Fee Locations'), feePayerMap: table('Fee Payer Map') });

const rows = (list) => (list.length ? `[\n${list.map((r) => `    ${JSON.stringify(r)},`).join('\n')}\n  ]` : '[]');
const counts = { payerMap: sel.payerMap.length, soharPayers: sel.soharPayers.length, feeLocations: sel.feeLocations.length, feePayerMap: sel.feePayerMap.length };
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
