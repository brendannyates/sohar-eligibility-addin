// Office Script "Rotation to master" (Clinic Workbook Pilot, v0.29). Run by the flow "Clinic Workbook - Rotation to master"
// (Excel Online (Business): Run script) against the master Sohar_Eligibility_Template.xlsx. No PHI in this file.
// The flow-owned tables (rules.js LEDGER_COLS and friends) are created once by the Sohar add-in (Admin > Set up the rotation
// tables). This script only writes rows in them, one row at a time: it never clears, sorts or rewrites a whole table,
// because the lookup flow appends to Needs Check while this runs.
//
// input (JSON text): {"op": "upsert", "table": "ledger" | "history" | "copay" | "cards", "rows": [{column: value}]}
//                    {"op": "needs-open"}  ->  the open Needs Check rows
//                    {"op": "needs-done", "rows": [{"Row Key", "Status", "Done At", "Result"}]}
// Returns JSON text: {"ok": true, "updated": n, "added": n, "skipped": n, "rows": [...]} or {"ok": false, "error": "..."}.

const TABLES: { [k: string]: { table: string, key: string } } = {
  ledger: { table: 'tblEligibilityChecks', key: 'Row Key' },
  history: { table: 'tblPatientHistory', key: 'Row Key' },
  copay: { table: 'tblCopayActivity', key: 'Row Key' },
  cards: { table: 'tblCardOnFile', key: 'MRN' },
  needs: { table: 'tblNeedsCheck', key: 'Row Key' },
};
const MAX_ROWS = 500; // a bigger batch is split by the caller

type Cell = string | number | boolean;
type Rec = { [k: string]: Cell };
interface Plan { updates: { index: number, values: Cell[] }[], adds: Cell[][], skipped: number }

function str(v: Cell | undefined | null): string { return v === undefined || v === null ? '' : String(v).trim(); }
// MRNs compare without leading zeros; every other key as typed.
function keyText(col: string, v: Cell | undefined): string { const s = str(v); return col === 'MRN' ? s.replace(/^0+(?=\d)/, '') : s; }

/**
 * Pure: where each incoming row goes. header + body are the table as it is; a row whose key is already there updates
 * that row (only the columns it carries; a blank incoming cell keeps what is there), else it is added. A row with no key
 * is skipped (and, with addNew false, so is a row whose key is not there). The table's one blank starter row (a header-only table keeps one) is filled before any row is added.
 */
function planUpsert(header: string[], body: Cell[][], keyCol: string, rows: Rec[], addNew: boolean = true): Plan {
  const k = header.indexOf(keyCol);
  if (k < 0) throw new Error(`The table has no "${keyCol}" column.`);
  const at: { [key: string]: number } = {};
  body.forEach((r, i) => { const v = keyText(keyCol, r[k]); if (v && at[v] === undefined) at[v] = i; });
  const blank = body.length === 1 && body[0].every((c) => str(c) === '') ? 0 : -1;
  const plan: Plan = { updates: [], adds: [], skipped: 0 };
  const seen: { [key: string]: number } = {};
  let usedBlank = false;
  for (const r of rows) {
    const key = keyText(keyCol, r[keyCol]);
    if (!key) { plan.skipped++; continue; }
    const i = at[key];
    if (i !== undefined) {
      const cur = body[i].slice();
      // the key cell stays as stored (an MRN keeps its leading zeros)
      header.forEach((h, j) => { if (j !== k && Object.prototype.hasOwnProperty.call(r, h) && str(r[h]) !== '') cur[j] = str(r[h]); });
      plan.updates.push({ index: i, values: cur });
      body[i] = cur;
      continue;
    }
    if (!addNew) { plan.skipped++; continue; } // needs-done: only rows that are there
    if (seen[key] !== undefined) { // the same key twice in one batch: the later row wins
      const row = plan.adds[seen[key]];
      header.forEach((h, j) => { if (Object.prototype.hasOwnProperty.call(r, h) && str(r[h]) !== '') row[j] = str(r[h]); });
      continue;
    }
    const values = header.map((h) => str(r[h]));
    if (blank === 0 && !usedBlank) { usedBlank = true; plan.updates.push({ index: 0, values }); at[key] = 0; body[0] = values; continue; }
    seen[key] = plan.adds.length;
    plan.adds.push(values);
  }
  return plan;
}

/** Pure: the open Needs Check rows as objects (Status blank or Open). */
function openNeeds(header: string[], body: Cell[][]): Rec[] {
  const s = header.indexOf('Status'), k = header.indexOf('Row Key');
  return body.filter((r) => str(r[k]) && (!str(r[s]) || /^open$/i.test(str(r[s]))))
    .map((r) => { const o: Rec = {}; header.forEach((h, j) => { o[h] = str(r[j]); }); return o; });
}

function main(workbook: ExcelScript.Workbook, input: string): string {
  let q: { op?: string, table?: string, rows?: Rec[] };
  try { q = JSON.parse(input); } catch (e) { return JSON.stringify({ ok: false, error: 'The input is not JSON.' }); }
  const op = str(q.op);
  const spec = TABLES[op === 'upsert' ? str(q.table) : 'needs'];
  if (!spec || ['upsert', 'needs-open', 'needs-done'].indexOf(op) < 0) return JSON.stringify({ ok: false, error: `Unknown op or table: ${op} ${str(q.table)}` });
  const t = workbook.getTable(spec.table);
  if (!t) return JSON.stringify({ ok: false, error: `${spec.table} is not in this workbook. Set up the rotation tables in the Sohar add-in first.` });
  const header = t.getHeaderRowRange().getValues()[0].map((h) => str(h));
  const bodyRange = t.getRangeBetweenHeaderAndTotal();
  const body = bodyRange.getValues() as Cell[][];
  if (op === 'needs-open') return JSON.stringify({ ok: true, rows: openNeeds(header, body) });
  const rows = (Array.isArray(q.rows) ? q.rows : []).slice(0, MAX_ROWS);
  const keep = op === 'needs-done' ? ['Row Key', 'Status', 'Done At', 'Result'] : header;
  const trimmed = rows.map((r) => { const o: Rec = {}; for (const h of keep) if (Object.prototype.hasOwnProperty.call(r, h)) o[h] = r[h]; return o; });
  // needs-done only updates rows that exist; it never adds one.
  const plan = planUpsert(header, body, spec.key, trimmed, op !== 'needs-done');
  const text = header.map(() => '@');
  for (const u of plan.updates) {
    const row = bodyRange.getRow(u.index);
    row.setNumberFormats([text]);
    row.setValues([u.values]);
  }
  if (plan.adds.length) {
    // Text first, so MRNs keep their leading zeros and dates stay as the rotation wrote them.
    for (let c = 0; c < header.length; c++) t.getColumn(header[c]).getRangeBetweenHeaderAndTotal().setNumberFormatLocal('@');
    t.addRows(-1, plan.adds);
  }
  return JSON.stringify({ ok: true, updated: plan.updates.length, added: plan.adds.length, skipped: plan.skipped + Math.max(0, (q.rows || []).length - MAX_ROWS) });
}
