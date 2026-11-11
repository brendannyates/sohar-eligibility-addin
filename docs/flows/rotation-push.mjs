// rotation-push.mjs (Clinic Workbook Pilot, v0.29): the eligibility-check rotation's side of "Clinic Workbook - Rotation to
// master". After a run writes Eligibility Checks.csv, push(...) sends the same rows to the master workbook's flow-owned tables,
// and needsOpen / needsDone run the needs-check loop. Pure apart from fetch. No PHI, URL or key in this file:
//   CLINIC_ROTATION_URL  the flow's HTTP POST URL (it carries a signature: treat it as a password)
//   CLINIC_ROTATION_KEY  the rotation key Brendan typed into the flow's "Rotation key" Compose card
// Both live in ~/.secrets (loaded into the environment by clients.js), never in this repo, a chat or a log.
// Wire-up (in rotate.js, after the CSV write):  import { pushLedger } from '<this file>'; await pushLedger(rows, process.env);

export const BATCH = 200;

const str = (v) => (v == null ? '' : String(v).trim());
/** M/D/YYYY, MM/DD/YYYY or YYYY-MM-DD -> YYYY-MM-DD ('' when unreadable). The master ledger keeps ISO days. */
export function isoDay(v) {
  const s = str(v);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}` : '';
}
/** The ledger's upsert key: MRN | Appt Date (ISO) | Time, the same key the CSV ledger upserts on. */
export const rowKey = (r) => `${str(r.MRN)}|${isoDay(r['Appt Date'])}|${str(r.Time)}`;

/** Ledger CSV rows (objects by header) -> master rows: ISO Appt Date, Row Key, and only the columns the table has. */
export function ledgerRows(rows, cols) {
  const keep = new Set(cols);
  return (rows || []).filter((r) => str(r.MRN) && isoDay(r['Appt Date'])).map((r) => {
    const o = { 'Row Key': rowKey(r) };
    for (const [k, v] of Object.entries(r)) if (keep.has(k) && k !== 'Row Key') o[k] = k === 'Appt Date' ? isoDay(v) : str(v);
    return o;
  });
}
export const batches = (rows, n = BATCH) => Array.from({ length: Math.ceil((rows || []).length / n) }, (_, i) => rows.slice(i * n, (i + 1) * n));

async function post(env, body, fetchFn = globalThis.fetch) {
  const url = str(env.CLINIC_ROTATION_URL), key = str(env.CLINIC_ROTATION_KEY);
  if (!url || !key) throw new Error('Set CLINIC_ROTATION_URL and CLINIC_ROTATION_KEY in ~/.secrets first.');
  const res = await fetchFn(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ v: 1, key, ...body }) });
  let json = null;
  try { json = await res.json(); } catch { json = null; }
  if (!res.ok || !json || json.ok !== true) throw new Error(`Rotation to master: ${res.status} ${json && json.error ? json.error : ''}`.trim());
  return json;
}
/** Upsert rows into one table (ledger | history | copay | cards), in batches. Returns {updated, added, skipped}. */
export async function push(table, rows, env, fetchFn) {
  const sum = { updated: 0, added: 0, skipped: 0 };
  for (const b of batches(rows)) {
    const r = await post(env, { op: 'upsert', table, rows: b }, fetchFn);
    sum.updated += r.updated || 0; sum.added += r.added || 0; sum.skipped += r.skipped || 0;
  }
  return sum;
}
export const pushLedger = (rows, env, cols, fetchFn) => push('ledger', ledgerRows(rows, cols), env, fetchFn);
/** The needs-check loop: open requests from the clinics, then mark each done with a short result (no PHI beyond the MRN). */
export async function needsOpen(env, fetchFn) { return (await post(env, { op: 'needs-open' }, fetchFn)).rows || []; }
export async function needsDone(done, env, fetchFn) {
  const rows = (done || []).map((d) => ({ 'Row Key': str(d['Row Key']), Status: 'Done', 'Done At': new Date().toISOString(), Result: str(d.Result).slice(0, 120) }));
  return rows.length ? post(env, { op: 'needs-done', rows }, fetchFn) : { ok: true, updated: 0 };
}
