// Office Script "Lookup summary" (Clinic Workbook Pilot, v0.29). Run every Sunday by "Clinic Workbook - Lookup audit summary"
// against Clinic Lookup Log.xlsx (tblLookupLog). Weeks run Sunday to Saturday. The summary is counts only: no MRNs, no
// reasons, so it can go to Brendan by email. The log itself stays in the admin-only workbook.
//
// input: weekStart "YYYY-MM-DD" (a Sunday, Pacific). Returns JSON text {subject, html, total, flagged}.

type Cell = string | number | boolean;
type Rec = { [k: string]: string };
const WATCH = 10; // unscheduled lookups by one person in one week worth a look

function str(v: Cell | undefined | null): string { return v === undefined || v === null ? '' : String(v).trim(); }
function esc(s: string): string { return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as { [k: string]: string })[c]); }
function addDays(iso: string, n: number): string { const d = new Date(`${iso}T12:00:00Z`); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }

/** Pure: rows (tblLookupLog as objects, When = ISO UTC) -> the week's counts. Days are Pacific calendar days. */
function summarize(rows: Rec[], weekStart: string, offsetHours: number): { subject: string, html: string, total: number, flagged: string[] } {
  const end = addDays(weekStart, 7);
  const local = (iso: string): string => { const t = Date.parse(iso); return isNaN(t) ? '' : new Date(t + offsetHours * 36e5).toISOString().slice(0, 10); };
  const week = rows.filter((r) => { const d = local(str(r.When)); return d >= weekStart && d < end; });
  const by: { [clinic: string]: { [outcome: string]: number } } = {};
  const staff: { [k: string]: number } = {};
  for (const r of week) {
    const c = str(r.Clinic) || '(bad key)', o = str(r.Outcome) || '(none)';
    by[c] = by[c] || {}; by[c][o] = (by[c][o] || 0) + 1;
    if (str(r.Tier) && str(r.Tier) !== 'scheduled') { const k = `${c} · ${str(r.Staff) || '(no name)'}`; staff[k] = (staff[k] || 0) + 1; }
  }
  const outcomes = ['ok', 'need-reason', 'need-note', 'not-found', 'queue', 'already-there', 'bad-mrn', 'bad-key'];
  const flagged = Object.keys(staff).filter((k) => staff[k] >= WATCH).sort();
  const rowsHtml = Object.keys(by).sort().map((c) => `<tr><td>${esc(c)}</td>${outcomes.map((o) => `<td>${by[c][o] || 0}</td>`).join('')}</tr>`).join('');
  const staffHtml = Object.keys(staff).sort((a, b) => staff[b] - staff[a]).map((k) => `<li>${esc(k)}: ${staff[k]}${staff[k] >= WATCH ? ' (look at these in the log)' : ''}</li>`).join('');
  const html = `<p>Clinic lookups for the week of ${weekStart} (Sunday to Saturday): ${week.length}.</p>`
    + `<table border="1" cellpadding="4" style="border-collapse:collapse"><tr><th>Clinic</th>${outcomes.map((o) => `<th>${o}</th>`).join('')}</tr>${rowsHtml || `<tr><td colspan="${outcomes.length + 1}">No lookups.</td></tr>`}</table>`
    + `<p>Lookups for patients without a visit booked at that clinic, by person:</p><ul>${staffHtml || '<li>None.</li>'}</ul>`
    + '<p>Counts only. The MRNs, reasons and notes are in Clinic Lookup Log.xlsx (admin only).</p>';
  return { subject: `Clinic lookups, week of ${weekStart}: ${week.length}${flagged.length ? `, ${flagged.length} to look at` : ''}`, html, total: week.length, flagged };
}

function main(workbook: ExcelScript.Workbook, weekStart: string, offsetHours: number): string {
  const t = workbook.getTable('tblLookupLog');
  if (!t) return JSON.stringify({ subject: 'Clinic lookups: the log table is missing', html: '<p>tblLookupLog is not in this workbook.</p>', total: 0, flagged: [] });
  const h = t.getHeaderRowRange().getValues()[0].map((x) => str(x));
  const rows = (t.getRangeBetweenHeaderAndTotal().getValues() as Cell[][]).filter((r) => r.some((c) => str(c) !== ''))
    .map((r) => { const o: Rec = {}; h.forEach((k, i) => { o[k] = str(r[i]); }); return o; });
  return JSON.stringify(summarize(rows, weekStart, offsetHours));
}
