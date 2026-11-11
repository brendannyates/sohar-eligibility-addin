// Builds "Future state - Clinic Workbook.html" (no PHI; made-up nothing: only system names). node make-future.mjs <out.html>
import { writeFileSync } from 'node:fs';
const out = process.argv[2];
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const W = 178, H = 62, GX = 232, GY = 84, PAD = 16;
// state: live (in use today), built (code merged or ready, set-up pending), planned. names: true where patient names exist.
function svg(d) {
  const pos = {};
  const rows = Math.max(...d.cols.map((c) => c.length));
  d.cols.forEach((col, i) => col.forEach((n, j) => { const off = (rows - col.length) / 2; pos[n.id] = { x: PAD + i * GX, y: PAD + (j + off) * GY, n }; }));
  const width = PAD * 2 + (d.cols.length - 1) * GX + W, height = PAD * 2 + (rows - 1) * GY + H;
  const edges = d.edges.map(([a, b, kind, label]) => {
    const A = pos[a], B = pos[b];
    const back = B.x < A.x, same = B.x === A.x;
    const x1 = same ? A.x + W / 2 : back ? A.x : A.x + W, y1 = same ? (B.y > A.y ? A.y + H : A.y) : A.y + H / 2;
    const x2 = same ? B.x + W / 2 : back ? B.x + W : B.x, y2 = same ? (B.y > A.y ? B.y : B.y + H) : B.y + H / 2;
    const mx = (x1 + x2) / 2;
    const path = same ? `M${x1},${y1} L${x2},${y2}` : `M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`;
    const lx = same ? x1 + 6 : mx, ly = same ? (y1 + y2) / 2 : (y1 + y2) / 2 - 6;
    return `<path class="e e-${kind}" d="${path}" marker-end="url(#ar-${kind})"/>${label ? `<text class="el" x="${lx}" y="${ly}" text-anchor="${same ? 'start' : 'middle'}">${esc(label)}</text>` : ''}`;
  }).join('');
  const nodes = Object.values(pos).map(({ x, y, n }) => `<g class="n n-${n.state}${n.names ? ' n-names' : ''}"><rect x="${x}" y="${y}" width="${W}" height="${H}" rx="10"/>
    <text class="nt" x="${x + 10}" y="${y + 22}">${esc(n.label)}${n.names ? ' 🔒' : ''}</text>${n.sub ? `<text class="ns" x="${x + 10}" y="${y + 42}">${esc(n.sub)}</text>` : ''}${n.sub2 ? `<text class="ns" x="${x + 10}" y="${y + 56}">${esc(n.sub2)}</text>` : ''}</g>`).join('');
  const markers = ['manual', 'flow', 'pq', 'addin'].map((k) => `<marker id="ar-${k}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" class="m-${k}"/></marker>`).join('');
  return `<svg viewBox="0 0 ${width} ${height}" role="img" aria-labelledby="${d.id}-t"><title id="${d.id}-t">${esc(d.title)}</title><defs>${markers}</defs>${edges}${nodes}</svg>`;
}
const N = (id, label, sub, state, extra) => ({ id, label, sub, state, ...(extra || {}) });

const D = [
  { id: 'tickets', title: 'Billing tickets', lede: 'BillingMaster stays the billing team\'s system of record. The master reads its open tickets each morning, and financial notes stay in the master.',
    cols: [
      [N('form', 'Microsoft Form', 'clinic staff submit', 'live')],
      [N('raw', 'Raw Data', 'Forms table', 'live'), N('inbox', 'billing@ inbox flow', 'Arietis replies', 'live')],
      [N('bm', 'BillingMaster', 'billing add-in sync', 'live', { sub2: 'digests, clinic recaps' })],
      [N('qt', 'qOpenTickets', 'Power Query, 7:30 AM', 'built', { sub2: 'open only, no notes' })],
      [N('tt', 'tblOpenTickets', 'master, hidden', 'built')],
      [N('pv', 'Patient view', 'Billing and balance', 'built'), N('tab', 'Billing tickets tab', 'by clinic, status, age', 'built')],
      [N('fn', 'Financial Notes', 'master, hidden', 'built', { sub2: 'cards scrubbed' })],
    ],
    edges: [['form', 'raw', 'flow', ''], ['raw', 'bm', 'addin', 'sync'], ['inbox', 'bm', 'flow', ''], ['bm', 'qt', 'pq', 'refresh'], ['qt', 'tt', 'pq', ''], ['tt', 'pv', 'addin', ''], ['tt', 'tab', 'addin', ''], ['pv', 'fn', 'addin', 'notes']],
    steps: ['Clinic staff submit a ticket in the Microsoft Form; Forms writes Raw Data.', 'The billing add-in copies it to BillingMaster; the billing@ flow records Arietis replies. Unchanged.',
      'At 7:30 AM the desktop refresh runs qOpenTickets: open tickets only (not Closed or Resolved, not deleted), without submitter, notes or attachments.',
      'The Sohar patient view shows each patient\'s open tickets with age and SLA; the Billing tickets tab shows them by clinic, status and age.', 'Staff add financial notes in the patient view. They are saved on the master\'s hidden Financial Notes sheet.'] },
  { id: 'ledger', title: 'Eligibility ledger, centralized', lede: 'The rotation keeps writing its CSVs. The master reads them each morning, so nobody imports anything by hand.',
    cols: [
      [N('luma', 'Luma', 'schedule, checklists, copays', 'live'), N('sohar', 'Sohar', 'verification, discovery', 'live')],
      [N('rot', '7 AM rotation', 'Claude, Brendan\'s PC', 'live')],
      [N('csv', 'Eligibility Checks.csv', 'ledger, one row per visit', 'live'), N('hist', 'Eligibility History.csv', 'one row per Sohar pull', 'live'), N('feed', 'Feed CSVs', 'trackers\' own copy', 'live')],
      [N('qq', 'qLedger, qHistory,', 'qCards, qCopays', 'built', { sub2: 'Power Query, 7:30 AM' }), N('trk', 'TMS / ESK / IVK', 'trackers', 'live')],
      [N('mt', 'Master tables', 'ledger, history,', 'built', { sub2: 'card on file, copays' })],
      [N('use', 'Sohar pane', 'patient view, exports', 'built')],
    ],
    edges: [['luma', 'rot', 'flow', ''], ['sohar', 'rot', 'flow', ''], ['rot', 'csv', 'flow', ''], ['rot', 'hist', 'flow', ''], ['rot', 'feed', 'flow', ''], ['csv', 'qq', 'pq', ''], ['hist', 'qq', 'pq', ''], ['feed', 'trk', 'flow', ''], ['qq', 'mt', 'pq', 'refresh'], ['mt', 'use', 'addin', '']],
    steps: ['At 7 AM the rotation reads Luma and Sohar and writes Eligibility Checks.csv and Eligibility History.csv (LASFIR and MRN only).', 'At 7:30 AM the desktop refresh loads them into the master: the ledger, patient history, card on file (newest answer) and copays.',
      'The Sohar pane and patient view read those tables. Allowed amounts stay in the master.', 'The TMS, ESK and IVK trackers keep their own feed, unchanged.'] },
  { id: 'appts', title: 'Appointment list, centralized', lede: 'One NextGen list a day, in one place. Names stay on one very hidden sheet; clinics get their slice with LASFIR only.',
    cols: [
      [N('ng', 'NextGen', 'Appointments Listing', 'live')],
      [N('drop', 'Data Drop', 'NextGen Appointments.xlsx', 'planned', { sub2: 'saved daily (or emailed)', names: true })],
      [N('qa', 'qNGAppts', 'Power Query, 7:30 AM', 'built')],
      [N('nga', 'NG Appointments', 'master, very hidden', 'built', { names: true })],
      [N('build', 'Sohar build', 'Patients, Review', 'live', { sub2: 'Sohar import CSVs', names: true }), N('slice', 'Clinic slice', 'LASFIR, no DOB', 'built')],
      [N('pub', 'Publish flow', 'clinic folder, archive', 'planned')],
      [N('cw', 'Clinic workbook', 'Start Here', 'built', { sub2: 'CTT access views' })],
    ],
    edges: [['ng', 'drop', 'manual', 'export'], ['drop', 'qa', 'pq', ''], ['qa', 'nga', 'pq', 'refresh'], ['nga', 'build', 'addin', 'daily list'], ['nga', 'slice', 'addin', 'export'], ['slice', 'pub', 'flow', ''], ['pub', 'cw', 'flow', '']],
    steps: ['Each morning the NextGen Appointments Listing is saved to Data Drop under one fixed name. By hand for now; an emailed scheduled report plus a flow, if NextGen can schedule it.',
      'At 7:30 AM qNGAppts finds the header row, stops at the footer and loads NG Appointments. That sheet is very hidden: it is the only place names and DOBs sit, because the Sohar import CSVs need them.',
      'In step 1 of the Sohar pane, "Use the daily NextGen list" builds Patients and the Sohar files from it, instead of an upload.',
      'The clinic export takes each clinic\'s rows, puts LASFIR in place of names and drops DOB, guarantor, member and policy numbers, details and staff names.',
      'The Publish flow puts each file in its clinic folder and archives yesterday\'s. The clinic opens Start Here and the Clinic Template Tool\'s templates, access (3NA), cadence and recall.'] },
];

const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Future state: the central eligibility workbook</title>
<style>
:root{color-scheme:light;--bg:#F2F4F8;--surface:#FFFFFF;--ink:#1A1D26;--muted:#5C6373;--line:#D7DBE5;--live:#0E7A3E;--built:#0F6CBD;--planned:#A8520A;
  --manual:#8A4F0F;--flow:#036C70;--pq:#6B3FA0;--addin:#1D6F42;--names:#B42318}
@media (prefers-color-scheme:dark){:root{color-scheme:dark;--bg:#13151A;--surface:#1C1F26;--ink:#E6E8EE;--muted:#9AA1B2;--line:#2D323D;--live:#5FD08A;--built:#6FB6FF;--planned:#F2A866;
  --manual:#E0A05E;--flow:#3FB8BC;--pq:#B79CFF;--addin:#4FBF7E;--names:#F38B80}}
*{box-sizing:border-box}html,body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif}
.wrap{max-width:1640px;margin:0 auto;padding:20px 18px 60px}h1{font-size:22px;font-weight:600;margin:0 0 6px}h2{font-size:18px;font-weight:600;margin:34px 0 6px}
.lede{color:var(--muted);max-width:80ch;margin:0 0 12px}.card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px;overflow-x:auto}
svg{display:block;min-width:900px;width:100%;height:auto}
.n rect{fill:var(--surface);stroke-width:2}.n-live rect{stroke:var(--live)}.n-built rect{stroke:var(--built);stroke-dasharray:0}.n-planned rect{stroke:var(--planned);stroke-dasharray:6 4}
.n-names rect{fill:color-mix(in srgb,var(--names) 10%,var(--surface))}.nt{font-size:13.5px;font-weight:600;fill:var(--ink)}.ns{font-size:11.5px;fill:var(--muted)}
.e{fill:none;stroke-width:2}.e-manual{stroke:var(--manual);stroke-dasharray:5 4}.e-flow{stroke:var(--flow)}.e-pq{stroke:var(--pq)}.e-addin{stroke:var(--addin)}
.m-manual{fill:var(--manual)}.m-flow{fill:var(--flow)}.m-pq{fill:var(--pq)}.m-addin{fill:var(--addin)}.el{font-size:11px;fill:var(--muted)}
.legend{display:flex;flex-wrap:wrap;gap:8px 18px;margin:10px 0 0;font-size:13px;color:var(--muted)}.legend span{display:inline-flex;align-items:center;gap:6px}
.sw{display:inline-block;width:22px;height:0;border-top:2px solid}.bx{display:inline-block;width:16px;height:12px;border:2px solid;border-radius:3px}
ol{margin:10px 0 0;padding-left:22px;max-width:90ch}li{margin:3px 0}
.note{font-size:13px;color:var(--muted);max-width:90ch}
</style></head><body><div class="wrap">
<h1>Future state: the Sohar Eligibility Template as the central workbook</h1>
<p class="lede">Brendan's master holds the appointment list, the eligibility ledger, open billing tickets, balances and financial notes. Clinic workbooks are a Start Here page plus the Clinic Template Tool's appointment and access views for now. Power Query refreshes the master at 7:30 AM in desktop Excel, or anyone presses Data &gt; Refresh All. No flow can refresh Power Query.</p>
<div class="legend" aria-label="Legend">
  <span><i class="bx" style="border-color:var(--live)"></i>Live today</span><span><i class="bx" style="border-color:var(--built)"></i>Built, set-up pending</span><span><i class="bx" style="border-color:var(--planned);border-style:dashed"></i>Planned</span>
  <span><i class="bx" style="border-color:var(--muted);background:color-mix(in srgb,var(--names) 18%,transparent)"></i>Patient names here 🔒</span>
  <span><i class="sw" style="border-color:var(--manual);border-top-style:dashed"></i>By hand</span><span><i class="sw" style="border-color:var(--flow)"></i>Flow or script</span>
  <span><i class="sw" style="border-color:var(--pq)"></i>Power Query refresh</span><span><i class="sw" style="border-color:var(--addin)"></i>Add-in</span>
</div>
${D.map((d, i) => `<h2>${i + 1}. ${esc(d.title)}</h2><p class="lede">${esc(d.lede)}</p><div class="card">${svg(d)}</div><ol>${d.steps.map((s) => `<li>${esc(s)}</li>`).join('')}</ol>`).join('\n')}
<h2>Not decided yet</h2>
<p class="note">Who saves the NextGen list daily, and whether NextGen can email it on a schedule. The NextGen patient balance report (MRN plus balance), or the Radial MCP as the balance source. When Clinic Day (eligibility in clinic files) comes back.</p>
<p class="note">Names exist only on the master's very hidden NG Appointments sheet, the Import sheet and the Sohar import CSVs. No sheet a clinic gets carries names or dates of birth.</p>
</div></body></html>
`;
writeFileSync(out, html);
console.log('wrote', out);
