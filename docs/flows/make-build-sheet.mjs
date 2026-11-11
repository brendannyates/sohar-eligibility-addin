// node docs/flows/make-build-sheet.mjs -> docs/flows/Clinic Workbook flows - Build Sheet.html
// One self-contained page for the v0.29 flows. The Office Scripts and the rotation client are read from docs/flows at
// build time, so the code on the page is the code the tests run. No PHI, keys or flow URLs in the page or here.
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
const here = (f) => new URL(f, import.meta.url);
const R = createRequire(import.meta.url)('../../rules.js');
const css = readFileSync(here('build-sheet.css'), 'utf8');
const CODE = {
  rotation: readFileSync(here('scripts/rotation-to-master.ts'), 'utf8'),
  lookup: readFileSync(here('scripts/clinic-lookup.ts'), 'utf8'),
  summary: readFileSync(here('scripts/lookup-summary.ts'), 'utf8'),
  push: readFileSync(here('rotation-push.mjs'), 'utf8'),
};
const COLS = { ledger: R.LEDGER_COLS, history: R.PATIENT_HISTORY_COLS, copay: R.COPAY_ACTIVITY_COLS, cards: R.CARD_ON_FILE_COLS, needs: R.NEEDS_CHECK_COLS, log: R.LOOKUP_LOG_COLS };
const data = JSON.stringify({ CODE, COLS }).replace(/<\//g, '<\\/');

const page = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Build sheet: Clinic Workbook flows v0.29</title>
<style>
${css}
.tabs{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 14px}.tabs button{font:inherit;border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:999px;padding:5px 12px;cursor:pointer}
.tabs button[aria-pressed="true"]{background:var(--btn);color:var(--btn-ink);border-color:var(--btn)}
pre.code{background:var(--code);border:1px solid var(--line);border-radius:8px;padding:10px;overflow:auto;max-height:420px;font-size:12.5px;white-space:pre}
table.t{border-collapse:collapse;font-size:13px;margin:6px 0}table.t td,table.t th{border:1px solid var(--line);padding:3px 6px;text-align:left}
.callout{border-left:4px solid var(--warn);background:var(--surface);padding:8px 12px;margin:10px 0;border-radius:6px}
</style></head>
<body><div class="wrap">
<header><h1>Build sheet: Clinic Workbook flows (Sohar v0.29, CTT 1.2.0)</h1>
<div class="actions" style="margin:0"><button class="btn" id="btnExport">Copy config</button><button class="btn" id="btnImportToggle">Load config</button></div></header>
<p class="lede">Five pieces make the San Rafael and Castro clinic files replace the template: the rotation writes into the master, clinic files are published each morning, the clinic lookup answers the Clinic Template Tool, the needs-check loop runs requested checks, and the audit log records every lookup. Each flow is plumbing: the decisions live in the Office Scripts and the add-ins, which have tests. Fill in the left side, then copy from the right side into the designer, card by card. Never type a key or a flow URL into this page.</p>
<div id="importBox" hidden style="margin-bottom:14px"><textarea id="importText" rows="3" placeholder="Paste a config copied earlier, then Load"></textarea>
<div class="actions"><button class="btn p" id="btnImport">Load</button><button class="btn" id="btnImportCancel">Cancel</button></div></div>
<div class="grid"><aside class="cfg" id="cfg"></aside><main id="out"></main></div>
</div>
<script>
const DATA = ${data};
const DEF = {
  site: 'StandardOperationProcedures-PCCPCT (pick it from the list)', library: 'Documents',
  root: '/Shared Documents/Reports & Trackers/Clinic Workbook Pilot/', master: 'Clinic Workbook Pilot/Sohar_Eligibility_Template.xlsx',
  logFile: 'Clinic Workbook Pilot/Admin/Clinic Lookup Log.xlsx', clinics: 'San Rafael, Castro', origin: 'https://brendannyates.github.io',
  owner: '', email: '', testClinic: 'Flow Test',
};
let cfg = Object.assign({}, DEF), tab = 'rotation';
const FIELDS = [
  ['site', 'SharePoint site address'], ['library', 'Document library'], ['root', 'Pilot folder path (as {Path} shows it, ending in /)'],
  ['master', 'Master workbook (path in the library)'], ['logFile', 'Lookup log workbook (admin-only folder)'], ['clinics', 'Pilot clinics, comma separated (exactly as Location is spelled)'],
  ['origin', 'Add-in site (the only origin the lookup answers)'], ['owner', 'Account that owns the flows (Brendan for the pilot)'], ['email', 'Weekly summary goes to'], ['testClinic', 'Test clinic name'],
];
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const q = (s) => String(s == null ? '' : s).replace(/'/g, "''");
const clinics = () => cfg.clinics.split(',').map((s) => s.trim()).filter(Boolean);
const STORE = new Map(); let nId = 0;
const reg = (text) => { const id = 'c' + (++nId); STORE.set(id, text); return id; };
const copyBtn = (text, label) => '<button class="btn" data-copy="' + reg(text) + '">' + (label || 'Copy') + '</button>';
const fx = (text) => '<code>' + esc(text) + '</code> <span class="small">(expression)</span> ' + copyBtn(text);
const txt = (text) => '<code>' + esc(text) + '</code> ' + copyBtn(text);
const pre = (text, label) => '<div class="actions">' + copyBtn(text, label || 'Copy code') + '</div><pre class="code">' + esc(text) + '</pre>';
const HT = ['var(--sched)', 'H', 'Request (premium)'], XL = ['var(--excel)', 'X', 'Excel Online (Business)'], DO = ['var(--dataop)', 'D', 'Data Operation'],
  CT = ['var(--control)', 'C', 'Control'], SP = ['var(--sp)', 'S', 'SharePoint'], OL = ['var(--outlook)', 'O', 'Office 365 Outlook'], SC = ['var(--sched)', 'R', 'Schedule'];
function card(k, kind, title, fields, note, child) {
  return '<div class="card" style="--c:' + k[0] + '"><div class="ch"><span class="ico" style="background:' + k[0] + '">' + k[1] + '</span><div><div class="kind">' + kind + '</div><div class="title">' + esc(title) + '</div><div class="conn">' + k[2] + '</div></div></div>'
    + fields.map((f) => '<div class="f"><div class="l">' + esc(f[0]) + '</div><div class="v">' + f[1] + '</div><div></div></div>').join('')
    + (child || '') + (note ? '<p class="note">' + note + '</p>' : '') + '</div>';
}
const link = '<div class="link"><span>+</span></div>';
const branch = (label, html) => '<div class="branch"><div class="bl">' + label + '</div>' + html + '</div>';
const SECURE = '<b>Settings &gt; Security:</b> Secure inputs On, Secure outputs On, so the run history never shows patient data or keys.';
const xlFile = (file) => [['Location', esc(cfg.site)], ['Document Library', esc(cfg.library)], ['File', 'pick <code>' + esc(file) + '</code> with the file picker']];
const CORS = () => '{"Access-Control-Allow-Origin":"' + cfg.origin + '","Content-Type":"application/json","Cache-Control":"no-store"}';
const tableBlock = (name, cols, note) => '<h3><code>' + esc(name) + '</code></h3><div class="actions">' + copyBtn(cols.join('\\t'), 'Copy for Excel') + '</div><table class="t"><tr>' + cols.map((c) => '<th>' + esc(c) + '</th>').join('') + '</tr></table><p class="small">' + note + '</p>';
const checklist = (items) => '<ul class="check">' + items.map((i) => '<li>' + i + '</li>').join('') + '</ul>';
// A new random key, shown only in this terminal: copy it into the flow's Compose card and ~/.secrets (or the front-desk setup).
const PS_NEWKEY = '$b = [byte[]]::new(24); [Security.Cryptography.RandomNumberGenerator]::Fill($b); [Convert]::ToBase64String($b)   # a new key, shown only here';
const PS_SECRET = (name, what) => '$env:' + name + " = Read-Host -MaskInput '" + what + "'   # typed here, once per terminal; never paste it anywhere else";

/* ---------- 1. Rotation to master ---------- */
function fRotation() {
  const flow = [
    card(HT, 'Trigger', 'When an HTTP request is received', [['Who can trigger the flow?', 'Anyone (the rotation key in the body is the check)'], ['Method', 'POST'],
      ['Request Body JSON Schema', txt('{"type":"object","properties":{"v":{"type":"integer"},"key":{"type":"string"},"op":{"type":"string"},"table":{"type":"string"},"rows":{"type":"array","items":{"type":"object"}}},"required":["key","op"]}')],
      ['Settings: Concurrency control', 'On, Degree of parallelism <b>1</b> (one write into the master at a time)']], SECURE + ' The URL appears on this card after the first Save. It carries a signature: it goes only into ~/.secrets as CLINIC_ROTATION_URL.'),
    link,
    card(DO, 'Action', 'Compose', [['Rename to', '<b>Rotation key</b> (exact: expressions say <code>outputs(\\'Rotation_key\\')</code>)'], ['Inputs', 'type the rotation key here yourself (make one with the PowerShell line in section 4)']], SECURE + ' This is the only place the key lives besides ~/.secrets.'),
    link,
    card(CT, 'Control', 'Condition', [['Rename to', '<b>Key matches</b>'], ['Condition', fx("equals(triggerBody()?['key'], outputs('Rotation_key'))") + ' is equal to <code>true</code>']], 'Put the expression on the left, choose "is equal to", and type true on the right as an expression (fx, then true).',
      branch('True', card(XL, 'Action inside True', 'Run script', xlFile(cfg.master).concat([['Script', 'Rotation to master (section 3)'], ['input', fx("string(removeProperty(triggerBody(), 'key'))")]]), SECURE + ' The script never clears or rewrites a table; it updates or adds one row at a time.')
        + link + card(HT, 'Action inside True', 'Response', [['Status Code', fx("if(equals(json(outputs('Run_script')?['body/result'])?['ok'], true), 200, 400)")], ['Headers', txt('{"Content-Type":"application/json"}')], ['Body', fx("json(outputs('Run_script')?['body/result'])")]], ''))
      + branch('False', card(HT, 'Action inside False', 'Response', [['Status Code', '401'], ['Body', txt('{"ok":false,"error":"bad-key"}')]], 'A wrong key gets nothing back and nothing is written.'))),
  ];
  return {
    intro: 'make.powerautomate.com, Create, Instant cloud flow, skip, then search the trigger <b>When an HTTP request is received</b>. Name it <b>Clinic Workbook - Rotation to master</b>. Four top-level cards. It replaces the CSV hand-off: the rotation posts its rows here after each run, and the same flow answers the needs-check loop (section "Needs-check loop").',
    flow: flow.join(''),
    after: 'After Save: copy the URL from the trigger into ~/.secrets as CLINIC_ROTATION_URL and the key as CLINIC_ROTATION_KEY. Each run shows one Run script call; with secure inputs on, the history shows only that it ran. Run the test in section 4 against a copy of the master first (point Run script at the copy, then back).',
    tables: '<p>The add-in makes these hidden, header-only tables in the master: Sohar pane, Admin, Clinic inbox, <b>Set up the rotation tables</b>. They belong to the flows: never sort them or type in them.</p>'
      + tableBlock('tblEligibilityChecks', DATA.COLS.ledger, 'Row Key is MRN|Appt Date (YYYY-MM-DD)|Time. Appt Date is kept as YYYY-MM-DD text. Allowed stays in the master and never goes to a clinic file.')
      + tableBlock('tblPatientHistory', DATA.COLS.history, 'One row per event (Kind: Verification, Discovery, Card on file, Visit, Copay ...). Row Key is the event id.')
      + tableBlock('tblCopayActivity', DATA.COLS.copay, 'Luma billing charges with amountPaid &gt; 0 and tracker copay lines. Desk collections come back from the clinics through the inbox instead.')
      + tableBlock('tblCardOnFile', DATA.COLS.cards, 'One row per MRN. Card On File is Yes / No from the Luma "Tokenized Card on File" checklist item or a tracker tick (Luma\\'s API lists no saved cards).')
      + tableBlock('tblNeedsCheck', DATA.COLS.needs, 'The lookup flow adds a row; the rotation marks it Done.'),
    code: '<h3>Office Script: Rotation to master</h3><p class="small">Excel on the web, Automate, New script, paste, save as <b>Rotation to master</b> in the flow owner\\'s OneDrive.</p>' + pre(DATA.CODE.rotation)
      + '<h3>Rotation client: rotation-push.mjs</h3><p class="small">For the eligibility-check rotation (Claude Code). Keep writing the CSV for the first two weeks; push after it.</p>' + pre(DATA.CODE.push),
    test: '<p>Make the rotation key once with the first line. Then, against the copy of the master: one synthetic row for the test clinic, and the same row again (it must update, not add).</p>' + pre([
      PS_NEWKEY, PS_SECRET('CLINIC_ROTATION_URL', 'Rotation flow URL'), PS_SECRET('CLINIC_ROTATION_KEY', 'Rotation key'),
      "$row = @{ 'Row Key' = '900901|2026-11-12|9:00 AM'; MRN = '900901'; LASFIR = 'TESPAT'; 'Appt Date' = '2026-11-12'; Time = '9:00 AM'; Location = '" + q(cfg.testClinic) + "'; 'Sohar Status' = 'complete.eligible'; Copay = '30' }",
      "$body = @{ v = 1; key = $env:CLINIC_ROTATION_KEY; op = 'upsert'; table = 'ledger'; rows = @($row) } | ConvertTo-Json -Depth 5",
      "Invoke-RestMethod -Method Post -Uri $env:CLINIC_ROTATION_URL -ContentType 'application/json' -Body $body   # first: added 1. Run again: updated 1, added 0",
      "$bad = @{ v = 1; key = 'wrong'; op = 'needs-open' } | ConvertTo-Json",
      "try { Invoke-RestMethod -Method Post -Uri $env:CLINIC_ROTATION_URL -ContentType 'application/json' -Body $bad } catch { $_.Exception.Response.StatusCode }   # 401",
    ].join('\\n'), 'Copy test'),
    check: checklist(['The flow is owned by ' + esc(cfg.owner || 'the pilot owner') + ' and a co-owner is added (for the move to a service account later).', 'Trigger concurrency is 1.', 'Secure inputs and outputs are on for the trigger, Rotation key and Run script.', 'The five tables exist in the master (Set up the rotation tables).', 'The test added one row and then updated it; the 401 test returned 401.', 'The test row is deleted from the copy, and Run script points at the real master.', 'Microsoft says a flow writing to a workbook open in Excel is not supported: run the test once with the master open in Excel on the web and once closed, and check the row both times.', 'Run script has a daily limit per account (about 1,600 calls): batches of 200 rows keep a full rotation under 10 calls.']),
    fallback: '<p>If the HTTP trigger is blocked (tenant policy or license), keep the CSV: the rotation already writes Eligibility Checks.csv. A Recurrence flow at 7:30 AM can read it from OneDrive (Get file content) and pass the text to a Run script that parses it. That parsing script is not written yet; until then the clinic files carry what the master already holds (Sohar Results), without the ledger estimates.</p>',
  };
}

/* ---------- 2. Publish clinic files (the master to clinic refresh) ---------- */
function fPublish() {
  const name = "triggerBody()?['{FilenameWithExtension}']";
  const loc = "trim(substring(" + name + ", 21, sub(lastIndexOf(" + name + ", ' '), 21)))";
  const flow = [
    card(SP, 'Trigger', 'When a file is created (properties only)', [['Site Address', esc(cfg.site)], ['Library Name', esc(cfg.library)], ['Folder', '<code>' + esc(cfg.root) + 'Exports</code> (admin-only)'],
      ['Trigger condition', txt("@and(startsWith(" + name + ", 'Clinic Eligibility - '), endsWith(" + name + ", '.xlsx'))")]], 'Only files the master exports start a run. Files named TEST FILE never do. Paste the condition as text under Settings.'),
    link,
    card(DO, 'Action', 'Compose', [['Rename to', '<b>Location</b>'], ['Inputs', fx(loc)]], '"Clinic Eligibility - " is 21 characters; the date after the last space is dropped. "Clinic Eligibility - San Rafael 2026-11-11.xlsx" gives San Rafael.'),
    link,
    card(CT, 'Control', 'Condition', [['Rename to', '<b>Pilot clinic</b>'], ['Condition', fx("contains(createArray(" + clinics().concat(cfg.testClinic).map((c) => "'" + q(c) + "'").join(', ') + "), outputs('Location'))") + ' is equal to <code>true</code>']], 'A location outside the pilot is left in Exports.',
      branch('True', card(SP, 'Action inside True', 'Get files (properties only)', [['Site Address', esc(cfg.site)], ['Library Name', esc(cfg.library)], ['Limit Entries to Folder', fx("concat('" + q(cfg.root) + "', outputs('Location'))")], ['Include Nested Items', 'No'], ['Filter Query', txt("startswith(FileLeafRef,'Clinic Eligibility - ')")]], '')
        + link + card(CT, 'Action inside True', 'Apply to each', [['Select an output', fx("body('Get_files_(properties_only)')?['value']")]], '',
          card(SP, 'Action inside the loop', 'Move file', [['Current Site Address', esc(cfg.site)], ['File to Move', fx("items('Apply_to_each')?['{Identifier}']")], ['Destination Site Address', esc(cfg.site)], ['Destination Folder', fx("concat('" + q(cfg.root) + "', outputs('Location'), '/Archive')")], ['If another file is already there', 'Replace']], 'Yesterday\\'s file goes to the clinic\\'s Archive folder, which the inbox flow ignores. Its Clinic Work rows already reached the inbox when they were saved.'))
        + link + card(SP, 'Action inside True', 'Copy file', [['Current Site Address', esc(cfg.site)], ['File to Copy', fx("triggerBody()?['{Identifier}']")], ['Destination Site Address', esc(cfg.site)], ['Destination Folder', fx("concat('" + q(cfg.root) + "', outputs('Location'))")], ['If another file is already there', 'Replace']], 'The new file keeps its exact name, Location and date, as the inbox flow expects.')
        + link + card(SP, 'Action inside True', 'Move file', [['File to Move', fx("triggerBody()?['{Identifier}']")], ['Destination Folder', '<code>' + esc(cfg.root) + 'Exports/Published</code>'], ['If another file is already there', 'Replace']], ''))
      + branch('False', '<p class="small">Nothing. The file stays in Exports.</p>')),
  ];
  return {
    intro: 'Automated cloud flow, trigger <b>When a file is created (properties only)</b>. Name it <b>Clinic Workbook - Publish clinic files</b>. This is the master to clinic refresh, and it never writes into a workbook: it moves whole files. You still export the files from the Sohar pane (with the lock password); save them into <code>Exports</code> and this flow puts each one in its clinic folder and archives yesterday\\'s.',
    flow: flow.join(''),
    after: 'Export before the clinics open (before 7 AM). A clinic that has yesterday\\'s file open sees it move to Archive; Excel on the web keeps it open, so tell staff to open the new one from the clinic folder each morning. The history push happens here: SE History rides in each new file, so nothing writes into a file staff have open.',
    tables: '<p>No tables. Folders: <code>Exports</code> and <code>Exports/Published</code> in the pilot folder (admin only), and <code>Archive</code> in each clinic folder (that clinic\\'s permissions).</p>',
    code: '<p>No code. The export is the Sohar pane (Admin, Clinic files, Clinic Template Tool file).</p>',
    test: '<p>Use the test clinic: export a synthetic Flow Test file with the demo data, rename it to the Clinic Eligibility pattern for ' + esc(cfg.testClinic) + ', and drop it into Exports. The Flow Test folder should end with the new file, the old one in Flow Test/Archive, and the export in Exports/Published.</p>',
    check: checklist(['Exports and Exports/Published are admin-only; each clinic folder (and its Archive) has only that clinic\\'s team.', 'The trigger condition is pasted as text and starts with @.', 'A file for a clinic outside the list stays in Exports.', 'The inbox flow still ignores files in Archive (its folder check).', 'Exports run before 7 AM.']),
    fallback: '<p>The pilot\\'s manual export: save each file straight into its clinic folder and move yesterday\\'s to Archive by hand.</p>',
  };
}

/* ---------- 3. Clinic lookup ---------- */
function fLookup() {
  const askFx = "json(base64ToString(triggerBody()?['$content']))";
  const inputFx = "string(setProperty(setProperty(setProperty(removeProperty(outputs('Ask'), 'key'), 'clinic', outputs('Clinic')), 'today', formatDateTime(convertFromUtc(utcNow(), 'Pacific Standard Time'), 'yyyy-MM-dd')), 'now', utcNow()))";
  const logRow = (src) => R_LOG.map((c) => [c, fx(src(c))]);
  const R_LOG = DATA.COLS.log;
  const okLog = logRow((c) => "outputs('Answer')?['log']?['" + c + "']");
  const badLog = R_LOG.map((c) => [c, c === 'When' ? fx('utcNow()') : c === 'Staff' ? fx("outputs('Ask')?['staff']") : c === 'Outcome' ? txt('bad-key') : c === 'Request Id' ? fx("outputs('Ask')?['requestId']") : '(leave blank)']);
  const flow = [
    card(HT, 'Trigger', 'When an HTTP request is received', [['Who can trigger the flow?', 'Anyone (the clinic key in the body is the check)'], ['Method', 'POST'], ['Request Body JSON Schema', 'leave blank (the add-in sends text/plain)'], ['Settings: Concurrency control', 'On, Degree of parallelism 5']], SECURE + ' The add-in posts text/plain so the browser sends no preflight. The URL goes only to the admin, who types it into each front-desk computer (Setup tab, Patient lookup).'),
    link,
    card(DO, 'Action', 'Compose', [['Rename to', '<b>Clinic keys</b>'], ['Inputs', 'type: <code>{"&lt;' + clinics().concat(cfg.testClinic).map((c) => esc(c) + ' key&gt;":"' + esc(c)).join('","&lt;') + '"}</code>, with each clinic\\'s own key in place of &lt;… key&gt;']], SECURE + ' The key decides the clinic; the request never names one. A lost key is changed here and on that clinic\\'s computers.'),
    link,
    card(DO, 'Action', 'Compose', [['Rename to', '<b>Ask</b>'], ['Inputs', fx(askFx)]], SECURE + ' If the first run fails here and the trigger output shows the body as plain JSON (no $content), use <code>json(string(triggerBody()))</code> instead.'),
    link,
    card(DO, 'Action', 'Compose', [['Rename to', '<b>Clinic</b>'], ['Inputs', fx("coalesce(outputs('Clinic_keys')?[outputs('Ask')?['key']], '')")]], 'Secure inputs On (they hold the keys).'),
    link,
    card(CT, 'Control', 'Condition', [['Rename to', '<b>Key known</b>'], ['Condition', fx("empty(outputs('Clinic'))") + ' is equal to <code>false</code>']], '',
      branch('True', card(XL, 'Action inside True', 'Run script', xlFile(cfg.master).concat([['Script', 'Clinic lookup (section 3)'], ['input', fx(inputFx)]]), SECURE + ' The script decides the tier, the fields and whether to queue a check. It reads the master and writes only a Needs Check row.')
        + link + card(DO, 'Action inside True', 'Compose', [['Rename to', '<b>Answer</b>'], ['Inputs', fx("json(outputs('Run_script')?['body/result'])")]], SECURE)
        + link + card(XL, 'Action inside True', 'Add a row into a table', xlFile(cfg.logFile).concat([['Table', '<code>tblLookupLog</code>']], okLog), 'Logged before answering: if the log cannot be written, the run fails and the add-in shows "The lookup could not be reached". No result values go in the log.')
        + link + card(HT, 'Action inside True', 'Response', [['Status Code', fx("outputs('Answer')?['status']")], ['Headers', txt(CORS())], ['Body', fx("outputs('Answer')?['body']")]], 'Access-Control-Allow-Origin is the add-in\\'s site only, so no other web page can read an answer.'))
      + branch('False', card(XL, 'Action inside False', 'Add a row into a table', xlFile(cfg.logFile).concat([['Table', '<code>tblLookupLog</code>']], badLog), '') + link
        + card(HT, 'Action inside False', 'Response', [['Status Code', '401'], ['Headers', txt(CORS())], ['Body', txt('{"ok":false,"error":"bad-key"}')]], ''))),
  ];
  return {
    intro: 'Instant cloud flow, trigger <b>When an HTTP request is received</b>, name <b>Clinic Workbook - Clinic lookup</b>. Six top-level cards. This is the only address the Clinic Template Tool calls (Compliance: see the Compliance note). It answers on demand and nothing it returns is saved in the clinic file. No Teams notices.',
    flow: flow.join(''),
    after: 'After Save, send Brendan\\'s Claude session only the host part of the URL (the part before /powerautomate or /workflows, never the rest): CTT\\'s lookup.js LOOKUP_HOST and the connect-src in taskpane.html are set to it in one change, and lookups switch on with the next CTT release.',
    tables: tableBlock('tblLookupLog', DATA.COLS.log, 'In ' + esc(cfg.logFile) + ', admin only. One row per call, bad keys included. Reason holds the reason and the note; no coverage, copay or other result values.'),
    code: '<h3>Office Script: Clinic lookup</h3><p class="small">Save as <b>Clinic lookup</b>. The tiers: on this clinic\\'s list (no reason); seen here in 12 months (reason); in the master but not here (reason and note); not in the master (request a check).</p>' + pre(DATA.CODE.lookup),
    test: '<p>Make each clinic\\'s key with the first line (one run per clinic, plus one for the test clinic). Then use the test clinic\\'s key for three scenarios. Each must log a row.</p>' + pre([
      PS_NEWKEY, PS_SECRET('CLINIC_LOOKUP_URL', 'Clinic lookup URL'), PS_SECRET('CLINIC_LOOKUP_KEY', cfg.testClinic + ' key'),
      "function Ask($o) { $o.key = $env:CLINIC_LOOKUP_KEY; $o.staff = 'Flow test'; $o.requestId = [guid]::NewGuid().ToString(); Invoke-RestMethod -Method Post -Uri $env:CLINIC_LOOKUP_URL -ContentType 'text/plain;charset=UTF-8' -Body ($o | ConvertTo-Json) }",
      "Ask @{ mrn = '900901'; purpose = 'scheduled' }                                     # tier scheduled, ok true (after the rotation test row)",
      "Ask @{ mrn = '900999'; purpose = 'unscheduled' }                                   # ok false, error need-reason",
      "Ask @{ mrn = '900999'; purpose = 'request-check'; reason = 'Walk-in or same-day add' }  # queued true; a Needs Check row appears",
    ].join('\\n'), 'Copy test'),
    check: checklist(['Every Compose that holds keys or patient data, the trigger and Run script have secure inputs and outputs on.', 'Headers allow only ' + esc(cfg.origin) + '.', 'Each clinic has its own key; the test clinic\\'s key is removed after testing.', 'Each test call added a row to tblLookupLog, including a wrong-key call.', 'The Log workbook sits in an admin-only folder.', 'Compliance has read the Compliance note (BAA for Power Automate + SharePoint, the network exception, the log).', 'Brendan has the host for LOOKUP_HOST; the full URL is only on the front-desk computers.']),
    fallback: '<p>Without the HTTP trigger, there is no live lookup. The clinic files still carry every scheduled patient\\'s summary and history; a patient who is not on the list waits for the next export after the clinic asks the admin.</p>',
  };
}

/* ---------- 4. Needs-check loop ---------- */
function fNeeds() {
  return {
    intro: 'No new flow: the loop uses the two above. A front desk asks for a check in the lookup (not in the master) and the lookup adds an Open row to tblNeedsCheck. Each morning the rotation asks the Rotation to master flow for the open rows (<code>op: needs-open</code>), runs them with the day\\'s discovery and verification (production only with Brendan\\'s OK, as always), and marks them Done (<code>op: needs-done</code>). The result reaches the clinic with the next file.',
    flow: card(HT, 'Uses', 'Clinic Workbook - Rotation to master', [['needs-open', txt('{"v":1,"key":"<rotation key>","op":"needs-open"}')], ['needs-done', txt('{"v":1,"key":"<rotation key>","op":"needs-done","rows":[{"Row Key":"<id>","Status":"Done","Done At":"<ISO>","Result":"Discovery sent"}]}')]], 'needs-done changes only Status, Done At and Result, and never adds a row. A request the rotation cannot run (no name or DOB on any report) stays Open with a Result saying why.'),
    after: 'The rotation needs a name and DOB to send a discovery. They come from the NextGen report, never from the clinic: a patient on no report stays Open, and the Result says to add them to the next report.',
    tables: '<p>tblNeedsCheck (section 2 of Rotation to master).</p>',
    code: '<h3>In rotate.js</h3>' + pre(["import { needsOpen, needsDone } from '<add-in repo>/docs/flows/rotation-push.mjs';", 'const open = await needsOpen(process.env);', '// add each open MRN to this run when the report has the patient; then:', "await needsDone(open.filter(r => sentThisRun.has(r.MRN)).map(r => ({ 'Row Key': r['Row Key'], Result: 'Discovery sent' })), process.env);"].join('\\n')),
    test: pre([PS_SECRET('CLINIC_ROTATION_URL', 'Rotation flow URL'), PS_SECRET('CLINIC_ROTATION_KEY', 'Rotation key'),
      "Invoke-RestMethod -Method Post -Uri $env:CLINIC_ROTATION_URL -ContentType 'application/json' -Body (@{ v = 1; key = $env:CLINIC_ROTATION_KEY; op = 'needs-open' } | ConvertTo-Json)   # lists the lookup test's request"].join('\\n'), 'Copy test'),
    check: checklist(['The lookup test\\'s request shows in needs-open.', 'After needs-done it no longer does, and the row keeps its MRN and requester.', 'Production checks still need Brendan\\'s OK.']),
    fallback: '<p>Without the flows, the clinic asks the admin, who adds the patient to the next NextGen report.</p>',
  };
}

/* ---------- 5. Audit log ---------- */
function fAudit() {
  const flow = [
    card(SC, 'Trigger', 'Recurrence', [['Interval', '1'], ['Frequency', 'Week'], ['On these days', 'Sunday'], ['At these hours', '7'], ['Time zone', '(UTC-08:00) Pacific Time (US &amp; Canada)']], 'Weeks run Sunday to Saturday; Sunday\\'s run covers the week just ended.'),
    link,
    card(DO, 'Action', 'Compose', [['Rename to', '<b>Week start</b>'], ['Inputs', fx("formatDateTime(addDays(convertFromUtc(utcNow(), 'Pacific Standard Time'), -7), 'yyyy-MM-dd')")]], ''),
    link,
    card(XL, 'Action', 'Run script', xlFile(cfg.logFile).concat([['Script', 'Lookup summary'], ['weekStart', fx("outputs('Week_start')")], ['offsetHours', fx("div(sub(ticks(convertFromUtc(utcNow(), 'Pacific Standard Time')), ticks(utcNow())), 36000000000)")]]), 'offsetHours is -8 or -7 (daylight time), so days fall on the Pacific calendar.'),
    link,
    card(OL, 'Action', 'Send an email (V2)', [['To', esc(cfg.email || 'the pilot owner')], ['Subject', fx("json(outputs('Run_script')?['body/result'])?['subject']")], ['Body', fx("json(outputs('Run_script')?['body/result'])?['html']")], ['Importance', 'Normal']], 'Counts only: clinics, outcomes and unscheduled lookups per person. No MRNs or reasons, so email is fine. Ten or more unscheduled lookups by one person in a week are flagged.'),
  ];
  return {
    intro: 'The log itself is written by the lookup flow (tblLookupLog). This scheduled flow, <b>Clinic Workbook - Lookup audit summary</b>, emails a weekly count so somebody reads it. Four cards.',
    flow: flow.join(''),
    after: 'Keep the log workbook as long as Compliance says (HIPAA keeps compliance documentation six years). Review flagged people with their practice lead.',
    tables: tableBlock('tblLookupLog', DATA.COLS.log, 'Make it in ' + esc(cfg.logFile) + ': paste the header in A1, Insert, Table, My table has headers, Table Design, Table Name tblLookupLog. Format the columns as Text.'),
    code: '<h3>Office Script: Lookup summary</h3>' + pre(DATA.CODE.summary),
    test: '<p>Run the flow by hand once (Test, Manually) after the lookup tests: the email should count them under the test clinic.</p>',
    check: checklist(['The log workbook is in an admin-only folder.', 'The email went to ' + esc(cfg.email || 'the owner') + ' and shows no MRNs.', 'A retention period is agreed with Compliance.']),
    fallback: '<p>Without the summary flow, open the log on Mondays and filter Outcome and Tier.</p>',
  };
}

const FLOWS = [['rotation', 'Rotation to master', fRotation], ['publish', 'Publish clinic files', fPublish], ['lookup', 'Clinic lookup', fLookup], ['needs', 'Needs-check loop', fNeeds], ['audit', 'Audit log', fAudit]];
function renderCfg() {
  document.getElementById('cfg').innerHTML = '<h2>What and where</h2>' + FIELDS.map(([k, l]) => '<label class="lbl" for="f_' + k + '">' + esc(l) + '</label><input type="text" id="f_' + k + '" data-k="' + k + '" value="' + esc(cfg[k]) + '">').join('')
    + '<h3>Fixed rules</h3><div class="rules small">Lookup tiers, minimum fields, the 90-day / 5-event history and the 10-a-week flag are in the Office Scripts (tested in the add-in repo), not in this page or a workbook. Keys and URLs never go here.</div>';
}
function render() {
  STORE.clear(); nId = 0;
  const f = FLOWS.find((x) => x[0] === tab), o = f[2]();
  document.getElementById('out').innerHTML = '<div class="tabs" role="group" aria-label="Flows">' + FLOWS.map(([id, label]) => '<button data-tab="' + id + '" aria-pressed="' + (id === tab) + '">' + esc(label) + '</button>').join('') + '</div>'
    + '<h2>1. The flow, step by step</h2><p>' + o.intro + '</p><div class="flow">' + o.flow + '</div><p class="after">' + o.after + '</p>'
    + '<h2>2. Workbook settings</h2>' + o.tables + '<h2>3. Code</h2>' + o.code + '<h2>4. Test</h2>' + o.test
    + '<h2>5. Before you call it done</h2>' + o.check + '<h2>6. Fallback</h2>' + o.fallback;
}
document.addEventListener('input', (e) => { const k = e.target.dataset && e.target.dataset.k; if (k) { cfg[k] = e.target.value; render(); } });
document.addEventListener('click', async (e) => {
  const t = e.target.closest('[data-tab]'); if (t) { tab = t.dataset.tab; render(); return; }
  const c = e.target.closest('[data-copy]');
  if (c) { try { await navigator.clipboard.writeText(STORE.get(c.dataset.copy)); c.textContent = 'Copied'; setTimeout(() => { c.textContent = 'Copy'; }, 1200); } catch (err) { c.textContent = 'Select and copy by hand'; } return; }
  if (e.target.id === 'btnExport') { try { await navigator.clipboard.writeText(JSON.stringify(cfg)); e.target.textContent = 'Config copied'; } catch (err) {} }
  if (e.target.id === 'btnImportToggle') document.getElementById('importBox').hidden = false;
  if (e.target.id === 'btnImportCancel') document.getElementById('importBox').hidden = true;
  if (e.target.id === 'btnImport') { try { cfg = Object.assign({}, DEF, JSON.parse(document.getElementById('importText').value)); renderCfg(); render(); document.getElementById('importBox').hidden = true; } catch (err) { alert('That is not a config copied from this page.'); } }
});
renderCfg(); render();
</script></body></html>
`;
const out = here('Clinic Workbook flows - Build Sheet.html');
writeFileSync(out, page);
console.log('wrote', decodeURIComponent(out.pathname));
