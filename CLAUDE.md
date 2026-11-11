# Working on sohar-eligibility-addin

Hard rules:
- `core.js`, `rules.js` and `xlsx.js` stay free of Office.js and DOM. Logic changes go there, with a test in `tests/`.
- Extra patient data travels on the non-column `x` property, never as new template columns.
- Add-in sheets are written only through `writeAddinTable` (fail-safe stage-and-swap). A new add-in table goes in `OWNED` in `excel.js`. Never write an add-in table in place, and never let formulas reference those sheets (each save replaces the sheet). Changes to save/read code must keep `tests/excel.test.mjs` green, including the failure sweeps.
- The one exception is a flow-owned table: `Clinic Work Inbox` (`tblClinicWorkInbox`, `excel.js` `INBOX`). A Power Automate flow appends to it, so the add-in never rewrites, swaps, sorts, clears or trims it (a stage-and-swap would race the flow and drop rows). It is not in `OWNED`. The add-in only creates it once (`createClinicInbox`, header only, fail-safe) and reads it (`readClinicInbox`). `tests/excel.test.mjs` runs every mutating export against a workbook holding it and fails if one touches it; keep new write paths in that list. Build guide: `docs/clinic-inbox-flow.md`.
  v0.29: the rotation's tables are flow-owned the same way (`excel.js` `FLOW_TABLES`: `tblEligibilityChecks` (the ledger),
  `tblPatientHistory`, `tblCopayActivity`, `tblCardOnFile`, `tblNeedsCheck`; columns in `rules.LEDGER_COLS` and friends).
  Created once, hidden (`createFlowTables`), read with `readFlowTable`, never written by the add-in. The same sweep covers them.
  Build sheets: `docs/flows/`.
- v0.29 clinic file v2 (`rules.clinicPack`): every patient with ANY upcoming visit at the location (guests too), only that
  clinic's appointments in the record; very hidden `SE Appointments`, `SE Estimates` (ledger rows, never `Allowed`),
  `SE History` (90 days, at most 5 per patient, scrubbed), `SE Copay` (card on file, newest OK to collect), `SE Fees` /
  `SE Fee Rules` (the clinic's fee tab and its rules, split across rows under the cell limit); the visible `Clinic Day`
  sheet (`CLINIC_DAY_COLS`, no MRN) is derived by `clinicDay` and CTT rewrites it. The structure lock is a hash only
  (`XL.protectionHash`); the password is typed in the pane each export and never stored, logged or sent. Clinic Work kinds
  `OK to collect` and `Collected` apply to the add-in-owned `Clinic Consent` and `Desk Collections` tables.
- Synthetic data only in tests, fixtures, demo mode, and commits. No real names, MRNs, member IDs, or exports.
  The one exception is the generated block in `payer-knowledge.js` (Brendan, 10/28): real payer names and Sohar payer
  IDs, no patient data, notes or contracted rates. Regenerate it with `node scripts/build-payer-knowledge.mjs <master
  workbook>` (the workbook stays outside the repo); never hand-edit it. The bar lives in `payer-knowledge.js`
  (`qualifies` / `select`). Workbook sheets are overrides: a sheet row wins, so Payer Map and Sohar Payers are optional
  tables (`templateCheck` skips them) and the Payer Map is created on its first save. The fee schedule never goes in it.
- `record.js` (v0.20) is pure too: one record per patient, Sohar Flow results, report profiles (gender identity,
  current gender, pronouns, guarantor) read by header, and the guarantor vs Sohar subscriber check. Profiles and checks
  reach the full-screen window as LASFIR and text only, and never go on a sheet. `Flow Results` (hidden) holds the
  Flow's CSV, which has LASFIR and MRN but no names.
  `Pre-Check` (hidden) holds the visit checklist ticks and the patient notes (one list per patient, never pruned; MRN key, LASFIR, the admin's name from Config "Admin
  names"); a tick wins over the report's "Appointment Pre-Check Complete". The add-in never writes to NextGen. Config "Report uploaded" stamps each report upload for the "Last
  updated" pill.
  The full-screen patient view (v0.22, sidebar + one section at a time) only lays out what `record.js` derives:
  `plansOf` (payer + member ID), `milestones` (first seen), `sinceLast`, `projection` (booked Flow estimates applied to
  the deductible and OOP max; sandbox and missing estimates left out) and `actionItems` / `leadLine`. New derivations go
  there, with a test, not in `dashboard.js`.
  v0.25: the pane and the main screen use the patient view's tokens (`--t-*` and the status colors on `:root` in
  `style.css`; the older names like `--pine` and `--ink` point at them). Prep by appointment is the handoff's list
  (one row button per patient). Style new screens with these tokens, not new colors.
- `audit.js` (v0.23, the patient view's Account audit) is pure too. The InstaMed, Arietis and NextGen files it reads carry
  names: they stay in the full-screen window's memory only (never a sheet, never `toPane`, never localStorage), and
  the copied summary is LASFIR + MRN with no names, file names or card numbers. Matching and balance rules live in
  `runAudit` with tests in `tests/audit.test.mjs`; the demo files come from `audit.demoRows` (made-up data).
  v0.23.1: "Save this audit" keeps the RESULTS, not the files: one summary row (`AUDIT_COLS`) and its lines
  (`AUDIT_LINE_COLS`) per patient on the hidden `Account Audits` / `Account Audit Lines` sheets (LASFIR + MRN, card
  digits scrubbed by `auditRecord`); saving again replaces that patient's rows (`mergeAudits`). "Export" builds one
  .xlsx in full screen (`exportSpec`: Summary, Findings, Visits, Re-quote, Payments).
  v0.24: `runAudit` balance.lines lists what each balance is made of (`chargeRow`: date of service, code, billed,
  plan paid, PR-1 deductible / PR-2 coinsurance / PR-3 copay and any other PR code, patient paid, balance); each
  balance in the patient view opens to it. 90867 is the MT (motor threshold) visit, 90869 the MTR (MT redo).
- The West Coast Patient Responsibility Calculator's code lives in `fees.js` (`calcEstimate`, `calcInputs`,
  `calcPatient`), but its built-in rates never go in this repo: staff drop the calculator .html once, and
  `parseCalculatorHtml` + `calculatorWins` put the rates on the hidden fee sheets. v0.29 (Brendan, 11/10): the calculator
  is the source of truth, so its rate replaces a saved one for the same tab + payer + level + code, each change goes on
  the hidden `Fee Schedule Log`, and rows only the saved schedule has (Gross, Cash, codes the calculator lacks) stay.
- Per-code rules (v0.24): an Estimate Defaults row with a `Code` says how that code counts for the fee payer and
  `Plan` (Sohar's plan name; Brendan: payer + plan). Most specific wins: payer + plan, then the payer with no Plan,
  then "All payers" (`CODE_CLASSES`: deductible applies or waived, coinsurance, copay). A row with no Code and no Plan
  is the plan-level default, as before. Staff classify codes in the patient view (saved with their name, `rememberCodeRule`). `estimateMath`,
  `calcEstimate` / `calcPatient` / `calcSeries` and the projection all follow them; with no rule, results are unchanged
  and the projection keeps its copay rule (an estimate equal to the copay counts toward the OOP max only).
- v0.26 (Brendan, 11/3: cut the noise). NextGen tasks (`core.buildTasks(results, {flags, guarantor, patients})`) are
  only for what is new or must change before the visit: coverage a discovery found, a policy newly ineligible against
  its earlier result (or ending before the visit), deductible / OOP max newly met against earlier results for the
  policy in the same plan year, the Medi-Cal / SFHP / credential / Kaiser auth flags (`TASK_FLAGS`), and a NextGen
  guarantor who is a different person from Sohar's subscriber. Never a task that restates benefits or stays open
  forever (unmatched results, rejected member IDs, no coverage found, payer mapping): those stay on the Worklist and
  Review. Interventions (TMS 90868, MT 90867, MTR 90869, esketamine G2083) are priced by code; `fees.upgradeInterventionSeeds`
  moves untouched old "none" rows. The back-room tabs go hidden (`excel.tidyTabs` / `TIDY_SHEETS`, OWNED `hidden`);
  account audit summaries and lines share the hidden Account Audits sheet (`audit.AUDIT_SHEET_COLS`). Learned,
  exact and edited rows are compiled into `payer-knowledge.js` by the build script, which is additive (workbook wins on
  the same key; `--fresh` rebuilds); the pane merges built-in Fee Locations, Fee Payer Map, Estimate Codes and Estimate
  Defaults under the sheets and saves only what is not built in (`sheetOnly`). Payer matches are assigned from one
  dropdown in step 1.
- Never add a network call. The one exception (v0.28) is `dev-updates.js` fetching the add-in's own `dev-updates.json`
  from the same origin, so `index.html` and `dashboard.html` have `connect-src 'self'` (`save.html` keeps `'none'`);
  `tests/who.test.mjs` fails on any other `fetch`. The only remote script is Office.js.
- v0.28: on open the pane shows a loading screen (`BOOT` steps in `app.js`), then "Enter your name" when nobody is
  picked today on this computer for this workbook. The workbook's people are in Office document settings
  (`who.USERS_SETTING`, no sheet), plus Config "Admin names"; today's pick is in localStorage per workbook (`who.whoKey`).
  Full screen gets the name in the payload and records it on ticks, notes, code rules and audits; its "Switch user"
  sends `who-pick` to the pane. `dev-updates.js` / `.css` are copied from the IVK Tracker; add a notice to
  `dev-updates.json` (newest first) for each change staff will notice.
  v0.28.2: the Dev Update bell shows only in full screen, on the `.du-home` right of the title (none in the pane).
  Light / dark comes from `theme.js` in each page's `<head>` (localStorage `soharTheme`, else Windows); the dark
  tokens are `:root[data-theme="dark"]` in `style.css`, never a `prefers-color-scheme` query. The pane passes its
  choice to full screen on the URL (`?theme=`) and both send `{ type: 'theme' }` when it changes. Toolbars are
  refresh, full screen, light / dark (full screen's ⟳ sends `refresh` to the pane).
  (The v0.29 clinic lookup webhook is called by the Clinic Template Tool only, never by this add-in.)
- `VERIFICATION_COLS`, `DISCOVERY_COLS` and `record.FLOW_IMPORT_COLS` (the "MHS Eligibility - Discovery vs NextGen" Flow's CSV
  template, v0.21) match the Sohar CSV templates exactly, in order. Do not rename or reorder. The Flow file has no
  rateTable column, so Flow runs from it never estimate cost.
- Output table columns (`PATIENT_COLS`, `REVIEW_COLS`) must match the template workbook. Changing the column count means shipping a new `Sohar_Eligibility_Template.xlsx`; a rename of the same column is applied in place by `replaceTable` (v0.12 Review Per Nbr -> MRN).
- The patient key is the MRN (`core.keyOf`, PER-<Per Nbr> fallback). Key on it, never on Per Nbr, in anything saved or sent to Sohar. `cfg.payerTests` is the Config settings; Payer Tests rows travel as `cfg.payerTestRows`.
- `payer-rules.js` is shared byte-identical with the Clinic Template Tool and ESK Tracker repos: never edit it here alone. The real Payer Credential Rules table (OneDrive CSV, imported to the editable `Payer Credential Rules` sheet) never goes in this repo, a test, a fixture or the demo. Besse approval is a flag only.
- `core.js`, `rules.js` and (v0.29) `fees.js` are also copied byte-identical into the Clinic Template Tool (it re-derives Clinic Eligibility files with them and prices visits from `SE Fees`). After changing any, copy all three into CTT in the same change and check with `git hash-object`. The Clinic Eligibility file (`rules.clinicPack`) never carries patient or guarantor names or DOBs (LASFIR instead), only the one clinic's patients, and never a `tblClinicMeta` table.
- No sheet (or download built from a sheet's columns) carries a patient's or subscriber's name or DOB (Brendan, 10/28): use `LASFIR` and keep the MRN. Names live in memory and on the very hidden `Import`, `Sohar Uploads` and `IR Uploads` sheets (`lockNameSheets`); only the Sohar import CSVs (`VERIFICATION_COLS` / `DISCOVERY_COLS` / `FLOW_IMPORT_COLS`) carry them. Sheets use the `*_SHEET_COLS` lists and `core.cellOf`. `tests/names.test.mjs` checks every sheet column list.
- Every file is LF (`core.autocrlf=false`, no `.gitattributes`). Never save CRLF: it rewrites whole files in the diff and breaks the byte-identical `git hash-object` check for the shared files.
- Routing changes (what counts as verification vs discovery, scope, the POS rule) need Brendan's explicit OK on the exact rule before coding.

Every release: bump `VERSION` in `app.js`, `package.json`, and the `?v=` on every local script/stylesheet in `index.html`, `dashboard.html` and `save.html`. Excel caches add-in files; without a new `?v=`, users keep running the old code.

Workflow: edit → `npm test` → `npm run serve` and check at about 350px and at full width (`?tab=clinic` opens the Clinic tab in demo mode) → commit with a plain-English message → push `main` → confirm the `pages` action is green.

Tell Brendan at the end: what changed, "Live now, close and reopen the pane", and whether `manifest.xml` or the template workbook changed (either one means a re-upload or a new workbook copy).
