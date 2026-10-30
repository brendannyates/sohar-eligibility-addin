# Working on sohar-eligibility-addin

Hard rules:
- `core.js`, `rules.js` and `xlsx.js` stay free of Office.js and DOM. Logic changes go there, with a test in `tests/`.
- Extra patient data travels on the non-column `x` property, never as new template columns.
- Add-in sheets are written only through `writeAddinTable` (fail-safe stage-and-swap). A new add-in table goes in `OWNED` in `excel.js`. Never write an add-in table in place, and never let formulas reference those sheets (each save replaces the sheet). Changes to save/read code must keep `tests/excel.test.mjs` green, including the failure sweeps.
- The one exception is a flow-owned table: `Clinic Work Inbox` (`tblClinicWorkInbox`, `excel.js` `INBOX`). A Power Automate flow appends to it, so the add-in never rewrites, swaps, sorts, clears or trims it (a stage-and-swap would race the flow and drop rows). It is not in `OWNED`. The add-in only creates it once (`createClinicInbox`, header only, fail-safe) and reads it (`readClinicInbox`). `tests/excel.test.mjs` runs every mutating export against a workbook holding it and fails if one touches it; keep new write paths in that list. Build guide: `docs/clinic-inbox-flow.md`.
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
  `Pre-Check` (hidden) holds the visit checklist ticks (MRN key); a tick wins over the report's "Appointment Pre-Check
  Complete". The add-in never writes to NextGen. Config "Report uploaded" stamps each report upload for the "Last
  updated" pill.
- The West Coast Patient Responsibility Calculator's code lives in `fees.js` (`calcEstimate`, `calcInputs`,
  `calcPatient`), but its built-in rates never go in this repo: staff drop the calculator .html once, and
  `parseCalculatorHtml` + `mergeSchedules` put the rates on the hidden fee sheets (a saved rate is never changed).
- Never add a network call. `connect-src 'none'` in `index.html` is deliberate; the only remote script is Office.js.
- `VERIFICATION_COLS` and `DISCOVERY_COLS` match the Sohar CSV templates exactly, in order. Do not rename or reorder.
- Output table columns (`PATIENT_COLS`, `REVIEW_COLS`) must match the template workbook. Changing the column count means shipping a new `Sohar_Eligibility_Template.xlsx`; a rename of the same column is applied in place by `replaceTable` (v0.12 Review Per Nbr -> MRN).
- The patient key is the MRN (`core.keyOf`, PER-<Per Nbr> fallback). Key on it, never on Per Nbr, in anything saved or sent to Sohar. `cfg.payerTests` is the Config settings; Payer Tests rows travel as `cfg.payerTestRows`.
- `payer-rules.js` is shared byte-identical with the Clinic Template Tool and ESK Tracker repos: never edit it here alone. The real Payer Credential Rules table (OneDrive CSV, imported to the editable `Payer Credential Rules` sheet) never goes in this repo, a test, a fixture or the demo. Besse approval is a flag only.
- `core.js` and `rules.js` are also copied byte-identical into the Clinic Template Tool (it re-derives Clinic Eligibility files with them). After changing either, copy both into CTT in the same change and check with `git hash-object`. The Clinic Eligibility file (`rules.clinicPack`) never carries patient or guarantor names or DOBs (LASFIR instead), only the one clinic's patients, and never a `tblClinicMeta` table.
- No sheet (or download built from a sheet's columns) carries a patient's or subscriber's name or DOB (Brendan, 10/28): use `LASFIR` and keep the MRN. Names live in memory and on the very hidden `Import`, `Sohar Uploads` and `IR Uploads` sheets (`lockNameSheets`); only the Sohar import CSVs (`VERIFICATION_COLS` / `DISCOVERY_COLS`) carry them. Sheets use the `*_SHEET_COLS` lists and `core.cellOf`. `tests/names.test.mjs` checks every sheet column list.
- Every file is LF (`core.autocrlf=false`, no `.gitattributes`). Never save CRLF: it rewrites whole files in the diff and breaks the byte-identical `git hash-object` check for the shared files.
- Routing changes (what counts as verification vs discovery, scope, the POS rule) need Brendan's explicit OK on the exact rule before coding.

Every release: bump `VERSION` in `app.js`, `package.json`, and the `?v=` on every local script/stylesheet in `index.html`, `dashboard.html` and `save.html`. Excel caches add-in files; without a new `?v=`, users keep running the old code.

Workflow: edit → `npm test` → `npm run serve` and check at about 350px and at full width (`?tab=clinic` opens the Clinic tab in demo mode) → commit with a plain-English message → push `main` → confirm the `pages` action is green.

Tell Brendan at the end: what changed, "Live now, close and reopen the pane", and whether `manifest.xml` or the template workbook changed (either one means a re-upload or a new workbook copy).
