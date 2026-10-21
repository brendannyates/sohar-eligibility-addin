# Working on sohar-eligibility-addin

Hard rules:
- `core.js`, `rules.js` and `xlsx.js` stay free of Office.js and DOM. Logic changes go there, with a test in `tests/`.
- Extra patient data travels on the non-column `x` property, never as new template columns.
- Add-in sheets are written only through `writeAddinTable` (fail-safe stage-and-swap). A new add-in table goes in `OWNED` in `excel.js`. Never write an add-in table in place, and never let formulas reference those sheets (each save replaces the sheet). Changes to save/read code must keep `tests/excel.test.mjs` green, including the failure sweeps.
- Synthetic data only in tests, fixtures, demo mode, and commits. No real names, MRNs, member IDs, or exports.
- Never add a network call. `connect-src 'none'` in `index.html` is deliberate; the only remote script is Office.js.
- `VERIFICATION_COLS` and `DISCOVERY_COLS` match the Sohar CSV templates exactly, in order. Do not rename or reorder.
- Output table columns (`PATIENT_COLS`, `REVIEW_COLS`) must match the template workbook. Changing the column count means shipping a new `Sohar_Eligibility_Template.xlsx`; a rename of the same column is applied in place by `replaceTable` (v0.12 Review Per Nbr -> MRN).
- The patient key is the MRN (`core.keyOf`, PER-<Per Nbr> fallback). Key on it, never on Per Nbr, in anything saved or sent to Sohar. `cfg.payerTests` is the Config settings; Payer Tests rows travel as `cfg.payerTestRows`.
- Routing changes (what counts as verification vs discovery, scope, the POS rule) need Brendan's explicit OK on the exact rule before coding.

Every release: bump `VERSION` in `app.js`, `package.json`, and the `?v=` on every local script/stylesheet in `index.html`, `dashboard.html` and `save.html`. Excel caches add-in files; without a new `?v=`, users keep running the old code.

Workflow: edit → `npm test` → `npm run serve` and check at about 350px and at full width (`?tab=clinic` opens the Clinic tab in demo mode) → commit with a plain-English message → push `main` → confirm the `pages` action is green.

Tell Brendan at the end: what changed, "Live now, close and reopen the pane", and whether `manifest.xml` or the template workbook changed (either one means a re-upload or a new workbook copy).
