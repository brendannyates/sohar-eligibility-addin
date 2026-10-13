# Working on sohar-eligibility-addin

Hard rules:
- `core.js` stays free of Office.js and DOM. Logic changes go there, with a test in `tests/core.test.mjs`.
- Synthetic data only in tests, fixtures, demo mode, and commits. No real names, MRNs, member IDs, or exports.
- Never add a network call. `connect-src 'none'` in `index.html` is deliberate; the only remote script is Office.js.
- `VERIFICATION_COLS` and `DISCOVERY_COLS` match the Sohar CSV templates exactly, in order. Do not rename or reorder.
- Output table columns (`PATIENT_COLS`, `REVIEW_COLS`) must match the template workbook. Changing either means shipping a new `Sohar_Eligibility_Template.xlsx` too.
- Routing changes (what counts as verification vs discovery, scope, the POS rule) need Brendan's explicit OK on the exact rule before coding.

Workflow: edit → `npm test` → `npm run serve` and check at about 350px and at full width → commit with a plain-English message → push `main` → confirm the `pages` action is green.

Tell Brendan at the end: what changed, "Live now, close and reopen the pane", and whether `manifest.xml` or the template workbook changed (either one means a re-upload or a new workbook copy).
