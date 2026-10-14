# Sohar Eligibility add-in

Excel task-pane add-in that turns a NextGen **Appointments Listing** into Sohar **discovery** and **verification** import files. It works only inside the companion workbook `Sohar_Eligibility_Template.xlsx`, which holds the payer map, the Sohar payer list, config, and the output tables.

Same stack as the Billing Tickets add-in: static files on GitHub Pages, Office.js, no build step, no server.

## What it does

1. **Upload or paste** the report. **Upload report file** takes `.csv` or `.xlsx` (xlsx needs ExcelApi 1.13) and replaces the `Import` sheet; pasting still works. Then it **scans** the `Import` sheet. Finds the NextGen header row under the title block, maps columns by name (synonyms live in `core.js`), and reports which fields are present.
2. **Map payers.** Any NextGen payer name not in `tblPayerMap` is listed with Sohar suggestions. Picking one writes the row to the Payer Map sheet, so it is remembered next time.
3. **Build.** One row per `Per Nbr`. Routing: verification when the primary payer maps (Exact, Probable, Confirmed) and `Pol Nbr (Cht Guar)` exists; discovery otherwise (no payer, Self Pay, unmapped, Review, or no member ID). Writes `Patients`, `Review`, `Verification`, `Discovery`.
4. **Export.** Downloads `sohar_verification_import_*.csv` and `sohar_discovery_import_*.csv` with the exact Sohar template columns. The same rows are on the sheets.
5. **Bring back results.** Upload the Sohar verification and discovery result files. Each row is matched to the appointment report on the `Import` sheet and written to `Sohar Results` (`tblSoharResults`, created by the add-in), soonest appointment first, with a `Work Queue` (Action needed, Unmatched, Pending, Clear) and a plain-English `Action`.
   - Verification results have no `id` column, so they match on member ID (ignoring case, punctuation and leading zeros), then on first + last name (marked "Name only. Confirm.").
   - Discovery results match on `id` (Per Nbr), then name + DOB.
   - Upcoming patients routed to an uploaded file type who are missing from it show as "Not in Sohar file".
   - Deductible and out-of-pocket columns are picked up if a Sohar file has them (in-network individual only; family and out-of-network columns are skipped). Sohar's standard verification and discovery exports don't include them today, so those columns stay blank until Sohar provides an export that does.
   - **Copy admin message** builds a note grouped by location listing only patients who need action, identified by LASFIR and Per Nbr. Send it inside Outlook or Teams only.

**Open full screen** opens the results worklist in an Office dialog (floating over the sheet in Excel on the web, a separate window on desktop) with filters for work queue, deductible/OOP met, location, appointment dates and search. It needs DialogApi 1.2. Data goes to the dialog over `messageChild`; nothing leaves Excel.

Scope `upcoming` keeps patients with a non-cancelled, non-no-show appointment on or after the as-of date. POS is `10` when the next appointment's Event contains a telehealth keyword (`Tele`), else `11`. State comes from a patient state column if the report has one, else Config > Location to State, else Default state.

## Deploy (first time)

1. Create a GitHub repo named `sohar-eligibility-addin` under the same owner as the billing add-in and push this folder to `main`. Settings → Pages → Source: **GitHub Actions**. The `pages` workflow runs the tests and publishes.
2. In `manifest.xml`, replace every `YOUR-GITHUB-USERNAME` with the GitHub owner (10 occurrences; search and replace). Nothing else in the manifest needs editing.
3. Distribute the manifest the same way as Billing Tickets: Microsoft 365 admin center → Integrated apps → Upload custom apps → `manifest.xml` → assign the security group. For a quick test first: Excel → Home → Add-ins → More add-ins → My Add-ins → Upload My Add-in.
4. Open `Sohar_Eligibility_Template.xlsx` and click **Sohar Eligibility** on the Home tab.

Code changes ship by pushing to `main`. Manifest changes need a re-upload in Integrated apps.

## Local preview

```
npm test          # unit tests on synthetic data
npm run serve     # http://localhost:8080 opens in demo mode (synthetic rows, nothing written)
```

## PHI

- The workbook holds names, DOBs, and member IDs once the report is pasted. Keep it in MHS OneDrive or SharePoint.
- The pane makes no network calls (`connect-src 'none'` in `index.html`). The only remote asset is Office.js from Microsoft's CDN. CSVs are generated in the browser and saved locally.
- GitHub Pages hosts code only and never receives workbook data.
- Tests and demo mode use synthetic data only. Never commit a real export.

If the pane ever fails to load inside Excel after a Microsoft update, the first thing to try is relaxing `connect-src` in the `<meta http-equiv="Content-Security-Policy">` tag to `https://*.microsoft.com https://*.office.com`.

## Files

| File | Role |
|---|---|
| `core.js` | Pure logic: header detection, name/date parsing, routing, exports, CSV. Shared with tests. |
| `excel.js` | Everything that reads or writes the workbook. Table names must match the template. |
| `app.js` | Pane behaviour, step flow, payer mapping UI, uploads, downloads, demo mode. |
| `dashboard.html`, `dashboard.js` | Full-screen results worklist shown in the Office dialog. |
| `index.html`, `style.css` | Pane markup and styles. |
| `manifest.xml` | Ribbon button and host URL. |
| `tests/*.test.mjs` | `node --test` suite. |
| `scripts/serve.mjs` | Static server for demo mode. |

## Template contract

Tables the add-in expects: `tblPayerMap` (Payer Map), `tblSoharPayers` (Sohar Payers), `tblConfig` and `tblLocationState` (Config), `tblPatients`, `tblReview`, `tblVerification`, `tblDiscovery`. Output tables are replaced on every build; their column counts must match `PATIENT_COLS`, `REVIEW_COLS`, `VERIFICATION_COLS`, `DISCOVERY_COLS` in `core.js`. If a column is added on either side, add it on both and ship a new template.

Refresh `tblSoharPayers` by pasting a new Sohar payer export over the table (keep the headers).
