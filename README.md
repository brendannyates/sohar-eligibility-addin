# Sohar Eligibility add-in

Excel task-pane add-in that turns a NextGen **Appointments Listing** into Sohar **discovery** and **verification** import files. It works only inside the companion workbook `Sohar_Eligibility_Template.xlsx`, which holds the payer map, the Sohar payer list, config, and the output tables.

Same stack as the Billing Tickets add-in: static files on GitHub Pages, Office.js, no build step, no server.

## The pane (v0.7)

Four steps, upload-first. Nothing needs pasting.

1. **Upload the NextGen report** (drop or choose; .xlsx or .csv). It is scanned and built right away. Unmapped insurance names show inline; mapping one rebuilds immediately. Scope, as-of date, rebuild and sheet links sit under Options.
2. **Save the Sohar import files**: `Sohar Ready Import - Verification YYYY-MM-DD.csv` and `Sohar Ready Import - Discovery YYYY-MM-DD.csv` (local date), saved to Downloads.
3. **Upload what Sohar sends back** (drop both outputs, or one at a time). If discovery found coverage, a callout offers `Sohar Ready Import - Verification (from Discovery) YYYY-MM-DD.csv` for the second run.
4. **Benefits and OOP max**: counts of OOP / deductible met and approaching, and the patients at or near their OOP max (LASFIR + Per Nbr, $ left, % met). Writes the `Benefits` sheet (`tblBenefits`, `BENEFIT_SHEET_COLS`), one row per verified patient, OOP closest to met first, with % columns as numbers and data bars. Benefits only exist after a **verification** output is uploaded; with discovery alone the step says so.

IR Portal and the secondary files are under collapsed sections. Full screen adds a **Benefits & OOP** tab.

**Saving on Excel on the web.** Every save tries a normal download first. If the frame blocks it, the note under the button offers **Save in a new tab** (`save.html`, same site, receives the CSV by `postMessage`, no network) or the Export sheet.

## What it does (detail)

1. **Upload or paste** the report. **Upload report file** takes `.csv` or `.xlsx` (xlsx needs ExcelApi 1.13) and replaces the `Import` sheet; pasting still works. Then it **scans** the `Import` sheet. Finds the NextGen header row under the title block, maps columns by name (synonyms live in `core.js`), and reports which fields are present.
2. **Map payers.** Any NextGen payer name not in `tblPayerMap` is listed with Sohar suggestions. Picking one writes the row to the Payer Map sheet, so it is remembered next time.
3. **Build.** One row per `Per Nbr`. Routing: verification when the primary payer maps (Exact, Probable, Confirmed) and `Pol Nbr (Cht Guar)` exists; discovery otherwise (no payer, Self Pay, unmapped, Review, or no member ID). Writes `Patients`, `Review`, `Verification`, `Discovery`.
4. **Export.** Downloads `sohar_verification_import_*.csv` and `sohar_discovery_import_*.csv` with the exact Sohar template columns. The same rows are on the sheets.
5. **Bring back results.** Upload Sohar's verification output and discovery output (the batch result files, .csv or .xlsx). Each row is matched to the appointment report on the `Import` sheet and written to `Sohar Results` (`tblSoharResults`, created and rebuilt by the add-in), soonest appointment first, with a `Work Queue` (Action needed, Unmatched, Pending, Clear) and a plain-English `Action`.
   - Matching: `id` (Per Nbr) first. Older verification exports without `id` fall back to member ID (ignoring case, punctuation and leading zeros), then name. Discovery falls back to name + DOB.
   - Each row carries the next visit (date, time, event, location, provider), the count of upcoming visits and the first six of them.
   - **Benefits** come from Sohar's in-network tier-one columns (`innCopay`, `innCoinsurance`, `innTotalDeductible`, `innRemainingDeductible`, `innOutOfPocket`, `innOutOfPocketRemaining`), plus plan, network type, coverage level (IND/FAM) and coverage dates. Tier 2/3 and out-of-network are ignored. Deductible and OOP status: Met (nothing remaining), Approaching (remaining at or under 20% of the total; change with a Config row `Benefit approaching %`), Not met, None ($0 deductible).
   - An eligible patient whose coverage ends before the next visit is flagged for action.
   - **Choose in NextGen** (discovery): the Payer Map reversed. Sohar's payer ID is looked up against Exact, Probable and Confirmed rows. If one NextGen insurance maps, it is named; if several do, all are listed and the admin picks by plan type; if none does, add it to the Payer Map. Discovery output has no member ID, so the action says to get it from the Sohar link.
   - **Copy action message** and **Copy benefits message** build notes grouped by location (LASFIR + Per Nbr only). Send inside Outlook or Teams only.
   - **Uploads add up.** Every Sohar result is saved on a hidden `Sohar Uploads` sheet (`tblSoharUploads`, one JSON record per row). A new upload adds to it; the same result uploaded again replaces the old copy (key: Per Nbr + payer for verification, Per Nbr for discovery). **Clear saved Sohar results** (click twice) starts fresh.
   - **Discovery → verification.** Discovery only finds the policy (payer, member ID); copay, deductible and OOP come only from a verification submission in Sohar. Every discovery result with coverage found and a matched patient goes on the `Discovery Verify` sheet (`tblDiscoveryVerify`, exactly `VERIFICATION_COLS`) with Sohar's discovered payerId, the patient's name, DOB, state and POS, and `id` = Per Nbr. Member ID comes from the discovery output if present, else from NextGen when NextGen already has that payer, else blank: type it on the sheet (it is kept across rebuilds by Per Nbr + payerId). **Download discovery verification import** reads the sheet. When that verification output is uploaded, the discovery row shows `Verified: <status>` in `Discovery Follow-up`, the verification row shows `From discovery` and carries the benefits, and the patient drops off the import. A verification result closes the discovery when it is for the same patient and the same policy: same member ID when both files have one, else the same Sohar payer.
   - **Download final tracker file (TMS + ESK)** (`sohar_tracker_import_*.csv`, `TRACKER_EXPORT_COLS`) is the one file both trackers import: the benefits columns, then `Service` (TMS / ESK from upcoming NextGen events), `Source` (Verification, or Discovery, then verification) and `Exported`. One row per patient from their latest Sohar verification. Patients whose discovered coverage still waits on a verification are left out and listed (LASFIR + Per Nbr), so download it again after that verification comes back. TMS Tracker: Import list > Import Sohar benefits. ESK Tracker: Import Sohar benefits, which fills the Benefits tab.
   - **Download benefits CSV** is one row per verified patient in `BENEFITS_EXPORT_COLS` order (`core.js`). This is the contract for the TMS Tracker import: add columns at the end only.

6. **Cross-check the IR Portal.** Upload the IR Portal workbook (the `Filtered` tab is read; headers can start in column C) or a CSV of it. Only insurance-related descriptions are kept: missing insurance, missing demographics or verify insurance, out of network, authorization. Each is matched to the Sohar results by MRN (ignoring punctuation and leading zeros), then last + first name + DOB, and written to `IR Cross-Ref` (`tblIrCrossRef`), leaving out items with no Sohar result, with the Sohar result, payer, plan, coverage dates, open NextGen tasks and a suggested action, soonest due date first. Coverage that started after the DOS is called out. **Copy IR message** groups items by facility with LASFIR and the last four MRN digits. Sohar does not check authorizations or provider participation, and the suggested actions say so.

**Load on open.** Nothing needs a manual scan or build. When the pane opens (or a report is uploaded, or Refresh from Import sheet is clicked after pasting) and the Import sheet holds a report, it scans it, builds Patients, Review, Verification and Discovery for the current as-of date, re-matches every saved Sohar upload (hidden `Sohar Uploads`) and re-runs the IR cross-check from the saved insurance-related IR items (hidden `IR Uploads`, `tblIrUploads`). Each refresh or upload writes Config `Last updated` and shows it in the header with the newest Sohar upload time.

**NextGen Tasks** (`tblNextGenTasks`, created by the add-in) is rebuilt with every results upload: one row per thing to change before the visit (add or update insurance, fix member ID or name, update benefits with the values to enter, add a payer to the Payer Map, find an unmatched result, resend to Sohar). `Done` marks carry over between uploads by `Task Key`.

**Open full screen** floats over the Excel page (an Office dialog with `displayInIframe: true`); it never opens a separate window. Menus and detail cards that pop out over the add-in blur what's behind them. Tabs: Prep by appointment (each visit day, soonest first, with result type, status, benefits and open tasks), Worklist (queue and benefit tiles and filters), NextGen tasks (tick Done; it writes to the sheet), Messages (open tasks, who needs action, deductible/OOP met or approaching), and Upload & save (drop the Sohar verification and discovery outputs; save every CSV). Needs DialogApi 1.2. Data moves between pane and window with `messageChild` / `messageParent` in pieces; nothing leaves Excel.

Excel on the web blocks downloads from the task pane, so on the web every Download button opens the full-screen window to save from there. The pane footer shows the platform and API support; include it when reporting a problem.

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
| `dashboard.html`, `dashboard.js` | Full-screen window (Office dialog): prep, worklist, NextGen tasks, messages, upload and save. |
| `index.html`, `style.css` | Pane markup and styles. |
| `manifest.xml` | Ribbon button and host URL. |
| `tests/*.test.mjs` | `node --test` suite. |
| `scripts/serve.mjs` | Static server for demo mode. |

## Template contract

Tables the add-in expects: `tblPayerMap` (Payer Map), `tblSoharPayers` (Sohar Payers), `tblConfig` and `tblLocationState` (Config), `tblPatients`, `tblReview`, `tblVerification`, `tblDiscovery`. Output tables are replaced on every build; their column counts must match `PATIENT_COLS`, `REVIEW_COLS`, `VERIFICATION_COLS`, `DISCOVERY_COLS` in `core.js`. If a column is added on either side, add it on both and ship a new template.

Refresh `tblSoharPayers` by pasting a new Sohar payer export over the table (keep the headers).
