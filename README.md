# Sohar Eligibility add-in

Excel task-pane add-in that turns a NextGen **Appointments Listing** into Sohar **discovery** and **verification** import files. It works only inside the companion workbook `Sohar_Eligibility_Template.xlsx`, which holds the payer map, the Sohar payer list, config, and the output tables.

Same stack as the Billing Tickets add-in: static files on GitHub Pages, Office.js, no build step, no server.

## The pane (v0.8)

Two tabs at the top: **Clinic** and **Admin**. The pane opens on Admin in the Sohar template and on Clinic in a clinic file (a workbook with `tblClinicMeta`). In a clinic file the Admin tab only says it needs the template.

**Clinic** is the pre-check view: every in-scope patient with an upcoming visit at the chosen location (remembered in Config `Clinic location`), grouped by day. Each card shows the next visit and provider, payer and secondary, which insurance to choose in NextGen, Sohar status and verified date, copay and coinsurance, deductible and OOP progress, flag pills, review items, and a **Re-verify in Sohar** link (`https://app.soharhealth.com/create?type=verification`, a plain link, opens a new tab). Chips filter by pill (Medi-Cal is the Medi-Cal focus; Kaiser/VA MM is the Kaiser, VA and TriWest med-management drill-down). **Copy email summary** uses LASFIR + Per Nbr only.

**Admin** is the four steps below, plus clinic files, payer evidence, IR and more files.

### Rules (v0.8)

- **R1 Report end.** Data stops at the first row after the header starting with `Total Appointments`. Rows with content below it (left over from pasting over a longer report) are ignored and a red warning says how many. A second warning fires when payer values match Rendering or Res Name values (shifted columns).
- **R2 Auto payer mapping.** NextGen payer names not on the Payer Map are matched to Sohar payers and written to the Payer Map once: `Auto-High` (exact name or alias), `Auto-Medium` (strong fuzzy match: score ≥ 0.75, ≥ 66% of words matched, clear winner by ≥ 0.15), `Auto-Low` (best guess only). Auto-High and Auto-Medium go to verification. A manual confidence on an existing row always wins. Self Pay and VIP names are never auto-matched. Step 1 has an optional **Review auto-mapped payers** list (Auto-Low first): Looks right, Change, Not a payer.
- **R3 Two insurances.** Low-confidence chart payer and a different last encounter (Fin Class or Pol Nbr): both are sent to verification with id = Per Nbr. The last-encounter payer comes from this report's Fin Class → most common confidently mapped NextGen payer. The Sohar results decide; the losing row is marked Superseded.
- **R4 VIP / Partial VIP** = MHS-sponsored. Verify any other insurance on file, else discovery only. Always a VIP pill.
- **R5 Kaiser / VA / TriWest** (one bucket; flags name the payer as it appears) med-management visits (not TMS, esketamine, ketamine, MT, meetings, testing) with no Auth Cd on the chart or last encounter: `Kaiser/VA MM no auth`.
- **Dual eligible**: Medicare primary with a Medi-Cal secondary keeps the Medi-Cal flag and gets the QMB treatment (do not collect Medicare cost-sharing; confirm with billing).
- **Auto-Low** payer names are listed at the top of step 1 as needing manual classification and add a High row to the Review sheet.
- **R6 Medi-Cal** other than Kaiser or SFHP (NextGen, secondary, Fin Class, or what Sohar returned): red `Medi-Cal` pill, plus `New Medi-Cal` when the last visit was not Medi-Cal. SFHP patients booked with a provider not on Config `SFHP providers` (names or last names, `;` separated) are flagged; an empty list asks to confirm the provider.
- **R7** UMR and Carelon (incl. Beacon) TMS visits with an NP rendering provider: `NP-TMS`.
- **R8 Secondary** on the chart: "may cover the deductible, copay or coinsurance" (stronger, medium, when Medicare/Noridian/QMB is involved). QMB anywhere: `QMB` pill, do not collect Medicare cost-sharing until billing confirms.
- **R9 Re-verify** when the latest Sohar verification is older than Config `Re-verify days` (default 7) or missing, for patients with an upcoming visit.

Flags are added to the **Review** sheet rows (same `REVIEW_COLS`) and a `Pills` column is added after the standard columns on **Sohar Results** and **Benefits**. Full screen shows pills and a Flag filter.

### Clinic files and payer evidence

- **Clinic files and emails** (Admin): one `Sohar Clinic - <Location> YYYY-MM-DD.xlsx` per location, saved to Downloads like the CSVs (save-tab fallback on the web). Sheets: `Clinic` (`tblClinic`, `rules.CLINIC_COLS`, run date, benefits, pills, review items, re-verify hyperlink), `Review` (`tblClinicReview`), hidden `Clinic Info` (`tblClinicMeta`: Kind = Sohar Clinic, Location, Run Date, Add-in Version). The file also carries the Office web-extension parts so Excel can open the pane with it; if Excel ignores them, open the add-in from the Home tab. **Copy email** per location: LASFIR + Per Nbr, grouped by date, pills and OOP/deductible met or approaching.
- **Payer Evidence** sheet (`tblPayerEvidence`), written every refresh and upserted by Per Nbr + NextGen payer + member ID prefix: Fin Class, prefix, location, ZIP3, state, confidence used, Sohar payer/plan/status, first and last seen. Admin shows a payer × prefix × location summary.
- **Copay history:** each clinic file and the tracker export keep the verification date with deductible and OOP status per patient, so a later copay-log report can join on Per Nbr + date.

### The four Admin steps

1. **Upload the NextGen report** (drop or choose; .xlsx or .csv). It is scanned and built right away; new insurance names are auto-mapped. Scope, as-of date, rebuild and sheet links sit under Options.
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

Clinic files open the pane by themselves only because the ribbon button's `<TaskpaneId>` is `Office.AutoShowTaskpaneWithDocument` (manifest 0.1.1.0 and later). Keep that value; Office needs it plus the web-extension parts the clinic file carries.

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
| `core.js` | Pure logic: header detection, name/date parsing, routing, auto payer matching, exports, CSV. Shared with tests. |
| `rules.js` | Pure clinic rules (R5-R10): flags and pills, Review rows, clinic sheet and email, payer evidence. |
| `xlsx.js` | Dependency-free .xlsx writer for clinic files (real tables, hyperlinks, hidden sheet, pane auto-open parts). |
| `save.html`, `save.js` | Fallback save tab for Excel on the web; takes CSV text or base64 (.xlsx). |
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
