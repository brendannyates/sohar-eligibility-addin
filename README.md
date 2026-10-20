# Sohar Eligibility add-in

Excel task-pane add-in that turns a NextGen **Appointments Listing** into Sohar **discovery** and **verification** import files. It works only inside the companion workbook `Sohar_Eligibility_Template.xlsx`, which holds the payer map, the Sohar payer list, config, and the output tables.

Same stack as the Billing Tickets add-in: static files on GitHub Pages, Office.js, no build step, no server.

## The pane (v0.8)

Two tabs at the top: **Clinic** and **Admin**. The pane opens on Admin in the Sohar template and on Clinic in a clinic file (a workbook with `tblClinicMeta`). In a clinic file the Admin tab only says it needs the template.

**Clinic** is the pre-check view: every in-scope patient with an upcoming visit at the chosen location (remembered in Config `Clinic location`), grouped by day. Each card shows the next visit and provider, payer and secondary, which insurance to choose in NextGen, Sohar status and verified date, copay and coinsurance, deductible and OOP progress, flag pills, review items, and a **Re-verify in Sohar** link (`https://app.soharhealth.com/create?type=verification`, a plain link, opens a new tab). Chips filter by pill (Medi-Cal is the Medi-Cal focus; Kaiser/VA MM is the Kaiser, VA and TriWest med-management drill-down). **Copy email summary** uses LASFIR + Per Nbr only.

**Admin** is the four steps below, plus clinic files, payer evidence, IR and more files.

### Rules (v0.8)

- **R1 Report end.** Data stops at the first row after the header starting with `Total Appointments`. Rows with content below it (left over from pasting over a longer report) are ignored and a red warning says how many. A second warning fires when payer values match Rendering or Res Name values (shifted columns).
- **R2 Auto payer mapping.** NextGen payer names not on the Payer Map are matched to Sohar payers and written to the Payer Map once: `Auto-High` (exact name or alias, or the same name after normalization as a Payer Map name labeled Exact, Probable, Confirmed or Learned), `Auto-Medium` (strong fuzzy match: score ≥ 0.75, ≥ 66% of words matched, clear winner by ≥ 0.15), `Auto-Low` (best guess only). Those confirmed Payer Map names also count as fuzzy candidates for their payerId. Auto-High, Auto-Medium and `Learned` (see Payer learning) go to verification. A manual confidence on an existing row always wins. Self Pay and VIP names are never auto-matched. Step 1 has an optional **Review auto-mapped payers** list: Looks right, Change, Not a payer.
- **R3 Two insurances.** Low-confidence chart payer and a different last encounter (Fin Class or Pol Nbr): both are sent to verification with id = Per Nbr. The last-encounter payer comes from this report's Fin Class → most common confidently mapped NextGen payer. The Sohar results decide; the losing row is marked Superseded.
- **R4 VIP / Partial VIP** = MHS-sponsored. Verify any other insurance on file, else discovery only. Always a VIP pill.
- **R5 Kaiser / VA / TriWest** (one bucket; flags name the payer as it appears) med-management visits (not TMS, esketamine, ketamine, MT, meetings, testing) with no Auth Cd on the chart or last encounter: `Kaiser/VA MM no auth`.
- **Dual eligible**: Medicare primary with a Medi-Cal secondary keeps the Medi-Cal flag and gets the QMB treatment (do not collect Medicare cost-sharing; confirm with billing).
- **Auto-Low, Review, or an unrecognized label** (anything that is not Exact, Probable, Confirmed, Auto-High, Auto-Medium, Auto-Low, Learned, Not a payer or Review, blank included): listed at the top of step 1 as needing manual classification, and each patient gets a High `Payer needs manual classification` row on the Review sheet. So do human-labeled payers Sohar disagrees with, and names with a Sohar suggestion.
- **R6 Medi-Cal** other than Kaiser or SFHP (NextGen, secondary, Fin Class, or what Sohar returned): red `Medi-Cal` pill, plus `New Medi-Cal` when the last visit was not Medi-Cal. SFHP patients booked with a provider not on Config `SFHP providers` (names or last names, `;` separated) are flagged; an empty list asks to confirm the provider.
- **R7** UMR and Carelon (incl. Beacon) TMS visits with an NP rendering provider: `NP-TMS`.
- **R8 Secondary** on the chart: "may cover the deductible, copay or coinsurance" (stronger, medium, when Medicare/Noridian/QMB is involved). QMB anywhere: `QMB` pill, do not collect Medicare cost-sharing until billing confirms.
- **R9 Re-verify**, for patients with an upcoming visit, uses the same due rule as the re-verify skip window below (`core.reverifyCheck`), per insurance sent (verifyRows; no verifyRows: any verification): missing, latest not `complete.eligible`, `Re-verify days` (default 7) or more calendar days old, coverage ending before the visit, or a next visit in a new month within `Re-verify days`. A name-only match never counts.

Flags are added to the **Review** sheet rows (same `REVIEW_COLS`) and a `Pills` column is added after the standard columns on **Sohar Results** and **Benefits**. Full screen shows pills and a Flag filter.

### Payer learning (v0.10)

The add-in records every payer classification and every Sohar outcome, and uses them to improve payer confidence. Rules approved by Brendan on 2026-10-20.

- **Payer Outcomes** (hidden sheet, `tblPayerOutcomes`, `PAYER_OUTCOME_COLS`): one row per Sohar outcome, written whenever results are merged (upload, refresh on open). Name matches (`Name only`, `Name + DOB`), `No match` and `No identifiers` rows are not recorded: someone else's result must not teach. Append-only: the key is the Sohar result key + Sohar Date, so re-uploading the same file adds nothing but a later re-check is a new outcome. **Clear saved Sohar results** does not clear it. It holds Per Nbr, so it stays hidden. Signals: verification `complete.eligible` = agree; `error.payer*` (or another non-member error saying the payer was wrong, unknown or not found) = payer error, except outages (unavailable, timeout, not responding, down, retry), which are neutral; `error.member.id`, `complete.ineligible` and anything else are neutral. The real Sohar `error.payer*` strings still need checking against a Sohar output. Discovery compares the payer found with the patient's Sohar payerId, else the Payer Map's guess for their NextGen payer (recorded as Sent payerId): agree, other payer, `found` (nothing to compare with; feeds suggestions only), neutral (nothing found). Last-encounter verifications are recorded under the proxy NextGen payer and the Lst Enc Fin Class. A verification of a payer discovery found is neutral (it says nothing about the NextGen name).
- **Rollup** (`core.learnPayers`), per NextGen payer name: distinct patients from each patient's latest outcome (per name + payerId for verification, per name for discovery), against the Payer Map's current payerId. Last-encounter verifications do not count for the proxy name (they test the Fin Class guess and feed Learned Fin Class only). A name on the Payer Map twice is never learned or suggested; it gets a Medium `Payer Map name listed twice` Review row. Evidence score = round(100 × (agree + 1) / (agree + conflicts + 2)), conflicts = payer-error + other-payer patients.
  - **Promote**: Auto-Low or Auto-Medium with a payerId, agree ≥ `Learn: promote patients` (3), no conflicts → `Learned`.
  - **Demote**: Auto-High, Auto-Medium or Learned with payer errors ≥ `Learn: demote payer errors` (2), or discovery patients ≥ `Learn: demote min patients` (3) with ≥ `Learn: demote discovery share` (50%) finding another payer → `Auto-Low`. Promotion needs zero conflicts, so a demoted row does not flip back by itself.
  - **Flag**: a human label (Exact, Probable, Confirmed, Not a payer, Review, anything else) meeting the demote rule is never changed; it gets a High `Sohar disagrees with Payer Map` Review row and shows in step 1 (also when no one in this report has that payer). **Looks right** keeps a human label as it is and notes `Sohar disagreement reviewed <when>, conflicts=N` (logged on the Payer Map Log); the flag stays quiet until the conflicts grow past N.
  - **Suggest**: an unmapped, Auto-Low or Review name whose discovery patients returned the same other payer for ≥ `Learn: suggest patients` (3) and ≥ `Learn: suggest share` (80%): step 1 shows **Use <payer>**, which writes Confirmed. Never applied automatically.
- **Re-score**: each build re-runs the matcher for Auto-Low rows only and upgrades them when the result is now Auto-Medium or Auto-High. Never downgrades, never touches other labels, never re-scores a row learning demoted (notes start `Demoted by Sohar learning`) or one with Sohar conflicts, and never into a payerId the ledger contradicts (any payer error or other-payer patient against it, or a suggestion for a different payer).
- **Learned Fin Class**: last-encounter verifications teach Fin Class → payerId when one payerId has agree ≥ promote patients, no payer errors, and no other payerId has any agreeing patient. Those rows say `Fin Class (learned)`; a learned Fin Class beats this report's guess.
- **Config** (optional rows, added on build when missing, with a plain description): `Learning` (On/Off; Off stops promote, demote, re-score and learned Fin Classes, but outcomes are still recorded), `Learn: promote patients`, `Learn: demote payer errors`, `Learn: demote discovery share` (50%, 0.5 or 50), `Learn: demote min patients`, `Learn: suggest patients`, `Learn: suggest share`.
- **Payer Map Log** (sheet, `tblPayerMapLog`): every Payer Map change, append-only: When, NextGen Payer, old and new payerId and confidence, By (Auto-match, Pane, Learning, Re-score, Suggestion, Sheet edit, Baseline), Score, Evidence. A build compares each Payer Map row with its last logged state and logs `Sheet edit` for changes made by hand; the first run logs one `Baseline` row per Payer Map row. No patient identifiers.
- **Payer Learning** (sheet, `tblPayerLearning`): the rollup, rewritten every build, one row per NextGen payer in the ledger or the Payer Map. No patient identifiers.
- **Payer Map writes** keep earlier Notes (new note first, older after ` | `, about 500 characters). `Last reviewed` is stamped only when a person classifies a payer in the pane; automatic writes keep it (new automatic rows leave it blank). It is always written as text `MM/DD/YYYY`.
- **error.payer.npi** (the provider's NPI setup at the payer) is neutral, never a payer error.
- Step 1 rows show the evidence (`Sohar: 9 agreed, 1 other payer, score 83`); the build note says `Learning: N promoted, N demoted, N suggestions, N re-scored`. After an upload adds outcomes, the add-in rebuilds so the Payer Map, Patients and imports follow; the Sohar Results sheet picks up a changed mapping on the next refresh.

### Re-verify skip window (v0.10)

Approved by Brendan on 2026-10-20. Config (optional rows, added on build when missing): `Skip recent verifications` (On), `Re-verify days` (7, shared with R9), `Re-verify in new month` (Off).

- A verification-routed insurance (one verifyRows entry: Sohar payerId + member ID) is **left out of the verification file** when Skip is On, the patient has no pending **Re-verify anyway** request, and the latest Sohar verification for the same patient, payerId and member ID is `complete.eligible`, dated (Sohar Date, local calendar day) fewer than `Re-verify days` calendar days ago (verified 1/31, today 2/7, 7 days: due). Exceptions that re-verify anyway: the next visit is in a different month or year than the verification and within `Re-verify days` of today (1/31 → visit 2/1: re-verify; visit 2/14: skipped until 2/7); with `Re-verify in new month` On, any month change (today or the next visit); Sohar's coverage end is before the next visit. Ineligible, error and pending results never skip. A different payerId or member ID (or a new last-encounter insurance) is never skipped. Discovery routing is unchanged.
- The history is the saved Sohar uploads matched again (Per Nbr, Sohar ref or Member ID; never a name match). If the uploads, Sohar Refs or requests cannot be read, nobody is skipped that build.
- Sohar verification exports have no payerId, so each new verification record on Sohar Uploads gets the payerId sent recorded when it is uploaded (`_sentpayerid` inside Record; matched by Per Nbr, Sohar ref or Member ID; never changed later). Sohar Results shows that payerId. A record saved before v0.10 has none: it counts toward a skip only when the Sohar payer name on the result and on the Payer Map are both present and equal.
- Skipped patients stay on **Patients** (Route stays verification; Route Reason ends `; skipped: verified eligible 10/18/2026, next check 10/25/2026`; `x.reverifySkip` = {verified, due, reason, partial}). With two insurances only the recently verified one is left out. Step 2 lists them (**Skipped (verified recently)**: LASFIR, Per Nbr, verified, next check, **Re-verify anyway**); the build note says `N skipped: verified within N days`.
- **Re-verify anyway** saves the request on the hidden `Re-verify Requests` sheet (`tblReverifyRequests`: Per Nbr, Requested) and rebuilds. A request is dropped on the next build once that patient has a Sohar verification dated after it.

### Sohar ref memory and results with no identifiers (v0.10)

- Sohar exports can arrive with `id`, `memberId`, names and DOB blank. Each Sohar check's refs (`requestId`, the id in its `https://app.soharhealth.com/verifications/<id>` link, `verificationId` on old exports) are remembered on the hidden `Sohar Refs` sheet (`tblSoharRefs`: Sohar Ref, Per Nbr, Kind, Result Key, First Matched), append-only, once the record matched by **Per Nbr or Member ID** (never by name). **Clear saved Sohar results** keeps it.
- A record with a blank `id` and a remembered ref matches that Per Nbr (`Match` = `Sohar ref`) and keeps the result key of the original record, so a re-export never duplicates a row on Sohar Uploads or the Payer Outcomes ledger. When it replaces the saved copy of the same Sohar check, the identifiers it lacks are kept from the earlier copy. A ref learned in the same pass (the original and its blank re-export in one upload, or an original saved before the Sohar Refs sheet existed) is used straight away: the upload is matched once more before anything is saved (`core.combineResults`).
- A result with no `id`, no remembered ref, and no member ID or name (verification) or no name or DOB (discovery) gets `Match` = `No identifiers` and an Action saying it is likely an export anomaly or a check run outside the batch (open the Sohar link). It gets no `Find patient for Sohar result` task; instead one High Review row `Sohar results with no patient identifiers` with the count, and a note in step 3. A result with identifiers that are not found stays an ordinary `No match`.

### Fee estimates (test, off by default) (v0.11)

A test feature: it estimates what a verified patient pays at the **next visit**, MD and NP, from an uploaded fee schedule. Estimates are not a guarantee.

**Turn on:** Config `Fee estimates` = `On`, then refresh (⟳). The add-in adds the row (`Off`) on the next build if it is missing. **Turn off:** set it back to `Off` and refresh. Off means v0.10 behaviour: no fee UI, no `Est MD` / `Est NP` columns on the Benefits sheet, no fee sheets read, nothing sent to the full-screen window. No template table or column changes either way, and no network calls.

**Upload:** Admin step 4 > *Fee estimates (test)*, or the Benefits & OOP tab in full screen. The `.xlsx` is read in the pane by a small pure-JS reader (`xlsx.js` `readWorkbook`, DEFLATE through the browser's `DecompressionStream`), never through Excel sheets. Re-uploading replaces the saved schedule completely; nothing is bundled in the add-in. The pane shows `Fee schedule: <file>, uploaded <date>, N tabs, N codes` and warns when it is more than 90 days old.

**Parsing** (`fees.js` `parseFeeSchedule`): visible tabs only (hidden tabs such as Summary or the payer tabs are skipped). Per tab: a `Locations:` / `Location:` line (split on commas, periods and "and"), the title's `Locality N`, the row with `CPT Code` (any column, any row), then Gross Charge, Cash and payer names over MD and Nurse / NP columns (a name carries over its merged cells or onto the NP column right after its MD column; "Imperial" over MD next to "Med Adv" over NP is one payer). `99417/G2212` (or `99417G2212`) gives both codes. `0`, `0.00`, blank, `NA`, `#REF!`, percentages and text are not rates (kept as a note). Code rows stop at the first row whose code cell is a note. Tabs with no MD / NP payer columns are skipped.

**Matching (never a guess):**
- Next visit location -> fee tab: the tab's Locations list and tab name, as whole words, case and punctuation ignored. A tie between tabs is no match. Telehealth (Config `Telehealth keywords`) uses the rendering provider's home clinic.
- Payer (Sohar payer from an eligible verification, else the NextGen payer) -> payer column on that tab: every word of the column name (or one side of a `/` name) must be in the payer name, the most specific wins, a Medicare / Advantage payer only matches a Medicare / Advantage column. Low confidence or a tie: no estimate, flagged *Map payer to a fee column*.
- Out of network (Sohar `networkType` OON / out of network): allowed = Gross Charge, with Sohar's out-of-network benefits.

**Sheets** (all add-in owned, saved fail-safe, safe to delete; created only when On):

| Sheet | Table | What |
|---|---|---|
| Fee Schedule (hidden) | `tblFeeSchedule` | Tab, Locality, Locations, Code, Description, Payer, Level (MD / NP / Gross / Cash), Rate, Note |
| Fee Schedule Info (hidden) | `tblFeeMeta` | File, upload time, counts, tab summary |
| Fee Locations | `tblFeeLocations` | NextGen Location -> Fee Tab |
| Provider Home Clinic | `tblProviderHome` | Provider -> Home Location (most frequent in-person location in the report) |
| Fee Payer Map | `tblFeePayerMap` | Payer + Fee Tab -> Fee Payer, Confidence. A row with a blank Fee Tab applies to every tab. |
| Fee Payer Map Log | `tblFeePayerMapLog` | Every Fee Payer Map change (Auto-match or Sheet edit) |
| Estimate Codes | `tblEstimateCodes` | Event keyword -> codes with units (`90868 x1; 99417 x2`). Seeded once, marked `default, review`; the longest keyword found in the event wins. Esketamine seeds use G2083: confirm 56 mg (G2082) vs 84 mg. |
| Estimate Defaults | `tblEstimateDefaults` | Deductible / Coinsurance / Copay toggles remembered per fee payer |

On the mapping sheets, change the value column and the add-in keeps it (the row becomes `manual`); delete a row to hand it back to automatic matching. Edits are read on refresh (⟳) or *Re-read fee sheets*.

**Math, per visit** (several codes: allowed amounts summed first): deductible (`Applies`: min(allowed, remaining) or `Waived`); coinsurance (`After deductible`: rate x (allowed - deductible part); `Before deductible`: rate x allowed, no deductible step, and it lowers the deductible left; or `None`); copay (`Added` by default only when Sohar returned one above $0, or `None`); then capped at the allowed amount and the OOP max remaining, never below $0, rounded to cents. Coinsurance accepts 0.2, 20 or 20%. Defaults come from Config (`Estimate deductible`, `Estimate coinsurance`, `Estimate copay`) until staff change them for a fee payer. A missing benefit is named and the estimate uses what is known; with no deductible, coinsurance or copay at all there is no patient amount.

**Where it shows (On only):** the Clinic tab card (MD / NP totals), the Benefits sheet (`Est MD`, `Est NP`), and full screen (Benefits tab columns; click a patient for the estimate panel with code and units editors, the three toggles and the line-by-line math). The window gets only the contexts and the rate tables it needs, through the chunked pane messages; it never reads Excel.

**Remove completely:** set `Fee estimates` = `Off`; delete the sheets above (unhide `Fee Schedule` and `Fee Schedule Info` first: right-click a sheet tab > Unhide); optionally delete the four `Fee estimates` / `Estimate ...` Config rows. To remove the code, revert to tag `v0.10.0`.

**Demo:** `npm run serve`, then `http://localhost:8080/?fees=1` (made-up fee schedule; `dashboard.html?fees=1#benefits` for the window on its own).

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
3. **Build.** One row per `Per Nbr`. Routing: verification when the primary payer maps (Exact, Probable, Confirmed, Auto-High, Auto-Medium, Learned) and `Pol Nbr (Cht Guar)` exists; discovery otherwise (no payer, Self Pay, unmapped, Auto-Low, Review or an unrecognized label, or no member ID). R3 and R4 above can add a last-encounter verification. Writes `Patients`, `Review`, `Verification`, `Discovery`.
4. **Export.** Downloads `sohar_verification_import_*.csv` and `sohar_discovery_import_*.csv` with the exact Sohar template columns. The same rows are on the sheets.
5. **Bring back results.** Upload Sohar's verification output and discovery output (the batch result files, .csv or .xlsx). Each row is matched to the appointment report on the `Import` sheet and written to `Sohar Results` (`tblSoharResults`, created and rebuilt by the add-in), soonest appointment first, with a `Work Queue` (Action needed, Unmatched, Pending, Clear) and a plain-English `Action`.
   - Matching: `id` (Per Nbr) first, then a remembered Sohar ref (see Sohar ref memory). Older verification exports without `id` fall back to member ID (ignoring case, punctuation and leading zeros), then name. Discovery falls back to name + DOB.
   - Each row carries the next visit (date, time, event, location, provider), the count of upcoming visits and the first six of them.
   - **Benefits** come from Sohar's in-network tier-one columns (`innCopay`, `innCoinsurance`, `innTotalDeductible`, `innRemainingDeductible`, `innOutOfPocket`, `innOutOfPocketRemaining`), plus plan, network type, coverage level (IND/FAM) and coverage dates. Tier 2/3 and out-of-network are ignored. Deductible and OOP status: Met (nothing remaining), Approaching (remaining at or under 20% of the total; change with a Config row `Benefit approaching %`), Not met, None ($0 deductible).
   - An eligible patient whose coverage ends before the next visit is flagged for action.
   - **Choose in NextGen** (discovery): the Payer Map reversed. Sohar's payer ID is looked up against verifiable rows (Exact, Probable, Confirmed, Auto-High, Auto-Medium, Learned). If one NextGen insurance maps, it is named; if several do, all are listed and the admin picks by plan type. If only guesses name it (Auto-Low, Review, or a label the add-in does not know), those are named and the task is **Confirm payer mapping**; if no row does, **Add payer to Payer Map**. **Coverage Differs** compares with the patient's Sohar payerId, or the Payer Map's guess when none was sent (Auto-Low). Discovery output has no member ID, so the action says to get it from the Sohar link.
   - **Copy action message** and **Copy benefits message** build notes grouped by location (LASFIR + Per Nbr only). Send inside Outlook or Teams only.
   - **Uploads add up.** Every Sohar result is saved on a hidden `Sohar Uploads` sheet (`tblSoharUploads`, one JSON record per row). A new upload adds to it; the same result uploaded again replaces the old copy (key: Per Nbr + payer for verification, Per Nbr for discovery). **Clear saved Sohar results** (click twice) starts fresh.
   - **Discovery → verification.** Discovery only finds the policy (payer, member ID); copay, deductible and OOP come only from a verification submission in Sohar. Every discovery result with coverage found and a matched patient goes on the `Discovery Verify` sheet (`tblDiscoveryVerify`, exactly `VERIFICATION_COLS`) with Sohar's discovered payerId, the patient's name, DOB, state and POS, and `id` = Per Nbr. Member ID comes from the discovery output if present, else from NextGen when NextGen already has that payer, else blank: type it on the sheet (it is kept across rebuilds by Per Nbr + payerId). **Download discovery verification import** reads the sheet. When that verification output is uploaded, the discovery row shows `Verified: <status>` in `Discovery Follow-up`, the verification row shows `From discovery` and carries the benefits, and the patient drops off the import. A verification result closes the discovery when it is for the same patient and the same policy: same member ID when both files have one, else the same Sohar payer.
   - **Download final tracker file (TMS + ESK)** (`sohar_tracker_import_*.csv`, `TRACKER_EXPORT_COLS`) is the one file both trackers import: the benefits columns, then `Service` (TMS / ESK from upcoming NextGen events), `Source` (Verification, or Discovery, then verification) and `Exported`. One row per patient from their latest Sohar verification. Patients whose discovered coverage still waits on a verification are left out and listed (LASFIR + Per Nbr), so download it again after that verification comes back. TMS Tracker: Import list > Import Sohar benefits. ESK Tracker: Import Sohar benefits, which fills the Benefits tab.
   - **Download benefits CSV** is one row per verified patient in `BENEFITS_EXPORT_COLS` order (`core.js`). This is the contract for the TMS Tracker import: add columns at the end only.

6. **Cross-check the IR Portal.** Upload the IR Portal workbook (the `Filtered` tab is read; headers can start in column C) or a CSV of it. Only insurance-related descriptions are kept: missing insurance, missing demographics or verify insurance, out of network, authorization. Each is matched to the Sohar results by MRN (ignoring punctuation and leading zeros), then last + first name + DOB, and written to `IR Cross-Ref` (`tblIrCrossRef`), leaving out items with no Sohar result, with the Sohar result, payer, plan, coverage dates, open NextGen tasks and a suggested action, soonest due date first. Coverage that started after the DOS is called out. **Copy IR message** groups items by facility with LASFIR and the last four MRN digits. Sohar does not check authorizations or provider participation, and the suggested actions say so.

**Load on open.** Nothing needs a manual scan or build. When the pane opens (or a report is uploaded, or Refresh from Import sheet is clicked after pasting) and the Import sheet holds a report, it scans it, builds Patients, Review, Verification and Discovery for the current as-of date, re-matches every saved Sohar upload (hidden `Sohar Uploads`) and re-runs the IR cross-check from the saved insurance-related IR items (hidden `IR Uploads`, `tblIrUploads`). Each refresh or upload writes Config `Last updated` and shows it in the header with the newest Sohar upload time.

**Fail-safe saves.** The add-in's own sheets (Sohar Results, NextGen Tasks, Benefits, Discovery Verify, Discovery from Verify, Export, IR Cross-Ref, Payer Evidence, Payer Map Log, Payer Learning, and the hidden Sohar Uploads, IR Uploads, Payer Outcomes, Sohar Refs and Re-verify Requests) are never rewritten in place. A save writes the new rows to a hidden `~saving <table>` sheet, makes them a table named `<table>_stg` only when every row is in, then swaps that sheet in for the old one. If a save fails before the swap, nothing changed and the message says so. If Excel closes mid-swap, the next read, save or pane open finishes it and the header says an interrupted save was finished. A report upload is written the same way (hidden `~saving Import`, then one copy), so Import is never left half-written. Because each save replaces the sheet, column widths or filters set by hand on those sheets reset, and **formulas must not point at them** (they would show #REF!). The template tables (Patients, Review, Verification, Discovery) are still rewritten in place; they are rebuilt from Import on every refresh. A structure-protected workbook cannot save these sheets; the message says to unprotect it.

**NextGen Tasks** (`tblNextGenTasks`, created by the add-in) is rebuilt with every results upload: one row per thing to change before the visit (add or update insurance, fix member ID or name, update benefits with the values to enter, add a payer to the Payer Map, confirm a guessed payer mapping, find an unmatched result, resend to Sohar). `Done` marks carry over between uploads by `Task Key`.

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
| `rules.js` | Pure clinic rules (R5-R10): flags and pills, Review rows, clinic sheet and email, payer evidence. Loads after `core.js` (R9 uses `core.reverifyCheck`). |
| `xlsx.js` | Dependency-free .xlsx writer for clinic files (real tables, hyperlinks, hidden sheet, pane auto-open parts), and a small reader (`readWorkbook`) for the fee schedule upload. |
| `fees.js` | Pure fee estimate logic (v0.11, a test, Off by default): fee schedule parsing, location / payer matching, estimate math, made-up demo schedule. Loads after `core.js`. |
| `save.html`, `save.js` | Fallback save tab for Excel on the web; takes CSV text or base64 (.xlsx). |
| `excel.js` | Everything that reads or writes the workbook. Table names must match the template. |
| `app.js` | Pane behaviour, step flow, payer mapping UI, uploads, downloads, demo mode. |
| `dashboard.html`, `dashboard.js` | Full-screen window (Office dialog): prep, worklist, NextGen tasks, messages, upload and save. |
| `index.html`, `style.css` | Pane markup and styles. |
| `manifest.xml` | Ribbon button and host URL. |
| `tests/*.test.mjs` | `node --test` suite. `tests/excel.test.mjs` runs the real `excel.js` against `tests/helpers/excel-mock.mjs` (a model of the Excel API with failure injection) and sweeps every save through a failure at every step. |
| `scripts/serve.mjs` | Static server for demo mode. |

## Template contract

Tables the add-in expects: `tblPayerMap` (Payer Map), `tblSoharPayers` (Sohar Payers), `tblConfig` and `tblLocationState` (Config), `tblPatients`, `tblReview`, `tblVerification`, `tblDiscovery`. Output tables are replaced on every build; their column counts must match `PATIENT_COLS`, `REVIEW_COLS`, `VERIFICATION_COLS`, `DISCOVERY_COLS` in `core.js`. If a column is added on either side, add it on both and ship a new template.

Refresh `tblSoharPayers` by pasting a new Sohar payer export over the table (keep the headers).
