# Power Query in the master workbook (Sohar v0.30)

The master (`Sohar_Eligibility_Template.xlsx`) pulls its sources with these queries. The Sohar add-in only reads what they load: it never writes, swaps or deletes those tables.

The files are generated. Run `node docs/powerquery/make-pq.mjs` after changing a column list in `rules.js`; `tests/powerquery.test.mjs` fails if you forget. Power Query can't run in Node, so set each query up once in Excel and check the row counts.

## Set up (once, in desktop Excel)

1. **pPaths.** Data > Get Data > From Other Sources > Blank Query. Open the Advanced Editor and paste `pPaths.pq`.
   - Replace each placeholder with the file's web address: open the file in OneDrive or SharePoint in the browser, then File > Info > Copy path, and drop the `?web=1` ending.
   - Name the query `pPaths`, then Close & Load To > Only Create Connection.
   - Keep those addresses in the workbook only. They name your OneDrive, so never commit them.
2. **The data queries.** Add them in this order: `qLedger` first, because `qCards` and `qCopays` read it, then `qNGAppts`, `qHistory`, `qCards`, `qCopays`, `qOpenTickets` and `qBalances`. For each one:
   - Blank Query > Advanced Editor > paste the file, and name the query as the file is named.
   - Close & Load To > Table > New worksheet.
   - Rename the sheet to the name in the file's `// Output:` line, and the table (Table Design > Table Name) to its table name. A refresh keeps both names.
   - Hide the sheet. The add-in makes NG Appointments very hidden, because it holds names and DOBs.
3. **Privacy.** File > Options and settings > Query Options > Privacy: set every source to Organizational. Each source is in its own query, which is what the privacy firewall needs.
4. **Sign in.** The first refresh asks you to sign in to each address. Use your work account.

`qBalances` waits on a NextGen report that has an MRN or Per Nbr column and a patient balance column; it doesn't exist yet. Leave it out until that file is in `Data Drop`.

## Refresh

- **By hand:** Data > Refresh All, in Excel on the web or desktop. The Sohar pane and full screen show "Data up to date / aging / stale" from each query's last refresh.
- **Every morning:** `scripts/refresh-master.ps1`, scheduled at 7:30 AM after the rotation (see `scripts/refresh-master.md`).
- **What doesn't work:** a Power Automate flow or an Office Script can't refresh Power Query (Microsoft: `refreshAllDataConnections` "does nothing" for Power Query), and an add-in can't start a refresh either.

| Query | Source | Table | Notes |
|---|---|---|---|
| qNGAppts | Data Drop/NextGen Appointments.xlsx | tblNGAppts | Finds the header row and stops at the "Total appointments" footer. MRN, Per Nbr and Beg Tm load as text. |
| qLedger | Eligibility Checks.csv | tblEligibilityChecks | Row Key = MRN \| Appt Date (YYYY-MM-DD) \| Time. Allowed never leaves the master. |
| qHistory | Eligibility History.csv | tblPatientHistory | One row per Sohar result. |
| qCards | qLedger | tblCardOnFile | The newest card-on-file answer per MRN. |
| qCopays | qLedger | tblCopayActivity | Collected or Luma copay per visit. |
| qOpenTickets | Patient Billing Escalation Forms.xlsx (BillingMaster) | tblOpenTickets | Open = Status not Closed/Resolved, and not deleted. No submitter, notes or attachments. |
| qBalances | Data Drop/NextGen Balances.xlsx | tblPatientBalances | Summed per patient. Names dropped. As Of = the refresh day. |
