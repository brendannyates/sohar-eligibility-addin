# Clinic inbox flow (Power Automate)

How a clinic's changes get from its Clinic Eligibility file back to the master Sohar workbook without anyone sending a file.

- Each clinic works its file (`Clinic Eligibility - <Location> YYYY-MM-DD.xlsx`) in the Clinic Template Tool, in its own folder of the
  **Clinic Eligibility** SharePoint library. Excel on the web saves on its own.
- This flow notices the save, reads the file's **Clinic Work** table (`tblClinicWork`), and appends the rows it has not copied yet to the
  **Clinic Work Inbox** table (`tblClinicWorkInbox`) in the master Sohar workbook.
- In Sohar, the Admin tab (Clinic files and emails) says "N new clinic changes waiting". **Apply clinic changes** writes them to
  Re-verify Requests, NextGen Tasks (Done and Clinic Note), Clinic Notes and Clinic Payer Notes. These are the same writes as
  "Bring back clinic files".

Neither add-in makes a network call; only this flow moves data. The flow never sees a password or key: it uses your own
SharePoint and Excel connections. This guide has no URLs, tenant names or patient data on purpose. Do not add any.

**Open item, before real data:** check with compliance that Power Automate and SharePoint (including flow run history) are
covered by MHS's Microsoft BAA. Until then, test only with a synthetic clinic file.

## What you need first

1. The master Sohar workbook saved in SharePoint or OneDrive for Business. The Excel Online (Business) connector only works with files in
   document libraries that Microsoft Graph supports (OneDrive for Business, SharePoint sites, Microsoft 365 Groups), in .xlsx format, up
   to 25 MB.
2. In the master workbook, open the Sohar pane, then Admin tab > Clinic files and emails > **Set up the clinic inbox**. This adds the
   visible sheet **Clinic Work Inbox** with the table **tblClinicWorkInbox**, header only, every column formatted as text:
   `Location, Kind, Key, MRN, LASFIR, Value, Note, By, When, File, Received`.
3. A SharePoint document library named **Clinic Eligibility** with one folder per clinic, directly under the library (no folders in
   between). Name each folder **exactly** as the location is spelled in Sohar (the flow compares them). Give each clinic team access
   to its own folder only. Edit access to a folder normally lets that team make subfolders in it, and the clinic file's location is
   only a cell on a hidden sheet. So a subfolder named after another clinic, holding a file whose location was changed to match, could
   pass a check on the folder name alone. Step 2d therefore checks the whole folder path: a file must sit directly in its clinic's
   folder. You can also hide the New Folder button (library settings > Advanced settings), but that only hides the button; the path
   check in step 2d is the real guard.
4. Decide which account owns the flow's SharePoint and Excel connections. Best is an account that never edits clinic files itself
   (ask IT for a shared service account). Step 1's third trigger condition skips changes that account makes. If you use your own
   account, your own edits to a clinic file will not reach the inbox until someone at that clinic saves it again.

### Rules for the Clinic Work Inbox sheet

- The flow adds rows; Sohar only reads them. **Do not sort the sheet, type in it, or add formulas that point at it.** Sohar never
  rewrites this table (unlike its other sheets), so it never fights the flow over it.
- Do not hide any of its columns: the connector cannot use a Filter Query on a table with hidden columns.
- Deleting old rows by hand is safe: Sohar ignores rows older than 120 days, and its brought-back ledger (the hidden **Clinic Work
  Applied** sheet) stops a row from applying twice, even if the flow copies it again.
- Rows that arrive twice (a retry, or the same file brought back by hand as well) also apply once.

## Build the flow

In Power Automate: **Create** > **Automated cloud flow**. Name it `Clinic inbox`. Pick the trigger below.

### 1. Trigger: SharePoint "When a file is created or modified (properties only)"

- **Site Address**: the site that holds the Clinic Eligibility library.
- **Library Name**: Clinic Eligibility.
- **Folder**: leave blank. Microsoft's description is "leave blank for the whole library".
- Open the trigger's **Settings**:
  - **Trigger conditions**: add three conditions. A run starts only when all three are true:
    - `@startsWith(triggerBody()?['{FilenameWithExtension}'], 'Clinic Eligibility - ')`
    - `@endsWith(triggerBody()?['{FilenameWithExtension}'], '.xlsx')`
    - `@not(equals(toLower(coalesce(triggerBody()?['Editor']?['Email'], '')), '<the flow connection account, lower case>'))`
  - Why the third one: Microsoft's Excel Online (Business) connector page says a file "may be modified and a new version may be
    visible in Version history of the file even when a 'read-only' action is executed." Steps 2a and 2e read the clinic file, so
    each run could change the file's Modified time and start the next run, over and over, with nothing new to copy. That would use
    up the owner's request limits and can get the flow throttled or turned off. The condition skips any change made by the flow's
    own account.
  - **Concurrency control**: On, **Degree of parallelism** 1, so two saves never write to the inbox at the same moment. Microsoft
    warns this cannot be undone on that trigger (you would have to rebuild the flow to turn it off).

Not sure (test it): Microsoft's page does not say in so many words that this trigger fires for files in subfolders. Its
"properties only" description says "library", and the folder triggers that skip subfolders are the deprecated "in a folder" ones.
Save a test file in one clinic folder and check the run starts. Also check the property names in the trigger's outputs on that
first run (`{FilenameWithExtension}`, `{Path}`, `{Identifier}`, and `Editor` with its `Email`). The expressions here use them.

Not sure (test it): Microsoft's trigger reference says only that its outputs are "dynamic" (the library's columns). `Editor` is the
usual internal name of the Modified By column, but check it on the first run, and check that its Email is filled in for the flow's
account. If Editor is missing, or a version the connector saved shows someone other than the flow's account, the third condition
cannot work. Then build the flow on a schedule instead: **Recurrence** (every 15 minutes, say) plus SharePoint **Get files
(properties only)** on the library, keeping files modified since the last run, then the same steps 2a to 2i for each file.

### 2. Add a **Scope** named `Try`. Put every step below inside it.

#### 2a. Read the file's location: Excel Online (Business) "List rows present in a table"

- **Location**: the same SharePoint site. **Document Library**: Clinic Eligibility.
- **File**: the trigger's **Identifier**. Microsoft documents this as a valid File value.
- **Table**: Enter custom value `tblSEInfo`. The file is picked at run time, so the dropdown cannot list its tables.
- **Filter Query**: `Setting eq 'Location'`
- Rename the step `SE Info`.

Not sure (test it): `tblSEInfo` sits on a hidden sheet. Microsoft's pages do not say whether the connector reads tables on hidden
sheets. The first test run will tell you.

#### 2b. Compose `Location`

`first(body('SE_Info')?['value'])?['Value']`

#### 2c. Compose `Root` and `Folder`

- `Root`: the library's own folder path with a `/` at the end, typed as plain text. Take it from the first test run: open the
  trigger's outputs, copy `{Path}`, and keep everything before the clinic folder's name. For a file in the San Rafael folder whose
  `{Path}` reads `Clinic Eligibility/San Rafael/`, `Root` is `Clinic Eligibility/`. The library's path can differ from its display
  name (for example if it was renamed), so copy it; do not guess.
- `Folder`: the file's whole folder path, without a `/` at the end:
  `if(endsWith(triggerBody()?['{Path}'], '/'), substring(triggerBody()?['{Path}'], 0, sub(length(triggerBody()?['{Path}']), 1)), triggerBody()?['{Path}'])`

#### 2d. Condition: the file sits directly in the folder of its own location

`toLower(outputs('Folder'))` **is equal to** `toLower(concat(outputs('Root'), trim(outputs('Location'))))`

- This compares the whole path, not only the last folder name. A file in a subfolder (`Clinic Eligibility/San Rafael/Walnut Creek/`),
  in another clinic's folder, or at the top of the library fails it.
- **If no**: Office 365 Outlook "Send an email (V2)" to yourself. Subject `Clinic inbox: location and folder disagree`. Body: the file
  name (`triggerBody()?['{FilenameWithExtension}']`) and the folder path only, no row data. Then **Terminate** with status Failed.
  This stops a file put in the wrong clinic's folder, or in a subfolder, from being applied as if that clinic made the changes.
- **If yes**: put steps 2e to 2i in this branch.

#### 2e. Read the clinic's changes: "List rows present in a table"

- Same Location, Library and File (the trigger's **Identifier**). **Table**: custom value `tblClinicWork`.
- **Settings** > **Pagination**: On, **Threshold** 5000. Without it the action returns only the first 256 rows.
- Rename the step `Clinic Work`.

The connector only sees rows inside the table. The Clinic Template Tool keeps every Clinic Work row inside `tblClinicWork`.

#### 2f. Rows already in the inbox for this file: "List rows present in a table"

- **Location / Document Library / File**: pick the master Sohar workbook with the file picker. **Table**: `tblClinicWorkInbox`.
- **Filter Query** (expression; it doubles any apostrophe in the file name):
  `concat('File eq ''', replace(triggerBody()?['{FilenameWithExtension}'], '''', ''''''), '''')`
- **Pagination**: On, **Threshold** 5000.
- Rename the step `Inbox`.

#### 2g. Data Operation "Select" `Known`

- **From**: `body('Inbox')?['value']`
- **Map** (switch to text mode): `concat(item()?['Kind'], '|', item()?['Key'], '|', item()?['When'])`

#### 2h. Data Operation "Filter array" `New`

- **From**: `body('Clinic_Work')?['value']`
- Condition (advanced mode):
  `@and(not(empty(item()?['Kind'])), not(contains(body('Known'), concat(item()?['Kind'], '|', item()?['Key'], '|', item()?['When']))))`

#### 2i. Apply to each `body('New')`: Excel Online (Business) "Add a row into a table"

- **Settings** > **Concurrency control**: leave it off, or set the degree to 1. Microsoft's limits page gives 1 as the default.
  One row at a time keeps the inbox in order.
- **File**: the master Sohar workbook (file picker). **Table**: `tblClinicWorkInbox`.
- Row fields:
  - **Location**: `outputs('Location')`
  - **Kind, Key, MRN, LASFIR, Value, Note, By, When**: `items('Apply_to_each')?['Kind']` and so on, one per column, copied as they are.
  - **File**: `triggerBody()?['{FilenameWithExtension}']`
  - **Received**: `utcNow()`

### 3. Secure inputs and outputs (do not skip)

Run history keeps every step's inputs and outputs for about a month. Microsoft's pages say 28 days in one place and 30 in
another. Without this step, MRNs and notes would sit in run history.

On each of these steps, open **Settings** > **Security** and turn on **Secure inputs** and **Secure outputs**: `SE Info`,
`Location`, `Clinic Work`, `Inbox`, `Known`, `New`, and **Add a row into a table**. Microsoft's guidance is to secure every
downstream step that handles the data as well, not only the step that first reads it. A secured step shows a lock icon.

### 4. Error email: Office 365 Outlook "Send an email (V2)" after the `Try` scope

- **Settings** > **Run after**: `Try` **has failed** and **has timed out** (clear "is successful").
- To: you. Subject: `Clinic inbox flow failed`. Body: the file name only (`triggerBody()?['{FilenameWithExtension}']`). **No row data,
  no MRNs, no notes.**

## Test (synthetic data only)

1. In a demo or test master workbook, set up the inbox. Export a Clinic Template Tool file for a made-up location. Put it in a test
   folder named exactly like that location.
2. Open it in the Clinic Template Tool, tick one task done with a note, and add one patient note. Wait for the flow run. The
   trigger polls, so this takes minutes.
3. Check that the run history shows locks, not values, on the secured steps. Check that at least two rows arrived in the inbox, and
   that Key, MRN and When are exactly as on the clinic file's Clinic Work sheet (leading zeros included). Ticking a task and adding
   its note are two saves of the same Task done row, so if a run lands between them you get that row twice (with different When
   times), three rows in all. That is fine: the newest Task done row carries both the tick and the note.
4. In Sohar: Admin tab > Clinic files and emails > **Check again** should say "2 new clinic changes waiting" (the task and the patient
   note; a tick and its note count once). Click **Apply clinic changes** and check NextGen Tasks (Done, Clinic Note) and Clinic Notes.
5. Save the clinic file again with no change. A run may happen; nothing is added, and Sohar still says nothing new.
6. Then leave the file alone for at least 30 minutes with nobody editing it. Check that no more runs start for it, and that the
   file's version history shows no new versions by the flow's account. If runs keep coming, the third trigger condition is not
   working: switch the flow off and build the scheduled version (end of step 1).
7. Move the file into a different clinic's folder, then open it there and change one cell (or upload a copy into that folder).
   Moving alone does not start a run (see Known limits). The run should stop and email you.
8. Put a copy of the file in a subfolder of its own clinic's folder (for example `Archive`) and change one cell. That run should also
   stop and email you, because the file is not directly in its clinic's folder.

## Known limits

- **Not instant.** The SharePoint trigger polls. Microsoft says a run can start a few minutes after the save, depending on the flow
  owner's license (5 minutes on Microsoft 365 plans, 15 on Free). A row added by "Add a row into a table" can take up to 30 seconds to
  show in the workbook.
- **Every autosave can start a run.** That is expected: step 2h skips rows already copied, so most runs add nothing. Each run
  still counts against the owner's Power Platform request limits.
- **The flow's own reads can change the file.** Microsoft says a read-only Excel action may still save a new version. Without the
  third trigger condition that could start runs on its own, forever. Test step 6 checks it stopped.
- **Moving a file does not start a run.** Microsoft's SharePoint connector notes say a move keeps the file's created and modified
  times, so "this action doesn't trigger any flows". Their page describes moves between libraries; a move inside the library most
  likely behaves the same. A file moved into the wrong folder is checked at its next save.
- **Rows that change nothing are not shown.** A note saved blank, or a note edited after a tick Sohar already applied, still reaches
  the inbox. Sohar files such rows away without counting them, so the waiting count can be lower than the number of new inbox rows.
- **25 MB** maximum file size for the Excel connector, and at most 5 MB per request.
- **Locks and editing at the same time.** Microsoft says a file may stay locked for updates for up to 6 minutes after the connector
  last used it. They also say editing one file from several clients at once (Excel desktop, Excel on the web, Power Automate) is
  "not supported" and can cause merge conflicts. The master workbook will usually be open in Excel while the flow adds rows.
  Sohar never writes the inbox table, which narrows the overlap, but nobody can promise it never conflicts. Watch the first weeks.
  If rows go missing or Excel reports a conflict, note the file name and time. Then compare the clinic file's Clinic Work sheet
  with the inbox.
- **Retries.** The default retry policy retries throttled (429), timed-out (408) and server (5xx) errors. Not confirmed: which
  error a lock returns, and whether that error is retried. If a run fails on a lock, the next save of that clinic
  file runs the flow again and step 2h copies whatever was missed. Microsoft also notes that a timeout followed by a retry can insert
  a row twice. Sohar applies a duplicate once, so that is harmless.
- **Filtered reads can lag.** Microsoft notes that List rows with a Filter Query can return slightly old data. At worst the flow
  copies a row twice; Sohar applies it once.
- **Throttling.** 100 connector calls per minute per connection. A file with many new rows adds them one call each and may slow
  down, with retries; it should not fail.
- **Queued runs with concurrency 1.** Only 1 run at a time, plus about 10 waiting. Microsoft says extra trigger events beyond
  that "might be re-tried" but may not run. A clinic's later saves catch up on anything missed. A change saved once and never again
  could wait until that file's next save.
- **Flow turned off.** Microsoft turns off a flow that fails continuously for 14 days. A flow not triggered for 90 days can also be
  turned off on some licenses. The owner gets an email first. If the flow owner leaves MHS, the flow drops to the lowest
  performance profile. Give it a co-owner.
- **Filter Query** supports only `eq`, `ne`, `contains`, `startswith`, `endswith` and alphanumeric column names. `File` and
  `Setting` are fine.

Sources checked (Microsoft Learn, 2026-10-23): Excel Online (Business) connector reference (known issues, limits, actions);
SharePoint connector reference (triggers; its outputs are listed only as dynamic); "Microsoft SharePoint Connector for Power
Automate" (known limitations: move files and flow runs); Power Automate "Limits of automated, scheduled, and instant flows"; "Optimize Power
Automate triggers"; "Secure data used in cloud flows" and "Manage sensitive input like passwords"; "Troubleshoot Power Automate
trigger problems"; Azure Logic Apps "Handle workflow errors and exceptions" and "Secure access and data in workflows".
