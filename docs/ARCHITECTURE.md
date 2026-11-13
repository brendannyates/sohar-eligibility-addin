# Eligibility system architecture (v0.31 contract)

Written 2026-11-12 (Brendan's rules of 11/12). It covers the Sohar Eligibility add-in and its central workbook, the Sohar
Flow "MHS Eligibility - Discovery vs NextGen", and the 7 AM rotation (`~/.claude/skills/eligibility-check/scripts/rotation`).
This is the contract the clinic app implements later. Synthetic examples only: no names, MRNs, member IDs, TINs or
contracted rates. When a rule changes, change it here and in the pure module that owns it, in the same PR.

## 1. Who does what

| Part | Decides | Writes |
| - | - | - |
| **Appointment import** (add-in, pane or full screen) | Which visits need a verification or a plan-year estimate, using the freshness rule (section 3), Config and Verification History | Sohar import CSVs and the Flow import CSV; then, from Sohar's results, `tblSoharResults`, `tblVerificationHistory` and `tblFlowResults` |
| **Sohar Flow** (production) | Discovery vs the NextGen policy, which policy to keep, network status, and the cost estimate when asked | Nothing local. Results are read back by `receiver.js poll` (run GETs) into `Flow Results.csv`, then the add-in imports that file |
| **7 AM rotation** | Only visits the import didn't cover under the same rule (section 3) | `Eligibility Checks.csv`, `Eligibility History.csv`, the add-in import CSVs, `NextGen Tasks (rotation).csv`, the digest |
| **Power Query** (central workbook) | Nothing | Loads the rotation's CSVs into the query tables (`tblEligibilityChecks`, `tblPatientHistory`, …) |

The rules live in pure modules, with tests: `core.js` (freshness, pool), `fees.js` (codes, plan year), `rules.js`
(service line, flags) and `record.js` (patient views). The rotation `require`s the same files from the add-in folder, so
the import and the rotation can't disagree.

## 2. Source of a check

Every check carries a **Source**: `Import` (Sohar CSV results from the appointment import), `Flow` (a Flow run),
`Rotation` (the 7 AM rotation), `Dashboard` (a Sohar dashboard export the add-in reads, e.g. a CSV import's results) or
`Webhook` (the Sohar webhook flow). A check from any source covers a visit. Source only says where the check came from.

## 3. Verification freshness (`core.reverifyDue(check, visit, line, cfg)`)

A visit at time *T* is covered by an eligible verification of the same policy made at *t* when
**t ≥ max(T − 72 h, the 1st of T's month at 00:00 Pacific)**:

- A visit on the 1st or later is never covered by a check from the previous month. The check waits until the 1st.
- **On the 1st of each month**, every active policy with a visit that month is re-verified (Config `Re-verify on the 1st`).
- **TMS treatments** (daily sessions): one check per week, Monday, plus the 1st. TMS F/U, MT/MTR and consults follow the
  provider rule.

Service lines come from `rules.serviceLine(event)`: `Provider`, `TMS treatment`, `ESK`, `IVK`. Config rows (added by
`ensureConfigRows`; editable in the Config section):

| Setting | Default | Meaning |
| - | - | - |
| `Re-verify window (hours)` | 72 | The *T − hours* part of the rule |
| `Re-verify on the 1st` | On | Month boundary: re-verify on or after the 1st |
| `Re-verify: Provider` | `Window` | `Window` (the rule above), `Weekly` (Monday), or a number of days |
| `Re-verify: TMS treatment` | `Weekly` | |
| `Re-verify: ESK` | `Window` | |
| `Re-verify: IVK` | `Window` | |
| `Include in import: <line>` | On | Off: that service line leaves the pool at the next import (the rotation still covers it) |

`Skip recent verifications` stays. `Re-verify days` and `Re-verify in new month` are replaced by the rows above and
migrate once: days become the window (in hours), and new month becomes `Re-verify on the 1st`.

## 4. Cost estimate (`fees.planYearEstimateDue`, `fees.visitCode`)

- **Once per policy per plan year, for every patient**, only when the policy is eligible and the OOP max isn't met. The
  estimate always quotes **90792, 99214 and 99215** (Config `Plan-year estimate codes`), so PR and coinsurance read as
  fee-schedule figures.
- **Plan year.** If Sohar's coverage start is January 1, the plan year is the calendar year. If not, the policy is
  estimated at both January 1 and its plan start date until the deductible shows which reset is real.
  - Deductible remaining reset to the total in January: the calendar year, and the plan date is marked unreliable.
  - Reset at the plan date: the plan year.
  - Stored per policy as `Reset Basis` = `Unknown`, `Calendar` or `Plan`.
- **No estimate:** VA, TriWest and TRICARE (expected $0), and the OOP max met ($0).
- **Code shown per visit** (`fees.visitCode(event, line, payer)`):

| Visit | Code |
| - | - |
| New patient, consult, TMS consult | 90792 |
| Med management or follow-up, 30 min; TMS F/U; telepsych | 99214 |
| 60 min (transition, ADHD testing review) | 99215 |
| IVK, not Kaiser/VA | 99214, plus the self-pay medication ($400 for infusions 1–6, $450 from 7; Config) |
| ESK session, TMS treatment, MT/MTR | none: verification only, and the tracker BIDF is the reference |

## 5. Tables

### `tblVerificationHistory` (new, add-in owned, hidden, append-only by passing old rows back)
One row per Sohar check the add-in processed (Source Import, Flow or Dashboard), sorted by `Pulled At`. A row is never
replaced. Re-reading the same Sohar ID updates that row's fields only. The columns are the same as the rotation's
`Eligibility History.csv` (`rotation/FEED.md`), so `record.js` can merge the add-in's table with `tblPatientHistory`
(the rotation and the webhook) into one timeline per policy:

`Pulled At, LASFIR, MRN, Per Nbr, Clinic, Kind, Sohar Status, Sohar Payer, Sohar payerId, Plan, Network, Coverage Level,
Coverage Start, Coverage End, Copay, Coinsurance, Deductible, Deductible Remaining, Deductible Status, OOP Max,
OOP Remaining, OOP Status, Benefit Flag, Sohar ID, Sohar Link, Related Sohar IDs, Policy Key, Projected Patient
Responsibility, Estimate Note, Source`

`Policy Key` = Sohar payerId + `#` + the last 4 characters of the member ID. The full member ID never goes on the sheet.
It tells two policies at the same payer apart, so an old policy's checks stay grouped under that policy after a
discovery finds a new one.

### `tblAccountAuditHistory` (new, add-in owned, hidden)
Every audit run, appended: `audit.AUDIT_COLS` + `Run At` + `Run By`. `tblAccountAudits` keeps showing the latest per
patient, as now.

### Flow import (`record.FLOW_IMPORT_COLS`, v27)
`id, patient.firstName, patient.lastName, patient.dateOfBirth, patient.state, placeOfServiceCode, specialtyCode,
chartPayerId, chartMemberId, providerGroupName, estimate, rateTable`

- **No `tin`.** The Flow template looks the TIN up by `providerGroupName`. (Sohar's network-status provider groups carry
  no TIN: their fields are `id`, `name`, `status`, `exceptions`.)
- **`providerGroupName`** by patient state:
  - CA: `MINDFUL HEALTH SOLUTIONS A PROFESSIONAL MEDICAL CORPORATION`
  - TX: `Mindful Health Solutions TX PA`
  - WA: `Mindful Health Solutions Wa Pc`
- **`estimate`:** `yes` only when section 4 says the policy is due.
- **`rateTable`:** a JSON string `{"<Sohar payerId>": [{"code": "99214", "contractedRate": 0.00}, …]}`, one entry for the
  chart payer and the payers discovery commonly finds for the patient's state, built with the existing fee lookups
  (Fee Locations, Fee Payer Map, `lookupFeePayer`, MD/NP level).
- The header is pinned to the production Flow's downloaded CSV template, and a test fails if they drift.

### `tblFlowResults` (`record.FLOW_COLS`)
Unchanged, plus `Source` (`Flow`), `Policy Key`, `Reset Basis`, and per-code PR `Est 90792`, `Est 99214`, `Est 99215`.

## 6. Flow input and output ("MHS Eligibility - Discovery vs NextGen", production)

Production flow id `470f1e51-cddf-4050-a706-c6fc213aece9` (template v1 = our v27). The sandbox flow `e8839155` runs the
same template (v32). Verified in sandbox: `rateTable` as a JSON string parses (`$eval` works); outputs and overrides must
start with `# Eligibility system architecture (v0.31 contract)

Written 2026-11-12 (Brendan's rules of 11/12). It covers the Sohar Eligibility add-in and its central workbook, the Sohar
Flow "MHS Eligibility - Discovery vs NextGen", and the 7 AM rotation (`~/.claude/skills/eligibility-check/scripts/rotation`).
This is the contract the clinic app implements later. Synthetic examples only: no names, MRNs, member IDs, TINs or
contracted rates. When a rule changes, change it here and in the pure module that owns it, in the same PR.

## 1. Who does what

| Part | Decides | Writes |
| - | - | - |
| **Appointment import** (add-in, pane or full screen) | Which visits need a verification or a plan-year estimate, using the freshness rule (section 3), Config and Verification History | Sohar import CSVs and the Flow import CSV; then, from Sohar's results, `tblSoharResults`, `tblVerificationHistory` and `tblFlowResults` |
| **Sohar Flow** (production) | Discovery vs the NextGen policy, which policy to keep, network status, and the cost estimate when asked | Nothing local. Results are read back by `receiver.js poll` (run GETs) into `Flow Results.csv`, then the add-in imports that file |
| **7 AM rotation** | Only visits the import didn't cover under the same rule (section 3) | `Eligibility Checks.csv`, `Eligibility History.csv`, the add-in import CSVs, `NextGen Tasks (rotation).csv`, the digest |
| **Power Query** (central workbook) | Nothing | Loads the rotation's CSVs into the query tables (`tblEligibilityChecks`, `tblPatientHistory`, …) |

The rules live in pure modules, with tests: `core.js` (freshness, pool), `fees.js` (codes, plan year), `rules.js`
(service line, flags) and `record.js` (patient views). The rotation `require`s the same files from the add-in folder, so
the import and the rotation can't disagree.

## 2. Source of a check

Every check carries a **Source**: `Import` (Sohar CSV results from the appointment import), `Flow` (a Flow run),
`Rotation` (the 7 AM rotation), `Dashboard` (a Sohar dashboard export the add-in reads, e.g. a CSV import's results) or
`Webhook` (the Sohar webhook flow). A check from any source covers a visit. Source only says where the check came from.

## 3. Verification freshness (`core.reverifyDue(check, visit, line, cfg)`)

A visit at time *T* is covered by an eligible verification of the same policy made at *t* when
**t ≥ max(T − 72 h, the 1st of T's month at 00:00 Pacific)**:

- A visit on the 1st or later is never covered by a check from the previous month. The check waits until the 1st.
- **On the 1st of each month**, every active policy with a visit that month is re-verified (Config `Re-verify on the 1st`).
- **TMS treatments** (daily sessions): one check per week, Monday, plus the 1st. TMS F/U, MT/MTR and consults follow the
  provider rule.

Service lines come from `rules.serviceLine(event)`: `Provider`, `TMS treatment`, `ESK`, `IVK`. Config rows (added by
`ensureConfigRows`; editable in the Config section):

| Setting | Default | Meaning |
| - | - | - |
| `Re-verify window (hours)` | 72 | The *T − hours* part of the rule |
| `Re-verify on the 1st` | On | Month boundary: re-verify on or after the 1st |
| `Re-verify: Provider` | `Window` | `Window` (the rule above), `Weekly` (Monday), or a number of days |
| `Re-verify: TMS treatment` | `Weekly` | |
| `Re-verify: ESK` | `Window` | |
| `Re-verify: IVK` | `Window` | |
| `Include in import: <line>` | On | Off: that service line leaves the pool at the next import (the rotation still covers it) |

`Skip recent verifications` stays. `Re-verify days` and `Re-verify in new month` are replaced by the rows above and
migrate once: days become the window (in hours), and new month becomes `Re-verify on the 1st`.

## 4. Cost estimate (`fees.planYearEstimateDue`, `fees.visitCode`)

- **Once per policy per plan year, for every patient**, only when the policy is eligible and the OOP max isn't met. The
  estimate always quotes **90792, 99214 and 99215** (Config `Plan-year estimate codes`), so PR and coinsurance read as
  fee-schedule figures.
- **Plan year.** If Sohar's coverage start is January 1, the plan year is the calendar year. If not, the policy is
  estimated at both January 1 and its plan start date until the deductible shows which reset is real.
  - Deductible remaining reset to the total in January: the calendar year, and the plan date is marked unreliable.
  - Reset at the plan date: the plan year.
  - Stored per policy as `Reset Basis` = `Unknown`, `Calendar` or `Plan`.
- **No estimate:** VA, TriWest and TRICARE (expected $0), and the OOP max met ($0).
- **Code shown per visit** (`fees.visitCode(event, line, payer)`):

| Visit | Code |
| - | - |
| New patient, consult, TMS consult | 90792 |
| Med management or follow-up, 30 min; TMS F/U; telepsych | 99214 |
| 60 min (transition, ADHD testing review) | 99215 |
| IVK, not Kaiser/VA | 99214, plus the self-pay medication ($400 for infusions 1–6, $450 from 7; Config) |
| ESK session, TMS treatment, MT/MTR | none: verification only, and the tracker BIDF is the reference |

## 5. Tables

### `tblVerificationHistory` (new, add-in owned, hidden, append-only by passing old rows back)
One row per Sohar check the add-in processed (Source Import, Flow or Dashboard), sorted by `Pulled At`. A row is never
replaced. Re-reading the same Sohar ID updates that row's fields only. The columns are the same as the rotation's
`Eligibility History.csv` (`rotation/FEED.md`), so `record.js` can merge the add-in's table with `tblPatientHistory`
(the rotation and the webhook) into one timeline per policy:

`Pulled At, LASFIR, MRN, Per Nbr, Clinic, Kind, Sohar Status, Sohar Payer, Sohar payerId, Plan, Network, Coverage Level,
Coverage Start, Coverage End, Copay, Coinsurance, Deductible, Deductible Remaining, Deductible Status, OOP Max,
OOP Remaining, OOP Status, Benefit Flag, Sohar ID, Sohar Link, Related Sohar IDs, Policy Key, Projected Patient
Responsibility, Estimate Note, Source`

`Policy Key` = Sohar payerId + `#` + the last 4 characters of the member ID. The full member ID never goes on the sheet.
It tells two policies at the same payer apart, so an old policy's checks stay grouped under that policy after a
discovery finds a new one.

### `tblAccountAuditHistory` (new, add-in owned, hidden)
Every audit run, appended: `audit.AUDIT_COLS` + `Run At` + `Run By`. `tblAccountAudits` keeps showing the latest per
patient, as now.

### Flow import (`record.FLOW_IMPORT_COLS`, v27)
`id, patient.firstName, patient.lastName, patient.dateOfBirth, patient.state, placeOfServiceCode, specialtyCode,
chartPayerId, chartMemberId, providerGroupName, estimate, rateTable`

- **No `tin`.** The Flow template looks the TIN up by `providerGroupName`. (Sohar's network-status provider groups carry
  no TIN: their fields are `id`, `name`, `status`, `exceptions`.)
- **`providerGroupName`** by patient state:
  - CA: `MINDFUL HEALTH SOLUTIONS A PROFESSIONAL MEDICAL CORPORATION`
  - TX: `Mindful Health Solutions TX PA`
  - WA: `Mindful Health Solutions Wa Pc`
- **`estimate`:** `yes` only when section 4 says the policy is due.
- **`rateTable`:** a JSON string `{"<Sohar payerId>": [{"code": "99214", "contractedRate": 0.00}, …]}`, one entry for the
  chart payer and the payers discovery commonly finds for the patient's state, built with the existing fee lookups
  (Fee Locations, Fee Payer Map, `lookupFeePayer`, MD/NP level).
- The header is pinned to the production Flow's downloaded CSV template, and a test fails if they drift.

### `tblFlowResults` (`record.FLOW_COLS`)
Unchanged, plus `Source` (`Flow`), `Policy Key`, `Reset Basis`, and per-code PR `Est 90792`, `Est 99214`, `Est 99215`.

 (a value starting with `(` comes back literally); `$` does not work inside filters; the estimate returns
lines 90792, 99214 and 99215 with `patientResponsibility`.

- **Input:** the Flow import columns above, as one object.
- **Steps:**
  1. Discovery.
  2. Compare with the chart policy (payer, member ID prefix).
  3. Verify the discovered policy, the chart policy or both.
  4. Pick the policy.
  5. Network status, matched case-insensitively.
  6. Cost estimate, only when `estimate = yes`, the picked policy is eligible, the provider group's status is INN
     (matched case-insensitively via `providerGroups{$uppercase(name):$}`), the group has a TIN in the template, and
     `rateTable` has the picked payer.
- **Output:** `id`, `providerGroupName`, `estimateRequested`, `ratesForPayer`, `rateTableParsed`, the discovery response,
  both verification responses, `pickedPolicy` (`verificationId`, `payerId`, `status`, `from`), `networkStatusResponse`,
  `providerGroupStatus`, and `costEstimate` (`lines` with per-line adjustments, `requestId`, `paidAmount`,
  `verificationId`, `patientResponsibility`, `postDeductiblePatientResponsibility`).
  Sohar's dashboard run export carries none of this, so results come only from `GET /v1/runs/{id}`.

## 7. Exports
The add-in's own exports are self-contained, interactive HTML pages in the add-in's design (tables, search, filters,
light and dark, Sohar links), saved to OneDrive or SharePoint only. Sohar import files stay CSV, the format Sohar requires.

## 8. Privacy
- **Sheets and CSVs:** LASFIR and MRN only. No names, DOBs or member IDs, except the Sohar import CSVs, which Sohar needs
  and which are saved to OneDrive only.
- **TINs:** only in Sohar (provider groups and the Flow template). Never in this repo, Config or a CSV.
- **Contracted rates:** only in hidden fee sheets and in the `rateTable` sent to Sohar. Never in a commit.
