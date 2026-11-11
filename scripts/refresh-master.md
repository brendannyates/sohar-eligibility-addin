# refresh-master.ps1: the 7:30 AM Power Query refresh

Microsoft has no unattended Power Query refresh for a workbook in Excel for the web:
- **Power Automate.** A flow's `refreshAllDataConnections` returns success and does nothing for Power Query.
- **Add-ins.** An add-in can't start a refresh.

So this script opens the master in **desktop Excel** on Brendan's PC, refreshes every query, saves and closes.

## What it does
1. **Skips** (exit 2), and closes without saving, when the workbook has no Power Query yet. It also skips when the workbook is open anywhere it can see: Excel's `~$` owner file, a file lock, or an Excel on this PC with the file open. It never closes or kills anyone's Excel.
2. **Refreshes.** It opens the workbook hidden, turns off background refresh for each connection, runs `RefreshAll`, then waits with `CalculateUntilAsyncQueriesDone` (timeout `-TimeoutMinutes`, default 10).
3. **Saves** only when every query finished (exit 0). Otherwise it closes without saving (exit 1), so a failed refresh never leaves half-loaded tables.
4. **Logs** one line to `Data Drop\refresh.log`:
   ```
   2026-11-11T07:30:41 ok 38s tblNGAppts=1054 tblEligibilityChecks=426 tblPatientHistory=407 tblCardOnFile=118 tblCopayActivity=61 tblOpenTickets=50 tblPatientBalances=missing
   2026-11-12T07:30:02 skipped 0s someone has it open (owner file)
   2026-11-13T07:31:10 failed 70s <the query's error message>
   ```
   The line holds counts only, never a cell value. `missing` means that query isn't set up yet (qBalances waits on its NextGen report).

## Try it
```powershell
.\refresh-master.ps1 -WhatIf
.\refresh-master.ps1 -Workbook "<a copy of the master>.xlsx" -Log "$env:TEMP\refresh.log"
```

## Schedule it (Windows Task Scheduler)
```powershell
$a = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<add-in repo>\scripts\refresh-master.ps1"'
$t = New-ScheduledTaskTrigger -Daily -At 7:30am
Register-ScheduledTask -TaskName 'Sohar master refresh' -Action $a -Trigger $t -Description 'Refresh the master workbook Power Query at 7:30 AM'
```
- The task runs as you, only when you're signed in, because desktop Excel needs your session and your Microsoft 365 sign-in.
- The PC has to be on.
- The Claude scheduled-tasks tool can run the same command instead.

## When it skips or fails
The Sohar pane and full screen show "Data aging" or "Data stale". Press Data > Refresh All yourself, then check `refresh.log` for the reason.

Microsoft doesn't recommend unattended Office automation. A refresh while someone co-authors the master can cause merge errors; that's why the script skips any open file and runs before the day starts.
