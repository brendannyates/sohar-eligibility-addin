<#
refresh-master.ps1 (Sohar v0.30): refresh the master workbook's Power Query tables in desktop Excel, unattended.
Microsoft gives no other way: a Power Automate flow or an Office Script can't refresh Power Query, and an add-in can't start
a refresh. Run it at 7:30 AM, after the 7 AM rotation writes its CSVs (see refresh-master.md).

It skips (exit 2) when the workbook is open anywhere it can see (a ~$ owner file, a lock, or an Excel already running with it).
It never kills Excel. On a failed refresh it closes WITHOUT saving (exit 1). The log has counts only: no cell values.
#>
[CmdletBinding()]
param(
  [string]$Workbook = (Join-Path $env:USERPROFILE 'OneDrive - Mindful Health Solutions\Projects\Clinic Workbook Pilot\Sohar_Eligibility_Template.xlsx'),
  [string]$Log = (Join-Path $env:USERPROFILE 'OneDrive - Mindful Health Solutions\Projects\Clinic Workbook Pilot\Data Drop\refresh.log'),
  [int]$TimeoutMinutes = 10,
  [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
$Tables = 'tblNGAppts', 'tblEligibilityChecks', 'tblPatientHistory', 'tblCardOnFile', 'tblCopayActivity', 'tblOpenTickets', 'tblPatientBalances'
$started = Get-Date

function Write-Line([string]$outcome, [string]$detail) {
  $secs = [math]::Round(((Get-Date) - $started).TotalSeconds)
  $line = '{0} {1} {2}s {3}' -f (Get-Date).ToString('s'), $outcome, $secs, $detail
  Write-Output $line
  if (-not $WhatIf) {
    $dir = Split-Path $Log -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }
    Add-Content -Path $Log -Value $line -Encoding utf8
  }
}

function Test-InUse([string]$path) {
  if (-not (Test-Path $path)) { return 'the workbook is not there' }
  $owner = Join-Path (Split-Path $path -Parent) ('~$' + (Split-Path $path -Leaf).Substring([math]::Min(2, (Split-Path $path -Leaf).Length - 1)))
  $owner2 = Join-Path (Split-Path $path -Parent) ('~$' + (Split-Path $path -Leaf))
  if ((Test-Path $owner) -or (Test-Path $owner2)) { return 'someone has it open (owner file)' }
  try { $fs = [IO.File]::Open($path, 'Open', 'ReadWrite', 'None'); $fs.Close() } catch { return 'the file is locked' }
  try {
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    foreach ($wb in @($xl.Workbooks)) { if ($wb.FullName -eq $path) { return 'it is open in Excel on this computer' } }
  } catch { }  # no Excel running: fine
  return ''
}

$why = Test-InUse $Workbook
if ($why) { Write-Line 'skipped' $why; exit 2 }
if ($WhatIf) { Write-Line 'whatif' "would refresh $([IO.Path]::GetFileName($Workbook))"; exit 0 }

$excel = $null; $wb = $null; $code = 1
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.EnableEvents = $false
  $wb = $excel.Workbooks.Open($Workbook, 0, $false)
  # Nothing to refresh yet: close without saving, so the master is never re-saved for no reason.
  if (@($wb.Queries).Count -eq 0) {
    Write-Line 'skipped' 'no Power Query in the workbook yet (docs/powerquery/README.md)'
    $code = 2
  } else {
    # Wait for each query in turn, so RefreshAll returns only when the data is in.
    foreach ($c in @($wb.Connections)) {
      try { if ($c.Type -eq 1) { $c.OLEDBConnection.BackgroundQuery = $false } } catch { }
    }
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $wb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()
    if ((Get-Date) -gt $deadline) { throw "the refresh took longer than $TimeoutMinutes minutes" }
    $errors = @()
    foreach ($q in @($wb.Queries)) { try { if ($q.Refreshing) { $errors += "$($q.Name) still refreshing" } } catch { } }
    $counts = foreach ($t in $Tables) {
      $n = $null
      foreach ($s in @($wb.Worksheets)) { try { $lo = $s.ListObjects.Item($t); $n = $lo.ListRows.Count; break } catch { } }
      if ($null -ne $n) { "$t=$n" } else { "$t=missing" }
    }
    if ($errors.Count) { throw ($errors -join '; ') }
    $wb.Save()
    Write-Line 'ok' ($counts -join ' ')
    $code = 0
  }
} catch {
  # Never save a half-refreshed workbook.
  Write-Line 'failed' ($_.Exception.Message -replace '\s+', ' ')
  $code = 1
} finally {
  if ($wb) { try { $wb.Close($false) } catch { } ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
  if ($excel) { try { $excel.Quit() } catch { } ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $code
