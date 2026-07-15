<#
  UpdateFundMaster.ps1
  Refreshes the MF_Master sheet (Fund Name / Scheme Name / ISIN) from the AMFI
  daily NAV file (NAVAll.txt). The Scheme Name dropdowns on the MutualFunds and
  MF_SIP sheets resize automatically. No Python needed; uses desktop Excel.

  Run it via UpdateFundMaster.bat (double-click). Keep the tracker workbook in
  this folder. Offline fallback: save the AMFI file as NAVAll.txt here and re-run.
#>
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Start-Transcript -Path (Join-Path $here "UpdateFundMaster_log.txt") -Force | Out-Null
$excel = $null; $wb = $null

try {
  # ---------- 1. get the AMFI NAV file ----------
  $navFile = Join-Path $here "NAVAll.txt"
  if (-not (Test-Path $navFile)) {
    $url = "https://www.amfiindia.com/spages/NAVAll.txt"
    Write-Host "Downloading AMFI NAV file..."
    try {
      Invoke-WebRequest -Uri $url -OutFile $navFile -UserAgent "Mozilla/5.0" -TimeoutSec 60
    } catch {
      throw "Could not download NAVAll.txt. Save it manually from $url into this folder and re-run. ($_)"
    }
  } else {
    Write-Host "Using local NAVAll.txt in this folder."
  }

  # ---------- 2. parse Fund House / Scheme Name / ISIN ----------
  $rows = New-Object System.Collections.ArrayList
  $house = ""
  foreach ($line in [System.IO.File]::ReadLines($navFile)) {
    $line = ("" + $line).Trim()
    if (-not $line) { continue }
    if ($line.IndexOf(';') -lt 0) {
      if ($line.Contains("Schemes(") -or $line.StartsWith("Scheme Code")) { continue }
      $house = $line
      continue
    }
    $p = $line.Split(';')
    if ($p.Count -lt 6 -or $p[0] -eq "Scheme Code") { continue }
    $isin = ""
    foreach ($idx in 1, 2) {
      $cand = ("" + $p[$idx]).Trim().ToUpper()
      if ($cand.Length -eq 12 -and $cand.StartsWith("INF")) { $isin = $cand; break }
    }
    if (-not $isin) { continue }
    [void]$rows.Add(@($house, $p[3].Trim(), $isin))
  }
  Write-Host ("Parsed {0} schemes with ISIN from AMFI." -f $rows.Count)
  if ($rows.Count -lt 1000) { throw "Suspiciously few rows parsed - is NAVAll.txt complete?" }

  # ---------- 3. write into MF_Master ----------
  $books = Get-ChildItem -Path $here -Filter "Family_Portfolio_Tracker.xlsx"
  if (-not $books) { $books = Get-ChildItem -Path $here -Filter "*Tracker*.xlsx" }
  if (-not $books) { throw "Family_Portfolio_Tracker.xlsx not found in this folder." }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($books[0].FullName)
  $ws = $null
  foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "MF_Master") { $ws = $sh } }
  if (-not $ws) { throw "MF_Master sheet not found in the workbook." }

  $oldLast = $ws.UsedRange.Rows.Count
  if ($oldLast -ge 4) { $ws.Range("A4:C" + $oldLast).ClearContents() | Out-Null }

  # keep the master sorted by scheme name - the dropdown type-ahead filter needs it
  $n = $rows.Count
  $keys = New-Object string[] $n
  $items = New-Object object[] $n
  for ($i = 0; $i -lt $n; $i++) { $keys[$i] = ("" + $rows[$i][1]); $items[$i] = $rows[$i] }
  [Array]::Sort($keys, $items, [System.StringComparer]::OrdinalIgnoreCase)
  $arr = New-Object 'object[,]' $n, 3
  for ($i = 0; $i -lt $n; $i++) {
    $arr[$i, 0] = $items[$i][0]
    $arr[$i, 1] = $items[$i][1]
    $arr[$i, 2] = $items[$i][2]
  }
  $ws.Range("A4:C" + (3 + $n)).Value2 = $arr
  $ws.Range("E2").Value2 = (Get-Date).ToString("dd-MM-yyyy")
  $wb.Save()
  Write-Host ("MF_Master refreshed: {0} schemes. Scheme dropdowns resize automatically." -f $n) -ForegroundColor Green
  $wb.Close($false); $wb = $null
}
catch { Write-Host "ERROR: $_" -ForegroundColor Red }
finally {
  if ($wb)    { try { $wb.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null 2>$null
  Stop-Transcript | Out-Null
  Read-Host "`nPress Enter to close"
}
