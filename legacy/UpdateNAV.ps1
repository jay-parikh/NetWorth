<#
  UpdateNAV.ps1
  Refreshes "Current NAV" on any sheet that has an ISIN column, matched by ISIN,
  from the AMFI daily NAV file (NAVAll.txt). No Python needed; uses desktop Excel.

  Run it via UpdateNAV.bat (double-click). Keep the tracker workbook in this folder.
  Offline fallback: save the AMFI file as NAVAll.txt in this folder and run again.
#>
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Start-Transcript -Path (Join-Path $here "UpdateNAV_log.txt") -Force | Out-Null
$excel = $null; $wb = $null

function Get-ColIndex($ws, $hr, $colc, [string[]]$wanted) {
  for ($c = 1; $c -le $colc; $c++) {
    $h = ("" + $ws.Cells.Item($hr, $c).Value2).Trim().ToLower()
    foreach ($w in $wanted) { if ($h -eq $w) { return $c } }
  }
  return 0
}


function Get-Xirr($flows) {   # $flows = array of @{d=[datetime]; a=[double]}
  $flows = @($flows | Where-Object { $_.a -ne 0 })
  if ($flows.Count -lt 2) { return $null }
  $pos=$false; $neg=$false
  foreach ($f in $flows) { if ($f.a -gt 0) { $pos=$true } elseif ($f.a -lt 0) { $neg=$true } }
  if (-not ($pos -and $neg)) { return $null }
  $flows = @($flows)
  $t0 = $flows[0].d; $t1 = $flows[0].d
  foreach ($f in $flows) { if ($f.d -lt $t0) { $t0 = $f.d }; if ($f.d -gt $t1) { $t1 = $f.d } }
  if ($t1 -eq $t0) { return $null }   # all flows on one date -> XIRR undefined
  $npv = { param($r) $s=0.0
    foreach ($f in $flows) { $s += $f.a / [math]::Pow(1.0+$r, (($f.d - $t0).TotalDays)/365.0) }
    return $s }
  $lo=-0.9999; $hi=10.0
  $flo=& $npv $lo; $fhi=& $npv $hi
  if ($flo*$fhi -gt 0) { return $null }
  for ($i=0; $i -lt 200; $i++) {
    $mid=($lo+$hi)/2.0; $fm=& $npv $mid
    if ([math]::Abs($fm) -lt 1e-7) { return $mid }
    if ($flo*$fm -lt 0) { $hi=$mid; $fhi=$fm } else { $lo=$mid; $flo=$fm }
  }
  return ($lo+$hi)/2.0
}

function Update-TrackerXirr($wb) {
  # Recompute per-fund + portfolio MF XIRR from the MF_SIP ledger and write plain values.
  $mf=$null; $sip=$null
  foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -eq "MutualFunds") { $mf=$ws }
    if ($ws.Name -eq "MF_SIP")      { $sip=$ws }
  }
  if (-not $mf -or -not $sip) { return }
  # gather transactions (Owner A, ISIN D, Date E, Amount F), rows 4..503
  $tx=@{}
  $srows=[math]::Min($sip.UsedRange.Rows.Count, 503)
  for ($r=4; $r -le $srows; $r++) {
    $own=("" + $sip.Cells.Item($r,1).Value2).Trim()
    $isin=("" + $sip.Cells.Item($r,4).Value2).Trim()
    $dv=$sip.Cells.Item($r,5).Value2; $amt=$sip.Cells.Item($r,6).Value2
    if ($own -and $isin -and $dv -and ($amt -is [double])) {
      $d=[datetime]::FromOADate([double]$dv)
      $key="$own|$isin"
      if (-not $tx.ContainsKey($key)) { $tx[$key]=New-Object System.Collections.ArrayList }
      [void]$tx[$key].Add(@{d=$d; a=-[double]$amt})
    }
  }
  $today=(Get-Date).Date
  $all=New-Object System.Collections.ArrayList
  $wrote=0
  for ($r=4; $r -le 63; $r++) {
    $own=("" + $mf.Cells.Item($r,1).Value2).Trim()
    $isin=("" + $mf.Cells.Item($r,4).Value2).Trim()
    if (-not ($own -and $isin)) { continue }
    $key="$own|$isin"
    $cur=$mf.Cells.Item($r,9).Value2      # Cur. val (Excel keeps this recalculated)
    $x=$null
    if (($cur -is [double]) -and $cur -gt 0 -and $tx.ContainsKey($key)) {
      $flows=New-Object System.Collections.ArrayList
      foreach ($f in $tx[$key]) { [void]$flows.Add($f); [void]$all.Add($f) }
      $term=@{d=$today; a=[double]$cur}
      [void]$flows.Add($term); [void]$all.Add($term)
      $x=Get-Xirr $flows
    }
    if ($x -ne $null) { $mf.Cells.Item($r,12).Value2=[double]$x; $wrote++ }
    else { $mf.Cells.Item($r,12).ClearContents() | Out-Null }
  }
  $port=Get-Xirr $all
  if ($port -ne $null) {
    $mf.Cells.Item(65,12).Value2=[double]$port
    $sip.Cells.Item(2,10).Value2=[double]$port
  }
  Write-Host ("  XIRR: wrote {0} fund value(s){1}" -f $wrote, $(if ($port -ne $null) { ", portfolio " + ("{0:P2}" -f $port) } else { "" }))
}

function Update-PortfolioXirr($wb) {
  # Portfolio-level XIRR across all asset classes -> Dashboard B4 + per-class C20:C24.
  # PPF has no contribution ledger: estimated at its Rate % from the as-on date.
  # Bonds need a Buy Date (column M); rows without one are skipped.
  $sh=@{}
  foreach ($ws in $wb.Worksheets) { $sh[$ws.Name]=$ws }
  $dash=$sh["Dashboard"]
  if (-not $dash) { return }
  $today=(Get-Date).Date
  $flows=@{}
  foreach ($k in @("Equity","MF","FD","PPF","Bonds")) { $flows[$k]=New-Object System.Collections.ArrayList }

  $eq=$sh["Equity"]
  if ($eq) {
    for ($r=4; $r -le 140; $r++) {
      $dv=$eq.Cells.Item($r,13).Value2; $inv=$eq.Cells.Item($r,10).Value2; $cv=$eq.Cells.Item($r,9).Value2
      if ($dv -and ($inv -is [double]) -and $inv -gt 0 -and ($cv -is [double])) {
        [void]$flows["Equity"].Add(@{d=[datetime]::FromOADate([double]$dv); a=-[double]$inv})
        [void]$flows["Equity"].Add(@{d=$today; a=[double]$cv})
      }
    }
  }

  $mf=$sh["MutualFunds"]; $sip=$sh["MF_SIP"]
  if ($mf -and $sip) {
    $tx=@{}
    $srows=[math]::Min($sip.UsedRange.Rows.Count, 503)
    for ($r=4; $r -le $srows; $r++) {
      $own=("" + $sip.Cells.Item($r,1).Value2).Trim()
      $isin=("" + $sip.Cells.Item($r,4).Value2).Trim()
      $dv=$sip.Cells.Item($r,5).Value2; $amt=$sip.Cells.Item($r,6).Value2
      if ($own -and $isin -and $dv -and ($amt -is [double])) {
        $key="$own|$isin"
        if (-not $tx.ContainsKey($key)) { $tx[$key]=New-Object System.Collections.ArrayList }
        [void]$tx[$key].Add(@{d=[datetime]::FromOADate([double]$dv); a=-[double]$amt})
      }
    }
    for ($r=4; $r -le 63; $r++) {
      $own=("" + $mf.Cells.Item($r,1).Value2).Trim()
      $isin=("" + $mf.Cells.Item($r,4).Value2).Trim()
      if (-not ($own -and $isin)) { continue }
      $cur=$mf.Cells.Item($r,9).Value2
      $key="$own|$isin"
      if (($cur -is [double]) -and $cur -gt 0 -and $tx.ContainsKey($key)) {
        foreach ($f in $tx[$key]) { [void]$flows["MF"].Add($f) }
        [void]$flows["MF"].Add(@{d=$today; a=[double]$cur})
      }
    }
  }

  $fd=$sh["FixedDeposits"]
  if ($fd) {
    for ($r=4; $r -le 53; $r++) {
      $p=$fd.Cells.Item($r,4).Value2; $sd=$fd.Cells.Item($r,6).Value2
      $md=$fd.Cells.Item($r,7).Value2; $cv=$fd.Cells.Item($r,9).Value2
      if (($p -is [double]) -and $p -gt 0 -and $sd -and ($cv -is [double])) {
        $end=$today
        if ($md) { $m=[datetime]::FromOADate([double]$md); if ($m -lt $end) { $end=$m } }
        [void]$flows["FD"].Add(@{d=[datetime]::FromOADate([double]$sd); a=-[double]$p})
        [void]$flows["FD"].Add(@{d=$end; a=[double]$cv})
      }
    }
  }

  $ppf=$sh["PPF"]
  if ($ppf) {
    for ($r=4; $r -le 43; $r++) {
      $bal=$ppf.Cells.Item($r,4).Value2; $ad=$ppf.Cells.Item($r,5).Value2; $rate=$ppf.Cells.Item($r,6).Value2
      if (($bal -is [double]) -and $bal -gt 0 -and $ad -and ($rate -is [double])) {
        $asOn=[datetime]::FromOADate([double]$ad)
        if ($asOn -gt $today) { continue }
        $grown=[double]$bal*[math]::Pow(1.0+[double]$rate/100.0, ($today-$asOn).TotalDays/365.0)
        [void]$flows["PPF"].Add(@{d=$asOn; a=-[double]$bal})
        [void]$flows["PPF"].Add(@{d=$today; a=$grown})
      }
    }
  }

  $bd=$sh["Bonds"]
  if ($bd) {
    for ($r=4; $r -le 53; $r++) {
      $bdt=$bd.Cells.Item($r,13).Value2; $inv=$bd.Cells.Item($r,10).Value2; $cv=$bd.Cells.Item($r,11).Value2
      if ($bdt -and ($inv -is [double]) -and $inv -gt 0 -and ($cv -is [double])) {
        [void]$flows["Bonds"].Add(@{d=[datetime]::FromOADate([double]$bdt); a=-[double]$inv})
        [void]$flows["Bonds"].Add(@{d=$today; a=[double]$cv})
      }
    }
  }

  $rowOf=@{Equity=20; MF=21; FD=22; PPF=23; Bonds=24}
  $all=New-Object System.Collections.ArrayList
  foreach ($k in @("Equity","MF","FD","PPF","Bonds")) {
    foreach ($f in $flows[$k]) { [void]$all.Add($f) }
    $x=Get-Xirr $flows[$k]
    if ($x -ne $null) { $dash.Cells.Item($rowOf[$k],3).Value2=[double]$x }
    else { $dash.Cells.Item($rowOf[$k],3).ClearContents() | Out-Null }
  }
  $px=Get-Xirr $all
  if ($px -ne $null) {
    $dash.Cells.Item(4,2).Value2=[double]$px
    Write-Host ("  Portfolio XIRR (all asset classes): {0:P2}" -f $px)
  } else { $dash.Cells.Item(4,2).ClearContents() | Out-Null }
}

try {
  # ---------- 1. get the AMFI NAV file ----------
  $navFile = Join-Path $here "NAVAll.txt"
  $local = Test-Path $navFile
  if (-not $local) {
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

  # ---------- 2. parse ISIN -> NAV ----------
  $ci = [System.Globalization.CultureInfo]::InvariantCulture
  $any = [System.Globalization.NumberStyles]::Any
  $nav = @{}
  foreach ($line in [System.IO.File]::ReadLines($navFile)) {
    if (-not $line -or $line.IndexOf(';') -lt 0) { continue }
    $p = $line.Split(';')
    if ($p.Count -lt 6) { continue }
    $navVal = 0.0
    if (-not [double]::TryParse($p[4], $any, $ci, [ref]$navVal)) { continue }
    if ($navVal -le 0) { continue }
    foreach ($idx in 1, 2) {
      $isin = ("" + $p[$idx]).Trim().ToUpper()
      if ($isin.Length -eq 12 -and $isin.StartsWith("INF")) { $nav[$isin] = $navVal }
    }
  }
  Write-Host ("Parsed {0} scheme NAVs from AMFI." -f $nav.Count)
  if ($nav.Count -eq 0) { throw "No NAVs parsed - is NAVAll.txt in the expected AMFI format?" }

  # ---------- 3. open the workbook(s) and update ----------
  $books = Get-ChildItem -Path $here -Filter "Family_Portfolio_Tracker.xlsx"
  if (-not $books) { $books = Get-ChildItem -Path $here -Filter "*Tracker*.xlsx" }
  if (-not $books) { throw "Family_Portfolio_Tracker.xlsx not found in this folder." }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false

  foreach ($bk in $books) {
    $wb = $excel.Workbooks.Open($bk.FullName)
    $touched = 0
    foreach ($ws in $wb.Worksheets) {
      $ur = $ws.UsedRange; $rows = $ur.Rows.Count; $colc = $ur.Columns.Count
      if ($rows -lt 2) { continue }
      # header row = first of rows 1..5 that has a cell equal to "ISIN"
      $hr = 0
      for ($r = 1; $r -le [math]::Min(5, $rows); $r++) {
        for ($c = 1; $c -le $colc; $c++) {
          $v = $ws.Cells.Item($r, $c).Value2
          if ($v -and ($v.ToString().Trim().ToLower() -eq "isin")) { $hr = $r; break }
        }
        if ($hr) { break }
      }
      if (-not $hr) { continue }
      $ciCol  = Get-ColIndex $ws $hr $colc @("isin")
      $navCol = Get-ColIndex $ws $hr $colc @("current nav", "nav", "current n.a.v.")
      if (-not $ciCol -or -not $navCol) { continue }

      $f = $hr + 1; $n = $rows - $f + 1
      if ($n -lt 1) { continue }
      $isinRng = $ws.Range($ws.Cells.Item($f, $ciCol),  $ws.Cells.Item($rows, $ciCol))
      $navRng  = $ws.Range($ws.Cells.Item($f, $navCol), $ws.Cells.Item($rows, $navCol))
      $isinArr = $isinRng.Value2
      $navArr  = $navRng.Value2                # 2-D array [1..n,1..1]
      $upd = 0
      for ($i = 1; $i -le $n; $i++) {
        $isin = ("" + $isinArr.GetValue($i, 1)).Trim().ToUpper()
        if ($isin -and $nav.ContainsKey($isin)) { $navArr.SetValue([double]$nav[$isin], $i, 1); $upd++ }
      }
      if ($upd -gt 0) { $navRng.Value2 = $navArr; $touched += $upd
        Write-Host ("  {0}!{1}: updated {2} NAV(s)" -f $bk.Name, $ws.Name, $upd) }
    }
    $excel.CalculateFull()
    Update-TrackerXirr $wb
    Update-PortfolioXirr $wb
    $wb.Save(); Write-Host ("Saved {0} ({1} NAVs)." -f $bk.Name, $touched)
    $wb.Close($false); $wb = $null
  }
  Write-Host "Done."
}
catch { Write-Host "ERROR: $_" -ForegroundColor Red }
finally {
  if ($wb)    { try { $wb.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null 2>$null
  Stop-Transcript | Out-Null
  Read-Host "`nPress Enter to close"
}
