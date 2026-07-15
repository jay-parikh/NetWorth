<#  UpdatePrices.ps1  --------------------------------------------------------
  Refreshes Closing Price Date / Closing Price / Prev. close in an Excel
  workbook from the NSE or BSE bhavcopy, matched by ISIN. No Python needed.

  DO NOT double-click this .ps1 file. Double-click  UpdatePrices.bat  instead.
  Everything is logged to  UpdatePrices_log.txt  and the window always waits.
--------------------------------------------------------------------------- #>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
$log = Join-Path $here "UpdatePrices_log.txt"
try { Start-Transcript -Path $log -Force | Out-Null } catch {}

function Pause-Exit([int]$code){
  try { Stop-Transcript | Out-Null } catch {}
  Write-Host ""; Read-Host "Press Enter to close"; exit $code
}

try {
  $ErrorActionPreference = "Stop"

function Get-Xirr($flows) {
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

function Update-EquityXirr($wb) {
  $eq=$null
  foreach ($ws in $wb.Worksheets) { if ($ws.Name -eq "Equity") { $eq=$ws } }
  if (-not $eq) { return }
  $today=(Get-Date).Date
  $flows=New-Object System.Collections.ArrayList
  for ($r=4; $r -le 140; $r++) {
    $dv=$eq.Cells.Item($r,13).Value2      # Cost date (M)
    $inv=$eq.Cells.Item($r,10).Value2     # Invested (J)
    $cv=$eq.Cells.Item($r,9).Value2       # Cur. val (I)
    if ($dv -and ($inv -is [double]) -and $inv -gt 0 -and ($cv -is [double])) {
      $d=[datetime]::FromOADate([double]$dv)
      [void]$flows.Add(@{d=$d; a=-[double]$inv})
      [void]$flows.Add(@{d=$today; a=[double]$cv})
    }
  }
  $x=Get-Xirr $flows
  # TOTAL row: find row in col C = TOTAL (search 100..200)
  for ($r=100; $r -le 200; $r++) {
    if (("" + $eq.Cells.Item($r,3).Value2).Trim() -eq "TOTAL") {
      if ($x -ne $null) { $eq.Cells.Item($r,14).Value2=[double]$x
        Write-Host ("  Equity portfolio XIRR: {0:P2}" -f $x) }
      break
    }
  }
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

function Update-StockMaster($wb, $info, $dateStr) {
  # Merge newly listed stocks from the bhavcopy into Stock_Master. Add-only:
  # existing ISINs keep their name, so Equity dropdown picks never break.
  if (-not $info -or $info.Count -eq 0) { return }
  $ws=$null
  foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "Stock_Master") { $ws=$sh } }
  if (-not $ws) { return }
  $last=$ws.UsedRange.Rows.Count
  $have=@{}   # isin -> @(symbol, name)
  $names=@{}
  if ($last -ge 4) {
    $arr=$ws.Range("A4:C"+$last).Value2
    for ($i=1; $i -le ($last-3); $i++) {
      $isin=("" + $arr[$i,3]).Trim()
      if (-not $isin) { continue }
      $nm=("" + $arr[$i,2]).Trim()
      $have[$isin]=@(("" + $arr[$i,1]).Trim(), $nm)
      $names[$nm]=1
    }
  }
  $new=0
  foreach ($isin in @($info.Keys)) {
    if ($have.ContainsKey($isin)) { continue }
    $nm=$info[$isin].Name; $sym=$info[$isin].Sym
    if (-not $nm) { continue }
    if ($names.ContainsKey($nm) -and $sym) { $nm="$nm ($sym)" }
    if ($names.ContainsKey($nm)) { continue }
    $have[$isin]=@($sym, $nm); $names[$nm]=1; $new++
  }
  $ws.Range("E2").Value2=[string]$dateStr
  if ($new -eq 0) {
    Write-Host ("  Stock master: {0} stocks (no new listings)" -f $have.Count)
    return
  }
  # ordinal sort by name - keeps prefix blocks contiguous for the dropdown filter
  $n=$have.Count
  $keys=New-Object string[] $n
  $vals=New-Object object[] $n
  $i=0
  foreach ($k in @($have.Keys)) {
    $keys[$i]=("" + $have[$k][1]); $vals[$i]=@($have[$k][0], $have[$k][1], $k); $i++
  }
  [Array]::Sort($keys, $vals, [System.StringComparer]::OrdinalIgnoreCase)
  $out=New-Object 'object[,]' $n, 3
  for ($i=0; $i -lt $n; $i++) { $out[$i,0]=$vals[$i][0]; $out[$i,1]=$vals[$i][1]; $out[$i,2]=$vals[$i][2] }
  if ($last -ge 4) { $ws.Range("A4:C"+$last).ClearContents() | Out-Null }
  $ws.Range("A4:C"+(3+$n)).Value2=$out
  Write-Host ("  Stock master: {0} stocks ({1} new)" -f $n, $new)
}


  $ISIN_KEYS  = @("isin","isin_code","isin no","isin no.","isin code")
  $CLOSE_KEYS = @("clspric","close","close price","closing price")
  $PREV_KEYS  = @("prvsclsgpric","prevclose","prev close","prev.","prev. close","previous close")
  $DATE_KEYS  = @("traddt","trading_date","closing price date")

  function Find-Col([string[]]$headers,[string[]]$wanted,[switch]$Close){
    for($i=0;$i -lt $headers.Count;$i++){ $h=(""+$headers[$i]).Trim().ToLower()
      if($Close -and $h -like "*date*"){ continue }
      foreach($w in $wanted){ if($h -eq $w){ return $i } } }
    return -1
  }
  function Get-ColName([string[]]$cols,[string[]]$wanted,[switch]$Close){
    foreach($c in $cols){ $l=$c.Trim().ToLower()
      if($Close -and $l -like "*date*"){ continue }
      foreach($w in $wanted){ if($l -eq $w){ return $c } } }
    foreach($c in $cols){ $l=$c.Trim().ToLower()
      if($Close -and $l -like "*date*"){ continue }
      if($Close -and ($l -like "*clspric*" -or $l -like "*clos*")){ return $c }
      if(-not $Close -and ($l -like "*prvsclsg*" -or $l -like "*prevclos*" -or ($l -like "*prev*" -and $l -like "*clos*"))){ return $c } }
    return $null
  }

  $in = Read-Host "Valuation date DD-MM-YYYY (blank = latest trading day)"

  function Get-Bhavcopy([datetime]$d){
    $ymd=$d.ToString("yyyyMMdd"); $dmy=$d.ToString("dd-MM-yyyy")
    foreach($f in @("bhavcopy_$dmy.csv","bhavcopy_$dmy.CSV")){
      if(Test-Path $f){ Write-Host "  using local $f"; return (Resolve-Path $f).Path } }
    $tmp=Join-Path $env:TEMP "bhav_$ymd.csv"
    try{ $u="https://www.bseindia.com/download/BhavCopy/Equity/BhavCopy_BSE_CM_0_0_0_${ymd}_F_0000.CSV"
      Invoke-WebRequest -Uri $u -OutFile $tmp -UserAgent "Mozilla/5.0" -Headers @{Referer="https://www.bseindia.com/"} -TimeoutSec 30
      if((Get-Item $tmp).Length -gt 1000){ Write-Host "  BSE ok ($dmy)"; return $tmp } }catch{ Write-Host "  BSE not available ($dmy)" }
    try{ $sess=$null
      Invoke-WebRequest -Uri "https://www.nseindia.com" -SessionVariable sess -UserAgent "Mozilla/5.0" -TimeoutSec 30 | Out-Null
      $zip=Join-Path $env:TEMP "bhav_$ymd.zip"
      $u="https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_${ymd}_F_0000.csv.zip"
      Invoke-WebRequest -Uri $u -OutFile $zip -WebSession $sess -UserAgent "Mozilla/5.0" -Headers @{Referer="https://www.nseindia.com/"} -TimeoutSec 30
      $dst=Join-Path $env:TEMP "bhav_$ymd"; Expand-Archive -Path $zip -DestinationPath $dst -Force
      $csv=Get-ChildItem $dst -Filter *.csv | Select-Object -First 1
      if($csv){ Write-Host "  NSE ok ($dmy)"; return $csv.FullName } }catch{ Write-Host "  NSE not available ($dmy)" }
    return $null
  }

  $csvPath=$null; $used=$null
  if($in.Trim() -ne ""){
    try{ $used=[datetime]::ParseExact($in.Trim(),"dd-MM-yyyy",$null) }
    catch{ Write-Host "Bad date format. Use DD-MM-YYYY." -ForegroundColor Yellow; Pause-Exit 1 }
    $csvPath=Get-Bhavcopy $used
  } else {
    for($k=0;$k -le 7;$k++){ $d=(Get-Date).AddDays(-$k)
      if($d.DayOfWeek -in "Saturday","Sunday"){ continue }
      Write-Host ("trying {0}..." -f $d.ToString("dd-MM-yyyy"))
      $csvPath=Get-Bhavcopy $d; if($csvPath){ $used=$d; break } }
  }
  if(-not $csvPath){
    Write-Host "`nCould not get a bhavcopy automatically." -ForegroundColor Yellow
    Write-Host "Download it from bseindia.com / nseindia.com, save here as bhavcopy_<DD-MM-YYYY>.csv, run again."
    Pause-Exit 1
  }
  $dateStr=$used.ToString("dd-MM-yyyy")

  $data = Import-Csv -LiteralPath $csvPath
  if(-not $data -or $data.Count -eq 0){ throw "bhavcopy file is empty: $csvPath" }
  $cols=$data[0].PSObject.Properties.Name
  $icN=Get-ColName $cols $ISIN_KEYS
  $ccN=Get-ColName $cols $CLOSE_KEYS -Close
  $pcN=Get-ColName $cols $PREV_KEYS
  $nmN=Get-ColName $cols @("fininstrmnm","security name","sc_name")
  $syN=Get-ColName $cols @("tckrsymb","symbol","sc_code")
  if(-not $icN -or -not $ccN){ throw ("ISIN/Close columns not found. Header: " + ($cols -join ",")) }
  Write-Host ("  columns -> ISIN='{0}'  Close='{1}'  Prev='{2}'" -f $icN,$ccN,$pcN)
  $prices=@{}
  $stockInfo=@{}
  foreach($row in $data){
    $isin=(""+$row.$icN).Trim().ToUpper(); if(-not $isin){ continue }
    if($nmN -and $isin.Length -eq 12 -and $isin.StartsWith("INE") -and -not $stockInfo.ContainsKey($isin)){
      $nm=(""+$row.$nmN).Trim()
      if($nm){
        $sy=""; if($syN){ $sy=(""+$row.$syN).Trim() }
        $stockInfo[$isin]=@{Name=$nm; Sym=$sy}
      }
    }
    $close=0.0; $prev=0.0
    [void][double]::TryParse((""+$row.$ccN),[ref]$close)
    if($pcN){ [void][double]::TryParse((""+$row.$pcN),[ref]$prev) }
    if($close -gt 0){ $prices[$isin]=@{Close=$close;Prev=$(if($prev -gt 0){$prev}else{$null})} }
  }
  Write-Host ("bhavcopy {0}: {1} priced ISINs" -f $dateStr,$prices.Count)
  if($prices.Count -eq 0){ throw "No priced rows parsed from bhavcopy." }

  $wbFile=Get-ChildItem -Filter "Family_Portfolio_Tracker.xlsx" | Select-Object -First 1
  if(-not $wbFile){ $wbFile=Get-ChildItem -Filter "*Tracker*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
  if(-not $wbFile){ throw "Family_Portfolio_Tracker.xlsx not found in this folder ($here)." }
  Write-Host ("workbook: {0}" -f $wbFile.Name)

  try{ $xl=New-Object -ComObject Excel.Application }
  catch{ throw "Could not start Excel. Desktop Microsoft Excel must be installed. ($_)" }
  $xl.Visible=$false; $xl.DisplayAlerts=$false
  $wb=$xl.Workbooks.Open($wbFile.FullName)
  $updated=0

  foreach($ws in $wb.Worksheets){
    $ur=$ws.UsedRange; $rows=$ur.Rows.Count; $colc=$ur.Columns.Count
    if($rows -lt 2){ continue }

    # header row = first of rows 1..5 containing a cell equal to "ISIN"
    $hr=-1
    for($r=1;$r -le [math]::Min(5,$rows);$r++){
      for($c=1;$c -le $colc;$c++){ $v=$ws.Cells.Item($r,$c).Value2
        if($v -and ($v.ToString().Trim().ToLower() -eq "isin")){ $hr=$r; break } }
      if($hr -gt 0){ break } }

    # update "valued <date>" title text (string assignment)
    for($r=1;$r -le [math]::Min(2,$rows);$r++){
      try{ $v=$ws.Cells.Item($r,1).Value2
        if($v -and $v.ToString() -match "valued\s+\d{2}-\d{2}-\d{4}"){
          $ws.Cells.Item($r,1).Value2 = [string]($v -replace "valued\s+\d{2}-\d{2}-\d{4}","valued $dateStr") } }catch{}
    }
    if($hr -lt 0){ continue }

    $hdrs=@(); for($c=1;$c -le $colc;$c++){ $hdrs += [string]$ws.Cells.Item($hr,$c).Value2 }
    $ci =(Find-Col $hdrs $ISIN_KEYS)+1
    $cCl=(Find-Col $hdrs $CLOSE_KEYS -Close)+1
    $cPv=(Find-Col $hdrs $PREV_KEYS)+1
    $cDt=(Find-Col $hdrs $DATE_KEYS)+1
    if($ci -le 0 -or $cCl -le 0){ continue }

    $f=$hr+1; $n=$rows-$f+1
    if($n -lt 1){ continue }

    # read whole columns as arrays
    $isinRng =$ws.Range($ws.Cells.Item($f,$ci),  $ws.Cells.Item($rows,$ci))
    $closeRng=$ws.Range($ws.Cells.Item($f,$cCl), $ws.Cells.Item($rows,$cCl))
    $isinArr =$isinRng.Value2
    $closeArr=$closeRng.Value2
    if($cPv -gt 0){ $prevRng=$ws.Range($ws.Cells.Item($f,$cPv),$ws.Cells.Item($rows,$cPv)); $prevArr=$prevRng.Value2 }
    if($cDt -gt 0){ $dateRng=$ws.Range($ws.Cells.Item($f,$cDt),$ws.Cells.Item($rows,$cDt)); $dateArr=$dateRng.Value2 }
    $before=$updated

    if($n -eq 1){
      $isin=(""+$isinArr).Trim()
      if($isin -and $prices.ContainsKey($isin)){
        $closeRng.Value2=[double]$prices[$isin].Close
        if($cPv -gt 0 -and $prices[$isin].Prev){ $prevRng.Value2=[double]$prices[$isin].Prev }
        if($cDt -gt 0){ $dateRng.Value2=[string]$dateStr }
        $updated++
      }
    } else {
      for($i=1;$i -le $n;$i++){
        $cell=$isinArr[$i,1]; if(-not $cell){ continue }
        $isin=(""+$cell).Trim(); if(-not $isin){ continue }
        if($prices.ContainsKey($isin)){
          $closeArr[$i,1]=[double]$prices[$isin].Close
          if($cPv -gt 0 -and $prices[$isin].Prev){ $prevArr[$i,1]=[double]$prices[$isin].Prev }
          if($cDt -gt 0){ $dateArr[$i,1]=[string]$dateStr }
          $updated++
        }
      }
      # write arrays back in one shot (robust for numbers)
      $closeRng.Value2=$closeArr
      if($cPv -gt 0){ $prevRng.Value2=$prevArr }
      if($cDt -gt 0){ $dateRng.Value2=$dateArr }
    }
    Write-Host ("  sheet '{0}': matched {1} (prev col {2}, date col {3})" -f $ws.Name,($updated-$before),$cPv,$cDt)
  }

  Update-StockMaster $wb $stockInfo $dateStr
  $xl.CalculateFull()
  Update-EquityXirr $wb
  Update-PortfolioXirr $wb
  $wb.Save(); $wb.Close($true); $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  Write-Host ("`nDONE. Cells updated: {0}. Valuation date: {1}" -f $updated,$dateStr) -ForegroundColor Green
  Pause-Exit 0
}
catch {
  Write-Host ""; Write-Host "ERROR:" -ForegroundColor Red
  Write-Host ($_ | Out-String) -ForegroundColor Red
  Write-Host "Details saved to UpdatePrices_log.txt" -ForegroundColor Yellow
  try{ if($wb){ $wb.Close($false) }; if($xl){ $xl.Quit() } }catch{}
  Pause-Exit 1
}
