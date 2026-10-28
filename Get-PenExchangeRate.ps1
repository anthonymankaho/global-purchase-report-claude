<#
.SYNOPSIS
  Fetches the month-end USD/PEN exchange rate from the Pacific Exchange Rate
  Service (fx.sauder.ubc.ca) and writes it into the monthly Report's own
  "Exchange Rate" tab, filling the PEN gap that `STANDARD EXCHANGE RATE.xls`
  doesn't cover (needed for MAZ). See SOP 5 "MAZ 的 PEN 匯率缺口".

.DESCRIPTION
  Step 1 - fetch: submits the same request as the site's own web form (Base
  Currency: U.S. Dollars, Target Currency: Peruvian New Soles, Exact Time
  Period, daily frequency, Price Notation, CSV output) via a direct HTTP
  POST - no browser automation, since this is a plain public data-retrieval
  form with no login required. Price Notation gives "USD price of 1 PEN"
  (e.g. 0.2945).

  Step 2 - write: opens the target month's Report file's "Exchange Rate"
  sheet and writes the rate into column M (PEN), processed the SAME way
  Build-GlobalPurchaseReport.ps1's own Update-ExchangeRateSheet function
  processes every other currency there:
    - Block 1 (columns C:M) holds "Foreign Currency => HKD" rates, read
      straight from STANDARD EXCHANGE RATE.xls for every other currency.
    - Block 2 (columns O:Z) is a LIVE FORMULA block that divides each
      Block 1 rate by that row's own USD=>HKD rate (column K) to get
      "Foreign Currency => USD" - never written as a static value (an
      earlier bug in Update-ExchangeRateSheet did that once and clobbered
      the live formulas; same lesson applies here, so this script only
      ever rewrites the FORMULA in Block 2, never a literal number).
  fx.sauder.ubc.ca only offers a USD-denominated PEN rate (no HKD
  cross-rate), so it's converted to the sheet's own HKD basis before
  writing:
      HKD-per-PEN = (USD-per-PEN) x (that row's own HKD-per-USD, column K)
  so column M matches every other currency's convention and Block 2's
  existing PEN=>USD formula (column Z) recomputes correctly on its own.

  If the service cannot be reached, a verified rate can be supplied with
  -UsdPerPen, -RateDate and -RateSource. The rate date must be within the
  report month's lookback window; the script validates these inputs before
  opening or changing the workbook.

.PARAMETER ReportMonth
  The reporting period in "YYYY-MM" format (e.g. "2026-07"). Also the row
  this looks up/writes in the Report's Exchange Rate sheet (column B).

.PARAMETER WorkingFolder
  Folder containing the REPORT subfolder. Defaults to this script's folder.

.PARAMETER ReportPath
  Full path to the target Report file. Defaults to
  WorkingFolder\REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_<ReportMonth>.xlsx
  (same naming convention as Build-GlobalPurchaseReport.ps1's own -OutputPath).

.PARAMETER LookbackDays
  How many trailing days to query before giving up looking for a trading day
  with data. Default 7 (covers any combination of weekend + one holiday).

.PARAMETER UsdPerPen
  Optional offline rate input, in USD per 1 PEN. Use only when the online
  service is unavailable and the value has been verified from a trusted source.

.PARAMETER RateDate
  Required with -UsdPerPen. Effective date in YYYY-MM-DD format; it must be
  within the report month's configured lookback window.

.PARAMETER RateSource
  Required with -UsdPerPen. Short description or URL identifying the source.

.EXAMPLE
  .\Get-PenExchangeRate.ps1 -ReportMonth "2026-07"
  Fetches USD/PEN as of 2026-07-31 (or the last trading day at/before it),
  converts it to HKD-per-PEN using that row's own USD rate, and writes it
  into column M of the 2026-07 row in
  REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_2026-07.xlsx.

.EXAMPLE
  .\Get-PenExchangeRate.ps1 -ReportMonth "2026-09" -UsdPerPen 0.2945 -RateDate "2026-09-30" -RateSource "Verified official rate source"
  Uses a verified USD-per-PEN quote when the online service cannot be reached.
#>
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^\d{4}-\d{2}$')]
  [string]$ReportMonth,

  [string]$WorkingFolder = $PSScriptRoot,
  [string]$ReportPath,
  [int]$LookbackDays = 7,
  [double]$UsdPerPen = 0,
  [string]$RateDate,
  [string]$RateSource
)

$ErrorActionPreference = "Stop"

if ($PSBoundParameters.ContainsKey('UsdPerPen') -and $UsdPerPen -le 0) {
    throw "-UsdPerPen must be a positive USD-per-1-PEN rate. No workbook changes have been made."
}

if (-not $ReportPath) {
    $ReportPath = Join-Path $WorkingFolder "REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_$ReportMonth.xlsx"
}
if (-not (Test-Path $ReportPath)) {
    throw "Report file not found: $ReportPath (run Build-GlobalPurchaseReport.ps1 -ReportMonth `"$ReportMonth`" first, or pass -ReportPath explicitly)."
}

# --- Step 1: fetch or accept a verified USD/PEN rate ---
$parts = $ReportMonth -split '-'
$year = [int]$parts[0]
$month = [int]$parts[1]
$monthEnd = Get-Date -Year $year -Month $month -Day ([DateTime]::DaysInMonth($year, $month))
$windowStart = $monthEnd.AddDays(-$LookbackDays)
$rateSourceLabel = "Pacific Exchange Rate Service (fx.sauder.ubc.ca)"

if ($UsdPerPen -gt 0) {
  if (-not $RateDate -or -not $RateSource) {
    throw "Offline rate mode requires -UsdPerPen, -RateDate (YYYY-MM-DD), and -RateSource. No workbook changes have been made."
  }
  try {
    $parsedRateDate = [DateTime]::ParseExact($RateDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
  }
  catch {
    throw "Invalid -RateDate '$RateDate'. Use YYYY-MM-DD, for example 2026-09-30. No workbook changes have been made."
  }
  if ($parsedRateDate.Date -lt $windowStart.Date -or $parsedRateDate.Date -gt $monthEnd.Date) {
    throw "Rate date $RateDate is outside the allowed report window ($($windowStart.ToString('yyyy-MM-dd')) to $($monthEnd.ToString('yyyy-MM-dd'))). No workbook changes have been made."
  }
  $rateDate = $parsedRateDate.ToString('yyyy/MM/dd')
  $rateSourceLabel = $RateSource
  Write-Host "Using manually supplied, verified USD per 1 PEN rate ($rateSourceLabel)." -ForegroundColor Yellow
}
else {
  if ($RateDate -or $RateSource) {
    throw "-RateDate and -RateSource can only be used together with a positive -UsdPerPen value. No workbook changes have been made."
  }
  $body = @{
    b  = "USD"
    c  = "PEN"
    rd = ""
    fd = $windowStart.Day
    fm = $windowStart.Month
    fy = $windowStart.Year
    ld = $monthEnd.Day
    lm = $monthEnd.Month
    ly = $monthEnd.Year
    y  = "daily"
    q  = "price"   # Price Notation: USD price of 1 PEN, not the default Volume Notation
    f  = "csv"
    o  = ""
  }

  Write-Host "Querying Pacific Exchange Rate Service: USD/PEN, $($windowStart.ToString('yyyy-MM-dd')) to $($monthEnd.ToString('yyyy-MM-dd'))..."
  try {
    $response = Invoke-WebRequest -Uri "https://fx.sauder.ubc.ca/cgi/fxdata" -Method Post -Body $body -UseBasicParsing
  }
  catch {
    $detail = $_.Exception.Message
    $webStatus = $_.Exception.Status
    if ($detail -match '(?i)remote name|name resolution|DNS|could not be resolved|無法解析遠端名稱' -or [string]$webStatus -eq 'NameResolutionFailure') {
      throw "Could not resolve fx.sauder.ubc.ca (DNS lookup failed). Check this PC's internet/DNS connection or VPN, then retry. No workbook changes have been made. If the site remains unreachable, verify the month-end USD per 1 PEN rate from a trusted source and rerun with -UsdPerPen <rate> -RateDate YYYY-MM-DD -RateSource <source>. Error: $detail"
    }
    throw "Could not retrieve USD/PEN from fx.sauder.ubc.ca. Check the network and retry. No workbook changes have been made. If the service remains unavailable, use a verified rate with -UsdPerPen <rate> -RateDate YYYY-MM-DD -RateSource <source>. Error: $detail"
  }

  # Data rows look like: 2461253,"2026/07/31","Fri",0.29453
  $dataLines = $response.Content -split "`r?`n" | Where-Object { $_ -match '^\d+,"[\d/]+","\w+",[\d.]+' }
  if (-not $dataLines) {
    throw "No USD/PEN data returned for $ReportMonth (queried $($windowStart.ToString('yyyy-MM-dd')) to $($monthEnd.ToString('yyyy-MM-dd'))). No workbook changes have been made."
  }

  $lastLine = $dataLines[-1]
  $fields = $lastLine -split ','
  $rateDate = $fields[1].Trim('"')
  $usdPerPen = [double]$fields[3]

  if ($rateDate -ne $monthEnd.ToString('yyyy/MM/dd')) {
    Write-Host "Note: $($monthEnd.ToString('yyyy-MM-dd')) had no trading data; using the most recent available date instead." -ForegroundColor Yellow
  }
}
Write-Host "USD/PEN as of $rateDate (USD per 1 PEN): $usdPerPen; source: $rateSourceLabel" -ForegroundColor Green

# --- Step 2: write into the Report's Exchange Rate sheet, Block 1 (HKD basis) ---
function Release-Com($obj) { if ($obj) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } }
function Col-Letter([int]$colIndex) {
    $letter = ""
    $n = $colIndex
    while ($n -gt 0) { $rem = ($n - 1) % 26; $letter = [char](65 + $rem) + $letter; $n = [int](($n - $rem - 1) / 26) }
    return $letter
}

$excel = $null
$wb = $null
try {
    $existingExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    if ($existingExcel) {
        Write-Host "Warning: EXCEL.EXE is already running ($($existingExcel.Count) process(es)). Please close all Excel windows for a clean write. Waiting 5s, then continuing anyway..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($ReportPath, 0, $false)
    if ($null -eq $wb) { throw "Workbook failed to open (Workbooks.Open returned null)." }
    if ($wb.ReadOnly) { throw "Opened read-only unexpectedly - check if '$ReportPath' is open elsewhere (e.g. in your own Excel)." }

    $ws = $wb.Sheets.Item("Exchange Rate")

    # Dynamically find PEN's and USD's Block 1 columns from the header row (3), same lookup
    # Update-ExchangeRateSheet uses - don't hardcode column letters in case the layout shifts.
    $colMap = @{}
    $c = 3
    while ($ws.Cells.Item(3, $c).Text.Trim() -ne "") {
        $colMap[$ws.Cells.Item(3, $c).Text.Trim()] = $c
        $c++
    }
    if (-not $colMap.ContainsKey("PEN")) { throw "Column header 'PEN' not found in row 3 of the 'Exchange Rate' sheet - layout may have changed." }
    if (-not $colMap.ContainsKey("USD")) { throw "Column header 'USD' not found in row 3 of the 'Exchange Rate' sheet - layout may have changed." }
    $penCol = $colMap["PEN"]
    $usdCol = $colMap["USD"]
    $penColLetter = Col-Letter $penCol
    $usdColLetter = Col-Letter $usdCol

    # Find the row for $ReportMonth in column B - identical lookup to Update-ExchangeRateSheet,
    # so this lands on exactly the row the Build script already filled Block 1 for every other
    # currency on (append after the last populated row only if genuinely not found).
    $lastRow = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row   # xlUp
    $targetRow = $null
    for ($r = 4; $r -le $lastRow; $r++) {
        if ($ws.Cells.Item($r, 2).Text.Trim() -eq $ReportMonth) { $targetRow = $r; break }
    }
    if (-not $targetRow) { $targetRow = $lastRow + 1; $ws.Cells.Item($targetRow, 2).Value2 = $ReportMonth }

    $hkdPerUsdRaw = $ws.Cells.Item($targetRow, $usdCol).Value2
    if ($hkdPerUsdRaw -isnot [double]) {
        throw "Row $targetRow (month $ReportMonth) has no numeric USD=>HKD rate in column $usdColLetter yet - run Build-GlobalPurchaseReport.ps1 for this month first (it fills Block 1 from STANDARD EXCHANGE RATE.xls), then re-run this script."
    }

    $hkdPerPen = [Math]::Round($usdPerPen * $hkdPerUsdRaw, 4)
    $ws.Cells.Item($targetRow, $penCol).Value2 = $hkdPerPen

    # Block 2's PEN=>USD cell sits 13 columns after its Block 1 counterpart (matches every other
    # currency's O:Z <- C:M offset in Update-ExchangeRateSheet) - always (re)write the FORMULA
    # rather than assume it's already there (same defensive practice as that function), and
    # NEVER write a static number here - that's the exact bug Update-ExchangeRateSheet's own
    # comment warns about (it clobbers the live formula for good).
    $block2Col = $penCol + 13
    $block2ColLetter = Col-Letter $block2Col
    $ws.Cells.Item($targetRow, $block2Col).Formula = "=IF($penColLetter$targetRow=`"`",`"`",ROUND(+$penColLetter$targetRow/$usdColLetter$targetRow,4))"

    # COM Value2/Formula writes don't reliably mark dependent cells dirty for recalculation
    # (same lesson as Process-MHK's Trimmed sheet) - force it before reading back / saving.
    $ws.Calculate()

    $writtenBackUsdPerPen = $ws.Cells.Item($targetRow, $block2Col).Value2

    $beforeSaveTime = Get-Date
    $wb.Save()
    Start-Sleep -Milliseconds 500
    $onDiskTime = (Get-Item $ReportPath).LastWriteTime
    if ($onDiskTime -lt $beforeSaveTime) { throw "Save() completed but the file's on-disk timestamp didn't update - it may not have actually saved." }

    Write-Host ""
    Write-Host "=== Written to Report ===" -ForegroundColor Cyan
    Write-Host "File          : $ReportPath"
    Write-Host "Sheet/Row     : 'Exchange Rate' row $targetRow (month $ReportMonth)"
    Write-Host "Column $penColLetter (PEN, HKD basis) <- $hkdPerPen   (= $usdPerPen USD/PEN x $hkdPerUsdRaw HKD/USD)"
    Write-Host "Rate source   : $rateSourceLabel"
    Write-Host "Column $block2ColLetter (PEN=>USD formula) recomputed to: $writtenBackUsdPerPen   (sanity check - should be close to the fetched $usdPerPen)"
    Write-Host "Saved successfully (confirmed via on-disk file timestamp)." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null; Release-Com $wb }
    if ($excel) { $excel.Quit(); Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    $stillRunning = $null
    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 1
        $stillRunning = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
        if (-not $stillRunning) { break }
    }
    if ($stillRunning) {
        Write-Host "Warning: EXCEL.EXE still running after this script ($($stillRunning.Count) process(es))." -ForegroundColor Yellow
        Write-Host "If this persists, close them manually via Task Manager before running this script again." -ForegroundColor Yellow
    }
}
