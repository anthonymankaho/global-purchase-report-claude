param(
    [string]$ReportPath = "C:\Users\User\Desktop\Global Purchase Report Claude\REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_2026-06.xlsx",
    [string]$RatePath = "C:\Users\User\Desktop\Global Purchase Report Claude\Exchange Rate\STANDARD EXCHANGE RATE.xls",
    [string]$Month = "2026-06"
)
$ErrorActionPreference = "Stop"

function Release-Com($obj) { if ($obj) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } }
function Get-TrueLastDataRow($ws, [int]$col) { return $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $rwb = $excel.Workbooks.Open($RatePath, 0, $true)
    # xlCalculationManual - avoid triggering a full recalc of the Report's 40,000+ rows (Q column
    # references Exchange Rate; U:AD have full-column MAXIFS/MINIFS) just from writing a few rate values.
    # Needs at least one workbook open first, or setting this property itself throws.
    $excel.Calculation = -4135
    $usdRates = @{}
    try {
        $rws = $rwb.Sheets.Item("Exchange rate table")
        if ($rws.Cells.Item(12, 1).Text.Trim() -ne "CURRENCY" -or $rws.Cells.Item(12, 5).Text.Trim() -ne "HK DOLLAR") {
            throw "Header row (12) doesn't match expected layout."
        }
        $r = 13
        while ($rws.Cells.Item($r, 3).Text.Trim() -ne "") {
            $code = $rws.Cells.Item($r, 3).Text.Trim()
            $hkdRate = $rws.Cells.Item($r, 5).Value2
            if ($code -and $hkdRate -is [double]) { $usdRates[$code] = $hkdRate }
            $r++
        }
    } finally { $rwb.Close($false) | Out-Null; Release-Com $rwb }

    $usdHkd = $usdRates["USD"]
    foreach ($code in @($usdRates.Keys)) { $usdRates[$code] = $usdRates[$code] / $usdHkd }
    Write-Host "Computed USD rates: $($usdRates.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" } | Out-String)"

    $wb = $excel.Workbooks.Open($ReportPath, 0, $false)
    try {
        $ws = $wb.Sheets.Item("Exchange Rate")
        $headerRow = 3
        $colMap = @{}
        $c = 15
        while ($ws.Cells.Item($headerRow, $c).Text.Trim() -ne "") {
            $code = $ws.Cells.Item($headerRow, $c).Text.Trim()
            if ($code -ne "yyyy-mm") { $colMap[$code] = $c }
            $c++
        }

        $lastRow = Get-TrueLastDataRow $ws 15
        Write-Host "Exchange Rate sheet last row (col O): $lastRow"
        $targetRow = $null
        for ($r = 4; $r -le $lastRow; $r++) {
            if ($ws.Cells.Item($r, 15).Text.Trim() -eq $Month) { $targetRow = $r; break }
        }
        if (-not $targetRow) { $targetRow = $lastRow + 1; $ws.Cells.Item($targetRow, 15).Value2 = $Month }
        Write-Host "Target row: $targetRow"

        $filled = 0; $skipped = @()
        foreach ($code in $colMap.Keys) {
            if ($usdRates.ContainsKey($code)) {
                $ws.Cells.Item($targetRow, $colMap[$code]).Value2 = $usdRates[$code]
                $filled++
            } else { $skipped += $code }
        }
        Write-Host "Filled $filled columns. Skipped (no source rate): $($skipped -join ', ')"

        $wb.Save() | Out-Null
        Write-Host "Saved: $ReportPath" -ForegroundColor Green
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }
}
catch { Write-Host "ERROR: $_" -ForegroundColor Red; exit 1 }
finally { $excel.Quit(); Release-Com $excel; [GC]::Collect() }
