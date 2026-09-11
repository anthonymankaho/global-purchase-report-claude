<#
.SYNOPSIS
    Second-pass formula fill for the Global Purchase Report - fills AC/AD/AE/AF/AG/AH/AI for the
    current month's new rows and builds the "New Vendors" review sheet. Run this AFTER manually
    copying AA2:AB2's formula down to the new rows in Excel.

.DESCRIPTION
    2026-08: per explicit instruction, Build-GlobalPurchaseReport.ps1 no longer fills AA/AB (Max/Min
    PO Price) at all - the report owner copies AA2:AB2's live formula down to the new rows manually in
    Excel and confirms it looks right before continuing. This script picks up from there: it fills the
    columns that depend on AA/AB (AC/AD/AE directly; AG/AH transitively via AC/AD/AE; AF/AI are grouped
    in here too since they were always filled in the same "second pass" as these), one column at a time
    with an immediate freeze-to-values after each (same technique as Build-GlobalPurchaseReport.ps1's
    Set-FormulaColumnsFrozen, and for the same reason - filling every column at once and calculating
    would recalculate the full dependency web across every new row simultaneously, which is what caused
    Excel to hang for 10+ minutes earlier in this project).

    Row range: found by matching column B (PO_REC_DATE, always populated - unlike the formula columns
    this script fills, which don't exist yet for new rows) against -ReportMonth, same technique as
    New-ReportAuditSummary.ps1. Before touching anything, sanity-checks that AA at the first matched row
    already has a formula (i.e. the manual step was actually done) - aborts with a clear message if not,
    rather than silently filling AC/AD/AE from blank AA/AB (which would just reproduce the stale-key
    #N/A bug this whole workflow change was meant to avoid).

    Dependency order for the columns this script fills:
        AC  <- needs Y (already filled) and R (already filled) and AA (manual step, done by now)
        AF  <- needs Y and U (already filled)
        AD  <- needs Y and AA (manual step)
        AE  <- needs Y and AB (manual step)
        AG  <- needs Y, AD, AC, AF
        AH  <- needs Y, AE, AC, AF
        AI  <- raw H only, independent, can go anywhere

.PARAMETER ReportMonth
    Month to process, "yyyy-MM" format, e.g. "2026-06". Used to find the row range via column B.

.PARAMETER ReportPath
    Path to the report. If omitted, auto-detects the latest "FIN_*_PUR_REPORT_*.xlsx" in
    WorkingFolder\REPORT (same file Build-GlobalPurchaseReport.ps1 would use/produce).

.EXAMPLE
    .\Complete-FormulaColumns.ps1 -ReportMonth "2026-06"
#>

param(
    [string]$WorkingFolder = $PSScriptRoot,
    [string]$ReportPath,
    [Parameter(Mandatory = $true)]
    [string]$ReportMonth
)

$ErrorActionPreference = "Stop"

if (-not $ReportPath) {
    $priorOutputs = Get-ChildItem (Join-Path $WorkingFolder "REPORT") -Filter "FIN_*_PUR_REPORT_*.xlsx" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending
    if (-not $priorOutputs) { throw "No 'FIN_*_PUR_REPORT_*.xlsx' file found in WorkingFolder\REPORT - pass -ReportPath explicitly." }
    $ReportPath = $priorOutputs[0].FullName
    Write-Host "Auto-detected ReportPath (latest monthly output): $ReportPath"
}

function Release-Com($obj) { if ($obj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } }
function Get-TrueLastDataRow($ws, [int]$col) { return $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }

function Get-RangeArray($ws, [int]$r1, [int]$c1, [int]$r2, [int]$c2) {
    $rng = $ws.Range($ws.Cells.Item($r1, $c1), $ws.Cells.Item($r2, $c2))
    $arr = $rng.Value2
    Release-Com $rng
    if ($r2 -eq $r1 -and $c2 -eq $c1) {
        $tmp = New-Object 'object[,]' 1, 1
        $tmp[0, 0] = $arr
        return , $tmp
    }
    return , $arr
}

function Set-RangeArray($ws, [int]$r1, [int]$c1, $arr) {
    $rows = $arr.GetLength(0)
    $cols = $arr.GetLength(1)
    $rng = $ws.Range($ws.Cells.Item($r1, $c1), $ws.Cells.Item($r1 + $rows - 1, $c1 + $cols - 1))
    $rng.Value2 = $arr
    Release-Com $rng
}

function Get-RowMonth($v, $epoch) {
    if ($v -is [double]) { return $epoch.AddDays($v).ToString("yyyy-MM") }
    if ($v -is [datetime]) { return $v.ToString("yyyy-MM") }
    if ($v -is [string] -and $v -ne "") { try { return ([datetime]$v).ToString("yyyy-MM") } catch { return $null } }
    return $null
}

# Same technique as Build-GlobalPurchaseReport.ps1's Set-FormulaColumnsFrozen: copy the row-2 master
# formula down over [startRow:endRow], force a recalc, then immediately freeze that column's result to
# static values before moving to the next column.
function Set-FormulaColumnsFrozen($ws, [int]$startRow, [int]$endRow, [string[]]$columns) {
    if ($endRow -lt $startRow) { return }
    $xlPasteFormulas = -4123
    $xlPasteValues = -4163
    foreach ($col in $columns) {
        $target = "$col$startRow`:$col$endRow"
        $ws.Range("$col`2").Copy() | Out-Null
        $ws.Range($target).PasteSpecial($xlPasteFormulas) | Out-Null
        $ws.Calculate() | Out-Null
        $ws.Range($target).Copy() | Out-Null
        $ws.Range($target).PasteSpecial($xlPasteValues) | Out-Null
        Write-Host "  [$col] filled, calculated, frozen"
    }
    $ws.Application.CutCopyMode = $false
}

# Scans [startRow:endRow] in the Data sheet for column U ("Std Vendor Name") = "NEW VENDOR" and copies
# any such rows in full to a "New Vendors" sheet (created if missing, cleared and repopulated if it
# already exists).
function Copy-NewVendorRowsToReviewSheet($wb, $dataWs, [int]$startRow, [int]$endRow) {
    if ($endRow -lt $startRow) { return }
    $lastCol = $dataWs.Cells.Item(1, $dataWs.Columns.Count).End(-4159).Column
    $block = Get-RangeArray $dataWs $startRow 1 $endRow $lastCol
    $n = $block.GetLength(0)
    $matchRows = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $n; $i++) {
        if ([string]$block[$i, 20] -eq "NEW VENDOR") { [void]$matchRows.Add($i) }   # U = column 21, array idx 20
    }

    if ($matchRows.Count -eq 0) {
        Write-Host "[NewVendorCheck] No 'NEW VENDOR' rows found in rows $startRow-$endRow."
        return
    }

    $sheetName = "New Vendors"
    $reviewWs = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq $sheetName) { $reviewWs = $s } }
    if ($reviewWs) {
        $reviewWs.Cells.Clear() | Out-Null
    } else {
        $reviewWs = $wb.Sheets.Add()
        $reviewWs.Name = $sheetName
    }

    $header = Get-RangeArray $dataWs 1 1 1 $lastCol
    Set-RangeArray $reviewWs 1 1 $header

    $result = New-Object 'object[,]' $matchRows.Count, $lastCol
    for ($i = 0; $i -lt $matchRows.Count; $i++) {
        for ($c = 0; $c -lt $lastCol; $c++) { $result[$i, $c] = $block[$matchRows[$i], $c] }
    }
    Set-RangeArray $reviewWs 2 1 $result
    Write-Host "[NewVendorCheck] Found $($matchRows.Count) 'NEW VENDOR' row(s) in rows $startRow-$endRow - copied to '$sheetName' sheet for review." -ForegroundColor Yellow
}

$existingExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($existingExcel) {
    Write-Host "Found $($existingExcel.Count) leftover EXCEL.EXE process(es) - closing them first to guarantee a clean COM session." -ForegroundColor Yellow
    Write-Host "(If you have unsaved work open in Excel, save it now - Ctrl+C to abort within 5 seconds.)" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    $existingExcel | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($ReportPath, [Type]::Missing, $false)
    try {
        $data = $wb.Sheets.Item("MHK+STC+MZQ+MIL+RACO Data")
        $trueLastRow = Get-TrueLastDataRow $data 1

        # Find the target month's row range via column B (PO_REC_DATE) - always populated, unlike the
        # formula columns this script is about to fill, which don't exist yet for brand-new rows.
        $range = $data.Range($data.Cells.Item(2, 2), $data.Cells.Item($trueLastRow, 2))
        $vals = $range.Value2
        Release-Com $range
        $n = $trueLastRow - 1
        $epoch = [datetime]'1899-12-30'

        $startRow = 0
        $endRow = 0
        for ($i = 1; $i -le $n; $i++) {
            $month = Get-RowMonth $vals[$i, 1] $epoch
            if ($month -eq $ReportMonth) {
                if ($startRow -eq 0) { $startRow = $i + 1 }
                $endRow = $i + 1
            }
        }

        if ($startRow -eq 0) {
            throw "No rows found matching ReportMonth '$ReportMonth' in column B (PO_REC_DATE). Nothing to fill."
        }
        Write-Host "Rows for $ReportMonth`: $startRow to $endRow"

        # Sanity check: AA at the first matched row must already have a formula (i.e. the manual
        # AA2:AB2-copy-down step was actually done) - otherwise AC/AD/AE would be built from blank
        # AA/AB, reproducing the exact stale-key #N/A bug this workflow change was meant to avoid.
        $aaFormula = $data.Range("AA$startRow").Formula
        if ([string]::IsNullOrEmpty($aaFormula)) {
            throw "AA$startRow has no formula yet. Copy AA2:AB2 down to rows $startRow-$endRow in Excel, calculate, and confirm it looks right BEFORE running this script."
        }
        Write-Host "AA$startRow formula check passed - manual AA/AB fill confirmed."

        Copy-NewVendorRowsToReviewSheet $wb $data $startRow $endRow
        Set-FormulaColumnsFrozen $data $startRow $endRow @("AC", "AF", "AD", "AE", "AG", "AH", "AI")

        $wb.Save()
        Write-Host "Saved: $ReportPath" -ForegroundColor Green
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($excel) { $excel.Quit(); Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "=== Done. Please open the report, run Data > Refresh All, and spot-check before sending. ===" -ForegroundColor Cyan
