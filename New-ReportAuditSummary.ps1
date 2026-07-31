<#
.SYNOPSIS
    Drafts a summary audit report for a Global Purchase Report: last data row, each entity's row
    subtotal for the target month, and the overall subtotal for that month.

.DESCRIPTION
    Opens the report read-only (never modifies it) and reports:
      1) Last report data row  - the true last row of the Data sheet (col A), ignoring any trailing
         blank/padding rows that UsedRange would otherwise include.
      2) Seven data subtotal   - row count per entity (MHK/STC/MZQ/MIL/RACO/TVS/MAZ, from the DB column)
         for the target month (from the month-tag column), so you can cross-check against what
         Build-GlobalPurchaseReport.ps1 printed when it ran.
      3) Current report subtotal - the sum of the seven entities' counts above, i.e. total rows added
         for that month.

.PARAMETER ReportPath
    Path to the report to audit. If omitted, auto-detects the latest "FIN_*_PUR_REPORT_*.xlsx" in
    WorkingFolder\REPORT (same file Build-GlobalPurchaseReport.ps1 would use/produce).

.PARAMETER ReportMonth
    Month to subtotal, format "yyyy-MM". If omitted, auto-detects the most recent month tag present
    in the Data sheet's month column (col X).

.PARAMETER OutputPath
    Where to write the summary file. Defaults to WorkingFolder\REPORT\AuditSummary_<month>.md

.EXAMPLE
    .\New-ReportAuditSummary.ps1
    .\New-ReportAuditSummary.ps1 -ReportMonth "2026-06"
#>

param(
    [string]$WorkingFolder = $PSScriptRoot,
    [string]$ReportPath,
    [string]$ReportMonth,
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (-not $ReportPath) {
    $priorOutputs = Get-ChildItem (Join-Path $WorkingFolder "REPORT") -Filter "FIN_*_PUR_REPORT_*.xlsx" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending
    if (-not $priorOutputs) { throw "No 'FIN_*_PUR_REPORT_*.xlsx' file found in WorkingFolder\REPORT - pass -ReportPath explicitly." }
    $ReportPath = $priorOutputs[0].FullName
    Write-Host "Auto-detected ReportPath (latest monthly output): $ReportPath"
}

function Release-Com($obj) { if ($obj) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } }
function Get-TrueLastDataRow($ws, [int]$col) { return $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row }

$existingExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($existingExcel) {
    Write-Host "Warning: EXCEL.EXE already running ($($existingExcel.Count)). Close all Excel windows first." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($ReportPath, 0, $true)   # read-only - this script never modifies the report
    try {
        $data = $wb.Sheets.Item("MHK+STC+MZQ+MIL+RACO Data")
        $trueLastRow = Get-TrueLastDataRow $data 1
        Write-Host "Last report data row: $trueLastRow"

        # Bulk-read column B (PO_REC_DATE, raw date - always populated) and column T (DB, entity marker)
        # in one call - avoids the per-cell-loop performance trap this project hit earlier. Deliberately
        # NOT using column X (month tag): that's computed by the Q2:AJ2 formulas, which may not have
        # been manually copied down to the newest rows yet (see SOP) - column B works either way.
        $range = $data.Range($data.Cells.Item(2, 2), $data.Cells.Item($trueLastRow, 20))
        $vals = $range.Value2
        $n = $trueLastRow - 1
        $epoch = [datetime]'1899-12-30'

        function Get-RowMonth($v) {
            if ($v -is [double]) { return $epoch.AddDays($v).ToString("yyyy-MM") }
            if ($v -is [datetime]) { return $v.ToString("yyyy-MM") }
            if ($v -is [string] -and $v -ne "") { try { return ([datetime]$v).ToString("yyyy-MM") } catch { return $null } }
            return $null
        }

        if (-not $ReportMonth) {
            # Most recent month present (data is appended chronologically, so the last row's date wins)
            $ReportMonth = Get-RowMonth $vals[$n, 1]
            Write-Host "Auto-detected ReportMonth (from last row's PO_REC_DATE): $ReportMonth"
        }

        $entities = @("MHK", "STC", "MZQ", "MIL", "RACO", "TVS", "MAZ")
        $counts = @{}
        foreach ($e in $entities) { $counts[$e] = 0 }
        $otherCount = 0

        for ($i = 1; $i -le $n; $i++) {
            $month = Get-RowMonth $vals[$i, 1]
            if ($month -ne $ReportMonth) { continue }
            $db = [string]$vals[$i, 19]
            if ($counts.ContainsKey($db)) { $counts[$db]++ } else { $otherCount++ }
        }

        $subtotal = ($counts.Values | Measure-Object -Sum).Sum

        $lines = @()
        $lines += "# Report Audit Summary"
        $lines += ""
        $lines += "Report: $ReportPath"
        $lines += "Month audited: $ReportMonth"
        $lines += ""
        $lines += "## 1) Last report data row"
        $lines += "$trueLastRow"
        $lines += ""
        $lines += "## 2) Per-entity row subtotal for $ReportMonth"
        foreach ($e in $entities) { $lines += "- $e : $($counts[$e])" }
        if ($otherCount -gt 0) { $lines += "- (other/unrecognized DB values) : $otherCount" }
        $lines += ""
        $lines += "## 3) Current report subtotal (sum of the entities above) for $ReportMonth"
        $lines += "$subtotal"

        Write-Host ""
        $lines | ForEach-Object { Write-Host $_ }

        if (-not $OutputPath) {
            $OutputPath = Join-Path $WorkingFolder "REPORT\AuditSummary_$ReportMonth.md"
        }
        $lines -join "`r`n" | Set-Content -Path $OutputPath -Encoding UTF8
        Write-Host ""
        Write-Host "Saved: $OutputPath" -ForegroundColor Green
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
