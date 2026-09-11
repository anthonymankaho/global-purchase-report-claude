<#
.SYNOPSIS
  Copies one named column's values (default ProdFamiDesc) from an OUTSIDETEMP-style
  workbook's Data sheet into the official monthly source report, matching rows by
  POSITION only after verifying every row's key columns (PO_NO, ITEM_CODE,
  VEND_INV_NO, REC_QTY) actually line up between the two files - abort-before-write
  if not.

.DESCRIPTION
  OUTSIDETEMP.xlsx/OUTSIDETEMP_VBA_PRACTICE.xlsm are originally synced FROM the
  official source report by Publish-OutsideReport.ps1 (straight row-by-row copy,
  same order), so their rows should still align 1:1 with the source report's rows.
  This script re-verifies that assumption row-by-row (not just spot-checked) before
  writing anything, since silently misaligning a column's data against the wrong
  transaction would be worse than not adding the column at all.

  Values only (not any formula) are copied - the source report doesn't need to know
  how the column was computed, just the resulting value.

.PARAMETER SourceReportPath
  The official monthly report to add/update the column in (modified in place -
  overwrites this file, per explicit user confirmation).

.PARAMETER OutsideTempPath
  Path to the workbook that already has the column filled in (OUTSIDETEMP.xlsx or
  OUTSIDETEMP_VBA_PRACTICE.xlsm, whichever one the column was actually added to).

.PARAMETER ColumnName
  The Data-sheet header text of the column to copy, e.g. "ProdFamiDesc" (default)
  or "ProdFamiCode".

.EXAMPLE
  .\Copy-ProdFamiDescToSource.ps1
  .\Copy-ProdFamiDescToSource.ps1 -ColumnName "ProdFamiCode" -OutsideTempPath "...\OUTSIDETEMP_VBA_PRACTICE.xlsm"
#>
param(
  [string]$SourceReportPath = "C:\Users\User\Desktop\Global Purchase Report Claude\Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-(1-7)2026_20260826_RV0.xlsx",
  [string]$OutsideTempPath = "C:\Users\User\Desktop\Global Purchase Report Claude\Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-OUTSIDETEMP.xlsx",
  [string]$ColumnName = "ProdFamiDesc"
)

$ErrorActionPreference = "Stop"

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
  throw "Excel is currently running - close all Excel windows first (this script needs read-write access and a second COM instance risks conflicting with an open workbook)."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wbSrc = $null; $wbOut = $null
try {
  Write-Host "Opening OUTSIDETEMP (read-only): $OutsideTempPath"
  $wbOut = $excel.Workbooks.Open($OutsideTempPath, [Type]::Missing, $true)
  $wsOut = $wbOut.Worksheets.Item("Data")
  $lastRowOut = $wsOut.Cells($wsOut.Rows.Count, 1).End(-4162).Row
  $lastColOut = $wsOut.Cells(1, $wsOut.Columns.Count).End(-4159).Column
  $famColOut = 0
  for ($c = 1; $c -le $lastColOut; $c++) {
    if ($wsOut.Cells(1, $c).Text.Trim() -eq $ColumnName) { $famColOut = $c; break }
  }
  if ($famColOut -eq 0) { throw "$ColumnName column not found in OUTSIDETEMP's Data sheet." }
  Write-Host "  $ColumnName is column $famColOut in OUTSIDETEMP ($($lastRowOut - 1) data rows)"

  # Key columns for row-alignment verification: PO_NO(1), VEND_INV_NO(6), ITEM_CODE(8), REC_QTY(12)
  $keyColsOut = $wsOut.Range($wsOut.Cells(2, 1), $wsOut.Cells($lastRowOut, 12)).Value2
  $famValsOut = $wsOut.Range($wsOut.Cells(2, $famColOut), $wsOut.Cells($lastRowOut, $famColOut)).Value2

  Write-Host "Opening source report (read-write): $SourceReportPath"
  $wbSrc = $excel.Workbooks.Open($SourceReportPath, [Type]::Missing, $false)
  $wsSrc = $wbSrc.Worksheets.Item("Data")
  $lastRowSrc = $wsSrc.Cells($wsSrc.Rows.Count, 1).End(-4162).Row
  $lastColSrc = $wsSrc.Cells(1, $wsSrc.Columns.Count).End(-4159).Column
  Write-Host "  Source report has $($lastRowSrc - 1) data rows, $lastColSrc columns"

  if (($lastRowOut - 1) -ne ($lastRowSrc - 1)) {
    throw "Row count mismatch: OUTSIDETEMP has $($lastRowOut - 1) data rows, source report has $($lastRowSrc - 1). Aborting - not safe to copy by position."
  }

  $keyColsSrc = $wsSrc.Range($wsSrc.Cells(2, 1), $wsSrc.Cells($lastRowSrc, 12)).Value2

  # ITEM_CODE can be stored as text in one file and as a number in the other
  # (leading zeros like "01253" vs 1253) - normalize by stripping leading zeros
  # from purely-numeric-looking values before comparing, so a display/typing
  # difference isn't mistaken for a real row misalignment.
  function Normalize($v) {
    $s = [string]$v
    if ($s -match '^\d+$') { return $s.TrimStart('0'); }
    return $s
  }

  Write-Host "Verifying row-by-row alignment (PO_NO, VEND_INV_NO, ITEM_CODE, REC_QTY)..."
  $rowCount = $lastRowOut - 1
  $mismatches = 0
  $benignItemCodeOnly = 0
  for ($r = 1; $r -le $rowCount; $r++) {
    $poOut = $keyColsOut.GetValue($r, 1); $poSrc = $keyColsSrc.GetValue($r, 1)
    $invOut = $keyColsOut.GetValue($r, 6); $invSrc = $keyColsSrc.GetValue($r, 6)
    $itemOut = $keyColsOut.GetValue($r, 8); $itemSrc = $keyColsSrc.GetValue($r, 8)
    $qtyOut = $keyColsOut.GetValue($r, 12); $qtySrc = $keyColsSrc.GetValue($r, 12)
    $poMatch = (Normalize $poOut) -eq (Normalize $poSrc)
    $invMatch = (Normalize $invOut) -eq (Normalize $invSrc)
    $itemMatch = (Normalize $itemOut) -eq (Normalize $itemSrc)
    $qtyMatch = (Normalize $qtyOut) -eq (Normalize $qtySrc)
    if (-not ($poMatch -and $invMatch -and $itemMatch -and $qtyMatch)) {
      $mismatches++
      if ($mismatches -le 10) {
        Write-Host "  MISMATCH at row $($r+1): OUTSIDETEMP=[$poOut,$invOut,$itemOut,$qtyOut] SOURCE=[$poSrc,$invSrc,$itemSrc,$qtySrc]" -ForegroundColor Red
      }
    }
  }
  if ($mismatches -gt 0) {
    throw "$mismatches row(s) did not align by position even after normalizing text-vs-number differences - aborting before writing anything. The two files' row order/content has diverged; positional copy is not safe."
  }
  Write-Host "  All $rowCount rows aligned correctly (after normalizing text-vs-number ITEM_CODE differences)." -ForegroundColor Green

  $famColSrc = 0
  for ($c = 1; $c -le $lastColSrc; $c++) {
    if ($wsSrc.Cells(1, $c).Text.Trim() -eq $ColumnName) { $famColSrc = $c; break }
  }
  if ($famColSrc -eq 0) {
    $famColSrc = $lastColSrc + 1
    $wsSrc.Cells(1, $famColSrc).Value2 = $ColumnName
    Write-Host "  Added new column $famColSrc '$ColumnName' to source report."
  } else {
    Write-Host "  $ColumnName already exists at column $famColSrc in source report - overwriting its values."
  }

  $wsSrc.Range($wsSrc.Cells(2, $famColSrc), $wsSrc.Cells($lastRowSrc, $famColSrc)).Value2 = $famValsOut

  $wbSrc.Save()
  Write-Host "Saved: $SourceReportPath" -ForegroundColor Green
} finally {
  if ($wbOut) { $wbOut.Close($false) }
  if ($wbSrc) { $wbSrc.Close($false) }
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  Remove-Variable wbSrc, wbOut, excel -ErrorAction SilentlyContinue
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
