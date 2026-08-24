<#
.SYNOPSIS
  Builds the interactive HTML version of the Global Purchase Report from the
  monthly Excel deliverable.

.DESCRIPTION
  1. Opens the source workbook read-only via Excel COM.
  2. Extracts the full Data sheet to JSON.
  3. Natively exports only the Exchange Rate sheet to HTML (the other tabs -
     Dashboard, GLID, Fittings, Packaging, Coating - are built from the Data
     JSON directly by build_report.js, not from a native Excel export).
  4. Runs build_report.js (Node) to assemble the final single-file HTML report.

.PARAMETER SourceExcelPath
  Path to the source "FIN_..._PUR_REPORT-(x-y)YYYY_YYYYMMDD_RVn.xlsx" file.
  If omitted, auto-detects the most recently modified file matching that
  naming pattern in the REPORT folder.

.PARAMETER OutputPath
  Path for the generated HTML file. If omitted, defaults to
  "REPORT\FIN_Global_Purchase -Report_html_<YYYYMM>.html", where YYYYMM is
  taken from the source file's own date stamp.
#>
param(
  [string]$SourceExcelPath,
  [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$reportDir = "C:\Users\User\Desktop\Global Purchase Report Claude\Global Purchase Report RPA\REPORT"

if (-not $SourceExcelPath) {
  $candidates = Get-ChildItem -Path $reportDir -Filter "FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-(*.xlsx" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^FIN_MHK\+STC\+MZQ\+MIL\+RACO_PUR_REPORT-\(\d+-\d+\)\d{4}_(\d{8})_RV?\d+\.xlsx$' }
  if (-not $candidates) {
    throw "No source report file found matching the expected naming pattern in: $reportDir"
  }
  $SourceExcelPath = ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}
if (-not (Test-Path $SourceExcelPath)) {
  throw "Source file not found: $SourceExcelPath"
}
Write-Host "Source: $SourceExcelPath"

$sourceFileName = Split-Path $SourceExcelPath -Leaf
# Filename encodes the reporting period as "(startMonth-endMonth)YYYY", e.g.
# "(1-6)2026" = Jan-Jun 2026. Use the period's END month/year for the output
# name, not the "_YYYYMMDD_" stamp (that's just when the file was generated).
if ($sourceFileName -notmatch '\(\d+-(\d+)\)(\d{4})') {
  throw "Could not find a (start-end)YYYY reporting-period pattern in source filename: $sourceFileName"
}
$endMonth = [int]$Matches[1]
$year = $Matches[2]
$yyyymm = "{0}{1:D2}" -f $year, $endMonth

if (-not $OutputPath) {
  $OutputPath = Join-Path $reportDir "FIN_Global_Purchase -Report_html_$yyyymm.html"
}
Write-Host "Output: $OutputPath"

$workDir = Join-Path $scriptDir "_work"
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($SourceExcelPath, [Type]::Missing, $true)  # ReadOnly

  # --- 1. Data sheet -> JSON ---
  Write-Host "Extracting Data sheet..."
  $ws = $wb.Worksheets.Item("Data")
  $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row       # xlUp
  $lastCol = $ws.Cells(1, $ws.Columns.Count).End(-4159).Column  # xlToLeft, from header row
  Write-Host "  $lastRow rows x $lastCol cols"

  $range = $ws.Range($ws.Cells(1, 1), $ws.Cells($lastRow, $lastCol))
  $vals = $range.Value2

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"header":[')
  $c = 1
  while ($c -le $lastCol) {
    if ($c -gt 1) { [void]$sb.Append(',') }
    $cellStr = ([string]$vals.GetValue(1, $c)).Replace('\', '\\').Replace('"', '\"')
    [void]$sb.Append('"' + $cellStr + '"')
    $c++
  }
  [void]$sb.Append('],"rows":[')

  $r = 2
  while ($r -le $lastRow) {
    if ($r -gt 2) { [void]$sb.Append(',') }
    [void]$sb.Append('[')
    $c = 1
    while ($c -le $lastCol) {
      if ($c -gt 1) { [void]$sb.Append(',') }
      $cellVal = $vals.GetValue($r, $c)
      if ($null -eq $cellVal) {
        [void]$sb.Append('null')
      } elseif ($cellVal -is [double]) {
        [void]$sb.Append([string]$cellVal)
      } else {
        $cellStr = [string]$cellVal
        $cellStr = $cellStr.Replace('\', '\\').Replace('"', '\"')
        $cellStr = [System.Text.RegularExpressions.Regex]::Replace($cellStr, '[\x00-\x1F]', ' ')
        [void]$sb.Append('"' + $cellStr + '"')
      }
      $c++
    }
    [void]$sb.Append(']')
    if (($r % 5000) -eq 0) { Write-Host "  ...$r rows processed" }
    $r++
  }
  [void]$sb.Append(']}')

  $noBomUtf8 = New-Object System.Text.UTF8Encoding $false
  [System.IO.File]::WriteAllText((Join-Path $workDir "data_sheet.json"), $sb.ToString(), $noBomUtf8)
  Write-Host "  Wrote data_sheet.json"

  # --- 2. Exchange Rate sheet -> native HTML export (only sheet kept as native) ---
  # Excel's "Save As Web Page" exports every VISIBLE sheet in the workbook
  # regardless of Worksheet.Select() - selection only changes which sheet is
  # active, not what gets included. To get just Exchange Rate, temporarily
  # hide every other currently-visible sheet first (safe: the workbook was
  # opened read-only and is never saved, so these visibility changes are
  # never written back to the source file).
  Write-Host "Exporting Exchange Rate sheet..."
  $hiddenSheets = @()
  foreach ($s in $wb.Worksheets) {
    if ($s.Name -ne "Exchange Rate" -and $s.Visible -eq -1) {  # xlSheetVisible
      $s.Visible = 0  # xlSheetHidden
      $hiddenSheets += $s.Name
    }
  }
  $htmlPath = Join-Path $workDir "report.htm"
  $wb.SaveAs($htmlPath, 44)  # xlHtml
  Write-Host "  Saved native export (hid: $($hiddenSheets -join ', '))"

  $wb.Close($false)
} finally {
  # Release COM references in reverse order of creation before Quit, so no
  # lingering RCW keeps an invisible EXCEL.EXE process alive after this script exits.
  if ($range) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null }
  if ($ws) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
  if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  Remove-Variable range, ws, wb, excel -ErrorAction SilentlyContinue
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

# --- 3. Assemble the final combined HTML report ---
Write-Host "Building combined HTML report..."
& node (Join-Path $scriptDir "build_report.js") $workDir $OutputPath
if ($LASTEXITCODE -ne 0) {
  throw "build_report.js failed with exit code $LASTEXITCODE"
}

Remove-Item $workDir -Recurse -Force
Write-Host "Done: $OutputPath"
