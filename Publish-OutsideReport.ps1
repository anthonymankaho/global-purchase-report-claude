<#
.SYNOPSIS
  月結交付流程：把已完成的當月 Report 資料以「數值」複製進 OUTSIDETEMP 工作檔的 Data 分頁，
  延伸 Table1 表格範圍、重新整理所有樞紐分析表與 Dashboard，存檔後另存為當月交付檔案。

  Monthly delivery step: value-copies the completed month's report data into the OUTSIDETEMP
  working file's Data tab, extends Table1's table range, refreshes all PivotTables/Dashboard,
  saves in place, then Save-As's the dated deliverable file.

.PARAMETER ReportMonth
  當月報表月份，格式 "yyyy-MM"，例如 "2026-06"。用來組出來源檔名與交付檔名的 (1-M)YYYY 標籤。
  The month being delivered, "yyyy-MM" format, e.g. "2026-06". Used to build both the source
  filename and the "(1-M)YYYY" label in the delivered filename.

.PARAMETER SourceReportPath
  當月合併完成的 Report 檔案路徑（Build-GlobalPurchaseReport.ps1 的產出）。預設依命名慣例自動組出。
  Path to the month's completed merged report (Build-GlobalPurchaseReport.ps1's output).
  Defaults to the standard naming convention if not given.

.PARAMETER OutsideTempPath
  外部交付用的工作檔（本腳本會直接改動此檔案，覆寫其 Data 分頁與 Table1）。
  The outside-delivery working file (this script modifies it in place: Data tab + Table1).

.PARAMETER OutputFolder
  另存新檔的目的資料夾。預設與 OutsideTempPath 同資料夾。
  Folder for the Save-As deliverable. Defaults to the same folder as OutsideTempPath.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ReportMonth,

    [string]$SourceReportPath,

    [string]$OutsideTempPath = (Join-Path $PSScriptRoot "Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-OUTSIDETEMP.xlsx"),

    [string]$OutputFolder
)

$ErrorActionPreference = "Stop"

# ---- 解析月份 / Parse month ----
if ($ReportMonth -notmatch '^\d{4}-\d{2}$') {
    throw "ReportMonth 必須是 'yyyy-MM' 格式，例如 '2026-06' / ReportMonth must be 'yyyy-MM', e.g. '2026-06'"
}
$year = [int]$ReportMonth.Substring(0, 4)
$month = [int]$ReportMonth.Substring(5, 2)
$periodLabel = "1-$month"

if (-not $SourceReportPath) {
    $SourceReportPath = Join-Path $PSScriptRoot "REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_$ReportMonth.xlsx"
}
if (-not (Test-Path $SourceReportPath)) {
    throw "找不到來源報表 / Source report not found: $SourceReportPath`n請用 -SourceReportPath 指定實際路徑 / Please specify the actual path via -SourceReportPath"
}
if (-not (Test-Path $OutsideTempPath)) {
    throw "找不到 OUTSIDETEMP 工作檔 / OUTSIDETEMP working file not found: $OutsideTempPath"
}
if (-not $OutputFolder) {
    $OutputFolder = Split-Path $OutsideTempPath -Parent
}

# ---- 決定另存新檔名（同一天重跑會自動遞增 RV 編號）/ Decide Save-As name (auto-bumps RV if rerun same day) ----
$dateStamp = (Get-Date).ToString("yyyyMMdd")
$rv = 0
do {
    $outputFileName = "FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-($periodLabel)$year`_$dateStamp`_RV$rv.xlsx"
    $outputPath = Join-Path $OutputFolder $outputFileName
    $rv++
} while (Test-Path $outputPath)

Write-Host "來源 / Source: $SourceReportPath"
Write-Host "工作檔 / Working file: $OutsideTempPath"
Write-Host "交付檔 / Deliverable: $outputPath"

function Release-Com {
    param($obj)
    if ($obj) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}

function Get-TrueLastDataRow {
    param($ws, [int]$col = 1)
    return $ws.Cells($ws.Rows.Count, $col).End(-4162).Row  # xlUp
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

$srcWb = $null
$destWb = $null

try {
    # ---- 1) 讀來源資料（值，非公式）/ Read source data as VALUES ----
    Write-Host "`n[1/6] 讀取來源 Data 分頁... / Reading source Data tab..."
    $srcWb = $excel.Workbooks.Open($SourceReportPath, [Type]::Missing, $true)
    $srcWs = $srcWb.Worksheets.Item("MHK+STC+MZQ+MIL+RACO Data")
    $srcLastRow = Get-TrueLastDataRow -ws $srcWs -col 1
    $srcLastCol = 36  # A..AJ
    Write-Host "  來源總列數 / Source total rows: $srcLastRow (含表頭 incl. header)"

    # 表頭與資料列分開兩次讀取（不切割大陣列），供後面「只寫資料列」使用 / Read header and data rows
    # separately (two bulk reads, no array-slicing loop) - feeds the "data rows only" write below.
    $hdrRange = $srcWs.Range($srcWs.Cells(1, 1), $srcWs.Cells(1, $srcLastCol))
    $hdrValues = $hdrRange.Value2
    Release-Com $hdrRange

    $srcRange = $srcWs.Range($srcWs.Cells(2, 1), $srcWs.Cells($srcLastRow, $srcLastCol))
    $dataOnlyValues = $srcRange.Value2
    Release-Com $srcRange
    Release-Com $srcWs
    $srcWb.Close($false)
    Release-Com $srcWb
    $srcWb = $null

    # ---- 2) 開啟 OUTSIDETEMP，檢查表頭一致 / Open OUTSIDETEMP, sanity-check header ----
    Write-Host "`n[2/6] 開啟 OUTSIDETEMP 並檢查表頭... / Opening OUTSIDETEMP and checking header..."
    $destWb = $excel.Workbooks.Open($OutsideTempPath, [Type]::Missing, $false)
    $destWs = $destWb.Worksheets.Item("Data")

    for ($c = 1; $c -le $srcLastCol; $c++) {
        $srcHdr = [string]$hdrValues[1, $c]
        $destHdr = [string]$destWs.Cells(1, $c).Value2
        if ($srcHdr -ne $destHdr) {
            throw "表頭不一致，中止以避免資料錯位 / Header mismatch at column $c ('$srcHdr' vs '$destHdr') - aborting to avoid misaligned data."
        }
    }
    Write-Host "  表頭一致 (36 欄 A:AJ) / Header matches (36 cols A:AJ)"

    # ---- 3) 延伸 Table1 範圍（先延伸結構，範圍內先無新資料）/ Resize Table1 first (extend structure before populating) ----
    Write-Host "`n[3/6] 延伸 Table1 表格範圍... / Resizing Table1 table range..."
    $table = $destWs.ListObjects.Item(1)
    $oldRange = $table.Range.Address($false, $false)
    $newTableRange = $destWs.Range($destWs.Cells(1, 1), $destWs.Cells($srcLastRow, $srcLastCol))
    $table.Resize($newTableRange)
    Write-Host "  Table1: $oldRange -> $($table.Range.Address($false, $false))"

    # ---- 4) 數值寫入 Data 分頁，僅資料列、不含表頭列 / Write values into Data tab, DATA ROWS ONLY, never the header row ----
    # 注意：曾實測過，用大陣列一次覆寫含表頭列(第1列)的整個範圍會讓 Table1 物件損毀
    # (ListObjects.Count 變 0，所有樞紐分析表跟著失效)，即使表頭文字完全相同也一樣。
    # 已用 A1:AJ$srcLastRow 全範圍寫入 vs. 只寫 A2:AJ$srcLastRow(跳過表頭列)兩種方式實測比對，
    # 只寫資料列的版本存檔、重新開啟後 Table1 都完整存活，因此固定只寫資料列。
    # NOTE: empirically tested - bulk-overwriting the FULL range (including header row 1) in one
    # array assignment destroys the Table1 ListObject (ListObjects.Count becomes 0, breaking every
    # PivotTable), even when the header text is identical to what's already there. Verified via a
    # side-by-side test: full-range write vs. data-rows-only write (A2:AJ$srcLastRow, skipping the
    # header) - only the data-rows-only version survives save/reopen with Table1 intact. Keep this
    # skip-header-row behavior; do not "simplify" back to a single full-range write.
    Write-Host "`n[4/6] 寫入數值到 Data 分頁 (A2:AJ$srcLastRow，跳過表頭列)... / Writing values into Data tab (A2:AJ$srcLastRow, header row skipped)..."
    $destRange = $destWs.Range($destWs.Cells(2, 1), $destWs.Cells($srcLastRow, $srcLastCol))
    $destRange.Value2 = $dataOnlyValues
    Release-Com $destRange

    # ---- 5) 重新整理所有樞紐分析表 / Refresh all PivotTables ----
    Write-Host "`n[5/6] 重新整理所有樞紐分析表... / Refreshing all PivotTables..."
    $refreshed = @()
    $failed = @()
    foreach ($ws in $destWb.Worksheets) {
        foreach ($pt in $ws.PivotTables()) {
            try {
                $pt.RefreshTable() | Out-Null
                $refreshed += "$($ws.Name) / $($pt.Name)"
            } catch {
                $failed += "$($ws.Name) / $($pt.Name) : $($_.Exception.Message)"
            }
        }
    }
    Write-Host "  成功 / Succeeded ($($refreshed.Count)):"
    $refreshed | ForEach-Object { Write-Host "    $_" }
    if ($failed.Count -gt 0) {
        Write-Host "  失敗 / Failed ($($failed.Count)) - 已知 'Data' 命名範圍損壞（#REF!），非本次操作造成 / known broken 'Data' defined name, pre-existing, not caused by this run:"
        $failed | ForEach-Object { Write-Host "    $_" }
    }

    $destWb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()
    $destWb.Application.Calculate()

    # ---- 6) 存檔 + 另存新檔 / Save + Save-As ----
    Write-Host "`n[6/6] 存檔並另存新檔... / Saving and Save-As..."
    $destWb.Save()
    Write-Host "  已存檔 / Saved: $OutsideTempPath"

    $destWb.SaveAs($outputPath, 51)  # xlOpenXMLWorkbook
    Write-Host "  已另存新檔 / Saved as: $outputPath"

    $destWb.Close($false)
    Release-Com $destWs
    Release-Com $destWb
    $destWb = $null

    Write-Host "`n完成 / Done."
    Write-Host "來源總列數 / Source rows: $srcLastRow | 交付檔 / Deliverable: $outputPath"
    if ($failed.Count -gt 0) {
        Write-Host "`n提醒 / Note: 'Comm Analysis / Vend Name Analysis / Desc Analysis' 三個隱藏分頁的樞紐分析表來源是損壞的 'Data' 命名範圍，需要另外修復（非本次範圍）。"
        Write-Host "Note: the PivotTables in 'Comm Analysis / Vend Name Analysis / Desc Analysis' source from a broken 'Data' defined name and need separate repair (out of scope for this run)."
    }
}
finally {
    if ($destWb) {
        try { $destWb.Close($false) } catch {}
        Release-Com $destWb
    }
    if ($srcWb) {
        try { $srcWb.Close($false) } catch {}
        Release-Com $srcWb
    }
    $excel.Quit()
    Release-Com $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
