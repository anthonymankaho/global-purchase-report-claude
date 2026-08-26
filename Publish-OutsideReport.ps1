<#
.SYNOPSIS
  月結交付流程：把已完成的當月 Report 資料以「數值」複製進 OUTSIDETEMP 工作檔的 Data 分頁，
  延伸 Table1 表格範圍、重新整理所有樞紐分析表與 Dashboard，存檔後另存為當月交付檔案；
  同一批來源資料也會同步進 OUTSIDETEMP_VBA_PRACTICE.xlsm（VBA 按鈕原型測試用的練習檔，僅
  存檔於原地，不產生另存新檔的交付版本）。

  Monthly delivery step: value-copies the completed month's report data into the OUTSIDETEMP
  working file's Data tab, extends Table1's table range, refreshes all PivotTables/Dashboard,
  saves in place, then Save-As's the dated deliverable file. The same source data is also
  synced into OUTSIDETEMP_VBA_PRACTICE.xlsm (the VBA-button-prototype practice copy) - saved
  in place only, no dated deliverable produced for it.

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

.PARAMETER VbaPracticePath
  VBA 按鈕原型的練習檔（結構與 OutsideTempPath 相同：Data 分頁 + Table1 + 同一組樞紐分析表）。
  同步邏輯與 OutsideTempPath 完全相同，但只存檔於原地，不產生另存新檔的交付版本。
  傳入空字串可略過這個檔案。
  The VBA-button-prototype practice file (same structure as OutsideTempPath: Data tab + Table1 +
  the same set of PivotTables). Synced with the identical logic as OutsideTempPath, but saved in
  place only - no dated deliverable is produced for it. Pass an empty string to skip it.

.PARAMETER OutputFolder
  另存新檔的目的資料夾。預設與 OutsideTempPath 同資料夾。
  Folder for the Save-As deliverable. Defaults to the same folder as OutsideTempPath.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ReportMonth,

    [string]$SourceReportPath,

    [string]$OutsideTempPath = (Join-Path $PSScriptRoot "Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-OUTSIDETEMP.xlsx"),

    [string]$VbaPracticePath = (Join-Path $PSScriptRoot "Global Purchase Report RPA\REPORT\OUTSIDETEMP_VBA_PRACTICE.xlsm"),

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
if ($VbaPracticePath -and -not (Test-Path $VbaPracticePath)) {
    throw "找不到 VBA 練習檔 / VBA practice file not found: $VbaPracticePath`n若不需要同步這個檔案，請傳入 -VbaPracticePath (空字串) / pass -VbaPracticePath (empty string) to skip it."
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
Write-Host "工作檔 1 / Working file 1: $OutsideTempPath"
if ($VbaPracticePath) { Write-Host "工作檔 2 / Working file 2: $VbaPracticePath" }
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

# 把來源的值寫進某一個目的檔的 Data 分頁：檢查表頭、延伸 Table1、寫入資料列、重新整理樞紐分析表。
# 回傳 refreshed/failed 清單，呼叫端自行決定是否存檔/另存新檔。
# Writes the source values into one destination file's Data tab: header check, Table1 resize,
# data-rows write, PivotTable refresh. Returns refreshed/failed lists; caller decides save/Save-As.
function Sync-DataIntoWorkbook {
    param($excel, [string]$label, [string]$path, $hdrValues, $dataOnlyValues, [int]$srcLastRow, [int]$srcLastCol)

    Write-Host "`n--- $label ---"
    Write-Host "  開啟並檢查表頭... / Opening and checking header..."
    $wb = $excel.Workbooks.Open($path, [Type]::Missing, $false)
    if ($wb.ReadOnly) {
        throw "$label 以唯讀開啟，可能已被其他 Excel 視窗鎖定 / $label opened read-only - it may be locked open in another Excel window: $path"
    }
    $ws = $wb.Worksheets.Item("Data")

    for ($c = 1; $c -le $srcLastCol; $c++) {
        $srcHdr = [string]$hdrValues[1, $c]
        $destHdr = [string]$ws.Cells(1, $c).Value2
        if ($srcHdr -ne $destHdr) {
            $wb.Close($false); Release-Com $wb
            throw "$label 表頭不一致，中止以避免資料錯位 / $label header mismatch at column $c ('$srcHdr' vs '$destHdr') - aborting to avoid misaligned data."
        }
    }
    Write-Host "  表頭一致 (36 欄 A:AJ) / Header matches (36 cols A:AJ)"

    Write-Host "  延伸 Table1 表格範圍... / Resizing Table1 table range..."
    $table = $ws.ListObjects.Item(1)
    $oldRange = $table.Range.Address($false, $false)
    $newTableRange = $ws.Range($ws.Cells(1, 1), $ws.Cells($srcLastRow, $srcLastCol))
    $table.Resize($newTableRange)
    Write-Host "  Table1: $oldRange -> $($table.Range.Address($false, $false))"

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
    Write-Host "  寫入數值到 Data 分頁 (A2:AJ$srcLastRow，跳過表頭列)... / Writing values into Data tab (A2:AJ$srcLastRow, header row skipped)..."
    $destRange = $ws.Range($ws.Cells(2, 1), $ws.Cells($srcLastRow, $srcLastCol))
    $destRange.Value2 = $dataOnlyValues
    Release-Com $destRange

    Write-Host "  重新整理所有樞紐分析表... / Refreshing all PivotTables..."
    $refreshed = @()
    $failed = @()
    foreach ($sheet in $wb.Worksheets) {
        foreach ($pt in $sheet.PivotTables()) {
            try {
                $pt.RefreshTable() | Out-Null
                $refreshed += "$($sheet.Name) / $($pt.Name)"
            } catch {
                $failed += "$($sheet.Name) / $($pt.Name) : $($_.Exception.Message)"
            }
        }
    }
    Write-Host "  成功 / Succeeded ($($refreshed.Count)); 失敗 / Failed ($($failed.Count))"
    if ($failed.Count -gt 0) { $failed | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow } }

    # PivotCharts 依附在各自的樞紐分析表上，RefreshTable() 後會自動同步更新，不需另外處理。
    # PivotCharts are bound to their PivotTable and update automatically once it refreshes - no
    # separate chart-refresh step is needed.
    $wb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()
    $wb.Application.Calculate()

    return @{ Wb = $wb; Ws = $ws; Refreshed = $refreshed; Failed = $failed }
}

$existingExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($existingExcel) {
    throw "偵測到 $($existingExcel.Count) 個正在執行的 EXCEL.EXE（可能是 OUTSIDETEMP 或練習檔已在您自己的視窗中開啟）。請先手動儲存並關閉所有 Excel 視窗，再重新執行本腳本 - 本腳本不會自動強制關閉，以免弄丟您未存檔的工作。`nDetected $($existingExcel.Count) running EXCEL.EXE process(es) - OUTSIDETEMP or the practice file may be open in your own window. Please save and close all Excel windows yourself, then re-run - this script will NOT force-close them, to avoid losing your unsaved work."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
# 手動計算模式：Table1.Resize()、寫入數萬列的值、每個 PivotTable.RefreshTable() 若在自動計算模式下，
# 每一步都會觸發整個活頁簿（Dashboard/Analysis 分頁等仍是公式）重算一次，這正是這個專案之前遇過的
# 「卡10分鐘以上」問題（Build-GlobalPurchaseReport.ps1 也用同樣方式防範）。改成手動後，只有
# Sync-DataIntoWorkbook 最後那一次明確的 Application.Calculate() 才會真正重算，其餘中間步驟都不會。
# Manual calculation mode: under Automatic, Table1.Resize(), writing tens of thousands of value rows,
# and each PivotTable.RefreshTable() would each independently trigger a full-workbook recalculation
# (Dashboard/Analysis tabs are still formulas) - the exact "10+ minute hang" pattern this project has
# hit before (Build-GlobalPurchaseReport.ps1 guards against it the same way). With Manual, only the
# single explicit Application.Calculate() at the end of Sync-DataIntoWorkbook actually recalculates;
# none of the intermediate steps do.
$excel.Calculation = -4135  # xlCalculationManual

$srcWb = $null
$destWb1 = $null
$destWb2 = $null

try {
    # ---- 1) 讀來源資料（值，非公式）/ Read source data as VALUES ----
    Write-Host "`n[1/3] 讀取來源 Data 分頁... / Reading source Data tab..."
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

    # ---- 2) 同步進 OUTSIDETEMP.xlsx，存檔後另存新檔 / Sync into OUTSIDETEMP.xlsx, save + Save-As ----
    Write-Host "`n[2/3] 同步 OUTSIDETEMP.xlsx... / Syncing OUTSIDETEMP.xlsx..."
    $result1 = Sync-DataIntoWorkbook -excel $excel -label "OUTSIDETEMP.xlsx" -path $OutsideTempPath `
        -hdrValues $hdrValues -dataOnlyValues $dataOnlyValues -srcLastRow $srcLastRow -srcLastCol $srcLastCol
    $destWb1 = $result1.Wb

    $destWb1.Save()
    Write-Host "  已存檔 / Saved: $OutsideTempPath"
    $destWb1.SaveAs($outputPath, 51)  # xlOpenXMLWorkbook
    Write-Host "  已另存新檔 / Saved as: $outputPath"

    $destWb1.Close($false)
    Release-Com $result1.Ws
    Release-Com $destWb1
    $destWb1 = $null

    # ---- 3) 同步進 OUTSIDETEMP_VBA_PRACTICE.xlsm，僅存檔於原地 / Sync into the VBA practice file, save in place only ----
    $result2 = $null
    if ($VbaPracticePath) {
        Write-Host "`n[3/3] 同步 OUTSIDETEMP_VBA_PRACTICE.xlsm... / Syncing OUTSIDETEMP_VBA_PRACTICE.xlsm..."
        $result2 = Sync-DataIntoWorkbook -excel $excel -label "OUTSIDETEMP_VBA_PRACTICE.xlsm" -path $VbaPracticePath `
            -hdrValues $hdrValues -dataOnlyValues $dataOnlyValues -srcLastRow $srcLastRow -srcLastCol $srcLastCol
        $destWb2 = $result2.Wb

        $destWb2.Save()
        Write-Host "  已存檔 / Saved: $VbaPracticePath"

        $destWb2.Close($false)
        Release-Com $result2.Ws
        Release-Com $destWb2
        $destWb2 = $null
    } else {
        Write-Host "`n[3/3] 已略過 VBA 練習檔（-VbaPracticePath 為空字串）/ Skipped VBA practice file (-VbaPracticePath was empty)."
    }

    Write-Host "`n完成 / Done."
    Write-Host "來源總列數 / Source rows: $srcLastRow | 交付檔 / Deliverable: $outputPath"
    $anyFailed = $result1.Failed.Count -gt 0 -or ($result2 -and $result2.Failed.Count -gt 0)
    if ($anyFailed) {
        Write-Host "`n提醒 / Note: 上方列出的樞紐分析表重新整理失敗，請檢查該分頁的資料來源設定。"
        Write-Host "Note: the PivotTable refresh failures listed above need their data-source settings checked."
    }
}
finally {
    if ($destWb2) {
        try { $destWb2.Close($false) } catch {}
        Release-Com $destWb2
    }
    if ($destWb1) {
        try { $destWb1.Close($false) } catch {}
        Release-Com $destWb1
    }
    if ($srcWb) {
        try { $srcWb.Close($false) } catch {}
        Release-Com $srcWb
    }
    $excel.Quit()
    Release-Com $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $stillRunning = $null
    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 1
        $stillRunning = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
        if (-not $stillRunning) { break }
    }
    if ($stillRunning) {
        Write-Host "`n警告 / Warning: 本腳本執行後仍有 EXCEL.EXE 在執行 ($($stillRunning.Count) 個)。若持續發生，請用工作管理員手動關閉後再重新執行。" -ForegroundColor Yellow
        Write-Host "Warning: EXCEL.EXE still running after this script ($($stillRunning.Count) process(es)). If this persists, close them manually via Task Manager before running this script again." -ForegroundColor Yellow
    }
}
