# PowerShell 指令速查 / PowerShell Commands Quick Reference

以下是本專案月結流程的所有 PowerShell 指令：
Here are all the PowerShell commands for this project's monthly pipeline:

> ⚠️ **最容易打錯的地方**：指令開頭一定要有 `.\`（一個點 + 反斜線），代表「目前資料夾」。
> 只打 `\腳本名稱`（斜線開頭，沒有點）或 `腳本名稱`（完全沒有前綴）都會出現
> `無法辨識...詞彙是否為 Cmdlet` 的錯誤，因為 PowerShell 預設不會從目前資料夾載入指令。
>
> ⚠️ **Most common mistake**: the command must start with `.\` (a dot + backslash), meaning "current folder".
> Typing `\ScriptName` (leading slash, no dot) or just `ScriptName` (no prefix at all) will both fail with
> a "not recognized as a Cmdlet" error, since PowerShell doesn't load commands from the current folder by default.

## 0. 切換到專案資料夾 + 設定當月月份 / Go to the project folder + set this month's period

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude"
$ReportMonth = "2026-07"
```

> ⚠️ **每個月只需要改這一行**：`$ReportMonth = "2026-07"` 改成 `"2026-08"`（或當月月份），下面所有指令都會自動套用，不用逐一修改。
> 但這個變數只在**同一個** PowerShell 視窗裡有效——如果中途關掉視窗重開，要記得先重新執行這一行，再繼續下面的步驟。
>
> ⚠️ **Only this one line changes each month**: update `$ReportMonth = "2026-07"` to `"2026-08"` (or the current period) — every command below picks it up automatically, no need to edit each one individually.
> This variable only lives in the **same** PowerShell window — if you close and reopen the window partway through, re-run this line first before continuing with the steps below.

## 1. 每月主要建置 / Monthly main build

```powershell
.\Build-GlobalPurchaseReport.ps1 -ReportMonth $ReportMonth
```

> ⚠️ 開始執行時會**自動強制關閉**所有正在執行的 EXCEL.EXE（不會詢問）。執行前請先自行儲存並關閉您開著的 Excel 檔案。
> ⚠️ Automatically **force-kills** any running EXCEL.EXE at startup, without asking. Save and close your own Excel files first.

Q:S 跟 U/V/W/Y/Z/AJ 會自動填好。**AA/AB 不會**——腳本執行完，畫面會印出新增列範圍，先手動把
`AA2:AB2` 複製到那個範圍、算過、確認無誤，再往下執行 Step 3。
Q:S and U/V/W/Y/Z/AJ get filled automatically. **AA/AB do not** — the script prints the new row range;
manually copy `AA2:AB2` down to that range, calculate, confirm it looks right, then proceed to Step 3.

## 2. 補 MAZ 的 PEN 匯率缺口（若當月 MAZ 有 PEN 幣別交易才需要）/ Fill MAZ's PEN rate gap (only if this month's MAZ has PEN-currency rows)

```powershell
.\Get-PenExchangeRate.ps1 -ReportMonth $ReportMonth
```

需等 Step 1 產出當月報表後才能執行。會從 Pacific Exchange Rate Service 抓 USD/PEN，換算成 HKD/PEN
（乘上該月 K欄的 USD 匯率），寫入報表「Exchange Rate」分頁 M欄；Z欄的 PEN→USD 公式會自動重算，不會
寫死數字。
Requires Step 1's output to already exist. Fetches USD/PEN from the Pacific Exchange Rate Service,
converts it to HKD/PEN (× that month's column K USD rate), writes it into the report's "Exchange Rate"
tab column M; column Z's PEN→USD formula recomputes on its own — never a hardcoded number.

## 3. 【人工填好 AA/AB 之後】填 AC/AD/AE/AF/AG/AH/AI + New Vendors 分頁 / [After manually filling AA/AB] fill AC–AI + New Vendors sheet

```powershell
.\Complete-FormulaColumns.ps1 -ReportMonth $ReportMonth
```

會自動用 B 欄（PO_REC_DATE）找出當月新增列範圍，檢查 AA 是否已經人工填好（沒填會直接報錯中止），
才繼續填 AC/AF/AD/AE/AG/AH/AI 並凍結成數值、建立「New Vendors」檢查分頁。
Auto-detects the month's new row range via column B (PO_REC_DATE), checks AA was manually filled first
(aborts with a clear error if not), then fills/freezes AC/AF/AD/AE/AG/AH/AI and builds the "New Vendors"
review sheet.

## 4. TVS 翻譯/商品代碼檢查（如有需要）/ TVS translation/commodity check (if needed)

```powershell
.\Update-TVSTranslations.ps1
```

> 若當月已透過 Excel Copilot（或其他方式）手動整理過 TVS 的 Commodity 代碼，這一步可以略過。
> Skip this step if TVS's Commodity codes were already cleaned up manually this month (e.g. via Excel Copilot).

## 5.（建議）產出稽核摘要報告 / (Recommended) generate an audit summary

只唯讀開啟報表，不會改動任何東西：報表最後列號、七實體當月列數、當月總列數。
Read-only, never modifies the report: last data row, per-entity row counts for the month, total row count.

```powershell
.\New-ReportAuditSummary.ps1 -ReportMonth $ReportMonth
```

## 6. 產出外部交付檔案 / Generate the outside-delivery deliverables

把當月報表 Data 分頁的數值複製進 **OUTSIDETEMP.xlsx** 和 **OUTSIDETEMP_VBA_PRACTICE.xlsm** 兩個檔案、
延伸各自的 Table1 表格範圍、重新整理所有樞紐分析表（PivotCharts 隨之自動更新）。OUTSIDETEMP.xlsx
存檔後會再另存為當月交付檔（`FIN_..._PUR_REPORT-(1-M)YYYY_YYYYMMDD_RV0.xlsx`，同一天重跑會自動遞增
RV 編號）；VBA 練習檔只存檔於原地，不產生另存新檔的交付版本。
Value-copies the month's report Data tab into both **OUTSIDETEMP.xlsx** and
**OUTSIDETEMP_VBA_PRACTICE.xlsm**, extends each Table1 range, refreshes all PivotTables (PivotCharts
update automatically). OUTSIDETEMP.xlsx is saved then Save-As'd as the dated deliverable (auto-bumps
the RV number if rerun the same day); the VBA practice file is saved in place only.

```powershell
.\Publish-OutsideReport.ps1 -ReportMonth $ReportMonth
```

> ⚠️ 這支腳本會直接改動 `OUTSIDETEMP.xlsx` 和 `OUTSIDETEMP_VBA_PRACTICE.xlsm` 本身（不是唯讀範本）。
> 執行前若偵測到任何 EXCEL.EXE 在跑（可能是這兩個檔案正開在您自己的視窗），腳本會直接報錯中止——
> **不會**自動強制關閉，請先自行存檔關閉所有 Excel 視窗再重新執行。
>
> ⚠️ This script modifies `OUTSIDETEMP.xlsx` and `OUTSIDETEMP_VBA_PRACTICE.xlsm` directly (not
> read-only templates). If any EXCEL.EXE is detected running beforehand (possibly these two files open
> in your own window), the script throws and aborts — it will **not** force-close anything; save and
> close all Excel windows yourself first, then re-run.

若只想同步 OUTSIDETEMP.xlsx、略過 VBA 練習檔：
To sync only OUTSIDETEMP.xlsx and skip the VBA practice file:

```powershell
.\Publish-OutsideReport.ps1 -ReportMonth $ReportMonth -VbaPracticePath ""
```

## 若要強制重跑同一個月（略過重複執行安全檢查）/ To force-rerun the same month (bypass the duplicate-run safety check)

```powershell
.\Build-GlobalPurchaseReport.ps1 -ReportMonth $ReportMonth -Force
```

## 若自動偵測失敗、需要手動指定路徑 / If auto-detection fails and paths need to be specified manually

RACO/TVS 樣板預設固定抓 `Trimmed Program\RACO(或 TVS) Act001 Trimmed.xlsx`（未命名日期的範本檔），一般不需要手動指定；只有在範本檔被搬到別的地方時才需要：
RACO/TVS templates default to the fixed `Trimmed Program\RACO(or TVS) Act001 Trimmed.xlsx` (the undated template) — you normally don't need to override this; only needed if the template file has been moved elsewhere:

```powershell
.\Build-GlobalPurchaseReport.ps1 `
  -ReportMonth $ReportMonth `
  -RacoTrimmedFile "C:\path\to\RACO Act001 Trimmed.xlsx" `
  -TvsTrimmedFile "C:\path\to\TVS Act001 Trimmed.xlsx"
```

## 一次性事前準備（僅需設定一次）/ One-time setup (only needed once)

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```
