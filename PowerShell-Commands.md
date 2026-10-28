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
Set-Location -LiteralPath "C:\Users\User\Desktop\Global Purchase Report Claude"; $ReportMonth = "2026-09"
```

> ⚠️ **每個月只需要改這一行**：把 `$ReportMonth = "2026-09"` 改成當月月份，下面所有指令都會自動套用，不用逐一修改。
> 請整行執行，或在多個指令之間保留換行/分號。不要把 `Set-Location` 和 `$ReportMonth` 直接接在一起；否則 PowerShell 會把變數指派當成 `cd` 的引數並報位置參數錯誤。
> 但這個變數只在**同一個** PowerShell 視窗裡有效——如果中途關掉視窗重開，要記得先重新執行這一行，再繼續下面的步驟。
>
> ⚠️ **Only this one line changes each month**: update `$ReportMonth = "2026-09"` to the reporting month — every command below picks it up automatically, no need to edit each one individually.
> Run the whole line, or keep a newline/semicolon between commands. Do not concatenate `Set-Location` directly with `$ReportMonth`; PowerShell will treat the assignment as an extra `cd` argument and report a positional-parameter error.
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

若建置曾在公式欄位填入時出現 `Range.PasteSpecial` 錯誤，請使用已更新的 `Build-GlobalPurchaseReport.ps1`，關閉 Excel 後以相同月份重跑上方指令。失敗的那次沒有存入當月資料；輸出路徑可能只有上一個月報表的複本。重跑會從自動偵測的上一個月報表重新建立當月輸出，通常不需要加 `-Force`。只有要刻意略過重複月份檢查時才使用下方的 `-Force` 指令。

If a previous build stopped with a `Range.PasteSpecial` error while filling formula columns, use the updated `Build-GlobalPurchaseReport.ps1`, close Excel, and rerun the command above for the same month. The failed run did not save the new month's data; the output path may contain only a copy of the prior-month report. The retry rebuilds the month's output from the auto-detected prior-month report. Usually do not add `-Force`; reserve it for intentionally bypassing the duplicate-month safety check.

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

若出現「無法解析遠端名稱 fx.sauder.ubc.ca」，代表電腦的 DNS/網路無法連到匯率網站；腳本會停止且不會改報表。先確認網路、DNS 或 VPN 後重試。若服務仍不可用，請從可信來源人工核對報表月底當天及往前 7 天範圍內的 **USD per 1 PEN** 匯率，再用以下參數執行（來源與匯率日期必填）：

If PowerShell reports that it cannot resolve `fx.sauder.ubc.ca`, the PC could not reach the rate site through DNS/network; the script stops without changing the report. Check internet, DNS or VPN and retry. If the service remains unavailable, verify the **USD per 1 PEN** rate from a trusted source for a date from month-end through the preceding 7 days, then run:

```powershell
.\Get-PenExchangeRate.ps1 -ReportMonth $ReportMonth -UsdPerPen 0.2945 -RateDate "2026-09-30" -RateSource "Verified source name or URL"
```

Replace the example rate, date and source with the verified quote. Do not enter PEN per USD; the input must be USD for one PEN.

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

先確認所有 Excel 視窗已儲存並關閉；腳本要求電腦上完全沒有 EXCEL.EXE 執行中。可先查看目前程序：
Before running, save and close every Excel window; the script requires that no EXCEL.EXE process is running. Check the current processes first:

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id, MainWindowTitle, StartTime
```

若仍列出程序，先確認對應活頁簿已儲存並正常關閉，再重新執行。沒有視窗標題的背景 Excel 程序也會讓腳本中止；請先確認沒有未存檔工作，再於工作管理員關閉殘留程序。不要用 `-Force` 嘗試繞過這項安全檢查。
If a process is still listed, make sure its workbook is saved and close Excel normally, then retry. A background Excel process with no window title also blocks the script; confirm there is no unsaved work before ending a confirmed orphan in Task Manager. Do not use `-Force` to bypass this safety check.

腳本會依來源 Data 分頁最後一個表頭動態判斷欄數，並在寫入前核對兩份工作檔的表頭與 Table 欄數；若欄位不一致會中止。樞紐分析表刷新失敗時也會中止該工作檔的存檔。執行前仍須確認 OUTSIDETEMP 與 VBA 練習檔的樞紐來源/定義名稱已修復且相符；不一致時先修復範本，不要略過檢查。
The script derives the column count from the source Data headers and checks both working files' headers and table widths before writing. It aborts on a mismatch, and refuses to save a workbook if any PivotTable refresh fails. Confirm the OUTSIDETEMP and VBA practice files' pivot sources and defined names are repaired and aligned before running; fix a mismatch in the working files first.

```powershell
.\Publish-OutsideReport.ps1 -ReportMonth $ReportMonth
```

> ⚠️ 這支腳本會直接改動 `OUTSIDETEMP.xlsx` 和 `OUTSIDETEMP_VBA_PRACTICE.xlsm` 本身（不是唯讀範本）。
> 執行前若偵測到任何 EXCEL.EXE 在跑（包含沒有視窗的背景程序），腳本會直接報錯中止——
> **不會**自動強制關閉。請先自行存檔並關閉活頁簿，再確認沒有殘留程序後重新執行。
>
> ⚠️ This script modifies `OUTSIDETEMP.xlsx` and `OUTSIDETEMP_VBA_PRACTICE.xlsm` directly (not
> read-only templates). If any EXCEL.EXE is detected beforehand (including a background process with no
> window), the script throws and aborts — it will **not** force-close anything. Save and close all
> workbooks, confirm no orphan process remains, then re-run.

若只想同步 OUTSIDETEMP.xlsx、略過 VBA 練習檔：
To sync only OUTSIDETEMP.xlsx and skip the VBA practice file:

```powershell
.\Publish-OutsideReport.ps1 -ReportMonth $ReportMonth -VbaPracticePath ""
```

## 6.5.（一次性/視需要）把 OUTSIDETEMP 算好的 ProdFamiDesc 複製回正式來源檔 / (One-off/as-needed) copy OUTSIDETEMP's computed ProdFamiDesc back into the official source file

只在來源 Excel 的「MHK+STC+MZQ+MIL+RACO Data」分頁還沒有 `ProdFamiDesc` 這欄（或該欄需要重新同步）時才需要執行。會先逐列核對 PO_NO/VEND_INV_NO/ITEM_CODE/REC_QTY 對齊無誤（含前導零正規化），才把值（不是公式）複製進正式來源檔；核對不過會直接中止，不寫入任何資料。
Only needed if the official source Excel's "MHK+STC+MZQ+MIL+RACO Data" tab doesn't yet have a `ProdFamiDesc` column (or it needs re-syncing). Verifies every row aligns by PO_NO/VEND_INV_NO/ITEM_CODE/REC_QTY (with leading-zero normalization) before copying the values (not the formula) into the official source file; aborts without writing anything if any row doesn't match.

```powershell
.\Global Purchase Report RPA\HTML_Report_Generator\Copy-ProdFamiDescToSource.ps1
```

> ⚠️ 這支腳本會直接改動並存檔正式月報來源檔（例如 `FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-(1-7)2026_20260826_RV0.xlsx`），執行前請先確認該檔案未在 Excel 中開啟（若偵測到 EXCEL.EXE 在跑會直接報錯中止）。預設路徑寫死在腳本的 `param()` 區塊，若檔名不同（例如換到下個月的來源檔），需用 `-SourceReportPath` 明確指定。
> ⚠️ This script modifies and saves the official monthly source file directly (e.g. `FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-(1-7)2026_20260826_RV0.xlsx`) — confirm it isn't open in Excel first (throws and aborts if EXCEL.EXE is detected running). Default paths are hardcoded in the script's `param()` block; pass `-SourceReportPath` explicitly if the filename differs (e.g. next month's source file).

## 7. 通知/發布給同事 / Notify/distribute to colleagues

```powershell
.\New-MonthlyEmailDraft.ps1 -ReportMonth $ReportMonth
.\New-MonthlySlackPost.ps1 -ReportMonth $ReportMonth -Channel "C0BSWRJU85T"
```

> ⚠️ Gmail 只會建立**草稿**，絕不自動寄出。**Slack 沒有草稿機制，一執行就會立刻真實發布**到頻道，所有成員馬上看得到。首次使用 Slack 前請先看 `SLACK_SETUP.md` 完成一次性設定。
> ⚠️ Gmail only creates a **draft**, never auto-sends. **Slack has no draft concept — it posts live the moment you run it**, visible to everyone in the channel immediately. See `SLACK_SETUP.md` for the one-time setup before first use.

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
