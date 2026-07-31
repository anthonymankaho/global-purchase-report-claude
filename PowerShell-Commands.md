# PowerShell 指令速查 / PowerShell Commands Quick Reference

以下是 SOP 和 README 裡列出的所有 PowerShell 指令：
Here are all the PowerShell prompts documented in the SOP and README:

> ⚠️ **最容易打錯的地方**：指令開頭一定要有 `.\`（一個點 + 反斜線），代表「目前資料夾」。
> 只打 `\腳本名稱`（斜線開頭，沒有點）或 `腳本名稱`（完全沒有前綴）都會出現
> `無法辨識...詞彙是否為 Cmdlet` 的錯誤，因為 PowerShell 預設不會從目前資料夾載入指令。
>
> ⚠️ **Most common mistake**: the command must start with `.\` (a dot + backslash), meaning "current folder".
> Typing `\ScriptName` (leading slash, no dot) or just `ScriptName` (no prefix at all) will both fail with
> a "not recognized as a Cmdlet" error, since PowerShell doesn't load commands from the current folder by default.

## 每月主要執行 / Monthly main run

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude"
.\Build-GlobalPurchaseReport.ps1 -ReportMonth "2026-06"
```

> Q2:AJ2 公式填入、New Vendors 檢查分頁都會自動完成，不需要額外人工步驟。
> Q2:AJ2 formula fill and the New Vendors review sheet both happen automatically — no extra manual step needed.

## TVS 翻譯/商品代碼檢查（如有需要）/ TVS translation/commodity check (if needed)

```powershell
.\Update-TVSTranslations.ps1
```

## （建議）產出稽核摘要報告 / (Recommended) generate an audit summary

只唯讀開啟報表，不會改動任何東西：報表最後列號、七實體當月列數、當月總列數。
Read-only, never modifies the report: last data row, per-entity row counts for the month, total row count.

```powershell
.\New-ReportAuditSummary.ps1
```

指定特定月份查核：
To check a specific month:
```powershell
.\New-ReportAuditSummary.ps1 -ReportMonth "2026-07"
```

## 產出外部交付檔案（OUTSIDETEMP → 當月交付檔）/ Generate the outside-delivery deliverable

把當月報表 Data 分頁的數值複製進 OUTSIDETEMP.xlsx、延伸其 Table1 表格範圍、重新整理所有樞紐分析表，
存檔後另存為當月交付檔（`FIN_..._PUR_REPORT-(1-M)YYYY_YYYYMMDD_RV0.xlsx`，同一天重跑會自動遞增 RV 編號）。
Value-copies the month's report Data tab into OUTSIDETEMP.xlsx, extends its Table1 range, refreshes
all PivotTables, saves it, then Save-As's the dated deliverable (auto-bumps the RV number if rerun
the same day).

```powershell
.\Publish-OutsideReport.ps1 -ReportMonth "2026-06"
```

> ⚠️ 這支腳本會直接改動 `Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-OUTSIDETEMP.xlsx`
> 本身（不是唯讀範本）。執行前建議先手動備份一份。三個隱藏分頁（Comm/Vend Name/Desc Analysis）裡
> 共 18 個樞紐分析表會重新整理失敗，這是既有的 `data` 命名範圍損壞問題，非本腳本造成，詳見 SOP 第 5 節。
>
> ⚠️ This script modifies `Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-OUTSIDETEMP.xlsx`
> directly (not a read-only template) - back it up manually before running, especially the first time.
> The 18 PivotTables across 3 hidden sheets (Comm/Vend Name/Desc Analysis) will fail to refresh due to
> a pre-existing broken `data` defined name - not caused by this script. See SOP §5.

## 若要強制重跑同一個月（略過重複執行安全檢查）/ To force-rerun the same month (bypass the duplicate-run safety check)

```powershell
.\Build-GlobalPurchaseReport.ps1 -ReportMonth "2026-06" -Force
```

## 若自動偵測失敗、需要手動指定路徑 / If auto-detection fails and paths need to be specified manually

RACO/TVS 樣板預設固定抓 `Trimmed Program\RACO(或 TVS) Act001 Trimmed.xlsx`（未命名日期的範本檔），一般不需要手動指定；只有在範本檔被搬到別的地方時才需要：
RACO/TVS templates default to the fixed `Trimmed Program\RACO(or TVS) Act001 Trimmed.xlsx` (the undated template) — you normally don't need to override this; only needed if the template file has been moved elsewhere:

```powershell
.\Build-GlobalPurchaseReport.ps1 `
  -ReportMonth "2026-06" `
  -RacoTrimmedFile "C:\path\to\RACO Act001 Trimmed.xlsx" `
  -TvsTrimmedFile "C:\path\to\TVS Act001 Trimmed.xlsx"
```

## 一次性事前準備（僅需設定一次）/ One-time setup (only needed once)

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```
