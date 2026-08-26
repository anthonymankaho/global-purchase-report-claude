# Global Purchase Report — SOP 精簡版 + 逐步指令 / SOP Quick Guide + Step-by-Step Commands

> 完整版 SOP 見 `SOP_Global Purchase Report採購報表處理流程.md`。這份是濃縮版，給熟悉流程後快速對照用。
> Full SOP: `SOP_Global Purchase Report採購報表處理流程.md`. This is the condensed version for quick reference once you know the flow.

---

## Part 1：重點摘要 / Key Points

**這是什麼 What this is：** 每月把 7 個實體（MHK/STC/MZQ/MIL/RACO/TVS/MAZ）的採購資料合併成一份報表，原本是 UiPath RPA，現在改用 PowerShell 重建（僅合併/產出報表這段，資料取得與寄送仍為人工）。
Monthly merge of 7 entities' purchase data into one report. Originally a UiPath RPA process, now rebuilt in PowerShell (merge/report-generation only — data collection and sending stay manual).

**人工 vs 自動 Manual vs Automated：**

| 階段 Stage | 方式 How | 說明 |
|---|---|---|
| 取得匯率表 Get exchange rate | 👤 人工 Manual | Gmail 下載改名放資料夾，見下方 SOP §2.1 |
| 取得7實體當月資料 Get each entity's data | 👤 人工 Manual | EPICOR匯出 或 廠商/財務提供檔案 |
| 合併＋產出報表 Merge + build report | 🤖 PowerShell | 本文件 Part 2 |
| 檢查最終報表 Review final report | 👤 人工 Manual | 打開報表確認 |
| 寄送/上傳 Send/upload | 👤 人工 Manual | Email / MMS，不自動化 Not automated |

**7個實體的資料特性 The 7 entities' data quirks：**

| 實體 | 來源 | 當月篩選欄 |
|---|---|---|
| MHK / STC | EPICOR匯出（已是單月）Already single-month | 不需篩選 None needed |
| MZQ | 廠商累積檔，分頁叫 **"MHK"**（歷史命名）Data sheet is literally named "MHK" | B欄 DATE |
| MIL | 廠商累積檔，分頁 **"MIL"** | B欄 Date |
| RACO | 廠商累積檔，實際只有 **"Sheet1"** 分頁 | B欄 DATE |
| TVS | 財務+Gemini草稿，22欄不是一對一對應 22 cols, not 1:1 mapped | G欄 Invoice Date |
| MAZ | 廠商累積檔，分頁叫 **"MHK"**；缺檔會自動略過不中斷 Missing file auto-skips, doesn't abort | B欄 Purchase Order Date |

**已知限制 Known limitations：**
- MAZ 的 PEN 匯率不在標準匯率表裡 → 用 `Get-PenExchangeRate.ps1` 補（見 Part 2 Step 2）
- TVS 沒有「缺檔自動略過」機制（不像 MAZ）—— TVS 資料沒到位，整個 Build 會失敗

---

## Part 2：逐步 PowerShell 指令 / Step-by-Step PowerShell Commands

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude"
$ReportMonth = "2026-07"   # <- 每月只改這一行 / change only this line each month
```

| # | 指令 Command | 用途 Purpose |
|---|---|---|
| 1 | `.\Build-GlobalPurchaseReport.ps1 -ReportMonth $ReportMonth` | 合併7實體+匯率，產出當月報表 Merge + build this month's report |
| — | 👤 手動複製 `AA2:AB2` 到新增列並計算 Manually copy `AA2:AB2` down + calculate | Max/Min PO Price，下一步的前提 Precondition for next step |
| 2 | `.\Get-PenExchangeRate.ps1 -ReportMonth $ReportMonth` | 補 MAZ 的 PEN 匯率缺口（若當月有 PEN 交易）Fill MAZ's PEN gap (if any PEN rows this month) |
| 3 | `.\Complete-FormulaColumns.ps1 -ReportMonth $ReportMonth` | 填/凍結 AC~AI + 建「New Vendors」複核分頁 Fill/freeze AC–AI + build New Vendors sheet |
| 4 | `.\Update-TVSTranslations.ps1`（如有需要 if needed） | 修正 TVS 翻譯欄與 Commodity Code |
| 5 | `.\New-ReportAuditSummary.ps1 -ReportMonth $ReportMonth` | （建議）唯讀稽核摘要 (Recommended) read-only audit summary |
| 6 | 👤 打開報表「全部重新整理」、檢查匯率/列數、確認無誤 Open report, Refresh All, check rate/row counts | 人工最終檢查 Manual final review |
| 7 | `.\Publish-OutsideReport.ps1 -ReportMonth $ReportMonth` | 產出外部交付檔（OUTSIDETEMP + VBA練習檔）Generate outside deliverables (OUTSIDETEMP + VBA practice file) |
| — | 👤 依公司規定寄送/上傳 Send/upload per company policy | 不自動化 Not automated |

> ⚠️ Step 1 開始會**自動強制關閉**所有 EXCEL.EXE；Step 7 若偵測到 EXCEL.EXE 在跑則**直接報錯中止**（不強制關閉）——兩支腳本行為不同，執行前都建議先自行存檔關閉 Excel。
> ⚠️ Step 1 **force-kills** all EXCEL.EXE at startup; Step 7 **throws and aborts** if any EXCEL.EXE is detected (does not force-close) — different behavior between the two, so save and close Excel yourself before either.

需要完整參數說明（強制重跑、手動指定範本路徑等）請見 `PowerShell-Commands.md`。
For full parameter details (force-rerun, manual template paths, etc.) see `PowerShell-Commands.md`.
