# Global Purchase Report (PowerShell 重建 / PowerShell Rebuild)

將原本由 UiPath RPA 執行的「Global Purchase Report」月報流程，改為人工資料匯出 + PowerShell 自動合併/產出報表。
Rebuilding the monthly "Global Purchase Report" from a UiPath RPA process into manual data export + PowerShell-automated merge/report generation.

**狀態 / Status**: ✅ 已完成並通過端對端測試，可正式使用 (Done, end-to-end tested, ready for regular use)

---

## 這是什麼 / What this is

舊 RPA 專案每月把 MHK、STC、MZQ、MIL、RACO、TVS、MAZAL 七個實體的採購資料，
從各自的 ERP/廠商匯出檔，合併、整理成一份最終報表，寄送並上傳到 Meyer Management System (MMS)。
The old RPA project merges 7 entities' (MHK/STC/MZQ/MIL/RACO/TVS/MAZAL) purchase data from ERP/vendor exports into one final report each month, then emails it and uploads it to MMS.

這次重建的範圍：**只重寫「合併/產出報表」的部分**。ERP 登入/匯出、Email 寄送、MMS 上傳維持人工；匯率表的「取得」是人工（Gmail 下載），但「轉貼進報表」已自動化。理由與作法詳見 [SOP 文件](<SOP_Global Purchase Report採購報表處理流程.md>)。
Scope of this rebuild: **only the merge/report-generation logic**. ERP login/export, email send, and MMS upload stay manual; exchange rate *retrieval* is manual (Gmail download) but *applying it into the report* is automated. See the [SOP doc](<SOP_Global Purchase Report採購報表處理流程.md>) for why and how.

---

## 怎麼用 / How to use it

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude"
.\Build-GlobalPurchaseReport.ps1 -ReportMonth "2026-06"
```

執行前請先完成 SOP 第 2 節的人工準備步驟（匯率表 + 七個實體的當月資料；MAZ 若當月沒收到可以先不放，腳本會自動略過）。執行完畢後會自動填好 Q2:AJ2 公式（一欄一欄填、算完立刻凍結成數值，避免整批重算卡住），並把當月新出現的「NEW VENDOR」列自動整理到「New Vendors」分頁供檢查。詳細步驟、預期畫面輸出、疑難排解都在 SOP 文件裡。
Complete the manual prep steps in SOP §2 first (exchange rate file + the seven entities' monthly data; MAZ can be left out if not received that month — the script skips it automatically). Q2:AJ2 formulas get filled automatically after it runs (one column at a time, frozen to values immediately after each to avoid a mass recalculation hang), and any new "NEW VENDOR" rows for the month are automatically collected into a "New Vendors" sheet for review. Full steps, expected output, and troubleshooting are in the SOP doc.

TVS 若翻譯/商品代碼有異常，另外執行：
If TVS's translation/commodity codes need attention, run separately:
```powershell
.\Update-TVSTranslations.ps1
```

（建議）產出稽核摘要報告（報表最後列號、六實體當月列數、當月總列數），只唯讀開啟，不會改動報表：
(Recommended) generate an audit summary (last data row, per-entity row counts for the month, total row count) — read-only, never modifies the report:
```powershell
.\New-ReportAuditSummary.ps1
```

若需要產出外部交付檔案（把數值複製進 OUTSIDETEMP.xlsx、延伸表格、重新整理樞紐分析表、另存當月交付檔）：
If an outside-delivery deliverable is needed (value-copy into OUTSIDETEMP.xlsx, extend the table, refresh all PivotTables, Save-As the dated deliverable):
```powershell
.\Publish-OutsideReport.ps1 -ReportMonth "2026-06"
```
> ⚠️ 這支腳本會直接改動 `OUTSIDETEMP.xlsx`（累加資料、延伸表格），不是唯讀範本。詳見 SOP 第 4 節步驟 8。
> ⚠️ This script modifies `OUTSIDETEMP.xlsx` directly (accumulates data, extends the table) — it is not a read-only template. See SOP §4 step 8 for details.

---

## 資料夾結構 / Folder structure

```
Global Purchase Report Claude/
├── README.md                                          ← 本文件
├── SOP_Global Purchase Report採購報表處理流程.md         ← 標準作業程序（人工步驟 + 執行步驟 + 已知限制）
├── Build-GlobalPurchaseReport.ps1                      ← 主執行腳本
├── Update-TVSTranslations.ps1                          ← TVS 翻譯/商品代碼檢查與修補（獨立腳本）
├── New-ReportAuditSummary.ps1                          ← 稽核摘要（唯讀，最後列號 + 七實體列數 + 總列數）
├── Publish-OutsideReport.ps1                           ← 外部交付檔案：複製數值進 OUTSIDETEMP、延伸表格、重新整理樞紐分析表、另存當月交付檔
├── Global Purchase Report RPA/                         ← 舊 RPA 專案的只讀複本，供分析參考（腳本不會寫入此資料夾）
├── Exchange Rate/                                      ← 每月放 STANDARD EXCHANGE RATE.xls[x]（人工取得，腳本自動讀取）
├── MHK/ STC/ MZQ/ MIL/ RACO/ TVS/ MAZ/                  ← 各實體當月原始資料（人工放入）+ 腳本產出的月份快照
└── REPORT/                                             ← 腳本產出的每月最終報表
```

- `MHK`/`STC` 資料來源：EPICOR 入口網站 ACT001 匯出（人工登入、篩選當月、匯出）。
- `MZQ`/`MIL`/`RACO`/`TVS`/`MAZ` 資料來源：廠商或財務部提供的累積檔（Jan–當月），腳本自動篩出當月資料。
- `MAZ`：若某月尚未收到資料，腳本會自動略過該實體並印出警告，不會中斷整個執行。

---

## 目前進度 / Current progress

| 項目 / Item | 狀態 / Status |
|---|---|
| 七個實體欄位對應（MHK/STC/MZQ/MIL/RACO/TVS/MAZ） | ✅ 全部完成並實測驗證（2026-06：6,816 列、0 筆公式錯誤）|
| TVS/MAZ 欄位對應 | ✅ 已用「比對舊 Report 已知正確列 vs 原始檔」的方式實測驗證（原沒有 RPA 自動化區段可參考）|
| PowerShell 主腳本 | ✅ 完成，含自動接續（免每月改參數）與重複執行安全檢查 |
| 端對端測試 | ✅ 通過 |
| 匯率表自動化 | ✅ 已完成（C:M 直接寫源頭值，O:Z 公式自動計算，不再靠複製上一列）|
| Q2:AJ2 公式填入新增列 | ✅ 已自動化（一欄一欄依相依順序填公式、算完立刻凍結成數值）|
| AA/AB（Max/Min PO Price） | ✅ 改用 PowerShell 分組計算，取代會隨資料量變慢的 MAXIFS/MINIFS |
| New Vendors 檢查分頁 | ✅ 已完成（自動彙整當月「NEW VENDOR」列）|
| RACO/TVS 樣板存檔衝突 | ✅ 已修正（改用固定範本檔，不再自動抓最新快照）|
| MIL 的 PART TYPE 欄位 | ✅ 已修正（數字 1-9 自動轉成 "B"，修好 AG 欄 #N/A 的問題）|
| 稽核摘要工具 | ✅ 已完成（`New-ReportAuditSummary.ps1`）|
| 外部交付檔案流程（OUTSIDETEMP → 當月交付檔） | ✅ 已完成（`Publish-OutsideReport.ps1`：數值複製、延伸 Table1、重新整理樞紐分析表、另存新檔，含自動遞增 RV 編號）|
| MAZ 的 PEN 匯率缺口 | 待處理（`STANDARD EXCHANGE RATE` 來源檔沒有 PEN 匯率）|
| OUTSIDETEMP.xlsx 的 `data` 命名範圍已損壞 | 待處理（既有問題，非本次造成；影響 3 個隱藏分頁共 18 個樞紐分析表）|

---

## 與其他專案的關係 / Related projects

- **MTC Purchase Report Claude**（`Desktop\MTC Purchase Report Claude\`）：同類型的先行專案，證實了「EPICOR 網站無可直接呼叫 API、瀏覽器自動化匯出會導致 Excel 殖屍程序」，本專案沿用其「人工匯出 + PowerShell 合併」的策略，但程式碼為重新撰寫，未共用其 `Common.ps1`。
- **Weekly Exchange Rate PowerShell Rebuild**：另一個獨立的匯率報表 RPA 遷移專案，範圍與作法不同（該專案的匯率來源有直接可呼叫的 API）。
