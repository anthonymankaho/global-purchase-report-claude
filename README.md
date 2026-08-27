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

執行前請先完成 SOP 第 2 節的人工準備步驟（匯率表 + 七個實體的當月資料；MAZ 若當月沒收到可以先不放，腳本會自動略過）。執行完畢後會自動填好 Q:S 跟 U/V/W/Y/Z/AJ 公式（一欄一欄填、算完立刻凍結成數值，避免整批重算卡住）。詳細步驟、預期畫面輸出、疑難排解都在 SOP 文件裡。
Complete the manual prep steps in SOP §2 first (exchange rate file + the seven entities' monthly data; MAZ can be left out if not received that month — the script skips it automatically). Q:S and U/V/W/Y/Z/AJ formulas get filled automatically after it runs (one column at a time, frozen to values immediately after each to avoid a mass recalculation hang). Full steps, expected output, and troubleshooting are in the SOP doc.

⚠️ **2026-08-04 起，AA/AB（Max/Min PO Price）改為全程人工處理**：主腳本執行完後，先手動把 `AA2:AB2` 複製到新增列並計算確認，再執行：
⚠️ **As of 2026-08-04, AA/AB (Max/Min PO Price) are handled entirely manually**: after the main script, manually copy `AA2:AB2` down to the new rows and calculate/confirm first, then run:
```powershell
.\Complete-FormulaColumns.ps1 -ReportMonth "2026-06"
```
這支腳本會自動找出當月新增列範圍、確認 AA 已經人工填好，才接著填 AC/AD/AE/AF/AG/AH/AI 並凍結、建立「New Vendors」檢查分頁。詳見 SOP 第4節步驟3.5/3.6。
This script auto-detects the month's new rows, confirms AA was filled manually, then fills/freezes AC/AD/AE/AF/AG/AH/AI and builds the "New Vendors" review sheet. See SOP §4 steps 3.5/3.6.

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

若當月 MAZ 有 PEN 幣別交易，需先補匯率缺口（見 SOP 第 5 節）：
If this month's MAZ has PEN-currency rows, fill the rate gap first (see SOP §5):
```powershell
.\Get-PenExchangeRate.ps1 -ReportMonth "2026-07"
```

若需要產出外部交付檔案（把數值複製進 OUTSIDETEMP.xlsx 和 OUTSIDETEMP_VBA_PRACTICE.xlsm、延伸表格、重新整理樞紐分析表、另存當月交付檔）：
If an outside-delivery deliverable is needed (value-copy into OUTSIDETEMP.xlsx and OUTSIDETEMP_VBA_PRACTICE.xlsm, extend the tables, refresh all PivotTables, Save-As the dated deliverable):
```powershell
.\Publish-OutsideReport.ps1 -ReportMonth "2026-07"
```
> ⚠️ 這支腳本會直接改動 `OUTSIDETEMP.xlsx` 和 `OUTSIDETEMP_VBA_PRACTICE.xlsm`（累加資料、延伸表格），不是唯讀範本；執行前若偵測到任何 EXCEL.EXE 在跑會直接報錯中止（不強制關閉）。詳見 SOP 第 4 節步驟 8。
> ⚠️ This script modifies `OUTSIDETEMP.xlsx` and `OUTSIDETEMP_VBA_PRACTICE.xlsm` directly (accumulates data, extends the tables) — not read-only templates; it throws and aborts if any EXCEL.EXE is detected running (does not force-close). See SOP §4 step 8 for details.

確認無誤後，用兩支獨立腳本通知/發布給相關同事（Gmail 草稿 + Slack 頻道貼文，詳見 SOP 第 9 節）：
Once confirmed, notify/distribute to colleagues via two independent scripts (Gmail draft + Slack channel post, see SOP §9 for details):
```powershell
.\New-MonthlyEmailDraft.ps1 -ReportMonth "2026-07"
.\New-MonthlySlackPost.ps1 -ReportMonth "2026-07" -Channel "C0BSWRJU85T"
```
> ⚠️ Gmail 只建立草稿，絕不自動寄出；**Slack 沒有草稿機制，執行後立刻真實發布到頻道**，第一次使用前請先看 `SLACK_SETUP.md` 完成一次性設定。
> ⚠️ Gmail only creates a draft, never auto-sends; **Slack has no draft concept — it posts live to the channel immediately**. See `SLACK_SETUP.md` for the one-time setup before first use.

完整逐步指令請見 `PowerShell-Commands.md` 或 `SOP-Quick-Guide.md`。
For the full step-by-step command list, see `PowerShell-Commands.md` or `SOP-Quick-Guide.md`.

---

## 資料夾結構 / Folder structure

```
Global Purchase Report Claude/
├── README.md                                          ← 本文件
├── SOP_Global Purchase Report採購報表處理流程.md         ← 標準作業程序（人工步驟 + 執行步驟 + 已知限制）
├── Build-GlobalPurchaseReport.ps1                      ← 主執行腳本（AA/AB 及其依賴欄位除外，見下）
├── Complete-FormulaColumns.ps1                         ← 手動填好 AA/AB 之後執行：AC/AD/AE/AF/AG/AH/AI + New Vendors 分頁
├── Update-TVSTranslations.ps1                          ← TVS 翻譯/商品代碼檢查與修補（獨立腳本）
├── New-ReportAuditSummary.ps1                          ← 稽核摘要（唯讀，最後列號 + 七實體列數 + 總列數）
├── Get-PenExchangeRate.ps1                             ← 補 MAZ 的 PEN 匯率缺口（抓 USD/PEN，換算寫入 Exchange Rate 分頁 M欄）
├── Publish-OutsideReport.ps1                           ← 外部交付檔案：複製數值進 OUTSIDETEMP.xlsx + OUTSIDETEMP_VBA_PRACTICE.xlsm、延伸表格、重新整理樞紐分析表、另存當月交付檔
├── New-MonthlyEmailDraft.ps1                           ← Gmail 草稿通知（比照舊RPA格式與收件人清單，只建草稿不寄出）
├── New-MonthlySlackPost.ps1                            ← Slack 頻道貼文（含互動式HTML報表附件，即時發布無草稿機制）
├── Save-SlackToken.ps1                                 ← 一次性設定：加密儲存 Slack Bot Token
├── SLACK_SETUP.md                                      ← Slack App 一次性設定步驟
├── GmailOAuth.Common.ps1                                ← Gmail OAuth 共用函式（沿用 Weekly Exchange Report 專案的授權）
├── Credentials/                                        ← OAuth/Bot Token（DPAPI 加密，已加入 .gitignore，不會被提交）
├── PowerShell-Commands.md                              ← 全部指令速查（`$ReportMonth` 變數，每月僅需改一行）
├── PowerShell-Runtime-Report.md                        ← 各腳本執行時間記錄
├── SOP-Quick-Guide.md                                  ← SOP 精簡版重點摘要 + 逐步指令表
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
| 外部交付檔案流程（OUTSIDETEMP → 當月交付檔） | ✅ 已完成（`Publish-OutsideReport.ps1`：數值複製、延伸 Table1、重新整理樞紐分析表、另存新檔，含自動遞增 RV 編號；2026-08-26 起同步擴充到 `OUTSIDETEMP_VBA_PRACTICE.xlsm`）|
| MAZ 的 PEN 匯率缺口 | ✅ 已完成（`Get-PenExchangeRate.ps1`：抓 USD/PEN 換算成 HKD/PEN 寫入，2026-07 月結已實測使用成功）|
| OUTSIDETEMP.xlsx 的 `data` 命名範圍已損壞 | ✅ 已修復（2026-08-25，人工手動重建：`data` 改指向 `=Table1[#All]`；3 個隱藏分頁共 18 個樞紐分析表已實測全部重新整理成功）|
| OUTSIDETEMP_VBA_PRACTICE.xlsm 的 `data` 命名範圍損壞/缺失 | ✅ 已修復（2026-08-27，人工在名稱管理員新增 `data` 指向 `=Table1[#All]`；25個樞紐分析表已實測全部重新整理成功）|
| 2026-07 月結 | ✅ 7個實體皆已用實際日期欄位驗證（非僅檔案時間戳記），主報表與交付檔均已產出 |
| 每月通知/發布自動化 | ✅ 已完成（`New-MonthlyEmailDraft.ps1` Gmail草稿 + `New-MonthlySlackPost.ps1` Slack頻道貼文含HTML報表附件，2026-08-27 實測成功，見 SOP 第9節）|
| 第2列母版公式效能優化（2026-08） | ✅ 已完成，U/V/W/AA/AB/AG/AH 改有界範圍；修正 ITEM_CODE 前導零不一致造成的 AG/AH #N/A（用 `LET()` 正規化）；完整公式表見 SOP 第3.1節 |
| AA/AB 改為全程人工處理（2026-08-04）| ✅ 已完成：`Build-GlobalPurchaseReport.ps1` 不再碰 AA/AB 及依賴欄位；新增 `Complete-FormulaColumns.ps1`，人工填好 AA/AB 後執行，自動檢查沒填會報錯中止，避免重演 key 不同步 #N/A 問題 |

---

## 與其他專案的關係 / Related projects

- **MTC Purchase Report Claude**（`Desktop\MTC Purchase Report Claude\`）：同類型的先行專案，證實了「EPICOR 網站無可直接呼叫 API、瀏覽器自動化匯出會導致 Excel 殖屍程序」，本專案沿用其「人工匯出 + PowerShell 合併」的策略，但程式碼為重新撰寫，未共用其 `Common.ps1`。
- **Weekly Exchange Rate PowerShell Rebuild**：另一個獨立的匯率報表 RPA 遷移專案，範圍與作法不同（該專案的匯率來源有直接可呼叫的 API）。

---

## HTML 互動版報表 / Interactive HTML report

`Global Purchase Report RPA\HTML_Report_Generator\` 底下的獨立工具，把每月的最終報表（`Global Purchase Report RPA\REPORT\` 裡的 `FIN_..._PUR_REPORT-(x-y)YYYY_YYYYMMDD_RVn.xlsx`）轉成單一自包含的互動式 HTML 檔案（Dashboard 圖表 + 篩選器 + 淺/深色模式；GLID/Fittings/Packaging/Coating 改成可依 Month/DB/Std Comm/Std Vendor Name/Item Code/Ignore-Consider 即時篩選的明細表，同料號多筆 DB/Vendor 資料可用 +/- 按鈕收合；Exchange Rate 改成跟 Data 頁一樣可排序/搜尋/匯出 CSV 的互動表格，只保留「Platium Buy Rate ⇒ USD」）。篩選是即時從 Data 分頁原始交易資料算出來的，不是讀 Excel 樞紐分析表快取，所以篩選會同時連動圖表跟明細表。來源檔案自動偵測是依檔名內編碼的報表期間排序，不是依檔案修改時間（避免檔案被開啟過就被誤判成「最新」）。

A standalone tool under `Global Purchase Report RPA\HTML_Report_Generator\` that converts the month's final report (`FIN_..._PUR_REPORT-(x-y)YYYY_YYYYMMDD_RVn.xlsx` in `Global Purchase Report RPA\REPORT\`) into a single self-contained interactive HTML file (Dashboard with charts + filters + light/dark mode; GLID/Fittings/Packaging/Coating rebuilt as item-level tables that respond live to Month/DB/Std Comm/Std Vendor Name/Item Code/Ignore-Consider filters, with a +/- toggle to collapse an item's multiple DB/Vendor rows down to one; Exchange Rate rebuilt as a sortable/searchable/CSV-exportable table like the Data tab, keeping only "Platium Buy Rate ⇒ USD"). Filters are computed live from the raw Data-sheet transactions rather than Excel's pivot cache, so they drive both the charts and the table together. Source-file auto-detection sorts by the reporting period encoded in the filename, not file modification time (so a file merely being opened doesn't get mistaken for "latest").

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude\Global Purchase Report RPA\HTML_Report_Generator"
.\Generate-PurchaseReportHtml.ps1
```

不帶參數時預設寫到固定的暫存檔 `FIN_Global_Purchase -Report_html_TEMP.html`——可以放心重複執行、反覆測試，不會動到正式的當月快照。要確認這次是要定案發布時，加上 `-Publish`，會另外把結果複製一份成當月報表期間命名的快照（如 `(1-6)2026` → `FIN_Global_Purchase -Report_html_202606.html`），存回同一個 `REPORT\` 資料夾。也可用 `-SourceExcelPath` / `-OutputPath` 指定特定檔案。（沿用專案既有的「TEMP 工作檔 + 定案才產出正式快照」慣例，避免每次測試都覆寫正式檔案。）

With no parameters it defaults to a fixed working file, `FIN_Global_Purchase -Report_html_TEMP.html` — safe to re-run any number of times while testing without touching the dated monthly snapshot. Add `-Publish` when you're actually finalizing this run: it additionally copies the result to a period-stamped snapshot (e.g. `(1-6)2026` → `FIN_Global_Purchase -Report_html_202606.html`) in the same `REPORT\` folder. Pass `-SourceExcelPath` / `-OutputPath` to target a specific file instead. (Follows this project's existing TEMP-working-file + publish-on-finalize convention, so test runs never overwrite the real snapshot.)

> ⚠️ 目前**刻意沒有**自動接在 `Build-GlobalPurchaseReport.ps1` 後面執行：AA/AB（Max/Min PO Price）要等人工填好、確認過才算報表定稿（見上方「怎麼用」與 SOP），自動串接的話有可能拿到還沒補完的報表去產生 HTML。等確認要接在月結流程的哪一步之後，再考慮自動化串接。
> This is **deliberately not** auto-chained onto `Build-GlobalPurchaseReport.ps1`: the report isn't final until AA/AB (Max/Min PO Price) are filled in manually and confirmed (see "How to use" above and the SOP) — auto-chaining risks generating HTML from an unfinished report. Revisit once it's clear exactly where in the monthly sequence this should trigger.

**（實驗性）Excel 按鈕觸發**：`HTML_Report_Generator\HtmlUpdateModule.bas` 是一份可手動匯入 Excel VBA 的巨集範本，讓使用者能用工作表上的一個按鈕觸發本腳本，已在沙盒環境驗證可行（不影響使用者原本開著的 Excel）。要正式使用需要把承載按鈕的檔案另存成 `.xlsm`，且因每月報表檔名不同、巨集不會自動延續到下個月，加上公司巨集安全政策是否允許仍待確認——尚未正式部署，詳見 SOP 第 8 節。
**(Experimental) Excel-button trigger**: `HTML_Report_Generator\HtmlUpdateModule.bas` is a manually-importable VBA macro that lets a worksheet button trigger this script — sandbox-verified as safe (doesn't conflict with the user's own open Excel session). Real use requires saving the host file as `.xlsm`, manually re-importing the macro each month (filenames change), and confirming your organization's macro-security policy allows it — not yet deployed for real use. See SOP §8 for details.
