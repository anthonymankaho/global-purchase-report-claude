# Global Purchase Report (PowerShell 重建 / PowerShell Rebuild)
Powershell Main prompt cd "C:\Users\User\Desktop\Global Purchase Report Claude"

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

## 月初資料索取 / Requesting monthly data from inter-group entities

在動手跑主流程之前，MZQ/MIL/RACO/TVS/MAZ 五個實體的當月採購資料通常需要先開口跟對方要。這支腳本會分別用五種語言（簡體中文/英文/義大利文/西班牙文）建立 5 封獨立的 Gmail 草稿，收件人、稱呼、語言各自對應該實體聯絡人，主旨/內文範本一致但已翻譯成自然道地的商務用語（非逐字翻譯）：
Before running the main pipeline, the monthly purchase data from the five entities MZQ/MIL/RACO/TVS/MAZ usually needs to be requested first. This script creates 5 separate Gmail drafts, one per entity, each in that contact's language (Simplified Chinese/English/Italian/Spanish), with the same subject/body template translated into natural business phrasing (not literal word-for-word):

```powershell
.\New-InterGroupPurchaseDataRequestDrafts.ps1 -ReportMonth "2026-09"
```

`-ReportMonth` 是要索取的採購資料月份；例如 `2026-09` 會要求 2026 年 9 月的資料。加 `-Only MZQ` 可只建立單一草稿方便測試。Cc 名單（Fiona Yip、Alex Wong、Steven Cheng）固定套用到全部 5 封，可用 `-Cc` 覆寫。只會呼叫 Gmail 的 `drafts.create`，絕不自動寄出。
`-ReportMonth` is the month of purchase data to request; for example, `2026-09` requests September 2026 data. Add `-Only MZQ` to create just one draft for testing. The Cc list (Fiona Yip, Alex Wong, Steven Cheng) applies to all 5 by default, override with `-Cc`. Only ever calls Gmail's `drafts.create`, never auto-sends.

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
$ReportMonth = "2026-09"  # 每月改成 YYYY-MM / set to the reporting month
.\Get-PenExchangeRate.ps1 -ReportMonth $ReportMonth
```

若 UBC 匯率網站無法解析 DNS，腳本會停止且不改報表；確認網路/DNS/VPN 後重試。仍無法連線時，可用可信來源核對的 USD per 1 PEN 匯率、日期與來源以 `-UsdPerPen`、`-RateDate`、`-RateSource` 傳入。詳見 PowerShell-Commands.md。
If DNS cannot resolve the UBC rate site, the script stops without changing the report. Check internet/DNS/VPN and retry; if it remains unavailable, pass a verified USD per 1 PEN quote, date and source using `-UsdPerPen`, `-RateDate` and `-RateSource`. See PowerShell-Commands.md.

若需要產出外部交付檔案（把數值複製進 OUTSIDETEMP.xlsx 和 OUTSIDETEMP_VBA_PRACTICE.xlsm、延伸表格、重新整理樞紐分析表、另存當月交付檔）：
If an outside-delivery deliverable is needed (value-copy into OUTSIDETEMP.xlsx and OUTSIDETEMP_VBA_PRACTICE.xlsm, extend the tables, refresh all PivotTables, Save-As the dated deliverable):
```powershell
.\Publish-OutsideReport.ps1 -ReportMonth "2026-07"
```
> ⚠️ 這支腳本會直接改動 `OUTSIDETEMP.xlsx` 和 `OUTSIDETEMP_VBA_PRACTICE.xlsm`（累加資料、延伸表格），不是唯讀範本；執行前若偵測到任何 EXCEL.EXE 在跑會直接報錯中止（不強制關閉）。詳見 SOP 第 4 節步驟 8。
> ⚠️ This script modifies `OUTSIDETEMP.xlsx` and `OUTSIDETEMP_VBA_PRACTICE.xlsm` directly (accumulates data, extends the tables) — not read-only templates; it throws and aborts if any EXCEL.EXE is detected running (does not force-close). See SOP §4 step 8 for details.
>
> 🔧 **交付腳本核對狀態**：2026-10-27 修正 Publish-OutsideReport.ps1 的固定36欄限制，改由來源表頭判斷欄數；寫入前核對兩份目的檔的表頭與 Table 欄數，並在樞紐刷新失敗時中止該檔存檔。程式尚未重新執行；2026-08唯讀檢視發現兩份工作檔的樞紐來源/定義名稱不一致，修復並驗證前仍不可用於正式交付。
> 🔧 **Outside-report script status**: On 2026-10-27, Publish-OutsideReport.ps1 was revised to derive column count from source headers, validate both destination headers/table widths before writing, and abort saving a workbook if pivot refresh fails. The revised script has not been rerun. An August 2026 read-only review found inconsistent pivot sources/defined names in the two working files; repair and verify them before using this script for a formal delivery.

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
├── New-InterGroupPurchaseDataRequestDrafts.ps1         ← 月初向 MZQ/MIL/RACO/TVS/MAZ 索取當月採購資料（5 封多語言 Gmail 草稿，只建草稿不寄出）
├── New-MonthlyEmailDraft.ps1                           ← Gmail 草稿通知（比照舊RPA格式與收件人清單，只建草稿不寄出）
├── New-MonthlySlackPost.ps1                            ← Slack 頻道貼文（含互動式HTML報表附件，即時發布無草稿機制）
├── Save-SlackToken.ps1                                 ← 一次性設定：加密儲存 Slack Bot Token
├── SLACK_SETUP.md                                      ← Slack App 一次性設定步驟
├── GmailOAuth.Common.ps1                                ← Gmail OAuth 共用函式（沿用 Weekly Exchange Report 專案的授權）
├── Credentials/                                        ← OAuth/Bot Token（DPAPI 加密，已加入 .gitignore，不會被提交）
├── PowerShell-Commands.md                              ← 全部指令速查（`$ReportMonth` 變數，每月僅需改一行）
├── PowerShell-Runtime-Report.md                        ← 各腳本執行時間記錄
├── SOP-Quick-Guide.md                                  ← SOP 精簡版重點摘要 + 逐步指令表
├── Global Purchase Report RPA/                         ← 舊 RPA 專案複本 + 對外交付流程的實際工作區（見下方 ⚠️ 說明，並非單純唯讀參考）
│   └── REPORT/                                         ← OUTSIDETEMP.xlsx、OUTSIDETEMP_VBA_PRACTICE.xlsm、每月日期版交付檔、HTML 報表都存在這裡
├── 3D Earth Dashboard/                                 ← （2026-10-29 新增）單一自包含 HTML 的「全球採購戰情室」3D 大螢幕看板＋build/ 重建原料，見該資料夾 README
├── Exchange Rate/                                      ← 每月放 STANDARD EXCHANGE RATE.xls[x]（人工取得，腳本自動讀取）
├── MHK/ STC/ MZQ/ MIL/ RACO/ TVS/ MAZ/                  ← 各實體當月原始資料（人工放入）+ 腳本產出的月份快照
└── REPORT/                                             ← 根目錄，Build-GlobalPurchaseReport.ps1 產出的每月主報表（與上面的巢狀 REPORT/ 是兩個不同資料夾，見下方說明）
```

- `MHK`/`STC` 資料來源：EPICOR 入口網站 ACT001 匯出（人工登入、篩選當月、匯出）。
- `MZQ`/`MIL`/`RACO`/`TVS`/`MAZ` 資料來源：廠商或財務部提供的累積檔（Jan–當月），腳本自動篩出當月資料。
- `MAZ`：若某月尚未收到資料，腳本會自動略過該實體並印出警告，不會中斷整個執行。

> ⚠️ **重要：專案裡有兩個用途不同的 REPORT 資料夾。** 根目錄 REPORT 放主報表；Global Purchase Report RPA\REPORT 放 OUTSIDETEMP、日期版 Excel 交付檔及 HTML。HTML 產生器只讀巢狀資料夾中符合命名規則的日期版 Excel，按報表期間、檔名日期及 RV 版本排序，不讀根目錄報表或 OUTSIDETEMP。沒有當月日期版檔時，HTML 仍可能使用舊月份資料。
> ⚠️ **There are two REPORT folders.** Root REPORT holds the master. Global Purchase Report RPA\REPORT holds OUTSIDETEMP, dated Excel deliverables and HTML. The HTML generator reads matching dated Excel files there and sorts by reporting period, filename date and revision. Ensure the deliverable is current before generating HTML.

> **2026-10-27 saved-file check:** master and delivery Data sheets have 39 columns and 50,262 rows; August has 6,298. Master, dated Excel, VBA practice workbook and HTML contain Jan–Aug 2026. TEMP HTML matches the published snapshot byte for byte. Pivot refresh was not tested.

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
| AA/AB（Max/Min PO Price） | ✅ 2026-08-04起由承辦人複製母版公式並檢查；主腳本不計算此欄。Complete-FormulaColumns.ps1 於確認後填入依賴欄位。|
| New Vendors 檢查分頁 | ✅ 已完成（自動彙整當月「NEW VENDOR」列）|
| RACO/TVS 樣板存檔衝突 | ✅ 已修正（改用固定範本檔，不再自動抓最新快照）|
| MIL 的 PART TYPE 欄位 | ✅ 已修正（數字 1-9 自動轉成 "B"，修好 AG 欄 #N/A 的問題）|
| 稽核摘要工具 | ✅ 已完成（`New-ReportAuditSummary.ps1`）|
| 外部交付檔案流程（OUTSIDETEMP → 當月交付檔） | ⚠️ Publish-OutsideReport.ps1 已改為動態來源欄數並加上表頭/Table寬度及樞紐刷新失敗檢查；尚未重跑。先修復並驗證兩份工作檔的樞紐來源/定義名稱，再用腳本發布。|
| MAZ 的 PEN 匯率缺口 | ✅ Get-PenExchangeRate.ps1 可補匯率；若建置時缺 PEN，另須在 Excel 重填/重算並凍結受影響列 Q:S，單純更新匯率表不會更新資料列。|
| External workbook pivot sources | August dated workbook has data = #REF!; its pivot caches refer to Table1, 表格1 and data, while its Data table is named 表格1. VBA practice workbook has no data name; caches mix A1:AK/A1:AM. These saved records differ from earlier repair notes. Excel refresh was not tested; verify before use. |
| VBA 練習檔定義名稱及樞紐快取 | 曾於2026-08人工修復，但2026-10唯讀檢視的檔案沒有 data 名稱且快取來源混用 A1:AK/A1:AM。Excel 刷新尚未重新驗證；詳見前項。|
| 2026-07 月結 | ✅ 7個實體皆已用實際日期欄位驗證（非僅檔案時間戳記），主報表與交付檔均已產出 |
| 每月通知/發布自動化 | ✅ 已完成（`New-MonthlyEmailDraft.ps1` Gmail草稿 + `New-MonthlySlackPost.ps1` Slack頻道貼文含HTML報表附件，2026-08-27 實測成功，見 SOP 第9節）|
| Product-family fields | August 2026 HTML has ProdFamiDesc in 347/454 MAZ rows and blank in all 3,274 TVS rows. The earlier note that all MAZ rows are blank is outdated. Other fields/months were not checked. Publish-OutsideReport.ps1 now derives source width dynamically; the revised publisher has not been rerun. |
| August 2026 saved state | Master, dated delivery and HTML cover Jan–Aug: 50,262 total rows and 6,298 for August. TEMP HTML matches the 202608 snapshot. Excel pivot caches and defined names still need review. |
| HTML 報表新增功能（2026-08-28～09-02；2026-09-09 更新） | ✅ Dashboard/GLID/Fittings/Data 新增 ProdFamiDesc 與 Part_MfgCat_Code 篩選；Dashboard 新增 SKU vs Part 家族分布圓餅圖（Top10+Other）；GLID/Fittings/Packaging/Coating 四頁明細表欄位可拖曳調整寬度（設定存在瀏覽器本機），且 2026-09-09 起預設欄寬會依使用者目前電腦/瀏覽器視窗可用寬度自動縮放，盡量讓四頁表格在同一畫面內看到全部欄位；Dashboard/GLID/Fittings/Packaging/Coating 的 ITEM_CODE/DB/Vendor 儲存格與 Dashboard 分類/供應商列都可點擊，直接跳到 Data 頁並套用篩選；KPI 卡片拆成 SKU/Part 兩區並各自顯示加權百分比；歡迎彈窗（intro modal）內容已更新並改用新版 localStorage 鍵值，舊使用者也會再看到一次 |
| 第2列母版公式效能優化（2026-08） | ✅ 已完成，U/V/W/AA/AB/AG/AH 改有界範圍；修正 ITEM_CODE 前導零不一致造成的 AG/AH #N/A（用 `LET()` 正規化）；完整公式表見 SOP 第3.1節 |
| AA/AB 改為全程人工處理（2026-08-04）| ✅ 已完成：`Build-GlobalPurchaseReport.ps1` 不再碰 AA/AB 及依賴欄位；新增 `Complete-FormulaColumns.ps1`，人工填好 AA/AB 後執行，自動檢查沒填會報錯中止，避免重演 key 不同步 #N/A 問題 |
| 月初資料索取自動化（2026-09-08） | ✅ 已完成：`New-InterGroupPurchaseDataRequestDrafts.ps1`，向 MZQ/MIL/RACO/TVS/MAZ 建立 5 封多語言 Gmail 草稿索取當月採購資料，`-ReportMonth` 自動往前推一個月 |
| 3D 地球儀互動看板（2026-10-29；2026-10-29 改名，去除「戰情室」字樣，正式名稱「全球採購報告 / Global Purchase Report」） | ✅ 已完成並端對端驗證：`3D Earth Dashboard\Global_Purchase_3D_WarRoom_TEMP.html`（約 14.4 MB 單一自包含 HTML，three.js 3D 地球＋KPI＋實體/品類/供應商下鑽＋明細/CSV＋**「中文/ENG」整屏語言切換**）。內嵌 2026-01～09 共 56,370 列，已與主報表獨立驗算對帳一致（$159.80M、6 月 6,816 列、8 月 6,298 列全對）。由 AI app-builder 的 React+MySQL 模板移植而來，無需安裝任何東西；狀態修訂僅存瀏覽器 localStorage，不回寫 Excel。語言偏好存 localStorage，`?lang=en`/`?lang=zh` 可直啟對應語言（也供無頭驗證）。每月重建方式見 `3D Earth Dashboard\README.md`；3D 地球技術已沉澱為個人 skill `globe-engine`（已推送 GitHub Skills + MyVault 個人 OS）。2026-10-30 加「AI 助理」面板（依 skill `qwen-assistant-static-html`）：僅連本機 127.0.0.1:8765 共用 MPL 的 Qwen 中繼、HTML 內零憑證、只送彙總不送明細、AI 建議篩選需人工按「套用」、同事沒開中繼時面板自動顯示離線；詳見該資料夾 README |

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

⚠️ **Product Family coverage varies by month**: August 2026 HTML has no TVS ProdFamiDesc values (3,274 rows); 347/454 MAZ rows have values. Other fields and months were not checked.
**New in 2026-08-28 to 09-02; updated 2026-09-09**: dashboard filters, family charts, drill-down links, resizable tables and weighted KPI cards. Check family coverage by month; August has no TVS ProdFamiDesc values and some MAZ values.

**一次性資料修補工具**：`Copy-ProdFamiDescToSource.ps1` 把 `OUTSIDETEMP.xlsx` 裡已經算好的欄位數值（不是公式），逐列核對 PO_NO/VEND_INV_NO/ITEM_CODE/REC_QTY 對齊無誤後，複製進正式月報來源檔案——核對不通過會直接中止、不寫入任何資料。透過 `-ColumnName` 參數可重複用來搬移 `ProdFamiDesc`（預設）、`ProdFamiCode`、`ProdFamilyFittingsCode` 這三個欄位（不存在就新增一欄，已存在就覆寫數值）。
**One-off data-patch tool**: `Copy-ProdFamiDescToSource.ps1` copies the already-computed values (not the formula) of a named column from `OUTSIDETEMP.xlsx` into the official monthly source file, only after verifying every row aligns by PO_NO/VEND_INV_NO/ITEM_CODE/REC_QTY — aborts before writing anything if any row doesn't match. The `-ColumnName` parameter lets it be reused for all three columns — `ProdFamiDesc` (default), `ProdFamiCode`, `ProdFamilyFittingsCode` — adding a new column if it doesn't exist yet, or overwriting values if it does.
