# Global Purchase Report 標準作業程序 (SOP)

**版本**: v1.5（新增外部交付檔案流程 `Publish-OutsideReport.ps1`：把月報數值複製進 OUTSIDETEMP.xlsx、延伸 Table1、重新整理樞紐分析表、另存當月交付檔；七個實體全部支援；Q2:AJ2 公式改回自動化，一欄一欄填完立刻凍結成數值，避免整批重算；AA/AB 改用 PowerShell 計算取代 MAXIFS/MINIFS；新增「New Vendors」新供應商檢查分頁；修正 MIL 的 PART TYPE 欄位問題）
**適用範圍**: 每月 Global Purchase Report（MHK/STC/MZQ/MIL/RACO/TVS/MAZ 採購資料彙總）產製流程
**取代對象**: 原本由 UiPath RPA 專案 "Global Purchase Report" 自動執行的合併/彙總/報表產出步驟

---

## 1. 流程總覽

| 階段 | 執行方式 | 說明 |
|---|---|---|
| 0. 取得當月匯率表 | **人工** | Gmail 搜尋、下載、改名，見第 2.1 節 |
| 1. 取得各實體當月採購資料 | **人工** | 依實體不同，分為「EPICOR 匯出」與「廠商/財務提供檔案」兩種來源，見第 2.2 節 |
| 2. 合併資料、產出最終報表 | **PowerShell 腳本**（`Build-GlobalPurchaseReport.ps1`） | 承辦人執行腳本，見第 4 節 |
| 3. 檢查最終報表 | **人工**確認 | 承辦人 |
| 4. 寄送報表 / 上傳 MMS | **人工**（不自動化） | 同原 RPA 步驟，本專案範圍不含 |

> 為什麼步驟 1、0 維持人工？因為同一個 EPICOR 入口網站（`portal.meyer-msc.com/EPICOR-SUB/`）在另一個已完成的 PowerShell 專案（MTC Purchase Report）測試中，曾經因為瀏覽器自動登入/匯出造成 Excel 程序卡死，風險較高；Gmail 附件抓取則是因為技術上只能在跟 Claude 的對話中操作，無法讓 PowerShell 腳本獨立呼叫（除非另外申請 Google API/OAuth，工程量遠大於人工操作一分鐘）。人工操作較穩定、成本也低，因此採用相同策略。

---

## 2. 資料來源

### 2.1 匯率表（Exchange Rate）

1. 開啟 Gmail（`anthonym@meyer-msc.com`），搜尋「AR / AP Exchange rate for MMM YYYY」（例：`AR / AP Exchange rate for JUN 2026`）。
2. 下載附件 `AR / AP Exchange rate for MMM YYYY.xlsx`。
3. 改名為 `STANDARD EXCHANGE RATE.xls`（或 `.xlsx`，兩種副檔名腳本都認得），放進：
   `Global Purchase Report Claude\Exchange Rate\STANDARD EXCHANGE RATE.xls`（每月覆蓋即可）

> **已自動化**：`Build-GlobalPurchaseReport.ps1` 執行時會自動讀取這個檔案的「Exchange rate table」分頁，換算成 USD 對應匯率（來源檔是「1 單位外幣 = X 港幣」，Report 要的是「1 單位外幣 = X 美金」，換算方式：來源匯率 ÷ 來源 USD 匯率），寫進 Report「Exchange Rate」分頁對應月份那一列。不需要再人工轉貼數字，只需要完成上面 1-3 步驟把檔案放進資料夾即可。

### 2.2 七個實體

| 實體 | 資料來源 | 檔案特性 | 當月篩選欄位 |
|---|---|---|---|
| MHK | EPICOR ACT001 報表（人工登入匯出） | 匯出時已篩當月，單月資料 | 不需篩選 |
| STC | EPICOR ACT001 報表（人工登入匯出） | 同上 | 不需篩選 |
| MZQ | 廠商提供「進倉」累積檔 | 實際資料在檔案內部一個叫 **"MHK"** 的分頁（歷史命名，非真的 MHK 資料），Jan-當月累積 | B 欄（DATE） |
| MIL | 廠商提供累積檔 | 實際資料在分頁 **"MIL"**，Jan-當月累積 | B 欄（Date） |
| RACO | 廠商提供檔案 | 實測目前來源檔只有 **"Sheet1"** 一個分頁（原設計參照的「RACO收货记录」分頁不存在，腳本已自動 fallback），Jan-當月累積 | B 欄（DATE） |
| TVS | 財務部提供、經 Gemini 整理的草稿表 | 22 欄，跟 Report 目標欄位**不是**一對一對應（詳見第 3 節），Jan-當月累積 | G 欄（Invoice Date） |
| MAZ | 廠商每月更新檔（放進 `MAZ\MAZ.xlsx`） | 實際資料在檔案內部一個叫 **"MHK"** 的分頁（歷史命名，同 MZQ 的情況），Jan-當月累積；**若某月尚未收到資料，腳本會自動略過並印出警告，不會中斷整個執行** | B 欄（Purchase Order Date） |

放置位置：`Global Purchase Report Claude\<實體>\<實體>.xlsx`（例：`MHK\MHK.xlsx`）。MAZ 資料夾名稱是 `MAZ`（不是 `MAZAL`），Report 裡的 DB 標記也是 `"MAZ"`（不是 `"MAZAL"`）。

---

## 3. 各實體欄位對應（供除錯/驗證參考）

- **MHK / STC**：原始 ERP 匯出 → 貼進各自的 `Trimmed Program\<實體> Act001 Trimmed.xlsx`（ACT001 分頁，欄位依實體重新對應）→ Trimmed 分頁用**既有公式**（不是腳本寫死的邏輯）自動算出其餘欄位 → 取 Trimmed!A:T 併入 Report。
  ⚠️ 這兩個實體的 Trimmed 樣板檔清除資料時，**只能清 A 欄**（`A3:A5000`），不能清到 A:T，否則會把公式列一併清掉，造成整批資料算出空白（測試時真的踩到這個坑，已修正）。
- **MZQ / MIL**：無 Trimmed 樣板，篩出當月後直接把來源欄位 A:P、T 複製進 Report（Q:S 由 Report 自己的公式計算）。
- **RACO**：來源 Sheet1 欄位跟 Report A:T 一對一對應。腳本每次都固定開啟同一個未命名日期的範本檔 `Trimmed Program\RACO Act001 Trimmed.xlsx`（唯讀），篩出當月資料貼進去，再另存成當月的日期檔（`RACO\RACO Act001 Trimmed 'MM-yyyy.xlsx`），最後併入 Report。範本檔跟另存的目的檔名永遠不同，不會有存檔衝突。TVS 也是同樣機制（範本在同一個資料夾）。
- **TVS**：來源欄位**不是**一對一對應，已用「比對舊 Report 裡已知正確的 TVS 列 vs 原始檔」的方式實測驗證出正確對應關係：

  | Report 欄位 | 來源欄位 |
  |---|---|
  | PO_NO | Purchase Order Number |
  | PO_REC_DATE | **Invoice Date**（不是 PO Date）|
  | PO_ISS_DATE | Purchase Order Date |
  | VENDOR_ID / VENDOR_NAME / VEND_INV_NO | Vendor ID / VendorName / Vendor_Invoice_No |
  | REFERENCE | REF_NO |
  | ITEM_CODE | Item Code |
  | PART_DESC | **Part Description（義大利文原文）**，不是翻譯後的英文欄位 |
  | COMM_CODE | **Unified Commodity Code**（VLOOKUP 算出來的那欄，不是 Commodity Italian/English）|
  | PBRT_TYPE | 固定值 `"B"` |
  | REC_QTY / UOM | Quantity / Unit |
  | PO_CUR | 固定值 `"EUR"` |
  | PO_PRICE_FCY | **用 Total÷Quantity 反推**（來源自己的單價欄位跟實際金額對不上）|
  | TOT_PO_PRICE_FCY | Total_PO_Price_PO CUR |

- **MAZ**：無 Trimmed 樣板（原 RPA 沒有 MAZAL 的自動化區段，這個對應關係是比照 TVS 的方式、用「比對舊 Report 裡已知正確的 MAZ 列（`OC-13367`、`OC-13500` 等 PO）vs 原始檔」實測驗證出來的）。比 TVS 簡單，來源欄位跟 Report 欄位幾乎一對一對應，只差一欄：

  | Report 欄位 | 來源欄位 |
  |---|---|
  | PO_NO ~ COMM_CODE（欄 1-10） | 來源欄 1-10（Purchase Order Number ~ Commodity），直接對應 |
  | PBRT_TYPE | 固定值 `"B"`（來源的 Division 欄不是這個用途，每筆樣本都是 "1"，跟 Report 的 "B" 對不上）|
  | REC_QTY ~ TOT_PO_PRICE_FCY（欄 12-16） | 來源欄 11-15（Quantity ~ Total_PO_Price_PO CUR），直接對應（因為欄 11 被 PBRT_TYPE 常數插入，往後位移一欄）|
  | PO_CUR | 來源欄 13（PO_Price_PO CUR，欄名雖然像價格，實際存的是幣別代碼，例如 "USD"、"PEN"）——**不是固定值**，MAZ 有多幣別交易 |
  | DB | 固定值 `"MAZ"` |

  ⚠️ MAZ 有 PEN（秘魯幣）交易，但目前 `STANDARD EXCHANGE RATE` 來源檔沒有 PEN 匯率，Report 的 Exchange Rate 分頁本來就有 PEN 欄位但長期空白，屬已知缺口（見第 5 節）。

- **Report 共用公式（Q:S 匯率換算、U:AD 廠商/料號/商品標準化與比價）**：所有實體都用同一組公式，公式母版固定在 Report 的**第 2 列**（`Q2:AJ2`）。
  ⚠️ **重要架構變更**：Report 樣板檔案裡，第 3 列到歷史最後一列（目前約 29,511 列）的 Q:AJ 已經從「即時公式」**手動轉成靜態數值**（只留第 2 列是公式），目的是避免每次新增資料時，Excel 觸發全部 3 萬多列重新計算（尤其 U:AD 的 MAXIFS/MINIFS 是整欄掃描，資料量大時會讓 Excel 卡住 10 分鐘以上甚至更久）。
  因此 `Build-GlobalPurchaseReport.ps1` **不會**自動把 Q2:AJ2 公式貼到新增列——這一步改為**人工**：腳本執行完畢後，畫面會印出新增列的範圍（例如「New rows 29512 to 36272 need Q2:AJ2 copied down manually in Excel」），承辦人自行在 Excel 裡把 Q2:AJ2 複製、貼到那個範圍即可。

---

## 4. 每月執行步驟

### 步驟 1：準備當月資料（見第 2 節）

依序完成：匯率表、七個實體資料檔，放進對應資料夾（MAZ 若當月尚未收到資料，可以先不放，腳本會自動略過並警告，不會中斷執行）。

### 步驟 2：開啟 PowerShell，切換到腳本資料夾

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude"
```

### 步驟 3：執行主腳本

```powershell
.\Build-GlobalPurchaseReport.ps1 -ReportMonth "2026-06"
```

> `-ReportMonth` 格式固定 `yyyy-MM`。其他路徑參數（樣板資料夾、上個月報表等）**都會自動偵測**，不需要每個月手動改——腳本會自動接續上個月自己產出的報表繼續往下加資料。若自動偵測到「這個月的資料似乎已經存在」，會直接中止並提示可以加 `-Force` 覆蓋執行（避免不小心重複執行造成資料重複）。

執行過程中畫面會依序顯示：

```
Auto-detected ReportTemplatePath (latest prior monthly output): ...
Auto-detected RacoTrimmedFile: ...
Auto-detected TvsTrimmedFile: ...
Auto-detected ExchangeRatePath: ...
[MHK] Processing ...
[MHK] Source rows: XXX (single-month export, no filter needed)
[STC] Processing ...
...
[TVS] New month's snapshot saved to ...
[MAZAL] Processing ... (或若 MAZ\MAZ.xlsx 不存在: "[MAZAL] ... not found - skipping (data not received yet this cycle)." 不會中斷執行)
Copying report template to output path...
Appended rows XXXXX to YYYYY into Report Data sheet.
[ExchangeRate] Processing ...
[ExchangeRate] Filled N Block 1 column(s) for 2026-06 at row NN.
Filling Q:S/U:AJ formulas for rows XXXXX to YYYYY (one column at a time)...
  [U] filled, calculated, frozen
  ...
[NewVendorCheck] Found N 'NEW VENDOR' row(s) in rows XXXXX-YYYYY - copied to 'New Vendors' sheet for review.
  [AA/AB] computed in PowerShell (grouped by item code + B/O type), written as values
  [AC] filled, calculated, frozen
  ...
Saved: ...\REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_2026-06.xlsx
=== Done. Please open the report, run Data > Refresh All, and spot-check before sending. ===
```

> Q2:AJ2 的公式填入已經**自動化**（一欄一欄填、算完立刻凍結成數值，避免整批一起重算卡住）。歷史列的公式不會被重跑，只處理這次新增的列。AA/AB（Max/Min PO Price）改用 PowerShell 計算後直接寫入數值，不再透過 Excel 的 MAXIFS/MINIFS 公式，速度更快、也不會隨資料量增加而變慢。X 欄（yyyy-mm 期間索引）維持人工在 Temp 參考檔案裡預先鋪好公式到 X70000，腳本不會去動它。

### 步驟 4：（如有需要）更新 TVS 翻譯/商品代碼

```powershell
.\Update-TVSTranslations.ps1
```

這支腳本會自動檢查/修補 TVS 的翻譯公式（含新詞彙的 TRANSLATE 備援）、稽核商品代碼、清除 AI 翻譯殘留的引註文字。TVS 資料若有異常（例如新料號翻不出來、代碼對不上標準表），可以先跑這支。

### 步驟 5：檢查「New Vendors」分頁

腳本執行時會自動把這次新增列裡，Std Vendor Name 算出 "NEW VENDOR"（代表廠商+實體的組合不在標準供應商名稱對照表裡）的整列，複製到報表裡的「New Vendors」分頁，方便直接檢查，不用在 3 萬多列的 Data 分頁裡手動篩選。這個分頁每次執行都會清空重填，只反映當月的新供應商，不會累積歷史資料。若有需要，把新供應商加進「Std Vendor Name」對照表後，下個月執行就不會再出現。

### 步驟 6：（建議）產出稽核摘要報告

```powershell
.\New-ReportAuditSummary.ps1
```

只唯讀開啟報表，不會修改任何東西。會列出：(1) 報表最後一列列號、(2) 當月六個實體各自的資料列數、(3) 當月總列數，並存成 `REPORT\AuditSummary_<月份>.md`。可以拿這份跟腳本執行時畫面上顯示的各實體列數互相核對。

### 步驟 7：開啟最終報表、人工檢查

1. 按「資料」→「全部重新整理」
2. 確認匯率（Q 欄）有沒有算出來——若當月匯率表還沒補（見第 2.1 節），這裡會是空白，屬正常現象，等匯率補齊後重新整理即可
3. 檢查各實體列數是否合理（可對照步驟 6 的稽核摘要，或腳本執行時畫面上顯示的數字）
4. 確認無誤後，依公司規定歸檔 / 寄送（Email 寄送、MMS 上傳仍為人工）

### 步驟 8：（如有需要）產出外部交付檔案（OUTSIDETEMP → 當月交付檔）

當內部 Report（步驟 3-7）確認無誤後，若需要產出給外部/其他系統使用的交付檔案（`FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-(1-M)YYYY_YYYYMMDD_RV0.xlsx`，內含 Dashboard、GLID、Fittings、Packaging、Coating 等樞紐分析表），執行：

```powershell
.\Publish-OutsideReport.ps1 -ReportMonth "2026-06"
```

這支腳本會：
1. 從步驟 3-7 產出的月報（`REPORT\FIN_..._PUR_REPORT_2026-06.xlsx`）讀取「MHK+STC+MZQ+MIL+RACO Data」分頁的**數值**（非公式）。
2. 開啟固定的工作檔 `Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-OUTSIDETEMP.xlsx`，先延伸其 `Data` 分頁的 `Table1` 表格範圍，再把數值寫入（**只寫資料列，不覆寫表頭列**——詳見下方「已知限制」的重要技術備註）。
3. 重新整理所有樞紐分析表（來源是 `Table1` 的會成功；來源是已損壞的 `data` 命名範圍的會失敗並印出清單，屬已知缺口，見第 5 節）。
4. 存檔（覆寫 OUTSIDETEMP.xlsx 本身），再另存新檔到同一個 REPORT 資料夾，檔名格式 `FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-(1-M)YYYY_YYYYMMDD_RV0.xlsx`（`M`=當月月份數字、`YYYY`=年、`YYYYMMDD`=執行當天日期；若同一天重跑，`RV` 編號會自動遞增，不會覆蓋前一次的交付檔）。

> ⚠️ `OUTSIDETEMP.xlsx` 是這個交付流程唯一的「工作檔」，每次執行都會被直接改動（累加資料、延伸表格），不是唯讀範本。執行前建議先手動備份一份，尤其是第一次執行或懷疑檔案有異常時。

---

## 5. 已知限制與注意事項

| 項目 | 說明 |
|---|---|
| **RACO 來源分頁名稱** | 目前來源檔只有 "Sheet1"，腳本會自動 fallback，但若廠商未來改回提供「RACO收货记录」這個分頁名稱，腳本一樣能正確處理（會優先找這個名稱） |
| **TVS 的 PBRT_TYPE/PO_CUR 為寫死常數** | 目前根據實測樣本都是 "B"/"EUR"，若未來 TVS 出現非歐元交易，此假設需要重新驗證 |
| **MAZ 的 PBRT_TYPE 為寫死常數，PO_CUR 不是** | PBRT_TYPE 固定 "B"；PO_CUR 是直接複製來源欄位（MAZ 有 USD、PEN 等多幣別交易，不能當常數） |
| **MIL 的 PART TYPE（K 欄）需要正規化** | 來源檔案 K 欄其實是「Div」部門代碼（1,2,6,7,8,9，偶有 "M"），不是 B/O 分類，曾經因此讓 AG 欄（Max Vendor）VLOOKUP 出現 #N/A。腳本已修正：數字 1-9 一律轉成 "B"（MZQ 不受影響，來源本來就是正確的 B/O） |
| **MAZ 的 PEN 匯率缺口** | `STANDARD EXCHANGE RATE` 來源檔沒有 PEN（秘魯幣）匯率，Report 的 Exchange Rate 分頁 PEN 欄位長期空白，PEN 交易的 Q:S 換算會受影響，需要另外取得 PEN 匯率來源才能補齊 |
| **Excel COM 穩定性** | 測試過程中曾遇到 Excel 卡死（`Not Responding`）與殘留背景程序，腳本已加上啟動前自動清除殘留 Excel 程序的邏輯，但仍建議執行前先手動確認工作管理員沒有殘留的 EXCEL.EXE |
| **MAZ 若某月無資料** | 腳本會自動偵測 `MAZ\MAZ.xlsx` 是否存在，不存在就跳過該實體並印出警告，不會中斷整個執行 |
| **Email 寄送 / MMS 上傳** | 不在本專案自動化範圍內，維持人工 |
| **OUTSIDETEMP.xlsx 的 `data` 命名範圍已損壞** | 工作簿裡有一個叫 `Data`（大小寫不分）的活頁簿層級命名範圍，`RefersTo` 顯示 `=#REF!#REF!`（已損壞的參照），這是 `Comm Analysis`/`Vend Name Analysis`/`Desc Analysis` 三個隱藏分頁裡共 18 個樞紐分析表的來源。`Publish-OutsideReport.ps1` 執行時這三個分頁的樞紐分析表會重新整理失敗並印出清單，**這是既有問題、不是這次腳本造成的**，需要另外手動修復命名範圍或重建這些樞紐分析表（目前不在本流程自動化範圍內） |
| **⚠️ 寫入 Excel Table（ListObject）時絕對不能覆寫表頭列** | 實測發現：用一個大陣列一次性覆寫整個表格範圍（含第 1 列表頭），即使表頭文字完全相同，也會讓 Excel 的 Table 物件損毀（`ListObjects.Count` 變成 0），導致所有依附在這個 Table 上的樞紐分析表全部失效。`Publish-OutsideReport.ps1` 已修正為「先延伸表格範圍、再只寫入第 2 列以後的資料列」，經測試存檔、重新開啟後 Table 都能正常存活。**未來如果修改這支腳本或寫類似的 Table 覆寫邏輯，絕對不要把表頭列（第 1 列）納入批次寫入的範圍** |

---

## 6. 疑難排解

| 錯誤訊息 / 現象 | 原因 | 解法 |
|---|---|---|
| `Data for <月份> already appears to exist...` | 這個月已經處理過，`-ReportTemplatePath` 自動接續到的檔案裡已有這個月資料 | 確認是否真的要重跑（例如資料有誤要重新產出），是的話加 `-Force`；不是的話不用理它，這是安全機制 |
| `您不能使用這個檔名存檔，因為當初您是以唯讀的方式開啟它的`（RACO 或 TVS 處理階段） | **已修正**（v1.1 前的舊版問題）：曾經用「自動抓最新快照檔」當樣板，同月重跑會撞到自己先前產生的檔案。v1.1 起改成固定開啟未命名日期的範本檔（`RACO/TVS Act001 Trimmed.xlsx`），來源與另存目的地永遠不同檔名，理論上不會再發生 | 若仍發生，檢查 `Trimmed Program\RACO Act001 Trimmed.xlsx`（或 TVS 版本）是否還存在、未被誤刪或改名 |
| `'RACO/TVS Act001 Trimmed.xlsx' not found...` | `Trimmed Program\` 資料夾裡找不到固定範本檔 | 確認 `Global Purchase Report RPA\Trimmed Program\RACO Act001 Trimmed.xlsx`（或 TVS 版本）存在，或用 `-RacoTrimmedFile`/`-TvsTrimmedFile` 手動指定完整路徑 |
| Excel 顯示 `Not Responding` / 執行很久沒反應 | 若歷史列還沒被凍結為靜態值（見第 3 節），大量公式（尤其 Report 的 MAXIFS/MINIFS 比價公式）需要重新計算，資料量大時可能要跑好幾分鐘，不一定是卡住 | 用工作管理員看 EXCEL.EXE 的 CPU 時間是否持續增加（代表仍在運算），先耐心等待；若確定真的卡住（CPU 時間不再增加）才強制結束。長期解法是確保歷史列維持凍結狀態，只留第 2 列公式 |
| `The RPC server is unavailable` | Excel COM 當機（已知風險） | 用工作管理員結束殘留的 EXCEL.EXE，重新執行腳本 |
| Q 欄（EXCH_RATE）空白或公式顯示 #VALUE! | 新增列還沒人工複製 Q2:AJ2 公式（見第 4 節步驟 5），或匯率表當月資料還沒補（見第 2.1 節） | 先確認步驟 5 有沒有做；再確認匯率表有沒有補齊，然後在 Excel 裡按「資料」→「全部重新整理」 |

---

## 7. 版本紀錄

| 日期 | 版本 | 修改內容 |
|---|---|---|
| 2026-07-24 | Draft v1 | 首版，涵蓋資料來源盤點、範圍決策（MAZAL 本月排除）、已知限制；PowerShell 腳本尚未撰寫 |
| 2026-07-28 | v1.0 | `Build-GlobalPurchaseReport.ps1` 完成並通過端對端測試（6 個實體、6,761 列、0 筆公式錯誤）；補上 TVS 正確欄位對應（實測反推）；修正 Trimmed 清除範圍、UsedRange 誤判、公式來源列等真實 bug；新增自動接續（免每月改參數）與重複執行安全檢查；新增匯率表人工取得流程與 `Exchange Rate` 資料夾 |
| 2026-07-28 | v1.1 | 匯率轉貼改為自動化（腳本讀取 `STANDARD EXCHANGE RATE.xls`，換算並寫入 Report）；因應 Report 樣板歷史列改為凍結靜態值（只留第 2 列公式）的架構調整，Q2:AJ2 公式改回人工複製到新增列，不再由腳本自動貼；新增 RACO/TVS 快照檔接續衝突的疑難排解 |
| 2026-07-29 | v1.2 | 修正 RACO/TVS 樣板選取邏輯：改成固定開啟未命名日期的範本檔（不再自動抓「最新快照檔」），徹底解決同月重跑時的存檔衝突；新增 `New-ReportAuditSummary.ps1`（稽核摘要：報表最後列號、六實體當月列數、當月總列數），列為每月建議執行步驟之一 |
| 2026-07-31 | v1.3 | 收到 MAZ 2026-06 資料後補上第 7 個實體：無 Trimmed 樣板、直接對應（比照 TVS 用歷史 Report 列反推驗證），PBRT_TYPE 固定 "B"、PO_CUR 為實際欄位（非常數，因 MAZ 有多幣別交易）；`MAZ\MAZ.xlsx` 不存在時自動略過、不中斷執行；`New-ReportAuditSummary.ps1` 加入 MAZ 到實體清單；移除已解決的「MAZAL 排除」章節，改記錄 PEN 匯率缺口等真正的已知限制 |
| 2026-07-31 | v1.4 | 修正 Exchange Rate 分頁的架構誤解：原本誤把美金換算結果直接寫死進 O:Z（本該是活公式），且完全沒填 C:M 的原始港幣匯率區塊，現已修正為 C:M 直接寫值、O:Z 一律重新寫入公式（不再靠複製上一列，因上一列也可能已被凍結成靜態值）；Q2:AJ2 公式填入改回自動化，一欄一欄依相依順序填公式、算完立刻凍結成數值再換下一欄，避免整批一起重算卡住 Excel；X 欄（yyyy-mm）改由人工在 Temp 參考檔預先鋪好公式到 X70000，腳本不再處理；AA/AB（Max/Min PO Price）改用 PowerShell 依料號+B/O 分組計算後直接寫值，取代原本會隨資料量變慢的 MAXIFS/MINIFS 公式；新增「New Vendors」分頁，自動彙整當月「NEW VENDOR」列供人工檢查；修正 MIL 來源檔 PART TYPE（K 欄）其實是部門代碼、導致 AG 欄 #N/A 的問題 |
| 2026-07-31 | v1.5 | 新增 `Publish-OutsideReport.ps1`：把當月報表「MHK+STC+MZQ+MIL+RACO Data」分頁的數值複製進外部交付工作檔 `OUTSIDETEMP.xlsx`、延伸 Data 分頁的 `Table1` 表格範圍、重新整理所有樞紐分析表、存檔後另存為當月交付檔（`FIN_..._PUR_REPORT-(1-M)YYYY_YYYYMMDD_RV0.xlsx`，同日重跑自動遞增 RV 編號）。開發過程中發現並修正一個嚴重問題：用大陣列一次性覆寫整個 Table 範圍（含表頭列）會讓 Table 物件損毀、造成所有樞紐分析表失效，即使表頭文字相同也一樣——已改為「先延伸表格、只寫資料列（跳過表頭列）」，並用一次性測試副本驗證表格能正常存活後才套用到正式檔案。另外發現 OUTSIDETEMP.xlsx 裡有個已損壞的 `data` 命名範圍（`=#REF!#REF!`），是 3 個隱藏分頁共 18 個樞紐分析表的來源，屬既有缺口、記錄在第 5 節 |
