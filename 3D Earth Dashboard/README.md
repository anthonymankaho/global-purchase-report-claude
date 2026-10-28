# 3D Earth Dashboard — 全球採購報告 (Global Purchase Report)

> 本產品正式名稱為「全球採購報告 / Global Purchase Report」(2026-10-29 定案,不使用「戰情室/War Room」等名稱;原始 React 模板標題曾被改成「戰情室」,移植時已一併去除)。
> Official product name is "Global Purchase Report / 全球採購報告" (finalized 2026-10-29; "War Room"-type names were removed, including the one the source React template had been relabeled to).

一面 1920×1080 風格的「指揮大螢幕」3D 採購報告看板,已轉成**單一自包含 HTML 檔**:
双击即可開啟,不需要 Node、MySQL、網路或架站。
A command-wall style 3D purchase "war room", converted to a **single self-contained HTML file** —
double-click to open; no Node, MySQL, network, or web server needed.

- 正式版檔案 / The deliverable: `Global_Purchase_3D_WarRoom_TEMP.html` (~14.4 MB)
- 來源 / Origin: 2026-10-28 由 AI app-builder 產出的 React+Vite+tRPC+MySQL 大螢幕模板
  (下載於 `C:\Users\User\Downloads\3D Earth Dashborad\app`),2026-10-29 由 Claude Code
  移植為純 HTML/JS。React 架構僅保留在來源專案,不在此資料夾。
  Originally a React 19 + Vite 7 + tRPC + MySQL app-builder export; ported to vanilla
  single-file HTML on 2026-10-29. The React scaffold stays in Downloads, not here.

## 畫面內容 / What's on screen

中央為 three.js 程序化 3D 地球(晝夜晨昏線、海岸線、城市燈、軌道環),七個實體以
城市錨點顯示(MHK→香港、STC→上海、MZQ→深圳、RACO→廣州、MIL→曼谷、TVS→羅馬、
MAZ→利馬),點節點即篩選該實體;兩翼:KPI 五卡、企業實體排行、品類圓環圖、
月度採購額×收貨數量趨勢、供應商 TOP8 排行、採購訊號(即時依篩選計算)、
12 筆/頁的採購明細表(搜尋/排序/明細抽屜/CSV 匯出)。深淺色主題、自動巡覽。
右上角**「中文 / ENG」按鈕**可即時整屏切換繁中／英文(含圖表圖例、訊號文案、下拉選單、
表頭、彈窗/抽屜、數字格式),語言偏好存 localStorage;可用 `?lang=en` / `?lang=zh` URL
參數直啟對應語言(也供無頭驗證用)。CSV 欄名固定英文(對齊 Excel 主檔表頭)。
A **中文/ENG toggle** in the top bar switches the whole screen (charts, signals, dropdowns,
table headers, modals, number formatting) between Traditional Chinese and English; the choice
is remembered in localStorage, and `?lang=en`/`?lang=zh` open directly in a language (also used
by the headless test). CSV column headers stay English to match the Excel master.
Center: procedural three.js globe (terminator, coastlines, city lights) with one node
per entity; click = filter. Wings: 5 KPI cards, entity ranking, commodity donut, monthly
trend, top-8 suppliers, live "procurement signals", and a searchable/sortable detail
table with row drawer and CSV export. Dark/light theme, auto-tour.

## 資料 / Data

- 內嵌快照:`2026-01 ~ 2026-09`,56,370 列,欄位即主報表的 20 欄
  (PO_NO … Std Vendor Name / Std Comm / yyyy-mm / ignore-consider / ProdFamiDesc/Code)。
- 已與主報表對帳:2026-06=6,816 列、2026-08=6,298 列、Jan–Aug 累計 50,262 列,
  與 README 2026-10-27 記錄完全一致(50,262+9月6,108=56,370)。
- Embedded Jan–Sep 2026 snapshot; verified against the master report row counts.
- 節點的**地理位置是示意對照**,金額/排行/明細都來自正式報表資料。
  Node *map positions are illustrative*; every number comes from the real report.

## AI 助理(選配,共用中繼)/ AI Assistant (optional, shared relay)

右上角「AI 助理」按鈕開啟對話面板,可依照**你目前的篩選範圍與彙總數字**問答(哪個月最高、供應商集中度、Ignore 占比…),並在助理建議收窄篩選時顯示「套用」按鈕(由你確認才生效)。
The top-bar AI Assistant answers from the CURRENT filters + aggregates on screen, and offers Apply-chips (never auto-applied) when narrowing would help.

- **本檔內零憑證 / no credential in this file**:面板只連本機 `http://127.0.0.1:8765`,API key 存於 `C:\Users\User\.mpl-qwen\proxy.config.json`(Desktop 同步範圍之外)。此 port 與 MPL_System 共用同一個中繼程序 — 不需也刻意不设第二個 proxy。
- **啟動 / to start**(只有報表承辦人需要):`C:\Users\User\.mpl-qwen\start-proxy.bat`(Node 常駐小視窗,關掉即離線)。
- **同事拿到 HTML 但沒開中繼**:面板顯示「AI 助理未啟動」,其餘一切照常 — 這是設計行為。
- **會送出的內容**(面板內有同等Disclosure):目前篩選條件 + 彙總 USD 金額 + 對話;不送明細列、不送任何憑證。用量:每帳號 1M tokens/月(relay 端計數、80% 警告、用完封鎖),7 秒內重複送出會被節流。
- 對話只存在記憶體,按「清空對話」或關閉面板重開即不殘留;稱呼(預設 Anthony)只用於本機用量記錄。
- 驗證(2026-10-30):成品檔 token 洩漏掃描=0;真實 chat 往返成功(MIL $54,809,248 對答正確、只依彙總回答、語言跟隨 UI);中繼關閉時面板優美降級。依 skill `qwen-assistant-static-html` 執行。

## 本機修訂(非資料庫)／Local edits (not a database)

原 React 版的"新增採購/改狀態"寫 MySQL;單檔版改為 **localStorage 覆寫層**:
修訂只存在這台電腦這個瀏覽器,即時套用到指標/圖表/匯出,**不會寫回 Excel 主檔**。
Excel/主報表仍是唯一正式來源。快照重建後若列數或數量總和變動,舊修訂自動作廢。
Edits (status toggle, added rows) persist only in that browser's localStorage and never
write back to the Excel master — the master report remains the single source of truth.
The overlay self-invalidates when the snapshot is rebuilt with different totals.

## 重新產出 / Rebuilding

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude\3D Earth Dashboard\build"
node build.mjs            # → ../Global_Purchase_3D_WarRoom_TEMP.html
node build.mjs 202610     #  additionally publishes ../Global_Purchase_3D_WarRoom_202610.html
```

更新每月資料 = 替換 `build/data/report-seed.json`(同 20 欄格式,fields+rows 陣列),
再跑上面指令。`build/` 內含全部原料:`src/`(移植後程式)、`vendor/`(three/echarts/字型/圖示)、
`styles-original/`(模板原 CSS)、`expected.mjs`(獨立驗算腳本)。
To add a month: replace `build/data/report-seed.json` with the same 20-field
`{fields,rows}` export, then rerun. `expected.mjs` independently recomputes the headline
KPIs from the seed for cross-checking the page.

## 驗證紀錄 / Verification (2026-10-29)

- 2026-10-29(同日稍後):新增「中文/ENG」整屏語言切換,headless 雙語渲染驗證
  (英文模式下 KPI、訊號文案、表頭、分頁皆為英文;數字縮位 EN 用 M、ZH 用 億/萬)。
  同時補強地球模組生命週期:`visibilitychange` 顯式暫停、`webglcontextlost/restored`
  自動重建(長時間掛背景的電腦可能丟失 WebGL context)。
- Headless Edge 1920×1080 實際渲染截圖:總採購額 $159.80M、MIL $54.81M (34.3%)、
  最大供應商 STC 15.7%、2026-09 環比 −7.4%、Ignore 32.6%、56,370 筆 / 4,698 頁 —
  全部與 `expected.mjs` 對 seed 的獨立重算一致;無 JS console 錯誤;
  WebGL 地球在無 GPU 環境亦成功啟動。
- Verified end-to-end against an independent recomputation; globe boots even on SwiftShader.

## 瀏覽器需求 / Browser requirements

需要支援 WebGL + ES modules 的現代瀏覽器(Edge/Chrome/Firefox 近年版本均可);
用 `file://` 直接開啟即可。公司 IE 或非常老的瀏覽器會顯示降級提示與靜態面板。
Any modern browser (Edge/Chrome/Firefox) with WebGL + ES modules; opens fine from `file://`.
