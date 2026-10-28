/* ── main: render + interaction, ported from App.tsx / PurchasePanels.tsx ── */
const $ = (sel) => document.querySelector(sel);
const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* lucide icon bodies, injected by build.mjs */
const ICONS = /*__ICONS__*/ {};
const icon = (name, size = 16, cls = '') =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="${size}" height="${size}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"${cls ? ` class="${cls}"` : ''} aria-hidden="true">${ICONS[name] || ''}</svg>`;

/* ── i18n: 中文 ↔ English, same 中文/ENG button pattern as the tabular report ─ */
const L = {
  zh: {
    htmlLang: 'zh-Hant', appTitle: '全球採購報告', brandKicker: 'GLOBAL PURCHASE REPORT · 2026', periodLabel: '資料期間',
    embedded: (n) => `內嵌資料 · ${n} 個月`, loading: '內嵌資料載入中…',
    addPurchase: '新增採購', analysisScope: '分析範圍',
    fEntity: '企業實體', fMonth: '月份', fComm: '標準品類', fVendor: '供應商', fStatus: '分析狀態',
    allEntity: '全部實體', allMonth: '全部月份', allComm: '全部品類', allVendor: '全部供應商', allStatus: '全部狀態',
    reset: '重設', entityPanel: '企業實體採購額', clickDrill: '點擊下鑽', categoryPanel: '品類採購結構',
    globePanel: '全球採購實體網絡', trendPanel: '月度採購額與收貨數量', supplierPanel: '供應商採購排行',
    signalsPanel: '採購訊號', signalsFoot: '依目前篩選資料即時計算，非外部即時風險通知。',
    globeBuilding: '正在建立全球實體網絡…', globeFailed: '3D 地球無法啟動，請使用實體排行下鑽',
    networkOverview: '網絡總覽', currentEntity: '目前實體', entitiesCount: (n) => `${n} 個實體`,
    globeHelp: '拖曳旋轉 · 點擊節點篩選',
    globeNote: '節點地理位置為示意對照；採購金額與排行均來自內嵌的正式月報資料。',
    currentScope: '目前範圍', spendUsd: '採購額 USD', receivedQty: '收貨數量', other: '其他',
    kpiSpend: '總採購額', vsPrev: '較前一資料月份', kpiPo: '採購單數',
    receiptLines: (n) => `${n} 筆收貨明細`, kpiQty: '收貨數量', kpiQtyNote: '所有目前篩選品項',
    kpiVendors: '供應商數', topVendor: '最大供應商', kpiItems: '採購料號',
    flags: (n) => `${n} 項注意`, lineCount: (n) => `${n} 筆明細`,
    poDetails: '採購明細', lines: (n) => `${n} 筆`, searchPh: '搜尋 PO、料號、品項或供應商',
    searchBtn: '搜尋', exportNow: '匯出目前明細', exporting: '匯出中…',
    thDate: '收貨日期', thPo: '採購單號', thEntity: '實體', thVendor: '供應商', thItem: '料號／品項',
    thComm: '品類', thQty: '數量', thUsd: '採購額 USD', thStatus: '狀態',
    noRows: '沒有符合條件的明細', clearFilters: '清除所有篩選',
    pager: (p, t) => `第 ${p} / ${t} 頁 · 每頁 12 筆`, prevPage: '上一頁', nextPage: '下一頁',
    viewPo: (po) => `查看採購單 ${po}`, dataSource: '資料來源', localEdits: '本機修訂', editsUnit: '筆',
    clearEdits: '清除本機修訂', toLight: '切換淺色主題', toDark: '切換深色主題', toggleLang: '切換語言 / Toggle language',
    /* drawer */
    totalSpend: '總採購額', receivedOn: (d) => `${d} 收貨`, dEntityMonth: '企業實體／月份',
    dStdVendor: '標準供應商', dVendor: '原始供應商', dComm: '品類', dFamily: '產品家族',
    unassigned: '未指派', dQty: '數量', dPrice: '美元單價', dUpdated: '最後更新', snapshot: '原始報表快照（未修訂）',
    statusH: '報表分析狀態', statusP: '狀態修訂保存在本機瀏覽器，會即時套用到所有指標、圖表與匯出；不會寫回 Excel 主檔。',
    statusSaved: (v) => `狀態已改為 ${v}（保存在本機瀏覽器）`,
    /* create modal */
    newPo: '新增採購明細', mPo: '採購單號', mRecDate: '收貨日期', mDataMonth: '資料月份',
    mVendor: '供應商名稱', mStdVendor: '標準供應商', mVendorPh: '原始供應商名稱', mStdPh: '集團供應商歸屬',
    mItem: '料號', mItemPh: 'ITEM CODE', mDesc: '品項描述', mDescPh: '料號品項與規格',
    mQty: '收貨數量', mPrice: '美元單價', mTotalCalc: '總採購額（USD）',
    invalid: '請輸入有效的數量與單價，總金額必須大於 0。', cancel: '取消',
    saveLocal: '新增並保存（本機）', poAdded: (po) => `採購單 ${po} 已新增（保存在本機瀏覽器）`,
    /* export */
    exported: (n) => `已匯出 ${n} 筆目前篩選明細`,
    clockLocale: 'zh-TW',
    aiBtn: 'AI 助理', aiTitle: '採購報告 AI 助理',
    aiHello: '你好!我可以依目前畫面(篩選範圍＋彙總數字)回答:各實體採購額、月度趨勢、供應商集中度、Ignore 占比等。資料只送到本機 AI 中繼,不含任何憑證。',
    aiPh: '例:8 月哪個實體採購額最高?', aiSend: '傳送', aiClear: '清空對話', aiThinking: '思考中…',
    aiTrunc: '(回覆因長度上限被截斷)', aiErrNet: '無法連線 AI 中繼服務。',
    aiOnline: 'AI 中繼已連線', aiQuota: (u, c) => `本月 ${integer(u)}/${integer(c)} tokens`,
    aiOffline: 'AI 助理未啟動:在本機執行 .mpl-qwen\\start-proxy.bat 後重試;報表其餘功能不受影響。',
    aiAppliedChip: (k, v) => `已套用篩選 ${k}=${v}`, aiApply: '套用',
    aiName: '名稱(僅本機用量紀錄)',
    aiDisclose: '送出內容:目前篩選條件＋彙總數字(各實體/月份/品類/供應商的 USD 金額)、對話。不會送出明細列、原始供應商全名或任何憑證。',
  },
  en: {
    htmlLang: 'en', appTitle: 'Global Purchase Report', brandKicker: 'MEYER GROUP · 2026', periodLabel: 'Reporting period',
    embedded: (n) => `Embedded · ${n} months`, loading: 'Loading embedded data…',
    addPurchase: 'Add PO', analysisScope: 'Analysis scope',
    fEntity: 'Entity', fMonth: 'Month', fComm: 'Std Comm', fVendor: 'Vendor', fStatus: 'Status',
    allEntity: 'All entities', allMonth: 'All months', allComm: 'All comms', allVendor: 'All vendors', allStatus: 'All statuses',
    reset: 'Reset', entityPanel: 'Spend by entity', clickDrill: 'Click to drill', categoryPanel: 'Commodity mix',
    globePanel: 'Global entity network', trendPanel: 'Monthly spend & received qty', supplierPanel: 'Vendor ranking',
    signalsPanel: 'Procurement signals', signalsFoot: 'Computed live from the current filters; not an external risk feed.',
    globeBuilding: 'Building the global entity network…', globeFailed: '3D globe unavailable — use the entity ranking',
    networkOverview: 'Network overview', currentEntity: 'Current entity', entitiesCount: (n) => `${n} entities`,
    globeHelp: 'Drag to rotate · click a node to filter',
    globeNote: 'Node positions are illustrative; amounts and rankings come from the embedded official monthly report.',
    currentScope: 'Current scope', spendUsd: 'Spend USD', receivedQty: 'Received qty', other: 'Other',
    kpiSpend: 'Total spend', vsPrev: 'vs prev. month', kpiPo: 'Purchase orders',
    receiptLines: (n) => `${n} receipt lines`, kpiQty: 'Received qty', kpiQtyNote: 'all items in current scope',
    kpiVendors: 'Vendors', topVendor: 'Top vendor', kpiItems: 'Item codes',
    flags: (n) => `${n} flags`, lineCount: (n) => `${n} lines`,
    poDetails: 'PO details', lines: (n) => `${n} lines`, searchPh: 'Search PO, item code, description or vendor',
    searchBtn: 'Search', exportNow: 'Export current view', exporting: 'Exporting…',
    thDate: 'Receipt date', thPo: 'PO number', thEntity: 'Entity', thVendor: 'Vendor', thItem: 'Item / description',
    thComm: 'Std Comm', thQty: 'Qty', thUsd: 'Spend USD', thStatus: 'Status',
    noRows: 'No lines match the filters', clearFilters: 'Clear all filters',
    pager: (p, t) => `Page ${p} of ${t} · 12 per page`, prevPage: 'Prev', nextPage: 'Next',
    viewPo: (po) => `View PO ${po}`, dataSource: 'Source', localEdits: 'Local edits', editsUnit: '',
    clearEdits: 'Clear local edits', toLight: 'Switch to light theme', toDark: 'Switch to dark theme', toggleLang: '切换语言 / Toggle language',
    /* drawer */
    totalSpend: 'Total spend', receivedOn: (d) => `received ${d}`, dEntityMonth: 'Entity / month',
    dStdVendor: 'Std vendor', dVendor: 'Vendor name', dComm: 'Std comm', dFamily: 'Product family',
    unassigned: 'Unassigned', dQty: 'Quantity', dPrice: 'Unit price USD', dUpdated: 'Last updated', snapshot: 'Master snapshot (unmodified)',
    statusH: 'Analysis status', statusP: 'Status edits live only in this browser and apply instantly to every KPI, chart and export; they never write back to the Excel master.',
    statusSaved: (v) => `Status set to ${v} (saved in this browser)`,
    /* create modal */
    newPo: 'New purchase record', mPo: 'PO number', mRecDate: 'Receipt date', mDataMonth: 'Data month',
    mVendor: 'Vendor name', mStdVendor: 'Std vendor', mVendorPh: 'original vendor name', mStdPh: 'group vendor mapping',
    mItem: 'Item code', mItemPh: 'ITEM CODE', mDesc: 'Description', mDescPh: 'item and specification',
    mQty: 'Received qty', mPrice: 'Unit price USD', mTotalCalc: 'Total spend (USD)',
    invalid: 'Enter a valid quantity and unit price — the total must be greater than 0.', cancel: 'Cancel',
    saveLocal: 'Save (local only)', poAdded: (po) => `PO ${po} added (saved in this browser)`,
    /* export */
    exported: (n) => `Exported ${n} filtered lines`,
    clockLocale: 'en-GB',
    aiBtn: 'AI Assistant', aiTitle: 'Purchase Report AI Assistant',
    aiHello: 'Hi! I can answer from what is on screen right now (your filters + aggregates): entity spend, monthly trends, vendor concentration, Ignore share, etc. Data only goes to the local relay - never any credentials.',
    aiPh: 'e.g. which entity had the highest spend in Aug?', aiSend: 'Send', aiClear: 'Clear chat', aiThinking: 'Thinking…',
    aiTrunc: '(reply truncated at the length cap)', aiErrNet: 'Cannot reach the AI relay service.',
    aiOnline: 'AI relay connected', aiQuota: (u, c) => `${integer(u)}/${integer(c)} tokens this month`,
    aiOffline: 'Assistant not running: start .mpl-qwen\\start-proxy.bat on this PC and retry; the rest of the report is unaffected.',
    aiAppliedChip: (k, v) => `Filter applied: ${k}=${v}`, aiApply: 'Apply',
    aiName: 'Name (local usage record only)',
    aiDisclose: 'Sent: current filters + aggregate USD figures (entity/month/commodity/vendor) + the conversation. Never sent: transaction rows, original vendor names, or any credential.',
  },
};

const initialLang = (() => {
  const q = new URLSearchParams(location.search).get('lang');
  if (q && L[q]) return q;                       /* ?lang=en also doubles as the headless test hook */
  try { return localStorage.getItem('gpr-lang') === 'en' ? 'en' : 'zh' } catch { return 'zh' }
})();

const state = {
  lang: initialLang,
  filters: {},
  search: '',
  page: 1,
  sortBy: 'poRecDate',
  sortDir: 'desc',
  touring: false,
  theme: (() => { try { return localStorage.getItem('gpr-theme') === 'light' ? 'light' : 'dark' } catch { return 'dark' } })(),
};
const T = () => L[state.lang];
NUM_LOCALE = state.lang === 'zh' ? 'zh-TW' : 'en';

let OPTS = buildOptions();
let OV = buildOverview(filteredRows(state.filters, ''), state.lang);
let donutChart = null, trendChart = null, globeApi = null, globeReady = false, globeTried = false;
let tourTimer = 0, toastTimer = 0;
let selectedId = null;   /* row id held by the open drawer */
let createOpen = false;

/* ── three.js bootstrap: the ESM build ships as inert JSON-encoded text in
   #gpr-three; we hand it to the engine as a blob module on first use, so the
   panels paint before the 3D work starts (same shape as the React original's
   dynamic import()). */
let threePromise = null;
const loadThree = () => {
  if (!threePromise) {
    threePromise = (async () => {
      const src = JSON.parse(document.getElementById('gpr-three').textContent);
      const url = URL.createObjectURL(new Blob([src], { type: 'text/javascript' }));
      try { return await import(url) } finally { setTimeout(() => URL.revokeObjectURL(url), 5000) }
    })();
  }
  return threePromise;
};

async function ensureGlobe() {
  if (globeTried) return;
  globeTried = true;
  try {
    const THREE = await loadThree();
    if (globeApi) return;
    globeApi = makeGlobe(THREE, $('#earth'), {
      routes: [],
      markers: globeMarkers(OV),
      onPick: (marker) => { if (marker.entity) updateFilter('entity', marker.entity) },
    });
    globeReady = true;
    $('#globe-fallback').classList.add('ready');
    if (state.filters.entity) globeApi.focus(ENTITY_HUBS[state.filters.entity] ?? '上海');
  } catch (error) {
    console.warn('Globe initialization failed', error);
    $('#globe-fallback').classList.add('error');
    $('#globe-status').textContent = T().globeFailed;
  }
}

const globeMarkers = (ov) => {
  if (!ov.entities.length) return [];
  const max = Math.max(...ov.entities.map((item) => item.value));
  return ov.entities.map((item) => ({
    city: ENTITY_HUBS[item.name] ?? '上海',
    entity: item.name,
    v: Math.max(12, item.value / max * 100),
    level: state.filters.entity === item.name ? 'warn' : undefined,
  }));
};

/* ── charts ─────────────────────────────────────────────────────────────── */
const donutOption = () => {
  const t = T();
  const rows = OV.categories;
  const top = rows.slice(0, 7);
  const other = rows.slice(7).reduce((sum, row) => sum + row.value, 0);
  const values = [...top.map((row) => ({ name: row.name, value: Math.round(row.value * 100) / 100 })),
    ...(other > 0 ? [{ name: t.other, value: Math.round(other * 100) / 100 }] : [])];
  return {
    animationDuration: 260,
    color: [css('--s1'), css('--s2'), css('--s3'), css('--s4'), css('--s5'), css('--s6'), css('--s3'), css('--faint')],
    tooltip: { trigger: 'item', backgroundColor: css('--raised'), borderColor: css('--line'), textStyle: { color: css('--ink') }, valueFormatter: (value) => usd(Number(value)) },
    legend: { bottom: 0, left: 'center', itemWidth: 8, itemHeight: 8, itemGap: 12, textStyle: { color: css('--muted'), fontSize: 11, fontFamily: 'Geist Mono Variable' } },
    series: [{ type: 'pie', radius: ['48%', '72%'], center: ['50%', '43%'], minAngle: 2, label: { show: false }, itemStyle: { borderColor: css('--panel'), borderWidth: 2 }, data: values }],
  };
};

const trendOption = () => {
  const t = T();
  return {
    animationDuration: 260,
    grid: { left: 54, right: 48, top: 28, bottom: 28 },
    tooltip: { trigger: 'axis', backgroundColor: css('--raised'), borderColor: css('--line'), textStyle: { color: css('--ink') }, valueFormatter: (value) => typeof value === 'number' ? decimal(value) : String(value) },
    legend: { top: 0, right: 6, textStyle: { color: css('--muted'), fontSize: 11 }, itemWidth: 12, itemHeight: 6 },
    xAxis: { type: 'category', data: OV.monthly.map((row) => monthLabel(row.name)), axisLine: { lineStyle: { color: css('--line') } }, axisTick: { show: false }, axisLabel: { color: css('--muted'), fontFamily: 'Geist Mono Variable', fontSize: 11 } },
    yAxis: [
      { type: 'value', axisLabel: { color: css('--muted'), formatter: (value) => `$${value / 1e6}M`, fontSize: 10 }, splitLine: { lineStyle: { color: css('--line-soft') } } },
      { type: 'value', axisLabel: { color: css('--muted'), formatter: (value) => compact(value), fontSize: 10 }, splitLine: { show: false } },
    ],
    series: [
      { name: t.spendUsd, type: 'bar', data: OV.monthly.map((row) => Math.round(row.value)), barMaxWidth: 28, itemStyle: { color: css('--s1'), borderRadius: [2, 2, 0, 0] } },
      { name: t.receivedQty, type: 'line', yAxisIndex: 1, data: OV.monthly.map((row) => Math.round(row.quantity)), smooth: true, symbolSize: 6, lineStyle: { color: css('--s2'), width: 2 }, itemStyle: { color: css('--s2') } },
    ],
  };
};

const paintCharts = () => {
  donutChart.setOption(donutOption(), true);
  trendChart.setOption(trendOption(), true);
};

/* ── theme + clock + language chrome ────────────────────────────────────── */
const applyTheme = () => {
  document.documentElement.dataset.theme = state.theme;
  try { localStorage.setItem('gpr-theme', state.theme) } catch { /* preference only */ }
  $('#theme-btn').innerHTML = icon(state.theme === 'dark' ? 'sun' : 'moon', 18);
  $('#theme-btn').setAttribute('aria-label', state.theme === 'dark' ? T().toLight : T().toDark);
  requestAnimationFrame(() => {
    dispatchEvent(new Event('kit:theme'));
    if (donutChart) paintCharts();
    globeApi?.retheme();
  });
};

/* Every static string in the skeleton flows through data-i18n / -placeholder. */
const applyStatic = () => {
  const t = T();
  document.documentElement.lang = t.htmlLang;
  document.querySelectorAll('[data-i18n]').forEach((el) => {
    const key = el.dataset.i18n;
    if (typeof t[key] === 'string') el.textContent = t[key];
  });
  document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => { el.placeholder = t[el.dataset.i18nPlaceholder] });
  $('#lang-btn').innerHTML = `<span class="lang-opt${t === L.zh ? ' on' : ''}">中文</span><span class="lang-sep">/</span><span class="lang-opt${t === L.en ? ' on' : ''}">ENG</span>`;
  $('#lang-btn').setAttribute('aria-label', t.toggleLang);
};

const setLang = (next) => {
  if (!L[next] || next === state.lang) return;
  state.lang = next;
  NUM_LOCALE = next === 'zh' ? 'zh-TW' : 'en';
  try { localStorage.setItem('gpr-lang', next) } catch { /* preference only */ }
  closeCreate(); closeDrawer();
  applyStatic();
  populateOptions();
  render();
  tickClock();
};

const tickClock = () => {
  const clock = new Date();
  const el = $('#clock');
  el.dateTime = clock.toISOString();
  el.firstChild.textContent = clock.toLocaleTimeString(T().clockLocale, { hour12: false });
  $('#clock-date').textContent = clock.toLocaleDateString(T().clockLocale, { year: 'numeric', month: '2-digit', day: '2-digit' });
};

/* ── toast ──────────────────────────────────────────────────────────────── */
const showToast = (message) => {
  $('#toast-root').innerHTML =
    `<div class="toast" role="status"><span>${esc(message)}</span><button type="button" aria-label="close">${icon('x', 15)}</button></div>`;
  $('#toast-root').firstElementChild.querySelector('button').addEventListener('click', () => { $('#toast-root').innerHTML = '' });
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { $('#toast-root').innerHTML = '' }, 4200);
};

/* ── renderers ──────────────────────────────────────────────────────────── */
const populateSelect = (sel, values, placeholder, current) => {
  sel.innerHTML = `<option value="">${esc(placeholder)}</option>` + values.map((v) => `<option${v === current ? ' selected' : ''}>${esc(v)}</option>`).join('');
  sel.value = current ?? '';
};

function populateOptions() {
  const t = T();
  OPTS = buildOptions();
  populateSelect($('#f-entity'), OPTS.entities, t.allEntity, state.filters.entity);
  populateSelect($('#f-month'), OPTS.months, t.allMonth, state.filters.month);
  populateSelect($('#f-comm'), OPTS.categories, t.allComm, state.filters.stdComm);
  populateSelect($('#f-vendor'), OPTS.vendors, t.allVendor, state.filters.vendor);
  $('#f-status').innerHTML = `<option value="">${esc(t.allStatus)}</option><option>Consider</option><option>Ignore</option>`;
  $('#f-status').value = state.filters.status ?? '';
  const first = OPTS.months[0] ?? '2026-01', last = OPTS.months.at(-1) ?? '2026-09';
  $('#period-label').textContent = `${first} — ${last}`;
  $('#connection-text').textContent = t.embedded(OPTS.months.length);
  $('#footer-period').textContent = first.slice(0, 4) === last.slice(0, 4)
    ? `${first.replace('-', '.')}-${last.slice(5)}` : `${first.replace('-', '.')}–${last.replace('-', '.')}`;
}

const renderKpis = () => {
  const t = T();
  const d = OV.summary;
  const cards = [
    { label: t.kpiSpend, value: usd(d.totalUsd), note: `${d.monthChange >= 0 ? '+' : ''}${d.monthChange.toFixed(1)}% ${t.vsPrev}`, name: 'circle-dollar-sign', tone: d.monthChange >= 0 ? 'positive' : 'negative' },
    { label: t.kpiPo, value: integer(d.purchaseOrders), note: t.receiptLines(integer(d.records)), name: 'database', tone: 'neutral' },
    { label: t.kpiQty, value: compact(d.totalQuantity), note: t.kpiQtyNote, name: 'package-search', tone: 'neutral' },
    { label: t.kpiVendors, value: integer(d.suppliers), note: `${t.topVendor} ${d.topVendorShare.toFixed(1)}%`, name: 'earth', tone: 'neutral' },
    { label: t.kpiItems, value: integer(d.items), note: `Ignore ${d.ignoreShare.toFixed(1)}%`, name: 'activity', tone: 'neutral' },
  ];
  $('#kpi-grid').innerHTML = cards.map((c) => `<article class="kpi-card"><div class="kpi-label"><span>${esc(c.label)}</span>${icon(c.name, 16)}</div><strong>${c.value}</strong><small class="${c.tone}">${esc(c.note)}</small></article>`).join('');
};

const renderEntityList = () => {
  $('#entity-list').innerHTML = OV.entities.map((row, index) => `
    <button type="button" class="rank-row${state.filters.entity === row.name ? ' selected' : ''}" data-entity="${esc(row.name)}">
      <span class="rank-no">${String(index + 1).padStart(2, '0')}</span><strong>${esc(row.name)}</strong>
      <span class="rank-track"><i style="width:${(row.value / OV.entities[0].value * 100).toFixed(2)}%"></i></span>
      <span class="rank-value">${usd(row.value)}</span><span class="rank-share">${(row.value / OV.summary.totalUsd * 100).toFixed(1)}%</span>
    </button>`).join('');
};

const renderSuppliers = () => {
  const t = T();
  $('#supplier-hint').textContent = `TOP ${Math.min(8, OV.vendors.length)}`;
  $('#supplier-list').innerHTML = OV.vendors.slice(0, 8).map((row, index) => `
    <button type="button" class="${state.filters.vendor === row.name ? 'selected' : ''}" data-vendor="${esc(row.name)}">
      <span class="supplier-index">${String(index + 1).padStart(2, '0')}</span>
      <span class="supplier-name"><strong title="${esc(row.name)}">${esc(row.name)}</strong><small>${esc(t.lineCount(integer(row.records)))}</small></span>
      <span class="supplier-value"><strong>${usd(row.value)}</strong><small>${(row.value / OV.summary.totalUsd * 100).toFixed(1)}%</small></span>
    </button>`).join('');
};

const renderAlerts = () => {
  $('#signal-count').textContent = T().flags(OV.alerts.filter((a) => a.severity !== 'info').length);
  $('#alert-list').innerHTML = OV.alerts.map((alert) => `
    <article class="alert-item ${alert.severity}"><div>${icon('alert-triangle', 15)}<strong>${esc(alert.title)}</strong></div><p>${esc(alert.detail)}</p></article>`).join('');
};

const renderReadout = () => {
  const t = T();
  $('#readout-label').textContent = state.filters.entity ? t.currentEntity : t.networkOverview;
  $('#readout-entity').textContent = state.filters.entity ?? t.entitiesCount(OV.entities.length);
  $('#readout-value').textContent = usd(OV.summary.totalUsd);
  $('#trend-hint').textContent = state.filters.entity || state.filters.vendor || state.filters.month || t.currentScope;
  if (!globeReady) $('#globe-status').textContent = t.globeBuilding;
};

function renderRecords() {
  const t = T();
  const rows = filteredRows(state.filters, state.search);
  const totalPages = Math.max(1, Math.ceil(rows.length / 12));
  if (state.page > totalPages) state.page = totalPages;
  const sorted = sortRows(rows, state.sortBy, state.sortDir);
  const pageRows = sorted.slice((state.page - 1) * 12, state.page * 12);

  $('#records-count').textContent = t.lines(integer(rows.length));
  $('#records-tbody').innerHTML = pageRows.map((row) => `
    <tr data-id="${row.id}">
      <td class="mono">${esc(row.poRecDate)}</td><td class="mono po-cell">${esc(row.poNo)}</td>
      <td class="optional-col"><span class="entity-tag">${esc(row.entity)}</span></td>
      <td><strong class="vendor-cell" title="${esc(row.stdVendor)}">${esc(row.stdVendor)}</strong></td>
      <td><div class="item-cell"><span>${esc(row.itemCode)}</span><small title="${esc(row.partDesc)}">${esc(row.partDesc)}</small></div></td>
      <td class="optional-col"><span class="tag">${esc(row.stdComm)}</span></td>
      <td class="n mono optional-col">${integer(row.quantity)} <small>${esc(row.uom ?? '')}</small></td>
      <td class="n mono amount-cell">${decimal(row.totalUsd)}</td>
      <td><span class="status-pill ${String(row.status).toLowerCase()}">${esc(row.status)}</span></td>
      <td><button type="button" class="row-open" aria-label="${esc(t.viewPo(row.poNo))}">${icon('chevron-right', 17)}</button></td>
    </tr>`).join('');
  $('#records-tbody').onclick = (event) => {
    const tr = event.target.closest('tr[data-id]');
    if (tr) openDrawer(Number(tr.dataset.id));
  };

  $('#table-empty').hidden = rows.length > 0;
  $('#pager-label').textContent = t.pager(state.page, totalPages);
  $('#page-prev').disabled = state.page <= 1;
  $('#page-next').disabled = state.page >= totalPages;
  $('#export-btn').disabled = rows.length === 0;
  $('#export-label').textContent = t.exportNow;

  document.querySelectorAll('[data-caret]').forEach((slot) => {
    slot.innerHTML = slot.dataset.caret !== state.sortBy ? '' : icon(state.sortDir === 'asc' ? 'arrow-up' : 'arrow-down', 12);
  });
}

const renderTourBtn = () => {
  const t = T();
  const btn = $('#tour-btn');
  btn.classList.toggle('active', state.touring);
  btn.innerHTML = `${icon(state.touring ? 'pause' : 'play', 14)}${state.touring ? (t === L.zh ? '暫停巡覽' : 'Pause tour') : (t === L.zh ? '自動巡覽' : 'Auto tour')}`;
};

function render() {
  OV = buildOverview(filteredRows(state.filters, ''), state.lang);
  const active = Object.values(state.filters).filter(Boolean).length + (state.search ? 1 : 0);
  $('#reset-btn').disabled = !active;
  $('#reset-label').textContent = active ? `${T().reset} (${active})` : T().reset;
  $('#search-clear').hidden = !state.search;
  renderKpis(); renderEntityList(); renderSuppliers(); renderAlerts(); renderReadout(); renderTourBtn();
  paintCharts(); renderRecords();
  if (globeApi) { globeApi.setMarkers(globeMarkers(OV)); if (state.filters.entity) globeApi.focus(ENTITY_HUBS[state.filters.entity] ?? '上海') }
  const n = localEditCount();
  $('#footer-edits').textContent = String(n);
  $('#edits-reset').hidden = !n;
}

/* ── interactions ──────────────────────────────────────────────────────── */
function updateFilter(key, value) {
  stopTour();
  if (value) state.filters[key] = value; else delete state.filters[key];
  state.page = 1;
  render();
}
const resetFilters = () => {
  stopTour();
  state.filters = {}; state.search = ''; state.page = 1;
  $('#search-input').value = '';
  render();
};

function startTour() {
  state.touring = true;
  clearInterval(tourTimer);
  tourTimer = setInterval(() => {
    if (!OV.entities.length) return;
    const index = OV.entities.findIndex((item) => item.name === state.filters.entity);
    const next = OV.entities[(index + 1) % OV.entities.length];
    state.filters.entity = next.name;   /* direct set — user interaction goes through updateFilter, which stops the tour */
    state.page = 1;
    syncSelects();
    render();
  }, 7000);
  renderTourBtn();
}
const stopTour = () => { if (!state.touring) return; state.touring = false; clearInterval(tourTimer); renderTourBtn() };
const syncSelects = () => {
  $('#f-entity').value = state.filters.entity ?? '';
  $('#f-month').value = state.filters.month ?? '';
  $('#f-comm').value = state.filters.stdComm ?? '';
  $('#f-vendor').value = state.filters.vendor ?? '';
  $('#f-status').value = state.filters.status ?? '';
};

/* ── drawer (PurchaseDrawer.tsx port) ───────────────────────────────────── */
const findRow = (id) => ALL_ROWS.find((r) => r.id === id);

function closeDrawer() { selectedId = null; $('#drawer-root').innerHTML = ''; if (!createOpen) document.body.classList.remove('overlay-open') }

function openDrawer(id) {
  const t = T();
  selectedId = id;
  const row = findRow(id);
  if (!row) return;
  const details = [
    [t.dEntityMonth, `${row.entity} · ${row.month}`],
    [t.dStdVendor, row.stdVendor],
    [t.dVendor, row.vendorName],
    [t.dComm, `${row.stdComm} · ${row.commCode || '—'}`],
    [t.dFamily, row.productFamily || t.unassigned],
    [t.dQty, `${integer(row.quantity)} ${row.uom || ''}`],
    [t.dPrice, `$${decimal(row.unitPriceUsd)}`],
    [t.dUpdated, row.updatedAt ? new Date(row.updatedAt).toLocaleString(t.clockLocale) : t.snapshot],
  ];
  $('#drawer-root').innerHTML = `
    <div class="overlay drawer-overlay">
      <aside class="drawer" role="dialog" aria-modal="true" aria-labelledby="drawer-title">
        <header class="overlay-head">
          <div><span class="eyebrow">PURCHASE ORDER</span><h2 id="drawer-title">${esc(row.poNo)}</h2></div>
          <button class="icon-btn" id="drawer-close" type="button" aria-label="close">${icon('x', 19)}</button>
        </header>
        <div class="drawer-body">
          <div class="record-amount"><span>${esc(t.totalSpend)}</span><strong>${usd(row.totalUsd)}</strong><small>US$${decimal(row.totalUsd)}</small></div>
          <div class="record-item"><span class="tag">${esc(row.itemCode)}</span><h3>${esc(row.partDesc)}</h3><p>${esc(t.receivedOn(row.poRecDate))}</p></div>
          <dl class="detail-list">${details.map(([label, value]) => `<div><dt>${esc(label)}</dt><dd>${esc(value)}</dd></div>`).join('')}</dl>
          <div class="status-editor">
            <span class="eyebrow">ANALYSIS STATUS</span>
            <h3>${esc(t.statusH)}</h3>
            <p>${esc(t.statusP)}</p>
            <div class="status-options">
              ${['Consider', 'Ignore'].map((v) => `<button type="button" class="status-set${row.status === v ? ' selected' : ''}" data-status="${v}"${row.status === v ? ' disabled' : ''}>${icon(row.status === v ? 'check' : 'chevron-right', 16)}${v}</button>`).join('')}
            </div>
          </div>
        </div>
      </aside>
    </div>`;
  document.body.classList.add('overlay-open');
  $('#drawer-close').addEventListener('click', closeDrawer);
  $('#drawer-root .overlay').addEventListener('mousedown', (e) => { if (e.target === e.currentTarget) closeDrawer() });
  $('#drawer-root').querySelectorAll('.status-set').forEach((btn) => btn.addEventListener('click', () => {
    const value = btn.dataset.status;
    if (row.status === value) return;
    setRowStatus(row, value);
    render();
    if (selectedId === row.id) openDrawer(row.id);
    showToast(T().statusSaved(value));
  }));
  $('#drawer-close').focus();
}

/* ── create modal (CreatePurchaseModal.tsx port; saved to localStorage) ── */
function closeCreate() { createOpen = false; $('#create-root').innerHTML = ''; if (!selectedId) document.body.classList.remove('overlay-open') }

function openCreate() {
  const t = T();
  createOpen = true;
  const defaultMonth = OPTS.months.at(-1) ?? '2026-09';
  $('#create-root').innerHTML = `
    <div class="overlay">
      <section class="modal" role="dialog" aria-modal="true" aria-labelledby="create-title">
        <header class="overlay-head">
          <div><span class="eyebrow">NEW PROCUREMENT RECORD</span><h2 id="create-title">${esc(t.newPo)}</h2></div>
          <button class="icon-btn" id="create-close" type="button" aria-label="close">${icon('x', 19)}</button>
        </header>
        <form id="create-form">
          <div class="form-grid">
            <label>${esc(t.mPo)}<input name="poNo" required maxlength="32" placeholder="PO-2026-001" /></label>
            <label>${esc(t.mRecDate)}<input name="poRecDate" type="date" required value="${defaultMonth}-28" /></label>
            <label>${esc(t.fEntity)}<select name="entity" required>${OPTS.entities.map((v) => `<option>${esc(v)}</option>`).join('')}</select></label>
            <label>${esc(t.mDataMonth)}<input name="month" type="month" required value="${defaultMonth}" /></label>
            <label>${esc(t.mVendor)}<input name="vendorName" required placeholder="${esc(t.mVendorPh)}" /></label>
            <label>${esc(t.mStdVendor)}<input name="stdVendor" required placeholder="${esc(t.mStdPh)}" /></label>
            <label>${esc(t.mItem)}<input name="itemCode" required placeholder="${esc(t.mItemPh)}" /></label>
            <label>${esc(t.fComm)}<select name="stdComm" required>${OPTS.categories.map((v) => `<option>${esc(v)}</option>`).join('')}</select></label>
            <label class="span-2">${esc(t.mDesc)}<input name="partDesc" required placeholder="${esc(t.mDescPh)}" /></label>
            <label>${esc(t.mQty)}<input name="quantity" type="number" min="0.01" step="0.01" required value="1" /></label>
            <label>${esc(t.mPrice)}<input name="unitPriceUsd" type="number" min="0" step="0.0001" required value="0" /></label>
            <label>${esc(t.fStatus)}<select name="status"><option>Consider</option><option>Ignore</option></select></label>
            <div class="calculated"><span>${esc(t.mTotalCalc)}</span><strong id="calc-total">0.00</strong></div>
          </div>
          <p class="form-error" id="form-error" role="alert" hidden></p>
          <footer class="form-actions">
            <button type="button" class="btn" id="create-cancel">${esc(t.cancel)}</button>
            <button type="submit" class="btn primary">${icon('plus', 16)}${esc(t.saveLocal)}</button>
          </footer>
        </form>
      </section>
    </div>`;
  document.body.classList.add('overlay-open');
  const form = $('#create-form');
  const calc = () => {
    const total = Number(form.quantity.value) * Number(form.unitPriceUsd.value);
    $('#calc-total').textContent = Number.isFinite(total) ? decimal(total) : '—';
  };
  form.quantity.addEventListener('input', calc);
  form.unitPriceUsd.addEventListener('input', calc);
  $('#create-close').addEventListener('click', closeCreate);
  $('#create-cancel').addEventListener('click', closeCreate);
  $('#create-root .overlay').addEventListener('mousedown', (e) => { if (e.target === e.currentTarget) closeCreate() });
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const err = $('#form-error');
    err.hidden = true;
    const qty = Number(form.quantity.value), price = Number(form.unitPriceUsd.value);
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isFinite(price) || price < 0 || qty * price <= 0) {
      err.textContent = T().invalid;
      err.hidden = false;
      return;
    }
    const text = (name) => String(form[name].value ?? '').trim();
    const row = addPurchase({
      poNo: text('poNo'), poRecDate: text('poRecDate'), vendorName: text('vendorName'),
      stdVendor: text('stdVendor'), itemCode: text('itemCode'), partDesc: text('partDesc'),
      entity: text('entity'), stdComm: text('stdComm'), month: text('month'),
      status: text('status') === 'Ignore' ? 'Ignore' : 'Consider',
      quantity: qty, unitPriceUsd: price, totalUsd: qty * price,
    });
    closeCreate();
    populateOptions();
    render();
    showToast(T().poAdded(row.poNo));
  });
  $('#create-close').focus();
}

/* ── CSV export (App.tsx exportCsv port: BOM + CRLF, English column names
   always — they match the Excel master's own headers) ───────────────────── */
function exportCsv() {
  const rows = exportRows(filteredRows(state.filters, state.search));
  const header = ['PO_NO', 'REC_DATE', 'ENTITY', 'MONTH', 'VENDOR', 'ITEM_CODE', 'PART_DESC', 'STD_COMM', 'QTY', 'UOM', 'TOTAL_USD', 'STATUS', 'PRODUCT_FAMILY'];
  const lines = rows.map((row) => [row.poNo, row.poRecDate, row.entity, row.month, row.stdVendor, row.itemCode,
    row.partDesc, row.stdComm, row.quantity, row.uom, row.totalUsd, row.status, row.productFamily]);
  const csv = '﻿' + [header, ...lines].map((line) => line.map(csvCell).join(',')).join('\r\n');
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `global-purchase-${new Date().toISOString().slice(0, 10)}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
  showToast(T().exported(integer(rows.length)));
}

/* ── boot ───────────────────────────────────────────────────────────────── */
function boot() {
  document.querySelectorAll('[data-icon]').forEach((el) => { el.outerHTML = icon(el.dataset.icon, Number(el.dataset.size) || 16) });
  $('#search-clear').innerHTML = icon('x', 14);

  applyStatic();
  applyTheme();
  tickClock();
  setInterval(tickClock, 1000);

  populateOptions();
  donutChart = echarts.init($('#donut-chart'));
  trendChart = echarts.init($('#trend-chart'));
  /* clicking a commodity slice filters the whole screen (App.tsx parity) */
  donutChart.on('click', (params) => {
    if (params.name && params.name !== T().other) updateFilter('stdComm', String(params.name));
  });
  addEventListener('resize', () => { donutChart.resize(); trendChart.resize() });

  /* filter selects */
  [['#f-entity', 'entity'], ['#f-month', 'month'], ['#f-comm', 'stdComm'], ['#f-vendor', 'vendor'], ['#f-status', 'status']]
    .forEach(([sel, key]) => $(sel).addEventListener('change', (e) => updateFilter(key, e.target.value)));
  $('#reset-btn').addEventListener('click', resetFilters);
  $('#table-reset').addEventListener('click', resetFilters);

  /* rank / supplier click-through */
  $('#entity-list').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-entity]'); if (!btn) return;
    updateFilter('entity', state.filters.entity === btn.dataset.entity ? '' : btn.dataset.entity);
  });
  $('#supplier-list').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-vendor]'); if (!btn) return;
    updateFilter('vendor', state.filters.vendor === btn.dataset.vendor ? '' : btn.dataset.vendor);
  });

  /* theme, language, create, tour */
  $('#theme-btn').addEventListener('click', () => { state.theme = state.theme === 'dark' ? 'light' : 'dark'; applyTheme() });
  $('#lang-btn').addEventListener('click', () => setLang(state.lang === 'zh' ? 'en' : 'zh'));
  $('#create-btn').addEventListener('click', openCreate);
  $('#tour-btn').addEventListener('click', () => { if (state.touring) stopTour(); else startTour() });
  $('#edits-reset').addEventListener('click', clearOverlay);

  /* escape closes topmost overlay */
  addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    if ($('#create-root').firstElementChild) closeCreate();
    else if ($('#drawer-root').firstElementChild) closeDrawer();
  });

  /* records */
  $('#search-form').addEventListener('submit', (e) => {
    e.preventDefault();
    state.search = $('#search-input').value.trim();
    state.page = 1;
    render();
  });
  $('#search-input').addEventListener('input', () => { $('#search-clear').hidden = !$('#search-input').value.trim() });
  $('#search-clear').addEventListener('click', () => { $('#search-input').value = ''; state.search = ''; state.page = 1; render() });
  document.querySelectorAll('[data-sort]').forEach((btn) => btn.addEventListener('click', () => {
    const col = btn.dataset.sort;
    if (state.sortBy === col) state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
    else { state.sortBy = col; state.sortDir = 'desc' }
    state.page = 1;
    renderRecords();
  }));
  $('#page-prev').addEventListener('click', () => { state.page = Math.max(1, state.page - 1); renderRecords() });
  $('#page-next').addEventListener('click', () => { state.page += 1; renderRecords() });
  $('#export-btn').addEventListener('click', exportCsv);

  render();
  ensureGlobe();
}

boot();
