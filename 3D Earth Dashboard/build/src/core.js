/* ── core: embedded data, local overlay, aggregation engine ────────────────
 * Ported from the builder app's api/queries/procurement.ts. The MySQL table
 * "purchases" becomes the frozen snapshot in #gpr-data; the tRPC mutations
 * (status edit, create) become a localStorage overlay, so nothing is ever
 * written back to the Excel master — this file stays a shareable snapshot. */

/* Design-token reader (same contract the original React app used). */
const css = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

/* ── number/date formatting — port of src/lib/format.ts, locale follows the
   UI language so EN mode shows "180M" not "1.8億" (main.js setLang updates it) */
let NUM_LOCALE = 'zh-TW';
const usd = (value, digits = 2) => {
  const absolute = Math.abs(value);
  if (absolute >= 1e9) return `$${(value / 1e9).toFixed(digits)}B`;
  if (absolute >= 1e6) return `$${(value / 1e6).toFixed(digits)}M`;
  if (absolute >= 1e3) return `$${(value / 1e3).toFixed(digits)}K`;
  return `$${value.toFixed(0)}`;
};
const compact = (value) => new Intl.NumberFormat(NUM_LOCALE, { notation: 'compact', maximumFractionDigits: 1 }).format(value);
const integer = (value) => new Intl.NumberFormat(NUM_LOCALE, { maximumFractionDigits: 0 }).format(value);
const decimal = (value) => new Intl.NumberFormat(NUM_LOCALE, { maximumFractionDigits: 2 }).format(value);
const monthLabel = (month) => { const [year, value] = month.split('-'); return `${year.slice(2)}.${value}` };
const csvCell = (value) => {
  const text = value === null || value === undefined ? '' : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
};

/* Entity → globe hub city. Mappings confirmed with the report owner.
 * Values must match CITY keys in globe.js exactly (that table keeps the
 * builder app's Simplified spellings — internal coordinates, not UI text). */
const ENTITY_HUBS = { MIL: '曼谷', MHK: '香港', STC: '上海', TVS: '罗马', MZQ: '深圳', MAZ: '利马', RACO: '广州' };

/* ── snapshot load ───────────────────────────────────────────────────────── */
const COLS = ['poNo', 'poRecDate', 'poIssDate', 'vendorId', 'vendorName', 'itemCode', 'partDesc',
  'commCode', 'quantity', 'uom', 'currency', 'unitPriceUsd', 'totalUsd', 'entity', 'stdVendor',
  'stdComm', 'month', 'status', 'productFamily', 'productFamilyCode'];

const readJson = (id) => JSON.parse(document.getElementById(id).textContent);
const RAW = readJson('gpr-data');
(() => {
  const expected = ['PO_NO', 'PO_REC_DATE', 'PO_ISS_DATE', 'VENDOR_ID', 'VENDOR_NAME', 'ITEM_CODE', 'PART_DESC',
    'COMM_CODE', 'REC_QTY', 'UOM', 'PO_CUR', 'PO_PRICE_USD', 'TOT_PO_PRICE_USD', 'DB', 'Std Vendor Name',
    'Std Comm', 'yyyy-mm', 'ignore / consider', 'ProdFamiDesc', 'ProdFamiCode'];
  if (RAW.fields.join('|') !== expected.join('|')) {
    throw new Error(`內嵌資料欄位與預期不符：${RAW.fields.join(',')}`);
  }
})();

const BASE_ROWS = RAW.rows.map((r, i) => {
  const row = { id: i + 1, updatedAt: null };
  COLS.forEach((key, c) => { row[key] = r[c] });
  row.poNo = String(row.poNo);
  row.quantity = Number(row.quantity) || 0;
  row.unitPriceUsd = Number(row.unitPriceUsd) || 0;
  row.totalUsd = Number(row.totalUsd) || 0;
  return row;
});

/* The fingerprint invalidates stored local edits if the snapshot is rebuilt
 * with different rows (monthly republish shifts positional ids). */
const FINGERPRINT = `${BASE_ROWS.length}:${BASE_ROWS.reduce((s, r) => s + r.quantity, 0).toFixed(2)}`;

const OVERLAY_KEY = 'gpr-local-v1';
let overlay = { fp: FINGERPRINT, edits: {}, added: [] };
try {
  const saved = JSON.parse(localStorage.getItem(OVERLAY_KEY) || 'null');
  if (saved && saved.fp === FINGERPRINT) overlay = saved;
} catch { /* corrupt or unavailable storage — start clean */ }

const saveOverlay = () => { try { localStorage.setItem(OVERLAY_KEY, JSON.stringify(overlay)) } catch { /* private mode etc. */ } };

let nextAddedId = -1;
const ALL_ROWS = BASE_ROWS.map((r) => {
  const st = overlay.edits[r.id];
  if (st) r.status = st;
  return r;
});
for (const a of overlay.added) { const row = { ...a }; ALL_ROWS.push(row); if (row.id <= nextAddedId) nextAddedId = row.id - 1 }

function setRowStatus(row, status) {
  row.status = status;
  if (row.id > 0) overlay.edits[row.id] = status;
  else { const a = overlay.added.find((x) => x.id === row.id); if (a) a.status = status }
  saveOverlay();
}

function addPurchase(form) {
  const row = {
    id: nextAddedId--, poNo: String(form.poNo), poRecDate: form.poRecDate, poIssDate: null, vendorId: null,
    vendorName: form.vendorName, itemCode: form.itemCode, partDesc: form.partDesc, commCode: form.stdComm,
    quantity: form.quantity, uom: 'PCS', currency: 'USD', unitPriceUsd: form.unitPriceUsd,
    totalUsd: form.totalUsd, entity: form.entity, stdVendor: form.stdVendor, stdComm: form.stdComm,
    month: form.month, status: form.status, productFamily: null, productFamilyCode: null,
    updatedAt: new Date().toISOString(),
  };
  ALL_ROWS.push(row);
  overlay.added.push(row);
  saveOverlay();
  return row;
}

const clearOverlay = () => { try { localStorage.removeItem(OVERLAY_KEY) } catch { /* ignore */ }; location.reload() };
const localEditCount = () => Object.keys(overlay.edits).length + overlay.added.length;

/* ── filtering — mirrors where()/filterConditions() on the backend ──────── */
function matches(row, f, search) {
  if (f.entity && row.entity !== f.entity) return false;
  if (f.month && row.month !== f.month) return false;
  if (f.stdComm && row.stdComm !== f.stdComm) return false;
  if (f.vendor && row.stdVendor !== f.vendor) return false;
  if (f.status && row.status !== f.status) return false;
  if (search) {
    const term = search.toLowerCase();
    const hay = `${row.poNo} ${row.itemCode} ${row.partDesc} ${row.vendorName} ${row.stdVendor}`.toLowerCase();
    if (!hay.includes(term)) return false;
  }
  return true;
}
const filteredRows = (f, search) => ALL_ROWS.filter((r) => matches(r, f, search));

/* ── overview aggregation — mirrors getProcurementOverview() ────────────── */
/* Procurement-signal copy, both languages (values are interpolated at build). */
const ALERT_TEXT = {
  zh: {
    vendorShare: (share) => `最大供應商占比 ${share}%`, vendorDetail: (name, money) => `${name} · ${money}，建議檢視供應集中度`,
    monthTitle: (month, dir, pct) => `${month}採購額 ${dir} ${pct}%`, monthDetail: (from, to) => `${from} → ${to}，較前一資料月份`,
    up: '成長', down: '下降', latest: '最新月份',
    entityTitle: (name) => `${name} 為最大採購實體`, entityDetail: (share, money) => `占目前篩選總額 ${share}% · ${money}`,
    ignoreTitle: (share) => `Ignore 金額占比 ${share}%`, ignoreDetail: '可切換 Consider／Ignore 狀態，核對報表分析口徑',
  },
  en: {
    vendorShare: (share) => `Top vendor share ${share}%`, vendorDetail: (name, money) => `${name} · ${money} — review supply concentration`,
    monthTitle: (month, dir, pct) => `${month} spend ${dir} ${pct}%`, monthDetail: (from, to) => `${from} → ${to} vs previous month`,
    up: 'up', down: 'down', latest: 'latest month',
    entityTitle: (name) => `${name} is the largest purchasing entity`, entityDetail: (share, money) => `${share}% of current scope · ${money}`,
    ignoreTitle: (share) => `Ignore share ${share}%`, ignoreDetail: 'Toggle Consider/Ignore to align the analysis scope',
  },
};
function sumBy(rows, keyFn) {
  const map = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    const cur = map.get(k) || { name: k, value: 0, records: 0, quantity: 0 };
    cur.value += r.totalUsd; cur.records += 1; cur.quantity += r.quantity;
    map.set(k, cur);
  }
  return [...map.values()];
}
const formatUsdM = (value) => `US$${(value / 1e6).toFixed(2)}M`;

function buildOverview(rows, lang = 'zh') {
  const uniq = (fn) => { const s = new Set(); for (const r of rows) s.add(fn(r)); return s };
  const totalUsd = rows.reduce((s, r) => s + r.totalUsd, 0);
  const totalQuantity = rows.reduce((s, r) => s + r.quantity, 0);
  const summary = {
    totalUsd, totalQuantity, records: rows.length,
    purchaseOrders: uniq((r) => r.poNo).size, suppliers: uniq((r) => r.stdVendor).size,
    items: uniq((r) => r.itemCode).size,
  };
  const monthly = sumBy(rows, (r) => r.month).sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  const byValueDesc = (list) => list.sort((a, b) => b.value - a.value);
  const entities = byValueDesc(sumBy(rows, (r) => r.entity));
  const categories = byValueDesc(sumBy(rows, (r) => r.stdComm));
  const vendors = byValueDesc(sumBy(rows, (r) => r.stdVendor)).slice(0, 10);
  const ignoredValue = rows.reduce((s, r) => s + (r.status === 'Ignore' ? r.totalUsd : 0), 0);

  const ignoreShare = totalUsd ? ignoredValue / totalUsd * 100 : 0;
  const topVendorShare = totalUsd && vendors.length ? vendors[0].value / totalUsd * 100 : 0;
  const topEntityShare = totalUsd && entities.length ? entities[0].value / totalUsd * 100 : 0;
  const last = monthly.at(-1), previous = monthly.at(-2);
  const monthChange = last && previous && previous.value ? (last.value - previous.value) / previous.value * 100 : 0;
  summary.monthChange = monthChange; summary.ignoreShare = ignoreShare; summary.topVendorShare = topVendorShare;

  const A = ALERT_TEXT[lang] ?? ALERT_TEXT.zh;
  const alerts = [];
  if (summary.records) {
    alerts.push({
      id: 'vendor-concentration', severity: topVendorShare > 20 ? 'critical' : 'warning',
      title: A.vendorShare(topVendorShare.toFixed(1)),
      detail: A.vendorDetail(vendors[0]?.name ?? '—', formatUsdM(vendors[0]?.value ?? 0)),
    });
    alerts.push({
      id: 'month-change', severity: monthChange < -10 ? 'critical' : monthChange < 0 ? 'warning' : 'info',
      title: A.monthTitle(last?.name ?? A.latest, monthChange >= 0 ? A.up : A.down, Math.abs(monthChange).toFixed(1)),
      detail: A.monthDetail(formatUsdM(previous?.value ?? 0), formatUsdM(last?.value ?? 0)),
    });
    alerts.push({
      id: 'entity-concentration', severity: topEntityShare > 35 ? 'warning' : 'info',
      title: A.entityTitle(entities[0]?.name ?? '—'),
      detail: A.entityDetail(topEntityShare.toFixed(1), formatUsdM(entities[0]?.value ?? 0)),
    });
    alerts.push({
      id: 'ignored-spend', severity: ignoreShare > 30 ? 'warning' : 'info',
      title: A.ignoreTitle(ignoreShare.toFixed(1)),
      detail: A.ignoreDetail,
    });
  }
  return { summary, monthly, entities, categories, vendors, alerts };
}

const buildOptions = () => ({
  entities: [...new Set(ALL_ROWS.map((r) => r.entity))].sort(),
  months: [...new Set(ALL_ROWS.map((r) => r.month))].sort(),
  categories: [...new Set(ALL_ROWS.map((r) => r.stdComm))].sort(),
  vendors: [...new Set(ALL_ROWS.map((r) => r.stdVendor))].sort(),
});

/* ── records — mirrors listPurchases()/exportPurchases() ────────────────── */
const SORT_KEYS = { poRecDate: 'poRecDate', totalUsd: 'totalUsd', quantity: 'quantity', poNo: 'poNo' };
function sortRows(rows, sortBy, sortDir) {
  const dir = sortDir === 'asc' ? 1 : -1;
  const numeric = sortBy === 'totalUsd' || sortBy === 'quantity';
  return rows.slice().sort((a, b) => {
    const x = a[sortBy], y = b[sortBy];
    const c = numeric ? (x - y) : String(x).localeCompare(String(y));
    return c !== 0 ? c * dir : b.id - a.id;   /* SQL tie-breaker: ORDER BY col, id DESC */
  });
}
/* backend exportPurchases(): ORDER BY poRecDate ASC, id ASC */
const exportRows = (rows) => rows.slice().sort((a, b) =>
  a.poRecDate < b.poRecDate ? -1 : a.poRecDate > b.poRecDate ? 1 : a.id - b.id);
