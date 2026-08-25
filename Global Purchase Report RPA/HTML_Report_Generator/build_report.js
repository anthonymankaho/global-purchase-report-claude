const fs = require('fs');
const path = require('path');

// build_report.js <workDir> <outputHtmlPath>
// workDir must contain data_sheet.json (Data sheet, full extraction) and
// exrate_platinum_usd.json (Exchange Rate's Platium Buy Rate USD table only),
// both produced by Generate-PurchaseReportHtml.ps1 in this folder.
const SCRIPT_DIR = __dirname; // vendor/ (jquery, DataTables, Chart.js) lives here
const WORK_DIR = process.argv[2];
const OUT_PATH = process.argv[3];
if (!WORK_DIR || !OUT_PATH) {
  console.error('Usage: node build_report.js <workDir> <outputHtmlPath>');
  process.exit(1);
}
const DIR = WORK_DIR;

function toDataUri(filePath, mime) {
  return 'data:' + mime + ';base64,' + fs.readFileSync(filePath).toString('base64');
}
function neutralize(jsonStr) { return jsonStr.replace(/<\/script/gi, '<\\/script'); }

// ---- Excel serial date -> yyyy-MM-dd ----
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
function excelSerialToDateStr(serial) {
  if (typeof serial !== 'number') return serial;
  const ms = EXCEL_EPOCH_MS + Math.round(serial) * 86400000;
  const d = new Date(ms);
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
}

// ---- load Data JSON (header + rows), convert date columns ----
let rawJson = fs.readFileSync(path.join(DIR, 'data_sheet.json'), 'utf8');
if (rawJson.charCodeAt(0) === 0xFEFF) rawJson = rawJson.slice(1);
const dataSheet = JSON.parse(rawJson);
const H = dataSheet.header;
const COL = {}; // name -> index
H.forEach((name, i) => { COL[name] = i; });
for (const row of dataSheet.rows) {
  if (typeof row[COL['PO_REC_DATE']] === 'number') row[COL['PO_REC_DATE']] = excelSerialToDateStr(row[COL['PO_REC_DATE']]);
  if (typeof row[COL['PO_ISS_DATE']] === 'number') row[COL['PO_ISS_DATE']] = excelSerialToDateStr(row[COL['PO_ISS_DATE']]);
}
console.log('Data loaded:', dataSheet.rows.length, 'rows,', H.length, 'cols');

// ---- vendor assets ----
const jqueryJs = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'jquery.min.js'), 'utf8');
const dataTablesJs = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'dataTables.min.js'), 'utf8');
const dataTablesCss = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'dataTables.min.css'), 'utf8');
const chartJs = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'chart.min.js'), 'utf8');

// =========================================================================
// Build a generic DataTables page for a (possibly filtered) subset of rows
// =========================================================================
function buildDataTablesHtml(header, rows, title, scrollX, decimalCols) {
  if (scrollX === undefined) scrollX = true;
  const headerJsonStr = neutralize(JSON.stringify(header));
  const rowsJsonStr = neutralize(JSON.stringify(rows));
  // decimalCols: column indices to always show to 4 decimal places (e.g. rate
  // tables), while sorting/filtering still use the raw numeric value.
  const columnDefsStr = (decimalCols && decimalCols.length)
    ? `[{ targets: ${JSON.stringify(decimalCols)}, render: function (data, type) { if (type === 'display' && typeof data === 'number') return data.toFixed(4); return data; } }]`
    : '[]';
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${title}</title>
<style>${dataTablesCss}
:root {
  --bg: #f4f6f9; --card-bg: #ffffff; --text: #222222; --text-muted: #666666;
  --accent: #005daa; --border: #cccccc; --border-light: #eeeeee; --shadow: rgba(0,0,0,0.1);
  --input-bg: #ffffff; --hover-bg: #f5f5f5;
}
body[data-theme="dark"] {
  --bg: #1a1d23; --card-bg: #262a33; --text: #e4e6eb; --text-muted: #a0a4ab;
  --accent: #6fa8ff; --border: #444a57; --border-light: #333844; --shadow: rgba(0,0,0,0.4);
  --input-bg: #1f232b; --hover-bg: #2f3440;
}
body { font-family: Calibri, Arial, sans-serif; font-size: 12px; margin: 8px; background: var(--bg); color: var(--text); }
table.dataTable { color: var(--text); }
th, td { white-space: nowrap; }
.summary-bar { margin-bottom: 8px; font-weight: bold; color: var(--accent); }
#exportBtn { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); }

body[data-theme="dark"] table.dataTable tbody tr { background-color: var(--card-bg); color: var(--text); }
body[data-theme="dark"] table.dataTable tbody tr.odd,
body[data-theme="dark"] table.dataTable tbody tr.even { background-color: var(--card-bg); }
body[data-theme="dark"] table.dataTable.stripe tbody tr.odd,
body[data-theme="dark"] table.dataTable.display tbody tr.odd { background-color: var(--hover-bg); }
body[data-theme="dark"] table.dataTable tbody tr:hover { background-color: var(--hover-bg) !important; }
body[data-theme="dark"] table.dataTable thead th,
body[data-theme="dark"] table.dataTable thead td { color: var(--text); border-bottom-color: var(--border); }
body[data-theme="dark"] table.dataTable tbody td { border-top-color: var(--border-light); }
body[data-theme="dark"] .dataTables_wrapper .dataTables_length,
body[data-theme="dark"] .dataTables_wrapper .dataTables_filter,
body[data-theme="dark"] .dataTables_wrapper .dataTables_info,
body[data-theme="dark"] .dataTables_wrapper .dataTables_processing,
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate { color: var(--text); }
body[data-theme="dark"] .dataTables_wrapper .dataTables_length select,
body[data-theme="dark"] .dataTables_wrapper .dataTables_filter input { background-color: var(--input-bg); color: var(--text); border-color: var(--border); }
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button { color: var(--text) !important; }
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button.current { background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important; }
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button:hover { background: var(--hover-bg) !important; color: var(--text) !important; }
</style>
</head>
<body>
<div class="summary-bar">${title} - ${rows.length.toLocaleString()} rows</div>
<button id="exportBtn">Export CSV (filtered)</button>
<table id="dataTable" class="display">
  <thead><tr id="dataTableHeadRow"></tr></thead>
  <tbody></tbody>
</table>
<script>${jqueryJs}</script>
<script>${dataTablesJs}</script>
<script>
var HEADER = ${headerJsonStr};
var ROWS = ${rowsJsonStr};
// Theme is controlled by the outer report shell (single global toggle), which
// calls this function directly on the iframe's window after (re)loading it.
function applyTheme(theme) {
  document.body.setAttribute('data-theme', theme);
}
$(function () {
  var headRow = document.getElementById('dataTableHeadRow');
  HEADER.forEach(function (h) {
    var th = document.createElement('th');
    th.textContent = h;
    headRow.appendChild(th);
  });
  var table = $('#dataTable').DataTable({
    data: ROWS, deferRender: true, pageLength: 50,
    lengthMenu: [25, 50, 100, 500, 1000], scrollX: ${scrollX ? 'true' : 'false'}, autoWidth: false,
    columnDefs: ${columnDefsStr}
  });
  document.getElementById('exportBtn').onclick = function () {
    var rowsData = table.rows({ search: 'applied' }).data();
    var lines = [HEADER.join(',')];
    for (var i = 0; i < rowsData.length; i++) {
      var r = rowsData[i];
      lines.push(r.map(function (v) {
        if (v === null || v === undefined) return '';
        var s = String(v);
        if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0) s = '"' + s.replace(/"/g, '""') + '"';
        return s;
      }).join(','));
    }
    var csv = '\\ufeff' + lines.join('\\r\\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'export.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
});
</script>
</body>
</html>`;
}

// =========================================================================
// GLID / Fittings / Packaging / Coating: filtered subsets of Data
// =========================================================================
const FILTER_SHEETS = [
  { name: 'GLID', stdComm: ['GLID'] },
  { name: 'Fittings', stdComm: ['HDL-', 'KNOB', 'LHDL', 'SHDL'] },
  { name: 'Packaging', stdComm: ['BHRE', 'LABL', 'STKR', 'U&C-', 'WACD'] },
  { name: 'Coating', stdComm: ['EXM-', 'INM-'] },
];

const stdCommIdx = COL['Std Comm'];
const statusIdx = COL['ignore / consider'];

const filteredSheetsHtml = {};

// =========================================================================
// GLID / Fittings / Packaging / Coating: item-level (ITEM_CODE x DB x Vendor)
// breakdown table, computed dynamically in the browser from the flat
// transaction-level Data rows (not from Excel's static pivot cache, which is
// already aggregated across all months and can't respond to the Month
// filter). Aggregation formula validated cell-for-cell against each sheet's
// real Excel PivotTable output: Avg PO Price(USD) = SUM(TOT_PO_PRICE_USD) /
// SUM(REC_QTY) (quantity-weighted, not a simple mean of unit prices); Max/Min
// Vendor+Price = the vendor and PO_PRICE_USD at that item's price extremes
// across all its lines, regardless of DB/Vendor grouping.
// =========================================================================
const PIVOT_HEADER = ['ITEM_CODE', 'Std Part Desc', 'Max_DB', 'Max Vendor', 'Max_PO_Price(USD)', 'Min_DB', 'Min Vendor', 'Min_PO_Price(USD)', '', 'DB', 'Vendor', 'Avg PO Price(USD)', 'Tot PO Price(USD)'];

function buildPivotPageHtml(pageName, stdCommLabel, itemRows, dbValues, monthValues, stdCommValues, vendorValues, itemCodeValues, statusValues, header, addColDivider) {
  const chartRowsJsonStr = neutralize(JSON.stringify(itemRows));
  const dbValuesJsonStr = neutralize(JSON.stringify(dbValues));
  const monthValuesJsonStr = neutralize(JSON.stringify(monthValues));
  const stdCommValuesJsonStr = neutralize(JSON.stringify(stdCommValues));
  const vendorValuesJsonStr = neutralize(JSON.stringify(vendorValues));
  const itemCodeValuesJsonStr = neutralize(JSON.stringify(itemCodeValues));
  const statusValuesJsonStr = neutralize(JSON.stringify(statusValues));
  const headerJsonStr = neutralize(JSON.stringify(header));

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${pageName}</title>
<style>${dataTablesCss}
  :root {
    --bg: #f4f6f9; --card-bg: #ffffff; --text: #222222; --text-muted: #666666;
    --accent: #005daa; --border: #cccccc; --border-light: #eeeeee; --shadow: rgba(0,0,0,0.1);
    --input-bg: #ffffff; --hover-bg: #f5f5f5;
  }
  body[data-theme="dark"] {
    --bg: #1a1d23; --card-bg: #262a33; --text: #e4e6eb; --text-muted: #a0a4ab;
    --accent: #6fa8ff; --border: #444a57; --border-light: #333844; --shadow: rgba(0,0,0,0.4);
    --input-bg: #1f232b; --hover-bg: #2f3440;
  }
  body { font-family: Calibri, Arial, sans-serif; font-size: 12px; margin: 8px; background: var(--bg); color: var(--text); }
  h3 { color: var(--accent); margin: 0 0 8px 0; font-size: 14px; }
  .page-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 8px; }
  #themeToggle { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--card-bg); color: var(--text); font-size: 12px; }
  .summary-bar { margin-bottom: 8px; font-weight: bold; color: var(--accent); font-size: 13px; }
  .filters { display: flex; gap: 16px; flex-wrap: wrap; background: var(--card-bg); padding: 10px; border-radius: 6px; margin-bottom: 12px; box-shadow: 0 1px 3px var(--shadow); }
  .filters label.fl { display: block; font-size: 11px; color: var(--text-muted); margin-bottom: 4px; }
  .charts { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 16px; }
  .chart-box { background: var(--card-bg); padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 420px; }
  canvas { max-height: 300px; }
  .table-box { background: var(--card-bg); padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); }
  table.dataTable { color: var(--text); }
  th, td { white-space: nowrap; }
  #exportBtn, #resetBtn { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); }
  body.col-divider #dataTable th:nth-child(2), body.col-divider #dataTable td:nth-child(2),
  body.col-divider #dataTable th:nth-child(5), body.col-divider #dataTable td:nth-child(5),
  body.col-divider #dataTable th:nth-child(8), body.col-divider #dataTable td:nth-child(8) { border-right: 2px solid var(--border); }
  #dataTable td:nth-child(5), #dataTable td:nth-child(8), #dataTable td:nth-child(12), #dataTable td:nth-child(13) { text-align: right; }
  #dataTable th:nth-child(9), #dataTable td:nth-child(9) { text-align: center; }

  .ms { position: relative; width: 180px; }
  .ms-toggle { width: 100%; text-align: left; padding: 5px 8px; background: var(--input-bg); color: var(--text); border: 1px solid var(--border); border-radius: 4px; cursor: pointer; font-size: 12px; }
  .ms-panel { position: absolute; z-index: 20; top: 100%; left: 0; background: var(--card-bg); border: 1px solid var(--border); border-radius: 4px; box-shadow: 0 2px 8px var(--shadow); max-height: 260px; overflow-y: auto; width: 240px; padding: 6px; }
  .ms-panel label { display: block; font-size: 12px; padding: 2px 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text); }
  .ms-panel label.ms-all-row { border-bottom: 1px solid var(--border-light); margin-bottom: 4px; padding-bottom: 6px; font-weight: bold; }
  .item-toggle { cursor: pointer; display: inline-block; width: 16px; text-align: center; border: 1px solid var(--border); border-radius: 2px; font-weight: bold; user-select: none; }
  .item-toggle:hover { background: var(--hover-bg); }

  body[data-theme="dark"] table.dataTable tbody tr { background-color: var(--card-bg); color: var(--text); }
  body[data-theme="dark"] table.dataTable tbody tr.odd,
  body[data-theme="dark"] table.dataTable tbody tr.even { background-color: var(--card-bg); }
  body[data-theme="dark"] table.dataTable.stripe tbody tr.odd,
  body[data-theme="dark"] table.dataTable.display tbody tr.odd { background-color: var(--hover-bg); }
  body[data-theme="dark"] table.dataTable tbody tr:hover { background-color: var(--hover-bg) !important; }
  body[data-theme="dark"] table.dataTable thead th,
  body[data-theme="dark"] table.dataTable thead td { color: var(--text); border-bottom-color: var(--border); }
  body[data-theme="dark"] table.dataTable tbody td { border-top-color: var(--border-light); }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_length,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_filter,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_info,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_processing,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate { color: var(--text); }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_length select,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_filter input { background-color: var(--input-bg); color: var(--text); border-color: var(--border); }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button { color: var(--text) !important; }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button.current { background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important; }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button:hover { background: var(--hover-bg) !important; color: var(--text) !important; }
</style>
</head>
<body class="${addColDivider ? 'col-divider' : ''}">
<div class="page-header">
  <h3>${pageName} - Std Comm: ${stdCommLabel}</h3>
</div>
<div class="filters">
  <div><label class="fl">Month</label><div class="ms" id="ms-month"></div></div>
  <div><label class="fl">DB (Entity)</label><div class="ms" id="ms-db"></div></div>
  <div><label class="fl">Std Comm</label><div class="ms" id="ms-stdcomm"></div></div>
  <div><label class="fl">Std Vendor Name</label><div class="ms" id="ms-vendor"></div></div>
  <div><label class="fl">Item Code</label><div class="ms" id="ms-itemcode"></div></div>
  <div><label class="fl">Ignore / Consider</label><div class="ms" id="ms-status"></div></div>
  <div style="align-self:flex-end"><button id="resetBtn">Reset filters</button></div>
</div>
<div class="charts">
  <div class="chart-box"><canvas id="chartGlidMonth"></canvas></div>
  <div class="chart-box"><canvas id="chartGlidDbPie"></canvas></div>
</div>
<div class="table-box">
  <div class="summary-bar" id="tableSummary"></div>
  <button id="exportBtn">Export CSV (filtered)</button>
  <table id="dataTable" class="display">
    <thead><tr id="dataTableHeadRow"></tr></thead>
    <tbody></tbody>
  </table>
</div>
<script>${jqueryJs}</script>
<script>${dataTablesJs}</script>
<script>${chartJs}</script>
<script>
var CHART_ROWS = ${chartRowsJsonStr}; // [DB, month, TOT_PO_PRICE_USD, ITEM_CODE, Std Part Desc, Std Vendor Name, PO_PRICE_USD, REC_QTY, Std Comm, ignore/consider]
var DB_VALUES = ${dbValuesJsonStr};
var MONTH_VALUES = ${monthValuesJsonStr};
var STDCOMM_VALUES = ${stdCommValuesJsonStr};
var VENDOR_VALUES = ${vendorValuesJsonStr};
var ITEMCODE_VALUES = ${itemCodeValuesJsonStr};
var STATUS_VALUES = ${statusValuesJsonStr};
var HEADER = ${headerJsonStr};
var DB_COLORS = ['#005daa', '#c0392b', '#27ae60', '#e67e22', '#8e44ad', '#16a085', '#d35400', '#7f8c8d'];
// Fixed DB -> color mapping (by full DB_VALUES list, not by whatever subset happens
// to be present after filtering) so the same DB always gets the same color across
// the bar chart, the pie chart, and between filter changes.
var DB_COLOR_MAP = {};
DB_VALUES.slice().sort().forEach(function (db, i) { DB_COLOR_MAP[db] = DB_COLORS[i % DB_COLORS.length]; });

// Item-level (ITEM_CODE x DB x Vendor) breakdown, computed from whatever rows
// are passed in (i.e. already filtered by Month/DB) - Avg is quantity-weighted
// (SUM(TOT_PO_PRICE_USD)/SUM(REC_QTY)), matching Excel's PivotTable exactly.
// Max_DB/Min_DB are new derived columns (not present in the source Excel):
// which DB the item-wide Max Vendor/Max Price (resp. Min Vendor/Min Price)
// transaction actually came from.
function computeItemPivot(rows) {
  var byItem = {};
  rows.forEach(function (r) {
    var item = r[3], desc = r[4], db = r[0], vendor = r[5], price = r[6], tot = r[2], qty = r[7];
    if (!byItem[item]) byItem[item] = { desc: desc, maxPrice: -Infinity, maxVendor: null, maxDb: null, minPrice: Infinity, minVendor: null, minDb: null, groups: {} };
    var it = byItem[item];
    if (typeof price === 'number') {
      if (price > it.maxPrice) { it.maxPrice = price; it.maxVendor = vendor; it.maxDb = db; }
      if (price < it.minPrice) { it.minPrice = price; it.minVendor = vendor; it.minDb = db; }
    }
    var key = db + '|' + vendor;
    if (!it.groups[key]) it.groups[key] = { db: db, vendor: vendor, sumQty: 0, sumTot: 0 };
    var g = it.groups[key];
    g.sumQty += (typeof qty === 'number' ? qty : 0);
    g.sumTot += (typeof tot === 'number' ? tot : 0);
  });
  var out = [];
  Object.keys(byItem).forEach(function (item) {
    var it = byItem[item];
    Object.keys(it.groups).forEach(function (key) {
      var g = it.groups[key];
      out.push([item, it.desc, it.maxDb, it.maxVendor, it.maxPrice, it.minDb, it.minVendor, it.minPrice, null, g.db, g.vendor, g.sumQty ? g.sumTot / g.sumQty : 0, g.sumTot]);
    });
  });
  return out;
}

function createMultiSelect(containerId, values, onChange, defaultSelected, labelMap) {
  var container = document.getElementById(containerId);
  var selected = new Set();
  var toggle = document.createElement('button');
  toggle.type = 'button';
  toggle.className = 'ms-toggle';
  container.appendChild(toggle);

  var panel = document.createElement('div');
  panel.className = 'ms-panel';
  panel.style.display = 'none';
  panel.addEventListener('click', function (e) { e.stopPropagation(); });
  container.appendChild(panel);

  var allRow = document.createElement('label');
  allRow.className = 'ms-all-row';
  var allCb = document.createElement('input');
  allCb.type = 'checkbox';
  allRow.appendChild(allCb);
  allRow.appendChild(document.createTextNode(' (Select All)'));
  panel.appendChild(allRow);

  // Normally every checkbox starts checked (no filtering). A caller can instead
  // pass defaultSelected (e.g. Ignore/Consider defaulting to just "Consider",
  // matching the report's historical Consider-only scope) to start with only
  // that subset checked. labelMap lets a caller show a friendlier description
  // while the checkbox's own value - what filtering matches against - stays
  // the raw underlying value.
  var checkboxes = [];
  values.forEach(function (v) {
    var row = document.createElement('label');
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = v;
    cb.checked = !defaultSelected || defaultSelected.indexOf(v) !== -1;
    row.appendChild(cb);
    row.appendChild(document.createTextNode(' ' + (labelMap && labelMap[v] ? labelMap[v] : v)));
    panel.appendChild(row);
    checkboxes.push(cb);
  });
  var initCheckedOnes = checkboxes.filter(function (c) { return c.checked; });
  if (initCheckedOnes.length === checkboxes.length || initCheckedOnes.length === 0) {
    selected = new Set();
    allCb.checked = true;
  } else {
    selected = new Set(initCheckedOnes.map(function (c) { return c.value; }));
    allCb.checked = false;
  }

  function displayLabel(v) { return (labelMap && labelMap[v]) ? labelMap[v] : v; }

  function updateToggleText() {
    if (selected.size === 0 || selected.size === values.length) {
      toggle.textContent = 'All ▾';
    } else if (selected.size === 1) {
      toggle.textContent = displayLabel(Array.from(selected)[0]) + ' ▾';
    } else {
      toggle.textContent = selected.size + ' selected ▾';
    }
  }

  allCb.addEventListener('change', function () {
    checkboxes.forEach(function (cb) { cb.checked = allCb.checked; });
    selected = new Set();
    updateToggleText();
    onChange();
  });

  checkboxes.forEach(function (cb) {
    cb.addEventListener('change', function () {
      var checkedOnes = checkboxes.filter(function (c) { return c.checked; });
      if (checkedOnes.length === checkboxes.length || checkedOnes.length === 0) {
        selected = new Set();
        allCb.checked = true;
      } else {
        selected = new Set(checkedOnes.map(function (c) { return c.value; }));
        allCb.checked = false;
      }
      updateToggleText();
      onChange();
    });
  });

  toggle.addEventListener('click', function (e) {
    e.stopPropagation();
    document.querySelectorAll('.ms-panel').forEach(function (p) { if (p !== panel) p.style.display = 'none'; });
    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
  });

  updateToggleText();

  return {
    getSelected: function () { return Array.from(selected); },
    reset: function () {
      selected = new Set();
      allCb.checked = true;
      checkboxes.forEach(function (cb) { cb.checked = true; });
      updateToggleText();
    }
  };
}

document.addEventListener('click', function () {
  document.querySelectorAll('.ms-panel').forEach(function (p) { p.style.display = 'none'; });
});

function matchesFilter(value, selectedArr) {
  return selectedArr.length === 0 || selectedArr.indexOf(value) !== -1;
}

function groupSum(rows, keyIdx, valIdx) {
  var map = {};
  rows.forEach(function (r) {
    var k = r[keyIdx];
    if (!map[k]) map[k] = 0;
    map[k] += (typeof r[valIdx] === 'number' ? r[valIdx] : 0);
  });
  return map;
}
function groupSumTwoKeys(rows, outerKeyIdx, innerKeyIdx, valIdx) {
  var map = {};
  rows.forEach(function (r) {
    var k1 = r[outerKeyIdx], k2 = r[innerKeyIdx];
    if (!map[k1]) map[k1] = {};
    if (!map[k1][k2]) map[k1][k2] = 0;
    map[k1][k2] += (typeof r[valIdx] === 'number' ? r[valIdx] : 0);
  });
  return map;
}

var msMonth = createMultiSelect('ms-month', MONTH_VALUES, function () { renderCharts(); });
var msDb = createMultiSelect('ms-db', DB_VALUES, function () { renderCharts(); });
var msStdComm = createMultiSelect('ms-stdcomm', STDCOMM_VALUES, function () { renderCharts(); });
var msVendor = createMultiSelect('ms-vendor', VENDOR_VALUES, function () { renderCharts(); });
var msItemCode = createMultiSelect('ms-itemcode', ITEMCODE_VALUES, function () { renderCharts(); });
// Mirrors Excel's PivotTable "ignore / consider" page filter - defaults to
// Consider only (this report's historical scope), but Ignore and (All) are
// selectable here just like in the Excel PivotTable.
// Ignore/Consider are internal jargon - show what they actually mean.
var msStatus = createMultiSelect('ms-status', STATUS_VALUES, function () { renderCharts(); }, ['Consider'], {
  Ignore: 'Ignore (Intercompany transaction)',
  Consider: 'Consider (Real transaction)'
});

// Per-item collapse state for the DB/Vendor/Avg/Tot breakdown rows: an item with
// multiple DB/Vendor sub-rows can be collapsed down to just its top (highest Avg
// PO Price(USD)) sub-row via the +/- toggle column rendered right before DB.
var collapsedItems = new Set();
var itemGroupCount = {};
var lastFilteredRows = [];

function redrawTable() {
  var allTableRows = computeItemPivot(lastFilteredRows);
  itemGroupCount = {};
  var maxAvgByItem = {};
  allTableRows.forEach(function (r) {
    var item = r[0], avg = r[11];
    itemGroupCount[item] = (itemGroupCount[item] || 0) + 1;
    if (!(item in maxAvgByItem) || avg > maxAvgByItem[item]) maxAvgByItem[item] = avg;
  });
  var tableRows = allTableRows.filter(function (r) {
    return !collapsedItems.has(r[0]) || r[11] === maxAvgByItem[r[0]];
  });
  document.getElementById('tableSummary').textContent = '${pageName} (Std Comm: ${stdCommLabel}) - ' + tableRows.length.toLocaleString() + ' rows';
  dataTable.clear();
  dataTable.rows.add(tableRows);
  dataTable.draw();

  var masterBtn = document.getElementById('collapseAllBtn');
  if (masterBtn) {
    var anyExpanded = Object.keys(itemGroupCount).some(function (item) {
      return itemGroupCount[item] > 1 && !collapsedItems.has(item);
    });
    masterBtn.textContent = anyExpanded ? '−' : '+';
    masterBtn.title = anyExpanded ? 'Collapse all items' : 'Expand all items';
  }
}

var charts = {};
function renderCharts() {
  var totalLineColor = document.body.getAttribute('data-theme') === 'dark' ? '#f0f0f0' : '#222';
  var selMonth = msMonth.getSelected();
  var selDb = msDb.getSelected();
  var selStdComm = msStdComm.getSelected();
  var selVendor = msVendor.getSelected();
  var selItemCode = msItemCode.getSelected();
  var selStatus = msStatus.getSelected();
  var rows = CHART_ROWS.filter(function (r) {
    return matchesFilter(r[1], selMonth) && matchesFilter(r[0], selDb) && matchesFilter(r[8], selStdComm) && matchesFilter(r[5], selVendor) && matchesFilter(String(r[3]), selItemCode) && matchesFilter(r[9], selStatus);
  });

  // Chart 1: By Month (GLID) - Total line + shadow bars per DB
  var monthSum = groupSum(rows, 1, 2);
  var monthLabels = Object.keys(monthSum).sort();
  var dbEntitiesPresent = Array.from(new Set(rows.map(function (r) { return r[0]; }))).filter(Boolean).sort();
  var monthByDbSum = groupSumTwoKeys(rows, 0, 1, 2); // { db: { month: sum } }

  var datasets = dbEntitiesPresent.map(function (db, i) {
    var byMonth = monthByDbSum[db] || {};
    var c = DB_COLOR_MAP[db] || DB_COLORS[i % DB_COLORS.length];
    return {
      type: 'bar', label: db,
      data: monthLabels.map(function (m) { return byMonth[m] || 0; }),
      backgroundColor: c + '55', borderColor: c + '55', stack: 'db', order: 2
    };
  });
  datasets.push({
    type: 'line', label: 'Total',
    data: monthLabels.map(function (m) { return monthSum[m]; }),
    borderColor: totalLineColor, backgroundColor: totalLineColor, borderWidth: 2.5,
    pointRadius: 4, pointBackgroundColor: totalLineColor, tension: 0.2, order: 1, fill: false
  });

  if (charts.month) charts.month.destroy();
  charts.month = new Chart(document.getElementById('chartGlidMonth'), {
    type: 'bar',
    data: { labels: monthLabels, datasets: datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        title: { display: true, text: 'By Month x ${pageName} - Total line + shadow bars per DB' },
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: function (ctx) {
              var v = ctx.parsed.y;
              var totalHere = monthSum[monthLabels[ctx.dataIndex]] || 0;
              var pct = totalHere ? (v / totalHere * 100).toFixed(1) : '0.0';
              return ctx.dataset.label + ': ' + v.toLocaleString(undefined, { maximumFractionDigits: 0 }) + ' (' + pct + '%)';
            }
          }
        }
      },
      scales: { x: { stacked: true }, y: { stacked: true } }
    }
  });

  // Chart 2: Pie - By DB of Total
  var dbSum = groupSum(rows, 0, 2);
  var dbLabels = Object.keys(dbSum).sort(function (a, b) { return dbSum[b] - dbSum[a]; });
  var dbTotalAll = dbLabels.reduce(function (s, l) { return s + dbSum[l]; }, 0);
  if (charts.dbPie) charts.dbPie.destroy();
  charts.dbPie = new Chart(document.getElementById('chartGlidDbPie'), {
    type: 'pie',
    data: {
      labels: dbLabels,
      datasets: [{
        data: dbLabels.map(function (l) { return dbSum[l]; }),
        backgroundColor: dbLabels.map(function (l, i) { return DB_COLOR_MAP[l] || DB_COLORS[i % DB_COLORS.length]; })
      }]
    },
    options: {
      responsive: true,
      plugins: {
        title: { display: true, text: 'By DB (${pageName}) - % of Total' },
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: function (ctx) {
              var v = ctx.parsed;
              var pct = dbTotalAll ? (v / dbTotalAll * 100).toFixed(1) : '0.0';
              return ctx.label + ': ' + v.toLocaleString(undefined, { maximumFractionDigits: 0 }) + ' (' + pct + '%)';
            }
          }
        }
      }
    }
  });

  // Item-level (ITEM_CODE x DB x Vendor) breakdown table - recomputed from the
  // same filtered rows as the charts, so Month/DB filters apply here too.
  lastFilteredRows = rows;
  redrawTable();
}

// Theme is controlled by the outer report shell (single global toggle), which
// calls this function directly on the iframe's window after (re)loading it.
function applyTheme(theme) {
  document.body.setAttribute('data-theme', theme);
  Chart.defaults.color = theme === 'dark' ? '#c8ccd4' : '#333333';
  Chart.defaults.borderColor = theme === 'dark' ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.1)';
}

var headRow = document.getElementById('dataTableHeadRow');
HEADER.forEach(function (h) {
  var th = document.createElement('th');
  th.appendChild(document.createTextNode(h));
  if (h === 'DB') {
    var masterBtn = document.createElement('span');
    masterBtn.id = 'collapseAllBtn';
    masterBtn.className = 'item-toggle';
    masterBtn.style.marginLeft = '6px';
    masterBtn.textContent = '−';
    masterBtn.title = 'Collapse all items';
    th.appendChild(masterBtn);
  }
  headRow.appendChild(th);
});
var dataTable = $('#dataTable').DataTable({
  data: [], deferRender: true, pageLength: 50,
  // scrollX off and autoWidth off: this table's 13 columns fit comfortably
  // without horizontal scroll, and scrollX's header/body column-width sync
  // (two separate DOM tables kept aligned via JS-computed widths, same as
  // the bug found and fixed on the Exchange Rate table) intermittently
  // miscalculated the ITEM_CODE column's width, spilling its nowrap text
  // into the next column. Relying on the browser's own table layout instead
  // is deterministic and immune to that failure mode.
  lengthMenu: [25, 50, 100, 500, 1000], scrollX: false, autoWidth: false,
  // Group by ITEM_CODE (col 0), then within each item's DB/Vendor rows,
  // order by Avg PO Price(USD) (col 11) highest to lowest.
  order: [[0, 'asc'], [11, 'desc']],
  // Columns 5, 8, 12, 13 (Max/Min/Avg/Tot price) shown to 4 decimal places;
  // sorting/filtering still use the raw numeric value, only the on-screen
  // text is formatted.
  columnDefs: [{
    // Own +/- column (col 9, right before DB) - not part of the source data
    // (always null), so it's rendered purely from the row's ITEM_CODE (row[0]).
    targets: [8],
    orderable: false,
    searchable: false,
    render: function (data, type, row) {
      if (type !== 'display') return '';
      var item = row[0];
      var count = itemGroupCount[item] || 1;
      if (count <= 1) return '';
      var collapsed = collapsedItems.has(item);
      var icon = collapsed ? '+' : '−';
      return '<span class="item-toggle" title="' + (collapsed ? 'Expand' : 'Collapse') + ' ' + count + ' DB/Vendor rows">' + icon + '</span>';
    }
  }, {
    targets: [4, 7, 11, 12],
    render: function (data, type) {
      if (type === 'display' && typeof data === 'number') return data.toFixed(4);
      return data;
    }
  }]
});

$('#dataTable tbody').on('click', '.item-toggle', function (e) {
  e.stopPropagation();
  var tr = $(this).closest('tr');
  var item = dataTable.row(tr).data()[0];
  if (collapsedItems.has(item)) collapsedItems.delete(item); else collapsedItems.add(item);
  redrawTable();
});

// Master +/- in the DB header: collapse every multi-row item at once ('close'),
// or clear all collapses at once ('release'/expand all). The icon reflects the
// action the next click performs, same convention as the per-row toggle.
document.getElementById('collapseAllBtn').addEventListener('click', function (e) {
  e.stopPropagation();
  var expandAll = this.textContent === '+';
  if (expandAll) {
    collapsedItems.clear();
  } else {
    Object.keys(itemGroupCount).forEach(function (item) {
      if (itemGroupCount[item] > 1) collapsedItems.add(item);
    });
  }
  redrawTable();
});

renderCharts();

document.getElementById('resetBtn').onclick = function () {
  msMonth.reset(); msDb.reset(); msStdComm.reset(); msVendor.reset(); msItemCode.reset(); msStatus.reset();
  renderCharts();
};

document.getElementById('exportBtn').onclick = function () {
  var rowsData = dataTable.rows({ search: 'applied' }).data();
  var lines = [HEADER.join(',')];
  for (var i = 0; i < rowsData.length; i++) {
    var r = rowsData[i];
    lines.push(r.map(function (v) {
      if (v === null || v === undefined) return '';
      var s = String(v);
      if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    }).join(','));
  }
  var csv = '\\ufeff' + lines.join('\\r\\n');
  var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = 'export.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
};
</script>
</body>
</html>`;
}

for (const f of FILTER_SHEETS) {
  const set = new Set(f.stdComm);
  // No longer hard-filtered to 'Consider' here - the Ignore/Consider split is now
  // a client-side filter (mirroring Excel's own PivotTable "ignore / consider"
  // page filter), defaulting to Consider only so today's scope is unchanged.
  const flatRows = dataSheet.rows.filter(r => set.has(r[stdCommIdx]));
  // rows: [DB, month, totUsd, ITEM_CODE, Std Part Desc, Std Vendor Name, PO_PRICE_USD, REC_QTY, Std Comm, ignore/consider]
  // (0,1,2 feed the charts; 3-8 feed the dynamic item-level table computed client-side
  // from these live rows, so all filters apply here too instead of relying on Excel's
  // static, already-all-months-aggregated pivot cache; 9 is filter-only, like 0/1/5/8)
  const itemRows = flatRows.map(r => [
    r[COL['DB']], r[COL['yyyy-mm']], typeof r[COL['TOT_PO_PRICE_USD']] === 'number' ? r[COL['TOT_PO_PRICE_USD']] : 0,
    r[COL['ITEM_CODE']], r[COL['Std Part Desc']], r[COL['Std Vendor Name']],
    typeof r[COL['PO_PRICE_USD']] === 'number' ? r[COL['PO_PRICE_USD']] : null,
    typeof r[COL['REC_QTY']] === 'number' ? r[COL['REC_QTY']] : 0,
    r[COL['Std Comm']], r[statusIdx]
  ]);
  const dbValues = Array.from(new Set(itemRows.map(r => r[0]))).filter(Boolean).sort();
  const monthValues = Array.from(new Set(itemRows.map(r => r[1]))).filter(Boolean).sort();
  const stdCommValues = Array.from(new Set(itemRows.map(r => r[8]))).filter(Boolean).sort();
  const vendorValues = Array.from(new Set(itemRows.map(r => r[5]))).filter(Boolean).sort();
  // ITEM_CODE is a mix of strings and raw numbers across sheets (e.g. Fittings has
  // bare numeric codes) - normalize to strings so it matches the multi-select's
  // checkbox values (always strings) on both sides.
  const itemCodeValues = Array.from(new Set(itemRows.map(r => String(r[3])))).filter(Boolean).sort();
  const statusValues = Array.from(new Set(itemRows.map(r => r[9]))).filter(Boolean).sort();
  // Column dividers between the 4 groups: ITEM_CODE/Std Part Desc | Max Vendor/Max
  // Price | Min Vendor/Min Price | DB/Vendor/Avg/Tot - confirmed identical layout
  // across all 4 pages, so applied to all of them now (no longer GLID-only).
  const addColDivider = true;
  filteredSheetsHtml[f.name] = buildPivotPageHtml(f.name, f.stdComm.join(', '), itemRows, dbValues, monthValues, stdCommValues, vendorValues, itemCodeValues, statusValues, PIVOT_HEADER, addColDivider);
  console.log(f.name + ':', itemRows.length, 'flat rows (dynamic item table + charts)');
}

// =========================================================================
// Dashboard: KPI cards + Chart.js charts + live filter dropdowns
// =========================================================================
const dbIdx = COL['DB'];
const poNoIdx = COL['PO_NO'];
const vendorIdIdx = COL['VENDOR_ID'];
const qtyIdx = COL['REC_QTY'];
const totUsdIdx = COL['TOT_PO_PRICE_USD'];
const monthIdx = COL['yyyy-mm'];

const dbValues = Array.from(new Set(dataSheet.rows.map(r => r[dbIdx]))).filter(Boolean).sort();
const monthValues = Array.from(new Set(dataSheet.rows.map(r => r[monthIdx]))).filter(Boolean).sort();
const reportTitle = monthValues.length
  ? 'Global Purchase Report - ' + monthValues[0] + (monthValues.length > 1 ? ' to ' + monthValues[monthValues.length - 1] : '')
  : 'Global Purchase Report';
const stdCommValues = Array.from(new Set(dataSheet.rows.map(r => r[stdCommIdx]))).filter(Boolean).sort();
const statusValues = Array.from(new Set(dataSheet.rows.map(r => r[statusIdx]))).filter(Boolean).sort();

const stdVendorNameIdx = COL['Std Vendor Name'];
const vendorValues = Array.from(new Set(dataSheet.rows.map(r => r[stdVendorNameIdx]))).filter(Boolean).sort();

// Minimal per-row projection to keep the Dashboard's own embedded dataset small
// (Dashboard only needs these fields for aggregation, not all 36 columns)
const dashRows = dataSheet.rows.map(r => [
  r[dbIdx], r[monthIdx], r[stdCommIdx], r[statusIdx], r[totUsdIdx], r[qtyIdx], r[poNoIdx], r[vendorIdIdx], r[stdVendorNameIdx]
]);
// dashRows column indices: 0=DB,1=month,2=stdComm,3=status,4=totUsd,5=qty,6=poNo,7=vendorId,8=stdVendorName

const dashRowsJsonStr = neutralize(JSON.stringify(dashRows));
const dbValuesJsonStr = neutralize(JSON.stringify(dbValues));
const monthValuesJsonStr = neutralize(JSON.stringify(monthValues));
const stdCommValuesJsonStr = neutralize(JSON.stringify(stdCommValues));
const statusValuesJsonStr = neutralize(JSON.stringify(statusValues));
const vendorValuesJsonStr = neutralize(JSON.stringify(vendorValues));

const dashboardHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Dashboard</title>
<style>${dataTablesCss}
  :root {
    --bg: #f4f6f9; --card-bg: #ffffff; --text: #222222; --text-muted: #666666;
    --accent: #005daa; --border: #cccccc; --border-light: #eeeeee; --shadow: rgba(0,0,0,0.1);
    --input-bg: #ffffff; --hover-bg: #f5f5f5;
  }
  body[data-theme="dark"] {
    --bg: #1a1d23; --card-bg: #262a33; --text: #e4e6eb; --text-muted: #a0a4ab;
    --accent: #6fa8ff; --border: #444a57; --border-light: #333844; --shadow: rgba(0,0,0,0.4);
    --input-bg: #1f232b; --hover-bg: #2f3440;
  }
  body { font-family: Calibri, Arial, sans-serif; margin: 12px; background: var(--bg); color: var(--text); }
  h2 { color: var(--accent); margin-top: 0; }
  h3 { color: var(--accent); margin: 0 0 8px 0; font-size: 14px; }
  .page-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
  #themeToggle { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--card-bg); color: var(--text); font-size: 12px; }
  .filters { display: flex; gap: 16px; flex-wrap: wrap; background: var(--card-bg); padding: 12px; border-radius: 6px; margin-bottom: 16px; box-shadow: 0 1px 3px var(--shadow); }
  .filters label.fl { display: block; font-size: 11px; color: var(--text-muted); margin-bottom: 4px; }
  .kpis { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
  .kpi { background: var(--card-bg); padding: 16px 24px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 180px; }
  .kpi .label { font-size: 12px; color: var(--text-muted); }
  .kpi .value { font-size: 24px; font-weight: bold; color: var(--accent); }
  .charts { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 16px; }
  .chart-box { background: var(--card-bg); padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 420px; }
  .chart-box.wide { flex-basis: 100%; }
  .table-box { background: var(--card-bg); padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); }
  canvas { max-height: 320px; }
  #resetBtn { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); }

  /* multi-select widget */
  .ms { position: relative; width: 180px; }
  .ms-toggle { width: 100%; text-align: left; padding: 5px 8px; background: var(--input-bg); color: var(--text); border: 1px solid var(--border); border-radius: 4px; cursor: pointer; font-size: 12px; }
  .ms-panel { position: absolute; z-index: 20; top: 100%; left: 0; background: var(--card-bg); border: 1px solid var(--border); border-radius: 4px; box-shadow: 0 2px 8px var(--shadow); max-height: 260px; overflow-y: auto; width: 240px; padding: 6px; }
  .ms-panel label { display: block; font-size: 12px; padding: 2px 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text); }
  .ms-panel label.ms-all-row { border-bottom: 1px solid var(--border-light); margin-bottom: 4px; padding-bottom: 6px; font-weight: bold; }
  table.dataTable { font-size: 12px; color: var(--text); }

  /* dark-mode overrides for DataTables' own chrome */
  body[data-theme="dark"] table.dataTable tbody tr { background-color: var(--card-bg); color: var(--text); }
  body[data-theme="dark"] table.dataTable tbody tr.odd,
  body[data-theme="dark"] table.dataTable tbody tr.even { background-color: var(--card-bg); }
  body[data-theme="dark"] table.dataTable.stripe tbody tr.odd,
  body[data-theme="dark"] table.dataTable.display tbody tr.odd { background-color: var(--hover-bg); }
  body[data-theme="dark"] table.dataTable tbody tr:hover { background-color: var(--hover-bg) !important; }
  body[data-theme="dark"] table.dataTable thead th,
  body[data-theme="dark"] table.dataTable thead td { color: var(--text); border-bottom-color: var(--border); }
  body[data-theme="dark"] table.dataTable tbody td { border-top-color: var(--border-light); }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_length,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_filter,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_info,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_processing,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate { color: var(--text); }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_length select,
  body[data-theme="dark"] .dataTables_wrapper .dataTables_filter input { background-color: var(--input-bg); color: var(--text); border-color: var(--border); }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button { color: var(--text) !important; }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button.current { background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important; }
  body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button:hover { background: var(--hover-bg) !important; color: var(--text) !important; }
</style>
</head>
<body>
<div class="page-header">
  <h2>PO Data Dashboard</h2>
</div>
<div class="filters">
  <div><label class="fl">DB (Entity)</label><div class="ms" id="ms-db"></div></div>
  <div><label class="fl">Month</label><div class="ms" id="ms-month"></div></div>
  <div><label class="fl">Std Comm</label><div class="ms" id="ms-stdcomm"></div></div>
  <div><label class="fl">Status</label><div class="ms" id="ms-status"></div></div>
  <div><label class="fl">Std Vendor Name</label><div class="ms" id="ms-vendor"></div></div>
  <div style="align-self:flex-end"><button id="resetBtn">Reset filters</button></div>
</div>
<div class="kpis">
  <div class="kpi"><div class="label">Total Purchase (USD)</div><div class="value" id="kpi-total">-</div></div>
  <div class="kpi"><div class="label">Total Orders (Count)</div><div class="value" id="kpi-orders">-</div></div>
  <div class="kpi"><div class="label">Total Quantity</div><div class="value" id="kpi-qty">-</div></div>
  <div class="kpi"><div class="label">Vendor Rows (Count)</div><div class="value" id="kpi-vendors">-</div></div>
</div>
<div class="charts">
  <div class="chart-box" style="flex:2; min-width:560px;"><canvas id="chartDB"></canvas></div>
  <div class="chart-box" style="flex:1; min-width:340px;"><canvas id="chartDBPie"></canvas></div>
</div>
<div class="charts">
  <div class="chart-box"><canvas id="chartTimeline"></canvas></div>
  <div class="chart-box"><canvas id="chartStdComm"></canvas></div>
</div>
<div class="table-box">
  <h3>By Vendor (Std Vendor Name)</h3>
  <table id="vendorTable" class="display">
    <thead><tr><th>Std Vendor Name</th><th>Sum of TOT_PO_PRICE_USD</th><th>Weight %</th><th>Sum of REC_QTY</th><th>Weight %</th><th>Count of PO_NO</th><th>Weight %</th></tr></thead>
    <tbody></tbody>
  </table>
</div>
<script>${jqueryJs}</script>
<script>${dataTablesJs}</script>
<script>${chartJs}</script>
<script>
// ROWS columns: 0=DB,1=month,2=stdComm,3=status,4=totUsd,5=qty,6=poNo,7=vendorId,8=stdVendorName
var ROWS = ${dashRowsJsonStr};
var DB_VALUES = ${dbValuesJsonStr};
var MONTH_VALUES = ${monthValuesJsonStr};
var STDCOMM_VALUES = ${stdCommValuesJsonStr};
var STATUS_VALUES = ${statusValuesJsonStr};
var VENDOR_VALUES = ${vendorValuesJsonStr};

// ---- Reusable multi-select widget: All / one / multi ----
// getSelected() returns [] to mean "All" (no filter applied).
function createMultiSelect(containerId, values, onChange, labelMap) {
  var container = document.getElementById(containerId);
  var selected = new Set(); // empty = All

  var toggle = document.createElement('button');
  toggle.type = 'button';
  toggle.className = 'ms-toggle';
  container.appendChild(toggle);

  var panel = document.createElement('div');
  panel.className = 'ms-panel';
  panel.style.display = 'none';
  panel.addEventListener('click', function (e) { e.stopPropagation(); });
  container.appendChild(panel);

  var allRow = document.createElement('label');
  allRow.className = 'ms-all-row';
  var allCb = document.createElement('input');
  allCb.type = 'checkbox';
  allCb.checked = true;
  allRow.appendChild(allCb);
  allRow.appendChild(document.createTextNode(' (Select All)'));
  panel.appendChild(allRow);

  // labelMap lets a caller (e.g. Status: Ignore/Consider) show a friendlier
  // description while the checkbox's own value - what filtering matches
  // against - stays the raw underlying value.
  var checkboxes = [];
  values.forEach(function (v) {
    var row = document.createElement('label');
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = v;
    cb.checked = true;
    row.appendChild(cb);
    row.appendChild(document.createTextNode(' ' + (labelMap && labelMap[v] ? labelMap[v] : v)));
    panel.appendChild(row);
    checkboxes.push(cb);
  });

  function displayLabel(v) { return (labelMap && labelMap[v]) ? labelMap[v] : v; }

  function updateToggleText() {
    if (selected.size === 0 || selected.size === values.length) {
      toggle.textContent = 'All ▾';
    } else if (selected.size === 1) {
      toggle.textContent = displayLabel(Array.from(selected)[0]) + ' ▾';
    } else {
      toggle.textContent = selected.size + ' selected ▾';
    }
  }

  allCb.addEventListener('change', function () {
    checkboxes.forEach(function (cb) { cb.checked = allCb.checked; });
    selected = new Set();
    if (allCb.checked) {
      // All checked = no filter (empty selected set means "All")
      selected = new Set();
    } else {
      selected = new Set(); // none checked = also treat as "All" (avoid an impossible zero-result state)
    }
    updateToggleText();
    onChange();
  });

  checkboxes.forEach(function (cb) {
    cb.addEventListener('change', function () {
      var checkedOnes = checkboxes.filter(function (c) { return c.checked; });
      if (checkedOnes.length === checkboxes.length || checkedOnes.length === 0) {
        selected = new Set();
        allCb.checked = true;
      } else {
        selected = new Set(checkedOnes.map(function (c) { return c.value; }));
        allCb.checked = false;
      }
      updateToggleText();
      onChange();
    });
  });

  toggle.addEventListener('click', function (e) {
    e.stopPropagation();
    document.querySelectorAll('.ms-panel').forEach(function (p) { if (p !== panel) p.style.display = 'none'; });
    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
  });

  updateToggleText();

  return {
    getSelected: function () { return Array.from(selected); },
    reset: function () {
      selected = new Set();
      allCb.checked = true;
      checkboxes.forEach(function (cb) { cb.checked = true; });
      updateToggleText();
    }
  };
}

document.addEventListener('click', function () {
  document.querySelectorAll('.ms-panel').forEach(function (p) { p.style.display = 'none'; });
});

var msDb, msMonth, msStdComm, msStatus, msVendor;
function onFilterChange() { render(); }
msDb = createMultiSelect('ms-db', DB_VALUES, onFilterChange);
msMonth = createMultiSelect('ms-month', MONTH_VALUES, onFilterChange);
msStdComm = createMultiSelect('ms-stdcomm', STDCOMM_VALUES, onFilterChange);
// Ignore/Consider are internal jargon - show what they actually mean.
msStatus = createMultiSelect('ms-status', STATUS_VALUES, onFilterChange, {
  Ignore: 'Ignore (Intercompany transaction)',
  Consider: 'Consider (Real transaction)'
});
msVendor = createMultiSelect('ms-vendor', VENDOR_VALUES, onFilterChange);

var charts = {};
var vendorDT = null;

function matchesFilter(value, selectedArr) {
  return selectedArr.length === 0 || selectedArr.indexOf(value) !== -1;
}

function getFiltered() {
  var selDb = msDb.getSelected();
  var selMonth = msMonth.getSelected();
  var selStdComm = msStdComm.getSelected();
  var selStatus = msStatus.getSelected();
  var selVendor = msVendor.getSelected();
  return ROWS.filter(function (r) {
    if (!matchesFilter(r[0], selDb)) return false;
    if (!matchesFilter(r[1], selMonth)) return false;
    if (!matchesFilter(r[2], selStdComm)) return false;
    if (!matchesFilter(r[3], selStatus)) return false;
    if (!matchesFilter(r[8], selVendor)) return false;
    return true;
  });
}

function groupSum(rows, keyIdx, valIdx) {
  var map = {};
  rows.forEach(function (r) {
    var k = r[keyIdx];
    if (!map[k]) map[k] = 0;
    map[k] += (typeof r[valIdx] === 'number' ? r[valIdx] : 0);
  });
  return map;
}
function groupCount(rows, keyIdx) {
  var map = {};
  rows.forEach(function (r) {
    var k = r[keyIdx];
    map[k] = (map[k] || 0) + 1;
  });
  return map;
}
function groupSumTwoKeys(rows, outerKeyIdx, innerKeyIdx, valIdx) {
  var map = {};
  rows.forEach(function (r) {
    var k1 = r[outerKeyIdx], k2 = r[innerKeyIdx];
    if (!map[k1]) map[k1] = {};
    if (!map[k1][k2]) map[k1][k2] = 0;
    map[k1][k2] += (typeof r[valIdx] === 'number' ? r[valIdx] : 0);
  });
  return map;
}

var SC_GROUP_ORDER = ['SKU-', 'GLID', 'FITTING', 'PACKAGING', 'COATING', 'OTHER'];
var SC_FITTING_CODES = ['HDL-', 'KNOB', 'LHDL', 'SHDL'];
var SC_PACKAGING_CODES = ['BHRE', 'LABL', 'STKR', 'U&C-', 'WACD'];
var SC_COATING_CODES = ['EXM-', 'INM-'];
function stdCommGroup(sc) {
  if (sc === 'SKU-') return 'SKU-';
  if (sc === 'GLID') return 'GLID';
  if (SC_FITTING_CODES.indexOf(sc) !== -1) return 'FITTING';
  if (SC_PACKAGING_CODES.indexOf(sc) !== -1) return 'PACKAGING';
  if (SC_COATING_CODES.indexOf(sc) !== -1) return 'COATING';
  return 'OTHER';
}

function render() {
  var rows = getFiltered();
  var totalLineColor = document.body.getAttribute('data-theme') === 'dark' ? '#f0f0f0' : '#222';

  var totalUsd = 0, totalQty = 0;
  rows.forEach(function (r) { totalUsd += (r[4] || 0); totalQty += (r[5] || 0); });
  document.getElementById('kpi-total').textContent = totalUsd.toLocaleString(undefined, { maximumFractionDigits: 0 });
  document.getElementById('kpi-orders').textContent = rows.length.toLocaleString();
  document.getElementById('kpi-qty').textContent = totalQty.toLocaleString(undefined, { maximumFractionDigits: 0 });
  document.getElementById('kpi-vendors').textContent = rows.length.toLocaleString();

  // Chart 1: by DB (X-axis) - shadow bars stacked by Month, + a bold Total-per-DB line
  // (mirror image of the By Month chart: there X=month/stack=DB, here X=DB/stack=month)
  var dbSumUsd = groupSum(rows, 0, 4);
  var dbLabels = Object.keys(dbSumUsd).sort(function(a,b){ return dbSumUsd[b]-dbSumUsd[a]; });
  var monthsPresent = Array.from(new Set(rows.map(function (r) { return r[1]; }))).filter(Boolean).sort();
  var dbByMonthSum = groupSumTwoKeys(rows, 1, 0, 4); // { month: { db: sum } }
  var MONTH_COLORS = ['#005daa', '#c0392b', '#27ae60', '#e67e22', '#8e44ad', '#16a085', '#d35400', '#7f8c8d'];
  var DB_COLORS = ['#005daa', '#c0392b', '#27ae60', '#e67e22', '#8e44ad', '#16a085', '#d35400', '#7f8c8d'];
  // Fixed DB -> color mapping (by full DB_VALUES list, not by whatever subset happens
  // to be present after filtering, and not by a value-sorted order) so the same DB
  // always gets the same color across the DB pie, the timeline chart, and the Std
  // Comm chart below, and stays stable across filter changes.
  var DB_COLOR_MAP = {};
  DB_VALUES.slice().sort().forEach(function (db, i) { DB_COLOR_MAP[db] = DB_COLORS[i % DB_COLORS.length]; });

  var dbChartDatasets = monthsPresent.map(function (m, i) {
    var byDb = dbByMonthSum[m] || {};
    var c = MONTH_COLORS[i % MONTH_COLORS.length];
    return {
      type: 'bar',
      label: m,
      data: dbLabels.map(function (db) { return byDb[db] || 0; }),
      backgroundColor: c + '55',
      borderColor: c + '55',
      stack: 'month',
      order: 2
    };
  });
  dbChartDatasets.push({
    type: 'line',
    label: 'Total',
    data: dbLabels.map(function (db) { return dbSumUsd[db]; }),
    borderColor: totalLineColor,
    backgroundColor: totalLineColor,
    borderWidth: 2.5,
    pointRadius: 4,
    pointBackgroundColor: totalLineColor,
    tension: 0.2,
    order: 1,
    fill: false
  });

  if (charts.db) charts.db.destroy();
  charts.db = new Chart(document.getElementById('chartDB'), {
    type: 'bar',
    data: { labels: dbLabels, datasets: dbChartDatasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        title: { display: true, text: 'By Entity (DB) - Total line + shadow bars per Month' },
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: function (ctx) {
              var v = ctx.parsed.y;
              var totalHere = dbSumUsd[dbLabels[ctx.dataIndex]] || 0;
              var pct = totalHere ? (v / totalHere * 100).toFixed(1) : '0.0';
              return ctx.dataset.label + ': ' + v.toLocaleString(undefined, { maximumFractionDigits: 0 }) + ' (' + pct + '%)';
            }
          }
        }
      },
      scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }
    }
  });

  // Pie chart: By Entity (DB) - % share of Sum of TOT_PO_PRICE_USD
  var dbTotalAll = dbLabels.reduce(function (s, l) { return s + dbSumUsd[l]; }, 0);
  if (charts.dbPie) charts.dbPie.destroy();
  charts.dbPie = new Chart(document.getElementById('chartDBPie'), {
    type: 'pie',
    data: {
      labels: dbLabels,
      datasets: [{
        data: dbLabels.map(function (l) { return dbSumUsd[l]; }),
        backgroundColor: dbLabels.map(function (l, i) { return DB_COLOR_MAP[l] || DB_COLORS[i % DB_COLORS.length]; })
      }]
    },
    options: {
      responsive: true,
      plugins: {
        title: { display: true, text: 'By Entity (DB) - %' },
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: function (ctx) {
              var v = ctx.parsed;
              var pct = dbTotalAll ? (v / dbTotalAll * 100).toFixed(1) : '0.0';
              return ctx.label + ': ' + v.toLocaleString(undefined, { maximumFractionDigits: 0 }) + ' (' + pct + '%)';
            }
          }
        }
      }
    }
  });

  // Chart 2: timeline (by month) - per-DB entities as shadow (semi-transparent) stacked bars,
  // Total as a bold line on top. Tooltip shows Total + every currently-visible DB's number
  // together (whichever DB entities remain after the top DB filter - all of them if unfiltered,
  // just the selected one(s) if the DB filter narrowed it down).
  var monthSum = groupSum(rows, 1, 4);
  var monthLabels = Object.keys(monthSum).sort();
  var dbEntitiesPresent = Array.from(new Set(rows.map(function (r) { return r[0]; }))).filter(Boolean).sort();
  var monthByDbSum = groupSumTwoKeys(rows, 0, 1, 4); // { db: { month: sum } }

  var timelineDatasets = dbEntitiesPresent.map(function (db, i) {
    var byMonth = monthByDbSum[db] || {};
    var c = DB_COLOR_MAP[db] || DB_COLORS[i % DB_COLORS.length];
    return {
      type: 'bar',
      label: db,
      data: monthLabels.map(function (m) { return byMonth[m] || 0; }),
      backgroundColor: c + '55', // "shadow" bar: same hue, low opacity
      borderColor: c + '55',
      stack: 'db',
      order: 2
    };
  });
  timelineDatasets.push({
    type: 'line',
    label: 'Total',
    data: monthLabels.map(function (l) { return monthSum[l]; }),
    borderColor: totalLineColor,
    backgroundColor: totalLineColor,
    borderWidth: 2.5,
    pointRadius: 4,
    pointBackgroundColor: totalLineColor,
    tension: 0.2,
    order: 1,
    fill: false
  });

  if (charts.timeline) charts.timeline.destroy();
  charts.timeline = new Chart(document.getElementById('chartTimeline'), {
    type: 'bar',
    data: { labels: monthLabels, datasets: timelineDatasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        title: { display: true, text: 'By Month (Total line + shadow bars per Entity)' },
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: function (ctx) {
              var v = ctx.parsed.y;
              var totalHere = monthSum[monthLabels[ctx.dataIndex]] || 0;
              var pct = totalHere ? (v / totalHere * 100).toFixed(1) : '0.0';
              return ctx.dataset.label + ': ' + v.toLocaleString(undefined, { maximumFractionDigits: 0 }) + ' (' + pct + '%)';
            }
          }
        }
      },
      scales: { x: { stacked: true }, y: { stacked: true } }
    }
  });

  // Chart 3: by Std Comm, grouped into SKU-, GLID, FITTING, PACKAGING, COATING, OTHER
  // Same "shadow" treatment as the other charts: per-DB entities as shadow (semi-transparent)
  // stacked bars, Total as a bold line on top.
  var scGroupSum = {};
  var scByDbSum = {}; // { db: { group: sum } }
  rows.forEach(function (r) {
    var g = stdCommGroup(r[2]);
    scGroupSum[g] = (scGroupSum[g] || 0) + (typeof r[4] === 'number' ? r[4] : 0);
    var db = r[0];
    if (!scByDbSum[db]) scByDbSum[db] = {};
    scByDbSum[db][g] = (scByDbSum[db][g] || 0) + (typeof r[4] === 'number' ? r[4] : 0);
  });
  var scLabels = SC_GROUP_ORDER;

  var scChartDatasets = dbEntitiesPresent.map(function (db, i) {
    var byGroup = scByDbSum[db] || {};
    var c = DB_COLOR_MAP[db] || DB_COLORS[i % DB_COLORS.length];
    return {
      type: 'bar',
      label: db,
      data: scLabels.map(function (g) { return byGroup[g] || 0; }),
      backgroundColor: c + '55',
      borderColor: c + '55',
      stack: 'db',
      order: 2
    };
  });
  scChartDatasets.push({
    type: 'line',
    label: 'Total',
    data: scLabels.map(function (g) { return scGroupSum[g] || 0; }),
    borderColor: totalLineColor,
    backgroundColor: totalLineColor,
    borderWidth: 2.5,
    pointRadius: 4,
    pointBackgroundColor: totalLineColor,
    tension: 0.2,
    order: 1,
    fill: false
  });

  if (charts.stdcomm) charts.stdcomm.destroy();
  charts.stdcomm = new Chart(document.getElementById('chartStdComm'), {
    type: 'bar',
    data: { labels: scLabels, datasets: scChartDatasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        title: { display: true, text: 'By Std Comm (Group) - Total line + shadow bars per Entity' },
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: function (ctx) {
              var v = ctx.parsed.y;
              var totalHere = scGroupSum[scLabels[ctx.dataIndex]] || 0;
              var pct = totalHere ? (v / totalHere * 100).toFixed(1) : '0.0';
              return ctx.dataset.label + ': ' + v.toLocaleString(undefined, { maximumFractionDigits: 0 }) + ' (' + pct + '%)';
            }
          }
        }
      },
      scales: { x: { stacked: true }, y: { stacked: true } }
    }
  });

  // By Vendor table
  var vSumUsd = groupSum(rows, 8, 4);
  var vSumQty = groupSum(rows, 8, 5);
  var vCountPo = groupCount(rows, 8);
  var vTotalUsd = Object.keys(vSumUsd).reduce(function (s, k) { return s + vSumUsd[k]; }, 0);
  var vTotalQty = Object.keys(vSumQty).reduce(function (s, k) { return s + (vSumQty[k] || 0); }, 0);
  var vTotalCount = Object.keys(vCountPo).reduce(function (s, k) { return s + (vCountPo[k] || 0); }, 0);
  var vendorTableData = Object.keys(vSumUsd).map(function (v) {
    var sumUsd = vSumUsd[v], sumQty = vSumQty[v] || 0, countPo = vCountPo[v] || 0;
    return [
      v,
      sumUsd, vTotalUsd ? (sumUsd / vTotalUsd * 100) : 0,
      sumQty, vTotalQty ? (sumQty / vTotalQty * 100) : 0,
      countPo, vTotalCount ? (countPo / vTotalCount * 100) : 0
    ];
  });

  if (vendorDT) {
    vendorDT.clear();
    vendorDT.rows.add(vendorTableData);
    vendorDT.draw();
  } else {
    vendorDT = $('#vendorTable').DataTable({
      data: vendorTableData,
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100],
      autoWidth: false,
      order: [[1, 'desc']],
      columns: [
        { title: 'Std Vendor Name' },
        { title: 'Sum of TOT_PO_PRICE_USD', render: function (d) { return Number(d).toLocaleString(undefined, { maximumFractionDigits: 0 }); } },
        { title: 'Weight %', render: function (d) { return Number(d).toFixed(1) + '%'; } },
        { title: 'Sum of REC_QTY', render: function (d) { return Number(d).toLocaleString(undefined, { maximumFractionDigits: 0 }); } },
        { title: 'Weight %', render: function (d) { return Number(d).toFixed(1) + '%'; } },
        { title: 'Count of PO_NO' },
        { title: 'Weight %', render: function (d) { return Number(d).toFixed(1) + '%'; } }
      ]
    });
  }
}

// Theme is controlled by the outer report shell (single global toggle), which
// calls this function directly on the iframe's window after (re)loading it.
function applyTheme(theme) {
  document.body.setAttribute('data-theme', theme);
  Chart.defaults.color = theme === 'dark' ? '#c8ccd4' : '#333333';
  Chart.defaults.borderColor = theme === 'dark' ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.1)';
}

document.getElementById('resetBtn').onclick = function () {
  msDb.reset(); msMonth.reset(); msStdComm.reset(); msStatus.reset(); msVendor.reset();
  render();
};

render();
</script>
</body>
</html>`;

console.log('Dashboard built (' + (dashboardHtml.length / 1024 / 1024).toFixed(2) + ' MB)');

// =========================================================================
// Data sheet (unchanged from previous build, full detail table)
// =========================================================================
const dataSheetHtml = buildDataTablesHtml(H, dataSheet.rows, 'Data');
console.log('Data sheet built (' + (dataSheetHtml.length / 1024 / 1024).toFixed(2) + ' MB)');

// =========================================================================
// Exchange Rate: only the "Platium Buy Rate : Foreign Currency => USD" table
// is kept (the HKD-rate table next to it is dropped per user decision), and
// it's rendered the same way as the Data sheet - a DataTables page with
// sorting/search/CSV export - instead of Excel's raw native HTML export.
// =========================================================================
let exRateJsonRaw = fs.readFileSync(path.join(DIR, 'exrate_platinum_usd.json'), 'utf8');
if (exRateJsonRaw.charCodeAt(0) === 0xFEFF) exRateJsonRaw = exRateJsonRaw.slice(1);
const exRateData = JSON.parse(exRateJsonRaw);
// scrollX disabled: unlike the 36-column Data sheet, this table's 12 columns
// fit comfortably without horizontal scrolling, and scrollX's header/body
// column-width sync (two separate DOM tables kept aligned via JS-computed
// widths) misaligned header text against data columns for a table this narrow.
// All 11 rate columns (everything but yyyy-mm at index 0) shown to 4 decimal
// places, matching the same convention used for the price columns elsewhere
// in this report.
const exRateDecimalCols = exRateData.header.map((h, i) => i).filter(i => i > 0);
const exRateHtml = buildDataTablesHtml(exRateData.header, exRateData.rows, 'Exchange Rate', false, exRateDecimalCols);

// =========================================================================
// Assemble final combined file
// =========================================================================
const sheetsData = [
  { name: 'Dashboard', html: dashboardHtml },
  { name: 'GLID', html: filteredSheetsHtml['GLID'] },
  { name: 'Fittings', html: filteredSheetsHtml['Fittings'] },
  { name: 'Packaging', html: filteredSheetsHtml['Packaging'] },
  { name: 'Coating', html: filteredSheetsHtml['Coating'] },
  { name: 'Exchange Rate', html: exRateHtml },
  { name: 'Data', html: dataSheetHtml },
];

const sheetsJsonStr = neutralize(JSON.stringify(sheetsData));

const combinedHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${reportTitle}</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; margin: 0; padding: 0; }
  #tabbar { display: flex; background: #005daa; align-items: stretch; }
  #tabs { display: flex; padding: 0; margin: 0; list-style: none; flex-wrap: wrap; flex: 1; }
  #tabs button { background: #005daa; color: white; border: none; padding: 10px 16px; cursor: pointer; font-size: 14px; }
  #tabs button.active { background: #1a3d6d; font-weight: bold; }
  #tabs button:hover { background: #1a3d6d; }
  #themeToggle { background: #1a3d6d; color: white; border: none; padding: 10px 16px; cursor: pointer; font-size: 13px; white-space: nowrap; }
  #themeToggle:hover { background: #12294a; }
  iframe { width: 100%; height: calc(100vh - 44px); border: none; display: block; }
</style>
</head>
<body>
<div id="tabbar">
  <div id="tabs"></div>
  <button id="themeToggle">🌙 Dark Mode</button>
</div>
<iframe id="sheetFrame" title="Report sheet"></iframe>
<script>
var SHEETS = ${sheetsJsonStr};
var tabsEl = document.getElementById('tabs');
var frameEl = document.getElementById('sheetFrame');
var themeBtn = document.getElementById('themeToggle');

var currentTheme = 'light';
try { currentTheme = localStorage.getItem('reportTheme') || 'light'; } catch (e) {}

function applyThemeToFrame() {
  var win = frameEl.contentWindow;
  if (!win || typeof win.applyTheme !== 'function') return;
  win.applyTheme(currentTheme);
  if (typeof win.renderCharts === 'function') win.renderCharts();
  else if (typeof win.render === 'function') win.render();
}

function updateThemeBtnLabel() {
  themeBtn.textContent = currentTheme === 'dark' ? '☀️ Light Mode' : '🌙 Dark Mode';
}

themeBtn.onclick = function () {
  currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
  try { localStorage.setItem('reportTheme', currentTheme); } catch (e) {}
  updateThemeBtnLabel();
  applyThemeToFrame();
};

frameEl.addEventListener('load', applyThemeToFrame);

function showSheet(i) {
  frameEl.srcdoc = SHEETS[i].html;
  Array.prototype.forEach.call(tabsEl.children, function (btn, idx) { btn.className = idx === i ? 'active' : ''; });
}
SHEETS.forEach(function (s, i) {
  var btn = document.createElement('button');
  btn.textContent = s.name;
  btn.onclick = function () { showSheet(i); };
  tabsEl.appendChild(btn);
});
updateThemeBtnLabel();
showSheet(0);
</script>
</body>
</html>`;

const outPath = OUT_PATH;
fs.writeFileSync(outPath, combinedHtml, 'utf8');
const stat = fs.statSync(outPath);
console.log('Wrote combined report:', outPath);
console.log('Total size:', (stat.size / 1024 / 1024).toFixed(2), 'MB');
