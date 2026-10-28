import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import {
  Activity,
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  ChevronLeft,
  ChevronRight,
  CircleDollarSign,
  Database,
  Download,
  Globe2,
  LoaderCircle,
  Moon,
  PackageSearch,
  Pause,
  Play,
  Plus,
  RotateCcw,
  Search,
  Sun,
  X,
} from "lucide-react";
import type { EChartsOption } from "echarts";
import { trpc } from "@/providers/trpc";
import { ENTITY_HUBS } from "@contracts/procurement";
import type { Purchase, PurchaseFilters } from "@contracts/procurement";
import { Chart } from "@/components/Chart";
import { CreatePurchaseModal, PurchaseDrawer } from "@/components/PurchasePanels";
import { compact, csvCell, decimal, integer, monthLabel, usd } from "@/lib/format";

interface GlobeMarker {
  city: string;
  v: number;
  entity: string;
  level?: "warn";
}
interface GlobeHandle {
  focus: (city: string) => void;
  setMarkers: (markers: GlobeMarker[]) => void;
  retheme: () => void;
  destroy: () => void;
}
interface GlobeModule {
  globe: (host: HTMLDivElement, options: {
    routes: [];
    markers: GlobeMarker[];
    onPick: (marker: GlobeMarker) => void;
  }) => Promise<GlobeHandle>;
}

const EMPTY_FILTERS: PurchaseFilters = {};
const css = (name: string) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

export default function App() {
  const utils = trpc.useUtils();
  const [filters, setFilters] = useState<PurchaseFilters>(EMPTY_FILTERS);
  const [page, setPage] = useState(1);
  const [searchDraft, setSearchDraft] = useState("");
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState<"poRecDate" | "totalUsd" | "quantity" | "poNo">("poRecDate");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");
  const [selected, setSelected] = useState<Purchase | null>(null);
  const [createOpen, setCreateOpen] = useState(false);
  const [toast, setToast] = useState("");
  const [exporting, setExporting] = useState(false);
  const [touring, setTouring] = useState(false);
  const [globeReady, setGlobeReady] = useState(false);
  const [globeError, setGlobeError] = useState(false);
  const [theme, setTheme] = useState<"dark" | "light">(() => {
    try { return localStorage.getItem("gpr-theme") === "light" ? "light" : "dark"; } catch { return "dark"; }
  });
  const [clock, setClock] = useState(() => new Date());
  const earthHost = useRef<HTMLDivElement>(null);
  const earth = useRef<GlobeHandle | null>(null);

  const optionsQuery = trpc.procurement.options.useQuery();
  const overviewQuery = trpc.procurement.overview.useQuery(filters, {
    placeholderData: (previous) => previous,
  });
  const recordsQuery = trpc.procurement.records.useQuery({
    ...filters,
    search,
    page,
    pageSize: 12,
    sortBy,
    sortDir,
  }, { placeholderData: (previous) => previous });
  const options = optionsQuery.data;
  const data = overviewQuery.data;
  const records = recordsQuery.data;

  const showToast = useCallback((message: string) => setToast(message), []);
  const closeCreate = useCallback(() => setCreateOpen(false), []);
  const closeDrawer = useCallback(() => setSelected(null), []);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    try { localStorage.setItem("gpr-theme", theme); } catch { /* UI preference only */ }
    requestAnimationFrame(() => {
      dispatchEvent(new Event("kit:theme"));
      earth.current?.retheme();
    });
  }, [theme]);

  useEffect(() => {
    const timer = setInterval(() => setClock(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(""), 4200);
    return () => clearTimeout(timer);
  }, [toast]);

  const globeEnabled = Boolean(data);
  useEffect(() => {
    let cancelled = false;
    let handle: GlobeHandle | null = null;
    if (!globeEnabled || !earthHost.current) return;
    import("./lib/globe.js").then((module) => {
      const globeModule = module as unknown as GlobeModule;
      return globeModule.globe(earthHost.current!, {
        routes: [],
        markers: [],
        onPick: (marker) => {
          if (!marker.entity) return;
          setFilters((current) => ({ ...current, entity: marker.entity }));
          setPage(1);
        },
      });
    }).then((instance) => {
      if (cancelled) { instance.destroy(); return; }
      handle = instance;
      earth.current = instance;
      setGlobeReady(true);
    }).catch((error) => {
      if (!cancelled) {
        console.warn("Globe initialization failed", error);
        setGlobeError(true);
      }
    });
    return () => {
      cancelled = true;
      handle?.destroy();
      earth.current = null;
    };
  }, [globeEnabled]);

  const markers = useMemo<GlobeMarker[]>(() => {
    if (!data?.entities.length) return [];
    const max = Math.max(...data.entities.map((item) => item.value));
    return data.entities.map((item) => ({
      city: ENTITY_HUBS[item.name] ?? "上海",
      entity: item.name,
      v: Math.max(12, item.value / max * 100),
      level: filters.entity === item.name ? "warn" : undefined,
    }));
  }, [data?.entities, filters.entity]);

  useEffect(() => { earth.current?.setMarkers(markers); }, [markers, globeReady]);
  useEffect(() => {
    if (filters.entity) earth.current?.focus(ENTITY_HUBS[filters.entity] ?? "上海");
  }, [filters.entity, globeReady]);

  useEffect(() => {
    if (!touring || !data?.entities.length) return;
    const timer = setInterval(() => {
      setFilters((current) => {
        const index = data.entities.findIndex((item) => item.name === current.entity);
        const next = data.entities[(index + 1) % data.entities.length];
        return { ...current, entity: next.name };
      });
      setPage(1);
    }, 7000);
    return () => clearInterval(timer);
  }, [touring, data?.entities]);

  const updateFilter = (key: keyof PurchaseFilters, value: string) => {
    setTouring(false);
    setFilters((current) => ({ ...current, [key]: value || undefined }));
    setPage(1);
  };
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS);
    setSearch("");
    setSearchDraft("");
    setPage(1);
    setTouring(false);
  };
  const activeFilterCount = Object.values(filters).filter(Boolean).length + (search ? 1 : 0);
  const totalPages = Math.max(1, Math.ceil((records?.total ?? 0) / 12));

  const categoryOption = useMemo<EChartsOption>(() => {
    const rows = data?.categories ?? [];
    const top = rows.slice(0, 7);
    const other = rows.slice(7).reduce((sum, row) => sum + row.value, 0);
    const values = [...top.map((row) => ({ name: row.name, value: row.value })),
      ...(other > 0 ? [{ name: "其他", value: other }] : [])];
    return {
      animationDuration: 260,
      color: [css("--s1"), css("--s2"), css("--s3"), css("--s4"), css("--s5"), css("--s6"), css("--s3"), css("--faint")],
      tooltip: { trigger: "item", backgroundColor: css("--raised"), borderColor: css("--line"), textStyle: { color: css("--ink") },
        valueFormatter: (value: unknown) => usd(Number(value)) },
      legend: { bottom: 0, left: "center", itemWidth: 8, itemHeight: 8, itemGap: 12,
        textStyle: { color: css("--muted"), fontSize: 11, fontFamily: "Geist Mono Variable" } },
      series: [{ type: "pie", radius: ["48%", "72%"], center: ["50%", "43%"], minAngle: 2,
        label: { show: false }, itemStyle: { borderColor: css("--panel"), borderWidth: 2 }, data: values }],
    };
  }, [data?.categories, theme]);

  const trendOption = useMemo<EChartsOption>(() => ({
    animationDuration: 260,
    grid: { left: 54, right: 48, top: 28, bottom: 28 },
    tooltip: { trigger: "axis", backgroundColor: css("--raised"), borderColor: css("--line"), textStyle: { color: css("--ink") },
      valueFormatter: (value: unknown) => typeof value === "number" ? decimal(value) : String(value) },
    legend: { top: 0, right: 6, textStyle: { color: css("--muted"), fontSize: 11 }, itemWidth: 12, itemHeight: 6 },
    xAxis: { type: "category", data: (data?.monthly ?? []).map((row) => monthLabel(row.name)),
      axisLine: { lineStyle: { color: css("--line") } }, axisTick: { show: false },
      axisLabel: { color: css("--muted"), fontFamily: "Geist Mono Variable", fontSize: 11 } },
    yAxis: [
      { type: "value", axisLabel: { color: css("--muted"), formatter: (value: number) => `$${value / 1e6}M`, fontSize: 10 },
        splitLine: { lineStyle: { color: css("--line-soft") } } },
      { type: "value", axisLabel: { color: css("--muted"), formatter: (value: number) => compact(value), fontSize: 10 }, splitLine: { show: false } },
    ],
    series: [
      { name: "採購額 USD", type: "bar", data: (data?.monthly ?? []).map((row) => row.value), barMaxWidth: 28,
        itemStyle: { color: css("--s1"), borderRadius: [2, 2, 0, 0] } },
      { name: "收貨數量", type: "line", yAxisIndex: 1, data: (data?.monthly ?? []).map((row) => row.quantity),
        smooth: true, symbolSize: 6, lineStyle: { color: css("--s2"), width: 2 }, itemStyle: { color: css("--s2") } },
    ],
  }), [data?.monthly, theme]);

  const sort = (column: typeof sortBy) => {
    if (sortBy === column) setSortDir((current) => current === "asc" ? "desc" : "asc");
    else { setSortBy(column); setSortDir("desc"); }
    setPage(1);
  };
  const sortIcon = (column: typeof sortBy) => sortBy !== column ? null : sortDir === "asc" ? <ArrowUp size={12} /> : <ArrowDown size={12} />;

  const submitSearch = (event: FormEvent) => {
    event.preventDefault();
    setSearch(searchDraft.trim());
    setPage(1);
  };

  const exportCsv = async () => {
    setExporting(true);
    try {
      const rows = await utils.procurement.export.fetch({ ...filters, search });
      const header = ["PO_NO", "REC_DATE", "ENTITY", "MONTH", "VENDOR", "ITEM_CODE", "PART_DESC", "STD_COMM", "QTY", "UOM", "TOTAL_USD", "STATUS", "PRODUCT_FAMILY"];
      const lines = rows.map((row) => [row.poNo, row.poRecDate, row.entity, row.month, row.stdVendor, row.itemCode,
        row.partDesc, row.stdComm, row.quantity, row.uom, row.totalUsd, row.status, row.productFamily]);
      const csv = "﻿" + [header, ...lines].map((line) => line.map(csvCell).join(",")).join("\r\n");
      const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `global-purchase-${new Date().toISOString().slice(0, 10)}.csv`;
      anchor.click();
      URL.revokeObjectURL(url);
      showToast(`已匯出 ${integer(rows.length)} 筆目前篩選明細`);
    } catch (error) {
      showToast(`匯出失敗：${error instanceof Error ? error.message : "請稍後再試"}`);
    } finally { setExporting(false); }
  };

  const summaryCards = data ? [
    { label: "總採購額", value: usd(data.summary.totalUsd), note: `${data.summary.monthChange >= 0 ? "+" : ""}${data.summary.monthChange.toFixed(1)}% 較前一資料月份`, icon: CircleDollarSign, tone: data.summary.monthChange >= 0 ? "positive" : "negative" },
    { label: "採購單數", value: integer(data.summary.purchaseOrders), note: `${integer(data.summary.records)} 筆收貨明細`, icon: Database, tone: "neutral" },
    { label: "收貨數量", value: compact(data.summary.totalQuantity), note: "所有目前篩選品項", icon: PackageSearch, tone: "neutral" },
    { label: "供應商數", value: integer(data.summary.suppliers), note: `最大供應商 ${data.summary.topVendorShare.toFixed(1)}%`, icon: Globe2, tone: "neutral" },
    { label: "採購料號", value: integer(data.summary.items), note: `Ignore ${data.summary.ignoreShare.toFixed(1)}%`, icon: Activity, tone: "neutral" },
  ] : [];

  return (
    <div className="app-shell">
      <div className="ambient-grid" aria-hidden="true" />
      <header className="topbar">
        <div className="brand-mark" aria-hidden="true"><Globe2 size={24} /></div>
        <div className="brand-copy">
          <span className="eyebrow">GLOBAL PURCHASE REPORT · 2026</span>
          <h1>全球採購戰情室</h1>
        </div>
        <div className="topbar-period"><span>資料期間</span><strong>{options?.months.length ? `${options.months[0]} — ${options.months.at(-1)}` : "2026-01 — 2026-09"}</strong></div>
        <div className={`connection ${overviewQuery.isError ? "offline" : ""}`}>
          <i />{overviewQuery.isError ? "資料庫連線異常" : "資料庫已連線"}
        </div>
        <time className="clock" dateTime={clock.toISOString()}>{clock.toLocaleTimeString("zh-TW", { hour12: false })}<small>{clock.toLocaleDateString("zh-TW", { year: "numeric", month: "2-digit", day: "2-digit" })}</small></time>
        <button className="icon-btn" onClick={() => setTheme((current) => current === "dark" ? "light" : "dark")} aria-label={theme === "dark" ? "切換淺色主題" : "切換深色主題"}>
          {theme === "dark" ? <Sun size={18} /> : <Moon size={18} />}
        </button>
        <button className="btn primary create-btn" onClick={() => setCreateOpen(true)}><Plus size={16} />新增採購</button>
      </header>

      <main>
        <section className="filterbar" aria-label="採購篩選">
          <div className="filter-title"><span className="eyebrow">ANALYSIS SCOPE</span><strong>分析範圍</strong></div>
          <label>企業實體<select value={filters.entity ?? ""} onChange={(event) => updateFilter("entity", event.target.value)}><option value="">全部實體</option>{options?.entities.map((value) => <option key={value}>{value}</option>)}</select></label>
          <label>月份<select value={filters.month ?? ""} onChange={(event) => updateFilter("month", event.target.value)}><option value="">全部月份</option>{options?.months.map((value) => <option key={value}>{value}</option>)}</select></label>
          <label>標準品類<select value={filters.stdComm ?? ""} onChange={(event) => updateFilter("stdComm", event.target.value)}><option value="">全部品類</option>{options?.categories.map((value) => <option key={value}>{value}</option>)}</select></label>
          <label>供應商<select value={filters.vendor ?? ""} onChange={(event) => updateFilter("vendor", event.target.value)}><option value="">全部供應商</option>{options?.vendors.map((value) => <option key={value}>{value}</option>)}</select></label>
          <label>分析狀態<select value={filters.status ?? ""} onChange={(event) => updateFilter("status", event.target.value)}><option value="">全部狀態</option><option>Consider</option><option>Ignore</option></select></label>
          <button className="btn reset-btn" onClick={resetFilters} disabled={!activeFilterCount}><RotateCcw size={15} />重設{activeFilterCount ? ` (${activeFilterCount})` : ""}</button>
        </section>

        {overviewQuery.isError ? (
          <section className="error-state"><AlertTriangle size={28} /><h2>無法讀取採購資料</h2><p>{overviewQuery.error.message}</p><button className="btn primary" onClick={() => overviewQuery.refetch()}>重新連線</button></section>
        ) : !data ? (
          <section className="loading-state"><LoaderCircle className="spin" size={28} /><span>正在計算全球採購指標…</span></section>
        ) : (
          <>
            <section className="kpi-grid" aria-label="採購關鍵指標">
              {summaryCards.map(({ label, value, note, icon: Icon, tone }) => <article className="kpi-card" key={label}>
                <div className="kpi-label"><span>{label}</span><Icon size={16} /></div><strong>{value}</strong><small className={tone}>{note}</small>
              </article>)}
            </section>

            <div className="cockpit-grid">
              <div className="left-rail">
                <section className="panel entity-panel">
                  <header className="panel-head"><div><span className="eyebrow">ENTITY PERFORMANCE</span><h2>企業實體採購額</h2></div><span className="panel-hint">點擊下鑽</span></header>
                  <div className="rank-list">
                    {data.entities.map((row, index) => <button key={row.name} className={`rank-row ${filters.entity === row.name ? "selected" : ""}`} onClick={() => updateFilter("entity", filters.entity === row.name ? "" : row.name)}>
                      <span className="rank-no">{String(index + 1).padStart(2, "0")}</span><strong>{row.name}</strong>
                      <span className="rank-track"><i style={{ width: `${row.value / data.entities[0].value * 100}%` }} /></span>
                      <span className="rank-value">{usd(row.value)}</span><span className="rank-share">{(row.value / data.summary.totalUsd * 100).toFixed(1)}%</span>
                    </button>)}
                  </div>
                </section>
                <section className="panel category-panel">
                  <header className="panel-head"><div><span className="eyebrow">COMMODITY MIX</span><h2>品類採購結構</h2></div><span className="panel-hint">USD</span></header>
                  <Chart option={categoryOption} label="品類採購額圓環圖" className="category-chart" onClick={(params) => { if (params.name && params.name !== "其他") updateFilter("stdComm", String(params.name)); }} />
                </section>
              </div>

              <div className="center-stage">
                <section className="panel globe-panel">
                  <header className="panel-head globe-head"><div><span className="eyebrow">GLOBAL ENTITY NETWORK</span><h2>全球採購實體網絡</h2></div>
                    <button className={`tour-btn ${touring ? "active" : ""}`} onClick={() => setTouring((current) => !current)}>{touring ? <Pause size={14} /> : <Play size={14} />}{touring ? "暫停巡覽" : "自動巡覽"}</button>
                  </header>
                  <div className="globe-stage">
                    <div className={`globe-fallback ${globeReady ? "ready" : ""} ${globeError ? "error" : ""}`}><Globe2 size={64} />{!globeReady && <span>{globeError ? "3D 地球無法啟動，請使用實體排行下鑽" : "正在建立全球實體網絡…"}</span>}</div>
                    <div ref={earthHost} id="earth" aria-label="可旋轉的全球採購實體地球" />
                    <div className="hud-corners" aria-hidden="true"><i /><i /><i /><i /></div>
                    <div className="globe-readout"><span>{filters.entity ? "目前實體" : "網絡總覽"}</span><strong>{filters.entity ?? `${data.entities.length} 個實體`}</strong><b>{usd(data.summary.totalUsd)}</b></div>
                    <span className="globe-help">拖曳旋轉 · 點擊節點篩選</span>
                  </div>
                  <footer className="globe-note">節點地理位置為 Demo 視覺示意；採購金額與排行均來自上傳報表。</footer>
                </section>
                <section className="panel trend-panel">
                  <header className="panel-head"><div><span className="eyebrow">MONTHLY PROCUREMENT</span><h2>月度採購額與收貨數量</h2></div><span className="panel-hint">{filters.entity || filters.vendor || "目前範圍"}</span></header>
                  <Chart option={trendOption} label="月度採購額長條圖與收貨數量折線圖" className="trend-chart" />
                </section>
              </div>

              <div className="right-rail">
                <section className="panel supplier-panel">
                  <header className="panel-head"><div><span className="eyebrow">SUPPLIER CONCENTRATION</span><h2>供應商採購排行</h2></div><span className="panel-hint">TOP {Math.min(8, data.vendors.length)}</span></header>
                  <div className="supplier-list">{data.vendors.slice(0, 8).map((row, index) => <button key={row.name} className={filters.vendor === row.name ? "selected" : ""} onClick={() => updateFilter("vendor", filters.vendor === row.name ? "" : row.name)}>
                    <span className="supplier-index">{String(index + 1).padStart(2, "0")}</span><span className="supplier-name"><strong title={row.name}>{row.name}</strong><small>{integer(row.records)} 筆明細</small></span>
                    <span className="supplier-value"><strong>{usd(row.value)}</strong><small>{(row.value / data.summary.totalUsd * 100).toFixed(1)}%</small></span>
                  </button>)}</div>
                </section>
                <section className="panel alerts-panel">
                  <header className="panel-head"><div><span className="eyebrow">PROCUREMENT SIGNALS</span><h2>採購訊號</h2></div><span className="signal-count">{data.alerts.filter((alert) => alert.severity !== "info").length} 項注意</span></header>
                  <div className="alert-list">{data.alerts.map((alert) => <article className={`alert-item ${alert.severity}`} key={alert.id}>
                    <div><AlertTriangle size={15} /><strong>{alert.title}</strong></div><p>{alert.detail}</p>
                  </article>)}</div>
                  <footer className="alert-foot">依目前篩選資料即時計算，非外部即時風險通知。</footer>
                </section>
              </div>
            </div>
          </>
        )}

        <section className="panel records-panel">
          <header className="records-head">
            <div><span className="eyebrow">PURCHASE ORDER DETAILS</span><h2>採購明細 <span>{integer(records?.total ?? 0)} 筆</span></h2></div>
            <form className="search-form" onSubmit={submitSearch}><Search size={15} /><input type="search" value={searchDraft} onChange={(event) => setSearchDraft(event.target.value)} placeholder="搜尋 PO、料號、品項或供應商" aria-label="搜尋採購明細" />{searchDraft && <button type="button" className="search-clear" onClick={() => { setSearchDraft(""); setSearch(""); }} aria-label="清除搜尋"><X size={14} /></button>}<button className="btn" type="submit">搜尋</button></form>
            <button className="btn export-btn" onClick={exportCsv} disabled={exporting || !records?.total}>{exporting ? <LoaderCircle className="spin" size={15} /> : <Download size={15} />}{exporting ? "匯出中…" : "匯出目前明細"}</button>
          </header>
          <div className="table-wrap">
            <table>
              <thead><tr>
                <th><button onClick={() => sort("poRecDate")}>收貨日期{sortIcon("poRecDate")}</button></th>
                <th><button onClick={() => sort("poNo")}>採購單號{sortIcon("poNo")}</button></th>
                <th className="optional-col">實體</th><th>供應商</th><th>料號／品項</th><th className="optional-col">品類</th>
                <th className="n optional-col"><button onClick={() => sort("quantity")}>數量{sortIcon("quantity")}</button></th>
                <th className="n"><button onClick={() => sort("totalUsd")}>採購額 USD{sortIcon("totalUsd")}</button></th><th>狀態</th><th><span className="sr-only">操作</span></th>
              </tr></thead>
              <tbody>
                {records?.rows.map((row) => <tr key={row.id} onClick={() => setSelected(row)}>
                  <td className="mono">{row.poRecDate}</td><td className="mono po-cell">{row.poNo}</td><td className="optional-col"><span className="entity-tag">{row.entity}</span></td>
                  <td><strong className="vendor-cell" title={row.stdVendor}>{row.stdVendor}</strong></td>
                  <td><div className="item-cell"><span>{row.itemCode}</span><small title={row.partDesc}>{row.partDesc}</small></div></td>
                  <td className="optional-col"><span className="tag">{row.stdComm}</span></td><td className="n mono optional-col">{integer(row.quantity)} <small>{row.uom}</small></td>
                  <td className="n mono amount-cell">{decimal(row.totalUsd)}</td><td><span className={`status-pill ${row.status.toLowerCase()}`}>{row.status}</span></td>
                  <td><button className="row-open" onClick={(event) => { event.stopPropagation(); setSelected(row); }} aria-label={`查看採購單 ${row.poNo}`}><ChevronRight size={17} /></button></td>
                </tr>)}
              </tbody>
            </table>
            {recordsQuery.isLoading && <div className="table-state"><LoaderCircle className="spin" size={22} />正在讀取明細…</div>}
            {records && !records.rows.length && <div className="table-state"><PackageSearch size={24} /><strong>沒有符合條件的明細</strong><button className="btn" onClick={resetFilters}>清除所有篩選</button></div>}
          </div>
          <footer className="records-footer"><span>第 {records?.page ?? page} / {totalPages} 頁 · 每頁 12 筆</span><div><button className="btn" disabled={page <= 1} onClick={() => setPage((current) => Math.max(1, current - 1))}><ChevronLeft size={15} />上一頁</button><button className="btn" disabled={page >= totalPages} onClick={() => setPage((current) => Math.min(totalPages, current + 1))}>下一頁<ChevronRight size={15} /></button></div></footer>
        </section>
      </main>

      <footer className="app-footer"><span>GLOBAL PURCHASE REPORT · DATABASE-BACKED DEMO</span><span>資料來源：FIN Global Purchase Report · 2026.01–09</span></footer>

      <CreatePurchaseModal open={createOpen} onClose={closeCreate} entities={options?.entities ?? []} categories={options?.categories ?? []} defaultMonth={options?.months.at(-1) ?? "2026-09"} onSaved={showToast} />
      <PurchaseDrawer record={selected} onClose={closeDrawer} onSaved={showToast} />
      {toast && <div className="toast" role="status"><span>{toast}</span><button onClick={() => setToast("")} aria-label="關閉通知"><X size={15} /></button></div>}
    </div>
  );
}
