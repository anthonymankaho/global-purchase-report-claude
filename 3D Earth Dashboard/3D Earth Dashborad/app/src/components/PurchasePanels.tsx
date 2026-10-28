import { useEffect, useRef, useState, type FormEvent } from "react";
import { Check, ChevronRight, LoaderCircle, Plus, X } from "lucide-react";
import { trpc } from "@/providers/trpc";
import type { Purchase } from "@contracts/procurement";
import { decimal, integer, usd } from "@/lib/format";

function useOverlay(open: boolean, onClose: () => void) {
  const closeRef = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    if (!open) return;
    const previous = document.activeElement as HTMLElement | null;
    const onKey = (event: KeyboardEvent) => { if (event.key === "Escape") onClose(); };
    document.addEventListener("keydown", onKey);
    document.body.classList.add("overlay-open");
    requestAnimationFrame(() => closeRef.current?.focus());
    return () => {
      document.removeEventListener("keydown", onKey);
      document.body.classList.remove("overlay-open");
      previous?.focus();
    };
  }, [open, onClose]);
  return closeRef;
}

interface CreateProps {
  open: boolean;
  onClose: () => void;
  entities: string[];
  categories: string[];
  defaultMonth: string;
  onSaved: (message: string) => void;
}

export function CreatePurchaseModal({ open, onClose, entities, categories, defaultMonth, onSaved }: CreateProps) {
  const closeRef = useOverlay(open, onClose);
  const utils = trpc.useUtils();
  const [quantity, setQuantity] = useState("1");
  const [unitPrice, setUnitPrice] = useState("0");
  const [error, setError] = useState("");
  const create = trpc.procurement.create.useMutation({
    onSuccess: async (record) => {
      await Promise.all([
        utils.procurement.overview.invalidate(),
        utils.procurement.records.invalidate(),
        utils.procurement.options.invalidate(),
      ]);
      onSaved(`採購單 ${record.poNo} 已新增並保存`);
      onClose();
    },
    onError: (reason) => setError(reason.message),
  });

  if (!open) return null;
  const total = Number(quantity) * Number(unitPrice);
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError("");
    const form = new FormData(event.currentTarget);
    const qty = Number(form.get("quantity"));
    const price = Number(form.get("unitPriceUsd"));
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isFinite(price) || price < 0 || qty * price <= 0) {
      setError("請輸入有效的數量與單價，總金額必須大於 0。 ");
      return;
    }
    const text = (key: string) => String(form.get(key) ?? "").trim();
    create.mutate({
      poNo: text("poNo"),
      poRecDate: text("poRecDate"),
      vendorName: text("vendorName"),
      stdVendor: text("stdVendor"),
      itemCode: text("itemCode"),
      partDesc: text("partDesc"),
      entity: text("entity"),
      stdComm: text("stdComm"),
      commCode: text("stdComm"),
      month: text("month"),
      status: text("status") as "Consider" | "Ignore",
      quantity: qty,
      unitPriceUsd: price,
      totalUsd: qty * price,
      uom: "PCS",
      currency: "USD",
    });
  };

  return (
    <div className="overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <section className="modal" role="dialog" aria-modal="true" aria-labelledby="create-title">
        <header className="overlay-head">
          <div><span className="eyebrow">NEW PROCUREMENT RECORD</span><h2 id="create-title">新增採購明細</h2></div>
          <button ref={closeRef} className="icon-btn" onClick={onClose} aria-label="關閉新增採購視窗"><X size={19} /></button>
        </header>
        <form onSubmit={submit}>
          <div className="form-grid">
            <label>採購單號<input name="poNo" required maxLength={32} placeholder="例：PO-2026-001" /></label>
            <label>收貨日期<input name="poRecDate" type="date" required defaultValue={`${defaultMonth}-28`} /></label>
            <label>企業實體<select name="entity" required defaultValue={entities[0]}>{entities.map((value) => <option key={value}>{value}</option>)}</select></label>
            <label>資料月份<input name="month" type="month" required defaultValue={defaultMonth} /></label>
            <label>供應商名稱<input name="vendorName" required placeholder="原始供應商名稱" /></label>
            <label>標準供應商<input name="stdVendor" required placeholder="集團供應商歸屬" /></label>
            <label>料號<input name="itemCode" required placeholder="ITEM CODE" /></label>
            <label>標準品類<select name="stdComm" required defaultValue={categories[0]}>{categories.map((value) => <option key={value}>{value}</option>)}</select></label>
            <label className="span-2">品項描述<input name="partDesc" required placeholder="料號品項與規格" /></label>
            <label>收貨數量<input name="quantity" type="number" min="0.01" step="0.01" required value={quantity} onChange={(event) => setQuantity(event.target.value)} /></label>
            <label>美元單價<input name="unitPriceUsd" type="number" min="0" step="0.0001" required value={unitPrice} onChange={(event) => setUnitPrice(event.target.value)} /></label>
            <label>分析狀態<select name="status" defaultValue="Consider"><option>Consider</option><option>Ignore</option></select></label>
            <div className="calculated"><span>總採購額（USD）</span><strong>{Number.isFinite(total) ? decimal(total) : "—"}</strong></div>
          </div>
          {error && <p className="form-error" role="alert">{error}</p>}
          <footer className="form-actions">
            <button type="button" className="btn" onClick={onClose}>取消</button>
            <button type="submit" className="btn primary" disabled={create.isPending}>
              {create.isPending ? <LoaderCircle className="spin" size={16} /> : <Plus size={16} />}
              {create.isPending ? "保存中…" : "新增並保存"}
            </button>
          </footer>
        </form>
      </section>
    </div>
  );
}

interface DrawerProps {
  record: Purchase | null;
  onClose: () => void;
  onSaved: (message: string) => void;
}

export function PurchaseDrawer({ record, onClose, onSaved }: DrawerProps) {
  const open = Boolean(record);
  const closeRef = useOverlay(open, onClose);
  const utils = trpc.useUtils();
  const [status, setStatus] = useState<"Consider" | "Ignore">("Consider");
  useEffect(() => { if (record) setStatus(record.status as "Consider" | "Ignore"); }, [record]);
  const update = trpc.procurement.updateStatus.useMutation({
    onSuccess: async (saved) => {
      setStatus(saved.status as "Consider" | "Ignore");
      await Promise.all([utils.procurement.overview.invalidate(), utils.procurement.records.invalidate()]);
      onSaved(`狀態已保存為 ${saved.status}`);
    },
    onError: (reason) => onSaved(`保存失敗：${reason.message}`),
  });
  if (!record) return null;
  const details = [
    ["企業實體／月份", `${record.entity} · ${record.month}`],
    ["標準供應商", record.stdVendor],
    ["原始供應商", record.vendorName],
    ["品類", `${record.stdComm} · ${record.commCode || "—"}`],
    ["產品家族", record.productFamily || "未指派"],
    ["數量", `${integer(record.quantity)} ${record.uom || ""}`],
    ["美元單價", `$${decimal(record.unitPriceUsd)}`],
    ["最後更新", new Date(record.updatedAt).toLocaleString("zh-TW")],
  ];
  return (
    <div className="overlay drawer-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <aside className="drawer" role="dialog" aria-modal="true" aria-labelledby="drawer-title">
        <header className="overlay-head">
          <div><span className="eyebrow">PURCHASE ORDER</span><h2 id="drawer-title">{record.poNo}</h2></div>
          <button ref={closeRef} className="icon-btn" onClick={onClose} aria-label="關閉採購明細"><X size={19} /></button>
        </header>
        <div className="drawer-body">
          <div className="record-amount"><span>總採購額</span><strong>{usd(record.totalUsd)}</strong><small>US${decimal(record.totalUsd)}</small></div>
          <div className="record-item"><span className="tag">{record.itemCode}</span><h3>{record.partDesc}</h3><p>{record.poRecDate} 收貨</p></div>
          <dl className="detail-list">{details.map(([label, value]) => <div key={label}><dt>{label}</dt><dd>{value}</dd></div>)}</dl>
          <div className="status-editor">
            <span className="eyebrow">ANALYSIS STATUS</span>
            <h3>報表分析狀態</h3>
            <p>狀態會寫入資料庫，並同步更新所有指標與圖表。</p>
            <div className="status-options">
              {(["Consider", "Ignore"] as const).map((value) => (
                <button key={value} className={status === value ? "selected" : ""} disabled={update.isPending}
                  onClick={() => update.mutate({ id: record.id, status: value })}>
                  {status === value ? <Check size={16} /> : <ChevronRight size={16} />}{value}
                </button>
              ))}
            </div>
            {update.isPending && <span className="saving"><LoaderCircle className="spin" size={14} /> 正在保存…</span>}
          </div>
        </div>
      </aside>
    </div>
  );
}
