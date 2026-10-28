import {
  and,
  asc,
  count,
  desc,
  eq,
  like,
  or,
  sql,
  type SQL,
} from "drizzle-orm";
import { purchases, type InsertPurchase } from "@db/schema";
import type {
  ProcurementAlert,
  PurchaseFilters,
  PurchaseRecordInput,
} from "@contracts/procurement";
import { getDb } from "./connection";

const money = sql<number>`coalesce(sum(${purchases.totalUsd}), 0)`;
const quantity = sql<number>`coalesce(sum(${purchases.quantity}), 0)`;

function filterConditions(filters: PurchaseRecordInput): SQL[] {
  const conditions: SQL[] = [];
  if (filters.entity) conditions.push(eq(purchases.entity, filters.entity));
  if (filters.month) conditions.push(eq(purchases.month, filters.month));
  if (filters.stdComm) conditions.push(eq(purchases.stdComm, filters.stdComm));
  if (filters.vendor) conditions.push(eq(purchases.stdVendor, filters.vendor));
  if (filters.status) conditions.push(eq(purchases.status, filters.status));
  const search = filters.search?.trim();
  if (search) {
    const term = `%${search}%`;
    const searchCondition = or(
      like(purchases.poNo, term),
      like(purchases.itemCode, term),
      like(purchases.partDesc, term),
      like(purchases.vendorName, term),
      like(purchases.stdVendor, term),
    );
    if (searchCondition) conditions.push(searchCondition);
  }
  return conditions;
}

function where(filters: PurchaseRecordInput) {
  const conditions = filterConditions(filters);
  return conditions.length ? and(...conditions) : undefined;
}

export async function getFilterOptions() {
  const db = getDb();
  const [entities, months, categories, vendors] = await Promise.all([
    db.selectDistinct({ value: purchases.entity }).from(purchases).orderBy(asc(purchases.entity)),
    db.selectDistinct({ value: purchases.month }).from(purchases).orderBy(asc(purchases.month)),
    db.selectDistinct({ value: purchases.stdComm }).from(purchases).orderBy(asc(purchases.stdComm)),
    db.selectDistinct({ value: purchases.stdVendor }).from(purchases).orderBy(asc(purchases.stdVendor)),
  ]);
  return {
    entities: entities.map((row) => row.value),
    months: months.map((row) => row.value),
    categories: categories.map((row) => row.value),
    vendors: vendors.map((row) => row.value),
  };
}

export async function getProcurementOverview(filters: PurchaseFilters) {
  const db = getDb();
  const condition = where(filters);
  const [totals, monthly, entities, categories, vendors, statuses] = await Promise.all([
    db.select({
      totalUsd: money,
      totalQuantity: quantity,
      records: count(),
      purchaseOrders: sql<number>`count(distinct ${purchases.poNo})`,
      suppliers: sql<number>`count(distinct ${purchases.stdVendor})`,
      items: sql<number>`count(distinct ${purchases.itemCode})`,
    }).from(purchases).where(condition),
    db.select({ name: purchases.month, value: money, quantity, records: count() })
      .from(purchases).where(condition).groupBy(purchases.month).orderBy(asc(purchases.month)),
    db.select({ name: purchases.entity, value: money, records: count() })
      .from(purchases).where(condition).groupBy(purchases.entity).orderBy(desc(money)),
    db.select({ name: purchases.stdComm, value: money, records: count() })
      .from(purchases).where(condition).groupBy(purchases.stdComm).orderBy(desc(money)),
    db.select({ name: purchases.stdVendor, value: money, records: count() })
      .from(purchases).where(condition).groupBy(purchases.stdVendor).orderBy(desc(money)).limit(10),
    db.select({ name: purchases.status, value: money, records: count() })
      .from(purchases).where(condition).groupBy(purchases.status),
  ]);

  const summary = totals[0] ?? {
    totalUsd: 0,
    totalQuantity: 0,
    records: 0,
    purchaseOrders: 0,
    suppliers: 0,
    items: 0,
  };
  const normalized = {
    totalUsd: Number(summary.totalUsd),
    totalQuantity: Number(summary.totalQuantity),
    records: Number(summary.records),
    purchaseOrders: Number(summary.purchaseOrders),
    suppliers: Number(summary.suppliers),
    items: Number(summary.items),
  };
  const monthlyRows = monthly.map((row) => ({ ...row, value: Number(row.value), quantity: Number(row.quantity) }));
  const entityRows = entities.map((row) => ({ ...row, value: Number(row.value) }));
  const categoryRows = categories.map((row) => ({ ...row, value: Number(row.value) }));
  const vendorRows = vendors.map((row) => ({ ...row, value: Number(row.value) }));
  const ignored = statuses.find((row) => row.name === "Ignore");
  const ignoreShare = normalized.totalUsd ? Number(ignored?.value ?? 0) / normalized.totalUsd * 100 : 0;
  const topVendorShare = normalized.totalUsd && vendorRows.length ? vendorRows[0].value / normalized.totalUsd * 100 : 0;
  const topEntityShare = normalized.totalUsd && entityRows.length ? entityRows[0].value / normalized.totalUsd * 100 : 0;
  const last = monthlyRows.at(-1);
  const previous = monthlyRows.at(-2);
  const monthChange = last && previous && previous.value ? (last.value - previous.value) / previous.value * 100 : 0;

  const alerts: ProcurementAlert[] = [];
  if (normalized.records) {
    alerts.push({
      id: "vendor-concentration",
      severity: topVendorShare > 20 ? "critical" : "warning",
      title: `最大供應商占比 ${topVendorShare.toFixed(1)}%`,
      detail: `${vendorRows[0]?.name ?? "—"} · ${formatUsd(vendorRows[0]?.value ?? 0)}，建議檢視供應集中度`,
    });
    alerts.push({
      id: "month-change",
      severity: monthChange < -10 ? "critical" : monthChange < 0 ? "warning" : "info",
      title: `${last?.name ?? "最新月份"}採購額 ${monthChange >= 0 ? "成長" : "下降"} ${Math.abs(monthChange).toFixed(1)}%`,
      detail: `${formatUsd(previous?.value ?? 0)} → ${formatUsd(last?.value ?? 0)}，較前一資料月份`,
    });
    alerts.push({
      id: "entity-concentration",
      severity: topEntityShare > 35 ? "warning" : "info",
      title: `${entityRows[0]?.name ?? "—"} 為最大採購實體`,
      detail: `占目前篩選總額 ${topEntityShare.toFixed(1)}% · ${formatUsd(entityRows[0]?.value ?? 0)}`,
    });
    alerts.push({
      id: "ignored-spend",
      severity: ignoreShare > 30 ? "warning" : "info",
      title: `Ignore 金額占比 ${ignoreShare.toFixed(1)}%`,
      detail: "可切換 Consider／Ignore 狀態，核對報表分析口徑",
    });
  }

  return {
    summary: { ...normalized, monthChange, ignoreShare, topVendorShare },
    monthly: monthlyRows,
    entities: entityRows,
    categories: categoryRows,
    vendors: vendorRows,
    alerts,
    period: { start: "2026-01", end: "2026-09" },
  };
}

function formatUsd(value: number) {
  return `US$${(value / 1_000_000).toFixed(2)}M`;
}

export async function listPurchases(input: PurchaseRecordInput & {
  page: number;
  pageSize: number;
  sortBy: "poRecDate" | "totalUsd" | "quantity" | "poNo";
  sortDir: "asc" | "desc";
}) {
  const db = getDb();
  const condition = where(input);
  const sortColumn = {
    poRecDate: purchases.poRecDate,
    totalUsd: purchases.totalUsd,
    quantity: purchases.quantity,
    poNo: purchases.poNo,
  }[input.sortBy];
  const order = input.sortDir === "asc" ? asc(sortColumn) : desc(sortColumn);
  const [rows, total] = await Promise.all([
    db.select().from(purchases).where(condition).orderBy(order, desc(purchases.id))
      .limit(input.pageSize).offset((input.page - 1) * input.pageSize),
    db.select({ value: count() }).from(purchases).where(condition),
  ]);
  return { rows, total: Number(total[0]?.value ?? 0), page: input.page, pageSize: input.pageSize };
}

export async function exportPurchases(filters: PurchaseRecordInput) {
  return getDb().select({
    poNo: purchases.poNo,
    poRecDate: purchases.poRecDate,
    entity: purchases.entity,
    month: purchases.month,
    stdVendor: purchases.stdVendor,
    itemCode: purchases.itemCode,
    partDesc: purchases.partDesc,
    stdComm: purchases.stdComm,
    quantity: purchases.quantity,
    uom: purchases.uom,
    totalUsd: purchases.totalUsd,
    status: purchases.status,
    productFamily: purchases.productFamily,
  }).from(purchases).where(where(filters)).orderBy(asc(purchases.poRecDate), asc(purchases.id));
}

export async function updatePurchaseStatus(id: number, status: "Consider" | "Ignore") {
  const db = getDb();
  await db.update(purchases).set({ status, updatedAt: new Date() }).where(eq(purchases.id, id));
  const [record] = await db.select().from(purchases).where(eq(purchases.id, id)).limit(1);
  return record;
}

export async function createPurchase(data: InsertPurchase) {
  const db = getDb();
  const [{ id }] = await db.insert(purchases).values(data).$returningId();
  const [record] = await db.select().from(purchases).where(eq(purchases.id, id)).limit(1);
  return record;
}
