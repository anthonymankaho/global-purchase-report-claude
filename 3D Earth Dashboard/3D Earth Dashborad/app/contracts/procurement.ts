export type { Purchase } from "../db/schema";

export interface PurchaseFilters {
  entity?: string;
  month?: string;
  stdComm?: string;
  vendor?: string;
  status?: "Consider" | "Ignore";
}

export interface PurchaseRecordInput extends PurchaseFilters {
  search?: string;
}

export interface ProcurementAlert {
  id: string;
  severity: "critical" | "warning" | "info";
  title: string;
  detail: string;
}

export const ENTITY_HUBS: Record<string, string> = {
  MIL: "曼谷",
  MHK: "香港",
  STC: "上海",
  TVS: "罗马",
  MZQ: "深圳",
  MAZ: "利马",
  RACO: "广州",
};
