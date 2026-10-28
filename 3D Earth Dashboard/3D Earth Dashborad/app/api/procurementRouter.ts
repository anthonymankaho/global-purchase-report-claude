import { z } from "zod";
import { createRouter, publicQuery } from "./middleware";
import {
  createPurchase,
  exportPurchases,
  getFilterOptions,
  getProcurementOverview,
  listPurchases,
  updatePurchaseStatus,
} from "./queries/procurement";

const filters = z.object({
  entity: z.string().optional(),
  month: z.string().regex(/^\d{4}-\d{2}$/).optional(),
  stdComm: z.string().optional(),
  vendor: z.string().optional(),
  status: z.enum(["Consider", "Ignore"]).optional(),
});

const recordFilters = filters.extend({ search: z.string().max(160).optional() });

export const procurementRouter = createRouter({
  options: publicQuery.query(() => getFilterOptions()),
  overview: publicQuery.input(filters).query(({ input }) => getProcurementOverview(input)),
  records: publicQuery.input(recordFilters.extend({
    page: z.number().int().min(1).default(1),
    pageSize: z.number().int().min(5).max(100).default(12),
    sortBy: z.enum(["poRecDate", "totalUsd", "quantity", "poNo"]).default("poRecDate"),
    sortDir: z.enum(["asc", "desc"]).default("desc"),
  })).query(({ input }) => listPurchases(input)),
  export: publicQuery.input(recordFilters).query(({ input }) => exportPurchases(input)),
  updateStatus: publicQuery.input(z.object({
    id: z.number().int().positive(),
    status: z.enum(["Consider", "Ignore"]),
  })).mutation(({ input }) => updatePurchaseStatus(input.id, input.status)),
  create: publicQuery.input(z.object({
    poNo: z.string().trim().min(1).max(32),
    poRecDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
    poIssDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/).optional(),
    vendorId: z.string().max(64).optional(),
    vendorName: z.string().trim().min(1).max(255),
    itemCode: z.string().trim().min(1).max(128),
    partDesc: z.string().trim().min(1).max(512),
    commCode: z.string().max(32).optional(),
    quantity: z.number().positive(),
    uom: z.string().max(16).optional(),
    currency: z.string().max(8).default("USD"),
    unitPriceUsd: z.number().nonnegative(),
    totalUsd: z.number().positive(),
    entity: z.string().trim().min(1).max(16),
    stdVendor: z.string().trim().min(1).max(255),
    stdComm: z.string().trim().min(1).max(32),
    month: z.string().regex(/^\d{4}-\d{2}$/),
    status: z.enum(["Consider", "Ignore"]),
    productFamily: z.string().max(255).optional(),
    productFamilyCode: z.string().max(64).optional(),
  })).mutation(({ input }) => createPurchase(input)),
});
