import {
  double,
  index,
  mysqlTable,
  serial,
  timestamp,
  varchar,
} from "drizzle-orm/mysql-core";

export const purchases = mysqlTable(
  "purchases",
  {
    id: serial("id").primaryKey(),
    poNo: varchar("po_no", { length: 32 }).notNull(),
    poRecDate: varchar("po_rec_date", { length: 10 }).notNull(),
    poIssDate: varchar("po_iss_date", { length: 10 }),
    vendorId: varchar("vendor_id", { length: 64 }),
    vendorName: varchar("vendor_name", { length: 255 }).notNull(),
    itemCode: varchar("item_code", { length: 128 }).notNull(),
    partDesc: varchar("part_desc", { length: 512 }).notNull(),
    commCode: varchar("comm_code", { length: 32 }),
    quantity: double("quantity").notNull(),
    uom: varchar("uom", { length: 16 }),
    currency: varchar("currency", { length: 8 }),
    unitPriceUsd: double("unit_price_usd").notNull(),
    totalUsd: double("total_usd").notNull(),
    entity: varchar("entity", { length: 16 }).notNull(),
    stdVendor: varchar("std_vendor", { length: 255 }).notNull(),
    stdComm: varchar("std_comm", { length: 32 }).notNull(),
    month: varchar("month", { length: 7 }).notNull(),
    status: varchar("status", { length: 16 }).notNull(),
    productFamily: varchar("product_family", { length: 255 }),
    productFamilyCode: varchar("product_family_code", { length: 64 }),
    createdAt: timestamp("created_at").defaultNow().notNull(),
    updatedAt: timestamp("updated_at").defaultNow().notNull(),
  },
  (table) => ({
    monthIdx: index("purchases_month_idx").on(table.month),
    entityIdx: index("purchases_entity_idx").on(table.entity),
    vendorIdx: index("purchases_vendor_idx").on(table.stdVendor),
    commIdx: index("purchases_comm_idx").on(table.stdComm),
    poIdx: index("purchases_po_idx").on(table.poNo),
    itemIdx: index("purchases_item_idx").on(table.itemCode),
  }),
);

export type Purchase = typeof purchases.$inferSelect;
export type InsertPurchase = typeof purchases.$inferInsert;
