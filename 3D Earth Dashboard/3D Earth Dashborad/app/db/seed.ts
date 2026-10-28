import { readFileSync } from "node:fs";
import { count } from "drizzle-orm";
import { getDb } from "../api/queries/connection";
import { purchases, type InsertPurchase } from "./schema";

interface SeedFile {
  fields: string[];
  rows: unknown[][];
}

const optional = (value: unknown) => value === null || value === undefined || value === ""
  ? undefined
  : String(value);
const required = (value: unknown) => String(value ?? "");
const numeric = (value: unknown) => Number(value ?? 0);

async function seed() {
  const db = getDb();
  const [existing] = await db.select({ value: count() }).from(purchases);
  if (Number(existing.value) > 0) {
    console.log(`Purchase table already contains ${existing.value} records; seed skipped.`);
    process.exit(0);
  }

  const source = JSON.parse(readFileSync(new URL("./report-seed.json", import.meta.url), "utf8")) as SeedFile;
  const indexes = Object.fromEntries(source.fields.map((field, index) => [field, index]));
  const get = (row: unknown[], field: string) => row[indexes[field]];
  const records: InsertPurchase[] = source.rows.map((row) => ({
    poNo: required(get(row, "PO_NO")),
    poRecDate: required(get(row, "PO_REC_DATE")),
    poIssDate: optional(get(row, "PO_ISS_DATE")),
    vendorId: optional(get(row, "VENDOR_ID")),
    vendorName: required(get(row, "VENDOR_NAME")),
    itemCode: required(get(row, "ITEM_CODE")),
    partDesc: required(get(row, "PART_DESC")),
    commCode: optional(get(row, "COMM_CODE")),
    quantity: numeric(get(row, "REC_QTY")),
    uom: optional(get(row, "UOM")),
    currency: optional(get(row, "PO_CUR")) ?? "USD",
    unitPriceUsd: numeric(get(row, "PO_PRICE_USD")),
    totalUsd: numeric(get(row, "TOT_PO_PRICE_USD")),
    entity: required(get(row, "DB")),
    stdVendor: required(get(row, "Std Vendor Name")),
    stdComm: required(get(row, "Std Comm")),
    month: required(get(row, "yyyy-mm")),
    status: required(get(row, "ignore / consider")),
    productFamily: optional(get(row, "ProdFamiDesc")),
    productFamilyCode: optional(get(row, "ProdFamiCode")),
  }));

  console.log(`Seeding ${records.length.toLocaleString()} procurement records...`);
  const batchSize = 400;
  for (let index = 0; index < records.length; index += batchSize) {
    await db.insert(purchases).values(records.slice(index, index + batchSize));
    if ((index / batchSize) % 20 === 0) {
      console.log(`${Math.min(index + batchSize, records.length).toLocaleString()} / ${records.length.toLocaleString()}`);
    }
  }
  console.log("Seed completed.");
  process.exit(0);
}

seed().catch((error) => {
  console.error(error);
  process.exit(1);
});
