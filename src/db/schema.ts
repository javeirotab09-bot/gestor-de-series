import {
  integer,
  pgTable,
  serial,
  text,
  timestamp,
  index,
  uniqueIndex,
} from "drizzle-orm/pg-core";

export const importBatches = pgTable("import_batches", {
  id: serial("id").primaryKey(),
  fileName: text("file_name").notNull(),
  totalRows: integer("total_rows").notNull().default(0),
  validRows: integer("valid_rows").notNull().default(0),
  createdAt: timestamp("created_at", { withTimezone: true }).defaultNow().notNull(),
});

export const seriesAssociations = pgTable(
  "series_associations",
  {
    id: serial("id").primaryKey(),
    importBatchId: integer("import_batch_id")
      .notNull()
      .references(() => importBatches.id, { onDelete: "cascade" }),
    serie1: text("serie_1").notNull(),
    serie2: text("serie_2").notNull(),
    codigoSap: text("codigo_sap"),
    descripcion: text("descripcion"),
    cantidad: integer("cantidad").notNull().default(1),
    createdAt: timestamp("created_at", { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => [
    index("series_assoc_import_batch_idx").on(table.importBatchId),
    index("series_assoc_serie1_idx").on(table.serie1),
    index("series_assoc_serie2_idx").on(table.serie2),
    uniqueIndex("series_assoc_import_serie1_unique").on(table.importBatchId, table.serie1),
    uniqueIndex("series_assoc_import_serie2_unique").on(table.importBatchId, table.serie2),
  ],
);

export const scanSessions = pgTable("scan_sessions", {
  id: serial("id").primaryKey(),
  importBatchId: integer("import_batch_id")
    .notNull()
    .references(() => importBatches.id, { onDelete: "restrict" }),
  lpn: text("lpn").notNull(),
  sku: text("sku").notNull(),
  packaging: text("packaging").notNull(),
  targetQuantity: integer("target_quantity").notNull().default(0),
  status: text("status").notNull().default("active"),
  startedAt: timestamp("started_at", { withTimezone: true }).defaultNow().notNull(),
  finishedAt: timestamp("finished_at", { withTimezone: true }),
});

export const scanEvents = pgTable(
  "scan_events",
  {
    id: serial("id").primaryKey(),
    sessionId: integer("session_id")
      .notNull()
      .references(() => scanSessions.id, { onDelete: "cascade" }),
    scannedSeries: text("scanned_series").notNull(),
    matchedSeries: text("matched_series"),
    serie1: text("serie_1"),
    serie2: text("serie_2"),
    codigoSap: text("codigo_sap"),
    descripcion: text("descripcion"),
    cantidad: integer("cantidad"),
    status: text("status").notNull(),
    message: text("message").notNull(),
    createdAt: timestamp("created_at", { withTimezone: true }).defaultNow().notNull(),
  },
  (table) => [
    index("scan_events_session_idx").on(table.sessionId),
    index("scan_events_scanned_series_idx").on(table.scannedSeries),
  ],
);
