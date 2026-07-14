import { z } from "zod";

// The AI produces one of these plans. It CANNOT invent arbitrary ops.
export const OpSchema = z.discriminatedUnion("op", [
  z.object({
    op: z.literal("merge"),
    keyColumn: z.string().min(1),
    strategy: z.enum(["union", "intersection"]).default("union"),
    highlightUnmatched: z.boolean().default(true),
  }),
  z.object({
    op: z.literal("dedupe"),
    strategy: z.enum(["key", "full_row"]).default("key"),
    keyColumn: z.string().optional(),
  }),
  z.object({
    op: z.literal("diff"),
    keyColumn: z.string().min(1),
    fileAIndex: z.number().int().min(0).default(0),
    fileBIndex: z.number().int().min(0).default(1),
  }),
  z.object({
    op: z.literal("intersection"),
    keyColumn: z.string().min(1),
    fileAIndex: z.number().int().min(0).default(0),
    fileBIndex: z.number().int().min(0).default(1),
  }),
  z.object({
    op: z.literal("summary"),
    includeCharts: z.boolean().default(true),
  }),
  z.object({
    op: z.literal("recalc"),
  }),
  z.object({
    op: z.literal("highlight_column"),
    column: z.string().min(1),
    rule: z.enum(["missing", "duplicate", "outlier"]),
  }),
]);

export const PlanSchema = z.object({
  summary: z.string().max(500).default(""),
  ops: z.array(OpSchema).min(1).max(10),
  columnMappings: z
    .array(
      z.object({
        canonical: z.string(),
        perFile: z.array(z.object({ fileIndex: z.number().int().min(0), column: z.string() })),
      }),
    )
    .default([]),
  warnings: z.array(z.string()).default([]),
});

export type Plan = z.infer<typeof PlanSchema>;
export type PlanOp = z.infer<typeof OpSchema>;

export const SheetMetaSchema = z.object({
  sheets: z.array(
    z.object({
      name: z.string(),
      rowCount: z.number(),
      columnCount: z.number(),
      headers: z.array(z.string()),
      sampleRows: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).max(3),
    }),
  ),
});

export type SheetMeta = z.infer<typeof SheetMetaSchema>;
