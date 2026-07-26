// Thin public facade. All implementation lives under ./engine/*.
// Kept here so existing imports (`@/lib/excel/engine.server`) keep resolving.

export type { CellValue, SheetGrid } from "./engine/shared/workbook";
export {
  readWorkbook,
  extractSheetMeta,
  sheetToGrid,
  writeGridToSheet,
  safeSheetName,
} from "./engine/shared/workbook";

export type { EngineFile, EngineResult } from "./engine/types";
export { runPlan } from "./engine/runPlan";

// Individual op functions (re-exported for tests and external callers).
export { opDedupe } from "./engine/ops/dedupe";
export { opSummary } from "./engine/ops/summary";
export { opIntersection } from "./engine/ops/intersection";
export { opDiff, type DiffResult } from "./engine/ops/diff";
export { opMerge } from "./engine/ops/merge";
export { opBulkLookup, type BulkLookupResult } from "./engine/ops/bulkLookup";
export { opMasterMerge, type MasterMergeResult } from "./engine/ops/masterMerge";
export { recalcFormulas } from "./engine/ops/highlight";
