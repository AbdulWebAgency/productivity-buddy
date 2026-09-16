import type { PlanOp } from "../../types";
import { resolveKeyColumn } from "../shared/key-resolution";
import type { CellValue, SheetGrid } from "../shared/workbook";
import type { OpHandler } from "../registry";

export function opDedupe(
  grid: SheetGrid,
  op: Extract<PlanOp, { op: "dedupe" }>,
): { grid: SheetGrid; removed: number; keyWarning?: string } {
  const seen = new Set<string>();
  const kept: CellValue[][] = [];
  let removed = 0;
  let keyIdx = -1;
  let keyWarning: string | undefined;
  if (op.keyColumn) {
    const r = resolveKeyColumn(grid.headers, op.keyColumn);
    if (r.status === "resolved") {
      keyIdx = r.index;
    } else if (r.status === "ambiguous") {
      keyWarning = `Dedupe key "${op.keyColumn}" is ambiguous (${r.candidates
        .map((c) => `"${c.header}"`)
        .join(", ")}) — removed full-row duplicates instead.`;
    } else {
      keyWarning = `Dedupe key "${op.keyColumn}" was not found — removed full-row duplicates instead.`;
    }
  }
  for (const r of grid.rows) {
    const key =
      op.strategy === "key" && keyIdx >= 0
        ? String(r[keyIdx] ?? "")
        : r.map((v) => (v == null ? "" : String(v))).join("\u0001");
    if (seen.has(key)) {
      removed++;
      continue;
    }
    seen.add(key);
    kept.push(r);
  }
  return { grid: { headers: grid.headers, rows: kept }, removed };
}

export const dedupeHandler: OpHandler<Extract<PlanOp, { op: "dedupe" }>> = (op, ctx) => {
  const { grid: dg, removed } = opDedupe(ctx.state.currentGrid, op);
  const rowsIn = ctx.state.currentGrid.rows.length + removed - dg.rows.length + dg.rows.length;
  // rowsIn = original count; preserve original log semantics
  const originalCount = ctx.state.currentGrid.rows.length;
  ctx.state.currentGrid = dg;
  ctx.opLogs.push({
    op: "dedupe",
    status: "ok",
    ms: Date.now() - ctx.started,
    strategy: op.strategy,
    keyColumn: op.keyColumn ?? null,
    rowsIn: originalCount,
    rowsOut: dg.rows.length,
    duplicatesRemoved: removed,
  });
  void rowsIn;
};
