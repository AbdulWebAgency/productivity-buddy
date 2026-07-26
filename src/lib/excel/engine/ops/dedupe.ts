import type { PlanOp } from "../../types";
import { resolveColumn } from "../shared/headers";
import type { CellValue, SheetGrid } from "../shared/workbook";
import { registerOp } from "../registry";

export function opDedupe(
  grid: SheetGrid,
  op: Extract<PlanOp, { op: "dedupe" }>,
): { grid: SheetGrid; removed: number } {
  const seen = new Set<string>();
  const kept: CellValue[][] = [];
  let removed = 0;
  const keyIdx = op.keyColumn ? resolveColumn(grid.headers, op.keyColumn) : -1;
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

registerOp<Extract<PlanOp, { op: "dedupe" }>>("dedupe", (op, ctx) => {
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
});
