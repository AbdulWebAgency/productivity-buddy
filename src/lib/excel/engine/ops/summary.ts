import type { PlanOp } from "../../types";
import type { CellValue, SheetGrid } from "../shared/workbook";
import { writeGridToSheet } from "../shared/workbook";
import type { OpHandler } from "../registry";

export function opSummary(grid: SheetGrid): { headers: string[]; rows: CellValue[][] } {
  const rows: CellValue[][] = [
    ["Metric", "Value"],
    ["Total rows", grid.rows.length],
    ["Total columns", grid.headers.length],
  ];
  grid.headers.forEach((h, i) => {
    const nums = grid.rows.map((r) => Number(r[i])).filter((n) => Number.isFinite(n));
    if (nums.length >= Math.max(3, grid.rows.length * 0.3)) {
      const sum = nums.reduce((a, b) => a + b, 0);
      rows.push([`${h} — sum`, sum]);
      rows.push([`${h} — avg`, Number((sum / nums.length).toFixed(4))]);
      rows.push([`${h} — min`, Math.min(...nums)]);
      rows.push([`${h} — max`, Math.max(...nums)]);
    }
  });
  return { headers: rows[0].map(String), rows: rows.slice(1) };
}

export const summaryHandler: OpHandler<Extract<PlanOp, { op: "summary" }>> = (_op, ctx) => {
  const s = opSummary(ctx.state.currentGrid);
  writeGridToSheet(ctx.outWb, "Summary", s);
  ctx.state.producedSheets++;
  ctx.opLogs.push({
    op: "summary",
    status: "ok",
    ms: Date.now() - ctx.started,
    metrics: s.rows.length,
    sheet: "Summary",
  });
};
