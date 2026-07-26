import type { PlanOp } from "../../types";
import { resolveColumn } from "../shared/headers";
import type { CellValue, SheetGrid } from "../shared/workbook";
import { writeGridToSheet } from "../shared/workbook";
import { registerOp, applyProjection } from "../registry";

export function opIntersection(
  a: SheetGrid,
  b: SheetGrid,
  op: Extract<PlanOp, { op: "intersection" }>,
  namesA: string,
  namesB: string,
): { grid: SheetGrid; stats: Record<string, number> } {
  const kA = resolveColumn(a.headers, op.keyColumn);
  const kB = resolveColumn(b.headers, op.keyColumn);
  if (kA < 0) throw new Error(`Key column "${op.keyColumn}" not found in ${namesA}`);
  if (kB < 0) throw new Error(`Key column "${op.keyColumn}" not found in ${namesB}`);

  const keysB = new Set<string>();
  for (const r of b.rows) {
    const k = String(r[kB] ?? "").trim().toLowerCase();
    if (k) keysB.add(k);
  }
  const rows: CellValue[][] = [];
  const emittedKeys = new Set<string>();
  for (const r of a.rows) {
    const kRaw = String(r[kA] ?? "").trim();
    const k = kRaw.toLowerCase();
    if (k && keysB.has(k) && !emittedKeys.has(k)) {
      emittedKeys.add(k);
      rows.push(r);
    }
  }
  return {
    grid: { headers: a.headers, rows },
    stats: { inA: a.rows.length, inB: b.rows.length, common: rows.length },
  };
}

registerOp<Extract<PlanOp, { op: "intersection" }>>("intersection", (op, ctx) => {
  const aIdx = op.fileAIndex;
  const bIdx = op.fileBIndex;
  const a = ctx.grids[aIdx];
  const b = ctx.grids[bIdx];
  if (!a || !b) {
    ctx.warnings.push(`intersection skipped: file index ${aIdx}/${bIdx} out of range`);
    ctx.opLogs.push({ op: "intersection", status: "skipped", reason: "file index out of range" });
    return;
  }
  if (aIdx === bIdx) {
    ctx.warnings.push("intersection skipped: same file on both sides");
    ctx.opLogs.push({ op: "intersection", status: "skipped", reason: "same file on both sides" });
    return;
  }
  const nameA = ctx.files[aIdx].name;
  const nameB = ctx.files[bIdx].name;
  const inter = opIntersection(a, b, op, nameA, nameB);
  const sheetName = "Common rows";
  const projected = applyProjection(ctx, sheetName, inter.grid, { alwaysKeep: [op.keyColumn] });
  writeGridToSheet(ctx.outWb, sheetName, projected);
  ctx.state.producedSheets++;
  ctx.opLogs.push({
    op: "intersection",
    status: "ok",
    ms: Date.now() - ctx.started,
    keyColumn: op.keyColumn,
    fileA: nameA,
    fileB: nameB,
    ...inter.stats,
    sheet: sheetName,
  });
});
