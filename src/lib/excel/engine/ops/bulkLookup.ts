import type { PlanOp } from "../../types";
import type { CellValue, SheetGrid } from "../shared/workbook";
import { writeGridToSheet } from "../shared/workbook";
import { registerOp, applyProjection } from "../registry";

const KEYISH_HINTS = ["id", "reg", "roll", "email", "mail", "phone", "mobile", "name", "code", "number"];

export type BulkLookupResult = {
  results: SheetGrid;
  notFound: SheetGrid;
  stats: Record<string, unknown>;
};

export function opBulkLookup(
  grid: SheetGrid,
  op: Extract<PlanOp, { op: "bulk_lookup" }>,
): BulkLookupResult {
  const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();
  const candidateIdxs: number[] = [];
  grid.headers.forEach((h, i) => {
    const n = norm(h);
    if (KEYISH_HINTS.some((hint) => n.includes(hint))) candidateIdxs.push(i);
  });
  if (candidateIdxs.length === 0) candidateIdxs.push(...grid.headers.map((_, i) => i));

  const queries = op.queries.map((q) => q.trim()).filter(Boolean);
  const resultsHeader = [...grid.headers, "Matched On", "Matched Query"];
  const resultRows: CellValue[][] = [];
  const notFound: string[] = [];
  const seenRows = new Set<number>();

  for (const q of queries) {
    const nq = norm(q);
    let matched = false;
    for (let ri = 0; ri < grid.rows.length; ri++) {
      const row = grid.rows[ri];
      for (const ci of candidateIdxs) {
        const cell = norm(row[ci]);
        if (!cell) continue;
        const isNameCol = /name/.test(norm(grid.headers[ci]));
        const hit = isNameCol ? cell.includes(nq) : cell === nq;
        if (hit) {
          if (!seenRows.has(ri)) {
            seenRows.add(ri);
            resultRows.push([...row, grid.headers[ci], q]);
          }
          matched = true;
          break;
        }
      }
    }
    if (!matched) notFound.push(q);
  }

  return {
    results: { headers: resultsHeader, rows: resultRows },
    notFound: { headers: ["Query"], rows: notFound.map((q) => [q]) },
    stats: {
      queries: queries.length,
      matched: queries.length - notFound.length,
      notFound: notFound.length,
      rowsReturned: resultRows.length,
    },
  };
}

registerOp<Extract<PlanOp, { op: "bulk_lookup" }>>("bulk_lookup", (op, ctx) => {
  const targetGrid = ctx.grids[op.fileIndex];
  if (!targetGrid) {
    ctx.warnings.push(`bulk_lookup skipped: fileIndex ${op.fileIndex} out of range`);
    ctx.opLogs.push({ op: "bulk_lookup", status: "skipped", reason: "file index out of range" });
    return;
  }
  const r = opBulkLookup(targetGrid, op);
  const projectedResults = applyProjection(ctx, "Results", r.results, {
    alwaysKeep: ["Matched On", "Matched Query"],
  });
  writeGridToSheet(ctx.outWb, "Results", projectedResults);
  // "Not Found" is a diagnostic query list — skip projection.
  writeGridToSheet(ctx.outWb, "Not Found", r.notFound);
  ctx.state.producedSheets += 2;
  ctx.opLogs.push({
    op: "bulk_lookup",
    status: "ok",
    ms: Date.now() - ctx.started,
    file: ctx.files[op.fileIndex]?.name,
    ...r.stats,
  });
});
