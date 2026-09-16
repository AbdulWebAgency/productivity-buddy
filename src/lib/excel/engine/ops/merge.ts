import type { PlanOp } from "../../types";
import { resolveKeyAcrossFiles } from "../shared/key-resolution";
import type { CellValue, SheetGrid } from "../shared/workbook";
import type { OpHandler } from "../registry";

export function opMerge(
  grids: SheetGrid[],
  fileNames: string[],
  op: Extract<PlanOp, { op: "merge" }>,
): {
  headers: string[];
  rows: CellValue[][];
  unmatchedRowSet: Set<number>;
  stats: Record<string, number>;
} {
  if (grids.length === 0) throw new Error("No input grids");

  const resolvedKeys = resolveKeyAcrossFiles(
    grids.map((g) => g.headers),
    op.keyColumn,
    fileNames,
  );
  const perFileKeyIdx = resolvedKeys.map((r) => r.index);

  const canonicalHeaders: string[] = [op.keyColumn];
  const columnMap: Map<string, number> = new Map();
  columnMap.set(op.keyColumn.toLowerCase(), 0);

  grids.forEach((g, fi) => {
    g.headers.forEach((h, ci) => {
      if (ci === perFileKeyIdx[fi]) return;
      const key = `${h}::${fi}`.toLowerCase();
      if (!columnMap.has(key)) {
        columnMap.set(key, canonicalHeaders.length);
        canonicalHeaders.push(`${h} (${fileNames[fi]})`);
      }
    });
  });
  canonicalHeaders.push("Sources");

  const byKey = new Map<string, { row: CellValue[]; sources: Set<string> }>();
  const seenPerFile: Set<string>[] = grids.map(() => new Set());

  grids.forEach((g, fi) => {
    const keyIdx = perFileKeyIdx[fi];
    g.rows.forEach((r) => {
      const keyVal = r[keyIdx];
      if (keyVal == null || keyVal === "") return;
      const key = String(keyVal);
      seenPerFile[fi].add(key);
      let entry = byKey.get(key);
      if (!entry) {
        entry = { row: new Array(canonicalHeaders.length).fill(null), sources: new Set() };
        entry.row[0] = keyVal;
        byKey.set(key, entry);
      }
      entry.sources.add(fileNames[fi]);
      g.headers.forEach((h, ci) => {
        if (ci === keyIdx) return;
        const targetIdx = columnMap.get(`${h}::${fi}`.toLowerCase())!;
        entry!.row[targetIdx] = r[ci];
      });
    });
  });

  const allKeys = Array.from(byKey.keys());
  const filtered =
    op.strategy === "intersection" ? allKeys.filter((k) => seenPerFile.every((s) => s.has(k))) : allKeys;

  const rows: CellValue[][] = [];
  const unmatched = new Set<number>();
  filtered.forEach((k, idx) => {
    const e = byKey.get(k)!;
    e.row[canonicalHeaders.length - 1] = Array.from(e.sources).join(", ");
    rows.push(e.row);
    if (op.highlightUnmatched && e.sources.size < grids.length) unmatched.add(idx);
  });

  return {
    headers: canonicalHeaders,
    rows,
    unmatchedRowSet: unmatched,
    stats: {
      inputFiles: grids.length,
      totalKeys: allKeys.length,
      merged: rows.length,
      unmatched: unmatched.size,
    },
  };
}

export const mergeHandler: OpHandler<Extract<PlanOp, { op: "merge" }>> = (op, ctx) => {
  if (ctx.grids.length < 2) {
    ctx.warnings.push("merge skipped: needs ≥2 files");
    ctx.opLogs.push({ op: "merge", status: "skipped", reason: "needs ≥2 files" });
    return;
  }
  const merged = opMerge(
    ctx.grids,
    ctx.files.map((f) => f.name),
    op,
  );
  ctx.state.currentGrid = { headers: merged.headers, rows: merged.rows };
  ctx.state.unmatched = merged.unmatchedRowSet;
  ctx.opLogs.push({
    op: "merge",
    status: "ok",
    ms: Date.now() - ctx.started,
    keyColumn: op.keyColumn,
    ...merged.stats,
  });
};
