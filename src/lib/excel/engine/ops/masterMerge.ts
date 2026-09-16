import type { PlanOp } from "../../types";
import { resolveKeyAcrossFiles, resolveKeyColumn } from "../shared/key-resolution";
import type { CellValue, SheetGrid } from "../shared/workbook";
import { safeSheetName, writeGridToSheet } from "../shared/workbook";
import { applyProjection, type OpHandler } from "../registry";

export type MasterMergeResult = {
  master: SheetGrid;
  missingPerFile: { fileName: string; keys: string[] }[];
  summary: SheetGrid;
  stats: Record<string, unknown>;
};

export function opMasterMerge(
  grids: SheetGrid[],
  fileNames: string[],
  op: Extract<PlanOp, { op: "master_merge" }>,
): MasterMergeResult {
  if (grids.length === 0) throw new Error("No input grids");
  // append ignores the key entirely, so tolerate an unresolvable key there.
  const keyIdxs =
    op.joinType === "append"
      ? grids.map((g) => {
          const r = resolveKeyColumn(g.headers, op.keyColumn);
          return r.status === "resolved" ? r.index : -1;
        })
      : resolveKeyAcrossFiles(
          grids.map((g) => g.headers),
          op.keyColumn,
          fileNames,
        ).map((r) => r.index);

  if (op.joinType === "append") {
    const headerSet: string[] = ["__Source"];
    const seen = new Set<string>(["__source"]);
    grids.forEach((g) =>
      g.headers.forEach((h) => {
        const k = h.trim().toLowerCase();
        if (!seen.has(k) && h) {
          seen.add(k);
          headerSet.push(h);
        }
      }),
    );
    const rows: CellValue[][] = [];
    grids.forEach((g, fi) => {
      g.rows.forEach((r) => {
        const out: CellValue[] = new Array(headerSet.length).fill(null);
        out[0] = fileNames[fi];
        g.headers.forEach((h, ci) => {
          const targetIdx = headerSet.findIndex(
            (x) => x.trim().toLowerCase() === h.trim().toLowerCase(),
          );
          if (targetIdx > 0) out[targetIdx] = r[ci];
        });
        rows.push(out);
      });
    });
    const summary: SheetGrid = {
      headers: ["Metric", "Value"],
      rows: [
        ["Join type", "append"],
        ["Files", grids.length],
        ["Total rows", rows.length],
      ],
    };
    return {
      master: { headers: headerSet, rows },
      missingPerFile: [],
      summary,
      stats: { joinType: "append", files: grids.length, rows: rows.length },
    };
  }

  const headers: string[] = [op.keyColumn];
  type Col = { sources: { fileIdx: number; sourceHeader: string }[] };
  const columnMeta: (Col | null)[] = [null];
  const usedNames = new Set<string>([op.keyColumn.toLowerCase()]);

  grids.forEach((g, fi) => {
    g.headers.forEach((h, ci) => {
      if (ci === keyIdxs[fi] || !h) return;
      const existingIndex = headers.findIndex(
        (x) => x.trim().toLowerCase() === h.trim().toLowerCase(),
      );
      if (existingIndex >= 0) {
        columnMeta[existingIndex]!.sources.push({ fileIdx: fi, sourceHeader: h });
      } else {
        usedNames.add(h.toLowerCase());
        headers.push(h);
        columnMeta.push({ sources: [{ fileIdx: fi, sourceHeader: h }] });
      }
    });
  });

  const byKey = new Map<string, { row: CellValue[]; seenIn: Set<number>; dupCount: number }>();
  const keyOrder: string[] = [];
  const seenPerFile: Set<string>[] = grids.map(() => new Set());

  grids.forEach((g, fi) => {
    const kIdx = keyIdxs[fi];
    g.rows.forEach((r) => {
      const rawKey = r[kIdx];
      if (rawKey == null || rawKey === "") return;
      const key = String(rawKey).trim();
      if (!key) return;
      seenPerFile[fi].add(key);
      let entry = byKey.get(key);
      if (!entry) {
        entry = { row: new Array(headers.length).fill(null), seenIn: new Set(), dupCount: 0 };
        entry.row[0] = rawKey;
        byKey.set(key, entry);
        keyOrder.push(key);
      }
      if (entry.seenIn.has(fi)) entry.dupCount++;
      entry.seenIn.add(fi);
      for (let ti = 1; ti < headers.length; ti++) {
        const cm = columnMeta[ti];
        if (!cm) continue;
        const source = cm.sources.find((s) => s.fileIdx === fi);
        if (!source) continue;
        const srcIdx = g.headers.indexOf(source.sourceHeader);
        if (srcIdx < 0) continue;
        const newVal = r[srcIdx];
        const cur = entry.row[ti];
        if (cur == null || cur === "") {
          entry.row[ti] = newVal;
        } else if (op.dupeStrategy === "latest") {
          entry.row[ti] = newVal;
        } else if (
          op.dupeStrategy === "merge" &&
          newVal != null &&
          newVal !== "" &&
          String(newVal) !== String(cur)
        ) {
          entry.row[ti] = `${String(cur)} | ${String(newVal)}`;
        }
      }
    });
  });

  const kept = keyOrder.filter((k) => {
    const e = byKey.get(k)!;
    if (op.joinType === "inner") return e.seenIn.size === grids.length;
    if (op.joinType === "left") return e.seenIn.has(0);
    return true;
  });

  const rows = kept.map((k) => byKey.get(k)!.row);

  const missingPerFile = grids.map((_, fi) => {
    const keys = kept.filter((k) => !byKey.get(k)!.seenIn.has(fi));
    return { fileName: fileNames[fi], keys };
  });

  const totalDup = Array.from(byKey.values()).reduce((s, e) => s + e.dupCount, 0);

  const summary: SheetGrid = {
    headers: ["Metric", "Value"],
    rows: [
      ["Key column", op.keyColumn],
      ["Join type", op.joinType],
      ["Duplicate strategy", op.dupeStrategy],
      ["Files merged", grids.length],
      ["Unique keys", keyOrder.length],
      ["Rows in master", rows.length],
      ["Duplicate key values across files", totalDup],
      ...missingPerFile.map((m) => [`Missing in ${m.fileName}`, m.keys.length] as CellValue[]),
    ],
  };

  return {
    master: { headers, rows },
    missingPerFile,
    summary,
    stats: {
      joinType: op.joinType,
      dupeStrategy: op.dupeStrategy,
      files: grids.length,
      uniqueKeys: keyOrder.length,
      rows: rows.length,
      duplicates: totalDup,
    },
  };
}

export const masterMergeHandler: OpHandler<Extract<PlanOp, { op: "master_merge" }>> = (op, ctx) => {
  if (ctx.grids.length < 2) {
    ctx.warnings.push("master_merge skipped: needs ≥2 files");
    ctx.opLogs.push({ op: "master_merge", status: "skipped", reason: "needs ≥2 files" });
    return;
  }
  const m = opMasterMerge(
    ctx.grids,
    ctx.files.map((f) => f.name),
    op,
  );
  const projectedMaster = applyProjection(ctx, "Master", m.master, { alwaysKeep: [op.keyColumn] });
  writeGridToSheet(ctx.outWb, "Master", projectedMaster);
  ctx.state.producedSheets++;
  // "Merge Summary" and "Missing in <file>" are diagnostic — skip projection.
  writeGridToSheet(ctx.outWb, "Merge Summary", m.summary);
  ctx.state.producedSheets++;
  for (const mp of m.missingPerFile) {
    if (mp.keys.length === 0) continue;
    const sheetName = safeSheetName(`Missing in ${mp.fileName}`);
    writeGridToSheet(ctx.outWb, sheetName, {
      headers: [op.keyColumn],
      rows: mp.keys.map((k) => [k]),
    });
    ctx.state.producedSheets++;
  }
  ctx.opLogs.push({
    op: "master_merge",
    status: "ok",
    ms: Date.now() - ctx.started,
    keyColumn: op.keyColumn,
    ...m.stats,
  });
};
