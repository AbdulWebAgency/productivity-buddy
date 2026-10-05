import type { PlanOp } from "../../types";
import { resolveKeyAcrossFiles } from "../shared/key-resolution";
import type { CellValue, SheetGrid } from "../shared/workbook";
import { safeSheetName, writeGridToSheet } from "../shared/workbook";
import { applyProjection, type OpHandler } from "../registry";
import { normalizeHeader } from "../shared/headers";

export type DiffResult = {
  missingInA: SheetGrid; // rows present in B but not A (add to A)
  missingInB: SheetGrid; // rows present in A but not B
  changed: SheetGrid;
  stats: {
    comparedA: number;
    comparedB: number;
    missingInA: number;
    missingInB: number;
    changed: number;
    resolvedKeyA?: string;
    resolvedKeyB?: string;
    emptyKeysA?: number;
    emptyKeysB?: number;
  };
};

export function opDiff(
  a: SheetGrid,
  b: SheetGrid,
  op: Extract<PlanOp, { op: "diff" }>,
  namesA: string,
  namesB: string,
): DiffResult {
  const [rA, rB] = resolveKeyAcrossFiles([a.headers, b.headers], op.keyColumn, [namesA, namesB]);
  const kA = rA.index;
  const kB = rB.index;

  const resolvedA = rA.header;
  const resolvedB = rB.header;

  const mapA = new Map<string, CellValue[]>();
  let emptyKeysA = 0;
  for (const r of a.rows) {
    const k = String(r[kA] ?? "").trim();
    if (k) mapA.set(k.toLowerCase(), r);
    else emptyKeysA++;
  }
  const mapB = new Map<string, CellValue[]>();
  let emptyKeysB = 0;
  for (const r of b.rows) {
    const k = String(r[kB] ?? "").trim();
    if (k) mapB.set(k.toLowerCase(), r);
    else emptyKeysB++;
  }

  if (mapA.size === 0) {
    throw new Error(
      `Column "${op.keyColumn}" resolved to "${resolvedA}" in ${namesA} but has no usable values. Pick a different key column.`,
    );
  }
  if (mapB.size === 0) {
    throw new Error(
      `Column "${op.keyColumn}" resolved to "${resolvedB}" in ${namesB} but has no usable values. Pick a different key column.`,
    );
  }

  const missingInA: CellValue[][] = [];
  const missingInB: CellValue[][] = [];
  const changed: CellValue[][] = [];

  const bHeaderIndexByNormalized = new Map<string, number>();

  for (let i = 0; i < b.headers.length; i++) {
    const normalized = normalizeHeader(b.headers[i]);
    if (!bHeaderIndexByNormalized.has(normalized)) {
      bHeaderIndexByNormalized.set(normalized, i);
    }
  }

  for (const [key, rowB] of mapB) {
    if (!mapA.has(key)) missingInA.push(rowB);
    else {
      const rowA = mapA.get(key)!;
      const changedColumns: string[] = [];
      const changedDetails: string[] = [];

      for (let i = 0; i < a.headers.length; i++) {
        if (i === kA) continue;

        const normalizedHeader = normalizeHeader(a.headers[i]);
        const j = bHeaderIndexByNormalized.get(normalizedHeader);

        if (j === undefined) continue;

        const valueA = String(rowA[i] ?? "").trim();
        const valueB = String(rowB[j] ?? "").trim();

        if (valueA !== valueB) {
          changedColumns.push(a.headers[i]);
          changedDetails.push(`${a.headers[i]}: ${valueA || "(blank)"} → ${valueB || "(blank)"}`);
        }
      }

      if (changedColumns.length > 0) {
        // Show the key as the user wrote it (the map key is lowercased for matching only).
        const displayKey = String(rowB[kB] ?? "").trim();
        changed.push([displayKey, changedColumns.join(", "), changedDetails.join("; ")]);
      }
    }
  }
  for (const [key, rowA] of mapA) {
    if (!mapB.has(key)) missingInB.push(rowA);
  }

  return {
    missingInA: { headers: b.headers, rows: missingInA },
    missingInB: { headers: a.headers, rows: missingInB },
    changed: { headers: [op.keyColumn, "Changed Columns", `Values (${namesA} → ${namesB})`], rows: changed },
    stats: {
      comparedA: mapA.size,
      comparedB: mapB.size,
      missingInA: missingInA.length,
      missingInB: missingInB.length,
      changed: changed.length,
      resolvedKeyA: resolvedA,
      resolvedKeyB: resolvedB,
      emptyKeysA,
      emptyKeysB,
    },
  };
}

export const diffHandler: OpHandler<Extract<PlanOp, { op: "diff" }>> = (op, ctx) => {
  const aIdx = op.fileAIndex;
  const bIdx = op.fileBIndex;
  const a = ctx.grids[aIdx];
  const b = ctx.grids[bIdx];
  if (!a || !b) {
    ctx.warnings.push(`diff skipped: fileAIndex=${aIdx} or fileBIndex=${bIdx} out of range`);
    ctx.opLogs.push({ op: "diff", status: "skipped", reason: "file index out of range" });
    return;
  }
  if (aIdx === bIdx) {
    ctx.warnings.push("diff skipped: fileAIndex == fileBIndex");
    ctx.opLogs.push({ op: "diff", status: "skipped", reason: "same file on both sides" });
    return;
  }
  const nameA = ctx.files[aIdx].name;
  const nameB = ctx.files[bIdx].name;
  const d = opDiff(a, b, op, nameA, nameB);
  const sheetOnlyInB = `Only in ${safeSheetName(nameB)}`; // rows from B whose key isn't in A
  const sheetOnlyInA = `Only in ${safeSheetName(nameA)}`; // rows from A whose key isn't in B
  if (d.missingInA.rows.length > 0) {
    const projected = applyProjection(ctx, sheetOnlyInB, d.missingInA, { alwaysKeep: [op.keyColumn] });
    writeGridToSheet(ctx.outWb, sheetOnlyInB, projected);
    ctx.state.producedSheets++;
  }
  if (d.missingInB.rows.length > 0) {
    const projected = applyProjection(ctx, sheetOnlyInA, d.missingInB, { alwaysKeep: [op.keyColumn] });
    writeGridToSheet(ctx.outWb, sheetOnlyInA, projected);
    ctx.state.producedSheets++;
  }
  if (d.changed.rows.length > 0) {
    // Diagnostic sheet — always [key, "Changed Columns", values], skip projection.
    writeGridToSheet(ctx.outWb, "Changed rows", d.changed);
    ctx.state.producedSheets++;
  }
  if (d.stats.missingInA === 0 && d.stats.missingInB === 0 && d.stats.changed === 0) {
    ctx.warnings.push(
      `${nameA} and ${nameB} have identical "${op.keyColumn}" values (${d.stats.comparedA} rows) — nothing missing on either side.`,
    );
  }
  if (d.stats.emptyKeysA && d.stats.emptyKeysA > 0) {
    ctx.warnings.push(`${nameA}: ${d.stats.emptyKeysA} rows had a blank "${op.keyColumn}" and were skipped.`);
  }
  if (d.stats.emptyKeysB && d.stats.emptyKeysB > 0) {
    ctx.warnings.push(`${nameB}: ${d.stats.emptyKeysB} rows had a blank "${op.keyColumn}" and were skipped.`);
  }
  ctx.opLogs.push({
    op: "diff",
    status: "ok",
    ms: Date.now() - ctx.started,
    keyColumn: op.keyColumn,
    fileA: nameA,
    fileB: nameB,
    ...d.stats,
    onlyInA: d.stats.missingInB,
    onlyInB: d.stats.missingInA,
    sheets: [sheetOnlyInB, sheetOnlyInA, ...(d.changed.rows.length ? ["Changed rows"] : [])],
  });
};
