// All-in-one server-only Excel engine.
// Reads .xlsx via ExcelJS, executes deterministic ops from a validated Plan,
// recalculates supported formulas via HyperFormula, writes back to .xlsx.

import ExcelJS from "exceljs";
import { HyperFormula } from "hyperformula";
import type { Plan, PlanOp, SheetMeta } from "./types";

export type CellValue = string | number | boolean | null;
export type SheetGrid = { headers: string[]; rows: CellValue[][] };

// ---------- Read ----------

export async function readWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

export function extractSheetMeta(wb: ExcelJS.Workbook): SheetMeta {
  const sheets = wb.worksheets.map((ws) => {
    const headers: string[] = [];
    const first = ws.getRow(1);
    first.eachCell({ includeEmpty: false }, (cell, col) => {
      headers[col - 1] = String(cell.value ?? "").trim();
    });
    const sampleRows: CellValue[][] = [];
    for (let r = 2; r <= Math.min(4, ws.rowCount); r++) {
      const row = ws.getRow(r);
      const out: CellValue[] = [];
      for (let c = 1; c <= headers.length; c++) {
        out.push(cellToValue(row.getCell(c).value));
      }
      sampleRows.push(out);
    }
    return {
      name: ws.name,
      rowCount: ws.rowCount,
      columnCount: ws.columnCount,
      headers: headers.filter(Boolean),
      sampleRows,
    };
  });
  return { sheets };
}

function cellToValue(v: ExcelJS.CellValue): CellValue {
  if (v == null) return null;
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    if ("text" in v && typeof v.text === "string") return v.text;
    if ("richText" in v && Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("");
    if ("result" in v) return cellToValue(v.result as ExcelJS.CellValue);
    if ("formula" in v) return null;
  }
  return String(v);
}

export function sheetToGrid(ws: ExcelJS.Worksheet): SheetGrid {
  const headers: string[] = [];
  const first = ws.getRow(1);
  const maxCol = ws.columnCount;
  for (let c = 1; c <= maxCol; c++) {
    headers.push(String(first.getCell(c).value ?? "").trim());
  }
  const rows: CellValue[][] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const out: CellValue[] = [];
    let any = false;
    for (let c = 1; c <= maxCol; c++) {
      const v = cellToValue(row.getCell(c).value);
      if (v !== null && v !== "") any = true;
      out.push(v);
    }
    if (any) rows.push(out);
  }
  return { headers, rows };
}

// ---------- Ops ----------

const HIGHLIGHT_UNMATCHED = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFE8B0" } };
const HIGHLIGHT_DUP = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFCCCC" } };
const HEADER_STYLE = {
  font: { bold: true, color: { argb: "FFFFFFFF" } },
  fill: { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF1F4033" } },
  alignment: { vertical: "middle" as const, horizontal: "left" as const },
};

function findColumnIndex(headers: string[], name: string): number {
  const norm = (s: string) => s.trim().toLowerCase().replace(/[_\-]+/g, " ").replace(/\s+/g, " ");
  const target = norm(name);
  // 1. exact normalized match
  let idx = headers.findIndex((h) => norm(h) === target);
  if (idx >= 0) return idx;
  // 2. substring either direction (e.g. "Name" ↔ "First Year Student Name")
  idx = headers.findIndex((h) => {
    const n = norm(h);
    return n.includes(target) || target.includes(n);
  });
  if (idx >= 0) return idx;
  // 3. token-subset match (all target words appear as whole words in header)
  const targetTokens = target.split(" ").filter(Boolean);
  if (targetTokens.length) {
    idx = headers.findIndex((h) => {
      const tokens = new Set(norm(h).split(" ").filter(Boolean));
      return targetTokens.every((t) => tokens.has(t));
    });
    if (idx >= 0) return idx;
  }
  return -1;
}

function styleHeader(ws: ExcelJS.Worksheet) {
  const row = ws.getRow(1);
  row.eachCell((cell) => {
    cell.font = HEADER_STYLE.font;
    cell.fill = HEADER_STYLE.fill;
    cell.alignment = HEADER_STYLE.alignment;
  });
  row.height = 22;
  row.commit();
}

function autoWidth(ws: ExcelJS.Worksheet, headers: string[], rows: CellValue[][]) {
  headers.forEach((h, i) => {
    const col = ws.getColumn(i + 1);
    const dataMax = rows.reduce((m, r) => Math.max(m, String(r[i] ?? "").length), 0);
    col.width = Math.min(Math.max(h.length + 2, dataMax + 2, 12), 42);
  });
}

// Merge multiple grids by key column.
export function opMerge(
  grids: SheetGrid[],
  fileNames: string[],
  op: Extract<PlanOp, { op: "merge" }>,
): { headers: string[]; rows: CellValue[][]; unmatchedRowSet: Set<number>; stats: Record<string, number> } {
  if (grids.length === 0) throw new Error("No input grids");

  const perFileKeyIdx = grids.map((g) => findColumnIndex(g.headers, op.keyColumn));
  perFileKeyIdx.forEach((idx, i) => {
    if (idx < 0) throw new Error(`Key column "${op.keyColumn}" not found in file "${fileNames[i]}"`);
  });

  // Canonical header list: keyColumn first, then union of other headers with source tag.
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

  // Merge rows keyed by keyColumn value.
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
    op.strategy === "intersection"
      ? allKeys.filter((k) => seenPerFile.every((s) => s.has(k)))
      : allKeys;

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

export function opDedupe(
  grid: SheetGrid,
  op: Extract<PlanOp, { op: "dedupe" }>,
): { grid: SheetGrid; removed: number } {
  const seen = new Set<string>();
  const kept: CellValue[][] = [];
  let removed = 0;
  const keyIdx = op.keyColumn ? findColumnIndex(grid.headers, op.keyColumn) : -1;
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
  };
};

export function opDiff(
  a: SheetGrid,
  b: SheetGrid,
  op: Extract<PlanOp, { op: "diff" }>,
  namesA: string,
  namesB: string,
): DiffResult {
  const kA = findColumnIndex(a.headers, op.keyColumn);
  const kB = findColumnIndex(b.headers, op.keyColumn);
  if (kA < 0) throw new Error(`Key column "${op.keyColumn}" not found in ${namesA}`);
  if (kB < 0) throw new Error(`Key column "${op.keyColumn}" not found in ${namesB}`);

  const mapA = new Map<string, CellValue[]>();
  for (const r of a.rows) {
    const k = String(r[kA] ?? "").trim();
    if (k) mapA.set(k, r);
  }
  const mapB = new Map<string, CellValue[]>();
  for (const r of b.rows) {
    const k = String(r[kB] ?? "").trim();
    if (k) mapB.set(k, r);
  }

  const missingInA: CellValue[][] = [];
  const missingInB: CellValue[][] = [];
  const changed: CellValue[][] = [];

  for (const [key, rowB] of mapB) {
    if (!mapA.has(key)) missingInA.push(rowB);
    else {
      const rowA = mapA.get(key)!;
      if (JSON.stringify(rowA) !== JSON.stringify(rowB)) {
        changed.push([key, JSON.stringify(rowA), JSON.stringify(rowB)]);
      }
    }
  }
  for (const [key, rowA] of mapA) {
    if (!mapB.has(key)) missingInB.push(rowA);
  }

  return {
    missingInA: { headers: b.headers, rows: missingInA },
    missingInB: { headers: a.headers, rows: missingInB },
    changed: {
      headers: [op.keyColumn, `Values in ${namesA}`, `Values in ${namesB}`],
      rows: changed,
    },
    stats: {
      comparedA: mapA.size,
      comparedB: mapB.size,
      missingInA: missingInA.length,
      missingInB: missingInB.length,
      changed: changed.length,
    },
  };
}

export function opSummary(grid: SheetGrid): { headers: string[]; rows: CellValue[][] } {
  const rows: CellValue[][] = [
    ["Metric", "Value"],
    ["Total rows", grid.rows.length],
    ["Total columns", grid.headers.length],
  ];
  // Numeric column stats
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

// Recalc formulas via HyperFormula. Preserves original formula strings when unsupported.
export function recalcFormulas(wb: ExcelJS.Workbook): { recalculated: number; skipped: number } {
  const sheetsData: Record<string, (string | number | boolean | null)[][]> = {};
  wb.worksheets.forEach((ws) => {
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const out: (string | number | boolean | null)[] = [];
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = row.getCell(c);
        const v = cell.value;
        if (v && typeof v === "object" && "formula" in v) {
          out.push("=" + (v as ExcelJS.CellFormulaValue).formula);
        } else {
          out.push(cellToValue(v) as string | number | boolean | null);
        }
      }
      data.push(out);
    }
    sheetsData[ws.name] = data;
  });

  let recalculated = 0;
  let skipped = 0;
  try {
    const hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: "gpl-v3" });
    wb.worksheets.forEach((ws, si) => {
      const sheetId = hf.getSheetId(ws.name);
      if (sheetId == null) return;
      for (let r = 1; r <= ws.rowCount; r++) {
        for (let c = 1; c <= ws.columnCount; c++) {
          const cell = ws.getRow(r).getCell(c);
          const v = cell.value;
          if (v && typeof v === "object" && "formula" in v) {
            try {
              const result = hf.getCellValue({ sheet: sheetId, row: r - 1, col: c - 1 });
              if (result != null && typeof result !== "object") {
                cell.value = { formula: (v as ExcelJS.CellFormulaValue).formula, result: result as ExcelJS.CellFormulaValue["result"] } as ExcelJS.CellFormulaValue;
                recalculated++;
              } else {
                skipped++;
              }
            } catch {
              skipped++;
            }
          }
        }
      }
      void si;
    });
    hf.destroy();
  } catch (e) {
    console.warn("HyperFormula init failed", e);
    skipped++;
  }
  return { recalculated, skipped };
}

// ---------- Write ----------

export function writeGridToSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  grid: SheetGrid,
  opts: { highlightRows?: Set<number>; headerStyle?: boolean } = {},
): ExcelJS.Worksheet {
  const existing = wb.getWorksheet(sheetName);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  ws.addRow(grid.headers);
  grid.rows.forEach((r) => ws.addRow(r));
  if (opts.headerStyle !== false) styleHeader(ws);
  autoWidth(ws, grid.headers, grid.rows);
  if (opts.highlightRows) {
    opts.highlightRows.forEach((rowIdx) => {
      const row = ws.getRow(rowIdx + 2); // +1 header, +1 1-indexed
      row.eachCell((cell) => {
        cell.fill = HIGHLIGHT_UNMATCHED;
      });
    });
  }
  return ws;
}

// ---------- Orchestrator ----------

export type EngineFile = { name: string; buffer: ArrayBuffer };

export type EngineResult = {
  buffer: Buffer;
  stats: Record<string, unknown>;
  warnings: string[];
};

export async function runPlan(files: EngineFile[], plan: Plan): Promise<EngineResult> {
  const warnings: string[] = [...plan.warnings];
  const opLogs: unknown[] = [];
  const stats: Record<string, unknown> = { ops: opLogs };

  const workbooks = await Promise.all(files.map((f) => readWorkbook(f.buffer)));
  const grids = workbooks.map((wb) => sheetToGrid(wb.worksheets[0]));

  // Determine output strategy.
  const hasMutating = plan.ops.some(
    (o) => o.op === "merge" || o.op === "dedupe" || o.op === "highlight_column",
  );
  // Merge/dedupe/highlight can preserve first-workbook styling by mutating a
  // copy of the first workbook. Diff/summary produce brand-new deliverables.
  const outWb = hasMutating ? workbooks[0] : new ExcelJS.Workbook();
  outWb.creator = "Ledgerly";
  outWb.created = new Date();

  let currentGrid: SheetGrid = grids[0];
  let unmatched: Set<number> | undefined;
  let producedSheets = 0;

  const startAll = Date.now();
  for (let i = 0; i < plan.ops.length; i++) {
    const op = plan.ops[i];
    const started = Date.now();
    const label = `op[${i}] ${op.op}`;
    console.info(`[engine] ${label} starting`);
    try {
      if (op.op === "merge") {
        if (grids.length < 2) {
          warnings.push("merge skipped: needs ≥2 files");
          opLogs.push({ op: "merge", status: "skipped", reason: "needs ≥2 files" });
          continue;
        }
        const merged = opMerge(grids, files.map((f) => f.name), op);
        currentGrid = { headers: merged.headers, rows: merged.rows };
        unmatched = merged.unmatchedRowSet;
        opLogs.push({
          op: "merge",
          status: "ok",
          ms: Date.now() - started,
          keyColumn: op.keyColumn,
          ...merged.stats,
        });
      } else if (op.op === "dedupe") {
        const { grid: dg, removed } = opDedupe(currentGrid, op);
        currentGrid = dg;
        opLogs.push({
          op: "dedupe",
          status: "ok",
          ms: Date.now() - started,
          strategy: op.strategy,
          keyColumn: op.keyColumn ?? null,
          rowsIn: currentGrid.rows.length + removed,
          rowsOut: currentGrid.rows.length,
          duplicatesRemoved: removed,
        });
      } else if (op.op === "diff") {
        const aIdx = op.fileAIndex;
        const bIdx = op.fileBIndex;
        const a = grids[aIdx];
        const b = grids[bIdx];
        if (!a || !b) {
          warnings.push(`diff skipped: fileAIndex=${aIdx} or fileBIndex=${bIdx} out of range`);
          opLogs.push({ op: "diff", status: "skipped", reason: "file index out of range" });
          continue;
        }
        if (aIdx === bIdx) {
          warnings.push("diff skipped: fileAIndex == fileBIndex");
          opLogs.push({ op: "diff", status: "skipped", reason: "same file on both sides" });
          continue;
        }
        const nameA = files[aIdx].name;
        const nameB = files[bIdx].name;
        const d = opDiff(a, b, op, nameA, nameB);
        const sheetMissingA = `Missing in ${safeSheetName(nameA)}`;
        const sheetMissingB = `Missing in ${safeSheetName(nameB)}`;
        writeGridToSheet(outWb, sheetMissingA, d.missingInA);
        writeGridToSheet(outWb, sheetMissingB, d.missingInB);
        if (d.changed.rows.length > 0) {
          writeGridToSheet(outWb, "Changed rows", d.changed);
          producedSheets++;
        }
        producedSheets += 2;
        opLogs.push({
          op: "diff",
          status: "ok",
          ms: Date.now() - started,
          keyColumn: op.keyColumn,
          fileA: nameA,
          fileB: nameB,
          ...d.stats,
          sheets: [sheetMissingA, sheetMissingB, ...(d.changed.rows.length ? ["Changed rows"] : [])],
        });
      } else if (op.op === "summary") {
        const s = opSummary(currentGrid);
        writeGridToSheet(outWb, "Summary", s);
        producedSheets++;
        opLogs.push({
          op: "summary",
          status: "ok",
          ms: Date.now() - started,
          metrics: s.rows.length,
          sheet: "Summary",
        });
      } else if (op.op === "highlight_column") {
        const idx = findColumnIndex(currentGrid.headers, op.column);
        if (idx < 0) {
          warnings.push(`highlight_column skipped: column "${op.column}" not found`);
          opLogs.push({ op: "highlight_column", status: "skipped", reason: `column "${op.column}" not found` });
          continue;
        }
        const highlightRows = new Set<number>();
        if (op.rule === "missing") {
          currentGrid.rows.forEach((r, ri) => {
            if (r[idx] == null || r[idx] === "") highlightRows.add(ri);
          });
        } else if (op.rule === "duplicate") {
          const counts = new Map<string, number>();
          currentGrid.rows.forEach((r) => {
            const k = String(r[idx] ?? "");
            counts.set(k, (counts.get(k) ?? 0) + 1);
          });
          currentGrid.rows.forEach((r, ri) => {
            if ((counts.get(String(r[idx] ?? "")) ?? 0) > 1) highlightRows.add(ri);
          });
        } else if (op.rule === "outlier") {
          const nums = currentGrid.rows.map((r) => Number(r[idx])).filter((n) => Number.isFinite(n));
          if (nums.length > 3) {
            const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
            const sd = Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / nums.length);
            currentGrid.rows.forEach((r, ri) => {
              const n = Number(r[idx]);
              if (Number.isFinite(n) && Math.abs(n - mean) > 2 * sd) highlightRows.add(ri);
            });
          }
        }
        unmatched = new Set([...(unmatched ?? []), ...highlightRows]);
        opLogs.push({
          op: "highlight_column",
          status: "ok",
          ms: Date.now() - started,
          column: op.column,
          rule: op.rule,
          matches: highlightRows.size,
        });
      } else if (op.op === "recalc") {
        const r = recalcFormulas(outWb);
        opLogs.push({ op: "recalc", status: "ok", ms: Date.now() - started, ...r });
      }
      console.info(`[engine] ${label} done in ${Date.now() - started}ms`);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      warnings.push(`${op.op} failed: ${msg}`);
      opLogs.push({ op: op.op, status: "failed", ms: Date.now() - started, error: msg });
      console.error(`[engine] ${label} failed: ${msg}`);
    }
  }

  if (hasMutating) {
    writeGridToSheet(outWb, "Result", currentGrid, { highlightRows: unmatched });
    producedSheets++;
    const resultWs = outWb.getWorksheet("Result");
    if (resultWs) {
      outWb.worksheets.splice(outWb.worksheets.indexOf(resultWs), 1);
      outWb.worksheets.unshift(resultWs);
    }
  }

  if (producedSheets === 0) {
    // No op produced output; add an explicit info sheet instead of silently
    // exporting the original file.
    writeGridToSheet(outWb, "No output", {
      headers: ["Notice"],
      rows: [
        ["No operation produced output. See warnings for details."],
        ...warnings.map((w) => [w] as CellValue[]),
      ],
    });
    warnings.push("No operation produced output.");
  }

  if (hasMutating && !plan.ops.some((o) => o.op === "recalc")) {
    const r = recalcFormulas(outWb);
    opLogs.push({ op: "recalc_auto", status: "ok", ...r });
  }

  stats.totalMs = Date.now() - startAll;
  stats.producedSheets = producedSheets;

  const out = await outWb.xlsx.writeBuffer();
  return { buffer: Buffer.from(out), stats, warnings };
}

function safeSheetName(name: string): string {
  // Excel sheet names: max 31 chars, no : \ / ? * [ ]
  return name.replace(/\.[^.]+$/, "").replace(/[\\/:?*[\]]/g, " ").slice(0, 25);
}
