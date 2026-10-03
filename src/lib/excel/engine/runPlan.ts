// Plan orchestrator. Reads inputs, dispatches ops through the registry,
// finalises the workbook, returns bytes + stats.
//
// Phase 2A: workbook I/O goes through the SheetJS compatibility layer.
// Formula recalculation still uses the legacy ExcelJS implementation via a
// serialize -> legacy load -> recalc -> serialize bridge (Phase 2B migrates it).
import type ExcelJS from "exceljs";
import type { Plan, PlanOp } from "../types";
import type { EngineFile, EngineResult } from "./types";
import {
  cloneWorkbook,
  createWorkbook,
  readLegacyWorkbook,
  readWorkbook,
  sheetToGrid,
  writeGridToSheet,
  writeWorkbook,
  type WorkbookHandle,
} from "./shared/workbook";
import type { CellValue } from "./shared/workbook";
import { applyProjection, type EngineState, type OpCtx } from "./registry";
import { recalcFormulas } from "./ops/highlight";
import { getOpHandler } from "./ops";

/** Legacy recalc bridge. Returns the (possibly) recalculated workbook bytes. */
async function legacyRecalc(
  bytes: Buffer,
): Promise<{ bytes: Buffer; stats: { recalculated: number; skipped: number } | null; error?: string }> {
  try {
    const legacy = await readLegacyWorkbook(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    const stats = recalcFormulas(legacy);
    const out = await legacy.xlsx.writeBuffer();
    return { bytes: Buffer.from(out), stats };
  } catch (e) {
    return { bytes, stats: null, error: e instanceof Error ? e.message : String(e) };
  }
}

export async function runPlan(files: EngineFile[], plan: Plan): Promise<EngineResult> {
  const warnings: string[] = [...plan.warnings];
  const opLogs: Record<string, unknown>[] = [];
  const stats: Record<string, unknown> = { ops: opLogs };

  const workbooks = await Promise.all(files.map((f) => readWorkbook(f.buffer)));
  const grids = workbooks.map((wb) => sheetToGrid(wb.worksheets[0]));

  // Determine output strategy.
  const hasMutating = plan.ops.some((o) => o.op === "merge" || o.op === "dedupe" || o.op === "highlight_column");
  // Merge/dedupe/highlight keep the first workbook's sheets by mutating an
  // independent copy of it. Diff/summary produce brand-new deliverables.
  const outWb: WorkbookHandle = hasMutating ? cloneWorkbook(workbooks[0]) : createWorkbook();
  outWb.creator = "Productivity Buddy";
  outWb.created = new Date();

  const state: EngineState = {
    currentGrid: grids[0],
    unmatched: undefined,
    producedSheets: 0,
  };

  const ctxBase = {
    files,
    grids,
    // OpCtx (registry.ts, not migrated) still types outWb as ExcelJS. Ops only
    // pass it to writeGridToSheet, which accepts both shapes.
    outWb: outWb as unknown as ExcelJS.Workbook,
    state,
    warnings,
    opLogs,
    projection: plan.projection ?? null,
    projectionEvents: [] as OpCtx["projectionEvents"],
  };

  const startAll = Date.now();
  for (let i = 0; i < plan.ops.length; i++) {
    const op = plan.ops[i];
    const started = Date.now();
    const label = `op[${i}] ${op.op}`;
    console.info(`[engine] ${label} starting`);
    try {
      const handler = getOpHandler(op.op);
      if (!handler) {
        warnings.push(`${op.op}: no handler registered`);
        opLogs.push({ op: op.op, status: "skipped", reason: "no handler registered" });
        continue;
      }
      const ctx: OpCtx = { ...ctxBase, started };
      await handler(op, ctx);
      console.info(`[engine] ${label} done in ${Date.now() - started}ms`);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      warnings.push(`${op.op} failed: ${msg}`);
      opLogs.push({ op: op.op, status: "failed", ms: Date.now() - started, error: msg });
      console.error(`[engine] ${label} failed: ${msg}`);
    }
  }

  if (hasMutating) {
    const alwaysKeep: string[] = [];
    for (const o of plan.ops) {
      if (o.op === "merge") {
        alwaysKeep.push((o as Extract<PlanOp, { op: "merge" }>).keyColumn, "Sources");
      } else if (o.op === "dedupe") {
        const k = (o as Extract<PlanOp, { op: "dedupe" }>).keyColumn;
        if (k) alwaysKeep.push(k);
      } else if (o.op === "highlight_column") {
        alwaysKeep.push((o as Extract<PlanOp, { op: "highlight_column" }>).column);
      }
    }
    const finalCtx: OpCtx = { ...ctxBase, started: Date.now() };
    const projectedResult = applyProjection(finalCtx, "Result", state.currentGrid, { alwaysKeep });
    writeGridToSheet(outWb, "Result", projectedResult, { highlightRows: state.unmatched });
    state.producedSheets++;
    // Result is always the first sheet.
    outWb.moveWorksheet("Result", 0);
  }

  if (state.producedSheets === 0) {
    writeGridToSheet(outWb, "No output", {
      headers: ["Notice"],
      rows: [["No operation produced output. See warnings for details."], ...warnings.map((w) => [w] as CellValue[])],
    });
    warnings.push("No operation produced output.");
  }

  let buffer = await writeWorkbook(outWb);

  const explicitRecalc = plan.ops.some((o) => o.op === "recalc");
  if (hasMutating || explicitRecalc) {
    const r = await legacyRecalc(buffer);
    if (r.stats) {
      buffer = r.bytes;
      if (hasMutating && !explicitRecalc) opLogs.push({ op: "recalc_auto", status: "ok", ...r.stats });
      else opLogs.push({ op: "recalc_bridge", status: "ok", ...r.stats });
    } else {
      // Legacy runtime unavailable: keep SheetJS output (cached formula results preserved).
      console.warn(`[engine] legacy recalc skipped: ${r.error}`);
      opLogs.push({ op: hasMutating && !explicitRecalc ? "recalc_auto" : "recalc_bridge", status: "skipped", reason: r.error });
    }
  }

  stats.totalMs = Date.now() - startAll;
  stats.producedSheets = state.producedSheets;
  if (ctxBase.projectionEvents.length > 0) {
    stats.projection = ctxBase.projectionEvents;
  }

  return { buffer, stats, warnings };
}
