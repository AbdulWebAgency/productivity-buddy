// Plan orchestrator. Reads inputs, dispatches ops through the registry,
// finalises the workbook, returns bytes + stats.
import type { Plan, PlanOp } from "../types";
import type { EngineFile, EngineResult } from "./types";
import { readWorkbook, sheetToGrid, writeGridToSheet } from "./shared/workbook";
import type { CellValue } from "./shared/workbook";
import { applyProjection, type EngineState, type OpCtx } from "./registry";
import { recalcFormulas } from "./ops/highlight";
import { getOpHandler } from "./ops";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const ExcelJSModule = require("exceljs") as typeof import("exceljs");

export async function runPlan(files: EngineFile[], plan: Plan): Promise<EngineResult> {
  const warnings: string[] = [...plan.warnings];
  const opLogs: Record<string, unknown>[] = [];
  const stats: Record<string, unknown> = { ops: opLogs };

  const workbooks = await Promise.all(files.map((f) => readWorkbook(f.buffer)));
  const grids = workbooks.map((wb) => sheetToGrid(wb.worksheets[0]));

  // Determine output strategy.
  const hasMutating = plan.ops.some((o) => o.op === "merge" || o.op === "dedupe" || o.op === "highlight_column");
  // Merge/dedupe/highlight can preserve first-workbook styling by mutating a
  // copy of the first workbook. Diff/summary produce brand-new deliverables.
  const outWb = hasMutating ? workbooks[0] : new ExcelJSModule.Workbook();
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
    outWb,
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
    // Compute alwaysKeep for the final "Result" sheet from the ops that fed it.
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
    // Row highlight indices are stable across projection (row order preserved).
    writeGridToSheet(outWb, "Result", projectedResult, { highlightRows: state.unmatched });
    state.producedSheets++;
    const resultWs = outWb.getWorksheet("Result");
    if (resultWs) {
      outWb.worksheets.splice(outWb.worksheets.indexOf(resultWs), 1);
      outWb.worksheets.unshift(resultWs);
    }
  }

  if (state.producedSheets === 0) {
    writeGridToSheet(outWb, "No output", {
      headers: ["Notice"],
      rows: [["No operation produced output. See warnings for details."], ...warnings.map((w) => [w] as CellValue[])],
    });
    warnings.push("No operation produced output.");
  }

  if (hasMutating && !plan.ops.some((o) => o.op === "recalc")) {
    const r = recalcFormulas(outWb);
    opLogs.push({ op: "recalc_auto", status: "ok", ...r });
  }

  stats.totalMs = Date.now() - startAll;
  stats.producedSheets = state.producedSheets;
  if (ctxBase.projectionEvents.length > 0) {
    stats.projection = ctxBase.projectionEvents;
  }

  const out = await outWb.xlsx.writeBuffer();
  return { buffer: Buffer.from(out), stats, warnings };
}
