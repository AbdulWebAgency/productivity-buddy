// Plan orchestrator. Reads inputs, dispatches ops through the registry,
// finalises the workbook, returns bytes + stats.
import ExcelJS from "exceljs";
import type { Plan } from "../types";
import type { EngineFile, EngineResult } from "./types";
import { readWorkbook, sheetToGrid, writeGridToSheet } from "./shared/workbook";
import type { CellValue } from "./shared/workbook";
import { getOpHandler, type EngineState, type OpCtx } from "./registry";
import { recalcFormulas } from "./ops/highlight";
import "./ops"; // side-effect: register all ops

export async function runPlan(files: EngineFile[], plan: Plan): Promise<EngineResult> {
  const warnings: string[] = [...plan.warnings];
  const opLogs: Record<string, unknown>[] = [];
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
  outWb.creator = "Productivity Buddy";
  outWb.created = new Date();

  const state: EngineState = {
    currentGrid: grids[0],
    unmatched: undefined,
    producedSheets: 0,
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
      const ctx: OpCtx = { files, grids, outWb, state, warnings, opLogs, started };
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
    writeGridToSheet(outWb, "Result", state.currentGrid, { highlightRows: state.unmatched });
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
  stats.producedSheets = state.producedSheets;

  const out = await outWb.xlsx.writeBuffer();
  return { buffer: Buffer.from(out), stats, warnings };
}
