// highlight_column + recalc live together — both are "post-processing" ops
// that don't produce their own output sheet.
import { HyperFormula } from "hyperformula";
import type { PlanOp } from "../../types";
import { resolveColumn } from "../shared/headers";
import { cellToValue, type WorkbookHandle } from "../shared/workbook";
import type { OpHandler } from "../registry";

/** Recalc formulas via HyperFormula on the workbook compatibility layer.
 *  Formulas are kept; only cached results change. Unsupported ones are left
 *  untouched and counted as skipped. */
export function recalcFormulas(wb: WorkbookHandle): { recalculated: number; skipped: number } {
  const sheets = wb.worksheets;
  const sheetsData: Record<string, (string | number | boolean | null)[][]> = {};
  sheets.forEach((ws) => {
    const data: (string | number | boolean | null)[][] = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const out: (string | number | boolean | null)[] = [];
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = ws.getCell(r, c);
        out.push(cell.formula != null ? "=" + cell.formula : cellToValue(cell.value));
      }
      data.push(out);
    }
    sheetsData[ws.name] = data;
  });

  let recalculated = 0;
  let skipped = 0;
  try {
    const hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: "gpl-v3" });
    sheets.forEach((ws) => {
      const sheetId = hf.getSheetId(ws.name);
      if (sheetId == null) return;
      for (let r = 1; r <= ws.rowCount; r++) {
        for (let c = 1; c <= ws.columnCount; c++) {
          if (ws.getCell(r, c).formula == null) continue;
          try {
            const result = hf.getCellValue({ sheet: sheetId, row: r - 1, col: c - 1 });
            if (result != null && typeof result !== "object" && ws.setFormulaResult(r, c, result)) {
              recalculated++;
            } else {
              skipped++;
            }
          } catch {
            skipped++;
          }
        }
      }
    });
    hf.destroy();
  } catch (e) {
    console.warn("HyperFormula init failed", e);
    skipped++;
  }
  return { recalculated, skipped };
}

export const highlightColumnHandler: OpHandler<Extract<PlanOp, { op: "highlight_column" }>> = (op, ctx) => {
  const idx = resolveColumn(ctx.state.currentGrid.headers, op.column);
  if (idx < 0) {
    ctx.warnings.push(`highlight_column skipped: column "${op.column}" not found`);
    ctx.opLogs.push({
      op: "highlight_column",
      status: "skipped",
      reason: `column "${op.column}" not found`,
    });
    return;
  }
  const highlightRows = new Set<number>();
  const cg = ctx.state.currentGrid;
  if (op.rule === "missing") {
    cg.rows.forEach((r, ri) => {
      if (r[idx] == null || r[idx] === "") highlightRows.add(ri);
    });
  } else if (op.rule === "duplicate") {
    const counts = new Map<string, number>();
    cg.rows.forEach((r) => {
      const k = String(r[idx] ?? "");
      counts.set(k, (counts.get(k) ?? 0) + 1);
    });
    cg.rows.forEach((r, ri) => {
      if ((counts.get(String(r[idx] ?? "")) ?? 0) > 1) highlightRows.add(ri);
    });
  } else if (op.rule === "outlier") {
    const nums = cg.rows.map((r) => Number(r[idx])).filter((n) => Number.isFinite(n));
    if (nums.length > 3) {
      const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
      const sd = Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / nums.length);
      cg.rows.forEach((r, ri) => {
        const n = Number(r[idx]);
        if (Number.isFinite(n) && Math.abs(n - mean) > 2 * sd) highlightRows.add(ri);
      });
    }
  }
  ctx.state.unmatched = new Set([...(ctx.state.unmatched ?? []), ...highlightRows]);
  ctx.opLogs.push({
    op: "highlight_column",
    status: "ok",
    ms: Date.now() - ctx.started,
    column: op.column,
    rule: op.rule,
    matches: highlightRows.size,
  });
};

export const recalcHandler: OpHandler<Extract<PlanOp, { op: "recalc" }>> = (_op, ctx) => {
  const r = recalcFormulas(ctx.outWb);
  ctx.opLogs.push({ op: "recalc", status: "ok", ms: Date.now() - ctx.started, ...r });
};
