import { describe, expect, test } from "bun:test";
import * as XLSX from "xlsx";
import { readWorkbook } from "../src/lib/excel/engine/shared/workbook";
import { recalcFormulas } from "../src/lib/excel/engine/ops/highlight";
import { runPlan } from "../src/lib/excel/engine/runPlan";
import { PlanSchema } from "../src/lib/excel/types";

// Formula cells deliberately carry STALE cached values so recalculation is observable.
function fixture(): ArrayBuffer {
  const b = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Id", "A", "B", "Total"],
    [1, 2, 3, null],
    [2, 4, 5, null],
  ]);
  ws["D2"] = { t: "n", v: 0, f: "B2+C2" };
  ws["D3"] = { t: "n", v: 0, f: "SUM(B3:C3)" };
  ws["E2"] = { t: "n", v: 7, f: "NOTAREALFUNCTION(B2)" };
  ws["!ref"] = "A1:E3";
  XLSX.utils.book_append_sheet(b, ws, "Data");
  return XLSX.write(b, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}
const plan = (ops: unknown[]) => PlanSchema.parse({ ops });
const cellOf = (buf: Buffer, sheet: string, addr: string) =>
  XLSX.read(buf, { type: "buffer", cellFormula: true }).Sheets[sheet][addr];

describe("recalcFormulas on SheetJS handles", () => {
  test("recalculates simple + multiple formulas, skips unsupported", async () => {
    const wb = await readWorkbook(fixture());
    const r = recalcFormulas(wb);
    const ws = wb.getWorksheet("Data")!;
    expect(ws.getCell(2, 4).result).toBe(5);
    expect(ws.getCell(3, 4).result).toBe(9);
    expect(r.recalculated).toBe(2);
    expect(r.skipped).toBe(1);
    // unsupported formula left untouched
    expect(ws.getCell(2, 5).formula).toBe("NOTAREALFUNCTION(B2)");
    expect(ws.getCell(2, 5).result).toBe(7);
  });

  test("formulas remain formulas with updated cached results", async () => {
    const wb = await readWorkbook(fixture());
    recalcFormulas(wb);
    const ws = wb.getWorksheet("Data")!;
    expect(ws.getCell(2, 4).value).toEqual({ formula: "B2+C2", result: 5 });
  });

  test("setFormulaResult refuses non-formula cells", async () => {
    const ws = (await readWorkbook(fixture())).getWorksheet("Data")!;
    expect(ws.setFormulaResult(2, 2, 99)).toBe(false);
    expect(ws.getCell(2, 2).value).toBe(2);
  });

  test("explicit recalc op through runPlan", async () => {
    // summary is non-mutating; recalc acts on the fresh output workbook
    const r = await runPlan([{ name: "f.xlsx", buffer: fixture() }], plan([{ op: "summary" }, { op: "recalc" }]));
    const logs = r.stats.ops as Record<string, unknown>[];
    expect(logs.some((l) => l.op === "recalc" && l.status === "ok")).toBe(true);
    expect(logs.some((l) => l.op === "recalc_auto")).toBe(false);
  });

  test("automatic recalc on mutating op updates source-sheet formulas", async () => {
    const r = await runPlan([{ name: "f.xlsx", buffer: fixture() }], plan([{ op: "dedupe", strategy: "full_row" }]));
    const logs = r.stats.ops as Record<string, unknown>[];
    const auto = logs.find((l) => l.op === "recalc_auto")!;
    expect(auto.status).toBe("ok");
    expect(auto.recalculated).toBe(2);
    const d2 = cellOf(r.buffer, "Data", "D2");
    expect(d2.f).toBe("B2+C2");
    expect(d2.v).toBe(5);
  });

  test("recalc-only plan returns the user's workbook with formulas kept and cached values refreshed", async () => {
    const r = await runPlan([{ name: "f.xlsx", buffer: fixture() }], plan([{ op: "recalc" }]));
    expect(r.warnings.join(" ")).not.toContain("No operation produced output");
    const out = XLSX.read(r.buffer, { type: "buffer", cellFormula: true });
    expect(out.SheetNames).toEqual(["Data"]); // no "No output" sheet, no extra Result sheet
    expect(out.Sheets["Data"]["D2"].f).toBe("B2+C2"); // formula preserved
    expect(out.Sheets["Data"]["D2"].v).toBe(5); // stale 0 refreshed
    expect(out.Sheets["Data"]["D3"].v).toBe(9);
    const rc = (r.stats.ops as Record<string, unknown>[]).find((l) => l.op === "recalc")!;
    expect(rc.recalculated).toBe(2);
    expect(rc.skipped).toBe(1);
  });

  test("recalculation runs on the SheetJS boundary with no legacy bridge", async () => {
    const src = await Bun.file("src/lib/excel/engine/ops/highlight.ts").text();
    const run = await Bun.file("src/lib/excel/engine/runPlan.ts").text();
    expect(src).toContain('from "../shared/workbook"');
    expect(run).not.toMatch(/legacyRecalc|readLegacy/);
  });
});
