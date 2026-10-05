import { describe, expect, test } from "bun:test";
import * as XLSX from "xlsx";
import { runPlan } from "../src/lib/excel/engine/runPlan";
import { PlanSchema } from "../src/lib/excel/types";

function file(name: string, aoa: (string | number | null)[][]) {
  const b = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(aoa), "Sheet1");
  const out = XLSX.write(b, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
  return { name, buffer: out };
}
function read(buf: Buffer) {
  const b = XLSX.read(buf, { type: "buffer" });
  const sheets: Record<string, unknown[][]> = {};
  for (const n of b.SheetNames) sheets[n] = XLSX.utils.sheet_to_json(b.Sheets[n], { header: 1 });
  return { names: b.SheetNames, sheets };
}
const plan = (ops: unknown[]) => PlanSchema.parse({ ops });

const A = file("Test 1.xlsx", [
  ["Name", "Reg No", "Score"],
  ["Ana", "R1", 10],
  ["Ben", "R2", 20],
  ["Cy", "R3", 30],
]);
// reordered columns, normalized header, changed value, extra/missing rows
const B = file("Test 2.xlsx", [
  ["score", "reg_no", "name"],
  [25, "R2", "Ben"],
  [30, "r3", "Cy"],
  [40, "R4", "Dee"],
]);

describe("runPlan on SheetJS layer", () => {
  test("diff: reordered cols, changed values, extra/missing rows", async () => {
    const r = await runPlan([A, B], plan([{ op: "diff", keyColumn: "Reg No" }]));
    const out = read(r.buffer);
    const onlyB = out.names.find((n) => n.includes("Test 2"))!;
    const onlyA = out.names.find((n) => n.includes("Test 1"))!;
    expect(out.sheets[onlyB].slice(1).flat()).toContain("R4");
    expect(out.sheets[onlyA].slice(1).flat()).toContain("R1");
    expect(out.names).toContain("Changed rows");
    expect(JSON.stringify(out.sheets["Changed rows"]).toLowerCase()).toContain("r2");
  });

  test("diff: Changed rows keeps the key's original case and shows old → new values", async () => {
    const OLD = file("old.xlsx", [
      ["Customer ID", "City", "Plan"],
      ["C003", "Chennai", "Pro"],
      ["C004", "Seoul", "Basic"],
    ]);
    const NEW = file("new.xlsx", [
      ["Customer ID", "City", "Plan"],
      ["C003", "Bengaluru", "Pro"],
      ["C004", "Seoul", "Basic"],
    ]);
    const r = await runPlan([OLD, NEW], plan([{ op: "diff", keyColumn: "Customer ID" }]));
    const rows = read(r.buffer).sheets["Changed rows"];
    expect(rows[0]).toEqual(["Customer ID", "Changed Columns", "Values (old.xlsx → new.xlsx)"]);
    expect(rows[1]).toEqual(["C003", "City", "City: Chennai → Bengaluru"]); // not "c003"
    expect(rows.length).toBe(2); // C004 is unchanged
  });

  test("non-mutating: summary produces fresh workbook", async () => {
    const r = await runPlan([A], plan([{ op: "summary" }]));
    const out = read(r.buffer);
    expect(out.names).toEqual(["Summary"]);
  });

  test("mutating: dedupe puts Result first and keeps source sheet", async () => {
    const D = file("d.xlsx", [
      ["Id", "V"],
      [1, "a"],
      [1, "a"],
      [2, "b"],
    ]);
    const r = await runPlan([D], plan([{ op: "dedupe", strategy: "full_row" }]));
    const out = read(r.buffer);
    expect(out.names[0]).toBe("Result");
    expect(out.names).toContain("Sheet1");
    expect(out.sheets.Result.length).toBe(3);
  });

  test("No output sheet when nothing produced", async () => {
    const r = await runPlan([A], plan([{ op: "intersection", keyColumn: "Nope", fileBIndex: 5 }]));
    const out = read(r.buffer);
    expect(out.names).toContain("No output");
    expect(r.warnings).toContain("No operation produced output.");
  });
});
