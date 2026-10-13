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

const EMP = file("employees.xlsx", [
  ["Employee ID", "Name", "Department", "Salary"],
  ["E001", "Anil", "Engineering", 62000],
  ["E002", "Bela", null, 54000],
  ["E002", "Bela2", "Sales", 55000],
  ["E004", "Dina", null, 51000],
]);

describe("highlight_column marks matches in a visible PB Flag column", () => {
  test("missing: flags only the blank-Department rows", async () => {
    const r = await runPlan([EMP], plan([{ op: "highlight_column", column: "Department", rule: "missing" }]));
    const rows = read(r.buffer).sheets.Result;
    expect(rows[0]).toEqual(["Employee ID", "Name", "Department", "Salary", "PB Flag"]);
    expect(rows[1][4]).toBeUndefined();
    expect(rows[2][4]).toBe("Missing Department");
    expect(rows[3][4]).toBeUndefined();
    expect(rows[4][4]).toBe("Missing Department");
  });

  test("two highlight steps share one PB Flag column and combine labels", async () => {
    const r = await runPlan(
      [EMP],
      plan([
        { op: "highlight_column", column: "Department", rule: "missing" },
        { op: "highlight_column", column: "Employee ID", rule: "duplicate" },
      ]),
    );
    const rows = read(r.buffer).sheets.Result;
    expect(rows[0].filter((h) => h === "PB Flag").length).toBe(1);
    expect(rows[1][4]).toBeUndefined();
    expect(rows[2][4]).toBe("Missing Department; Duplicate Employee ID");
    expect(rows[3][4]).toBe("Duplicate Employee ID");
    expect(rows[4][4]).toBe("Missing Department");
  });

  test("no matches: no PB Flag column is added", async () => {
    const r = await runPlan([EMP], plan([{ op: "highlight_column", column: "Name", rule: "missing" }]));
    const rows = read(r.buffer).sheets.Result;
    expect(rows[0]).toEqual(["Employee ID", "Name", "Department", "Salary"]);
    const log = (r.stats.ops as Record<string, unknown>[]).find((l) => l.op === "highlight_column")!;
    expect(log.matches).toBe(0);
  });
});