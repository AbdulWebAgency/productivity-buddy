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

// File 0 = directory (priya, zoe). File 1 = staff (priya, yara). Only priya is in both.
const DIRECTORY = file("directory.xlsx", [
  ["Email", "Department"],
  ["priya@acme.example", "Engineering"],
  ["zoe@acme.example", "Marketing"],
]);
const STAFF = file("staff.xlsx", [
  ["Employee ID", "Email"],
  ["E01", "priya@acme.example"],
  ["E08", "yara@acme.example"],
]);

describe("master_merge reports rows left out by the join", () => {
  test("left join: warns about keys that only exist outside file 0 and says so in the summary", async () => {
    const r = await runPlan(
      [DIRECTORY, STAFF],
      plan([{ op: "master_merge", keyColumn: "Email", joinType: "left", dupeStrategy: "first" }]),
    );
    const out = read(r.buffer);
    expect(out.sheets.Master.length).toBe(3); // header + priya + zoe
    expect(JSON.stringify(out.sheets.Master)).not.toContain("yara");
    expect(r.warnings.join(" ")).toContain("yara@acme.example");
    expect(r.warnings.join(" ")).toContain("left out by the left join");
    expect(JSON.stringify(out.sheets["Merge Summary"])).toContain("Keys left out by the left join");
  });

  test("outer join: nothing is left out, so no warning and no extra summary row", async () => {
    const r = await runPlan(
      [DIRECTORY, STAFF],
      plan([{ op: "master_merge", keyColumn: "Email", joinType: "outer", dupeStrategy: "first" }]),
    );
    const out = read(r.buffer);
    expect(out.sheets.Master.length).toBe(4); // header + priya + zoe + yara
    expect(r.warnings.join(" ")).not.toContain("left out");
    expect(JSON.stringify(out.sheets["Merge Summary"])).not.toContain("left out");
  });
});