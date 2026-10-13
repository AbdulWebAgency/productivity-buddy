import { describe, expect, test } from "bun:test";
import { buildOutputName } from "../src/lib/workspace/output-name";
import { PlanSchema } from "../src/lib/excel/types";

const plan = (ops: unknown[]) => PlanSchema.parse({ ops });

describe("buildOutputName", () => {
  test("one file: named after the file and the result", () => {
    const p = plan([{ op: "dedupe", strategy: "key", keyColumn: "Email" }]);
    expect(buildOutputName(p, ["customers_with_duplicates.xlsx"], 1)).toBe(
      "customers_with_duplicates_deduplicated_v1.xlsx",
    );
  });

  test("two files are joined with _and_", () => {
    const p = plan([{ op: "diff", keyColumn: "Email" }]);
    expect(buildOutputName(p, ["customers_old.xlsx", "customers_new.xlsx"], 1)).toBe(
      "customers_old_and_customers_new_comparison_v1.xlsx",
    );
  });

  test("three or more files are summarised as N_files and the version is used", () => {
    const p = plan([{ op: "master_merge", keyColumn: "Product ID", joinType: "append", dupeStrategy: "first" }]);
    expect(buildOutputName(p, ["sales_jan.xlsx", "sales_feb.xlsx", "sales_mar.xlsx"], 2)).toBe(
      "3_files_master_v2.xlsx",
    );
  });

  test("recalc-only plan is named after the first file only", () => {
    const p = plan([{ op: "recalc" }]);
    expect(buildOutputName(p, ["formulas stale.xlsx", "other.xlsx"], 1)).toBe("formulas_stale_recalculated_v1.xlsx");
  });

  test("messy names are cleaned and the extension is case-insensitive", () => {
    const p = plan([{ op: "summary", includeCharts: false }]);
    expect(buildOutputName(p, ["My Report (final) 2026.XLSX"], 3)).toBe("My_Report_final_2026_summary_v3.xlsx");
  });

  test("a name with no usable characters falls back to 'workbook'", () => {
    const p = plan([{ op: "summary", includeCharts: false }]);
    expect(buildOutputName(p, ["日本語.xlsx"], 1)).toBe("workbook_summary_v1.xlsx");
  });

  test("very long names are cut to 60 characters", () => {
    const p = plan([{ op: "dedupe", strategy: "key", keyColumn: "Email" }]);
    const name = buildOutputName(p, ["a".repeat(100) + ".xlsx"], 1);
    expect(name).toBe("a".repeat(60) + "_deduplicated_v1.xlsx");
  });
});