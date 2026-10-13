// Builds download file names that describe the result,
// e.g. "customers_with_duplicates_deduplicated_v1.xlsx".
import type { Plan } from "../excel/types";

export function buildOutputName(plan: Plan, inputNames: string[], version: number): string {
  const kinds: string[] = plan.ops.map((o) => o.op);
  const has = (k: string) => kinds.includes(k);
  const slug = has("master_merge")
    ? "master"
    : has("bulk_lookup")
      ? "lookup"
      : has("merge")
        ? "merged"
        : has("intersection")
          ? "common_rows"
          : has("diff")
            ? "comparison"
            : has("dedupe")
              ? "deduplicated"
              : has("summary")
                ? "summary"
                : has("recalc")
                  ? "recalculated"
                  : has("highlight_column")
                    ? "flagged"
                    : "result";

  // A recalc-only plan refreshes just the first file, so name it after that file only.
  const recalcOnly = kinds.length > 0 && kinds.every((k) => k === "recalc");
  const used = recalcOnly ? inputNames.slice(0, 1) : inputNames;

  const stem = (n: string) =>
    n
      .replace(/\.(xlsx|xlsm|xls|csv)$/i, "")
      .replace(/[^a-zA-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  const stems = used.map(stem).filter((s) => s.length > 0);
  const source = stems.length === 0 ? "workbook" : used.length >= 3 ? `${used.length}_files` : stems.join("_and_");

  return `${source.slice(0, 60)}_${slug}_v${version}.xlsx`;
}