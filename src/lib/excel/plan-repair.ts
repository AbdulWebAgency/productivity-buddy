// Repairs common AI planner output variations into a Zod-valid Plan.
// The planner uses generateText (JSON via prompt) rather than a strict schema,
// so we normalize field names, casing, and file references before validation.

import { PlanSchema, type Plan } from "./types";

type FileForAi = { index: number; name: string; sheets: { name: string; headers: string[] }[] };

export function repairAndParsePlan(
  rawText: string,
  files: FileForAi[],
): { plan: Plan; repairs: string[] } {
  const repairs: string[] = [];

  const jsonText = extractJson(rawText);
  if (!jsonText) {
    throw new Error(
      `Planner did not return JSON. First 300 chars: ${rawText.slice(0, 300)}`,
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    throw new Error(
      `Planner returned invalid JSON: ${(e as Error).message}. First 300 chars: ${jsonText.slice(0, 300)}`,
    );
  }

  const normalized = normalizePlan(parsed, files, repairs);
  const result = PlanSchema.safeParse(normalized);
  if (!result.success) {
    const issues = result.error.issues
      .slice(0, 5)
      .map((i) => `${i.path.join(".") || "(root)"}: ${i.message}`)
      .join("; ");
    throw new Error(
      `Planner output failed schema validation after repair (${issues}). Raw: ${JSON.stringify(normalized).slice(0, 400)}`,
    );
  }
  return { plan: result.data, repairs };
}

// Pull JSON out of arbitrary text (markdown fences, prose wrappers).
function extractJson(text: string): string | null {
  if (!text) return null;
  const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = (fenced?.[1] ?? text).trim();
  const start = candidate.indexOf("{");
  const end = candidate.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  return candidate.slice(start, end + 1);
}

type LooseObj = Record<string, unknown>;

function isObj(v: unknown): v is LooseObj {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

const KNOWN_OPS = new Set([
  "merge",
  "dedupe",
  "diff",
  "summary",
  "recalc",
  "highlight_column",
]);

function normalizePlan(input: unknown, files: FileForAi[], repairs: string[]): LooseObj {
  const plan: LooseObj = isObj(input) ? { ...input } : {};

  // ops may live under a different key.
  let ops = plan.ops ?? plan.operations ?? plan.steps ?? plan.actions;
  if (isObj(ops)) ops = [ops];
  if (!Array.isArray(ops)) ops = [];

  const normalizedOps = (ops as unknown[])
    .map((raw) => normalizeOp(raw, files, repairs))
    .filter((o): o is LooseObj => o !== null);

  plan.ops = normalizedOps;
  if (typeof plan.summary !== "string") plan.summary = "";
  if (!Array.isArray(plan.warnings)) plan.warnings = [];
  if (!Array.isArray(plan.columnMappings)) plan.columnMappings = [];
  return plan;
}

function normalizeOp(raw: unknown, files: FileForAi[], repairs: string[]): LooseObj | null {
  if (!isObj(raw)) return null;
  const o: LooseObj = { ...raw };

  // Normalize the tag.
  let tag = String(o.op ?? o.type ?? o.action ?? o.kind ?? "").toLowerCase().trim();
  if (tag === "compare" || tag === "difference") tag = "diff";
  if (tag === "deduplicate" || tag === "unique") tag = "dedupe";
  if (tag === "join" || tag === "combine") tag = "merge";
  if (tag === "highlight") tag = "highlight_column";
  if (tag === "recalculate") tag = "recalc";
  if (!KNOWN_OPS.has(tag)) {
    repairs.push(`Dropped unknown op "${String(o.op ?? o.type ?? "")}"`);
    return null;
  }
  o.op = tag;

  // Normalize column key aliases across ops.
  const columnAliases = ["keyColumn", "key_column", "primaryColumn", "primary_column", "column", "onColumn", "on"];
  const pickColumn = (): string | undefined => {
    for (const k of columnAliases) {
      const v = (o as LooseObj)[k];
      if (typeof v === "string" && v.trim()) return v.trim();
    }
    // Nested { file1: { primaryColumn }, file2: { primaryColumn } }
    for (const nk of ["file1", "file2", "fileA", "fileB", "source", "target"]) {
      const nested = (o as LooseObj)[nk];
      if (isObj(nested)) {
        for (const k of columnAliases) {
          const v = nested[k];
          if (typeof v === "string" && v.trim()) return v.trim();
        }
      }
    }
    return undefined;
  };

  const resolveFileIndex = (v: unknown): number | undefined => {
    if (typeof v === "number" && Number.isInteger(v)) return v;
    if (isObj(v)) {
      const nested =
        v.fileIndex ?? v.index ?? v.file_index ?? v.fileName ?? v.name ?? v.file;
      return resolveFileIndex(nested);
    }
    if (typeof v === "string") {
      const norm = v.trim().toLowerCase();
      const match = files.find((f) => f.name.toLowerCase() === norm);
      if (match) return match.index;
      const partial = files.find(
        (f) => f.name.toLowerCase().includes(norm) || norm.includes(f.name.toLowerCase()),
      );
      if (partial) return partial.index;
    }
    return undefined;
  };

  if (tag === "diff") {
    const keyColumn = pickColumn();
    if (!keyColumn) {
      repairs.push("diff op missing keyColumn — dropping");
      return null;
    }
    let a =
      resolveFileIndex(o.fileAIndex) ??
      resolveFileIndex(o.fileIndexA) ??
      resolveFileIndex(o.file1) ??
      resolveFileIndex(o.fileA) ??
      resolveFileIndex(o.source) ??
      resolveFileIndex(o.target);
    let b =
      resolveFileIndex(o.fileBIndex) ??
      resolveFileIndex(o.fileIndexB) ??
      resolveFileIndex(o.file2) ??
      resolveFileIndex(o.fileB);
    if (a === undefined) a = 0;
    if (b === undefined) b = a === 0 && files.length > 1 ? 1 : Math.max(0, files.length - 1);
    if (a === b) {
      repairs.push("diff had identical file indices — using files 0 and 1");
      a = 0;
      b = 1;
    }
    return { op: "diff", keyColumn, fileAIndex: a, fileBIndex: b };
  }

  if (tag === "merge") {
    const keyColumn = pickColumn();
    if (!keyColumn) {
      repairs.push("merge op missing keyColumn — dropping");
      return null;
    }
    const strategy = o.strategy === "intersection" ? "intersection" : "union";
    const highlightUnmatched = o.highlightUnmatched !== false;
    return { op: "merge", keyColumn, strategy, highlightUnmatched };
  }

  if (tag === "dedupe") {
    const keyColumn = pickColumn();
    const strategy = o.strategy === "full_row" ? "full_row" : keyColumn ? "key" : "full_row";
    const out: LooseObj = { op: "dedupe", strategy };
    if (keyColumn && strategy === "key") out.keyColumn = keyColumn;
    return out;
  }

  if (tag === "highlight_column") {
    const column = pickColumn();
    const ruleRaw = String(o.rule ?? o.type ?? "").toLowerCase();
    const rule =
      ruleRaw === "missing" || ruleRaw === "duplicate" || ruleRaw === "outlier"
        ? ruleRaw
        : "missing";
    if (!column) {
      repairs.push("highlight_column missing column — dropping");
      return null;
    }
    return { op: "highlight_column", column, rule };
  }

  if (tag === "summary") {
    return { op: "summary", includeCharts: o.includeCharts !== false };
  }

  if (tag === "recalc") {
    return { op: "recalc" };
  }

  return null;
}
