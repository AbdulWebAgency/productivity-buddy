// Deterministic planner. Skips the AI when the intent + column headers are
// unambiguous. Returns null when human/AI judgement is needed.

import type { Plan, PlanOp } from "./types";

export type FileForAi = {
  index: number;
  name: string;
  sheets: { name: string; headers: string[] }[];
};

const KEY_HINTS = [
  "registration",
  "reg no",
  "regno",
  "roll",
  "student id",
  "employee id",
  "user id",
  "email",
  "phone",
  "mobile",
  "code",
  "id",
  "number",
  "sl no",
  "serial",
];

function norm(s: string): string {
  return s.trim().toLowerCase().replace(/[._\-]+/g, " ").replace(/\s+/g, " ");
}

function fileHeaders(f: FileForAi): string[] {
  // Prefer first sheet — engine also uses first sheet only.
  return f.sheets[0]?.headers ?? [];
}

// Columns that appear (case/punctuation-insensitive) in every file's first sheet.
function commonColumns(files: FileForAi[]): { display: string; perFile: string[] }[] {
  if (files.length === 0) return [];
  const perFileMaps = files.map((f) => {
    const m = new Map<string, string>();
    fileHeaders(f).forEach((h) => m.set(norm(h), h));
    return m;
  });
  const base = perFileMaps[0];
  const out: { display: string; perFile: string[] }[] = [];
  for (const [k, display] of base) {
    const perFile: string[] = [display];
    let all = true;
    for (let i = 1; i < perFileMaps.length; i++) {
      const hit = perFileMaps[i].get(k);
      if (!hit) { all = false; break; }
      perFile.push(hit);
    }
    if (all) out.push({ display, perFile });
  }
  return out;
}

function pickKeyColumn(shared: { display: string; perFile: string[] }[]): string | null {
  if (shared.length === 0) return null;
  // Score by hint match; prefer more specific hints.
  const scored = shared.map((c) => {
    const n = norm(c.display);
    let score = 0;
    for (const h of KEY_HINTS) if (n.includes(h)) score += h.length;
    return { c, score };
  });
  scored.sort((a, b) => b.score - a.score);
  if (scored[0].score > 0) return scored[0].c.display;
  // Single shared column? use it.
  if (shared.length === 1) return shared[0].display;
  return null;
}

function detectOp(intent: string): "diff" | "merge" | "dedupe" | null {
  const s = intent.toLowerCase();
  if (/\b(missing|not in|absent|only in|present in .* (?:but|and) .* not|difference|compare|diff)\b/.test(s)) return "diff";
  if (/\b(merge|combine|join|consolidat|union|intersect)\b/.test(s)) return "merge";
  if (/\b(dedup|duplicate|unique)\b/.test(s)) return "dedupe";
  return null;
}

export type DeterministicResult =
  | { kind: "plan"; plan: Plan; notes: string[] }
  | { kind: "needs_clarification"; reason: string; sharedColumns: string[]; candidateKeys: string[] };

export function tryDeterministicPlan(
  intent: string | null | undefined,
  files: FileForAi[],
): DeterministicResult | null {
  const trimmed = (intent ?? "").trim();
  if (!trimmed) return null;
  const opKind = detectOp(trimmed);
  if (!opKind) return null;

  const shared = commonColumns(files);
  const sharedNames = shared.map((s) => s.display);

  if (opKind === "diff") {
    if (files.length < 2) return null;
    const key = pickKeyColumn(shared);
    if (!key) {
      return {
        kind: "needs_clarification",
        reason: `Could not find a shared key column across "${files[0]?.name}" and "${files[1]?.name}". Please tell me which column to match on.`,
        sharedColumns: sharedNames,
        candidateKeys: sharedNames,
      };
    }
    const ops: PlanOp[] = [{ op: "diff", keyColumn: key, fileAIndex: 0, fileBIndex: 1 }];
    return {
      kind: "plan",
      plan: {
        summary: `Deterministic diff on "${key}" between "${files[0].name}" and "${files[1].name}".`,
        ops,
        columnMappings: [
          {
            canonical: key,
            perFile: files.map((f, i) => ({ fileIndex: i, column: shared.find((s) => s.display === key)?.perFile[i] ?? key })),
          },
        ],
        warnings: [],
      },
      notes: [`Skipped AI: exact header match on "${key}".`],
    };
  }

  if (opKind === "merge") {
    if (files.length < 2) return null;
    const key = pickKeyColumn(shared);
    if (!key) {
      return {
        kind: "needs_clarification",
        reason: `Could not find a shared key column across the uploaded files. Please tell me which column to merge on.`,
        sharedColumns: sharedNames,
        candidateKeys: sharedNames,
      };
    }
    const ops: PlanOp[] = [
      { op: "merge", keyColumn: key, strategy: "union", highlightUnmatched: true },
    ];
    return {
      kind: "plan",
      plan: {
        summary: `Deterministic merge on "${key}".`,
        ops,
        columnMappings: [],
        warnings: [],
      },
      notes: [`Skipped AI: exact header match on "${key}".`],
    };
  }

  if (opKind === "dedupe") {
    const key = pickKeyColumn(shared);
    const ops: PlanOp[] = key
      ? [{ op: "dedupe", strategy: "key", keyColumn: key }]
      : [{ op: "dedupe", strategy: "full_row" }];
    return {
      kind: "plan",
      plan: {
        summary: key ? `Deterministic dedupe on "${key}".` : "Deterministic full-row dedupe.",
        ops,
        columnMappings: [],
        warnings: [],
      },
      notes: ["Skipped AI: deterministic dedupe."],
    };
  }

  return null;
}
