// Intent-driven deterministic planner.
// 1. Classify user intent into a concept (intersection, difference, merge, ...)
// 2. Resolve columns / files / keys deterministically
// 3. Only ask for clarification when genuinely ambiguous
//    (e.g. multiple equally-strong keys, or missing intent).

import type { Plan, PlanOp } from "./types";
import { classifyIntent, type Intent } from "./intent";

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
  "id no",
  "id",
  "number",
  "sl no",
  "serial",
];

function norm(s: string): string {
  return s.trim().toLowerCase().replace(/[._\-]+/g, " ").replace(/\s+/g, " ");
}

function fileHeaders(f: FileForAi): string[] {
  return f.sheets[0]?.headers ?? [];
}

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

// Score shared columns by "keyness". A single top-scoring column is used
// automatically. Ties across multiple strong keys trigger clarification.
function scoreKey(display: string): number {
  const n = norm(display);
  let score = 0;
  for (const h of KEY_HINTS) if (n.includes(h)) score = Math.max(score, h.length);
  return score;
}

function pickBestKey(
  shared: { display: string; perFile: string[] }[],
  preferred?: string | null,
): { key: string | null; tiedCandidates: string[] } {
  if (shared.length === 0) return { key: null, tiedCandidates: [] };

  if (preferred) {
    const match = shared.find((s) => norm(s.display) === norm(preferred));
    if (match) return { key: match.display, tiedCandidates: [] };
    // Preferred not in shared — still return it; engine's fuzzy findColumnIndex
    // will try to resolve per file.
    return { key: preferred, tiedCandidates: [] };
  }

  if (shared.length === 1) return { key: shared[0].display, tiedCandidates: [] };

  const scored = shared
    .map((c) => ({ c, score: scoreKey(c.display) }))
    .sort((a, b) => b.score - a.score);

  const top = scored[0];
  if (top.score === 0) {
    // No hint match at all → ambiguous; return all shared as candidates
    return { key: null, tiedCandidates: shared.map((s) => s.display) };
  }
  const tied = scored.filter((s) => s.score === top.score).map((s) => s.c.display);
  if (tied.length === 1) return { key: top.c.display, tiedCandidates: [] };
  return { key: null, tiedCandidates: tied };
}

export type DeterministicResult =
  | { kind: "plan"; plan: Plan; notes: string[]; intent: Intent }
  | {
      kind: "needs_clarification";
      reason: string;
      sharedColumns: string[];
      candidateKeys: string[];
      pendingIntent: Intent;
      pendingSide?: "A" | "B" | "either";
    }
  | { kind: "capabilities" };

export type PlannerOptions = {
  // A prior unresolved intent from earlier in the conversation. If the current
  // message is a short key-name reply (e.g. "Name"), we resume that intent.
  resumeIntent?: Intent;
  resumeSide?: "A" | "B" | "either";
  // A user-specified key override (from "use X instead" or plain key name).
  preferredKey?: string | null;
};

export function tryDeterministicPlan(
  intent: string | null | undefined,
  files: FileForAi[],
  opts: PlannerOptions = {},
): DeterministicResult | null {
  const trimmed = (intent ?? "").trim();
  if (!trimmed) return null;

  const classified = classifyIntent(trimmed);
  let effectiveIntent: Intent = classified.intent;
  let side = classified.side;

  // Resume prior unresolved intent when the current message doesn't classify
  // (e.g. user is just replying with a column name).
  if (effectiveIntent === "unknown" && opts.resumeIntent) {
    effectiveIntent = opts.resumeIntent;
    side = opts.resumeSide;
  }

  if (effectiveIntent === "capabilities") return { kind: "capabilities" };
  if (effectiveIntent === "unknown") return null;

  const shared = commonColumns(files);
  const sharedNames = shared.map((s) => s.display);

  const picked = pickBestKey(shared, opts.preferredKey ?? null);
  const buildKeyOp = (): { key: string | null; tied: string[] } => {
    if (picked.key) return { key: picked.key, tied: [] };
    return { key: null, tied: picked.tiedCandidates.length ? picked.tiedCandidates : sharedNames };
  };

  if (effectiveIntent === "intersection") {
    if (files.length < 2) return null;
    const { key, tied } = buildKeyOp();
    if (!key) {
      return {
        kind: "needs_clarification",
        reason: `To find rows present in both "${files[0]?.name}" and "${files[1]?.name}", pick a column to match on.`,
        sharedColumns: sharedNames,
        candidateKeys: tied,
        pendingIntent: "intersection",
      };
    }
    const ops: PlanOp[] = [
      { op: "intersection", keyColumn: key, fileAIndex: 0, fileBIndex: 1 },
    ];
    return {
      kind: "plan",
      intent: "intersection",
      plan: {
        summary: `Rows present in both "${files[0].name}" and "${files[1].name}", matched on "${key}".`,
        ops,
        columnMappings: [],
        warnings: [],
      },
      notes: [`Deterministic intersection on "${key}".`],
    };
  }

  if (effectiveIntent === "difference") {
    if (files.length < 2) return null;
    const { key, tied } = buildKeyOp();
    if (!key) {
      return {
        kind: "needs_clarification",
        reason: `To compare "${files[0]?.name}" and "${files[1]?.name}", pick a column to match on.`,
        sharedColumns: sharedNames,
        candidateKeys: tied,
        pendingIntent: "difference",
        pendingSide: side,
      };
    }
    const ops: PlanOp[] = [{ op: "diff", keyColumn: key, fileAIndex: 0, fileBIndex: 1 }];
    return {
      kind: "plan",
      intent: "difference",
      plan: {
        summary: `Difference between "${files[0].name}" and "${files[1].name}" on "${key}".`,
        ops,
        columnMappings: [],
        warnings: side ? [`Requested side: only in file ${side}.`] : [],
      },
      notes: [`Deterministic difference on "${key}".`],
    };
  }

  if (effectiveIntent === "merge") {
    if (files.length < 2) return null;
    const { key, tied } = buildKeyOp();
    if (!key) {
      return {
        kind: "needs_clarification",
        reason: `Which column should I merge on?`,
        sharedColumns: sharedNames,
        candidateKeys: tied,
        pendingIntent: "merge",
      };
    }
    const ops: PlanOp[] = [
      { op: "merge", keyColumn: key, strategy: "union", highlightUnmatched: true },
    ];
    return {
      kind: "plan",
      intent: "merge",
      plan: {
        summary: `Merge all files on "${key}".`,
        ops,
        columnMappings: [],
        warnings: [],
      },
      notes: [`Deterministic merge on "${key}".`],
    };
  }

  if (effectiveIntent === "master_merge") {
    if (files.length < 2) {
      return {
        kind: "needs_clarification",
        reason: "Upload at least 2 files before creating a master sheet.",
        sharedColumns: sharedNames,
        candidateKeys: [],
        pendingIntent: "master_merge",
      };
    }
    const { key, tied } = buildKeyOp();
    if (!key) {
      return {
        kind: "needs_clarification",
        reason: `Which column should I use as the master key across all ${files.length} files?`,
        sharedColumns: sharedNames,
        candidateKeys: tied,
        pendingIntent: "master_merge",
      };
    }
    const jt = /\bappend\b/i.test(trimmed)
      ? "append"
      : /\binner\s*join\b|\bkeep\s+(only\s+)?common\b/i.test(trimmed)
        ? "inner"
        : /\bleft\s*join\b/i.test(trimmed)
          ? "left"
          : "outer";
    const ds = /\bkeep\s+latest\b|\blast\s+wins\b/i.test(trimmed)
      ? "latest"
      : /\bmerge\s+dup|\bcombine\s+dup/i.test(trimmed)
        ? "merge"
        : "first";
    return {
      kind: "plan",
      intent: "master_merge",
      plan: {
        summary: `Create master sheet across ${files.length} files on "${key}" (${jt}, keep ${ds}).`,
        ops: [{ op: "master_merge", keyColumn: key, joinType: jt, dupeStrategy: ds }],
        columnMappings: [],
        warnings: [],
      },
      notes: [`Master merge on "${key}" (${jt}).`],
    };
  }

  if (effectiveIntent === "bulk_lookup") {
    // Parse queries from the raw text: everything after the first newline, or
    // comma/whitespace-separated tokens on the same line after the trigger phrase.
    const lines = trimmed.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    let queries: string[] = [];
    let targetName: string | null = null;
    const firstLine = lines[0] ?? "";
    const inMatch = firstLine.match(/\bin\s+([^\n:]+?)(?::|$)/i);
    if (inMatch) targetName = inMatch[1].trim();
    if (lines.length > 1) {
      queries = lines.slice(1).flatMap((l) => l.split(/[,;\t]+/).map((s) => s.trim())).filter(Boolean);
    } else {
      const after = firstLine.split(/[:\-–]\s*/).slice(1).join(" ");
      queries = after
        .split(/[,;\s]+/)
        .map((s) => s.trim())
        .filter((s) => s && !/^(in|from|the|file)$/i.test(s));
    }
    // pick file: named or default to last (most recent upload)
    let fileIndex = files.length - 1;
    if (targetName) {
      const tn = norm(targetName);
      const hit = files.find((f) => norm(f.name).includes(tn) || tn.includes(norm(f.name)));
      if (hit) fileIndex = hit.index;
    }
    if (queries.length === 0) {
      return {
        kind: "needs_clarification",
        reason:
          "Paste the IDs, names, or emails to look up — one per line — after the request. Example:\n\nBulk lookup in students.xlsx:\nABC001\nABC002",
        sharedColumns: sharedNames,
        candidateKeys: [],
        pendingIntent: "bulk_lookup",
      };
    }
    return {
      kind: "plan",
      intent: "bulk_lookup",
      plan: {
        summary: `Bulk lookup of ${queries.length} value(s) in "${files[fileIndex]?.name ?? "file"}".`,
        ops: [{ op: "bulk_lookup", fileIndex, queries }],
        columnMappings: [],
        warnings: [],
      },
      notes: [`Bulk lookup (${queries.length} queries).`],
    };
  }


  if (effectiveIntent === "dedupe") {
    const { key } = buildKeyOp();
    const ops: PlanOp[] = key
      ? [{ op: "dedupe", strategy: "key", keyColumn: key }]
      : [{ op: "dedupe", strategy: "full_row" }];
    return {
      kind: "plan",
      intent: "dedupe",
      plan: {
        summary: key ? `Remove duplicates using "${key}".` : "Remove full-row duplicates.",
        ops,
        columnMappings: [],
        warnings: [],
      },
      notes: ["Deterministic dedupe."],
    };
  }

  if (effectiveIntent === "summary") {
    return {
      kind: "plan",
      intent: "summary",
      plan: {
        summary: "Add a summary sheet with row counts and numeric column stats.",
        ops: [{ op: "summary", includeCharts: false }],
        columnMappings: [],
        warnings: [],
      },
      notes: ["Deterministic summary."],
    };
  }

  if (effectiveIntent === "clean") {
    return {
      kind: "plan",
      intent: "clean",
      plan: {
        summary: "Remove full-row duplicates and blank rows.",
        ops: [{ op: "dedupe", strategy: "full_row" }],
        columnMappings: [],
        warnings: ["Blank-row cleanup applied via full-row dedupe."],
      },
      notes: ["Deterministic clean."],
    };
  }

  // formula_help doesn't produce a plan — let the conversational fallback handle it
  return null;
}
