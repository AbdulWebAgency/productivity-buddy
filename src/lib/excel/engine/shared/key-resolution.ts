// Strict key-column resolution.
//
// `resolveColumn` in ./headers is deliberately permissive: it is used for
// projection, highlighting and planner gates, where a loose guess is useful
// and a miss is harmless. Key columns have the opposite failure preference —
// a wrong match silently produces wrong joins — so they get their own
// resolver that refuses rather than guesses.
//
// Differences from resolveColumn:
//  - full-token containment instead of raw substring ("id" no longer matches
//    "Valid Until", "name" no longer matches "Filename")
//  - every header is scored; extra tokens in the header are penalised
//  - near-ties return `ambiguous` instead of silently taking header order
//  - never returns a bare index, so callers must handle failure explicitly

import { normalizeHeader } from "./headers";

export type KeyCandidate = { index: number; header: string; score: number };

export type KeyResolution =
  | { status: "resolved"; index: number; header: string; confidence: number }
  | { status: "ambiguous"; candidates: KeyCandidate[] }
  | { status: "not_found" };

/** Minimum score a fuzzy candidate must reach to be considered at all. */
const MIN_SCORE = 0.34;
/** Top two candidates within this distance are treated as ambiguous. */
const TIE_EPSILON = 0.1;

function tokens(s: string): string[] {
  return normalizeHeader(s).split(" ").filter(Boolean);
}

/**
 * Score a header against a target name.
 * 1 = identical token sets. 0 = not a candidate.
 * Requires every target token to appear as a full token in the header, then
 * divides by the header's token count so extra tokens ("Reg No Verified Flag")
 * score lower than a tight match ("Reg No").
 */
export function scoreKeyMatch(header: string, target: string): number {
  const h = tokens(header);
  const t = tokens(target);
  if (h.length === 0 || t.length === 0) return 0;
  const hSet = new Set(h);
  if (!t.every((tok) => hSet.has(tok))) return 0;
  return t.length / h.length;
}

export function resolveKeyColumn(headers: string[], name: string): KeyResolution {
  const target = normalizeHeader(name);
  if (!target) return { status: "not_found" };

  const exact = headers.findIndex((h) => normalizeHeader(h) === target);
  if (exact >= 0) {
    return { status: "resolved", index: exact, header: headers[exact], confidence: 100 };
  }

  const scored: KeyCandidate[] = headers
    .map((h, index) => ({ index, header: h, score: scoreKeyMatch(h, name) }))
    .filter((c) => c.score >= MIN_SCORE)
    .sort((a, b) => b.score - a.score);

  if (scored.length === 0) return { status: "not_found" };
  if (scored.length > 1 && scored[0].score - scored[1].score < TIE_EPSILON) {
    return { status: "ambiguous", candidates: scored.slice(0, 5) };
  }
  const top = scored[0];
  return {
    status: "resolved",
    index: top.index,
    header: top.header,
    confidence: Math.round(top.score * 100),
  };
}

/** Resolve or throw with a message that names the file and the ambiguity. */
export function requireKeyColumn(headers: string[], name: string, fileName: string): {
  index: number;
  header: string;
  confidence: number;
} {
  const r = resolveKeyColumn(headers, name);
  if (r.status === "resolved") return { index: r.index, header: r.header, confidence: r.confidence };
  if (r.status === "ambiguous") {
    const list = r.candidates.map((c) => `"${c.header}"`).join(", ");
    throw new Error(
      `Key column "${name}" is ambiguous in "${fileName}" — it could be ${list}. Tell me which column to use.`,
    );
  }
  throw new Error(`Key column "${name}" not found in "${fileName}".`);
}

export type PerFileKey = { fileName: string; index: number; header: string; confidence: number };

/**
 * Resolve the same key name across several files and verify the files agree.
 * Prevents file A landing on "Student ID" while file B lands on
 * "Old Student ID" — the join would succeed but be wrong.
 */
export function resolveKeyAcrossFiles(
  headerSets: string[][],
  name: string,
  fileNames: string[],
): PerFileKey[] {
  const resolved = headerSets.map((headers, i) =>
    ({ fileName: fileNames[i] ?? `file ${i + 1}`, ...requireKeyColumn(headers, name, fileNames[i] ?? `file ${i + 1}`) }),
  );
  const normalized = resolved.map((r) => normalizeHeader(r.header));
  const first = normalized[0];
  const mismatch = resolved.findIndex((_, i) => normalized[i] !== first);
  if (mismatch > 0) {
    throw new Error(
      `Key column "${name}" resolved to different columns: "${resolved[0].header}" in "${resolved[0].fileName}" but "${resolved[mismatch].header}" in "${resolved[mismatch].fileName}". Tell me the exact column name to match on.`,
    );
  }
  return resolved;
}
