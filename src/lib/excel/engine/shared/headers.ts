// Canonical header normalization and matching utilities.
// Single source of truth used by the planner, inspector, AI conversation
// layer, and the Excel engine. Behaviour matches the pre-refactor duplicates.

export const KEY_HINTS = [
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

/** Trim, lowercase, collapse `._-` and whitespace runs. */
export function normalizeHeader(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .replace(/[._\-]+/g, " ")
    .replace(/\s+/g, " ");
}

/** Score how "key-like" a header is by longest KEY_HINTS substring match. */
export function scoreKey(header: string): number {
  const n = normalizeHeader(header);
  let score = 0;
  for (const h of KEY_HINTS) if (n.includes(h)) score = Math.max(score, h.length);
  return score;
}

/**
 * Resolve a header name against a list of headers using progressively looser
 * strategies: exact-normalized → substring either direction → token subset.
 * Returns -1 if not found.
 */
export function resolveColumn(headers: string[], name: string): number {
  const target = normalizeHeader(name);
  let idx = headers.findIndex((h) => normalizeHeader(h) === target);
  if (idx >= 0) return idx;
  idx = headers.findIndex((h) => {
    const n = normalizeHeader(h);
    return n.includes(target) || target.includes(n);
  });
  if (idx >= 0) return idx;
  const targetTokens = target.split(" ").filter(Boolean);
  if (targetTokens.length) {
    idx = headers.findIndex((h) => {
      const tokens = new Set(normalizeHeader(h).split(" ").filter(Boolean));
      return targetTokens.every((t) => tokens.has(t));
    });
    if (idx >= 0) return idx;
  }
  return -1;
}
