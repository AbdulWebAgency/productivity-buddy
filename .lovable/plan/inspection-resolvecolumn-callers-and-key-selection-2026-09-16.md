# Inspection: resolveColumn callers and key selection

No files changed. Findings below, plus a recommendation for a separate key resolver.

## 1. Callers of resolveColumn

All import from `src/lib/excel/engine/shared/headers.ts`.

| Caller | Line | Input | Expectation |
| --- | --- | --- | --- |
| `ops/diff.ts` | 31-32 | plan `keyColumn` vs each file's headers | exactly one correct key column per side; throws if -1 |
| `ops/intersection.ts` | 14-15 | plan `keyColumn` vs A and B | same; throws if -1 |
| `ops/merge.ts` | 18 | `keyColumn` per grid | per-file key index; -1 tolerated (row falls back) |
| `ops/masterMerge.ts` | 20 | `keyColumn` per grid | key index per file for the join |
| `ops/dedupe.ts` | 13 | optional `keyColumn` | -1 silently degrades to full-row dedupe |
| `ops/highlight.ts` | 72 | display column name | index of a column to style; cosmetic |
| `shared/projection.ts` | 79, 93 | user-requested output column names, plus `alwaysKeep` | best-effort display-column match; -1 recorded as unresolved |
| `deterministic-plan.ts` | 77 | projection candidates vs union of headers | boolean "does any candidate look like a real column" gate |

Two distinct expectations are being served by one function:
- **Key resolution** (diff, intersection, merge, masterMerge, dedupe): must be exact and identity-preserving; a wrong match silently produces wrong joins.
- **Display/projection resolution** (projection, highlight, planner gate): loose matching is desirable and failures are harmless.

## 2. Is the current fuzzy matching safe?

`resolveColumn` tries, in order: exact normalized equality, substring in either direction, then token-subset. Safe for projection/highlight. Not safe as a key resolver:

- **Substring either direction is over-permissive.** Target `"id"` matches `"Candidate ID"`, `"Grid Ref"`, `"Valid Until"` — whichever comes first in header order. `"name"` matches `"Filename"`, `"Surname"`.
- **First-match wins, no scoring.** With `"Student ID"` and `"Student ID (old)"` present, the answer depends on column order, not similarity. There is no ambiguity signal to the caller.
- **Token-subset ignores extra tokens.** Target `"reg no"` matches `"Reg No Verified Flag"`.
- **Asymmetric across files.** diff/intersection/merge resolve the same key name independently per file, so file A can land on `Student ID` and file B on `Old Student ID` with no cross-check. This is the highest-severity risk: the operation succeeds and returns plausible but wrong rows.
- **No type/uniqueness check.** A column resolved as a key is never verified against the inspector's uniqueness data.

## 3. Correctness risks in the current key selection

`commonColumns` / `pickBestKey` / `scoreKey` in `deterministic-plan.ts`:

- `commonColumns` uses strict normalized equality only, while the engine resolves fuzzily. So the planner can decide "no shared columns → ask the user", then the engine happily fuzzy-matches something. Planner and engine disagree about what a match is.
- `scoreKey` scores by longest KEY_HINTS substring, so `"id"` (2) loses to `"number"` (6) — `"Invoice Number"` beats `"Student ID"`. Hint length is a poor proxy for key quality.
- Score ties return `key: null` and force a clarification even when uniqueness data would settle it immediately.
- `pickBestKey` returns a user-supplied `preferredKey` verbatim even when it matches no header at all; the failure then surfaces from deep inside the op as a thrown error rather than as a clarification.
- Key choice ignores everything the inspector already computed (`likelyKeys`, uniqueness, duplicate counts, primary sheet), so the planner is strictly less informed than the inspection layer.
- Keys compare via `String(...).trim().toLowerCase()` in the ops. Numeric IDs read as `1001` vs text `"1001"` still match, but `1001.0` or Excel-formatted values do not — no normalization of numeric/date key values.

## 4. Recommendation: a separate `resolveKeyColumn`

Warranted. The two use cases have opposite failure preferences: projection should guess, keys should refuse.

**Where it sits:** a new module `src/lib/excel/engine/shared/key-resolution.ts`, next to `headers.ts`, depending on it. `headers.ts` keeps `normalizeHeader`, `scoreKey`, `resolveColumn` unchanged so projection/highlight/planner-gate behaviour is untouched.

Proposed shape:

```ts
type KeyResolution =
  | { status: "resolved"; index: number; header: string; confidence: number }
  | { status: "ambiguous"; candidates: { index: number; header: string; score: number }[] }
  | { status: "not_found" };

resolveKeyColumn(headers: string[], name: string): KeyResolution
```

Rules: exact normalized match wins outright; otherwise score all headers (token overlap ratio, penalise extra tokens, require full-token containment rather than raw substring) and return `ambiguous` when the top two are close, instead of picking one. Never return a bare index.

**Adoption points**, in order:
1. `ops/diff.ts`, `ops/intersection.ts` — replace both `resolveColumn` calls; also add a cross-file agreement check so A and B must land on equivalent headers before the op runs.
2. `ops/merge.ts`, `ops/masterMerge.ts` — same per-grid resolution, with per-file resolution reported in the op log.
3. `ops/dedupe.ts` — resolve, and on `ambiguous`/`not_found` warn rather than silently switching to full-row dedupe.
4. `deterministic-plan.ts` — `commonColumns` uses the same resolver so planner and engine agree; `pickBestKey` consults inspector uniqueness before falling back to `scoreKey`, and validates `preferredKey` against real headers before returning it.

Projection, highlight and the planner's projection gate keep using `resolveColumn`.

## Suggested first step

Introduce `key-resolution.ts` with unit tests over adversarial header sets (`Student ID` / `Old Student ID`, `Name` / `Filename`, `Reg No` / `Reg No Verified`), then adopt in diff and intersection only, verifying against a real workspace run before touching merge/master_merge.
