// Engine-agnostic output projection.
//
// The projection layer narrows/reorders columns of a SheetGrid based on a
// user-supplied `Projection` object. It knows nothing about diff, merge,
// intersection or any specific operation — invariants like "always keep the
// key column" are the responsibility of the individual op handlers, which
// pass those columns via `opts.alwaysKeep`.
//
// Sprint 3.1 only implements `columns`. The `Projection` type is defined as
// an object (not a bare `string[]`) so future capabilities such as
// excludeColumns, renameColumns, reorderColumns, or computedColumns can be
// added without another schema redesign.

import type { CellValue, SheetGrid } from "./workbook";
import { resolveColumn } from "./headers";

export type Projection = {
  /** Ordered list of requested output columns (user-friendly names). */
  columns?: string[];
  // Reserved for future extension — intentionally not yet implemented:
  // excludeColumns?: string[];
  // renameColumns?: Record<string, string>;
  // reorderColumns?: string[];
  // computedColumns?: Array<{ name: string; formula: string }>;
};

export type ProjectionMeta = {
  /** True when the returned grid was actually narrowed. */
  applied: boolean;
  /** Column names as the caller requested (post-trim, pre-resolve). */
  requested: string[];
  /** Header names as they appear in the source grid, in requested order. */
  resolved: string[];
  /** Requested names with no fuzzy match against the source headers. */
  unresolved: string[];
  /** Columns force-included by the op handler on top of `resolved`. */
  alwaysKept: string[];
};

export function hasProjection(p?: Projection | null): boolean {
  if (!p) return false;
  return Array.isArray(p.columns) && p.columns.some((c) => typeof c === "string" && c.trim().length > 0);
}

/**
 * Narrow `grid` to the requested columns.
 *
 * - Preserves the order the caller requested.
 * - Uses the shared fuzzy `resolveColumn`, so per-file header variance is
 *   handled the same way it is everywhere else in the engine.
 * - `opts.alwaysKeep` columns are appended after the requested ones and only
 *   added if they exist in the source grid and aren't already included.
 * - Safety net: if no requested column resolves AND no alwaysKeep column
 *   exists, returns the untouched grid (never an empty sheet) and reports
 *   `applied: false` with the unresolved names.
 */
export function projectGrid(
  grid: SheetGrid,
  projection?: Projection | null,
  opts: { alwaysKeep?: string[] } = {},
): { grid: SheetGrid; meta: ProjectionMeta } {
  const noopMeta: ProjectionMeta = {
    applied: false,
    requested: [],
    resolved: [],
    unresolved: [],
    alwaysKept: [],
  };
  if (!hasProjection(projection)) return { grid, meta: noopMeta };

  const requested = projection!.columns!.map((c) => c.trim()).filter(Boolean);

  const indices: number[] = [];
  const seen = new Set<number>();
  const resolved: string[] = [];
  const unresolved: string[] = [];

  for (const name of requested) {
    const idx = resolveColumn(grid.headers, name);
    if (idx < 0) {
      unresolved.push(name);
      continue;
    }
    if (seen.has(idx)) continue;
    seen.add(idx);
    indices.push(idx);
    resolved.push(grid.headers[idx]);
  }

  const alwaysKept: string[] = [];
  for (const name of opts.alwaysKeep ?? []) {
    if (!name) continue;
    const idx = resolveColumn(grid.headers, name);
    if (idx < 0 || seen.has(idx)) continue;
    seen.add(idx);
    indices.push(idx);
    alwaysKept.push(grid.headers[idx]);
  }

  if (indices.length === 0) {
    // Nothing resolved and no alwaysKeep hit — surface the failure via meta
    // but return the original grid so the caller doesn't ship an empty sheet.
    return {
      grid,
      meta: { applied: false, requested, resolved: [], unresolved, alwaysKept: [] },
    };
  }

  const headers = indices.map((i) => grid.headers[i]);
  const rows: CellValue[][] = grid.rows.map((r) => indices.map((i) => r[i] ?? null));
  return {
    grid: { headers, rows },
    meta: { applied: true, requested, resolved, unresolved, alwaysKept },
  };
}
