// Operation registry. Each op self-registers a handler; runPlan dispatches
// through the map instead of a hard-coded switch.
import type ExcelJS from "exceljs";
import type { PlanOp } from "../types";
import type { EngineFile } from "./types";
import type { SheetGrid } from "./shared/workbook";
import type { Projection, ProjectionMeta } from "./shared/projection";
import { projectGrid } from "./shared/projection";

/** Mutable per-run state shared across op handlers. */
export type EngineState = {
  currentGrid: SheetGrid;
  unmatched?: Set<number>;
  producedSheets: number;
};

export type ProjectionEvent = { sheet: string; meta: ProjectionMeta };

export type OpCtx = {
  files: EngineFile[];
  grids: SheetGrid[];
  outWb: ExcelJS.Workbook;
  state: EngineState;
  warnings: string[];
  opLogs: Record<string, unknown>[];
  started: number;
  /** Top-level output shaping requested by the plan (may be undefined). */
  projection?: Projection | null;
  /** Per-sheet projection outcomes; consumed by summarizers, never by ops. */
  projectionEvents: ProjectionEvent[];
};

export type OpHandler<Op extends PlanOp = PlanOp> = (op: Op, ctx: OpCtx) => void | Promise<void>;

const registry = new Map<PlanOp["op"], OpHandler>();

export function registerOp<Op extends PlanOp>(id: Op["op"], handler: OpHandler<Op>): void {
  registry.set(id, handler as OpHandler);
}

export function getOpHandler(id: PlanOp["op"]): OpHandler | undefined {
  return registry.get(id);
}

/**
 * Context-aware wrapper around `projectGrid`. Ops call this before writing a
 * user-facing sheet; diagnostic sheets (Summary, Missing keys, Not Found,
 * Changed rows) should skip projection and write the grid directly.
 *
 * The projection layer itself stays engine-agnostic — `alwaysKeep` is where
 * the op declares its own invariants (key column, "Sources", etc.).
 */
export function applyProjection(
  ctx: OpCtx,
  sheetName: string,
  grid: SheetGrid,
  opts: { alwaysKeep?: string[] } = {},
): SheetGrid {
  const projection = ctx.projection ?? null;
  const { grid: out, meta } = projectGrid(grid, projection, opts);

  const projectionWasRequested = !!(projection && (projection.columns?.length ?? 0) > 0);
  if (projectionWasRequested) {
    ctx.projectionEvents.push({ sheet: sheetName, meta });
    if (meta.unresolved.length > 0) {
      ctx.warnings.push(
        `Requested column${meta.unresolved.length === 1 ? "" : "s"} not found in "${sheetName}": ${meta.unresolved.join(", ")}. Skipped for that sheet.`,
      );
    }
    if (!meta.applied && meta.requested.length > 0 && meta.resolved.length === 0) {
      ctx.warnings.push(
        `None of the requested columns matched in "${sheetName}"; kept the full sheet instead.`,
      );
    }
  }
  return out;
}
