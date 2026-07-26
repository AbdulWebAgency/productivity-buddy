// Operation registry. Each op self-registers a handler; runPlan dispatches
// through the map instead of a hard-coded switch.
import type ExcelJS from "exceljs";
import type { PlanOp } from "../types";
import type { EngineFile } from "./types";
import type { SheetGrid } from "./shared/workbook";

/** Mutable per-run state shared across op handlers. */
export type EngineState = {
  currentGrid: SheetGrid;
  unmatched?: Set<number>;
  producedSheets: number;
};

export type OpCtx = {
  files: EngineFile[];
  grids: SheetGrid[];
  outWb: ExcelJS.Workbook;
  state: EngineState;
  warnings: string[];
  opLogs: Record<string, unknown>[];
  started: number;
};

export type OpHandler<Op extends PlanOp = PlanOp> = (op: Op, ctx: OpCtx) => void | Promise<void>;

const registry = new Map<PlanOp["op"], OpHandler>();

export function registerOp<Op extends PlanOp>(id: Op["op"], handler: OpHandler<Op>): void {
  registry.set(id, handler as OpHandler);
}

export function getOpHandler(id: PlanOp["op"]): OpHandler | undefined {
  return registry.get(id);
}
