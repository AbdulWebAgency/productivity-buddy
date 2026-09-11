// Explicit static operation table.
//
// Handlers are imported as *values* and referenced directly in OP_HANDLERS, so
// they can never be dropped by a bundler or lost to import ordering. Do NOT go
// back to side-effect registration (`import "./foo"`): package.json declares
// `"sideEffects": false`, which lets the bundler discard modules whose only
// purpose is a top-level register call.
import type { PlanOp } from "../../types";
import type { OpHandler } from "../registry";

import { dedupeHandler } from "./dedupe";
import { summaryHandler } from "./summary";
import { intersectionHandler } from "./intersection";
import { diffHandler } from "./diff";
import { mergeHandler } from "./merge";
import { bulkLookupHandler } from "./bulkLookup";
import { masterMergeHandler } from "./masterMerge";
import { highlightColumnHandler, recalcHandler } from "./highlight";

/** Every plan op id must have exactly one handler here (compiler-enforced). */
export const OP_HANDLERS: Record<PlanOp["op"], OpHandler> = {
  dedupe: dedupeHandler as OpHandler,
  summary: summaryHandler as OpHandler,
  intersection: intersectionHandler as OpHandler,
  diff: diffHandler as OpHandler,
  merge: mergeHandler as OpHandler,
  bulk_lookup: bulkLookupHandler as OpHandler,
  master_merge: masterMergeHandler as OpHandler,
  highlight_column: highlightColumnHandler as OpHandler,
  recalc: recalcHandler as OpHandler,
};

export function getOpHandler(id: PlanOp["op"]): OpHandler | undefined {
  return OP_HANDLERS[id];
}

export const REGISTERED_OPS = Object.keys(OP_HANDLERS) as PlanOp["op"][];
