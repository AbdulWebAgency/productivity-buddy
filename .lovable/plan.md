# Sprint 3.0 — Architecture Refactor Plan

Goal: improve internal structure of the Excel engine and planner without changing any behaviour, prompts, UI, or data.

## Current state (measured)

- `src/lib/excel/engine.server.ts` — **1,069 lines**. Contains 8 op implementations (`opMerge`, `opDedupe`, `opDiff`, `opIntersection`, `opSummary`, `opMasterMerge`, `opBulkLookup`, `recalcFormulas`), workbook I/O (`readWorkbook`, `extractSheetMeta`, `sheetToGrid`, `writeGridToSheet`), styling helpers (`styleHeader`, `autoWidth`), header matching (`findColumnIndex`), and the `runPlan` orchestrator.
- Header normalization is duplicated in **4 files** with slightly different rules:
  - `engine.server.ts::findColumnIndex` — fuzzy resolver (normalize + substring + token subset)
  - `deterministic-plan.ts::norm` — planner-side normalization
  - `inspector.server.ts::norm` — inspector normalization
  - `suggestions.ts` — inline `.trim().toLowerCase()` for shared-header detection
- Op dispatch happens as a `switch` inside `runPlan`; `types.ts` defines the Zod discriminated union; `plan-repair.ts` has its own op-name allowlist; `friendly-errors.ts` has a per-op summarizer. Adding an op today requires touching 4–5 unrelated files.

## Target structure

```text
src/lib/excel/
  engine.server.ts         # thin: re-exports runPlan + public API only
  engine/
    runPlan.ts             # orchestrator; iterates registry
    registry.ts            # OperationRegistry: id -> { schema, run, summarize? }
    ops/
      diff.ts
      merge.ts
      intersection.ts
      masterMerge.ts
      bulkLookup.ts
      summary.ts
      dedupe.ts
      highlight.ts         # highlight_column + recalc live here
    shared/
      headers.ts           # canonical normalize() + resolveColumn() + scoreKey()
      workbook.ts          # readWorkbook, sheetToGrid, writeGridToSheet, extractSheetMeta, safeSheetName
      styling.ts           # styleHeader, autoWidth
  header-detection.ts      # unchanged (row-scan heuristic)
  deterministic-plan.ts    # switches to shared/headers
  intent.ts                # unchanged
  plan-repair.ts           # unchanged for now (op list stays in types.ts)
  types.ts                 # unchanged public shape
```

`src/lib/workspace/inspector.server.ts` and `suggestions.ts` also switch to `engine/shared/headers.ts`.

## Operation Registry shape

```ts
// engine/registry.ts
export type OpRunCtx = { files: EngineFile[]; wb: ExcelJS.Workbook; result: EngineResult };
export type OpHandler<Op extends PlanOp = PlanOp> = (op: Op, ctx: OpRunCtx) => Promise<void> | void;
export interface OperationDef<Op extends PlanOp = PlanOp> {
  id: Op["op"];
  run: OpHandler<Op>;
}
export const operationRegistry = new Map<PlanOp["op"], OperationDef>();
export function registerOp<Op extends PlanOp>(def: OperationDef<Op>): void;
```

Each `ops/*.ts` file self-registers on import; `runPlan.ts` imports the barrel `ops/index.ts` once, then dispatches `operationRegistry.get(op.op).run(op, ctx)`. Op schemas continue to live in `types.ts` — moving them would risk changing the discriminated union type surface and breaking plan-repair. This keeps the registry a pure runtime concern.

## Unified header utility (`engine/shared/headers.ts`)

Single source of truth exposing:
- `normalizeHeader(s: string): string` — the current `norm()` (trim, lowercase, collapse `._-`, collapse whitespace). Replaces the 4 duplicates.
- `resolveColumn(headers: string[], name: string): number` — current `findColumnIndex` (exact → normalized → substring → token-subset).
- `scoreKey(header: string): number` — current planner + inspector heuristic (KEY_HINTS-based).
- `KEY_HINTS` — the one shared array.

Behaviour is preserved by taking the existing implementations verbatim; no algorithmic change.

## Risks & mitigations

- **Risk: engine barrel splitting can accidentally drop a helper referenced across ops.** Mitigation: move file-by-file, keep `engine.server.ts` re-exporting the same public symbols so no importer changes, run typecheck after each move.
- **Risk: header hint list changes ranking behaviour.** Mitigation: use `deterministic-plan.ts`'s KEY_HINTS (longer list) as canonical; inspector's shorter list is a strict subset — verified.
- **Risk: self-registering ops rely on import side effects; tree-shaking or SSR-only import order could drop one.** Mitigation: single explicit barrel `ops/index.ts` imported by `runPlan.ts` (side-effect import), listed statically.
- **Risk: `.server.ts` filename guard.** Files under `engine/` that touch ExcelJS must keep the `.server.ts` suffix or live under an already-guarded path. Plan: name each op file `diff.ts` etc. but keep the whole `engine/` folder imported only from `engine.server.ts` (which is already server-only), so the import-protection chain is preserved.
- **Risk: circular imports between registry and ops.** Mitigation: `registry.ts` holds only the Map + `registerOp`; ops import types from `types.ts` and helpers from `shared/*`, never from `runPlan.ts`.

## Execution order (small, verifiable steps)

1. **Create `engine/shared/`** — extract `workbook.ts`, `styling.ts`, `headers.ts` from `engine.server.ts`. Replace originals with re-exports. Typecheck.
2. **Point planner/inspector/suggestions at `shared/headers.ts`** — delete their local `norm`/KEY_HINTS. Typecheck.
3. **Introduce `engine/registry.ts` + `engine/runPlan.ts`** — move current `runPlan` body in, but keep the `switch` temporarily. Re-export from `engine.server.ts`. Typecheck.
4. **Extract ops one at a time** in this order (safest → riskiest): `dedupe`, `summary`, `intersection`, `diff`, `merge`, `bulkLookup`, `masterMerge`, `highlight` (+ `recalc`). After each: register the op, delete its `switch` arm, typecheck.
5. **Delete the switch** once every op is registered; `runPlan` becomes a pure registry loop.
6. **Final pass**: shrink `engine.server.ts` to a re-export shim so all existing imports (`@/lib/excel/engine.server`) keep resolving unchanged.

Verification after every step: `tsgo` typecheck + a manual smoke of upload → diff → download in the preview to confirm identical output.

## Out of scope (explicitly)

- No prompt changes, no UI changes, no schema/migration changes, no new ops, no perf work, no test additions. `plan-repair.ts` and `intent.ts` are left alone.

Awaiting approval before step 1.
