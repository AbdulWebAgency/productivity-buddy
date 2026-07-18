# Master Sheet + Bulk Lookup

Add two new deterministic actions to the existing Workspace — no redesign, no AI dependency, reuses the current engine, planner, versioning, and download flow.

## What the user gets

Two new items appear in the workspace's suggestion chip row (and as commands the chat also recognises):

1. **Create Master Sheet** — one-click multi-file merge (2–10 files) → downloadable Master workbook with summary + missing sheets.
2. **Bulk Lookup** — paste a list of IDs / names / emails, pick a workspace file or the latest master, get a filtered result workbook.

No changes to the three-pane layout, colors, typography, auth, or version-history UI.

## Create Master Sheet

Flow (deterministic, AI is only a fallback for header confidence):

1. User clicks the chip or types "create master sheet" / "merge all files".
2. New planner path `planMaster()` runs:
   - Inspects each file's first sheet headers (already cached in `workspace_files.inspector`).
   - Builds candidate merge keys by scoring header names against known aliases:
     `student id ~ id ~ reg no ~ registration number ~ roll no ~ roll number ~ email ~ email address ~ mobile ~ phone`.
     Score = alias match (3) + appears in ≥2 files (2) + high uniqueness in each file (2) + ID-shaped values (1).
   - **One strong key** (score ≥ 6 and unique in every file it appears in): auto-select, produce plan, run.
   - **Multiple candidates within 1 point**: emit a `clarify` message with chips ("Use Reg No", "Use Email", …). Reuses the existing clarify pipeline in `sendMessage`.
   - **No shared key**: emit a friendly error explaining which files lack a shared key.
3. Executes a new engine op `master_merge` (see below) → writes a single `.xlsx` workspace version labelled `Master sheet (key: <col>)`.

Options exposed as follow-up chat commands (deterministic parse — no AI):
- "use left join" / "inner join" / "append" — default is full outer.
- "keep first" / "keep latest" / "merge duplicates" — default is keep first, with a warning row count.

## Bulk Lookup

1. User clicks "Bulk Lookup" chip → composer is pre-filled with:
   ```
   Bulk lookup in <filename>:
   <paste IDs, names, or emails here — one per line>
   ```
2. `sendMessage` deterministically detects `bulk lookup` intent, parses the target filename (defaults to the newest master version if omitted) and the list of query values.
3. New engine op `bulk_lookup` scans the chosen workbook's first sheet for matches across all "key-ish" columns (id, email, name, roll, reg) — case- and whitespace-insensitive, substring match for names.
4. Produces a workspace version `Bulk lookup (<n> queries)` — a workbook with:
   - `Results` sheet: matched rows with source columns preserved + `Matched On` column.
   - `Not Found` sheet: queries with no match.

## Engine ops (extend, don't duplicate)

`src/lib/excel/engine.server.ts`:

- `opMasterMerge({ keyColumn, joinType, dupeStrategy })` — builds one wide table by outer/inner/left-joining every file's first sheet on the resolved key column (using existing fuzzy column resolver). Emits:
  - `Master` sheet — union of all columns, prefixed with file short-name on collision.
  - `Missing in <file>` sheets — key values present in the union but not in that file.
  - `Merge Summary` sheet — total rows, matched, per-file missing counts, duplicate count, key used, join type.
- `opBulkLookup({ file, queries })` — reads the target file, indexes candidate key columns, returns matched rows.

Both go through `runPlan`, so the existing per-op logging, timing, stats, warnings, and "no output" guard rails already apply. Downloads, version rows, and stats JSON reuse the current pipeline.

## Types + plan schema

`src/lib/excel/types.ts` — add two ops to `OpSchema`:
```ts
{ op: 'master_merge', keyColumn, joinType: 'outer'|'inner'|'left'|'append', dupeStrategy: 'first'|'latest'|'merge' }
{ op: 'bulk_lookup', fileIndex, queries: string[] }
```
`plan-repair.ts` — recognise these ops; if fields missing, fill defaults (outer, first).

## Planner glue

`src/lib/excel/deterministic-plan.ts`:
- Add `planMaster(files)` and `planBulkLookup(files, text)`.
- Add intent classifier keywords: `master sheet`, `combine all files`, `merge everything`, `bulk lookup`, `bulk search`, `lookup these`.
- Keep existing diff/intersection/dedupe intents untouched.

## Suggestions

`src/lib/workspace/suggestions.ts` — when `files.length >= 2` push:
- `Create Master Sheet` → prompt `Create a master sheet by merging all uploaded files.`
- `Bulk Lookup` (always when at least one file) → prompt `Bulk lookup in <newest file>:\n`

The chips already render in the composer; nothing else in the UI changes.

## Out of scope

No AI merging, no OCR, no PDFs, no auth changes, no payment flows, no MCP work, no redesign.

## Files touched

- `src/lib/excel/types.ts` — 2 new op schemas
- `src/lib/excel/engine.server.ts` — `opMasterMerge`, `opBulkLookup`, wire into `runPlan`
- `src/lib/excel/deterministic-plan.ts` — new intents + planners
- `src/lib/excel/plan-repair.ts` — repair defaults for new ops
- `src/lib/excel/intent.ts` — `master_merge`, `bulk_lookup` intent labels
- `src/lib/workspace.functions.ts` — describePlan / planLabel / summarizeRun cases
- `src/lib/workspace/suggestions.ts` — two new chips
- `src/routes/_authenticated/app.w.$workspaceId.tsx` — render new op summaries in `PlanCard` (2 extra `op.op === …` branches, no layout changes)

Everything else — versions panel, download, inspector, chat rendering, auth, storage buckets — is untouched.
