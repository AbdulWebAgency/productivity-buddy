# Ledgerly → AI Office Workspace

Goal: keep the deterministic planner, ExcelJS engine, job/storage/auth layers exactly as they are. Rebuild the surface around them into a persistent, chat-driven workspace where uploaded workbooks stay attached, every operation produces a new version, and the AI acts as a coworker instead of a one-shot function.

## What stays untouched
- `src/lib/excel/engine.server.ts` — execution engine
- `src/lib/excel/deterministic-plan.ts` — planner-first logic
- `src/lib/excel/plan-repair.ts`, `types.ts` — schema + repair
- `src/lib/ai-gateway.server.ts` — gateway helper
- Supabase auth, storage buckets (`excel-uploads`, `excel-outputs`), RLS
- `runPlan`, `extractSheetMeta`, and all engine ops

## New concept: Workspace
A workspace is a long-lived container that owns files, a chat, and a chain of versions. Each user message can trigger a plan → execution → new version. The workbook of the latest version is what the AI reasons about next turn.

### Data model (one migration)
```
workspaces          (id, user_id, name, created_at, updated_at)
workspace_files     (id, workspace_id, storage_path, original_name, size_bytes,
                     sheet_meta jsonb, inspector jsonb, created_at)
workspace_messages  (id, workspace_id, role, parts jsonb, created_at)   -- AI SDK UIMessage[]
workspace_versions  (id, workspace_id, parent_version_id, label,
                     plan jsonb, stats jsonb, warnings jsonb,
                     output_path, output_name, size_bytes, created_at)
```
All tables: RLS scoped to `user_id` via workspace ownership, GRANTs to `authenticated` + `service_role`, plus `has_role(admin)` read policy. Existing `excel_jobs` tables stay for backward compatibility; new flow writes only to workspace tables.

Storage layout: `${userId}/${workspaceId}/inputs/*.xlsx`, `${userId}/${workspaceId}/versions/${versionId}.xlsx`.

### Server functions (all `createServerFn` + `requireSupabaseAuth`)
- `createWorkspace`, `listWorkspaces`, `getWorkspace`, `deleteWorkspace`
- `registerWorkspaceFiles` — reuses existing `extractSheetMeta`, additionally computes an **Inspector** report (rows, cols, blank rows, duplicate keys, formula cells, likely primary keys via `KEY_HINTS`, warnings). Stored on the file row.
- `suggestActions` — pure deterministic: given files + inspector, returns 4–6 chip suggestions ("Merge on Registration Number", "Find missing rows", "Dedupe", "Summary sheet").
- `planFromMessage` — runs deterministic planner first (existing). If ambiguous, returns a `needs_mapping` payload with candidate columns so UI can render the mapping picker; only escalates to Gemini when the deterministic planner returns `null` and no shared columns exist.
- `runPlanOnWorkspace` — resolves the *latest version's workbook* (or original inputs if v0) as engine inputs, runs `runPlan`, uploads output, inserts a new `workspace_versions` row linked to `parent_version_id`, returns human-readable summary.
- `restoreVersion` — creates a new version whose file is a copy of a prior version (non-destructive undo).
- `getVersionDownloadUrl` — signed URL, same pattern as `getDownloadUrl`.
- `chat` server route at `src/routes/api/chat.ts` — `streamText` with `openai/gpt-5.5`, tools: `propose_plan`, `run_plan`, `describe_file`, `list_versions`. Tools call the same server functions; model never touches the engine directly. Persists messages via `onFinish`.

### Frontend: `/app/w/$workspaceId`
Three-pane layout (desktop) collapsing to tabs on mobile:

```text
┌─────────────┬──────────────────────────┬──────────────┐
│  Files      │       Chat (AI SDK)      │  Versions    │
│  + Inspector│  transcript + composer   │  timeline    │
│  chips      │  suggested action chips  │  preview /   │
│             │  plan cards + run btn    │  download    │
└─────────────┴──────────────────────────┴──────────────┘
```

Built with AI Elements (`conversation`, `message`, `prompt-input`, `tool`, `shimmer`) per the chat-ui contract. Threaded route derives `workspaceId` from URL; chat `id` = `workspaceId`; messages persisted server-side.

Key UI components:
- **FilesPanel** — upload dropzone (reuses existing `registerJobFiles` upload flow, retargeted to workspace bucket path), lists files with inspector accordion.
- **InspectorCard** — per-file: sheets, rows, cols, headers, likely keys, warnings.
- **SuggestedActions** — chips above composer, click → prefills composer.
- **PlanCard** (tool render) — friendly bullet list of ops + estimated runtime + Run/Edit/Cancel buttons; no raw JSON.
- **ColumnMappingDialog** — rendered when `planFromMessage` returns `needs_mapping`; user picks per-file column; confirmation re-invokes `planFromMessage` with overrides.
- **VersionsPanel** — vertical timeline (v0 Original → v1 Merge → …). Each node: label, stats badges, Preview / Download / Restore.
- **PreviewSheet** — modal with tabs: Summary (ops, rows added/removed/changed, warnings), Sheets (first 50 rows of each output sheet via lightweight parse), Execution Details (collapsible raw logs).
- **ErrorToast** — maps engine/planner errors through a friendly-message helper. No raw Zod issues.

### Chat behavior rules
- System prompt tells the model: files + inspector + last version stats are ground truth; always call `propose_plan` before `run_plan`; never invent columns.
- `propose_plan` tool returns the friendly PlanCard payload — user must click Run to execute (`needsApproval` pattern via UI, not tool-level).
- After execution, assistant streams the transparent reasoning template ("I compared 2 workbooks using Registration Number. 312 matched…").
- Undo = "restore v{n-1}" natural language → tool call `restoreVersion`.

### Friendly error mapping (`src/lib/excel/errors.ts`)
Central function converting known engine/planner errors:
- missing shared column → "I couldn't find a shared column. Pick one: …"
- unknown op → "I don't know how to do that yet. Try: merge, compare, dedupe, summary."
- Zod issue → generic "I couldn't understand that request. Could you rephrase?"

### Migration path
- Keep `/app` (jobs list) and `/app/$jobId` routes working — mark them "Legacy" in nav.
- New default landing: `/app` shows Workspaces list + "New workspace" CTA.
- Old job creation route redirects to new workspace creation.

## Files created / edited

**New**
- `supabase/migrations/…_workspaces.sql`
- `src/lib/workspace.functions.ts` (all server fns above)
- `src/lib/workspace/inspector.server.ts` (analysis)
- `src/lib/workspace/suggestions.ts` (deterministic action suggestions)
- `src/lib/workspace/friendly-errors.ts`
- `src/routes/api/chat.ts` (streaming chat)
- `src/routes/_authenticated/app.w.$workspaceId.tsx`
- `src/routes/_authenticated/app.workspaces.tsx` (list + new)
- `src/components/workspace/{FilesPanel,InspectorCard,SuggestedActions,PlanCard,ColumnMappingDialog,VersionsPanel,PreviewSheet,ChatWindow}.tsx`
- `src/components/ai-elements/*` (installed via `ai-elements add`)

**Edited**
- `src/routes/_authenticated/app.index.tsx` — becomes Workspaces list
- `src/start.ts` — no change needed (auth attacher already wired)

## Out of scope (call out to user)
- Cross-session memory beyond the workspace (each workspace is its own memory unit — matches your "workspace memory" ask).
- Real-time collaboration / multi-user workspaces.
- Non-Excel formats (PDF, DOCX) — architecture leaves room but not implemented this pass.
- Charts inside preview modal (stats only; the engine's chart op still runs into the file).

## Verification before finishing
1. `tsgo` clean.
2. Create workspace → upload 2 xlsx → inspector shows sheets/keys.
3. "Find students missing from Main" → deterministic plan → PlanCard → Run → v1 appears in timeline → preview shows Missing Students sheet.
4. "Actually dedupe too" → new plan uses v1 as input → v2 created.
5. "Undo" → v3 = copy of v1.
6. Reload page → chat, files, versions all restore.
7. Ambiguous request → column mapping dialog appears instead of error.

Confirm and I'll build it in one pass.
