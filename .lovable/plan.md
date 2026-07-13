
# Excel Automation SaaS — MVP Plan

## Recommendation: Pure JS backend now, Python-ready later

Given your scale (2–10 files, 1–15 MB, 500–20k rows) and audience (teachers, small offices), a **pure-JS backend inside Lovable** is the right MVP. It stays on one platform, avoids DevOps, and comfortably handles those file sizes. We'll design the job/worker boundary so you can later swap in an external Python service for enterprise-scale workloads without touching the UI.

### Why not Python now
Lovable's server runtime is Cloudflare Workers-style — no Python, no `pandas`/`openpyxl`, no LibreOffice, no native binaries. Adding Python means hosting a second service elsewhere, which is real ops overhead you don't need at MVP scale.

### Platform limits to design around
- **~15 MB / 20k rows per file** is comfortable in JS; beyond ~50 MB or 100k+ rows we'd want the Python path.
- **No true formula recalculation** in the Worker. We use HyperFormula for common formulas (SUM, IF, VLOOKUP, arithmetic, text, date). Rare/exotic formulas keep the formula string; Excel recalculates on open.
- **Per-request CPU/time limits.** Large jobs run as async background jobs, not inside the upload request.
- **No persistent disk.** Files live in Lovable Cloud Storage; workers stream them.

---

## MVP Scope (deterministic ops + AI reasoning layer)

Deterministic engine (JS, no LLM in the hot path):
1. Merge multiple workbooks on a user-chosen key column (e.g. Registration Number)
2. Auto-detect matching columns across files (header similarity + sample-value overlap)
3. Preserve formatting, headers, column widths, merged cells, number formats
4. Highlight unmatched / missing rows (fill color + a "Status" column)
5. Detect and remove duplicates (by key or full-row hash)
6. Compare worksheets and produce a diff report (added / removed / changed rows)
7. Apply conditional formatting rules
8. Generate summary sheet + basic charts (bar, line, pie)
9. Recalculate supported formulas via HyperFormula; leave the rest as formula strings
10. Export a clean downloadable `.xlsx`

AI reasoning layer (Lovable AI, server-side only):
- Interprets the user's natural-language request → produces a **typed job plan** (JSON) listing which deterministic ops to run and with what parameters
- Explains broken formulas and proposes a fix (returns the corrected formula string; user confirms before it's written)
- Suggests column mappings when auto-detect is ambiguous
- Never touches cell data directly — it only picks and parameterizes deterministic ops

---

## Architecture

```text
Browser ──uploads──▶ Lovable Cloud Storage (bucket: excel-uploads/{user}/{jobId}/)
   │                         │
   │                         ▼
   │              excel_jobs row (queued)
   │                         │
   ▼                         ▼
Job UI (poll/realtime) ◀── Server fn: processJob (reads inputs, runs engine, writes output)
                                     │
                                     ├─ ExcelJS  → read/write .xlsx, preserve styles
                                     ├─ HyperFormula → recalc supported formulas
                                     └─ Lovable AI → plan + formula-fix reasoning
                                     │
                                     ▼
                         Lovable Cloud Storage (excel-outputs/{user}/{jobId}/result.xlsx)
```

Key boundary: `processJob` is called via a server function and does all the heavy lifting behind one interface. Later, that function's body can be replaced with a `fetch()` to an external Python worker — the UI, storage, and job table stay identical.

---

## Data model (Lovable Cloud / Postgres)

- `excel_jobs`: `id`, `user_id`, `kind` (merge / dedupe / diff / format / summary), `status` (queued/running/succeeded/failed), `params jsonb`, `ai_plan jsonb`, `error text`, `output_path text`, timestamps
- `excel_job_files`: `id`, `job_id`, `role` (input/output), `storage_path`, `original_name`, `size_bytes`, `sheet_meta jsonb`
- RLS: user can only see their own jobs/files. Roles table (`user_roles` + `has_role`) for future admin views.
- Storage buckets: `excel-uploads` (private), `excel-outputs` (private, signed URLs for download).

---

## User flow (MVP)

1. **Upload** — drag & drop 2–10 files, shown in a job composer.
2. **Describe intent** — free-text ("Merge these by Registration Number, remove duplicates, highlight rows missing in file 2").
3. **AI plan preview** — server function calls Lovable AI with the file headers + user intent, returns a structured plan (ops list + column mappings + confidence). User can tweak mappings inline.
4. **Run** — job is queued; UI subscribes/polls status.
5. **Result** — download the `.xlsx`, view a summary panel (rows merged, duplicates removed, unmatched count, warnings).

---

## Technical Details

### Packages
- `exceljs` — read/write `.xlsx` with styles, merged cells, images, charts, conditional formatting
- `hyperformula` — formula recalculation engine (MIT/GPL dual-licensed; MIT for typical SaaS use)
- `zod` — validate AI plan + all server-function inputs
- Lovable AI Gateway via `ai` + `@ai-sdk/openai-compatible` (already the platform default)

### Server surface (TanStack Start)
- `src/lib/excel.functions.ts`
  - `createJob({ kind, fileMeta[], intent })` → returns `jobId`, uploads via signed URLs
  - `getJob(jobId)` / `listJobs()`
  - `planJob(jobId)` — calls AI, stores `ai_plan`
  - `runJob(jobId)` — orchestrates engine ops, writes output
  - `getDownloadUrl(jobId)`
- `src/lib/excel/engine/*.server.ts` (server-only helpers, not client-imported)
  - `merge.ts`, `dedupe.ts`, `diff.ts`, `format.ts`, `summary.ts`, `formulas.ts`
- All protected with `requireSupabaseAuth`; storage access uses `supabaseAdmin` inside handlers only.

### Auth
- Email + password to start (Lovable Cloud native). Google sign-in optional. Roles table wired but only `user` role at MVP.

### Async job execution
- MVP: `runJob` executes inline within the server function (Workers allow ~30s of CPU per request on standard tiers; 15 MB files fit). If we hit ceilings, we split into chunked steps or move to the external-worker path.
- UI polls `getJob` every 2s while `status in ('queued','running')`.

### Formula handling
- On read: preserve all formula strings.
- Recalc pass: attempt via HyperFormula; on unsupported functions or errors, keep original string and flag in output summary.
- On open in Excel, unrecalculated cells recompute automatically.

### Formatting preservation
- ExcelJS preserves styles, number formats, merged cells, column widths, row heights, conditional formatting rules, and basic chart definitions. Complex pivots and some chart types have partial support — we surface a warning in the job result rather than silently dropping them.

### Security
- Signature-verified signed upload URLs (Lovable Cloud Storage)
- Zod validation on every server-fn input
- Per-file size cap (25 MB) and per-job total cap (150 MB) enforced server-side
- AI plan is a strict Zod schema — the LLM cannot invoke arbitrary operations

### Later: swap-in Python worker
When you need `pandas`/`openpyxl`/LibreOffice at scale:
- Stand up a Python service (FastAPI on Fly.io/Render/Modal)
- Replace `runJob`'s engine calls with an authenticated `fetch` to the Python service (HMAC-signed request; the worker reads input from and writes output back to the same Lovable Cloud Storage bucket)
- No UI, DB schema, or auth changes required

---

## Out of scope for MVP (explicit)
- Pivot tables round-trip fidelity, advanced chart types (waterfall, radar), macros/VBA
- Files > 25 MB per file or > 100k rows per sheet (queue-and-warn)
- Real-time collaborative editing
- Payments / plans (add later once usage patterns are clear)

---

## Deliverables in build phase
1. Enable Lovable Cloud; migrations for `excel_jobs`, `excel_job_files`, roles; storage buckets + RLS
2. Auth (email/password) + `_authenticated` route layout
3. Upload composer + job list + job detail pages
4. Engine modules (merge / dedupe / diff / format / summary / formulas)
5. AI plan server fn + plan-preview UI with editable column mappings
6. Job runner + download flow
7. Design pass (I'll ask for a direction before styling)

Approve this and I'll switch to build mode and start with Cloud enablement + schema.
