// Server functions for Excel automation. Client-safe module path.
import { createServerFn } from "@tanstack/react-start";
import { requireSupabaseAuth } from "@/integrations/supabase/auth-middleware";
import { z } from "zod";
import ExcelJS from "exceljs";
import { PlanSchema, SheetMetaSchema } from "./excel/types";

const MAX_FILE_BYTES = 25 * 1024 * 1024; // 25 MB
const MAX_JOB_BYTES = 150 * 1024 * 1024; // 150 MB
const MAX_FILES_PER_JOB = 10;

// --- Create job (row + return jobId for storage upload path). ---

export const createJob = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) =>
    z
      .object({
        name: z.string().min(1).max(200).default("Untitled job"),
        intent: z.string().max(4000).optional(),
        kind: z.enum(["merge", "dedupe", "diff", "format", "summary", "auto"]).default("auto"),
      })
      .parse(input),
  )
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { data: job, error } = await supabase
      .from("excel_jobs")
      .insert({
        user_id: userId,
        name: data.name,
        intent: data.intent ?? null,
        kind: data.kind,
        status: "draft",
      })
      .select("id")
      .single();
    if (error) throw new Error(error.message);
    return { jobId: job.id };
  });

// --- Register uploaded files (client uploads to storage, then calls this). ---

const RegisterFileSchema = z.object({
  jobId: z.string().uuid(),
  files: z
    .array(
      z.object({
        storagePath: z.string().min(1),
        originalName: z.string().min(1),
        sizeBytes: z.number().int().nonnegative().max(MAX_FILE_BYTES),
      }),
    )
    .min(1)
    .max(MAX_FILES_PER_JOB),
});

export const registerJobFiles = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => RegisterFileSchema.parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const totalBytes = data.files.reduce((s, f) => s + f.sizeBytes, 0);
    if (totalBytes > MAX_JOB_BYTES) {
      throw new Error(`Total upload exceeds ${MAX_JOB_BYTES / 1024 / 1024} MB limit.`);
    }

    // Verify job ownership.
    const { data: job, error: jobErr } = await supabase
      .from("excel_jobs")
      .select("id,user_id")
      .eq("id", data.jobId)
      .single();
    if (jobErr || !job) throw new Error("Job not found");
    if (job.user_id !== userId) throw new Error("Forbidden");

    // Load each file, extract sheet meta.
    const { extractSheetMeta } = await import("./excel/engine.server");
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");

    type Json = import("@/integrations/supabase/types").Json;
    const rows: {
      job_id: string;
      user_id: string;
      role: "input";
      storage_path: string;
      original_name: string;
      size_bytes: number;
      sheet_meta: Json;
    }[] = [];

    for (const f of data.files) {
      const { data: blob, error: dlErr } = await supabaseAdmin.storage.from("excel-uploads").download(f.storagePath);
      if (dlErr || !blob) throw new Error(`Download failed for ${f.originalName}: ${dlErr?.message}`);
      const buf = await blob.arrayBuffer();
      let meta: unknown = null;
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf);
        meta = extractSheetMeta(wb);
      } catch (e) {
        console.warn(`Could not parse ${f.originalName}:`, e);
        meta = { error: (e as Error).message };
      }
      rows.push({
        job_id: data.jobId,
        user_id: userId,
        role: "input",
        storage_path: f.storagePath,
        original_name: f.originalName,
        size_bytes: f.sizeBytes,
        sheet_meta: meta as Json,
      });
    }

    const { error: insErr } = await supabase.from("excel_job_files").insert(rows);
    if (insErr) throw new Error(insErr.message);
    return { count: rows.length };
  });

// --- List jobs ---

export const listJobs = createServerFn({ method: "GET" })
  .middleware([requireSupabaseAuth])
  .handler(async ({ context }) => {
    const { supabase, userId } = context;
    const { data, error } = await supabase
      .from("excel_jobs")
      .select("id,name,kind,status,created_at,completed_at,error,output_name")
      .eq("user_id", userId)
      .order("created_at", { ascending: false })
      .limit(50);
    if (error) throw new Error(error.message);
    return data;
  });

// --- Get single job with files ---

export const getJob = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ jobId: z.string().uuid() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { data: job, error } = await supabase
      .from("excel_jobs")
      .select("*")
      .eq("id", data.jobId)
      .eq("user_id", userId)
      .single();
    if (error || !job) throw new Error("Job not found");
    const { data: files, error: fErr } = await supabase
      .from("excel_job_files")
      .select("*")
      .eq("job_id", data.jobId)
      .order("created_at", { ascending: true });
    if (fErr) throw new Error(fErr.message);
    return { job, files };
  });

// --- Plan job via AI ---

export const planJob = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ jobId: z.string().uuid() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const apiKey = process.env.LOVABLE_API_KEY;
    if (!apiKey) throw new Error("AI gateway not configured");

    const { data: job } = await supabase.from("excel_jobs").select("id,intent,user_id").eq("id", data.jobId).single();
    if (!job || job.user_id !== userId) throw new Error("Job not found");

    const { data: files } = await supabase
      .from("excel_job_files")
      .select("original_name,sheet_meta")
      .eq("job_id", data.jobId)
      .eq("role", "input")
      .order("created_at", { ascending: true });
    if (!files || files.length === 0) throw new Error("Add at least one input file first.");

    await supabase.from("excel_jobs").update({ status: "planning" }).eq("id", data.jobId);

    try {
      const filesForAi = files.map((f, i) => {
        const parsed = SheetMetaSchema.safeParse(f.sheet_meta);
        return {
          index: i,
          name: f.original_name,
          sheets: parsed.success ? parsed.data.sheets : [],
        };
      });

      // 1. Try deterministic planner first — skip AI when unambiguous.
      const { tryDeterministicPlan } = await import("./excel/deterministic-plan");
      const det = tryDeterministicPlan(job.intent, filesForAi);
      if (det?.kind === "plan") {
        const plan = { ...det.plan, warnings: [...det.plan.warnings, ...det.notes] };
        await supabase
          .from("excel_jobs")
          .update({ ai_plan: plan, status: "queued", warnings: plan.warnings })
          .eq("id", data.jobId);
        return plan;
      }
      if (det?.kind === "needs_clarification") {
        const msg = `${det.reason} Shared columns detected: ${det.sharedColumns.join(", ") || "(none)"}.`;
        await supabase
          .from("excel_jobs")
          .update({ status: "failed", error: `Needs clarification: ${msg}` })
          .eq("id", data.jobId);
        throw new Error(msg);
      }

      // 2. Fall back to AI planner for ambiguous cases.
      const { createLovableAiGatewayProvider } = await import("./ai-gateway.server");
      const { generateText, NoObjectGeneratedError } = await import("ai");

      const columnCatalog = filesForAi.flatMap((f) =>
        f.sheets.flatMap((s) =>
          s.headers.map((h) => ({ fileIndex: f.index, fileName: f.name, sheet: s.name, column: h })),
        ),
      );

      const schemaGuide = `Return STRICT JSON matching this TypeScript type — no prose, no markdown, no code fences:
{
  "summary": string,
  "ops": Array<Op>,
  "columnMappings": Array<{ canonical: string, perFile: Array<{ fileIndex: number, column: string }> }>,
  "warnings": string[]
}
type Op =
  | { "op": "merge", "keyColumn": string, "strategy": "union"|"intersection", "highlightUnmatched": boolean }
  | { "op": "dedupe", "strategy": "key"|"full_row", "keyColumn"?: string }
  | { "op": "diff", "keyColumn": string, "fileAIndex": number, "fileBIndex": number }
  | { "op": "intersection", "keyColumn": string, "fileAIndex": number, "fileBIndex": number }
  | { "op": "summary", "includeCharts": boolean }
  | { "op": "recalc" }
  | { "op": "highlight_column", "column": string, "rule": "missing"|"duplicate"|"outlier" };

Rules:
- "op" values are lowercase.
- For "rows in B missing from A", use op "diff" with fileAIndex=A, fileBIndex=B — the engine emits a "Missing in A" sheet listing rows only in B.
- keyColumn / column MUST match a header string from the column catalog exactly (case-insensitive).
- fileAIndex / fileBIndex are integers (0-based) from files[].index.
- Never invent columns, and never emit file1/file2/outputSheetName-style keys.`;

      const gateway = createLovableAiGatewayProvider(apiKey);
      const model = gateway("google/gemini-2.5-flash");
      const promptPayload = JSON.stringify({
        intent: job.intent ?? "Auto-merge and dedupe reasonable data.",
        files: filesForAi,
        columnCatalog,
      });

      let rawText = "";
      try {
        const result = await generateText({
          model,
          system: `You are an Excel automation planner. Produce a strict deterministic plan.\n\n${schemaGuide}`,
          prompt: promptPayload,
        });
        rawText = result.text;
      } catch (e) {
        if (NoObjectGeneratedError.isInstance(e)) rawText = e.text ?? "";
        else throw e;
      }

      const { repairAndParsePlan } = await import("./excel/plan-repair");
      const { plan, repairs } = repairAndParsePlan(rawText, filesForAi);
      if (plan.ops.length === 0) {
        throw new Error(
          `AI planner produced no operations. Please rephrase your intent or specify the key column. Repairs: ${repairs.join("; ") || "none"}`,
        );
      }
      const mergedWarnings = [...(plan.warnings ?? []), ...repairs];

      await supabase
        .from("excel_jobs")
        .update({ ai_plan: plan, status: "queued", warnings: mergedWarnings })
        .eq("id", data.jobId);
      return { ...plan, warnings: mergedWarnings };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      await supabase
        .from("excel_jobs")
        .update({ status: "failed", error: `Planning failed: ${msg}` })
        .eq("id", data.jobId);
      throw new Error(msg);
    }
  });

// --- Run job ---

export const runJob = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ jobId: z.string().uuid(), plan: PlanSchema.optional() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { runPlan } = await import("./excel/engine.server");

    const { data: job } = await supabase.from("excel_jobs").select("*").eq("id", data.jobId).single();
    if (!job || job.user_id !== userId) throw new Error("Job not found");

    const plan = data.plan ?? PlanSchema.parse(job.ai_plan);
    if (!plan) throw new Error("No plan available. Run planJob first.");

    const { data: files } = await supabase
      .from("excel_job_files")
      .select("storage_path,original_name")
      .eq("job_id", data.jobId)
      .eq("role", "input")
      .order("created_at", { ascending: true });
    if (!files || files.length === 0) throw new Error("No input files.");

    await supabase.from("excel_jobs").update({ status: "running", error: null }).eq("id", data.jobId);

    try {
      const engineFiles = await Promise.all(
        files.map(async (f) => {
          const { data: blob, error } = await supabaseAdmin.storage.from("excel-uploads").download(f.storage_path);
          if (error || !blob) throw new Error(`Download failed: ${f.original_name}`);
          return { name: f.original_name, buffer: await blob.arrayBuffer() };
        }),
      );

      const result = await runPlan(engineFiles, plan);

      const outputName = `${job.name.replace(/[^a-z0-9-_ ]/gi, "").slice(0, 40) || "result"}.xlsx`;
      const outputPath = `${userId}/${data.jobId}/output/${outputName}`;
      const { error: upErr } = await supabaseAdmin.storage.from("excel-outputs").upload(outputPath, result.buffer, {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        upsert: true,
      });
      if (upErr) throw new Error(`Upload failed: ${upErr.message}`);

      type Json2 = import("@/integrations/supabase/types").Json;
      await supabase
        .from("excel_jobs")
        .update({
          status: "succeeded",
          output_path: outputPath,
          output_name: outputName,
          stats: result.stats as Json2,
          warnings: result.warnings as Json2,
          completed_at: new Date().toISOString(),
        })
        .eq("id", data.jobId);

      await supabase.from("excel_job_files").insert({
        job_id: data.jobId,
        user_id: userId,
        role: "output",
        storage_path: outputPath,
        original_name: outputName,
        size_bytes: result.buffer.length,
        sheet_meta: null,
      });

      return { ok: true, outputName };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      await supabase.from("excel_jobs").update({ status: "failed", error: msg }).eq("id", data.jobId);
      throw new Error(msg);
    }
  });

// --- Signed download URL ---

export const getDownloadUrl = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ jobId: z.string().uuid() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { data: job } = await supabase
      .from("excel_jobs")
      .select("output_path,output_name,user_id")
      .eq("id", data.jobId)
      .single();
    if (!job || job.user_id !== userId || !job.output_path) throw new Error("No output available");
    const { data: signed, error } = await supabaseAdmin.storage
      .from("excel-outputs")
      .createSignedUrl(job.output_path, 300, { download: job.output_name ?? "result.xlsx" });
    if (error || !signed) throw new Error(error?.message ?? "Signing failed");
    return { url: signed.signedUrl, name: job.output_name ?? "result.xlsx" };
  });

// --- Delete job ---

export const deleteJob = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ jobId: z.string().uuid() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { data: files } = await supabase.from("excel_job_files").select("storage_path,role").eq("job_id", data.jobId);
    if (files) {
      const uploads = files.filter((f) => f.role === "input").map((f) => f.storage_path);
      const outputs = files.filter((f) => f.role === "output").map((f) => f.storage_path);
      if (uploads.length) await supabaseAdmin.storage.from("excel-uploads").remove(uploads);
      if (outputs.length) await supabaseAdmin.storage.from("excel-outputs").remove(outputs);
    }
    const { error } = await supabase.from("excel_jobs").delete().eq("id", data.jobId).eq("user_id", userId);
    if (error) throw new Error(error.message);
    return { ok: true };
  });
