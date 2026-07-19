// Workspace server functions: persistent AI workspaces with chat, files, versions.
import { createServerFn } from "@tanstack/react-start";
import { requireSupabaseAuth } from "@/integrations/supabase/auth-middleware";
import { z } from "zod";
import { PlanSchema, type Plan } from "./excel/types";

const MAX_FILE_BYTES = 25 * 1024 * 1024;
const MAX_FILES_PER_WORKSPACE = 10;

// ---------- Workspace CRUD ----------

export const createWorkspace = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) =>
    z.object({ name: z.string().min(1).max(200).default("Untitled workspace") }).parse(input),
  )
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { data: row, error } = await supabase
      .from("workspaces")
      .insert({ user_id: userId, name: data.name })
      .select("id")
      .single();
    if (error) throw new Error(error.message);
    // Seed with a welcome assistant message.
    await supabase.from("workspace_messages").insert({
      workspace_id: row.id,
      user_id: userId,
      role: "assistant",
      content:
        "Hi — I'm Productivity Buddy. Drop one or more spreadsheets on the left and I'll take a look, then we can chat about what to do with them.",

    });
    return { id: row.id };
  });

export const listWorkspaces = createServerFn({ method: "GET" })
  .middleware([requireSupabaseAuth])
  .handler(async ({ context }) => {
    const { supabase, userId } = context;
    const { data, error } = await supabase
      .from("workspaces")
      .select("id,name,created_at,updated_at")
      .eq("user_id", userId)
      .order("updated_at", { ascending: false })
      .limit(100);
    if (error) throw new Error(error.message);
    return data;
  });

export const getWorkspace = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ workspaceId: z.string().uuid() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { data: ws } = await supabase
      .from("workspaces")
      .select("*")
      .eq("id", data.workspaceId)
      .eq("user_id", userId)
      .single();
    if (!ws) throw new Error("Workspace not found");
    const [{ data: files }, { data: messages }, { data: versions }] = await Promise.all([
      supabase
        .from("workspace_files")
        .select("id,original_name,size_bytes,storage_path,sheet_meta,inspector,created_at")
        .eq("workspace_id", data.workspaceId)
        .order("created_at", { ascending: true }),
      supabase
        .from("workspace_messages")
        .select("id,role,content,tool_data,created_at")
        .eq("workspace_id", data.workspaceId)
        .order("created_at", { ascending: true }),
      supabase
        .from("workspace_versions")
        .select("id,version_number,label,plan,stats,warnings,output_name,size_bytes,created_at")
        .eq("workspace_id", data.workspaceId)
        .order("version_number", { ascending: true }),
    ]);
    return { workspace: ws, files: files ?? [], messages: messages ?? [], versions: versions ?? [] };
  });

export const deleteWorkspace = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => z.object({ workspaceId: z.string().uuid() }).parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { data: files } = await supabase
      .from("workspace_files")
      .select("storage_path")
      .eq("workspace_id", data.workspaceId);
    const { data: versions } = await supabase
      .from("workspace_versions")
      .select("output_path")
      .eq("workspace_id", data.workspaceId);
    const inPaths = (files ?? []).map((f) => f.storage_path).filter(Boolean);
    const outPaths = (versions ?? []).map((v) => v.output_path).filter(Boolean) as string[];
    if (inPaths.length) await supabaseAdmin.storage.from("excel-uploads").remove(inPaths);
    if (outPaths.length) await supabaseAdmin.storage.from("excel-outputs").remove(outPaths);
    const { error } = await supabase
      .from("workspaces")
      .delete()
      .eq("id", data.workspaceId)
      .eq("user_id", userId);
    if (error) throw new Error(error.message);
    return { ok: true };
  });

// ---------- Files ----------

const RegisterFilesSchema = z.object({
  workspaceId: z.string().uuid(),
  files: z
    .array(
      z.object({
        storagePath: z.string().min(1),
        originalName: z.string().min(1),
        sizeBytes: z.number().int().nonnegative().max(MAX_FILE_BYTES),
      }),
    )
    .min(1)
    .max(MAX_FILES_PER_WORKSPACE),
});

export const registerWorkspaceFiles = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) => RegisterFilesSchema.parse(input))
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { data: ws } = await supabase
      .from("workspaces")
      .select("id,user_id")
      .eq("id", data.workspaceId)
      .single();
    if (!ws || ws.user_id !== userId) throw new Error("Workspace not found");

    const ExcelJS = (await import("exceljs")).default;
    const { extractSheetMeta } = await import("./excel/engine.server");
    const { inspectWorkbook } = await import("./workspace/inspector.server");
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");

    type Json = import("@/integrations/supabase/types").Json;
    const rows: {
      workspace_id: string;
      user_id: string;
      storage_path: string;
      original_name: string;
      size_bytes: number;
      sheet_meta: Json;
      inspector: Json;
    }[] = [];

    for (const f of data.files) {
      const { data: blob, error } = await supabaseAdmin.storage
        .from("excel-uploads")
        .download(f.storagePath);
      if (error || !blob) throw new Error(`Download failed for ${f.originalName}: ${error?.message}`);
      const buf = await blob.arrayBuffer();
      let meta: unknown = null;
      let inspector: unknown = null;
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf);
        meta = extractSheetMeta(wb);
        inspector = inspectWorkbook(wb);
      } catch (e) {
        meta = { error: (e as Error).message };
      }
      rows.push({
        workspace_id: data.workspaceId,
        user_id: userId,
        storage_path: f.storagePath,
        original_name: f.originalName,
        size_bytes: f.sizeBytes,
        sheet_meta: meta as Json,
        inspector: inspector as Json,
      });
    }
    const { error: insErr } = await supabase.from("workspace_files").insert(rows);
    if (insErr) throw new Error(insErr.message);
    await supabase
      .from("workspaces")
      .update({ updated_at: new Date().toISOString() })
      .eq("id", data.workspaceId);

    // Auto-greeting: after every upload batch, post a warm inspection summary
    // so the assistant behaves like a coworker who actually looked at the files.
    try {
      const { data: allFiles } = await supabase
        .from("workspace_files")
        .select("original_name,sheet_meta,inspector")
        .eq("workspace_id", data.workspaceId)
        .order("created_at", { ascending: true });
      const { buildInspectionGreeting } = await import("./workspace/ai-conversation.server");
      const ctx = (allFiles ?? []).map((f, i) => ({
        index: i,
        name: f.original_name,
        sheets:
          (f.sheet_meta as { sheets?: { name: string; headers: string[] }[] } | null)?.sheets ?? [],
        inspector: f.inspector as never,
      }));
      const greeting = buildInspectionGreeting(ctx);
      await supabase.from("workspace_messages").insert({
        workspace_id: data.workspaceId,
        user_id: userId,
        role: "assistant",
        content: greeting,
      });
    } catch {
      // Non-fatal: file registration succeeded.
    }
    return { count: rows.length };
  });


export const removeWorkspaceFile = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) =>
    z.object({ workspaceId: z.string().uuid(), fileId: z.string().uuid() }).parse(input),
  )
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { data: f } = await supabase
      .from("workspace_files")
      .select("storage_path,user_id")
      .eq("id", data.fileId)
      .single();
    if (!f || f.user_id !== userId) throw new Error("File not found");
    if (f.storage_path) await supabaseAdmin.storage.from("excel-uploads").remove([f.storage_path]);
    await supabase.from("workspace_files").delete().eq("id", data.fileId);
    return { ok: true };
  });

// ---------- Chat / planning ----------
//
// The conversation layer is LLM-first: the model behaves as a natural
// coworker, sees the full workspace context (files + inspector + recent
// history), and only hands off to the deterministic planner/engine when
// it emits a validated plan directive. The deterministic planner is kept
// as an offline fallback when the AI gateway is unavailable.

type PlannerFile = { index: number; name: string; sheets: { name: string; headers: string[] }[] };

type FileForAiFull = PlannerFile & { inspector: unknown };

async function loadFilesForAi(
  supabase: import("@supabase/supabase-js").SupabaseClient<import("@/integrations/supabase/types").Database>,
  workspaceId: string,
): Promise<FileForAiFull[]> {
  const { data: files } = await supabase
    .from("workspace_files")
    .select("original_name,sheet_meta,inspector")
    .eq("workspace_id", workspaceId)
    .order("created_at", { ascending: true });
  return (files ?? []).map((f, i) => {
    const meta = f.sheet_meta as { sheets?: { name: string; headers: string[] }[] } | null;
    return {
      index: i,
      name: f.original_name,
      sheets: meta?.sheets ?? [],
      inspector: f.inspector,
    };
  });
}

async function loadRecentHistory(
  supabase: import("@supabase/supabase-js").SupabaseClient<import("@/integrations/supabase/types").Database>,
  workspaceId: string,
): Promise<{ role: "user" | "assistant"; content: string }[]> {
  const { data } = await supabase
    .from("workspace_messages")
    .select("role,content,created_at")
    .eq("workspace_id", workspaceId)
    .order("created_at", { ascending: false })
    .limit(20);
  const rows = (data ?? []).slice().reverse();
  return rows
    .filter((r) => r.role === "user" || r.role === "assistant")
    .map((r) => ({ role: r.role as "user" | "assistant", content: r.content ?? "" }));
}

export const sendMessage = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) =>
    z.object({ workspaceId: z.string().uuid(), text: z.string().min(1).max(4000) }).parse(input),
  )
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { data: ws } = await supabase
      .from("workspaces")
      .select("id,user_id")
      .eq("id", data.workspaceId)
      .single();
    if (!ws || ws.user_id !== userId) throw new Error("Workspace not found");

    // Persist the user message BEFORE calling the model so it appears in
    // history and the assistant can reference it naturally.
    await supabase.from("workspace_messages").insert({
      workspace_id: data.workspaceId,
      user_id: userId,
      role: "user",
      content: data.text,
    });

    const filesForAi = await loadFilesForAi(supabase, data.workspaceId);
    if (filesForAi.length === 0) {
      const reply =
        "Go ahead and drop one or more spreadsheets on the left — as soon as they're up, I'll take a look and suggest what to do.";
      await supabase.from("workspace_messages").insert({
        workspace_id: data.workspaceId,
        user_id: userId,
        role: "assistant",
        content: reply,
      });
      return { kind: "text" as const, text: reply };
    }

    const apiKey = process.env.LOVABLE_API_KEY;

    // ---- LLM-first path ---------------------------------------------------
    if (apiKey) {
      try {
        const { runConversation } = await import("./workspace/ai-conversation.server");
        const history = await loadRecentHistory(supabase, data.workspaceId);
        const ctx = filesForAi.map((f) => ({
          index: f.index,
          name: f.name,
          sheets: f.sheets,
          inspector: (f.inspector as never) ?? null,
        }));
        const result = await runConversation(apiKey, ctx, history);
        type Json = import("@/integrations/supabase/types").Json;

        if (result.plan) {
          await supabase.from("workspace_messages").insert({
            workspace_id: data.workspaceId,
            user_id: userId,
            role: "assistant",
            content: result.reply || "Here's what I'll do — hit **Run** when you're ready.",
            tool_data: { kind: "plan", plan: result.plan } as unknown as Json,
          });
          return { kind: "plan" as const, plan: result.plan };
        }

        // Pure conversational reply.
        await supabase.from("workspace_messages").insert({
          workspace_id: data.workspaceId,
          user_id: userId,
          role: "assistant",
          content: result.reply,
        });
        return { kind: "text" as const, text: result.reply };
      } catch (err) {
        console.error("LLM conversation failed, falling back to deterministic planner:", err);
        // fall through to deterministic path below
      }
    }

    // ---- Deterministic fallback (no API key or LLM error) -----------------
    const { tryDeterministicPlan } = await import("./excel/deterministic-plan");
    const { extractKeyOverride } = await import("./excel/intent");
    const preferredKey = extractKeyOverride(data.text);
    const det = tryDeterministicPlan(data.text, filesForAi, { preferredKey });
    type Json = import("@/integrations/supabase/types").Json;

    if (det?.kind === "capabilities" || !det) {
      const reply =
        "I can compare files, find missing or common records, merge on a shared key, build a master sheet, run bulk lookups, dedupe, summarize, and clean up blank rows. Tell me in your own words what you want.";
      await supabase.from("workspace_messages").insert({
        workspace_id: data.workspaceId,
        user_id: userId,
        role: "assistant",
        content: reply,
      });
      return { kind: "text" as const, text: reply };
    }

    if (det.kind === "plan") {
      const plan: Plan = { ...det.plan, warnings: [...det.plan.warnings, ...det.notes] };
      const reply = `Here's what I'll do:\n\n${describePlan(plan)}\n\nReview and hit **Run** to apply it.`;
      await supabase.from("workspace_messages").insert({
        workspace_id: data.workspaceId,
        user_id: userId,
        role: "assistant",
        content: reply,
        tool_data: { kind: "plan", plan } as unknown as Json,
      });
      return { kind: "plan" as const, plan };
    }

    // needs_clarification
    const reply = `${det.reason}${
      det.candidateKeys.length ? ` Candidates I can see: ${det.candidateKeys.join(", ")}.` : ""
    }`;
    await supabase.from("workspace_messages").insert({
      workspace_id: data.workspaceId,
      user_id: userId,
      role: "assistant",
      content: reply,
      tool_data: {
        kind: "clarify",
        candidates: det.candidateKeys,
        sharedColumns: det.sharedColumns,
        pendingIntent: det.pendingIntent,
        pendingSide: det.pendingSide ?? null,
      } as unknown as Json,
    });
    return { kind: "clarify" as const, candidates: det.candidateKeys };
  });



function describePlan(plan: Plan): string {
  const lines: string[] = [];
  for (const op of plan.ops) {
    if (op.op === "merge") lines.push(`• Merge all files on **${op.keyColumn}** (${op.strategy})`);
    else if (op.op === "dedupe")
      lines.push(`• Remove duplicates${op.keyColumn ? ` by **${op.keyColumn}**` : " (full row)"}`);
    else if (op.op === "diff") lines.push(`• Compare files on **${op.keyColumn}** and list differences`);
    else if (op.op === "intersection")
      lines.push(`• Find rows present in both files, matched on **${op.keyColumn}**`);
    else if (op.op === "summary") lines.push(`• Add a summary sheet`);
    else if (op.op === "highlight_column")
      lines.push(`• Highlight ${op.rule} values in **${op.column}**`);
    else if (op.op === "recalc") lines.push(`• Recalculate formulas`);
    else if (op.op === "master_merge")
      lines.push(
        `• Build a **master sheet** across all files on **${op.keyColumn}** (${op.joinType} join, keep ${op.dupeStrategy})`,
      );
    else if (op.op === "bulk_lookup")
      lines.push(`• Bulk lookup of **${op.queries.length}** value(s) in file #${op.fileIndex + 1}`);
  }
  return lines.join("\n");
}


// ---------- Run plan → new version ----------

export const runProposedPlan = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) =>
    z
      .object({
        workspaceId: z.string().uuid(),
        plan: PlanSchema,
        label: z.string().max(120).optional(),
      })
      .parse(input),
  )
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { runPlan } = await import("./excel/engine.server");
    const { summarizeRun, friendlyError } = await import("./workspace/friendly-errors");

    const { data: ws } = await supabase
      .from("workspaces")
      .select("id,user_id,name")
      .eq("id", data.workspaceId)
      .single();
    if (!ws || ws.user_id !== userId) throw new Error("Workspace not found");

    const { data: files } = await supabase
      .from("workspace_files")
      .select("storage_path,original_name")
      .eq("workspace_id", data.workspaceId)
      .order("created_at", { ascending: true });
    if (!files || files.length === 0) throw new Error("No input files.");

    try {
      const engineFiles = await Promise.all(
        files.map(async (f) => {
          const { data: blob, error } = await supabaseAdmin.storage
            .from("excel-uploads")
            .download(f.storage_path);
          if (error || !blob) throw new Error(`Download failed: ${f.original_name}`);
          return { name: f.original_name, buffer: await blob.arrayBuffer() };
        }),
      );

      const result = await runPlan(engineFiles, data.plan);

      // Determine next version number
      const { data: prior } = await supabase
        .from("workspace_versions")
        .select("version_number")
        .eq("workspace_id", data.workspaceId)
        .order("version_number", { ascending: false })
        .limit(1);
      const nextVersion = (prior?.[0]?.version_number ?? 0) + 1;
      const label = data.label ?? planLabel(data.plan);
      const outputName = `${ws.name.replace(/[^a-z0-9\-_ ]/gi, "").slice(0, 40) || "result"}-v${nextVersion}.xlsx`;
      const outputPath = `${userId}/workspaces/${data.workspaceId}/v${nextVersion}-${crypto.randomUUID()}.xlsx`;
      const { error: upErr } = await supabaseAdmin.storage
        .from("excel-outputs")
        .upload(outputPath, result.buffer, {
          contentType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          upsert: true,
        });
      if (upErr) throw new Error(`Upload failed: ${upErr.message}`);

      type Json = import("@/integrations/supabase/types").Json;
      const { data: version, error: vErr } = await supabase
        .from("workspace_versions")
        .insert({
          workspace_id: data.workspaceId,
          user_id: userId,
          version_number: nextVersion,
          label,
          plan: data.plan as unknown as Json,
          stats: result.stats as Json,
          warnings: result.warnings as Json,
          output_path: outputPath,
          output_name: outputName,
          size_bytes: result.buffer.length,
        })
        .select("id")
        .single();
      if (vErr) throw new Error(vErr.message);

      const fallbackSummary = summarizeRun(result.stats, result.warnings);
      const { summarizeExecution } = await import("./workspace/ai-conversation.server");
      const conversational = await summarizeExecution(
        process.env.LOVABLE_API_KEY,
        label,
        result.stats,
        result.warnings,
        fallbackSummary,
      );
      await supabase.from("workspace_messages").insert({
        workspace_id: data.workspaceId,
        user_id: userId,
        role: "assistant",
        content: `${conversational}\n\n_Saved as **v${nextVersion} · ${label}** — download below._`,

        tool_data: {
          kind: "result",
          versionId: version.id,
          versionNumber: nextVersion,
          stats: result.stats,
          warnings: result.warnings,
        } as unknown as Json,
      });
      await supabase
        .from("workspaces")
        .update({ updated_at: new Date().toISOString() })
        .eq("id", data.workspaceId);
      return { ok: true, versionId: version.id, versionNumber: nextVersion };
    } catch (e) {
      const nice = friendlyError(e);
      await supabase.from("workspace_messages").insert({
        workspace_id: data.workspaceId,
        user_id: userId,
        role: "assistant",
        content: `I couldn't complete that. ${nice}`,
      });
      throw new Error(nice);
    }
  });

function planLabel(plan: Plan): string {
  const kinds = plan.ops.map((o) => o.op);
  if (kinds.includes("master_merge")) {
    const mm = plan.ops.find((o) => o.op === "master_merge") as
      | Extract<Plan["ops"][number], { op: "master_merge" }>
      | undefined;
    return mm ? `Master sheet (key: ${mm.keyColumn})` : "Master sheet";
  }
  if (kinds.includes("bulk_lookup")) {
    const bl = plan.ops.find((o) => o.op === "bulk_lookup") as
      | Extract<Plan["ops"][number], { op: "bulk_lookup" }>
      | undefined;
    return bl ? `Bulk lookup (${bl.queries.length} queries)` : "Bulk lookup";
  }
  if (kinds.includes("merge")) return "Merge";
  if (kinds.includes("intersection")) return "Common rows";
  if (kinds.includes("diff")) return "Compare";
  if (kinds.includes("dedupe")) return "Dedupe";
  if (kinds.includes("summary")) return "Summary";
  return kinds.join(" + ") || "Run";
}


export const getVersionDownloadUrl = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((input: unknown) =>
    z.object({ workspaceId: z.string().uuid(), versionId: z.string().uuid() }).parse(input),
  )
  .handler(async ({ data, context }) => {
    const { supabase, userId } = context;
    const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
    const { data: v } = await supabase
      .from("workspace_versions")
      .select("output_path,output_name,user_id,workspace_id")
      .eq("id", data.versionId)
      .single();
    if (!v || v.user_id !== userId || v.workspace_id !== data.workspaceId || !v.output_path)
      throw new Error("Version not found");
    const { data: signed, error } = await supabaseAdmin.storage
      .from("excel-outputs")
      .createSignedUrl(v.output_path, 300, { download: v.output_name ?? "result.xlsx" });
    if (error || !signed) throw new Error(error?.message ?? "Signing failed");
    return { url: signed.signedUrl, name: v.output_name ?? "result.xlsx" };
  });
