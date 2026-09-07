// Conversational AI layer for Productivity Buddy.
// The deterministic planner and engine are unchanged — this file makes the
// assistant feel like a natural coworker: it inspects files, greets the user,
// carries context turn-to-turn, and decides when to hand off to the planner.

import type { Plan } from "../excel/types";
import { PlanSchema } from "../excel/types";
import { repairAndParsePlan } from "../excel/plan-repair";

type SheetMeta = { name: string; headers: string[] };
type InspectorSheet = {
  name: string;
  rows: number;
  columns: number;
  headers: string[];
  likelyKeys: string[];
  blankRows: number;
  formulaCells: number;
  duplicateKeyValues: number;
  worksheetType?: string;
  confidence?: number;
  likelyPrimaryTable?: boolean;
  rank?: number;
  hidden?: boolean;
};
type InspectorReport = { sheets: InspectorSheet[]; warnings: string[]; primaryDataSheet?: string | null } | null;

export type WorkspaceFileCtx = {
  index: number;
  name: string;
  sheets: SheetMeta[];
  inspector: InspectorReport;
};

export type ChatTurn = { role: "user" | "assistant"; content: string };

// ---------- Deterministic greeting after upload ----------

import { normalizeHeader as norm, scoreKey } from "@/lib/excel/engine/shared/headers";
import { primarySheetOf } from "./primary-sheet";

// Always reason over the recommended data sheet, never a pivot/summary sheet.
function mainSheet(f: WorkspaceFileCtx): InspectorSheet | undefined {
  return primarySheetOf(f.inspector);
}

function keysSharedAcross(files: WorkspaceFileCtx[]): string[] {
  if (files.length === 0) return [];
  const perFile = files.map((f) => {
    const sheet = mainSheet(f);
    const headers = sheet?.headers ?? f.sheets[0]?.headers ?? [];
    return new Set(headers.map(norm));
  });
  const [first, ...rest] = perFile;
  if (!first) return [];
  const displayFor = new Map<string, string>();
  const s0 = mainSheet(files[0])?.headers ?? files[0].sheets[0]?.headers ?? [];
  s0.forEach((h) => displayFor.set(norm(h), h));
  const shared: string[] = [];
  for (const k of first) {
    if (rest.every((s) => s.has(k))) shared.push(displayFor.get(k) ?? k);
  }
  return shared;
}

export function pickBestSharedKey(files: WorkspaceFileCtx[]): string | null {
  const shared = keysSharedAcross(files);
  if (shared.length === 0) return null;
  const sorted = [...shared].sort((a, b) => scoreKey(b) - scoreKey(a));
  return sorted[0] ?? null;
}

export function buildInspectionGreeting(files: WorkspaceFileCtx[]): string {
  if (files.length === 0) {
    return "Upload one or more spreadsheets and I'll take a look right away.";
  }
  const lines: string[] = [
    files.length === 1
      ? "I've finished inspecting your file."
      : `I've finished inspecting your ${files.length} uploaded files.`,
    "",
  ];
  for (const f of files) {
    const sheet = mainSheet(f);
    const rows = sheet?.rows ?? 0;
    const cols = sheet?.headers?.length ?? f.sheets[0]?.headers?.length ?? 0;
    const suffix =
      rows > 0
        ? ` — ${rows.toLocaleString()} row${rows === 1 ? "" : "s"}, ${cols} column${cols === 1 ? "" : "s"}`
        : cols > 0
          ? ` — ${cols} column${cols === 1 ? "" : "s"}`
          : "";
    const sheetNote = sheet && (f.inspector?.sheets?.length ?? 0) > 1 ? ` (using sheet "${sheet.name}")` : "";
    lines.push(`• **${f.name}**${suffix}${sheetNote}`);
    const skipped = (f.inspector?.sheets ?? []).filter(
      (s) => s.name !== sheet?.name && s.worksheetType && s.worksheetType !== "DATA" && s.worksheetType !== "UNKNOWN",
    );
    if (skipped.length) {
      lines.push(
        `  ↳ ignoring ${skipped.map((s) => `_${s.name}_ (${(s.worksheetType ?? "other").toLowerCase()})`).join(", ")}`,
      );
    }
  }
  const shared = keysSharedAcross(files);
  const best = pickBestSharedKey(files);
  if (files.length >= 2 && best) {
    lines.push("");
    lines.push(`✅ All ${files.length} files share **${best}**, so they're ready to be matched using that key.`);
    const others = shared.filter((s) => s !== best).slice(0, 3);
    if (others.length) {
      lines.push(`I also found these shared columns: ${others.map((o) => `_${o}_`).join(", ")}.`);
    }
  } else if (files.length >= 2 && shared.length === 0) {
    lines.push("");
    lines.push(
      "The files don't share an obvious column name. Tell me which columns to match on and I'll handle the rest.",
    );
  }
  // Warnings worth surfacing
  const warns = files.flatMap((f) => f.inspector?.warnings ?? []).slice(0, 3);
  if (warns.length) {
    lines.push("");
    lines.push(`Heads up: ${warns.join("; ")}.`);
  }
  lines.push("");
  lines.push("What would you like me to do?");
  lines.push("");
  lines.push(
    "You can now ask me to create a master sheet, compare files, find missing records, remove duplicates, or perform a bulk lookup.",
  );
  return lines.join("\n");
}

// ---------- LLM conversation ----------

function fileCatalog(files: WorkspaceFileCtx[]): string {
  return files
    .map((f) => {
      const insp = mainSheet(f);
      const rows = insp?.rows ?? 0;
      const headers = (insp?.headers ?? f.sheets[0]?.headers ?? []).slice(0, 40);
      const keys = insp?.likelyKeys?.slice(0, 3) ?? [];
      const dupes = insp?.duplicateKeyValues ?? 0;
      const all = f.inspector?.sheets ?? [];
      const others = all
        .filter((s) => s.name !== insp?.name)
        .map((s) => `${s.name} [${s.worksheetType ?? "UNKNOWN"}${s.confidence != null ? ` ${s.confidence}%` : ""}]`);
      return [
        `#${f.index} "${f.name}" — ${rows} rows`,
        insp
          ? `  primary data sheet: "${insp.name}" [${insp.worksheetType ?? "DATA"}${insp.confidence != null ? ` ${insp.confidence}%` : ""}]`
          : null,
        `  columns: ${headers.join(" | ")}`,
        keys.length ? `  likely keys: ${keys.join(", ")}` : null,
        dupes > 0 ? `  ⚠ ${dupes} duplicate values in ${keys[0] ?? "key"}` : null,
        others.length ? `  other worksheets (do NOT use for merge/compare): ${others.join(", ")}` : null,
      ]
        .filter(Boolean)
        .join("\n");
    })
    .join("\n\n");
}

const SYSTEM_PROMPT = `You are Productivity Buddy, an intelligent AI coworker who helps people work with Excel workbooks. You are NOT a wizard or command parser — you talk like a warm, capable colleague who happens to have a powerful spreadsheet engine at your disposal.

## How you behave
- Speak naturally, in the first person. Keep replies short (2–5 sentences) unless the user asks for detail.
- USE THE CONVERSATION HISTORY. If a previous turn established the task, the file, or the key column, don't re-ask.
- Be proactive: recommend the best matching key when files clearly share one; explain trade-offs briefly ("Name has duplicates, Registration No is unique").
- Never ask "which operation?" if the task is obvious from context. Never ask the same question twice.
- Refer to files by their real names, not "file #1". You always know what "them", "these", "the two files" refer to based on the file catalog and history.
- When the user just wants to chat, explain, or asks what you can do — reply conversationally without proposing a plan.

## When to propose a plan
Only when the user's intent is clear enough to execute. Confirm what you're about to do in one sentence, then attach a machine-readable plan block. The user will see a "Run" button.

To propose a plan, END your reply with a fenced block using the language tag \`plan\`:

\`\`\`plan
{
  "summary": "Short human summary",
  "ops": [ { "op": "master_merge", "keyColumn": "Registration Number", "joinType": "outer", "dupeStrategy": "first" } ],
  "warnings": []
}
\`\`\`

Supported ops (use exact op names and casing):
- { "op": "merge", "keyColumn": "...", "strategy": "union"|"intersection", "highlightUnmatched": true }
- { "op": "master_merge", "keyColumn": "...", "joinType": "outer"|"inner"|"left"|"append", "dupeStrategy": "first"|"latest"|"merge" }
- { "op": "intersection", "keyColumn": "...", "fileAIndex": 0, "fileBIndex": 1 }
- { "op": "diff", "keyColumn": "...", "fileAIndex": 0, "fileBIndex": 1 }
- { "op": "dedupe", "strategy": "key"|"full_row", "keyColumn": "..." }
- { "op": "summary", "includeCharts": false }
- { "op": "highlight_column", "column": "...", "rule": "missing"|"duplicate"|"outlier" }
- { "op": "bulk_lookup", "fileIndex": 0, "queries": ["...", "..."] }
- { "op": "recalc" }

Rules:
- Use column names EXACTLY as they appear in the file catalog. Prefer a column marked "likely key".
- Never invent columns. If the requested column doesn't exist in any file, ask which one to use.
- One plan block per reply, at the very end. Nothing after the closing fence.
- If a key is ambiguous (multiple equally-good columns), don't emit a plan — ask a single, specific question naming the candidates.
- If the user is just chatting / asking a question / not requesting an action, don't emit a plan.

## Choosing the right op for "missing / present / compare / not in" requests
When the user asks anything like "who's in X but missing from Y", "list students present in allocation but not in main", "find the difference", "what's missing", "compare these two" — ALWAYS use \`diff\` (not merge, not intersection). The diff op ALWAYS produces two sheets so both directions are visible: "Only in <fileA>" and "Only in <fileB>". Pick fileAIndex/fileBIndex from the catalog #numbers; don't stress about the order — both sides are reported. In your reply, name the sheet the user is looking for by its real name ("Only in <filename>"). If the user says "add missing rows into X", follow the diff with a \`master_merge\` using joinType "outer" (or "left" if they only want to enrich X).

## Worksheets
Each file's catalog entry names its **primary data sheet** plus any other worksheets with a classification (DATA, PIVOT, SUMMARY, DOCUMENTATION, EMPTY, UNKNOWN) and a confidence score. Reason only over DATA worksheets — the columns listed in the catalog come from the primary data sheet. Never use a PIVOT, SUMMARY or DOCUMENTATION worksheet for merge, compare, intersection or missing-record work. You may mention those sheets when relevant (e.g. explaining that you skipped them), and if the user explicitly names a different worksheet, say which one you'll use.

## Choosing key columns
- If a column is flagged as "likely key" in the catalog for BOTH files, use it.
- Prefer registration numbers / IDs / emails over names — names collide.
- Only ask the user to pick a key if there is genuine ambiguity. If they've already answered once in history, use that answer and don't re-ask.

## Requested output columns
If — and only if — the user explicitly asks for specific output columns (e.g. "give me only Name, Reg No and Email", "just show emails", "with columns X, Y"), add a top-level "projection" object to the plan:

  "projection": { "columns": ["Name", "Reg No", "Email"] }

Use the column names as the user wrote them; the engine fuzzy-matches per file. NEVER add projection when the user hasn't asked for specific columns — omitting it exports every column, which is the default.`;

export type ConversationResult = {
  reply: string; // user-visible text (plan block stripped)
  plan?: Plan; // present when the model proposed a valid plan
  planRepairs?: string[];
  planError?: string; // present when a plan block was found but couldn't be repaired
};

// Extract ```plan ... ``` block; return { visibleText, jsonRaw }.
function extractPlanBlock(text: string): { visible: string; json: string | null } {
  const re = /```plan\s*([\s\S]*?)```/i;
  const m = text.match(re);
  if (!m) return { visible: text.trim(), json: null };
  const visible = (text.slice(0, m.index) + text.slice((m.index ?? 0) + m[0].length)).replace(/\n{3,}/g, "\n\n").trim();
  return { visible, json: m[1].trim() };
}

export async function runConversation(
  apiKey: string,
  files: WorkspaceFileCtx[],
  history: ChatTurn[],
): Promise<ConversationResult> {
  const { createLovableAiGatewayProvider } = await import("../ai-gateway.server");
  const { generateText } = await import("ai");
  const gateway = createLovableAiGatewayProvider(apiKey);

  const catalog = fileCatalog(files);
  const contextBlock = files.length
    ? `## Uploaded files (authoritative)\n\n${catalog}`
    : "## Uploaded files\n\n(none yet — the user hasn't uploaded any spreadsheets)";

  // Cap history to keep prompt bounded.
  const recent = history.slice(-16);
  const chatMessages = recent.map((t) => ({ role: t.role, content: t.content }));

  const { text } = await generateText({
    model: gateway("google/gemini-2.5-flash"),
    system: `${SYSTEM_PROMPT}\n\n${contextBlock}`,
    messages: chatMessages,
  });

  const raw = (text ?? "").trim();
  const { visible, json } = extractPlanBlock(raw);

  if (!json) {
    return { reply: visible || "Let me know what you'd like to do next." };
  }

  // Try strict parse first, then plan-repair as a fallback.
  try {
    const parsed = JSON.parse(json);
    const result = PlanSchema.safeParse(parsed);
    if (result.success) {
      return { reply: visible, plan: result.data };
    }
  } catch {
    // fall through to repair
  }
  try {
    const { plan, repairs } = repairAndParsePlan(json, files);
    return { reply: visible, plan, planRepairs: repairs };
  } catch (e) {
    return {
      reply:
        visible ||
        "I had trouble putting together a valid plan for that. Could you rephrase or tell me which column to match on?",
      planError: e instanceof Error ? e.message : String(e),
    };
  }
}

// ---------- Post-run natural summary ----------

export async function summarizeExecution(
  apiKey: string | undefined,
  label: string,
  stats: Record<string, unknown>,
  warnings: string[],
  fallback: string,
): Promise<string> {
  if (!apiKey) return fallback;
  try {
    const { createLovableAiGatewayProvider } = await import("../ai-gateway.server");
    const { generateText } = await import("ai");
    const gateway = createLovableAiGatewayProvider(apiKey);
    const { text } = await generateText({
      model: gateway("google/gemini-2.5-flash"),
      system:
        "You are Productivity Buddy summarizing the outcome of an Excel operation to the user. Write 2–5 short lines. Warm, natural, first person. Mention concrete numbers (matched, missing, duplicates, rows, files). When the op was a diff, ALWAYS name each 'Only in <file>' sheet and its row count explicitly (e.g. 'The \"Only in allocation.xlsx\" sheet has 12 rows — those are the students present in allocation but not in main'). If a sheet has 0 rows, say so and explain what that means. Include any warnings clearly. Do NOT dump raw JSON. Do NOT mention 'the engine' or internal ops. End with a brief note that they can download the new version below.",
      prompt: `Operation: ${label}\n\nStats (JSON): ${JSON.stringify(stats).slice(0, 4000)}\n\nWarnings: ${warnings.slice(0, 8).join("; ") || "none"}`,
    });
    const t = text.trim();
    return t || fallback;
  } catch {
    return fallback;
  }
}
