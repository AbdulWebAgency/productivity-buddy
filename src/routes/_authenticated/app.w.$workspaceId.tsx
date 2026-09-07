import ReactMarkdown from "react-markdown";
import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import { useEffect, useMemo, useRef, useState } from "react";
import { supabase } from "@/integrations/supabase/client";
import {
  getWorkspace,
  registerWorkspaceFiles,
  removeWorkspaceFile,
  sendMessage,
  runProposedPlan,
  getVersionDownloadUrl,
} from "@/lib/workspace.functions";
import { suggestActions } from "@/lib/workspace/suggestions";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Badge } from "@/components/ui/badge";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { Upload, X, Loader2, FileSpreadsheet, Play, Download, ArrowLeft, Sparkles, History, Send } from "lucide-react";
import { toast } from "sonner";
import type { Plan } from "@/lib/excel/types";

export const Route = createFileRoute("/_authenticated/app/w/$workspaceId")({
  component: WorkspacePage,
});

const MAX_FILE_MB = 25;

type WorkspaceData = Awaited<ReturnType<typeof getWorkspace>>;
type Message = WorkspaceData["messages"][number];
type FileRow = WorkspaceData["files"][number];
type VersionRow = WorkspaceData["versions"][number];

function WorkspacePage() {
  const { workspaceId } = Route.useParams();
  const qc = useQueryClient();
  const fetchWs = useServerFn(getWorkspace);
  const registerFn = useServerFn(registerWorkspaceFiles);
  const removeFileFn = useServerFn(removeWorkspaceFile);
  const sendFn = useServerFn(sendMessage);
  const runFn = useServerFn(runProposedPlan);
  const downloadFn = useServerFn(getVersionDownloadUrl);

  const key = ["workspace", workspaceId];
  const { data, isPending } = useQuery({
    queryKey: key,
    queryFn: () => fetchWs({ data: { workspaceId } }),
  });

  const [text, setText] = useState("");
  const [uploading, setUploading] = useState(false);
  const chatBottomRef = useRef<HTMLDivElement>(null);
  const composerRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    chatBottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [data?.messages.length]);

  useEffect(() => {
    composerRef.current?.focus();
  }, [workspaceId]);

  const send = useMutation({
    mutationFn: (t: string) => sendFn({ data: { workspaceId, text: t } }),
    onSuccess: () => qc.invalidateQueries({ queryKey: key }),
    onError: (e) => toast.error(e instanceof Error ? e.message : "Failed"),
  });

  const runPlan = useMutation({
    mutationFn: (plan: Plan) => runFn({ data: { workspaceId, plan } }),
    onSuccess: () => {
      toast.success("New version created");
      qc.invalidateQueries({ queryKey: key });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Run failed"),
  });

  async function handleUpload(files: FileList | null) {
    if (!files || files.length === 0) return;
    const chosen = Array.from(files);
    const invalid = chosen.find((f) => f.size > MAX_FILE_MB * 1024 * 1024);
    if (invalid) {
      toast.error(`${invalid.name} exceeds ${MAX_FILE_MB} MB`);
      return;
    }
    setUploading(true);
    try {
      const { data: userData } = await supabase.auth.getUser();
      if (!userData.user) throw new Error("Not signed in");
      const uploaded: { storagePath: string; originalName: string; sizeBytes: number }[] = [];
      for (const [i, file] of chosen.entries()) {
        const safe = file.name.replace(/[^a-z0-9.\-_ ]/gi, "");
        const path = `${userData.user.id}/workspaces/${workspaceId}/${Date.now()}-${i}-${safe}`;
        const { error } = await supabase.storage.from("excel-uploads").upload(path, file, {
          upsert: true,
          contentType: file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        if (error) throw new Error(error.message);
        uploaded.push({ storagePath: path, originalName: file.name, sizeBytes: file.size });
      }
      await registerFn({ data: { workspaceId, files: uploaded } });
      toast.success(`Uploaded ${uploaded.length} file(s)`);
      qc.invalidateQueries({ queryKey: key });
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Upload failed");
    } finally {
      setUploading(false);
    }
  }

  const removeFile = useMutation({
    mutationFn: (fileId: string) => removeFileFn({ data: { workspaceId, fileId } }),
    onSuccess: () => qc.invalidateQueries({ queryKey: key }),
  });

  const suggestions = useMemo(() => {
    if (!data) return [];
    return suggestActions({
      files: data.files.map((f) => ({
        name: f.original_name,
        inspector: (f.inspector as never) ?? null,
      })),
    });
  }, [data]);

  async function doDownload(versionId: string) {
    try {
      const { url, name } = await downloadFn({ data: { workspaceId, versionId } });
      const a = document.createElement("a");
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : "Download failed");
    }
  }

  function submit() {
    const t = text.trim();
    if (!t) return;
    setText("");
    send.mutate(t);
  }

  if (isPending || !data) {
    return (
      <div className="flex items-center gap-2 text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> Loading workspace…
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="flex items-center gap-3">
        <Button asChild variant="ghost" size="sm">
          <Link to="/app/workspaces">
            <ArrowLeft className="mr-2 h-4 w-4" /> Workspaces
          </Link>
        </Button>
        <h1 className="text-2xl">{data.workspace.name}</h1>
      </div>

      <div className="grid gap-4 lg:grid-cols-[280px_1fr_280px]">
        {/* LEFT: files + inspector */}
        <div className="space-y-3">
          <FilesPanel
            files={data.files}
            uploading={uploading}
            onUpload={handleUpload}
            onRemove={(id) => removeFile.mutate(id)}
          />
        </div>

        {/* CENTER: chat */}
        <div className="flex min-h-[70vh] flex-col rounded-lg border border-border bg-card">
          <div className="flex-1 space-y-4 overflow-y-auto p-4">
            {data.messages.map((m) => (
              <MessageBubble
                key={m.id}
                message={m}
                onRun={(plan) => runPlan.mutate(plan)}
                onDownload={doDownload}
                running={runPlan.isPending}
                versions={data.versions}
              />
            ))}
            {send.isPending && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <Loader2 className="h-3.5 w-3.5 animate-spin" /> Thinking…
              </div>
            )}
            <div ref={chatBottomRef} />
          </div>

          {suggestions.length > 0 && data.files.length > 0 && (
            <div className="flex flex-wrap gap-2 border-t border-border px-4 py-2">
              {suggestions.map((s) => (
                <button
                  key={s.label}
                  onClick={() => setText(s.prompt)}
                  className="rounded-full border border-border bg-secondary/50 px-3 py-1 text-xs hover:bg-secondary"
                >
                  <Sparkles className="mr-1 inline h-3 w-3" />
                  {s.label}
                </button>
              ))}
            </div>
          )}

          <div className="flex items-end gap-2 border-t border-border p-3">
            <Textarea
              ref={composerRef}
              rows={2}
              value={text}
              onChange={(e) => setText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  submit();
                }
              }}
              placeholder={
                data.files.length === 0
                  ? "Upload files first, then chat here…"
                  : "Ask me to compare, merge, dedupe, or explain…"
              }
              className="min-h-[52px] resize-none"
              disabled={send.isPending}
            />
            <Button size="icon" onClick={submit} disabled={send.isPending || !text.trim()} title="Send">
              <Send className="h-4 w-4" />
            </Button>
          </div>
        </div>

        {/* RIGHT: versions */}
        <div className="space-y-3">
          <VersionsPanel versions={data.versions} onDownload={doDownload} />
        </div>
      </div>
    </div>
  );
}

// ---------- Files panel ----------

function FilesPanel({
  files,
  uploading,
  onUpload,
  onRemove,
}: {
  files: FileRow[];
  uploading: boolean;
  onUpload: (files: FileList | null) => void;
  onRemove: (id: string) => void;
}) {
  return (
    <Card>
      <CardHeader className="pb-2">
        <CardTitle className="flex items-center gap-2 text-sm">
          <FileSpreadsheet className="h-4 w-4" /> Files
        </CardTitle>
      </CardHeader>
      <CardContent className="space-y-3">
        <label
          htmlFor="ws-file-input"
          className="flex cursor-pointer flex-col items-center justify-center gap-1 rounded-md border-2 border-dashed border-border bg-secondary/40 p-4 text-center text-xs transition-colors hover:bg-secondary"
        >
          {uploading ? (
            <Loader2 className="h-5 w-5 animate-spin text-primary" />
          ) : (
            <Upload className="h-5 w-5 text-primary" />
          )}
          <div className="font-medium">Upload .xlsx</div>
          <div className="text-muted-foreground">Up to {MAX_FILE_MB} MB each</div>
          <Input
            id="ws-file-input"
            type="file"
            accept=".xlsx"
            multiple
            className="hidden"
            onChange={(e) => {
              onUpload(e.target.files);
              e.target.value = "";
            }}
          />
        </label>

        {files.length === 0 ? (
          <div className="rounded-md bg-muted/40 p-3 text-xs text-muted-foreground">No files yet. Upload to start.</div>
        ) : (
          <Accordion type="multiple" className="w-full">
            {files.map((f) => (
              <AccordionItem key={f.id} value={f.id} className="border-b-0">
                <div className="flex items-center gap-1">
                  <AccordionTrigger className="flex-1 py-2 text-left text-xs hover:no-underline">
                    <div className="min-w-0 flex-1">
                      <div className="truncate font-medium">{f.original_name}</div>
                      <div className="text-[10px] text-muted-foreground">{(f.size_bytes / 1024).toFixed(0)} KB</div>
                    </div>
                  </AccordionTrigger>
                  <button
                    onClick={() => onRemove(f.id)}
                    className="mr-1 rounded p-1 text-muted-foreground hover:bg-muted hover:text-destructive"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                </div>
                <AccordionContent>
                  <InspectorSummary inspector={f.inspector} />
                </AccordionContent>
              </AccordionItem>
            ))}
          </Accordion>
        )}
      </CardContent>
    </Card>
  );
}

function InspectorSummary({ inspector }: { inspector: unknown }) {
  const insp = inspector as {
    sheets?: {
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
    }[];
    warnings?: string[];
    primaryDataSheet?: string | null;
  } | null;
  if (!insp?.sheets?.length) {
    return <div className="text-[11px] text-muted-foreground">No inspection data.</div>;
  }
  return (
    <div className="space-y-2 text-[11px]">
      {insp.sheets.map((s) => (
        <div key={s.name} className="rounded bg-muted/40 p-2">
          <div className="font-medium">{s.name}</div>
          <div className="text-muted-foreground">
            {s.rows} rows · {s.columns} cols
            {s.blankRows > 0 && ` · ${s.blankRows} blank`}
            {s.formulaCells > 0 && ` · ${s.formulaCells} formulas`}
          </div>
          {s.likelyKeys.length > 0 && (
            <div className="mt-1">
              <span className="text-muted-foreground">Likely key: </span>
              <span className="font-medium">{s.likelyKeys[0]}</span>
            </div>
          )}
          {s.duplicateKeyValues > 0 && (
            <div className="text-amber-600">{s.duplicateKeyValues} duplicate key values</div>
          )}
        </div>
      ))}
    </div>
  );
}

// ---------- Message bubble ----------

function MessageBubble({
  message,
  onRun,
  onDownload,
  running,
  versions,
}: {
  message: Message;
  onRun: (plan: Plan) => void;
  onDownload: (versionId: string) => void;
  running: boolean;
  versions: VersionRow[];
}) {
  const isUser = message.role === "user";
  const tool = message.tool_data as
    | { kind: "plan"; plan: Plan }
    | { kind: "clarify"; candidates: string[]; sharedColumns: string[] }
    | { kind: "result"; versionId: string; versionNumber: number; stats: Record<string, unknown>; warnings: string[] }
    | null;

  return (
    <div className={isUser ? "flex justify-end" : "flex justify-start"}>
      <div
        className={
          isUser
            ? "max-w-[80%] rounded-2xl rounded-br-sm bg-primary px-4 py-2 text-primary-foreground"
            : "max-w-[85%] space-y-2 text-sm text-foreground"
        }
      >
        <div className="text-sm">
          <ReactMarkdown
            components={{
              p: ({ children }) => <p className="whitespace-pre-wrap mb-2 last:mb-0">{children}</p>,
              ul: ({ children }) => <ul className="list-disc pl-5 mb-2">{children}</ul>,
              ol: ({ children }) => <ol className="list-decimal pl-5 mb-2">{children}</ol>,
              li: ({ children }) => <li className="mb-1">{children}</li>,
              strong: ({ children }) => <strong className="font-semibold">{children}</strong>,
              code: ({ children }) => <code className="rounded bg-muted px-1 py-0.5 text-xs">{children}</code>,
            }}
          >
            {message.content}
          </ReactMarkdown>
        </div>

        {tool?.kind === "plan" && <PlanCard plan={tool.plan} onRun={() => onRun(tool.plan)} running={running} />}

        {tool?.kind === "result" && (
          <ResultCard
            versionId={tool.versionId}
            versionNumber={tool.versionNumber}
            versions={versions}
            onDownload={onDownload}
            stats={tool.stats}
          />
        )}
      </div>
    </div>
  );
}

function PlanCard({ plan, onRun, running }: { plan: Plan; onRun: () => void; running: boolean }) {
  return (
    <Card className="border-primary/30 bg-secondary/40">
      <CardContent className="space-y-3 p-3">
        <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
          <Sparkles className="h-3.5 w-3.5" /> Execution plan
        </div>
        <ul className="space-y-1 text-sm">
          {plan.ops.map((op, i) => (
            <li key={i}>
              {op.op === "merge" && (
                <>
                  Merge all files on <b>{op.keyColumn}</b> ({op.strategy})
                </>
              )}
              {op.op === "dedupe" && (
                <>
                  Remove duplicates{" "}
                  {op.keyColumn ? (
                    <>
                      by <b>{op.keyColumn}</b>
                    </>
                  ) : (
                    "(full row)"
                  )}
                </>
              )}
              {op.op === "diff" && (
                <>
                  Compare files on <b>{op.keyColumn}</b> and list differences
                </>
              )}
              {op.op === "summary" && <>Add a summary sheet</>}
              {op.op === "highlight_column" && (
                <>
                  Highlight {op.rule} values in <b>{op.column}</b>
                </>
              )}
              {op.op === "recalc" && <>Recalculate formulas</>}
              {op.op === "master_merge" && (
                <>
                  Create master sheet on <b>{op.keyColumn}</b> ({op.joinType} join, keep {op.dupeStrategy})
                </>
              )}
              {op.op === "bulk_lookup" && (
                <>
                  Bulk lookup of <b>{op.queries.length}</b> value(s) in file #{op.fileIndex + 1}
                </>
              )}
            </li>
          ))}
        </ul>
        {plan.projection?.columns && plan.projection.columns.length > 0 && (
          <div className="text-xs text-muted-foreground">
            Columns: <span className="font-medium text-foreground">{plan.projection.columns.join(", ")}</span>
          </div>
        )}
        {plan.warnings.length > 0 && (
          <div className="text-xs text-muted-foreground">Notes: {plan.warnings.join("; ")}</div>
        )}
        <Button size="sm" onClick={onRun} disabled={running}>
          {running ? <Loader2 className="mr-2 h-3.5 w-3.5 animate-spin" /> : <Play className="mr-2 h-3.5 w-3.5" />}
          Run
        </Button>
      </CardContent>
    </Card>
  );
}

function ResultCard({
  versionId,
  versionNumber,
  versions,
  onDownload,
  stats,
}: {
  versionId: string;
  versionNumber: number;
  versions: VersionRow[];
  onDownload: (id: string) => void;
  stats: Record<string, unknown>;
}) {
  const v = versions.find((x) => x.id === versionId);
  const ops = Array.isArray(stats.ops) ? (stats.ops as Record<string, unknown>[]) : [];
  return (
    <Card className="border-emerald-500/30 bg-emerald-50/40 dark:bg-emerald-950/20">
      <CardContent className="space-y-2 p-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Badge variant="secondary">v{versionNumber}</Badge>
            <span className="text-sm font-medium">{v?.label ?? "Result"}</span>
          </div>
          <Button size="sm" variant="outline" onClick={() => onDownload(versionId)}>
            <Download className="mr-2 h-3.5 w-3.5" /> Download
          </Button>
        </div>
        {ops.length > 0 && (
          <details className="text-xs text-muted-foreground">
            <summary className="cursor-pointer">Execution details</summary>
            <pre className="mt-2 overflow-x-auto rounded bg-muted/40 p-2 text-[10px]">
              {JSON.stringify(ops, null, 2)}
            </pre>
          </details>
        )}
      </CardContent>
    </Card>
  );
}

// ---------- Versions panel ----------

function VersionsPanel({ versions, onDownload }: { versions: VersionRow[]; onDownload: (id: string) => void }) {
  return (
    <Card>
      <CardHeader className="pb-2">
        <CardTitle className="flex items-center gap-2 text-sm">
          <History className="h-4 w-4" /> Versions
        </CardTitle>
      </CardHeader>
      <CardContent>
        {versions.length === 0 ? (
          <div className="rounded-md bg-muted/40 p-3 text-xs text-muted-foreground">
            No versions yet. Run a plan to create one.
          </div>
        ) : (
          <ol className="space-y-2">
            {versions.map((v) => (
              <li key={v.id} className="rounded-md border border-border bg-card p-2">
                <div className="flex items-center justify-between">
                  <div className="text-xs">
                    <Badge variant="secondary" className="mr-1">
                      v{v.version_number}
                    </Badge>
                    <span className="font-medium">{v.label}</span>
                  </div>
                  <button
                    onClick={() => onDownload(v.id)}
                    className="rounded p-1 text-muted-foreground hover:bg-muted hover:text-foreground"
                    title="Download"
                  >
                    <Download className="h-3.5 w-3.5" />
                  </button>
                </div>
                <div className="mt-1 text-[10px] text-muted-foreground">{new Date(v.created_at).toLocaleString()}</div>
              </li>
            ))}
          </ol>
        )}
      </CardContent>
    </Card>
  );
}
