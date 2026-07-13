import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import { getJob, planJob, runJob, getDownloadUrl, deleteJob } from "@/lib/excel.functions";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { toast } from "sonner";
import {
  ArrowLeft,
  Sparkles,
  Play,
  Download,
  Loader2,
  FileSpreadsheet,
  AlertTriangle,
  Trash2,
  CheckCircle2,
} from "lucide-react";

export const Route = createFileRoute("/_authenticated/app/$jobId")({
  component: JobDetail,
});

function JobDetail() {
  const { jobId } = Route.useParams();
  const navigate = useNavigate();
  const qc = useQueryClient();

  const fetchJob = useServerFn(getJob);
  const plan = useServerFn(planJob);
  const run = useServerFn(runJob);
  const download = useServerFn(getDownloadUrl);
  const del = useServerFn(deleteJob);

  const { data, isPending, error } = useQuery({
    queryKey: ["job", jobId],
    queryFn: () => fetchJob({ data: { jobId } }),
    refetchInterval: (q) => {
      const s = q.state.data?.job.status;
      return s && ["planning", "running"].includes(s) ? 2000 : false;
    },
  });

  const planMut = useMutation({
    mutationFn: () => plan({ data: { jobId } }),
    onSuccess: () => {
      toast.success("Plan ready — review and run.");
      qc.invalidateQueries({ queryKey: ["job", jobId] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Planning failed"),
  });

  const runMut = useMutation({
    mutationFn: () => run({ data: { jobId } }),
    onSuccess: () => {
      toast.success("Job started");
      qc.invalidateQueries({ queryKey: ["job", jobId] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Run failed"),
  });

  const dl = useMutation({
    mutationFn: () => download({ data: { jobId } }),
    onSuccess: ({ url }) => {
      window.open(url, "_blank");
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Download failed"),
  });

  const delMut = useMutation({
    mutationFn: () => del({ data: { jobId } }),
    onSuccess: () => {
      toast.success("Job deleted");
      qc.invalidateQueries({ queryKey: ["jobs"] });
      navigate({ to: "/app" });
    },
  });

  if (isPending) {
    return (
      <div className="flex items-center gap-2 text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> Loading job…
      </div>
    );
  }
  if (error || !data) {
    return (
      <Alert variant="destructive">
        <AlertDescription>{error instanceof Error ? error.message : "Job not found"}</AlertDescription>
      </Alert>
    );
  }

  const { job, files } = data;
  const inputFiles = files.filter((f) => f.role === "input");
  const canPlan = inputFiles.length > 0 && ["draft", "failed"].includes(job.status);
  const canRun = job.ai_plan && !["running", "planning"].includes(job.status);
  const isBusy = ["planning", "running"].includes(job.status);
  const done = job.status === "succeeded";

  return (
    <div className="mx-auto max-w-3xl">
      <Button asChild variant="ghost" size="sm" className="mb-4 -ml-2">
        <Link to="/app">
          <ArrowLeft className="mr-1 h-4 w-4" /> All jobs
        </Link>
      </Button>

      <div className="flex items-start justify-between gap-4">
        <div>
          <h1 className="text-3xl">{job.name}</h1>
          <p className="mt-1 text-sm text-muted-foreground capitalize">
            {job.kind} · {job.status}
          </p>
        </div>
        <Button variant="ghost" size="icon" onClick={() => { if (confirm("Delete this job?")) delMut.mutate(); }}>
          <Trash2 className="h-4 w-4 text-muted-foreground" />
        </Button>
      </div>

      {job.error ? (
        <Alert variant="destructive" className="mt-6">
          <AlertTriangle className="h-4 w-4" />
          <AlertDescription>{job.error}</AlertDescription>
        </Alert>
      ) : null}

      {job.intent ? (
        <Card className="mt-6">
          <CardHeader>
            <CardTitle className="text-base">Intent</CardTitle>
          </CardHeader>
          <CardContent className="whitespace-pre-wrap text-sm">{job.intent}</CardContent>
        </Card>
      ) : null}

      <Card className="mt-6">
        <CardHeader>
          <CardTitle className="text-base">Input files ({inputFiles.length})</CardTitle>
        </CardHeader>
        <CardContent>
          <ul className="divide-y">
            {inputFiles.map((f) => {
              const meta = f.sheet_meta as { sheets?: { name: string; rowCount: number; headers: string[] }[] } | null;
              return (
                <li key={f.id} className="flex items-start gap-3 py-3 first:pt-0 last:pb-0">
                  <FileSpreadsheet className="mt-0.5 h-5 w-5 text-primary" />
                  <div className="flex-1 min-w-0">
                    <div className="truncate font-medium">{f.original_name}</div>
                    <div className="text-xs text-muted-foreground">
                      {(f.size_bytes / 1024 / 1024).toFixed(2)} MB
                      {meta?.sheets?.[0] ? ` · ${meta.sheets[0].rowCount} rows · ${meta.sheets[0].headers.length} cols` : ""}
                    </div>
                    {meta?.sheets?.[0]?.headers?.length ? (
                      <div className="mt-1 flex flex-wrap gap-1">
                        {meta.sheets[0].headers.slice(0, 8).map((h) => (
                          <Badge key={h} variant="outline" className="text-[10px] font-normal">
                            {h}
                          </Badge>
                        ))}
                        {meta.sheets[0].headers.length > 8 ? (
                          <Badge variant="outline" className="text-[10px] font-normal">
                            +{meta.sheets[0].headers.length - 8} more
                          </Badge>
                        ) : null}
                      </div>
                    ) : null}
                  </div>
                </li>
              );
            })}
          </ul>
        </CardContent>
      </Card>

      {job.ai_plan ? (
        <Card className="mt-6">
          <CardHeader>
            <CardTitle className="text-base flex items-center gap-2">
              <Sparkles className="h-4 w-4 text-accent" /> AI plan
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-3 text-sm">
            <p>{(job.ai_plan as { summary?: string }).summary}</p>
            <ol className="ml-5 list-decimal space-y-1 font-mono text-xs">
              {((job.ai_plan as { ops?: Record<string, unknown>[] }).ops ?? []).map((op, i) => (
                <li key={i}>{JSON.stringify(op)}</li>
              ))}
            </ol>
            {((job.ai_plan as { warnings?: string[] }).warnings ?? []).map((w, i) => (
              <Alert key={i}>
                <AlertTriangle className="h-4 w-4" />
                <AlertDescription>{w}</AlertDescription>
              </Alert>
            ))}
          </CardContent>
        </Card>
      ) : null}

      {done ? (
        <Card className="mt-6 border-primary/40">
          <CardContent className="flex items-center justify-between gap-4 p-5">
            <div className="flex items-center gap-3">
              <div className="grid h-10 w-10 place-items-center rounded-md bg-primary/10">
                <CheckCircle2 className="h-5 w-5 text-primary" />
              </div>
              <div>
                <div className="font-medium">Ready to download</div>
                <div className="text-xs text-muted-foreground">{job.output_name}</div>
              </div>
            </div>
            <Button onClick={() => dl.mutate()} disabled={dl.isPending}>
              {dl.isPending ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}
              Download .xlsx
            </Button>
          </CardContent>
        </Card>
      ) : null}

      {job.stats && Object.keys(job.stats).length ? (
        <Card className="mt-6">
          <CardHeader>
            <CardTitle className="text-base">Run stats</CardTitle>
          </CardHeader>
          <CardContent>
            <pre className="overflow-x-auto rounded bg-secondary p-3 text-xs">
              {JSON.stringify(job.stats, null, 2)}
            </pre>
          </CardContent>
        </Card>
      ) : null}

      <div className="mt-8 flex flex-wrap items-center gap-3">
        <Button onClick={() => planMut.mutate()} disabled={!canPlan || planMut.isPending || isBusy} variant={job.ai_plan ? "outline" : "default"}>
          {planMut.isPending ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Sparkles className="mr-2 h-4 w-4" />}
          {job.ai_plan ? "Re-plan" : "Plan with AI"}
        </Button>
        <Button onClick={() => runMut.mutate()} disabled={!canRun || runMut.isPending || isBusy}>
          {runMut.isPending || isBusy ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Play className="mr-2 h-4 w-4" />}
          {isBusy ? job.status : "Run job"}
        </Button>
      </div>
    </div>
  );
}
