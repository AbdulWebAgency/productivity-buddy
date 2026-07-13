import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import { listJobs, deleteJob } from "@/lib/excel.functions";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { FileSpreadsheet, Plus, Trash2, Loader2, ArrowRight } from "lucide-react";
import { formatDistanceToNow } from "date-fns";
import { toast } from "sonner";

export const Route = createFileRoute("/_authenticated/app/")({
  component: JobsList,
});

const STATUS_META: Record<string, { label: string; variant: "default" | "secondary" | "destructive" | "outline" }> = {
  draft: { label: "Draft", variant: "outline" },
  queued: { label: "Ready", variant: "secondary" },
  planning: { label: "Planning", variant: "secondary" },
  running: { label: "Running", variant: "secondary" },
  succeeded: { label: "Done", variant: "default" },
  failed: { label: "Failed", variant: "destructive" },
};

function JobsList() {
  const fetchJobs = useServerFn(listJobs);
  const runDelete = useServerFn(deleteJob);
  const qc = useQueryClient();
  const navigate = useNavigate();

  const { data, isPending } = useQuery({
    queryKey: ["jobs"],
    queryFn: () => fetchJobs(),
    refetchInterval: (q) => {
      const jobs = q.state.data;
      if (!jobs) return false;
      return jobs.some((j) => ["queued", "planning", "running"].includes(j.status)) ? 2500 : false;
    },
  });

  const del = useMutation({
    mutationFn: (jobId: string) => runDelete({ data: { jobId } }),
    onSuccess: () => {
      toast.success("Job deleted");
      qc.invalidateQueries({ queryKey: ["jobs"] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Delete failed"),
  });

  return (
    <div>
      <div className="mb-8 flex items-end justify-between gap-4">
        <div>
          <h1 className="text-4xl">Your jobs</h1>
          <p className="mt-1 text-muted-foreground">
            Recent Excel automations. Runs stream live status.
          </p>
        </div>
        <Button onClick={() => navigate({ to: "/app/new" })}>
          <Plus className="mr-2 h-4 w-4" /> New job
        </Button>
      </div>

      {isPending ? (
        <div className="flex items-center gap-2 text-muted-foreground">
          <Loader2 className="h-4 w-4 animate-spin" /> Loading…
        </div>
      ) : !data || data.length === 0 ? (
        <EmptyState />
      ) : (
        <div className="grid gap-3">
          {data.map((job) => {
            const meta = STATUS_META[job.status] ?? { label: job.status, variant: "outline" as const };
            return (
              <Card key={job.id}>
                <CardContent className="flex items-center gap-4 p-4">
                  <div className="grid h-10 w-10 place-items-center rounded-md bg-secondary text-secondary-foreground">
                    <FileSpreadsheet className="h-5 w-5" />
                  </div>
                  <Link
                    to="/app/$jobId"
                    params={{ jobId: job.id }}
                    className="flex-1 min-w-0"
                  >
                    <div className="truncate font-medium">{job.name}</div>
                    <div className="mt-0.5 flex items-center gap-2 text-xs text-muted-foreground">
                      <span className="capitalize">{job.kind}</span>
                      <span>·</span>
                      <span>{formatDistanceToNow(new Date(job.created_at), { addSuffix: true })}</span>
                      {job.error ? (
                        <>
                          <span>·</span>
                          <span className="truncate text-destructive">{job.error}</span>
                        </>
                      ) : null}
                    </div>
                  </Link>
                  <Badge variant={meta.variant} className="capitalize">
                    {meta.label}
                  </Badge>
                  <Button asChild variant="ghost" size="sm">
                    <Link to="/app/$jobId" params={{ jobId: job.id }}>
                      Open <ArrowRight className="ml-1 h-3.5 w-3.5" />
                    </Link>
                  </Button>
                  <Button
                    variant="ghost"
                    size="icon"
                    onClick={() => {
                      if (confirm(`Delete "${job.name}"?`)) del.mutate(job.id);
                    }}
                    disabled={del.isPending}
                  >
                    <Trash2 className="h-4 w-4 text-muted-foreground" />
                  </Button>
                </CardContent>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

function EmptyState() {
  return (
    <Card>
      <CardContent className="flex flex-col items-center gap-3 py-16 text-center">
        <div className="grid h-14 w-14 place-items-center rounded-xl bg-secondary">
          <FileSpreadsheet className="h-7 w-7 text-primary" />
        </div>
        <div>
          <h3 className="text-lg font-semibold">No jobs yet</h3>
          <p className="mt-1 text-sm text-muted-foreground">
            Upload two or more spreadsheets to merge, dedupe or diff them.
          </p>
        </div>
        <Button asChild className="mt-2">
          <Link to="/app/new">
            <Plus className="mr-2 h-4 w-4" /> Create your first job
          </Link>
        </Button>
      </CardContent>
    </Card>
  );
}
