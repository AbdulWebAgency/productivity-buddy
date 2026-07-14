import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import {
  listWorkspaces,
  createWorkspace,
  deleteWorkspace,
} from "@/lib/workspace.functions";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Plus, Trash2, MessageSquare, Loader2, ArrowRight } from "lucide-react";
import { formatDistanceToNow } from "date-fns";
import { toast } from "sonner";
import { useState } from "react";

export const Route = createFileRoute("/_authenticated/app/workspaces")({
  component: WorkspacesList,
});

function WorkspacesList() {
  const fetchList = useServerFn(listWorkspaces);
  const runCreate = useServerFn(createWorkspace);
  const runDelete = useServerFn(deleteWorkspace);
  const qc = useQueryClient();
  const navigate = useNavigate();
  const [name, setName] = useState("");

  const { data, isPending } = useQuery({
    queryKey: ["workspaces"],
    queryFn: () => fetchList(),
  });

  const create = useMutation({
    mutationFn: async () => runCreate({ data: { name: name.trim() || "Untitled workspace" } }),
    onSuccess: ({ id }) => {
      setName("");
      qc.invalidateQueries({ queryKey: ["workspaces"] });
      navigate({ to: "/app/w/$workspaceId", params: { workspaceId: id } });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Failed"),
  });

  const del = useMutation({
    mutationFn: (workspaceId: string) => runDelete({ data: { workspaceId } }),
    onSuccess: () => {
      toast.success("Workspace deleted");
      qc.invalidateQueries({ queryKey: ["workspaces"] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Delete failed"),
  });

  return (
    <div>
      <div className="mb-8">
        <h1 className="text-4xl">Workspaces</h1>
        <p className="mt-1 text-muted-foreground">
          Persistent AI conversations with your uploaded spreadsheets.
        </p>
      </div>

      <Card className="mb-6">
        <CardContent className="flex gap-2 p-4">
          <Input
            placeholder="New workspace name (e.g. Semester 1 rosters)"
            value={name}
            onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") create.mutate();
            }}
          />
          <Button onClick={() => create.mutate()} disabled={create.isPending}>
            {create.isPending ? (
              <Loader2 className="mr-2 h-4 w-4 animate-spin" />
            ) : (
              <Plus className="mr-2 h-4 w-4" />
            )}
            Create
          </Button>
        </CardContent>
      </Card>

      {isPending ? (
        <div className="flex items-center gap-2 text-muted-foreground">
          <Loader2 className="h-4 w-4 animate-spin" /> Loading…
        </div>
      ) : !data || data.length === 0 ? (
        <Card>
          <CardContent className="flex flex-col items-center gap-3 py-16 text-center">
            <div className="grid h-14 w-14 place-items-center rounded-xl bg-secondary">
              <MessageSquare className="h-7 w-7 text-primary" />
            </div>
            <div>
              <h3 className="text-lg font-semibold">No workspaces yet</h3>
              <p className="mt-1 text-sm text-muted-foreground">
                Create one to start chatting with your spreadsheets.
              </p>
            </div>
          </CardContent>
        </Card>
      ) : (
        <div className="grid gap-3">
          {data.map((w) => (
            <Card key={w.id}>
              <CardContent className="flex items-center gap-4 p-4">
                <div className="grid h-10 w-10 place-items-center rounded-md bg-secondary">
                  <MessageSquare className="h-5 w-5" />
                </div>
                <Link
                  to="/app/w/$workspaceId"
                  params={{ workspaceId: w.id }}
                  className="min-w-0 flex-1"
                >
                  <div className="truncate font-medium">{w.name}</div>
                  <div className="mt-0.5 text-xs text-muted-foreground">
                    Updated {formatDistanceToNow(new Date(w.updated_at), { addSuffix: true })}
                  </div>
                </Link>
                <Button asChild variant="ghost" size="sm">
                  <Link to="/app/w/$workspaceId" params={{ workspaceId: w.id }}>
                    Open <ArrowRight className="ml-1 h-3.5 w-3.5" />
                  </Link>
                </Button>
                <Button
                  variant="ghost"
                  size="icon"
                  onClick={() => {
                    if (confirm(`Delete "${w.name}"?`)) del.mutate(w.id);
                  }}
                  disabled={del.isPending}
                >
                  <Trash2 className="h-4 w-4 text-muted-foreground" />
                </Button>
              </CardContent>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}
