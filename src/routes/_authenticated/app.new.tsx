import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useState } from "react";
import { useMutation } from "@tanstack/react-query";
import { useServerFn } from "@tanstack/react-start";
import { supabase } from "@/integrations/supabase/client";
import { createJob, registerJobFiles } from "@/lib/excel.functions";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Upload, X, Loader2, ArrowRight, FileSpreadsheet } from "lucide-react";
import { toast } from "sonner";

export const Route = createFileRoute("/_authenticated/app/new")({
  component: NewJob,
});

const MAX_FILE_MB = 25;

function NewJob() {
  const navigate = useNavigate();
  const [name, setName] = useState("");
  const [intent, setIntent] = useState("");
  const [files, setFiles] = useState<File[]>([]);
  const createJobFn = useServerFn(createJob);
  const registerFn = useServerFn(registerJobFiles);

  const create = useMutation({
    mutationFn: async () => {
      if (files.length === 0) throw new Error("Add at least one spreadsheet.");
      const { data: userData } = await supabase.auth.getUser();
      if (!userData.user) throw new Error("Not signed in");

      const { jobId } = await createJobFn({
        data: { name: name.trim() || "Untitled job", intent: intent.trim() || undefined, kind: "auto" },
      });

      const uploaded: { storagePath: string; originalName: string; sizeBytes: number }[] = [];
      for (const [i, file] of files.entries()) {
        const safe = file.name.replace(/[^a-z0-9.\-_ ]/gi, "");
        const path = `${userData.user.id}/${jobId}/input/${String(i).padStart(2, "0")}-${safe}`;
        const { error } = await supabase.storage.from("excel-uploads").upload(path, file, {
          upsert: true,
          contentType: file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        if (error) throw new Error(`Upload ${file.name}: ${error.message}`);
        uploaded.push({ storagePath: path, originalName: file.name, sizeBytes: file.size });
      }

      await registerFn({ data: { jobId, files: uploaded } });
      return jobId;
    },
    onSuccess: (jobId) => {
      toast.success("Files uploaded — planning next.");
      navigate({ to: "/app/$jobId", params: { jobId } });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : "Upload failed"),
  });

  function onPickFiles(list: FileList | null) {
    if (!list) return;
    const chosen = Array.from(list);
    const invalid = chosen.find((f) => f.size > MAX_FILE_MB * 1024 * 1024);
    if (invalid) {
      toast.error(`${invalid.name} exceeds ${MAX_FILE_MB} MB limit`);
      return;
    }
    setFiles((prev) => [...prev, ...chosen].slice(0, 10));
  }

  return (
    <div className="mx-auto max-w-2xl">
      <h1 className="text-4xl">New job</h1>
      <p className="mt-1 text-muted-foreground">
        Upload workbooks, describe what should happen, and Ledgerly plans the ops.
      </p>

      <Card className="mt-8">
        <CardHeader>
          <CardTitle>Details</CardTitle>
        </CardHeader>
        <CardContent className="space-y-6">
          <div className="space-y-2">
            <Label htmlFor="name">Job name</Label>
            <Input
              id="name"
              placeholder="Class A + B roster merge"
              value={name}
              onChange={(e) => setName(e.target.value)}
            />
          </div>

          <div className="space-y-2">
            <Label htmlFor="intent">What should happen?</Label>
            <Textarea
              id="intent"
              rows={4}
              placeholder="Merge these files by Registration Number, remove duplicates, flag rows only in one file, and add a summary sheet."
              value={intent}
              onChange={(e) => setIntent(e.target.value)}
            />
            <p className="text-xs text-muted-foreground">
              Plain English. The AI will map this to concrete operations.
            </p>
          </div>

          <div className="space-y-2">
            <Label>Spreadsheets</Label>
            <label
              htmlFor="file-input"
              className="flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed border-border bg-secondary/40 p-8 text-center transition-colors hover:bg-secondary"
            >
              <Upload className="h-6 w-6 text-primary" />
              <div className="text-sm font-medium">Click to upload .xlsx files</div>
              <div className="text-xs text-muted-foreground">
                Up to 10 files · {MAX_FILE_MB} MB each
              </div>
              <input
                id="file-input"
                type="file"
                accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                multiple
                className="hidden"
                onChange={(e) => {
                  onPickFiles(e.target.files);
                  e.target.value = "";
                }}
              />
            </label>

            {files.length > 0 && (
              <ul className="mt-2 divide-y rounded-md border border-border">
                {files.map((f, i) => (
                  <li key={i} className="flex items-center gap-3 px-3 py-2 text-sm">
                    <FileSpreadsheet className="h-4 w-4 text-primary" />
                    <span className="flex-1 truncate">{f.name}</span>
                    <span className="text-xs text-muted-foreground">
                      {(f.size / 1024 / 1024).toFixed(2)} MB
                    </span>
                    <button
                      type="button"
                      onClick={() => setFiles((prev) => prev.filter((_, j) => j !== i))}
                      className="text-muted-foreground hover:text-destructive"
                    >
                      <X className="h-4 w-4" />
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </div>

          <Button
            onClick={() => create.mutate()}
            disabled={create.isPending || files.length === 0}
            className="w-full"
            size="lg"
          >
            {create.isPending ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" /> Uploading…
              </>
            ) : (
              <>
                Continue <ArrowRight className="ml-2 h-4 w-4" />
              </>
            )}
          </Button>
        </CardContent>
      </Card>
    </div>
  );
}
