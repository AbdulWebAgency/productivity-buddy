import { createFileRoute, Link } from "@tanstack/react-router";
import { FileSpreadsheet, GitMerge, Sparkles, ShieldCheck, ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";

export const Route = createFileRoute("/")({
  component: Landing,
});

function Landing() {
  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="mx-auto flex max-w-6xl items-center justify-between px-6 py-6">
        <Link to="/" className="flex items-center gap-2 font-serif text-xl">
          <span className="grid h-8 w-8 place-items-center rounded-md bg-primary text-primary-foreground">
            <FileSpreadsheet className="h-4 w-4" />
          </span>
          Productivity Buddy
        </Link>
        <nav className="flex items-center gap-3">
          <Link to="/auth" className="text-sm text-muted-foreground hover:text-foreground">
            Sign in
          </Link>
          <Button asChild size="sm">
            <Link to="/auth">Get started</Link>
          </Button>
        </nav>
      </header>

      <main className="mx-auto max-w-6xl px-6">
        <section className="pt-16 pb-24 md:pt-24 md:pb-32">
          <div className="max-w-3xl">
            <div className="mb-6 inline-flex items-center gap-2 rounded-full border border-border bg-card px-3 py-1 text-xs text-muted-foreground">
              <Sparkles className="h-3.5 w-3.5 text-accent" />
              AI-planned, deterministically executed
            </div>
            <h1 className="text-5xl leading-[1.05] md:text-6xl">
              Excel work,
              <br />
              <span className="text-primary">without the busywork.</span>
            </h1>
            <p className="mt-6 max-w-2xl text-lg text-muted-foreground">
              Upload spreadsheets, describe what needs to happen, and get a clean{" "}
              <code className="rounded bg-secondary px-1.5 py-0.5 font-mono text-sm">.xlsx</code> back — with
              formatting, headers and formulas intact. Built for teachers, office staff and small teams juggling messy
              workbooks.
            </p>
            <div className="mt-8 flex flex-wrap gap-3">
              <Button asChild size="lg">
                <Link to="/auth">
                  Start free <ArrowRight className="ml-2 h-4 w-4" />
                </Link>
              </Button>
              <Button asChild size="lg" variant="outline">
                <a href="#capabilities">See what it does</a>
              </Button>
            </div>
          </div>
        </section>

        <section id="capabilities" className="border-t border-border py-20">
          <h2 className="text-3xl md:text-4xl">Capabilities</h2>
          <p className="mt-3 max-w-2xl text-muted-foreground">
            The AI reasons about what to do. The engine does the work deterministically — no invented cells, no LLM
            hallucinations in your data.
          </p>
          <div className="mt-10 grid gap-6 md:grid-cols-2 lg:grid-cols-3">
            {CAPABILITIES.map((c) => (
              <div
                key={c.title}
                className="rounded-2xl border border-border bg-card p-6 transition-shadow hover:shadow-sm"
              >
                <c.icon className="h-5 w-5 text-primary" />
                <h3 className="mt-4 text-lg font-semibold">{c.title}</h3>
                <p className="mt-2 text-sm text-muted-foreground">{c.body}</p>
              </div>
            ))}
          </div>
        </section>

        <section className="border-t border-border py-20">
          <div className="grid gap-10 md:grid-cols-2 md:items-center">
            <div>
              <h2 className="text-3xl md:text-4xl">Your files, your rules.</h2>
              <p className="mt-4 text-muted-foreground">
                Files are stored privately per user, processed in-region and deleted on request. Row-level security
                means no one else can see your workbooks — not even other Productivity Buddy users.
              </p>
              <div className="mt-6 flex items-center gap-3 text-sm text-muted-foreground">
                <ShieldCheck className="h-5 w-5 text-primary" />
                Private storage · Signed downloads · Auditable job history
              </div>
            </div>
            <div className="rounded-2xl border border-border bg-card p-6 font-mono text-sm shadow-sm">
              <div className="mb-3 text-xs uppercase tracking-wide text-muted-foreground">Example job</div>
              <pre className="whitespace-pre-wrap text-foreground">{`intent:
  "Merge Class A and Class B by
   Registration Number, flag rows
   only in one file, remove
   duplicates."

plan:
  1. merge → key: RegNo
  2. dedupe → strategy: key
  3. highlight → unmatched rows
output: result.xlsx  (styles kept)`}</pre>
            </div>
          </div>
        </section>
      </main>

      <footer className="border-t border-border">
        <div className="mx-auto flex max-w-6xl flex-col items-start justify-between gap-4 px-6 py-8 text-sm text-muted-foreground md:flex-row md:items-center">
          <div>© {new Date().getFullYear()} Productivity Buddy</div>
          <div className="flex items-center gap-1 font-serif text-base text-foreground">
            <FileSpreadsheet className="h-4 w-4" /> Productivity Buddy
          </div>
        </div>
      </footer>
    </div>
  );
}

const CAPABILITIES = [
  {
    icon: GitMerge,
    title: "Merge workbooks",
    body: "Join multiple files on a shared key (Registration Number, Email, Order ID). Auto-detects likely matches.",
  },
  {
    icon: FileSpreadsheet,
    title: "Preserve formatting",
    body: "Styles, headers, column widths, merged cells and number formats survive the round-trip.",
  },
  {
    icon: Sparkles,
    title: "Explain formulas",
    body: "Paste a broken formula, get a plain-language explanation plus a proposed fix you can accept.",
  },
  {
    icon: Sparkles,
    title: "Dedupe & diff",
    body: "Find duplicate rows by key or full content. Compare two sheets and see added, removed and changed rows.",
  },
  {
    icon: Sparkles,
    title: "Highlights & summaries",
    body: "Conditional formatting for outliers, generated summary sheet with counts, totals and basic charts.",
  },
  {
    icon: Sparkles,
    title: "Recalculated on export",
    body: "Common formulas (SUM, IF, VLOOKUP, arithmetic) are recomputed before download so cells aren't blank.",
  },
];
