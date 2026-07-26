import { createFileRoute, Outlet, redirect, Link, useNavigate, useRouterState } from "@tanstack/react-router";
import { supabase } from "@/integrations/supabase/client";
import { FileSpreadsheet, LogOut, MessageSquare, ListChecks } from "lucide-react";
import { Button } from "@/components/ui/button";
import { useQueryClient } from "@tanstack/react-query";

export const Route = createFileRoute("/_authenticated")({
  ssr: false,
  beforeLoad: async () => {
    const { data, error } = await supabase.auth.getUser();
    if (error || !data.user) throw redirect({ to: "/auth" });
    return { user: data.user };
  },
  component: ProtectedLayout,
});

function ProtectedLayout() {
  const { user } = Route.useRouteContext();
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const pathname = useRouterState({ select: (s) => s.location.pathname });

  async function signOut() {
    await queryClient.cancelQueries();
    queryClient.clear();
    await supabase.auth.signOut();
    navigate({ to: "/auth", replace: true });
  }

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <Link to="/app/workspaces" className="flex items-center gap-2 font-serif text-lg">
            <span className="grid h-8 w-8 place-items-center rounded-md bg-primary text-primary-foreground">
              <FileSpreadsheet className="h-4 w-4" />
            </span>
            Productivity Buddy
          </Link>
          <nav className="flex items-center gap-1">
            <Button
              asChild
              variant={pathname.startsWith("/app/workspaces") || pathname.startsWith("/app/w/") ? "secondary" : "ghost"}
              size="sm"
            >
              <Link to="/app/workspaces">
                <MessageSquare className="mr-2 h-4 w-4" />
                Workspaces
              </Link>
            </Button>
            <Button asChild variant={pathname === "/app" ? "secondary" : "ghost"} size="sm">
              <Link to="/app">
                <ListChecks className="mr-2 h-4 w-4" />
                Legacy jobs
              </Link>
            </Button>
          </nav>
          <div className="flex items-center gap-3">
            <span className="hidden text-xs text-muted-foreground md:inline">{user.email}</span>
            <Button variant="ghost" size="sm" onClick={signOut} title="Sign out">
              <LogOut className="h-4 w-4" />
            </Button>
          </div>
        </div>
      </header>
      <main className="mx-auto max-w-6xl px-6 py-8">
        <Outlet />
      </main>
    </div>
  );
}
