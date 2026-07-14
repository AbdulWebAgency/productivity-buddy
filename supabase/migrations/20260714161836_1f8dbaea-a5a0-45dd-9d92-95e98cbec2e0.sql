
CREATE TABLE public.workspaces (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL,
  name text NOT NULL DEFAULT 'Untitled workspace',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.workspaces TO authenticated;
GRANT ALL ON public.workspaces TO service_role;
ALTER TABLE public.workspaces ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Users manage own workspaces" ON public.workspaces FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());
CREATE POLICY "Admins view all workspaces" ON public.workspaces FOR SELECT TO authenticated
  USING (has_role(auth.uid(), 'admin'::app_role));
CREATE TRIGGER workspaces_updated BEFORE UPDATE ON public.workspaces
  FOR EACH ROW EXECUTE FUNCTION public.tg_set_updated_at();

CREATE TABLE public.workspace_files (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  workspace_id uuid NOT NULL REFERENCES public.workspaces(id) ON DELETE CASCADE,
  user_id uuid NOT NULL,
  storage_path text NOT NULL,
  original_name text NOT NULL,
  size_bytes bigint NOT NULL DEFAULT 0,
  sheet_meta jsonb,
  inspector jsonb,
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.workspace_files TO authenticated;
GRANT ALL ON public.workspace_files TO service_role;
ALTER TABLE public.workspace_files ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Users manage own workspace files" ON public.workspace_files FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());
CREATE INDEX workspace_files_ws_idx ON public.workspace_files(workspace_id);

CREATE TABLE public.workspace_messages (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  workspace_id uuid NOT NULL REFERENCES public.workspaces(id) ON DELETE CASCADE,
  user_id uuid NOT NULL,
  role text NOT NULL CHECK (role IN ('user','assistant','system')),
  content text NOT NULL DEFAULT '',
  tool_data jsonb,
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.workspace_messages TO authenticated;
GRANT ALL ON public.workspace_messages TO service_role;
ALTER TABLE public.workspace_messages ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Users manage own workspace messages" ON public.workspace_messages FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());
CREATE INDEX workspace_messages_ws_idx ON public.workspace_messages(workspace_id, created_at);

CREATE TABLE public.workspace_versions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  workspace_id uuid NOT NULL REFERENCES public.workspaces(id) ON DELETE CASCADE,
  user_id uuid NOT NULL,
  parent_version_id uuid REFERENCES public.workspace_versions(id) ON DELETE SET NULL,
  version_number int NOT NULL,
  label text NOT NULL DEFAULT '',
  plan jsonb,
  stats jsonb NOT NULL DEFAULT '{}'::jsonb,
  warnings jsonb NOT NULL DEFAULT '[]'::jsonb,
  output_path text,
  output_name text,
  size_bytes bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.workspace_versions TO authenticated;
GRANT ALL ON public.workspace_versions TO service_role;
ALTER TABLE public.workspace_versions ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Users manage own workspace versions" ON public.workspace_versions FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());
CREATE INDEX workspace_versions_ws_idx ON public.workspace_versions(workspace_id, version_number);
