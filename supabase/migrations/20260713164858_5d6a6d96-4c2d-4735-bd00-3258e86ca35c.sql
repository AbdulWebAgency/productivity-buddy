
-- Roles
CREATE TYPE public.app_role AS ENUM ('admin', 'user');

CREATE TABLE public.user_roles (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  role public.app_role NOT NULL DEFAULT 'user',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (user_id, role)
);

GRANT SELECT ON public.user_roles TO authenticated;
GRANT ALL ON public.user_roles TO service_role;
ALTER TABLE public.user_roles ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users read own roles" ON public.user_roles
  FOR SELECT TO authenticated
  USING (user_id = auth.uid());

CREATE OR REPLACE FUNCTION public.has_role(_user_id UUID, _role public.app_role)
RETURNS BOOLEAN
LANGUAGE SQL STABLE SECURITY DEFINER SET search_path = public
AS $$
  SELECT EXISTS (SELECT 1 FROM public.user_roles WHERE user_id = _user_id AND role = _role);
$$;

-- Job status + kind enums
CREATE TYPE public.job_status AS ENUM ('draft', 'queued', 'planning', 'running', 'succeeded', 'failed');
CREATE TYPE public.job_kind AS ENUM ('merge', 'dedupe', 'diff', 'format', 'summary', 'auto');

-- Jobs
CREATE TABLE public.excel_jobs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  name TEXT NOT NULL DEFAULT 'Untitled job',
  kind public.job_kind NOT NULL DEFAULT 'auto',
  status public.job_status NOT NULL DEFAULT 'draft',
  intent TEXT,
  params JSONB NOT NULL DEFAULT '{}'::jsonb,
  ai_plan JSONB,
  warnings JSONB NOT NULL DEFAULT '[]'::jsonb,
  stats JSONB NOT NULL DEFAULT '{}'::jsonb,
  error TEXT,
  output_path TEXT,
  output_name TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  completed_at TIMESTAMPTZ
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.excel_jobs TO authenticated;
GRANT ALL ON public.excel_jobs TO service_role;
ALTER TABLE public.excel_jobs ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users manage own jobs" ON public.excel_jobs
  FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());

CREATE POLICY "Admins view all jobs" ON public.excel_jobs
  FOR SELECT TO authenticated
  USING (public.has_role(auth.uid(), 'admin'));

CREATE INDEX excel_jobs_user_created_idx ON public.excel_jobs(user_id, created_at DESC);

-- Job files
CREATE TYPE public.file_role AS ENUM ('input', 'output');

CREATE TABLE public.excel_job_files (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  job_id UUID NOT NULL REFERENCES public.excel_jobs(id) ON DELETE CASCADE,
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  role public.file_role NOT NULL DEFAULT 'input',
  storage_path TEXT NOT NULL,
  original_name TEXT NOT NULL,
  size_bytes BIGINT NOT NULL DEFAULT 0,
  sheet_meta JSONB,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.excel_job_files TO authenticated;
GRANT ALL ON public.excel_job_files TO service_role;
ALTER TABLE public.excel_job_files ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users manage own job files" ON public.excel_job_files
  FOR ALL TO authenticated
  USING (user_id = auth.uid()) WITH CHECK (user_id = auth.uid());

CREATE POLICY "Admins view all job files" ON public.excel_job_files
  FOR SELECT TO authenticated
  USING (public.has_role(auth.uid(), 'admin'));

CREATE INDEX excel_job_files_job_idx ON public.excel_job_files(job_id);

-- updated_at trigger
CREATE OR REPLACE FUNCTION public.tg_set_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql AS $$
BEGIN NEW.updated_at = now(); RETURN NEW; END $$;

CREATE TRIGGER excel_jobs_set_updated_at
BEFORE UPDATE ON public.excel_jobs
FOR EACH ROW EXECUTE FUNCTION public.tg_set_updated_at();
