
-- 1. Move has_role out of the exposed public API schema into a private schema.
CREATE SCHEMA IF NOT EXISTS private;
REVOKE ALL ON SCHEMA private FROM PUBLIC;
GRANT USAGE ON SCHEMA private TO authenticated, service_role;

CREATE OR REPLACE FUNCTION private.has_role(_user_id uuid, _role public.app_role)
RETURNS boolean
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = public
AS $$
  SELECT EXISTS (SELECT 1 FROM public.user_roles WHERE user_id = _user_id AND role = _role);
$$;

REVOKE ALL ON FUNCTION private.has_role(uuid, public.app_role) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION private.has_role(uuid, public.app_role) TO authenticated, service_role;

-- Update policies to reference the private function
DROP POLICY IF EXISTS "Admins view all jobs" ON public.excel_jobs;
CREATE POLICY "Admins view all jobs" ON public.excel_jobs
  FOR SELECT USING (private.has_role(auth.uid(), 'admin'::public.app_role));

DROP POLICY IF EXISTS "Admins view all job files" ON public.excel_job_files;
CREATE POLICY "Admins view all job files" ON public.excel_job_files
  FOR SELECT USING (private.has_role(auth.uid(), 'admin'::public.app_role));

DROP POLICY IF EXISTS "Admins view all workspaces" ON public.workspaces;
CREATE POLICY "Admins view all workspaces" ON public.workspaces
  FOR SELECT USING (private.has_role(auth.uid(), 'admin'::public.app_role));

-- Drop the public-schema copy so it is no longer reachable via PostgREST.
DROP FUNCTION IF EXISTS public.has_role(uuid, public.app_role);

-- 2. Add owner-scoped INSERT/UPDATE storage policies for excel-outputs bucket.
CREATE POLICY "Users write own excel outputs"
  ON storage.objects FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'excel-outputs' AND (storage.foldername(name))[1] = auth.uid()::text);

CREATE POLICY "Users update own excel outputs"
  ON storage.objects FOR UPDATE TO authenticated
  USING (bucket_id = 'excel-outputs' AND (storage.foldername(name))[1] = auth.uid()::text)
  WITH CHECK (bucket_id = 'excel-outputs' AND (storage.foldername(name))[1] = auth.uid()::text);
