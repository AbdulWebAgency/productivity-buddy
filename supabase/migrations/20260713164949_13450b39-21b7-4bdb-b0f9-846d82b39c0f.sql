
-- Restrict has_role EXECUTE to authenticated only
REVOKE EXECUTE ON FUNCTION public.has_role(UUID, public.app_role) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.has_role(UUID, public.app_role) TO authenticated, service_role;

-- Storage policies: user-scoped folders {user_id}/...
CREATE POLICY "Users read own excel uploads" ON storage.objects
  FOR SELECT TO authenticated
  USING (bucket_id = 'excel-uploads' AND (storage.foldername(name))[1] = auth.uid()::text);

CREATE POLICY "Users write own excel uploads" ON storage.objects
  FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'excel-uploads' AND (storage.foldername(name))[1] = auth.uid()::text);

CREATE POLICY "Users update own excel uploads" ON storage.objects
  FOR UPDATE TO authenticated
  USING (bucket_id = 'excel-uploads' AND (storage.foldername(name))[1] = auth.uid()::text);

CREATE POLICY "Users delete own excel uploads" ON storage.objects
  FOR DELETE TO authenticated
  USING (bucket_id = 'excel-uploads' AND (storage.foldername(name))[1] = auth.uid()::text);

CREATE POLICY "Users read own excel outputs" ON storage.objects
  FOR SELECT TO authenticated
  USING (bucket_id = 'excel-outputs' AND (storage.foldername(name))[1] = auth.uid()::text);

CREATE POLICY "Users delete own excel outputs" ON storage.objects
  FOR DELETE TO authenticated
  USING (bucket_id = 'excel-outputs' AND (storage.foldername(name))[1] = auth.uid()::text);
