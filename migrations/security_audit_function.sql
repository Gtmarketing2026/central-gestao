-- Funcao de auditoria de seguranca: detecta tabela sem RLS, policy liberada pro papel public/anon
-- (o mesmo tipo de furo corrigido em fix_rls_public_role_leak.sql) e view sem security_invoker
-- (o mesmo tipo de furo corrigido no vazamento das views de negocio, ver memoria da usuaria).
-- SECURITY DEFINER pra poder ler pg_catalog, mas EXECUTE e revogado de anon/authenticated -
-- so o service_role (chamado pela rotina automatica) pode rodar. Nunca exponha isso pro anon.
create or replace function public.security_audit()
returns table(check_key text, severity text, detail text)
language sql
security definer
set search_path = public
as $$
  select
    'rls_disabled:' || t.relname,
    'critical',
    'Tabela ' || t.relname || ' esta com RLS desabilitado (dados acessiveis sem nenhuma policy).'
  from pg_class t
  join pg_namespace n on n.oid = t.relnamespace
  where n.nspname = 'public' and t.relkind = 'r' and not t.relrowsecurity

  union all

  select
    'policy_public_role:' || p.tablename || ':' || p.policyname,
    'critical',
    'Tabela ' || p.tablename || ', policy "' || p.policyname || '" liberada pro papel public/anon (using: ' || coalesce(p.qual, 'true') || ').'
  from pg_policies p
  where p.schemaname = 'public'
    and (p.roles && array['public','anon']::name[])
    and coalesce(p.qual, 'true') <> 'false'

  union all

  select
    'view_no_security_invoker:' || c.relname,
    'warning',
    'View ' || c.relname || ' nao esta marcada security_invoker (roda com permissao do dono, pode vazar dado mesmo com RLS na tabela base).'
  from pg_class c
  join pg_namespace n on n.oid = c.relnamespace
  where n.nspname = 'public' and c.relkind = 'v'
    and coalesce((select option_value from pg_options_to_table(c.reloptions) where option_name = 'security_invoker'), 'false') <> 'true';
$$;

revoke all on function public.security_audit() from public;
revoke all on function public.security_audit() from anon;
revoke all on function public.security_audit() from authenticated;
grant execute on function public.security_audit() to service_role;
