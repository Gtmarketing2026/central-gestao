-- Defesa em profundidade: o frontend sempre opera como `authenticated`.
-- O rastreamento público escreve exclusivamente por Edge Function com service_role.
-- Portanto o papel `anon` não precisa de privilégios diretos em tabelas, sequências ou funções da aplicação.
do $$
declare r record;
begin
  for r in select schemaname, tablename from pg_tables where schemaname in ('public','midia') loop
    execute format('revoke all privileges on table %I.%I from anon', r.schemaname, r.tablename);
  end loop;
  for r in select sequence_schema, sequence_name from information_schema.sequences where sequence_schema in ('public','midia') loop
    execute format('revoke all privileges on sequence %I.%I from anon', r.sequence_schema, r.sequence_name);
  end loop;
  for r in
    select n.nspname schema_name, p.proname function_name, pg_get_function_identity_arguments(p.oid) args
    from pg_proc p join pg_namespace n on n.oid=p.pronamespace
    where n.nspname in ('public','midia')
  loop
    execute format('revoke execute on function %I.%I(%s) from anon', r.schema_name, r.function_name, r.args);
  end loop;
end $$;

-- Evita que novas funções sejam executáveis pelo público por padrão.
alter default privileges in schema public revoke execute on functions from public, anon;
alter default privileges in schema midia revoke execute on functions from public, anon;
alter default privileges in schema public revoke all on tables from anon;
alter default privileges in schema midia revoke all on tables from anon;
alter default privileges in schema public revoke all on sequences from anon;
alter default privileges in schema midia revoke all on sequences from anon;

comment on schema public is 'Acesso anônimo direto revogado em 2026-08-12; entradas públicas somente por Edge Functions validadas.';
