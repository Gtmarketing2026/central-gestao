-- Completa o isolamento das tabelas secundárias que também guardam dados de clientes.
do $$
declare t text; pol record;
begin
  foreach t in array array['agent_chats','agent_knowledge','briefing','capi_events','checkout_events','journey_sync','order_aggregates','raiox','track_events','track_links','tracking_config','wa_agent_sessions','wa_instances'] loop
    if to_regclass('public.'||t) is null then continue; end if;
    for pol in select policyname from pg_policies where schemaname='public' and tablename=t and roles @> array['authenticated']::name[] loop
      execute format('drop policy if exists %I on public.%I',pol.policyname,t);
    end loop;
    execute format('create policy %I on public.%I for select to authenticated using (public.app_has_aal2() and public.app_can_access_client(client_id))',t||'_read_scope',t);
    execute format('create policy %I on public.%I for all to authenticated using (public.app_has_aal2() and public.app_can_access_client(client_id) and (public.app_is_admin() or public.app_has_permission(''data.write''))) with check (public.app_has_aal2() and public.app_can_access_client(client_id) and (public.app_is_admin() or public.app_has_permission(''data.write'')))',t||'_write_scope',t);
  end loop;
end $$;

do $$
declare t text; pol record;
begin
  foreach t in array array['briefing_analise','briefing_curadoria','briefing_ficha','briefing_mensagem'] loop
    if to_regclass('public.'||t) is null then continue; end if;
    for pol in select policyname from pg_policies where schemaname='public' and tablename=t and roles @> array['authenticated']::name[] loop
      execute format('drop policy if exists %I on public.%I',pol.policyname,t);
    end loop;
    execute format('create policy %I on public.%I for select to authenticated using (public.app_has_aal2() and exists(select 1 from public.briefing b where b.id=briefing_id and public.app_can_access_client(b.client_id)))',t||'_read_scope',t);
    execute format('create policy %I on public.%I for all to authenticated using (public.app_has_aal2() and exists(select 1 from public.briefing b where b.id=briefing_id and public.app_can_access_client(b.client_id)) and (public.app_is_admin() or public.app_has_permission(''data.write''))) with check (public.app_has_aal2() and exists(select 1 from public.briefing b where b.id=briefing_id and public.app_can_access_client(b.client_id)) and (public.app_is_admin() or public.app_has_permission(''data.write'')))',t||'_write_scope',t);
  end loop;
end $$;

-- Agenda sincronizada é vinculada à tarefa; respeita o mesmo cliente da tarefa.
do $$ declare pol record; begin
  if to_regclass('public.calendar_events') is not null then
    for pol in select policyname from pg_policies where schemaname='public' and tablename='calendar_events' and roles @> array['authenticated']::name[] loop execute format('drop policy if exists %I on public.calendar_events',pol.policyname); end loop;
    create policy calendar_events_read_scope on public.calendar_events for select to authenticated
      using (public.app_has_aal2() and exists(select 1 from public.tasks t where t.id=task_id));
    create policy calendar_events_write_scope on public.calendar_events for all to authenticated
      using (public.app_has_aal2() and (public.app_is_admin() or public.app_has_permission('task.write')) and exists(select 1 from public.tasks t where t.id=task_id))
      with check (public.app_has_aal2() and (public.app_is_admin() or public.app_has_permission('task.write')) and exists(select 1 from public.tasks t where t.id=task_id));
  end if;
end $$;

-- Bibliotecas globais podem ser lidas pela equipe, mas somente administradas por quem tem permissão global.
do $$
declare t text; pol record;
begin
  foreach t in array array['report_layouts'] loop
    if to_regclass('public.'||t) is null then continue; end if;
    for pol in select policyname from pg_policies where schemaname='public' and tablename=t and roles @> array['authenticated']::name[] loop execute format('drop policy if exists %I on public.%I',pol.policyname,t); end loop;
    execute format('create policy %I on public.%I for select to authenticated using (public.app_has_aal2() and public.app_is_active())',t||'_read_active',t);
    execute format('create policy %I on public.%I for all to authenticated using (public.app_has_aal2() and (public.app_is_admin() or public.app_has_permission(''global.write''))) with check (public.app_has_aal2() and (public.app_is_admin() or public.app_has_permission(''global.write'')))',t||'_write_admin',t);
  end loop;
end $$;

-- Automações da agência são globais; gestores não administram a operação inteira.
drop policy if exists aa_all on public.andreia_automations;
drop policy if exists andreia_automations_admin on public.andreia_automations;
create policy andreia_automations_admin on public.andreia_automations for all to authenticated
using (public.app_has_aal2() and public.app_is_admin())
with check (public.app_has_aal2() and public.app_is_admin());

-- Payload bruto de webhook pode conter telefone, nome e mensagem: fica somente no backend/service role.
drop policy if exists wl_auth on public.wa_webhook_log;
revoke all on public.wa_webhook_log from authenticated;

