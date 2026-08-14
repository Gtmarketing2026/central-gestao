-- Controle de acesso da Central de Gestão: perfis, escopo por cliente e auditoria.
create table if not exists public.app_users (
  user_id uuid primary key references auth.users(id) on delete restrict,
  email text not null,
  name text not null default '',
  role text not null default 'gestor' check (role in ('master','admin','gestor')),
  status text not null default 'active' check (status in ('invited','active','inactive')),
  all_clients boolean not null default false,
  permissions jsonb not null default '{"menus":[],"actions":[]}'::jsonb,
  mfa_required boolean not null default true,
  protected boolean not null default false,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists app_users_email_lower_idx on public.app_users(lower(email));

create table if not exists public.app_user_clients (
  user_id uuid not null references public.app_users(user_id) on delete cascade,
  client_id text not null references public.clients(id) on delete cascade,
  created_at timestamptz not null default now(),
  primary key (user_id, client_id)
);

create table if not exists public.app_access_audit (
  id bigint generated always as identity primary key,
  actor_user_id uuid references auth.users(id) on delete set null,
  target_user_id uuid references auth.users(id) on delete set null,
  action text not null,
  before_data jsonb,
  after_data jsonb,
  created_at timestamptz not null default now()
);

alter table public.app_users enable row level security;
alter table public.app_user_clients enable row level security;
alter table public.app_access_audit enable row level security;

create or replace function public.app_current_role()
returns text language sql stable security definer set search_path=public
as $$ select coalesce((select role from public.app_users where user_id=auth.uid() and status='active'),'') $$;

create or replace function public.app_is_active()
returns boolean language sql stable security definer set search_path=public
as $$ select exists(select 1 from public.app_users where user_id=auth.uid() and status='active') $$;

create or replace function public.app_is_master()
returns boolean language sql stable security definer set search_path=public
as $$ select public.app_current_role()='master' $$;

create or replace function public.app_is_admin()
returns boolean language sql stable security definer set search_path=public
as $$ select public.app_current_role() in ('master','admin') $$;

create or replace function public.app_has_aal2()
returns boolean language sql stable
as $$ select coalesce(auth.jwt()->>'aal','')='aal2' $$;

create or replace function public.app_has_permission(code text)
returns boolean language sql stable security definer set search_path=public
as $$
  select exists(
    select 1 from public.app_users u
    where u.user_id=auth.uid() and u.status='active'
      and (
        u.role in ('master','admin')
        or coalesce(u.permissions->'actions','[]'::jsonb) ? '*'
        or coalesce(u.permissions->'actions','[]'::jsonb) ? code
      )
  )
$$;

create or replace function public.app_can_access_client(cid text)
returns boolean language sql stable security definer set search_path=public
as $$
  select exists(
    select 1 from public.app_users u
    where u.user_id=auth.uid() and u.status='active'
      and (
        u.role in ('master','admin') or u.all_clients
        or exists(select 1 from public.app_user_clients uc where uc.user_id=u.user_id and uc.client_id=cid)
      )
  )
$$;

revoke all on function public.app_current_role() from public;
revoke all on function public.app_is_active() from public;
revoke all on function public.app_is_master() from public;
revoke all on function public.app_is_admin() from public;
revoke all on function public.app_has_aal2() from public;
revoke all on function public.app_has_permission(text) from public;
revoke all on function public.app_can_access_client(text) from public;
grant execute on function public.app_current_role() to authenticated;
grant execute on function public.app_is_active() to authenticated;
grant execute on function public.app_is_master() to authenticated;
grant execute on function public.app_is_admin() to authenticated;
grant execute on function public.app_has_aal2() to authenticated;
grant execute on function public.app_has_permission(text) to authenticated;
grant execute on function public.app_can_access_client(text) to authenticated;

drop policy if exists app_users_self on public.app_users;
create policy app_users_self on public.app_users for select to authenticated
using (user_id=auth.uid() or public.app_is_master());
drop policy if exists app_user_clients_self on public.app_user_clients;
create policy app_user_clients_self on public.app_user_clients for select to authenticated
using (user_id=auth.uid() or public.app_is_master());
drop policy if exists app_access_audit_master on public.app_access_audit;
create policy app_access_audit_master on public.app_access_audit for select to authenticated
using (public.app_is_master() and public.app_has_aal2());

-- O primeiro usuário é imutável e protegido. O usuário já existente permanece administrador.
insert into public.app_users(user_id,email,name,role,status,all_clients,permissions,mfa_required,protected)
select id,email,'Andreia','master','active',true,'{"menus":["*"],"actions":["*"]}'::jsonb,true,true
from auth.users where lower(email)='contato@cursosonlinevrrj.com.br'
on conflict(user_id) do update set role='master',status='active',all_clients=true,mfa_required=true,protected=true,updated_at=now();

insert into public.app_users(user_id,email,name,role,status,all_clients,permissions,mfa_required,protected)
select id,email,coalesce(raw_user_meta_data->>'name','Dionathan'),'admin','active',true,'{"menus":["*"],"actions":["*"]}'::jsonb,true,false
from auth.users where lower(email)='dionymorr@gmail.com'
on conflict(user_id) do nothing;

-- Proteção contra remoção/rebaixamento do master, inclusive por SQL acidental.
create or replace function public.protect_master_user()
returns trigger language plpgsql security definer set search_path=public as $$
begin
  if old.protected and (tg_op='DELETE' or new.role<>'master' or new.status<>'active' or not new.protected) then
    raise exception 'O usuário master protegido não pode ser removido, rebaixado ou desativado.';
  end if;
  if tg_op='UPDATE' then new.updated_at=now(); end if;
  return case when tg_op='DELETE' then old else new end;
end $$;
drop trigger if exists trg_protect_master_user on public.app_users;
create trigger trg_protect_master_user before update or delete on public.app_users
for each row execute function public.protect_master_user();

grant select on public.app_users,public.app_user_clients,public.app_access_audit to authenticated;
revoke insert,update,delete on public.app_users,public.app_user_clients,public.app_access_audit from authenticated;

-- Restringe as tabelas principais por cliente. Service role continua com bypass para webhooks e crons.
drop policy if exists auth on public.clients;
drop policy if exists clients_read_scope on public.clients;
drop policy if exists clients_write_scope on public.clients;
create policy clients_read_scope on public.clients for select to authenticated
using (public.app_has_aal2() and public.app_can_access_client(id));
create policy clients_write_scope on public.clients for all to authenticated
using (public.app_has_aal2() and public.app_can_access_client(id) and (public.app_is_admin() or public.app_has_permission('clients.write')))
with check (public.app_has_aal2() and public.app_can_access_client(id) and (public.app_is_admin() or public.app_has_permission('clients.write')));

do $$
declare t text; col text; pol record;
begin
  foreach t in array array['tasks','finance','shows','wallet','wa_conversations','wa_messages','wa_journey','lead_people','lead_identities','lead_touchpoints','rd_conversions','report_analysis','creative_miner_items','channel_metrics_daily'] loop
    if to_regclass('public.'||t) is null then continue; end if;
    col := case when exists(select 1 from information_schema.columns where table_schema='public' and table_name=t and column_name='client_id') then 'client_id' else 'client' end;
    for pol in select policyname from pg_policies where schemaname='public' and tablename=t and roles @> array['authenticated']::name[] loop
      execute format('drop policy if exists %I on public.%I',pol.policyname,t);
    end loop;
    execute format('create policy %I on public.%I for select to authenticated using (public.app_has_aal2() and public.app_can_access_client(%I))',t||'_read_scope',t,col);
    execute format('create policy %I on public.%I for all to authenticated using (public.app_has_aal2() and public.app_can_access_client(%I) and (public.app_is_admin() or public.app_has_permission(''data.write''))) with check (public.app_has_aal2() and public.app_can_access_client(%I) and (public.app_is_admin() or public.app_has_permission(''data.write'')))',t||'_write_scope',t,col,col);
  end loop;
end $$;

-- Dados globais operacionais: leitura para usuários ativos com 2FA; alteração controlada.
do $$
declare t text; pol record;
begin
  foreach t in array array['team','report_templates','notifications','creditors'] loop
    if to_regclass('public.'||t) is null then continue; end if;
    for pol in select policyname from pg_policies where schemaname='public' and tablename=t and roles @> array['authenticated']::name[] loop
      execute format('drop policy if exists %I on public.%I',pol.policyname,t);
    end loop;
    execute format('create policy %I on public.%I for select to authenticated using (public.app_has_aal2() and public.app_is_active())',t||'_read_active',t);
    execute format('create policy %I on public.%I for all to authenticated using (public.app_has_aal2() and (public.app_is_admin() or public.app_has_permission(''global.write''))) with check (public.app_has_aal2() and (public.app_is_admin() or public.app_has_permission(''global.write'')))',t||'_write_admin',t);
  end loop;
end $$;

-- Configuração global contém segredos legados: somente o master em AAL2 lê ou altera pelo navegador.
drop policy if exists auth on public.account_config;
drop policy if exists account_config_master on public.account_config;
create policy account_config_master on public.account_config for all to authenticated
using (public.app_is_master() and public.app_has_aal2())
with check (public.app_is_master() and public.app_has_aal2());
