-- Custos operacionais e telemetria agregada do sistema.
create table if not exists public.system_cost_services (
  service_key text primary key,
  name text not null,
  category text not null default 'sistema',
  active boolean not null default true,
  monthly_fixed numeric not null default 0,
  input_cost_per_million numeric not null default 0,
  output_cost_per_million numeric not null default 0,
  allocation_mode text not null default 'equal' check (allocation_mode in ('equal','usage','direct','overhead')),
  client_id text references public.clients(id) on delete set null,
  notes text,
  updated_at timestamptz not null default now()
);

insert into public.system_cost_services(service_key,name,category,allocation_mode,notes) values
 ('gemini','Gemini / IA','ia','usage','Informe os preços de entrada e saída por 1 milhão de tokens.'),
 ('supabase','Supabase','infraestrutura','usage','Plano, banco, armazenamento e Edge Functions.'),
 ('uazapi','WhatsApp / UAZAPI','crm','usage','Instâncias e infraestrutura do CRM.'),
 ('apify','Apify','criativos','usage','Mineração de referências externas.'),
 ('github','GitHub','infraestrutura','equal','Código e GitHub Pages.'),
 ('cloudflare','Cloudflare','infraestrutura','equal','DNS, proxy e páginas institucionais.'),
 ('vercel','Vercel','infraestrutura','equal','Hospedagem e serviços web, quando aplicável.'),
 ('dominios','Domínios','infraestrutura','equal','Rateio mensal do custo anual dos domínios.')
on conflict (service_key) do nothing;

create table if not exists public.system_usage_events (
  id bigint generated always as identity primary key,
  occurred_at timestamptz not null default now(),
  client_id text references public.clients(id) on delete set null,
  service_key text not null,
  action text not null default 'request',
  input_units bigint not null default 0,
  output_units bigint not null default 0,
  quantity numeric not null default 1,
  meta jsonb not null default '{}'::jsonb
);
create index if not exists idx_system_usage_events_date on public.system_usage_events(occurred_at desc);
create index if not exists idx_system_usage_events_client on public.system_usage_events(client_id,occurred_at desc);

alter table public.system_cost_services enable row level security;
drop policy if exists system_cost_services_auth on public.system_cost_services;
create or replace function public.is_system_cost_admin()
returns boolean language sql security definer stable set search_path=public as $$
  select exists(select 1 from public.secure_credential_admins a where lower(a.email)=lower(coalesce(auth.jwt()->>'email','')));
$$;
revoke all on function public.is_system_cost_admin() from public,anon;
grant execute on function public.is_system_cost_admin() to authenticated;
create policy system_cost_services_auth on public.system_cost_services for all to authenticated using (public.is_system_cost_admin()) with check (public.is_system_cost_admin());
revoke all on public.system_cost_services from anon;
grant select,insert,update,delete on public.system_cost_services to authenticated;

alter table public.system_usage_events enable row level security;
revoke all on public.system_usage_events from anon, authenticated;

create or replace function public.system_cost_usage_summary(p_from date, p_to date)
returns table(
  client_id text, client_name text, client_status text,
  messages bigint, conversations bigint, ai_classifications bigint,
  andreia_answers bigint, briefing_analyses bigint, creative_analyses bigint,
  capa_cases bigint, tracked_events bigint, memories bigint
)
language sql security definer stable set search_path=public as $$
  with
  msg as (select client_id,count(*)::bigint n from wa_messages where ts>=p_from::timestamptz and ts<(p_to+1)::timestamptz group by client_id),
  conv as (select client_id,count(*)::bigint n from wa_conversations where last_at>=p_from::timestamptz and last_at<(p_to+1)::timestamptz group by client_id),
  cls as (select client_id,count(*)::bigint n from wa_journey where source='ia' and created_at>=p_from::timestamptz and created_at<(p_to+1)::timestamptz group by client_id),
  ans as (select x.client_id,count(*) filter(where m.item->>'role'='assistant')::bigint n from agent_chats x cross join lateral jsonb_array_elements(coalesce(x.messages,'[]'::jsonb)) m(item) where x.updated_at>=p_from::timestamptz and x.updated_at<(p_to+1)::timestamptz group by x.client_id),
  bri as (select b.client_id,count(*)::bigint n from briefing_analise a join briefing b on b.id=a.briefing_id where a.gerado_em>=p_from::timestamptz and a.gerado_em<(p_to+1)::timestamptz group by b.client_id),
  cre as (select client_id,count(*)::bigint n from creative_miner_items where status='analyzed' and updated_at>=p_from::timestamptz and updated_at<(p_to+1)::timestamptz group by client_id),
  capa as (select client_id,coalesce(sum(audited),0)::bigint n from crm_capa_audits where created_at>=p_from::timestamptz and created_at<(p_to+1)::timestamptz group by client_id),
  trk as (select client_id,count(*)::bigint n from lead_touchpoints where ts>=p_from::timestamptz and ts<(p_to+1)::timestamptz group by client_id),
  mem as (select client_id,count(*)::bigint n from andreia_memory where active=true group by client_id)
  select c.id,c.name,c.status,coalesce(msg.n,0),coalesce(conv.n,0),coalesce(cls.n,0),coalesce(ans.n,0),coalesce(bri.n,0),coalesce(cre.n,0),coalesce(capa.n,0),coalesce(trk.n,0),coalesce(mem.n,0)
  from clients c left join msg on msg.client_id=c.id left join conv on conv.client_id=c.id left join cls on cls.client_id=c.id left join ans on ans.client_id=c.id left join bri on bri.client_id=c.id left join cre on cre.client_id=c.id left join capa on capa.client_id=c.id left join trk on trk.client_id=c.id left join mem on mem.client_id=c.id
  where c.status<>'Encerrado' and public.is_system_cost_admin() order by c.name;
$$;

create or replace function public.system_ai_usage_summary(p_from date, p_to date)
returns table(requests bigint,input_tokens bigint,output_tokens bigint,database_bytes bigint)
language sql security definer stable set search_path=public as $$
  select count(*),coalesce(sum(input_units),0)::bigint,coalesce(sum(output_units),0)::bigint,pg_database_size(current_database())
  from system_usage_events where public.is_system_cost_admin() and service_key in ('gemini','openai') and occurred_at>=p_from::timestamptz and occurred_at<(p_to+1)::timestamptz;
$$;

revoke all on function public.system_cost_usage_summary(date,date) from public,anon;
revoke all on function public.system_ai_usage_summary(date,date) from public,anon;
grant execute on function public.system_cost_usage_summary(date,date) to authenticated;
grant execute on function public.system_ai_usage_summary(date,date) to authenticated;
