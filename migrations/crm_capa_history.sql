-- Histórico operacional da Auditoria CAPA.
-- Não armazena o texto integral das conversas nem dados pessoais do lead.
create table if not exists public.crm_capa_audits (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  stage text not null,
  days integer not null default 30,
  min_hours integer not null default 24,
  requested integer not null default 5,
  audited integer not null default 0,
  average_score numeric not null default 0,
  aggregate jsonb not null default '{}'::jsonb,
  created_by text,
  created_at timestamptz not null default now()
);

create table if not exists public.crm_capa_cases (
  id text primary key,
  audit_id text not null references public.crm_capa_audits(id) on delete cascade,
  client_id text not null references public.clients(id) on delete cascade,
  conversation_id text not null references public.wa_conversations(id) on delete cascade,
  stage text not null,
  channel text not null default 'organico',
  score numeric not null default 0,
  diagnosis text,
  break_point text,
  themes jsonb not null default '[]'::jsonb,
  recommended_message text,
  follow_up text,
  traffic_actions jsonb not null default '[]'::jsonb,
  commercial_actions jsonb not null default '[]'::jsonb,
  process_actions jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists idx_capa_audits_client_date on public.crm_capa_audits(client_id, created_at desc);
create index if not exists idx_capa_cases_client_date on public.crm_capa_cases(client_id, created_at desc);
create index if not exists idx_capa_cases_conversation on public.crm_capa_cases(conversation_id, created_at desc);

alter table public.crm_capa_audits enable row level security;
alter table public.crm_capa_cases enable row level security;
revoke all on public.crm_capa_audits from anon;
revoke all on public.crm_capa_cases from anon;
revoke all on public.crm_capa_audits from authenticated;
revoke all on public.crm_capa_cases from authenticated;

-- Deliberadamente sem policy para anon/authenticated: o navegador nunca lê estas
-- tabelas diretamente. A função dynamic-responder valida a sessão e acessa com
-- service_role, mantendo o histórico encapsulado no backend.

comment on table public.crm_capa_cases is 'Resultado estruturado e sem transcrição integral das auditorias CAPA.';
