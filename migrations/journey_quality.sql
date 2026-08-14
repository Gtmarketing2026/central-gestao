-- Qualidade de navegação coletada pelo pixel próprio.
-- Guarda apenas identificadores técnicos anônimos e dados agregáveis de navegação.
create table if not exists public.journey_quality_events (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  created_at timestamptz not null default now(),
  event_type text not null,
  session_id text,
  anon_id text,
  channel text,
  source text,
  medium text,
  campaign text,
  term text,
  content text,
  page text,
  title text,
  referrer text,
  active_seconds integer not null default 0
);

create index if not exists journey_quality_client_date_idx
  on public.journey_quality_events(client_id, created_at desc);
create index if not exists journey_quality_session_idx
  on public.journey_quality_events(client_id, session_id, created_at);

alter table public.journey_quality_events enable row level security;
drop policy if exists journey_quality_auth on public.journey_quality_events;
revoke all on public.journey_quality_events from anon;
revoke all on public.journey_quality_events from authenticated;

comment on table public.journey_quality_events is
  'Eventos anonimizados de página e tempo ativo, enviados somente após consentimento. Sem acesso anon/authenticated; leitura apenas via service_role e retorno agregado.';
