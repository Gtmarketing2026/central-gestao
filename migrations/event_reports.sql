-- Eventos: projetos recorrentes, edicoes, snapshots consolidados e versoes de relatorio.
-- Acesso exclusivo pela Edge Function autenticada (service_role); nada direto no navegador.
create table if not exists public.event_projects (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  name text not null,
  description text,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique(client_id, name)
);

create table if not exists public.event_editions (
  id text primary key,
  project_id text not null references public.event_projects(id) on delete cascade,
  client_id text not null references public.clients(id) on delete cascade,
  name text not null,
  year integer,
  capture_start date,
  capture_end date,
  event_date date,
  sales_start date,
  sales_end date,
  status text not null default 'draft',
  config jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.event_snapshots (
  id text primary key,
  edition_id text not null references public.event_editions(id) on delete cascade,
  client_id text not null references public.clients(id) on delete cascade,
  period_start date,
  period_end date,
  metrics jsonb not null default '{}'::jsonb,
  sources jsonb not null default '{}'::jsonb,
  collected_at timestamptz not null default now()
);

create table if not exists public.event_report_versions (
  id text primary key,
  edition_id text not null references public.event_editions(id) on delete cascade,
  snapshot_id text references public.event_snapshots(id) on delete set null,
  version_no integer not null,
  status text not null default 'draft',
  title text not null,
  instructions text,
  report_data jsonb not null default '{}'::jsonb,
  created_by text,
  created_at timestamptz not null default now(),
  unique(edition_id, version_no)
);

create index if not exists idx_event_projects_client on public.event_projects(client_id, created_at desc);
create index if not exists idx_event_editions_client on public.event_editions(client_id, event_date desc);
create index if not exists idx_event_snapshots_edition on public.event_snapshots(edition_id, collected_at desc);
create index if not exists idx_event_versions_edition on public.event_report_versions(edition_id, version_no desc);

alter table public.event_projects enable row level security;
alter table public.event_editions enable row level security;
alter table public.event_snapshots enable row level security;
alter table public.event_report_versions enable row level security;
revoke all on public.event_projects, public.event_editions, public.event_snapshots, public.event_report_versions from anon, authenticated;

comment on table public.event_snapshots is 'Fotografia consolidada por edicao; revisoes do relatorio nao duplicam metricas.';
