create table if not exists public.security_events (
  id bigint generated always as identity primary key,
  kind text not null,
  route text not null default '',
  actor_hash text not null default '',
  detail text not null default '',
  created_at timestamptz not null default now()
);
create index if not exists security_events_created_idx on public.security_events(created_at desc);
create index if not exists security_events_kind_idx on public.security_events(kind,created_at desc);
alter table public.security_events enable row level security;
revoke all on table public.security_events from anon, authenticated;
revoke all on sequence public.security_events_id_seq from anon, authenticated;
comment on table public.security_events is 'Tentativas bloqueadas; identificador de origem somente em hash diário, sem IP em texto.';
