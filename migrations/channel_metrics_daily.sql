-- Historico diario de metricas por canal (Meta/Google/TikTok/Pinterest), pra alimentar a aba "Banco de Dados",
-- exportacao CSV e conexao direta do Power BI. Preenchida por rotina automatica (cron), nao editada manualmente.
create table if not exists public.channel_metrics_daily (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  channel text not null, -- 'meta' | 'google' | 'tiktok' | 'pinterest'
  date date not null,
  spend numeric default 0,
  impressions bigint default 0,
  clicks bigint default 0,
  reach bigint default 0,
  purchases numeric default 0,
  revenue numeric default 0,
  leads numeric default 0,
  conversas numeric default 0,
  video_views numeric default 0,
  engajamentos numeric default 0,
  updated_at timestamptz default now(),
  unique(client_id, channel, date)
);
create index if not exists idx_cmd_client_date on public.channel_metrics_daily(client_id, date);
create index if not exists idx_cmd_channel_date on public.channel_metrics_daily(channel, date);
alter table public.channel_metrics_daily enable row level security;
drop policy if exists cmd_auth on public.channel_metrics_daily;
create policy cmd_auth on public.channel_metrics_daily for all to authenticated using (true) with check (true);
