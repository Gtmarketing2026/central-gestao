-- Historico diario de metricas por canal (Meta/Google/TikTok/Pinterest/GA4), pra alimentar a aba "Banco de Dados",
-- exportacao CSV e conexao direta do Power BI. Preenchida por rotina automatica (cron), nao editada manualmente.
create table if not exists public.channel_metrics_daily (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  channel text not null, -- 'meta' | 'google' | 'tiktok' | 'pinterest' | 'ga4'
  date date not null,
  source_medium text not null default '', -- so preenchido pro GA4 (origem/midia); meta/google/tiktok/pinterest ficam ''
  campaign text not null default '', -- nome da campanha (meta/google = granularidade real; ga4 = sessionCampaignName)
  adset text not null default '', -- conjunto de anuncios (meta) / grupo de anuncios (google) - so meta/google por enquanto
  ad_content text not null default '', -- so preenchido pro GA4 por enquanto (conteudo do anuncio manual/utm_content)
  spend numeric default 0,
  impressions bigint default 0,
  clicks bigint default 0,
  ctr numeric default 0,
  cpm numeric default 0,
  reach bigint default 0,
  purchases numeric default 0,
  revenue numeric default 0,
  leads numeric default 0,
  conversas numeric default 0,
  video_views numeric default 0,
  engajamentos numeric default 0,
  updated_at timestamptz default now(),
  constraint channel_metrics_daily_uniq unique (client_id, channel, date, source_medium, campaign, adset, ad_content)
);
create index if not exists idx_cmd_client_date on public.channel_metrics_daily(client_id, date);
create index if not exists idx_cmd_channel_date on public.channel_metrics_daily(channel, date);
create index if not exists idx_cmd_campaign on public.channel_metrics_daily(client_id, campaign);
alter table public.channel_metrics_daily enable row level security;
drop policy if exists cmd_auth on public.channel_metrics_daily;
create policy cmd_auth on public.channel_metrics_daily for all to authenticated using (true) with check (true);
