-- Banco de dados centralizado de midia paga (schema midia): substitui/evolui channel_metrics_daily.
-- Granularidade nivel-anuncio, NULL != 0, resultado em EAV. Spec completa em docs/spec-banco-dados-midia.md.
-- Fase 1 do checklist: schema + dimensoes + fatos vazios + views que reaproveitam wa_conversations/lead_touchpoints.
-- Coletor (Fase 2+) ainda nao grava aqui - roda em paralelo com channel_metrics_daily ate validar.

create schema if not exists midia;

-- ===== DIMENSOES =====

create table if not exists midia.dim_plataforma (
  id text primary key,               -- 'meta' | 'google' | 'tiktok' | 'pinterest' | 'instagram_organico' | 'ga4' | 'whatsapp'
  nome text not null,
  tipo text not null,                -- 'ads' | 'organico' | 'analytics' | 'crm'
  ativa boolean not null default true
);
insert into midia.dim_plataforma (id, nome, tipo) values
  ('meta','Meta Ads','ads'), ('google','Google Ads','ads'),
  ('instagram_organico','Instagram Organico','organico'),
  ('ga4','Google Analytics 4','analytics'), ('whatsapp','WhatsApp','crm')
on conflict (id) do nothing;

create table if not exists midia.dim_conta (
  id uuid primary key default gen_random_uuid(),
  client_id text not null references public.clients(id) on delete cascade,
  plataforma_id text not null references midia.dim_plataforma(id),
  conta_externa_id text not null,     -- act_XXXX, customer id, ig business id, ga4 property id...
  nome text not null default '',
  moeda text not null default 'BRL',
  timezone text not null default 'America/Sao_Paulo',
  ativa boolean not null default true,
  criado_em timestamptz not null default now(),
  unique (client_id, plataforma_id, conta_externa_id)
);

create table if not exists midia.dim_campanha (
  id uuid primary key default gen_random_uuid(),
  conta_id uuid not null references midia.dim_conta(id) on delete cascade,
  campanha_externa_id text not null,
  nome text not null default '',
  objetivo_bruto text,                 -- objective (Meta) / advertising_channel_type (Google)
  objetivo_tipo text,                  -- classificado: 'trafego'|'conversao'|'engajamento'|'awareness'|...
  objetivo_rotulo text,
  atualizado_em timestamptz not null default now(),
  unique (conta_id, campanha_externa_id)
);

create table if not exists midia.dim_grupo (
  id uuid primary key default gen_random_uuid(),
  campanha_id uuid not null references midia.dim_campanha(id) on delete cascade,
  grupo_externo_id text not null,      -- adset_id (Meta) / ad_group.id (Google)
  nome text not null default '',
  otimizacao text,                     -- optimization_goal (Meta) - nullable, Google nao tem equivalente direto
  destino text,                        -- destination_type (Meta: MESSENGER/WHATSAPP/INSTAGRAM_DIRECT) - nullable
  unique (campanha_id, grupo_externo_id)
);

create table if not exists midia.dim_anuncio (
  id uuid primary key default gen_random_uuid(),
  grupo_id uuid not null references midia.dim_grupo(id) on delete cascade,
  anuncio_externo_id text not null,    -- ad_id (Meta) / ad_group_ad.ad.id (Google)
  nome text not null default '',
  formato text,                        -- media_type/ad.type quando disponivel - nullable
  thumbnail_url text,
  criado_em timestamptz not null default now(),
  unique (grupo_id, anuncio_externo_id)
);

create table if not exists midia.dim_conteudo_organico (
  id uuid primary key default gen_random_uuid(),
  conta_id uuid not null references midia.dim_conta(id) on delete cascade,
  post_externo_id text not null,
  tipo_midia text,                     -- media_type (IMAGE/VIDEO/CAROUSEL_ALBUM)
  permalink text,
  legenda text,
  publicado_em timestamptz,
  thumbnail_url text,
  unique (conta_id, post_externo_id)
);

-- ===== FATOS =====

create table if not exists midia.fact_performance (
  id uuid primary key default gen_random_uuid(),
  anuncio_id uuid not null references midia.dim_anuncio(id) on delete cascade,
  data date not null,
  moeda text not null default 'BRL',
  investimento numeric not null default 0,   -- sempre reportado pela API quando ha gasto
  impressoes bigint,                         -- NULL so se a plataforma nao reportar nada nesse dia
  cliques bigint,
  alcance bigint,                            -- NULL sempre pra Google Ads (API nao fornece)
  frequencia numeric,                        -- NULL sempre pra Google Ads (API nao fornece)
  atualizado_em timestamptz not null default now(),
  unique (anuncio_id, data)
);
create index if not exists idx_fact_performance_data on midia.fact_performance(data);

create table if not exists midia.fact_resultado (
  id uuid primary key default gen_random_uuid(),
  anuncio_id uuid not null references midia.dim_anuncio(id) on delete cascade,
  data date not null,
  tipo_resultado text not null,   -- 'compra'|'lead'|'conversa'|'add_to_cart'|'iniciar_checkout'|'cadastro'|'instalacao'|'ligacao'|'outro'
  quantidade numeric not null default 0,
  valor numeric,                  -- NULL quando o tipo de resultado nao tem valor monetario associado
  origem_deteccao text not null,  -- 'meta_actions'|'google_conversion_action'|'manual'
  moeda text not null default 'BRL',
  unique (anuncio_id, data, tipo_resultado)
);
create index if not exists idx_fact_resultado_tipo on midia.fact_resultado(tipo_resultado, data);

create table if not exists midia.fact_video (
  id uuid primary key default gen_random_uuid(),
  anuncio_id uuid not null references midia.dim_anuncio(id) on delete cascade,
  data date not null,
  visualizacoes bigint,           -- video_view generico
  thruplay bigint,                -- Meta: video_thruplay_watched_actions
  vis_3s bigint,                  -- Google: video_trueview_views usado como proxy quando thruplay ausente
  ret_25pct numeric, ret_50pct numeric, ret_75pct numeric, ret_100pct numeric,  -- NULL sempre hoje (nao coletado ainda)
  tempo_medio_assistido numeric,  -- NULL sempre hoje (nao coletado ainda)
  unique (anuncio_id, data)
);

create table if not exists midia.fact_engajamento (
  id uuid primary key default gen_random_uuid(),
  anuncio_id uuid not null references midia.dim_anuncio(id) on delete cascade,
  data date not null,
  engajamentos_total bigint,      -- post_engagement (Meta) / engagements (Google)
  unique (anuncio_id, data)
);

create table if not exists midia.fact_conteudo_organico_metricas (
  id uuid primary key default gen_random_uuid(),
  conteudo_id uuid not null references midia.dim_conteudo_organico(id) on delete cascade,
  data_coleta date not null,
  curtidas bigint, comentarios bigint, compartilhamentos bigint, salvamentos bigint,
  alcance bigint, visualizacoes bigint,     -- visualizacoes NULL pra post que nao e video/reels
  unique (conteudo_id, data_coleta)
);

create table if not exists midia.fact_seguidores_snapshot (
  id uuid primary key default gen_random_uuid(),
  conta_id uuid not null references midia.dim_conta(id) on delete cascade,
  data date not null,
  total_seguidores bigint not null,
  unique (conta_id, data)
);

create table if not exists midia.fact_analytics_ga4 (
  id uuid primary key default gen_random_uuid(),
  client_id text not null references public.clients(id) on delete cascade,
  propriedade_id text not null,
  data date not null,
  origem text not null default '',      -- sessionSource
  midia_texto text not null default '', -- sessionMedium
  campanha_texto text not null default '', -- sessionCampaignName (join fraco, por nome)
  conteudo_texto text not null default '', -- sessionManualAdContent
  sessoes bigint,
  usuarios bigint,
  compras numeric,
  receita numeric,
  moeda text not null default 'BRL',
  unique (client_id, propriedade_id, data, origem, midia_texto, campanha_texto, conteudo_texto)
);

-- ===== VIEWS (aproveitam o que ja existe, sem duplicar tabela) =====

create or replace view midia.vw_whatsapp_diario as
select
  client_id,
  date(created_at) as data,
  origin_type,
  count(*) as conversas_iniciadas,
  count(*) filter (where stage in ('mql','sql','comprou','posvenda')) as qualificados,
  count(*) filter (where stage in ('comprou','posvenda')) as vendas
from public.wa_conversations
group by client_id, date(created_at), origin_type;

create or replace view midia.vw_jornada_caminhos as
with ordenado as (
  select
    lt.client_id, lt.person_id,
    string_agg(
      '[' || coalesce(lt.channel,'direto') || '] ' || coalesce(lt.campaign, lt.label, lt.kind),
      ' -> ' order by lt.ts
    ) as caminho,
    max(lp.value) as valor_pessoa
  from public.lead_touchpoints lt
  join public.lead_people lp on lp.id = lt.person_id
  group by lt.client_id, lt.person_id
)
select client_id, caminho, count(*) as ocorrencias, sum(valor_pessoa) as valor_total
from ordenado
group by client_id, caminho
order by ocorrencias desc;

-- ===== RLS =====
alter table midia.dim_conta enable row level security;
alter table midia.dim_campanha enable row level security;
alter table midia.dim_grupo enable row level security;
alter table midia.dim_anuncio enable row level security;
alter table midia.dim_conteudo_organico enable row level security;
alter table midia.fact_performance enable row level security;
alter table midia.fact_resultado enable row level security;
alter table midia.fact_video enable row level security;
alter table midia.fact_engajamento enable row level security;
alter table midia.fact_conteudo_organico_metricas enable row level security;
alter table midia.fact_seguidores_snapshot enable row level security;
alter table midia.fact_analytics_ga4 enable row level security;

do $$ declare t text; begin
  for t in select unnest(array[
    'dim_conta','dim_campanha','dim_grupo','dim_anuncio','dim_conteudo_organico',
    'fact_performance','fact_resultado','fact_video','fact_engajamento',
    'fact_conteudo_organico_metricas','fact_seguidores_snapshot','fact_analytics_ga4'
  ]) loop
    execute format('drop policy if exists midia_auth on midia.%I', t);
    execute format('create policy midia_auth on midia.%I for all to authenticated using (true) with check (true)', t);
  end loop;
end $$;

-- ===== acesso Power BI (read-only, sem tocar em public) =====
grant usage on schema midia to powerbi_reader;
grant select on all tables in schema midia to powerbi_reader;
alter default privileges in schema midia grant select on tables to powerbi_reader;
