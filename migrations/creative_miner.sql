-- Minerador de Criativos: referências externas/próprias, leitura semântica e conceitos estáticos.
-- Dados acessíveis somente pelo backend protegido; o navegador usa a Edge Function autenticada.
create table if not exists public.creative_miner_items (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  source_type text not null, -- instagram_official | apify | manual
  source_url text not null,
  profile text,
  media_type text not null default 'post',
  caption text,
  media_url text,
  thumbnail_url text,
  published_at timestamptz,
  metrics jsonb not null default '{}'::jsonb,
  analysis jsonb,
  concepts jsonb not null default '[]'::jsonb,
  status text not null default 'captured',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique(client_id, source_url)
);
create index if not exists idx_creative_miner_client on public.creative_miner_items(client_id, created_at desc);
alter table public.creative_miner_items enable row level security;
revoke all on public.creative_miner_items from anon;
revoke all on public.creative_miner_items from authenticated;
comment on table public.creative_miner_items is 'Biblioteca privada de referências do Minerador; acesso exclusivo pelo backend autenticado.';
