-- Conexão por cliente do TikTok Ads e Pinterest Ads (OAuth, mesmo molde do RD Station/Nuvemshop).
alter table public.clients add column if not exists tiktok_config jsonb default null;
alter table public.clients add column if not exists pinterest_config jsonb default null;
