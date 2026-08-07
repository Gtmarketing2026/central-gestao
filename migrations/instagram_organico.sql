-- Conta do Instagram Business vinculada por cliente, pra puxar posts organicos + metricas reais
-- (alimenta a Curadoria de Conteudo do Agente de Briefing Criativo e, depois, a aba Social).
alter table public.clients add column if not exists instagram_business_id text not null default '';
alter table public.clients add column if not exists instagram_username text not null default '';
