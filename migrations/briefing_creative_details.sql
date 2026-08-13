-- Mantém thumbnails, links e KPIs reais usados em cada análise para reabrir o histórico completo.
alter table public.briefing_analise
  add column if not exists criativos_json jsonb not null default '{}'::jsonb;
