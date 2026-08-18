-- Minerador de Criativos: motivo de seleção (disclaimer no card) e funil desejado no momento da captura.
alter table public.creative_miner_items
  add column if not exists selection_reason text,
  add column if not exists funnel_target text;
