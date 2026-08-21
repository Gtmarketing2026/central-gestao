-- Avaliacao de qualidade: guardar POR QUE uma conversa foi descartada.
--
-- A avaliacao pegava qualquer conversa do periodo e dava nota: papo com a propria equipe, fornecedor,
-- assunto pessoal e conversa de 3 mensagens entravam junto e puxavam a media pra baixo. Agora cada caso
-- guarda se era mesmo atendimento a cliente/lead e, quando nao era, o motivo -- em vez de sumir calado.
alter table public.crm_capa_cases
  add column if not exists relevant boolean not null default true,
  add column if not exists irrelevance_reason text;
