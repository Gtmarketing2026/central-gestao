-- Segmentacao livre por produto/programa do cliente (ex.: "ENEM" x "Concurso" na Curso Fernanda Pessoa).
-- Texto livre, nao catalogo fixo - filtra os criativos por nome de campanha/conjunto que contenha o termo.
alter table public.briefing add column if not exists produto text not null default '';
