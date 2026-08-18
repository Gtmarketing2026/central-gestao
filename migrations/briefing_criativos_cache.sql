-- Criativos do periodo, guardados por um tempo curto.
--
-- Buscar os anuncios de um cliente grande passa de 80 segundos so no Meta (7 contas, 1.637 anuncios
-- em 79 dias). Fazer isso e a leitura da IA na MESMA chamada estourava o limite da Edge Function e a
-- tela mostrava "Edge Function returned a non-2xx status code". Agora a busca acontece numa chamada,
-- guarda aqui, e a leitura da IA vem numa segunda chamada que so le daqui — cada uma cabe no limite.
create table if not exists public.briefing_criativos_cache (
  client_id  text        not null,
  since      date        not null,
  until      date        not null,
  payload    jsonb       not null,
  gerado_em  timestamptz not null default now(),
  primary key (client_id, since, until)
);

-- RLS ligada e NENHUMA policy, de proposito: quem le e escreve e so a Edge Function, que usa a chave
-- de servico e passa por cima da RLS. Sem isso, a chave publica da API alcancaria o gasto de anuncio
-- de todos os clientes.
alter table public.briefing_criativos_cache enable row level security;
