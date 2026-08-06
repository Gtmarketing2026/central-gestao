-- Agente de Briefing Criativo: analisa os criativos que rodaram (Meta/Google) por etapa de funil,
-- opcionalmente cura conteudo organico pra recorte, e monta fichas de producao. Ver spec em
-- spec-agente-briefing-claude-code.md (nao versionado no repo, guardado pela usuaria).
create table if not exists public.briefing (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  criado_por text references public.team(id) on delete set null,
  periodo_inicio date not null,
  periodo_fim date not null,
  objetivo text not null default '',
  publico text not null default '',
  angulo text not null default '',
  promessa text not null default '',
  funil text not null default '', -- Topo | Meio | Fundo
  formatos_json jsonb not null default '[]',
  canais_json jsonb not null default '[]',
  variacoes int not null default 4,
  referencia text not null default '',
  prazo date,
  curadoria_ativa boolean not null default false,
  status text not null default 'rascunho', -- rascunho | gerando | pronto | aprovado | entregue
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now()
);

create table if not exists public.briefing_analise (
  id text primary key,
  briefing_id text not null references public.briefing(id) on delete cascade,
  leitura text not null default '',
  funis_json jsonb not null default '[]',
  padroes_json jsonb not null default '[]',
  criativos_analisados int not null default 0,
  criativos_em_leitura int not null default 0,
  funil_inferido_pct numeric not null default 0, -- % dos criativos que precisou de inferencia pra achar o funil
  gerado_em timestamptz not null default now()
);

create table if not exists public.briefing_curadoria (
  id text primary key,
  briefing_id text not null references public.briefing(id) on delete cascade,
  leitura text not null default '',
  candidatos_json jsonb not null default '[]',
  gerado_em timestamptz not null default now()
);

create table if not exists public.briefing_ficha (
  id text primary key,
  briefing_id text not null references public.briefing(id) on delete cascade,
  codigo text not null default '',
  titulo text not null default '',
  prioridade text not null default 'P2',
  rota text not null default 'nova', -- nova | recorte
  funil text not null default '',
  canal text not null default '',
  formato text not null default '',
  objetivo text not null default '',
  referencia text not null default '',
  publico text not null default '',
  angulo text not null default '',
  promessa text not null default '',
  roteiro_json jsonb not null default '{}',
  copy_json jsonb not null default '{}',
  especificacoes text not null default '',
  obrigatorio_json jsonb not null default '[]',
  proibido_json jsonb not null default '[]',
  prazo date,
  ordem int not null default 0,
  status text not null default 'fila' -- fila | producao | entregue
);

create table if not exists public.briefing_mensagem (
  id text primary key,
  briefing_id text not null references public.briefing(id) on delete cascade,
  papel text not null default 'gestor', -- gestor | agente
  texto text not null default '',
  criado_em timestamptz not null default now()
);

create index if not exists idx_briefing_client on public.briefing(client_id);
create index if not exists idx_briefing_analise_briefing on public.briefing_analise(briefing_id);
create index if not exists idx_briefing_curadoria_briefing on public.briefing_curadoria(briefing_id);
create index if not exists idx_briefing_ficha_briefing on public.briefing_ficha(briefing_id);
create index if not exists idx_briefing_mensagem_briefing on public.briefing_mensagem(briefing_id);

alter table public.briefing enable row level security;
alter table public.briefing_analise enable row level security;
alter table public.briefing_curadoria enable row level security;
alter table public.briefing_ficha enable row level security;
alter table public.briefing_mensagem enable row level security;

drop policy if exists briefing_auth on public.briefing;
create policy briefing_auth on public.briefing for all to authenticated using (true) with check (true);
drop policy if exists briefing_analise_auth on public.briefing_analise;
create policy briefing_analise_auth on public.briefing_analise for all to authenticated using (true) with check (true);
drop policy if exists briefing_curadoria_auth on public.briefing_curadoria;
create policy briefing_curadoria_auth on public.briefing_curadoria for all to authenticated using (true) with check (true);
drop policy if exists briefing_ficha_auth on public.briefing_ficha;
create policy briefing_ficha_auth on public.briefing_ficha for all to authenticated using (true) with check (true);
drop policy if exists briefing_mensagem_auth on public.briefing_mensagem;
create policy briefing_mensagem_auth on public.briefing_mensagem for all to authenticated using (true) with check (true);
