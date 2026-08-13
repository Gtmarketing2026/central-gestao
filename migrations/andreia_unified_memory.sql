-- Memória seletiva da AndréIA: somente decisões/preferências explícitas, sem transcrições completas.
create table if not exists public.andreia_memory (
  id text primary key,
  client_id text references public.clients(id) on delete cascade,
  scope text not null default 'client',
  kind text not null default 'decision',
  content text not null,
  source text not null,
  created_at timestamptz not null default now(),
  active boolean not null default true
);
create index if not exists idx_andreia_memory_scope on public.andreia_memory(client_id, active, created_at desc);
alter table public.andreia_memory enable row level security;
revoke all on public.andreia_memory from anon;
revoke all on public.andreia_memory from authenticated;
comment on table public.andreia_memory is 'Memórias explícitas e minimizadas da AndréIA; acesso exclusivo pelo backend protegido.';
