-- Controle do import incremental das vendas da planilha (só as linhas novas)
create table if not exists public.journey_sync (
  client_id text not null,
  tab text not null,
  last_row int default 2,
  last_run timestamptz,
  resultado jsonb,
  primary key (client_id, tab)
);
alter table public.journey_sync enable row level security;
drop policy if exists journey_sync_auth on public.journey_sync;
create policy journey_sync_auth on public.journey_sync for all to authenticated using (true) with check (true);
insert into public.journey_sync (client_id, tab, last_row) values ('mqu2mawxst76','SITE',330002)
  on conflict (client_id,tab) do nothing;
