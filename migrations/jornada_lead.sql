-- ===== JORNADA DO LEAD: grafo de identidade + fluxo de toques =====
create table if not exists public.lead_people (
  id text primary key,
  client_id text,
  name text, phone text, email text,
  first_seen timestamptz, last_seen timestamptz,
  stage text, value numeric default 0,
  created_at timestamptz default now(), updated_at timestamptz default now()
);
create index if not exists lead_people_client_idx on public.lead_people(client_id);

create table if not exists public.lead_identities (
  id text primary key,
  client_id text, person_id text not null,
  kind text not null,
  value text not null,
  first_seen timestamptz default now(), last_seen timestamptz default now()
);
create unique index if not exists lead_identities_uq on public.lead_identities(client_id, kind, value);
create index if not exists lead_identities_person_idx on public.lead_identities(person_id);

create table if not exists public.lead_touchpoints (
  id text primary key,
  client_id text, person_id text,
  ts timestamptz not null,
  kind text not null,
  channel text,
  source text, medium text, campaign text, adset text, ad text, term text, content text,
  page text, referrer text,
  label text, value numeric default 0,
  ref_table text, ref_id text,
  meta jsonb,
  created_at timestamptz default now()
);
create unique index if not exists lead_touchpoints_ref_uq on public.lead_touchpoints(ref_table, ref_id);
create index if not exists lead_touchpoints_person_idx on public.lead_touchpoints(person_id, ts);
create index if not exists lead_touchpoints_client_ts_idx on public.lead_touchpoints(client_id, ts);

alter table public.lead_people enable row level security;
alter table public.lead_identities enable row level security;
alter table public.lead_touchpoints enable row level security;
drop policy if exists lead_people_auth on public.lead_people;
drop policy if exists lead_identities_auth on public.lead_identities;
drop policy if exists lead_touchpoints_auth on public.lead_touchpoints;
create policy lead_people_auth on public.lead_people for all to authenticated using (true) with check (true);
create policy lead_identities_auth on public.lead_identities for all to authenticated using (true) with check (true);
create policy lead_touchpoints_auth on public.lead_touchpoints for all to authenticated using (true) with check (true);
