alter table public.wa_conversations
  add column if not exists import_batch_id text,
  add column if not exists import_source text;

create table if not exists public.crm_import_batches (
  id text primary key,
  client_id text not null references public.clients(id) on delete cascade,
  instance_id text references public.wa_instances(id) on delete set null,
  file_name text not null,
  row_count integer not null default 0,
  added_count integer not null default 0,
  duplicate_count integer not null default 0,
  invalid_count integer not null default 0,
  status text not null default 'active',
  protected_count integer not null default 0,
  deleted_count integer not null default 0,
  created_at timestamptz not null default now(),
  deleted_at timestamptz
);

create index if not exists idx_wa_conversations_import_batch on public.wa_conversations(import_batch_id);
create index if not exists idx_crm_import_batches_client on public.crm_import_batches(client_id, created_at desc);

alter table public.crm_import_batches enable row level security;
drop policy if exists crm_import_batches_auth on public.crm_import_batches;
-- Sem política para anon/authenticated: acesso somente pelo service role da Edge Function.
