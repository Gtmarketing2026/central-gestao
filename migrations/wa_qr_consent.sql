create table if not exists public.wa_qr_consents (
  id text primary key,
  instance_id text not null references public.wa_instances(id) on delete cascade,
  client_id text references public.clients(id) on delete cascade,
  authorized_by text not null,
  authorized_role text,
  authorized_email text,
  term_version text not null default '2026-08-13',
  purpose text not null,
  accepted_at timestamptz not null default now(),
  revoked_at timestamptz,
  ip_hash text,
  user_agent text,
  created_at timestamptz not null default now()
);
create index if not exists idx_wa_qr_consents_instance on public.wa_qr_consents(instance_id, accepted_at desc);
alter table public.wa_qr_consents enable row level security;
revoke all on public.wa_qr_consents from anon;
revoke all on public.wa_qr_consents from authenticated;
