create table if not exists public.secure_credentials (
  id text primary key,
  secret_cipher text not null,
  updated_at timestamptz not null default now()
);

alter table public.secure_credentials enable row level security;
revoke all on table public.secure_credentials from anon, authenticated;
comment on table public.secure_credentials is 'Credenciais criptografadas, acessíveis somente pelas Edge Functions com service role';

create table if not exists public.secure_credential_admins (
  email text primary key,
  created_at timestamptz not null default now()
);
alter table public.secure_credential_admins enable row level security;
revoke all on table public.secure_credential_admins from anon, authenticated;
insert into public.secure_credential_admins(email) values ('contato@cursosonlinevrrj.com.br') on conflict do nothing;
