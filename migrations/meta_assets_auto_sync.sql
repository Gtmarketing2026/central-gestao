alter table public.clients add column if not exists instagram_accounts_excluded jsonb not null default '[]'::jsonb;
comment on column public.clients.instagram_accounts_excluded is 'IDs de Instagram removidos manualmente; a sincronização automática não os readiciona.';
