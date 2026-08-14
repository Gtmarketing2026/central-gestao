alter table public.clients add column if not exists youtube_config jsonb not null default '{}'::jsonb;
comment on column public.clients.youtube_config is 'Metadados públicos do canal YouTube; refresh token fica criptografado em secure_credentials.';
