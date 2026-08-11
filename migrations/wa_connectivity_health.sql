alter table public.wa_instances
  add column if not exists health_fail_count integer not null default 0,
  add column if not exists health_last_alert_at timestamptz,
  add column if not exists health_last_ok_at timestamptz,
  add column if not exists health_last_recovery_at timestamptz;

comment on column public.wa_instances.health_fail_count is 'Falhas consecutivas do monitor antes de confirmar uma queda';
comment on column public.wa_instances.health_last_alert_at is 'Último alerta confirmado de desconexão';
comment on column public.wa_instances.health_last_ok_at is 'Última verificação saudável diretamente no provedor';
comment on column public.wa_instances.health_last_recovery_at is 'Última recuperação já comunicada e sincronizada';
