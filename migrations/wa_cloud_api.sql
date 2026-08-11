alter table public.wa_instances
  add column if not exists waba_id text,
  add column if not exists phone_number_id text,
  add column if not exists meta_app_id text,
  add column if not exists verified_name text,
  add column if not exists quality_rating text;

create unique index if not exists wa_instances_phone_number_id_uniq
  on public.wa_instances(phone_number_id)
  where phone_number_id is not null and phone_number_id <> '';

comment on column public.wa_instances.waba_id is 'ID público da conta do WhatsApp Business';
comment on column public.wa_instances.phone_number_id is 'ID público do número na Cloud API';
comment on column public.wa_instances.meta_app_id is 'ID público do aplicativo Meta';
comment on column public.wa_instances.verified_name is 'Nome verificado consultado automaticamente na Meta';
comment on column public.wa_instances.quality_rating is 'Qualidade do número consultada automaticamente na Meta';
