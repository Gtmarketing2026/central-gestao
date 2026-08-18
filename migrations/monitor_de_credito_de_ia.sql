-- Nem a OpenAI nem o Google expõem SALDO pela chave da API: a OpenAI só dá custo (e com chave de
-- admin), o Google exige credencial de faturamento do Cloud. Então o monitor não lê saldo — ele
-- acompanha o GASTO contra um teto que a gestora define, e o consumo do dia contra o limite do
-- plano gratuito do Gemini. É o que dá pra fazer sem pedir credencial nova, e resolve o problema
-- real: avisar ANTES de parar.
create table if not exists public.ai_credit_watch (
  provider text primary key,
  teto_usd_mes numeric not null default 0,      -- 0 = sem teto definido
  limite_dia_requests int not null default 0,   -- free tier do Gemini (0 = sem limite)
  avisar_em int[] not null default '{50,80,100}',
  ultimo_aviso jsonb not null default '{}'::jsonb,
  atualizado_em timestamptz not null default now()
);
alter table public.ai_credit_watch enable row level security;
drop policy if exists ai_credit_watch_admin on public.ai_credit_watch;
create policy ai_credit_watch_admin on public.ai_credit_watch
  for all to authenticated
  using (public.is_system_cost_admin()) with check (public.is_system_cost_admin());

insert into public.ai_credit_watch (provider, teto_usd_mes, limite_dia_requests) values
  ('openai', 0, 0),
  ('gemini', 0, 500)
on conflict (provider) do nothing;

-- gasto do mês por provedor, já multiplicado pelo preço de cada modelo
create or replace function public.ai_gasto_mes()
returns table(provider text, usd numeric, requests bigint, requests_hoje bigint)
language sql stable security definer set search_path to 'public' as $$
  select e.service_key,
         round(coalesce(sum(e.input_units * p.input_usd_per_million / 1e6
                          + e.output_units * p.output_usd_per_million / 1e6), 0)::numeric, 4),
         count(*)::bigint,
         count(*) filter (where e.occurred_at >= date_trunc('day', now()))::bigint
  from system_usage_events e
  left join ai_model_prices p on p.model = e.meta->>'model'
  where e.service_key in ('gemini','openai')
    and e.occurred_at >= date_trunc('month', now())
  group by e.service_key;
$$;

-- checagem periodica: gasto do mes x teto, e chamadas de hoje x limite diario do plano gratuito.
-- avisa uma vez por faixa (50/80/100%) por mes, no sino e no grupo do WhatsApp da AndreIA.
-- chamada pela rota /ia-credito/tick (tracking.ts), a cada 4h via pg_cron.
-- cron.schedule() com o mesmo jobname substitui o agendamento anterior (upsert), então rodar de novo é seguro.
select cron.schedule(
  'ia-credito-tick',
  '0 */4 * * *',
  $cron$select net.http_get(
    url:='https://mocrfqmdjwvyhqvdpimm.supabase.co/functions/v1/tracking/ia-credito/tick',
    headers:=jsonb_build_object('Content-Type','application/json','x-cron-secret',
      (select decrypted_secret from vault.decrypted_secrets where name='internal_cron_secret' limit 1))
  )$cron$
);
