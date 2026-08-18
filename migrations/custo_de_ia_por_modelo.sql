-- A tela de Custos do Sistema calculava a IA usando SÓ a linha "gemini", com preço zerado, e não
-- existia linha da OpenAI: o custo de IA aparecia como R$ 0 enquanto a fatura corria por fora.
-- Preço tem que ser POR MODELO: gpt-4o custa ~16x o gpt-4o-mini, então uma média por provedor mente.
create table if not exists public.ai_model_prices (
  model text primary key,
  provider text not null,
  input_usd_per_million numeric not null default 0,
  output_usd_per_million numeric not null default 0,
  updated_at timestamptz not null default now()
);
alter table public.ai_model_prices enable row level security;
drop policy if exists ai_model_prices_admin on public.ai_model_prices;
create policy ai_model_prices_admin on public.ai_model_prices
  for all to authenticated
  using (public.is_system_cost_admin()) with check (public.is_system_cost_admin());

insert into public.ai_model_prices (model, provider, input_usd_per_million, output_usd_per_million) values
  ('gpt-4o',                'openai', 2.50, 10.00),
  ('gpt-4o-2024-08-06',     'openai', 2.50, 10.00),
  ('gpt-4o-mini',           'openai', 0.15,  0.60),
  ('gpt-4o-mini-2024-07-18','openai', 0.15,  0.60),
  ('gemini-flash-latest',   'gemini', 0.075, 0.30),
  ('gemini-3.5-flash',      'gemini', 0.075, 0.30),
  ('gemini-3.5-flash-lite', 'gemini', 0.0375,0.15),
  -- Whisper e cobrado por MINUTO, nao por token. A telemetria (dynamic-responder.ts, _regUsoIa)
  -- grava a DURACAO do audio em input_units, entao o preco dele fica "USD por 1 milhao de
  -- SEGUNDOS" (US$ 100 = US$ 0,006/min) — assim a mesma formula de custo serve sem caso especial.
  ('whisper-1',             'openai', 100,   0)
on conflict (model) do nothing;

-- consumo por modelo, pra tela conseguir multiplicar cada um pelo seu preço
create or replace function public.system_ai_usage_by_model(p_from date, p_to date)
returns table(service_key text, model text, requests bigint, input_tokens bigint, output_tokens bigint)
language sql stable security definer set search_path to 'public' as $$
  select e.service_key,
         coalesce(nullif(e.meta->>'model',''),'(desconhecido)') as model,
         count(*)::bigint,
         coalesce(sum(e.input_units),0)::bigint,
         coalesce(sum(e.output_units),0)::bigint
  from system_usage_events e
  where public.is_system_cost_admin()
    and e.service_key in ('gemini','openai')
    and e.occurred_at >= p_from::timestamptz
    and e.occurred_at < (p_to+1)::timestamptz
  group by 1,2
  order by 4 desc;
$$;
