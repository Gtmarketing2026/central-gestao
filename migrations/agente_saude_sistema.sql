-- Agente de saude do sistema (roda 05:30) — fotografia unica de cron, erros internos,
-- seguranca, frescor dos canais e contas de midia paradas.
-- Lido por systemHealthTick() no dynamic-responder; resultado vai pra account_config.data.health_report.
--
-- 14/08/2026: "contas de midia paradas" so considera conta que REALMENTE investiu nos ultimos 30 dias.
-- Antes, conta antiga sem nenhum investimento ficava listada pra sempre (o dado nunca mais fica "novo").
create or replace function public.system_health_snapshot()
returns jsonb
language sql
security definer
set search_path to 'public', 'pg_catalog'
as $function$
select jsonb_build_object(
  'cron_falhas_24h', (
    select coalesce(jsonb_agg(jsonb_build_object('job', j.jobname, 'falhas', d.n, 'ultima', d.ultima)), '[]'::jsonb)
    from (select jobid, count(*) n, max(start_time) ultima from cron.job_run_details
          where start_time > now() - interval '24 hours' and status <> 'succeeded' group by jobid) d
    join cron.job j using (jobid)
  ),
  'http_erros_24h', (
    select coalesce(jsonb_agg(jsonb_build_object('status', coalesce(status_code, 0), 'qtd', n) order by n desc), '[]'::jsonb)
    from (select status_code, count(*) n from net._http_response
          where created > now() - interval '24 hours' and (status_code is null or status_code >= 400)
          group by status_code) e
  ),
  'http_total_24h', (select count(*) from net._http_response where created > now() - interval '24 hours'),
  'seguranca_eventos_24h', (
    select coalesce(jsonb_agg(jsonb_build_object('tipo', kind, 'qtd', n) order by n desc), '[]'::jsonb)
    from (select kind, count(*) n from public.security_events
          where created_at > now() - interval '24 hours' group by kind) s
  ),
  'sync_canais', (
    select coalesce(jsonb_agg(jsonb_build_object('canal', channel, 'ultima_data', mx)), '[]'::jsonb)
    from (select channel, max(date) mx from public.channel_metrics_daily group by channel) c
  ),
  'contas_midia_paradas', (
    select coalesce(jsonb_agg(jsonb_build_object(
             'cliente', cl.name, 'plataforma', co.plataforma_id,
             'ultima_data', t.mx, 'gasto_30d', round(t.gasto_30d, 2)) order by t.mx desc), '[]'::jsonb)
    from (
      select co2.id, max(fp.data) mx,
             sum(fp.investimento) filter (where fp.data > current_date - 30) gasto_30d
      from midia.dim_conta co2
      join midia.dim_campanha dc on dc.conta_id = co2.id
      join midia.dim_grupo dg on dg.campanha_id = dc.id
      join midia.dim_anuncio da on da.grupo_id = dg.id
      join midia.fact_performance fp on fp.anuncio_id = da.id
      where co2.ativa and co2.plataforma_id in ('meta','google')
      group by co2.id
      having max(fp.data) < current_date - 3
         -- so avisa de conta que estava rodando de verdade: sem investimento nos ultimos
         -- 30 dias a conta esta encerrada/parada de proposito, nao e problema de saude.
         and coalesce(sum(fp.investimento) filter (where fp.data > current_date - 30), 0) > 0
    ) t
    join midia.dim_conta co on co.id = t.id
    join public.clients cl on cl.id = co.client_id
  )
);
$function$;
