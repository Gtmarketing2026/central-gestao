-- Correcao de seguranca: 8 tabelas tinham a policy de RLS com "to public" em vez de "to authenticated".
-- Em Postgres, "to public" cobre TODO mundo, inclusive o papel "anon" do Supabase (chave publica embutida
-- no frontend) -- ou seja, qualquer pessoa na internet, sem login, conseguia ler/gravar essas tabelas via
-- REST API. Confirmado ao vivo: wa_conversations, wa_messages, capi_events, calendar_events,
-- andreia_automations e wa_agent_sessions vazavam dados reais de clientes (nome, telefone, mensagens);
-- wa_instances vazava o uaz_token (token da integracao com a uazapi) em texto puro.
-- Ja aplicado direto em producao em 2026-08-05; este arquivo documenta a mudanca pro historico do repo.
alter policy aa_all on public.andreia_automations to authenticated;
alter policy cal_all on public.calendar_events to authenticated;
alter policy capi_all on public.capi_events to authenticated;
alter policy rl_all on public.report_layouts to authenticated;
alter policy was_all on public.wa_agent_sessions to authenticated;
alter policy wa_conversations_all on public.wa_conversations to authenticated;
alter policy wa_instances_all on public.wa_instances to authenticated;
alter policy wa_messages_all on public.wa_messages to authenticated;
