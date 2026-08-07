-- Conta(s) do Instagram Business vinculada(s) por cliente, pra puxar posts organicos + metricas reais
-- (alimenta a Curadoria de Conteudo do Agente de Briefing Criativo e, depois, a aba Social).
-- Um cliente pode ter mais de um perfil (achado real: "Curso Fernanda Pessoa" tem 2 Paginas/perfis
-- diferentes na mesma conta de anuncios) - por isso lista, nao campo unico, no molde de meta_account_id.
alter table public.clients add column if not exists instagram_accounts jsonb not null default '[]';
-- migra o teste pontual das colunas antigas (unico registro usado, ver commit anterior) e remove elas.
update public.clients
  set instagram_accounts = jsonb_build_array(jsonb_build_object('id', instagram_business_id, 'username', instagram_username))
  where coalesce(instagram_business_id, '') <> '';
alter table public.clients drop column if exists instagram_business_id;
alter table public.clients drop column if exists instagram_username;
