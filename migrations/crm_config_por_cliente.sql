alter table public.clients
  add column if not exists crm_config jsonb;

-- Preserva o comportamento atual, mas cria uma cópia independente para cada cliente.
-- Depois desta migração, alterações feitas em um cliente não afetam os demais.
update public.clients c
set crm_config = jsonb_build_object(
  'stages', coalesce(a.data->'crm_stages', '[]'::jsonb),
  'fields', coalesce(a.data->'crm_fields', '[]'::jsonb),
  'min_confidence', coalesce((a.data->>'crm_min_confidence')::numeric, 70)
)
from public.account_config a
where a.id = 'main'
  and c.crm_config is null;

