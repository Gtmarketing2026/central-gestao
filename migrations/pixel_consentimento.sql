-- LGPD: o site do cliente pode ja ter aviso/banner de cookies proprio (ex: plataforma de e-commerce que ja
-- inclui, ou o cliente ja contratou algo). Se ja tem, o nosso pixel roda normal (confia no aviso generico que
-- ja existe no site - e assim que a maioria dos scripts de terceiro/analytics ja se comporta). Se nao tem
-- (padrao, cliente novo), o pixel mostra um banner leve proprio antes de gravar cookie/mandar evento.
alter table public.clients add column if not exists pixel_banner_proprio boolean not null default false;
