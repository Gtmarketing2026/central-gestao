-- Mais de um numero de WhatsApp por empresa.
--
-- Ate aqui a conversa era identificada por (cliente + telefone do lead), com UNIQUE nesse par: dois
-- numeros da MESMA empresa colidiam -- o mesmo lead falando nos dois virava uma conversa so, e a
-- resposta saia por qualquer um dos numeros. Agora a conversa (e a mensagem) sabem por qual numero
-- entraram.
alter table public.wa_conversations add column if not exists instance_id text;
alter table public.wa_messages      add column if not exists instance_id text;

-- Preenchimento do que ja existe: hoje cada cliente tem exatamente UMA instancia, entao nao ha duvida.
-- A agencia (client_id null) tem duas e fica sem preenchimento de proposito -- nao da pra adivinhar,
-- e a chave abaixo trata null como valor, mantendo essas conversas funcionando como antes.
update public.wa_conversations c
   set instance_id = u.id
  from (select client_id, min(id) as id from public.wa_instances
         where client_id is not null group by client_id having count(*) = 1) u
 where c.client_id = u.client_id and c.instance_id is null;

update public.wa_messages m
   set instance_id = c.instance_id
  from public.wa_conversations c
 where m.conversation_id = c.id and m.instance_id is null and c.instance_id is not null;

-- A unicidade passa a incluir o numero. coalesce(...,'') porque no Postgres NULL nao colide com NULL:
-- sem isso, conversa sem instancia (agencia, importacao antiga) poderia duplicar.
alter table public.wa_conversations drop constraint if exists wa_conversations_client_id_chat_id_key;
create unique index if not exists wa_conversations_client_chat_inst_key
  on public.wa_conversations (client_id, chat_id, coalesce(instance_id, ''));

create index if not exists wa_conv_inst on public.wa_conversations (instance_id, last_at desc);
create index if not exists wa_msg_inst  on public.wa_messages (instance_id);
