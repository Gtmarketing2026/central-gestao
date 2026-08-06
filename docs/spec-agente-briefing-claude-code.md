# Agente de Briefing Criativo — Especificação de Implementação
### Sistema GT Marketing · Documento para execução via Claude Code

---

## Como usar

Coloque este arquivo na raiz do repositório do GT Marketing e aponte o Claude Code para ele. O documento descreve **o que construir e por quê**. A stack deve ser descoberta no repositório, nunca presumida.

Há um protótipo funcional de referência (`agente-briefing-app.html`) com a interface, o fluxo e os prompts já validados. Use como referência de comportamento, não como código a portar — ele foi feito para rodar isolado no navegador.

---

## 1. Antes de escrever qualquer código

Execute nesta ordem e **apresente o plano antes de implementar**:

1. Mapeie a stack: linguagem, framework, ORM, banco, filas, autenticação, padrão de migrations e de testes
2. Localize a ingestão de performance de Meta e Google que já existe. Responda com precisão:
   - Em que tabelas os dados de anúncio ficam
   - Qual a granularidade disponível (anúncio, conjunto, campanha)
   - Quais métricas existem por anúncio: investimento, impressões, alcance, frequência, cliques, CTR, CPM, CPC, resultados, custo por resultado, e métricas de vídeo se houver
   - Com que frequência atualiza
   - Qual janela de atribuição de conversão está em uso
3. Descubra **como identificar o funil de cada anúncio**. Abra 30 a 40 nomes de anúncio e de campanha reais, de clientes diferentes. Verifique se existe padrão que permita derivar topo, meio ou fundo. Reporte o que encontrou antes de seguir — este ponto define a qualidade de toda a análise
4. Verifique se existe cadastro de cliente, usuário e permissão para reaproveitar
5. Confirme como o sistema já chama LLMs, se chama. Reaproveite o cliente existente

**Não crie infraestrutura paralela à que já existe.** Se há ingestão funcionando, este agente consome dela.

---

## 2. O que é

Um agente que recebe os campos de briefing preenchidos pelo gestor, lê sozinho os criativos que rodaram no período, analisa a performance por funil, opcionalmente procura material orgânico aproveitável, e entrega o briefing pronto para o time produzir.

Substitui o preenchimento manual do briefing padrão da agência, mantendo os mesmos oito campos do escopo comercial.

### Fluxo

```
Gestor preenche os campos
        ▼
[1] ANÁLISE DOS CRIATIVOS          ← obrigatória, dados vêm da ingestão
    por funil, melhores e piores,
    pontos positivos e negativos
        ▼
[2] CURADORIA DE CONTEÚDO          ← opcional, acionada por toggle
    material de Instagram e YouTube
    que pode virar recorte
        ▼
[3] BRIEFING
    fichas de produção, marcando
    quais são recorte e quais são nova
        ▼
[4] CHAT DE AJUSTES                ← itera sem refazer do zero
```

---

## 3. Dependências de dados

### 3.1 Criativos do período — obrigatório e automático

O gestor **não cola** dados. O agente busca da ingestão existente, filtrando por cliente e período.

Campos necessários por anúncio: identificador, nome, campanha, conjunto, funil, formato, investimento, impressões, alcance, frequência, cliques, CTR, CPM, resultados, custo por resultado. Para vídeo, quando disponível: reproduções de 3 segundos e retenção em 25, 50, 75 e 100 por cento.

**Métrica ausente nunca é estimada.** Marque como não disponível e siga sem ela.

**Elegibilidade:** anúncio com investimento zero ou volume irrelevante no período não entra na análise. Piso sugerido, ajustável:

| Funil | Piso |
|---|---|
| Topo | ≥ 2.000 impressões |
| Meio | ≥ 1.500 impressões |
| Fundo | investimento ≥ 1,5× custo por resultado alvo |

Anúncios abaixo do piso aparecem como "em leitura" e não são classificados nem criticados.

### 3.2 Funil por criativo

Ordem de resolução:

1. Campo próprio, se existir no sistema
2. Parsing da nomenclatura do anúncio ou da campanha, conforme o padrão descoberto no passo 1.3
3. Inferência pelo LLM a partir do nome e da campanha, marcada com `(inferido)` na saída

Se mais de 30% dos anúncios caírem na inferência, o relatório exibe aviso de confiabilidade reduzida.

### 3.3 Conteúdo orgânico — opcional

Quando o toggle de curadoria estiver ligado. Fontes: Instagram Graph API para os perfis dos clientes conectados ao Business Manager da agência, e YouTube Data API para os canais dos clientes.

Se a integração ainda não existir, implemente a etapa aceitando entrada manual e deixe o ponto de integração isolado e documentado.

---

## 4. Interface

### Painel de entrada

| Campo | Tipo | Obrigatório |
|---|---|---|
| Cliente | seleção | Sim |
| Período analisado | intervalo de datas | Sim |
| Objetivo | texto | Sim |
| Público | texto | Não |
| Ângulo | texto ou seleção do catálogo | Sim |
| Promessa | texto | Não |
| Funil | seleção única: topo, meio, fundo | Sim |
| Formato | múltipla: estático, motion, vídeo, UGC, carrossel | Não |
| Canal | múltipla: Meta, YouTube, TikTok, Display, Search, Pinterest | Não |
| Variações | número | Sim |
| Referência | texto | Não |
| Prazo | data | Não |
| Curadoria de conteúdo | toggle, desligado por padrão | — |

Campo vazio não bloqueia além dos obrigatórios — o agente preenche o que faltar a partir da análise e sinaliza o que inferiu.

### Painel de saída

Três seções numeradas: análise, curadoria quando ativa, demanda de produção. Ao final, o chat de ajustes.

Cada ficha de produção é um bloco expansível com etiquetas de prioridade, rota, funil, canal e formato visíveis fechado.

---

## 5. Etapa 1 — Análise dos criativos

Obrigatória. Roda antes de qualquer geração de briefing.

### Regras

- **Comparação apenas dentro do funil.** Custo por resultado de topo e de fundo não são grandezas equivalentes
- Máximo 2 melhores e 2 piores por funil
- **Todo criativo listado recebe pontos positivos e negativos**, inclusive os melhores e os piores. O melhor tem algo a corrigir; o pior quase sempre tem algo aproveitável
- Os pontos são escritos como instrução para quem produz, não como descrição de métrica. "Abre com rosto em close antes de qualquer texto" e não "teve CTR de 1,2%"
- Padrão apoiado em menos de 3 peças é hipótese e deve ser rotulado

### Prompt

```
Voce e analista de criativos de trafego pago. Analise os criativos que rodaram e devolva a
leitura para o time de producao, que precisa entender o que funcionou antes de criar peca nova.

Cliente: {cliente}
Periodo: {periodo}
Objetivo da campanha: {objetivo}

Criativos que rodaram:
{dados}

Organize a analise POR FUNIL: topo, meio e fundo. Inclua apenas os funis presentes nos dados.
Se o funil nao estiver explicito, infira pela campanha ou pelo nome do anuncio e acrescente
(inferido) ao lado do codigo.

Compare cada criativo apenas dentro do proprio funil. Custo por resultado de topo e de fundo
nao sao grandezas equivalentes.

Para cada criativo, traga pontos positivos e pontos negativos, mesmo nos melhores e nos piores:
o melhor criativo tem algo a corrigir e o pior quase sempre tem algo aproveitavel. Escreva os
pontos como instrucao util para quem vai produzir, nao como descricao de metrica.

{REGRAS_DE_LINGUAGEM}

Limites: no maximo 2 melhores e 2 piores por funil, no maximo 2 pontos positivos e 2 negativos
por criativo, frases curtas de uma linha. No maximo 3 padroes no total. Padrao apoiado em menos
de 3 pecas deve ter hipotese true.

Responda APENAS com JSON valido, sem markdown, sem preambulo:
{"leitura":"<um paragrafo curto sobre o periodo inteiro>",
 "funis":[{"funil":"Topo","leitura":"<uma frase sobre este funil>","total_pecas":0,
   "melhores":[{"codigo":"","metricas":"","positivos":[""],"negativos":[""]}],
   "piores":[{"codigo":"","metricas":"","positivos":[""],"negativos":[""]}]}],
 "padroes":[{"afirmacao":"","evidencia":"","pecas":0,"hipotese":false}]}
```

---

## 6. Etapa 2 — Curadoria de conteúdo

Opcional, acionada pelo toggle. Procura material orgânico que possa virar recorte em vez de produção nova.

### Critérios de pontuação

Ordem de peso, e a ordem importa mais que os números:

1. **Percentual de alcance de não-seguidores** — o melhor preditor disponível de performance com público frio
2. **Compartilhamentos sobre alcance** — ressonância
3. **Salvamentos sobre alcance** — utilidade percebida, bom sinal de meio
4. **Retenção ou tempo médio de visualização**
5. **Autocontenção** — o trecho se entende sem o contexto anterior

Normalize contra a mediana do próprio perfil, nunca contra benchmark externo. Corte em 60.

### Bloqueios de uso

Verificar antes de sugerir. Peça reprovada na subida depois de editada é o pior desperdício possível.

| Bloqueio | Consequência |
|---|---|
| Áudio de biblioteca do Instagram | **Impeditivo.** Licença orgânica não cobre uso comercial |
| Marca d'água de outra plataforma | Impeditivo na prática, entrega penalizada |
| Rosto de terceiro sem cessão de imagem | Conferência humana obrigatória |
| Resolução abaixo de 1080p vertical | Descartar |
| Oferta ou preço desatualizado | Não serve como peça permanente |

### Prompt

```
Voce e curador de conteudo para trafego pago. Avalie o material organico abaixo e diga o que
pode virar criativo pago por recorte, em vez de producao nova.

Cliente: {cliente}
Angulo desejado: {angulo}
Funil desejado: {funil}
Canais: {canais}

Conteudos disponiveis:
{conteudos}

Criterios de pontuacao de 0 a 100, nesta ordem de peso:
1. Percentual de alcance de nao-seguidores. E o melhor preditor de performance com publico frio.
2. Compartilhamentos sobre alcance. Indica ressonancia.
3. Salvamentos sobre alcance. Indica utilidade percebida, bom sinal de meio.
4. Retencao ou tempo medio de visualizacao.
5. Autocontencao: o trecho se entende sem o contexto anterior.

Bloqueios que impedem o uso e devem ser reportados:
- Audio de biblioteca do Instagram. A licenca organica nao cobre uso comercial.
- Marca dagua de outra plataforma.
- Rosto de terceiro sem cessao de imagem, que exige conferencia humana.
- Resolucao abaixo de 1080p vertical.
- Oferta ou preco desatualizado na peca.

{REGRAS_DE_LINGUAGEM}

Traga no maximo 5 candidatos, do maior para o menor score. Item com score abaixo de 60 nao entra.
Se o material tiver timecode ou duracao, sugira janela de corte; se nao tiver, descreva o trecho.

Responda APENAS com JSON valido, sem markdown, sem preambulo:
{"leitura":"<uma frase sobre o inventario disponivel>",
 "candidatos":[{"titulo":"","origem":"","score":0,"funil":"","corte":"","aproveitar":"",
   "complemento":"","bloqueios":""}]}
```

---

## 7. Etapa 3 — Briefing

Consome a análise e, quando houver, a curadoria.

### Regras

- Ficha sem roteiro e sem copy não é briefing, é pedido. Preencher os dois sempre
- Quando houver candidato de curadoria adequado, a ficha sai como **rota recorte**: no lugar do roteiro vai a direção de corte, o reenquadramento para a proporção de destino e o que precisa ser gravado a mais
- Recorte custa fração de produção nova. Prefira recorte quando houver candidato adequado
- Headline com no máximo 40 caracteres
- Especificações incluem proporção, duração e requisitos do canal: safe zone, legenda, som

### Prompt

```
Voce e o agente de briefing criativo de uma operacao de trafego pago. Gere {n} fichas de
producao completas.

Campos do briefing:
Cliente: {cliente}
Objetivo: {objetivo}
Publico: {publico}
Angulo: {angulo}
Promessa: {promessa}
Funil: {funil}
Formatos: {formatos}
Canais: {canais}
Referencia: {referencia}
Prazo: {prazo}

Analise dos criativos que rodaram:
{analise_json}
Use os padroes identificados para orientar os roteiros. Referencie os codigos vencedores quando
fizer sentido.

[quando houver curadoria]
Material organico aprovado na curadoria:
{candidatos_json}
Sempre que um candidato cobrir a necessidade da ficha, gere a ficha como rota recorte em vez de
producao nova: preencha rota com recorte, informe o material de origem em referencia, e no lugar
do roteiro descreva a direcao de corte, o reenquadramento para a proporcao de destino e o que
precisa ser gravado a mais. Recorte custa fracao de producao nova, entao prefira recorte quando
houver candidato adequado.

{REGRAS_DE_LINGUAGEM}

Cada ficha precisa de roteiro e copy preenchidos. Headline com no maximo 40 caracteres. Nas
especificacoes, inclua proporcao, duracao e requisitos do canal (safe zone, legenda, som).

Responda APENAS com JSON valido, sem markdown, sem preambulo:
{"fichas":[{"codigo":"AD01","titulo":"","prioridade":"P1","rota":"nova","funil":"","canal":"",
 "formato":"","objetivo":"","referencia":"","publico":"","angulo":"","promessa":"",
 "roteiro":{"hook":"","desenvolvimento":"","prova":"","cta":""},
 "copy":{"headline":"","texto":"","cta":""},
 "especificacoes":"","obrigatorio":[""],"proibido":[""],"prazo":""}]}
```

**Dimensionamento de tokens:** briefing com muitas fichas completas é longo. Dimensione o limite de saída com folga ou gere em lotes de 3 a 4 fichas, concatenando. Truncamento silencioso é o modo de falha mais provável desta etapa — valide que o JSON fechou antes de renderizar.

---

## 8. Etapa 4 — Chat de ajustes

Itera sobre o briefing já gerado sem refazer as etapas anteriores. Mantém histórico da conversa.

```
Voce e o agente de briefing criativo. Este e o briefing atual:
{fichas_json}

Campos originais: {form_json}
Analise: {analise_json}
Conteudo disponivel para recorte: {candidatos_json}

Ajuste solicitado pelo gestor: {mensagem}

{REGRAS_DE_LINGUAGEM}

Aplique o ajuste e devolva o briefing completo atualizado. Mantenha o que nao foi pedido para
mudar.

Responda APENAS com JSON valido no mesmo formato, sem markdown:
{"fichas":[...],"resposta":"<uma frase sobre o que voce mudou>"}
```

---

## 9. Regras de linguagem

Bloco `{REGRAS_DE_LINGUAGEM}` referenciado em todos os prompts. Invioláveis — valem para o cliente e para o time.

```
Regras de linguagem obrigatorias: nunca use os termos qualificado, lead quente, intencao de
compra ou engajado. Nunca afirme causalidade a partir de correlacao. Nao use adjetivo sem numero
que o sustente. Escreva em portugues do Brasil, tom tecnico e direto, sem linguagem de venda.
```

Motivo: termos de qualidade de lead exigem tracking que confirme, e afirmá-los sem isso cria expectativa que o relatório não sustenta. É regra estabelecida da operação.

---

## 10. Modelo de dados

Adapte às convenções do repositório.

```sql
briefing (
  id, cliente_id, criado_por,
  periodo_inicio, periodo_fim,
  objetivo, publico, angulo, promessa,
  funil, formatos_json, canais_json,
  variacoes, referencia, prazo,
  curadoria_ativa,
  status,                    -- rascunho | gerando | pronto | aprovado | entregue
  criado_em, atualizado_em
)

briefing_analise (
  id, briefing_id,
  leitura, funis_json, padroes_json,
  criativos_analisados, criativos_em_leitura,
  funil_inferido_pct,        -- para o aviso de confiabilidade
  gerado_em
)

briefing_curadoria (
  id, briefing_id,
  leitura, candidatos_json,
  gerado_em
)

briefing_ficha (
  id, briefing_id, codigo, titulo,
  prioridade, rota,          -- nova | recorte
  funil, canal, formato,
  objetivo, referencia, publico, angulo, promessa,
  roteiro_json, copy_json,
  especificacoes, obrigatorio_json, proibido_json,
  prazo, ordem,
  status                     -- fila | producao | entregue
)

briefing_mensagem (
  id, briefing_id, papel,    -- gestor | agente
  texto, criado_em
)
```

Guarde a análise junto do briefing. Ela é o registro do que foi visto no momento da decisão — reprocessar depois com dados atualizados produz outra leitura e quebra a rastreabilidade.

---

## 11. API

```
POST   /briefings                          cria rascunho com os campos
POST   /briefings/{id}/gerar               dispara o fluxo completo (assíncrono)
GET    /briefings/{id}                     briefing com análise, curadoria e fichas
GET    /briefings/{id}/status              estado da geração
POST   /briefings/{id}/mensagem            ajuste via chat
PATCH  /briefings/{id}/fichas/{fid}        edição manual de campo
POST   /briefings/{id}/aprovar             congela e envia para a fila de produção
GET    /briefings/{id}/html                render do documento
GET    /clientes/{id}/criativos            criativos elegíveis do período
```

Geração passa de 30 segundos. Fila com status e webhook, nunca requisição síncrona. Exponha o progresso por etapa — analisando, curando, montando — porque a espera sem feedback parece travamento.

---

## 12. Tratamento de erro

| Situação | Comportamento |
|---|---|
| Nenhum criativo elegível no período | Não gere briefing às cegas. Informe e ofereça ampliar o período |
| JSON inválido do LLM | Uma tentativa de reparo, depois falha explícita. Nunca renderize parcial |
| Resposta truncada | Detecte pelo JSON não fechado. Regenere em lotes menores |
| Métrica ausente | Marque não disponível, siga sem ela, declare na saída |
| Funil inferido acima de 30% | Gere normalmente com aviso de confiabilidade reduzida |
| Integração de conteúdo indisponível | Curadoria falha isolada. Briefing sai como produção nova |

Falha de curadoria nunca derruba o briefing. É etapa opcional e deve degradar sozinha.

---

## 13. Critérios de aceite

1. Gestor preenche os campos obrigatórios e recebe briefing completo sem colar dado nenhum
2. Análise sai agrupada por funil, com melhores e piores, e todo criativo listado tem pontos positivos e negativos
3. Comparação de performance nunca cruza funis diferentes
4. Toda ficha tem roteiro e copy preenchidos
5. Com curadoria ligada e material disponível, ao menos parte das fichas sai como rota recorte, com direção de corte no lugar do roteiro
6. Bloqueio de áudio de biblioteca aparece antes de a peça ser sugerida
7. Chat de ajustes altera o solicitado e preserva o resto
8. Nenhum termo da lista proibida aparece na saída, verificado por teste automatizado
9. Análise fica persistida junto do briefing e não é recalculada em leituras posteriores
10. Falha de curadoria não impede a geração do briefing

---

## 14. Fora de escopo

Não construir agora, mesmo que a estrutura pareça pedir:

- Sugestão automática de quantidade a partir do planejamento
- Cadastro de ângulos com histórico consolidado
- Bancos de criativos com piso de estoque e rodízio
- Modo automático que propõe a hipótese sozinho
- Transcrição e segmentação automática de vídeo longo

Estes itens estão especificados em documentos separados e dependem deste agente funcionando primeiro.

---

## 15. Ordem sugerida

| Etapa | Escopo |
|---|---|
| 1 | Leitura dos criativos da ingestão + resolução de funil |
| 2 | Análise por funil, com persistência |
| 3 | Geração do briefing e render das fichas |
| 4 | Chat de ajustes |
| 5 | Curadoria de conteúdo, com entrada manual |
| 6 | Integração Instagram e YouTube substituindo a entrada manual |

Etapas 1 a 3 já entregam valor sozinhas: o gestor deixa de montar briefing na mão e o time passa a receber a leitura de performance junto. Rode assim com um cliente antes de seguir.
