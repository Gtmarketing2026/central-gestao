import { google } from "npm:googleapis@144";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

const BLOCKED_TAB_PATTERNS = [/site/i, /cadastro/i, /cpf/i, /lead/i];
function isBlockedTab(tab: string): boolean { return BLOCKED_TAB_PATTERNS.some((re) => re.test(tab)); }

function parseNumberBR(v: unknown): number {
  if (v == null) return 0;
  let s = String(v).replace(/R\$/gi, "").trim();
  if (!s || s === "-") return 0;
  s = s.replace(/\s/g, "");
  if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", ".");
  else if (s.includes(",")) s = s.replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

async function aggregateOrdersFromSheet(sheets: any, spreadsheetId: string, tab: string) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets(properties(title,gridProperties))" });
  const sheetMeta = (meta.data.sheets || []).find((s: any) => s.properties.title === tab);
  if (!sheetMeta) throw new Error(`Aba "${tab}" nao encontrada nessa planilha`);
  const rowCount = sheetMeta.properties.gridProperties.rowCount;
  const startRow = Math.max(2, rowCount - 60000);
  const batch = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges: [`'${tab}'!A1:Z1`, `'${tab}'!A${startRow}:Z${rowCount}`] });
  const headerRow = (batch.data.valueRanges[0].values || [[]])[0] || [];
  const header = headerRow.map((h: string) => String(h || "").toLowerCase().trim());
  const dateIdx = header.findIndex((h: string) => h === "data" || h === "day");
  const statusIdx = header.findIndex((h: string) => h === "status");
  const totalIdx = header.findIndex((h: string) => h === "total" || h === "valor");
  if (dateIdx === -1 || statusIdx === -1 || totalIdx === -1) throw new Error('Nao encontrei as colunas "Data", "Status" e "Total" nessa aba');
  const rows = batch.data.valueRanges[1].values || [];
  const entries: any[] = [];
  for (const row of rows) {
    const rawDate = row[dateIdx];
    if (!rawDate) continue;
    const m = String(rawDate).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) continue;
    entries.push({ date: `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`, status: row[statusIdx] || "Desconhecido", total: parseNumberBR(row[totalIdx]) });
  }
  return { entries, rowsScanned: rows.length, sheetRowCount: rowCount };
}

async function aggregateOrdersTabs(sheets: any, spreadsheetIds: string[], tab: string) {
  const agg: Record<string, any> = {};
  let rowsScanned = 0, sheetRowCount = 0;
  const errors: string[] = [];
  for (const spreadsheetId of spreadsheetIds) {
    try {
      const { entries, rowsScanned: rs, sheetRowCount: src } = await aggregateOrdersFromSheet(sheets, spreadsheetId, tab);
      rowsScanned += rs; sheetRowCount += src;
      for (const e of entries) {
        const key = e.date + "|" + e.status;
        if (!agg[key]) agg[key] = { date: e.date, status: e.status, count: 0, total: 0 };
        agg[key].count += 1; agg[key].total += e.total;
      }
    } catch (err) { errors.push(`${spreadsheetId}: ${(err as Error).message}`); }
  }
  if (!Object.keys(agg).length && errors.length) throw new Error(errors.join(" | "));
  return { aggregated: true, rows: Object.values(agg), rowsScanned, sheetRowCount, partialErrors: errors.length ? errors : undefined };
}

async function callOpenAI(body: any) {
  const apiKey = Deno.env.get("OPENAI_API_KEY");
  if (!apiKey) throw new Error("OPENAI_API_KEY nao configurada");
  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { "Authorization": `Bearer ${apiKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({ model: "gpt-4o-mini", temperature: 0.6, max_tokens: 1000, ...body }),
  });
  const json = await resp.json();
  if (!resp.ok) throw new Error(json.error?.message || "Erro na API da OpenAI");
  return json;
}

const DNA_SHAPE = `{
  "identidade": {"marca": "", "promessa": "", "posicionamento": "", "tom": "", "sobre": ""},
  "produtos": [{"nome": "", "dorQueResolve": "", "desejo": "", "personaAlvo": ""}],
  "personas": [{"titulo": "", "descricao": "", "transformacao": "de [estado atual] para [estado desejado]", "estadoAtual": "", "dores": [""], "desejos": [""], "tensoes": "", "crencas": ""}],
  "diretrizes": {"tom": "", "palavrasRessoam": [""], "palavrasProibidas": [""], "abordagens": [""], "beneficios": [""], "sempreFim": [""]}
}`;
const DNA_SYSTEM = `Voce e uma estrategista de marketing e copywriting senior. A partir do material fornecido sobre um cliente/negocio (briefing, site, questionario, material institucional), voce monta o "DNA" do cliente: identidade da marca, produtos/servicos, personas detalhadas e diretrizes de copy.

Responda SOMENTE com um JSON valido no formato exato abaixo (sem markdown, sem comentarios, em portugues do Brasil):
${DNA_SHAPE}

Regras:
- Preencha com base no material; nao invente fatos concretos (nomes, precos), mas PODE inferir dores/desejos/tom coerentes com o segmento.
- personas: crie de 2 a 4 personas ricas. Cada uma com titulo curto e descritivo, descricao de 1-2 frases, transformacao (de X para Y), estadoAtual (o "antes" concreto), 4-6 dores e 4-6 desejos especificos, tensoes recorrentes e crencas/mitos a quebrar.
- produtos: liste os produtos/servicos identificados; se so houver um negocio, crie 1-3 entradas. personaAlvo deve referenciar o titulo de uma das personas.
- diretrizes: tom de comunicacao, 6-12 palavras que ressoam, 4-8 palavras proibidas, 4-6 abordagens de copy, 4-6 beneficios principais, e 1-3 frases para "sempre no fim da copy" (CTA/assinatura).
- Se um campo nao tiver base, deixe string vazia ou array vazio, nunca invente dado factual.`;

async function extractDna(text: string, direcionamento: string) {
  let user = `Material do cliente:\n${String(text || "").slice(0, 24000)}`;
  if (direcionamento) user += `\n\nDirecionamento do gestor (leve em conta): ${direcionamento}`;
  const json = await callOpenAI({ messages: [{ role: "system", content: DNA_SYSTEM }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 3500, temperature: 0.7 });
  const content = json.choices?.[0]?.message?.content || "{}";
  return JSON.parse(content);
}

async function refineDna(dna: any, instrucao: string) {
  const sys = `Voce edita o DNA de um cliente (JSON). Aplique a instrucao do gestor ao DNA atual e devolva o DNA COMPLETO atualizado, no MESMO formato JSON, sem markdown. Formato:\n${DNA_SHAPE}\nMantenha tudo que nao foi pedido para mudar. Portugues do Brasil.`;
  const user = `DNA atual:\n${JSON.stringify(dna)}\n\nInstrucao: ${instrucao}`;
  const json = await callOpenAI({ messages: [{ role: "system", content: sys }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 3500, temperature: 0.5 });
  return JSON.parse(json.choices?.[0]?.message?.content || "{}");
}

async function fetchUrlText(url: string) {
  const r = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0" } });
  const html = await r.text();
  return html.replace(/<script[\s\S]*?<\/script>/gi, " ").replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 24000);
}

async function generateAnalysis(m: any, chat: any[], styleExamples: string[]) {
  let system = `Voce e uma gestora de trafego pago senior, especialista em performance (Meta Ads, Google Ads, funil de vendas e e-commerce). Escreve analises gerenciais mensais claras, diretas e acionaveis. Baseie-se SEMPRE nos numeros reais fornecidos, nunca invente dados. Responda apenas com o texto da analise, em portugues, sem markdown e sem titulos, em 2 a 4 paragrafos curtos.`;
  if (Array.isArray(styleExamples) && styleExamples.length) {
    system += `\n\nO gestor humano tem um estilo proprio de escrever. Imite o tom, o tamanho e a estrutura destes exemplos de analises anteriores dele:\n` + styleExamples.map((s, i) => `--- Exemplo ${i + 1} ---\n${s}`).join("\n\n");
  }
  const messages: any[] = [{ role: "system", content: system }];
  messages.push({ role: "user", content: `Dados do mes para o cliente "${m.clientName}", referente a ${m.mesLabel}:\n${JSON.stringify(m, null, 2)}\n\nGere a analise gerencial mensal.` });
  if (Array.isArray(chat)) for (const t of chat) messages.push({ role: t.role === "user" ? "user" : "assistant", content: String(t.text || "") });
  const json = await callOpenAI({ model: "gpt-4o", messages, max_tokens: 1200 });
  return json.choices?.[0]?.message?.content || "";
}

const AGENT_TOOLS = [
  { type: "function", function: { name: "criar_tarefa", description: "Cria uma nova tarefa/atividade no sistema da agencia.", parameters: { type: "object", properties: {
    nome: { type: "string" }, cliente: { type: "string" }, responsavel: { type: "string" },
    prioridade: { type: "string", enum: ["alta", "media", "baixa"] }, prazo: { type: "string", description: "YYYY-MM-DD" }, urgente: { type: "boolean" },
  }, required: ["nome"] } } },
  { type: "function", function: { name: "concluir_tarefa", description: "Marca uma tarefa existente como concluida (dar baixa).", parameters: { type: "object", properties: {
    nome: { type: "string" }, cliente: { type: "string" },
  }, required: ["nome"] } } },
  { type: "function", function: { name: "pausar_meta", description: "Pausa um anuncio, conjunto ou campanha no Meta Ads. Use quando algo esta drenando verba sem retorno. Pegue o 'id' e o 'nivel' da lista metaEntidades do snapshot.", parameters: { type: "object", properties: {
    id: { type: "string", description: "id do objeto no Meta (campaign/adset/ad)" }, nivel: { type: "string", enum: ["campanha", "conjunto", "anuncio"] }, nome: { type: "string", description: "nome legivel para o card de confirmacao" },
  }, required: ["id", "nivel", "nome"] } } },
  { type: "function", function: { name: "reativar_meta", description: "Reativa (liga) um anuncio, conjunto ou campanha pausado no Meta Ads. Use os dados de metaEntidades.", parameters: { type: "object", properties: {
    id: { type: "string" }, nivel: { type: "string", enum: ["campanha", "conjunto", "anuncio"] }, nome: { type: "string" },
  }, required: ["id", "nivel", "nome"] } } },
  { type: "function", function: { name: "ajustar_orcamento", description: "Ajusta o orcamento diario de uma campanha ou conjunto no Meta Ads. Informe OU 'percentual' (ex: 20 para subir 20%, -30 para descer 30%) OU 'novoOrcamentoDiario' em reais. Use os dados de metaEntidades.", parameters: { type: "object", properties: {
    id: { type: "string" }, nivel: { type: "string", enum: ["campanha", "conjunto"] }, nome: { type: "string" }, percentual: { type: "number" }, novoOrcamentoDiario: { type: "number", description: "em reais" },
  }, required: ["id", "nivel", "nome"] } } },
  { type: "function", function: { name: "duplicar_campanha", description: "Duplica uma campanha vencedora no Meta Ads (a copia nasce PAUSADA por seguranca). Use para escalar. Pegue o id da campanha em metaEntidades.", parameters: { type: "object", properties: {
    id: { type: "string" }, nome: { type: "string" },
  }, required: ["id", "nome"] } } },
];

async function runAgent(a: any) {
  let system = `Voce e a AndreIA, uma SUPER gestora de trafego (nivel "Jarvis") de uma agencia de performance de elite. Voce pensa e recomenda no nivel dos melhores gestores do Brasil (Pedro Sobral e outros que a agencia treinou em voce via BASE DE CONHECIMENTO abaixo). Voce olha TODOS OS PILARES e conecta eles: (1) TRAFEGO PAGO (estrutura de campanha, publico, leilao, orcamento, escala), (2) CRIATIVO (angulos, hook, formato, fadiga/saturacao, o que testar), (3) SITE/PAGINA e FUNIL/CRO (conversao, checkout, oferta, prova social, velocidade). Uma metrica ruim num pilar quase sempre tem causa em outro — diga qual e por que. Seja uma consultora tecnica de verdade: especifica, com numeros do snapshot, priorizada, e com o "porque" por tras (nao conselho generico de manual).

⚠️ REGRA #1 (INEGOCIAVEL) — NUNCA JULGUE UMA CAMPANHA/ANUNCIO POR ROAS SE O OBJETIVO DELA NAO FOR VENDA.
Cada anuncio no snapshot tem 'objetivo' (tipo + metrica de sucesso), 'metricaDoObjetivo' e 'avaliacao' (BOM/RUIM/observar JA calculado pelo objetivo correto). USE a 'avaliacao' e a 'metricaDoObjetivo' — NAO recalcule por ROAS. Exemplos:
- objetivo TRAFEGO: sucesso = CPC baixo e CTR saudavel. ROAS 0 aqui e NORMAL e NAO significa "drenando verba". NUNCA sugira pausar campanha de trafego so porque nao teve venda.
- objetivo MENSAGENS: sucesso = conversas iniciadas e custo por conversa.
- objetivo VIDEO/DISTRIBUICAO: sucesso = views e custo por view.
- objetivo ENGAJAMENTO: sucesso = engajamentos e custo por engajamento.
- objetivo LEADS: sucesso = CPL e volume de leads.
- objetivo CONVERSAO/VENDAS: ai sim ROAS/CPA/compras.
Ao falar de um anuncio, SEMPRE diga o objetivo dele e avalie pela metricaDoObjetivo. So chame de "drenando verba" quando a 'avaliacao' for RUIM.

FOCO PRINCIPAL: analisar os resultados (os big numbers) e RECOMENDAR OTIMIZACOES TECNICAS de campanha, funil e pagina. Gestao de tarefas e secundaria.

OBJETIVO DO CLIENTE MANDA (leia 'objetivosDoCliente' e 'temVenda' no snapshot): analise SO pelas metricas do objetivo dele. Se 'temVenda' for false (cliente sem objetivo de venda/conversao — ex: Mensagens, Trafego, Video, Alcance, Engajamento): NAO cite ROAS, faturamento, receita, CPA nem "nenhuma venda registrada"; NAO liste metricas de venda zeradas. Foque na metrica-chave do objetivo (ex: custo por conversa e nº de conversas p/ Mensagens; CPL e leads p/ Leads; custo por view p/ Video; CPC/CTR p/ Trafego) + CTR/CPC/CPM de eficiencia. Em 'metasCliente', o status ja diz: 'atingida' = ok; 'abaixo_do_alvo_ruim' = piorou numa metrica onde MAIOR e melhor; 'acima_do_alvo_ruim' = piorou numa metrica onde MENOR e melhor (custos). Nunca diga so "abaixo da meta" sem dizer se isso e bom ou ruim.

DOCUMENTOS: Voce PODE montar documentos (relatorios, propostas, briefings, planos de acao, resumos executivos). Quando pedirem um documento/relatorio/PDF/Word, NUNCA diga que nao consegue gerar arquivos — escreva o CONTEUDO COMPLETO e bem formatado em markdown (titulos com #, ## e ###, listas, **negrito**, e tabelas em markdown com | quando fizer sentido) direto na resposta. Ao terminar, avise: "Pronto — clique em '📄 Baixar como documento' abaixo pra salvar em PDF ou Word, escolhendo o layout." O sistema converte sua resposta no layout da agencia (temas GT) automaticamente. Estruture como documento de verdade: titulo, secoes, e quando for relatorio de cliente siga a logica dos nossos templates (visao geral -> resultados por objetivo -> funil -> recomendacoes/proximos passos).

Baseie-se SOMENTE nos dados do snapshot (KPIs do relatorio do cliente, canais, funil, pedidos). Nunca invente numeros; se faltar um dado, diga que nao esta disponivel. Seja direta, especifica e priorize; nada de conselho generico de manual.

Metodo de analise:
1. Leia os KPIs: investimento, CTR, CPC, CPM, ROAS, ticket medio, pedidos e o funil (impressoes -> cliques -> checkout -> venda).
2. Diagnostique ONDE esta o gargalo:
   - CTR baixo (feed abaixo de 1%, search abaixo de 2%): problema de criativo/oferta/segmentacao. Recomende novos angulos de criativo, hook nos primeiros 3s, revisar publico e a headline.
   - CTR ok mas poucas vendas / checkout baixo: problema de pagina ou oferta. Recomende CRO: headline mais clara, prova social, velocidade e mobile, reduzir friccao, revisar oferta/garantia, preco/parcelamento.
   - Checkout iniciado mas nao converte: friccao no checkout, formas de pagamento, confianca.
   - ROAS bom e estavel: escalar (subir orcamento ~20% por vez, duplicar campanhas vencedoras). ROAS caindo: pausar o que nao performa, revisar publicos saturados, renovar criativos.
   - CPM subindo: saturacao de publico ou leilao concorrido; teste novos publicos/criativos.
3. Compare os canais entre si e realoque verba para quem tem melhor ROAS/CPA.
4. Traga recomendacoes priorizadas (o que fazer primeiro), usando os numeros reais do snapshot.

REGRA CRITICA — ANALISE POR OBJETIVO DA CAMPANHA (nunca julgue tudo como venda):
Cada campanha/anuncio no snapshot tem um campo 'objetivo' (tipo + metrica de sucesso). Avalie SEMPRE pela metrica do objetivo dela, NAO por ROAS/vendas cegamente. Um ROAS 0 numa campanha de trafego/engajamento/alcance NAO significa que ela esta ruim — ela nem tem venda como meta. Playbook:
- objetivo 'conversao' (Vendas/Conversoes): ai sim julgue por ROAS, CPA e nº de compras. ROAS baixo/0 com gasto alto = ruim; ROAS bom = escalar.
- objetivo 'trafego' (Trafego/Cliques): julgue por CPC e CTR e volume de cliques. Bom = CPC baixo e CTR saudavel (feed >1%). NAO recomende pausar por falta de venda; se o CPC/CTR estao bons, a campanha esta cumprindo o objetivo.
- objetivo 'engajamento': julgue por custo por engajamento, CTR e alcance. Venda nao e a meta.
- objetivo 'leads': julgue por CPL (custo por lead = investimento / nº de leads) e volume de leads. Nao por ROAS de compra.
- objetivo 'alcance'/awareness: julgue por CPM, alcance e frequencia (frequencia alta = saturacao). Nao por venda.
- objetivo 'video': julgue por custo por ThruPlay/visualizacao e CPM.
- objetivo 'mensagens': julgue por custo por conversa iniciada.
Sempre diga explicitamente o objetivo da campanha e por qual metrica voce a esta avaliando. Se propuser pausar/ajustar, so faca sentido dentro do objetivo (ex: pausar uma campanha de CONVERSAO com gasto alto e 0 compras — nunca uma de trafego so porque nao vendeu).

PERIODOS E BENCHMARK:
- Se o snapshot tiver 'periodos' (ultimos7dias, ultimos30dias), use esses numeros quando perguntarem sobre 7 ou 30 dias.
- Se tiver 'benchmarkProprioCliente' (variacao % dos ultimos 30d vs os 30d anteriores), use como BENCHMARK DO PROPRIO CLIENTE — diga o que melhorou/piorou vs o historico dele (ex: "CTR subiu 12%, CPC caiu 8% vs o mes anterior").
- Se tiver 'metasCliente' (metas definidas pelo gestor + status atingida/abaixo vs ultimos 30d), esse e o BENCHMARK-ALVO OFICIAL do cliente: priorize ele. Diga claramente o que bateu a meta e o que ficou abaixo, com o numero da meta e o real.
- COMPARACAO DE MERCADO (referencias gerais do Meta Ads, use como parametro aproximado, nunca como verdade absoluta e sempre considerando o nicho): CTR no feed bom > 1% (otimo > 2%); CPC saudavel geralmente < R$2 (varia muito por nicho); frequencia > 3-4 no periodo indica saturacao; taxa de conversao de LP e-commerce tipica 1-3%; checkout->compra saudavel 30-50%; ROAS bom depende da margem, mas < 1 e prejuizo e > 2-3 costuma ser saudavel em e-commerce. Ao comparar com mercado, diga "acima/abaixo da media de mercado" com o numero.

Se o snapshot tiver 'dnaCliente' (identidade, produtos, personas com dores/desejos, diretrizes de copy), USE como base ao sugerir angulos de criativo, headlines, copies e publico — respeite o tom, as palavras que ressoam e evite as proibidas.

Voce tambem pode EXECUTAR acoes quando o gestor pedir explicitamente: criar/concluir tarefas E acoes reais no Meta Ads (pausar_meta, reativar_meta, ajustar_orcamento, duplicar_campanha). Para as acoes do Meta, use SEMPRE o 'id' e o 'nivel' que estao na lista 'metaEntidades' do snapshot (campanhas, conjuntos e anuncios com id, status e orcamento atuais) — nunca invente ids. O sistema mostra um card de confirmacao antes de executar; entao apenas PROPONHA a acao chamando a funcao e explique o porque em texto; nunca afirme que ja executou. So proponha acao no Meta quando o gestor pedir ou quando os dados claramente justificarem (ex: anuncio com gasto alto e 0 compras -> propor pausar). Seu valor principal continua sendo a analise tecnica.

===== FORMATO DAS RESPOSTAS (OBRIGATORIO) =====
O gestor nao tem tempo de ler textao. Toda resposta deve ser ESCANEAVEL:
- Comece com 1 linha de resposta direta (a conclusao primeiro, nao no final).
- Estruture com titulos curtos em **negrito** e blocos separados por linha em branco.
- Estrategias e planos: SEMPRE em passo a passo numerado (1. 2. 3.) ou checklist (☐), um item por linha, cada item com no maximo 2 linhas.
- Numeros e metricas: em bullets "• Metrica: valor (leitura)", nunca dissolvidos num paragrafo.
- Nada de paragrafos com mais de 3 linhas. Pode ser extensa SE necessario, desde que dividida em blocos claros.
- Feche com "**Proximo passo:**" quando fizer sentido (1 acao concreta).

RESUMO PARA CLIENTE (quando pedirem resumo/relatorio pro cliente): escreva PRONTO PRA COLAR NO WHATSAPP:
- Saudacao curta + periodo. Emojis com moderacao (📊 ✅ 🎯).
- Bullets curtos com os numeros que importam pro cliente (sem jargao tecnico: nada de CPM/CTR sem explicar).
- 1 bloco "O que faremos agora" com 2-3 acoes.
- Maximo ~15 linhas. Tom profissional e proximo, sem markdown de titulo (#), so *negrito estilo WhatsApp* com um asterisco.`;

  if (Array.isArray(a.knowledge) && a.knowledge.length) {
    system += `\n\n===== BASE DE CONHECIMENTO (JARVIS) =====\nEstes sao os metodos e frameworks dos gestores que a agencia treinou em voce (Pedro Sobral e outros). Eles sao a SUA forma de pensar: aplique estes principios, benchmarks e mentalidade em TODA analise e recomendacao, citando o raciocinio quando util. Nao os ignore.\n` +
      a.knowledge.map((k: any, i: number) => `--- Fonte ${i + 1}: ${k.title || "material"} ---\n${String(k.text || "").slice(0, 14000)}`).join("\n\n");
  }

  const messages: any[] = [{ role: "system", content: system }];
  messages.push({ role: "user", content: `Snapshot atual (dados reais do sistema):\n${JSON.stringify(a.snapshot, null, 2)}` });
  if (Array.isArray(a.history)) for (const t of a.history) messages.push({ role: t.role === "user" ? "user" : "assistant", content: String(t.text || "") });

  // URLs enviadas pelo gestor: busca o conteúdo da página e entrega como contexto (análise de sites)
  if (Array.isArray(a.urls) && a.urls.length) {
    for (const u of a.urls.slice(0, 3)) {
      try { const t = await fetchUrlText(String(u)); messages.push({ role: "user", content: `Conteúdo da página ${u} (texto extraído):\n${t.slice(0, 12000)}` }); }
      catch (_e) { messages.push({ role: "user", content: `(Não consegui acessar a URL ${u} — avise o gestor.)` }); }
    }
  }
  // Anexos: imagens vão como vision (gpt-4o); PDFs/textos já chegam extraídos do front
  if (Array.isArray(a.anexos) && a.anexos.length) {
    const imgs: any[] = [];
    for (const ax of a.anexos.slice(0, 4)) {
      if (ax.tipo === "imagem" && ax.dataUrl) imgs.push({ type: "image_url", image_url: { url: String(ax.dataUrl) } });
      else if (ax.texto) messages.push({ role: "user", content: `Anexo "${ax.nome || "arquivo"}" (texto extraído):\n${String(ax.texto).slice(0, 15000)}` });
    }
    if (imgs.length) messages.push({ role: "user", content: [{ type: "text", text: "Imagem(ns) anexada(s) pelo gestor — analise:" }, ...imgs] });
  }

  const json = await callOpenAI({ model: "gpt-4o", messages, tools: AGENT_TOOLS, tool_choice: "auto", max_tokens: 2000, temperature: 0.5 });
  const msg = json.choices?.[0]?.message || {};
  const actions: any[] = [];
  if (Array.isArray(msg.tool_calls)) {
    for (const tc of msg.tool_calls) {
      try { actions.push({ name: tc.function.name, args: JSON.parse(tc.function.arguments || "{}") }); } catch (_e) { /* ignora */ }
    }
  }
  return { answer: msg.content || "", actions };
}

// Normaliza o objetivo da campanha do Meta em um tipo + metrica de sucesso, para analise correta.
function metaObjetivo(obj: string) {
  const o = String(obj || "").toUpperCase();
  const map: Record<string, { tipo: string; rotulo: string; metrica: string }> = {
    OUTCOME_SALES: { tipo: "conversao", rotulo: "Vendas/Conversão", metrica: "ROAS, CPA, nº de compras" },
    CONVERSIONS: { tipo: "conversao", rotulo: "Conversões", metrica: "ROAS, CPA, nº de compras" },
    CATALOG_SALES: { tipo: "conversao", rotulo: "Vendas de catálogo", metrica: "ROAS, CPA" },
    PRODUCT_CATALOG_SALES: { tipo: "conversao", rotulo: "Vendas de catálogo", metrica: "ROAS, CPA" },
    OUTCOME_LEADS: { tipo: "leads", rotulo: "Cadastros (Leads)", metrica: "CPL (custo por lead), nº de leads" },
    LEAD_GENERATION: { tipo: "leads", rotulo: "Geração de leads", metrica: "CPL, nº de leads" },
    OUTCOME_TRAFFIC: { tipo: "trafego", rotulo: "Tráfego", metrica: "CPC, CTR, cliques no link" },
    LINK_CLICKS: { tipo: "trafego", rotulo: "Cliques no link", metrica: "CPC, CTR" },
    OUTCOME_ENGAGEMENT: { tipo: "engajamento", rotulo: "Engajamento", metrica: "custo por engajamento, CTR, alcance" },
    POST_ENGAGEMENT: { tipo: "engajamento", rotulo: "Engajamento", metrica: "custo por engajamento, CTR" },
    PAGE_LIKES: { tipo: "engajamento", rotulo: "Curtidas de página", metrica: "custo por curtida" },
    EVENT_RESPONSES: { tipo: "engajamento", rotulo: "Respostas a evento", metrica: "custo por resposta" },
    VIDEO_VIEWS: { tipo: "video", rotulo: "Visualizações de vídeo", metrica: "custo por ThruPlay/view, CPM" },
    MESSAGES: { tipo: "mensagens", rotulo: "Mensagens", metrica: "custo por conversa iniciada" },
    OUTCOME_AWARENESS: { tipo: "alcance", rotulo: "Reconhecimento/Alcance", metrica: "CPM, alcance, frequência" },
    BRAND_AWARENESS: { tipo: "alcance", rotulo: "Reconhecimento de marca", metrica: "CPM, alcance" },
    REACH: { tipo: "alcance", rotulo: "Alcance", metrica: "CPM, alcance, frequência" },
    APP_INSTALLS: { tipo: "app", rotulo: "Instalações de app", metrica: "custo por instalação" },
    OUTCOME_APP_PROMOTION: { tipo: "app", rotulo: "Promoção de app", metrica: "custo por instalação/evento" },
  };
  return { codigo: o || null, ...(map[o] || { tipo: "outro", rotulo: obj || "Não informado", metrica: "métrica do objetivo" }) };
}

async function metaAdsInsights(m: any) {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  // aceita: accounts [{id,name}], accountIds [id], ou accountId (compat)
  let accounts: { id: string; name: string }[] = [];
  if (Array.isArray(m.accounts) && m.accounts.length) accounts = m.accounts.map((a: any) => ({ id: String(a.id).replace(/^act_/, ""), name: a.name || "" }));
  else if (Array.isArray(m.accountIds) && m.accountIds.length) accounts = m.accountIds.map((id: any) => ({ id: String(id).replace(/^act_/, ""), name: "" }));
  else if (m.accountId) accounts = [{ id: String(m.accountId).replace(/^act_/, ""), name: "" }];
  if (!accounts.length) throw new Error("accountId(s) obrigatorio");
  const multi = accounts.length > 1;
  const ver = "v21.0";
  const base = `https://graph.facebook.com/${ver}`;
  let range = "";
  if (m.since && m.until) range = `&time_range=${encodeURIComponent(JSON.stringify({ since: m.since, until: m.until }))}`;
  else range = `&date_preset=${m.datePreset || "last_30d"}`;
  const fields = "spend,impressions,clicks,ctr,cpc,cpm,reach,frequency,actions,action_values,purchase_roas";

  async function fetchInsights(acct: string, level: string, extra = "") {
    let lvlFields = "";
    if (level === "campaign") lvlFields = ",campaign_name,campaign_id";
    else if (level === "ad") lvlFields = ",campaign_name,campaign_id,adset_name,adset_id,ad_name,ad_id";
    let url: string | null = `${base}/act_${acct}/insights?level=${level}&fields=${fields}${lvlFields}${range}${extra}&use_unified_attribution_setting=true&limit=200&access_token=${token}`;
    const out: any[] = [];
    for (let i = 0; i < 20 && url; i++) {
      const r = await fetch(url);
      const j = await r.json();
      if (j.error) throw new Error(j.error.message);
      out.push(...(j.data || []));
      url = j.paging?.next || null;
    }
    return out;
  }
  // objetivo por campaign_id (uma chamada por conta)
  async function fetchObjectives(acct: string) {
    const map: Record<string, any> = {};
    let url: string | null = `${base}/act_${acct}/campaigns?fields=id,objective&limit=200&access_token=${token}`;
    for (let i = 0; i < 20 && url; i++) {
      const r = await fetch(url);
      const j = await r.json();
      if (j.error) break;
      for (const c of (j.data || [])) map[c.id] = metaObjetivo(c.objective);
      url = j.paging?.next || null;
    }
    // Local de conversão: nomenclatura engana (ex: campanha com "venda" no nome mas objetivo Mensagem).
    // Lê destination_type/optimization_goal dos conjuntos e refina o tipo por campanha.
    try {
      let aurl: string | null = `${base}/act_${acct}/adsets?fields=campaign_id,destination_type,optimization_goal&limit=500&access_token=${token}`;
      for (let i = 0; i < 10 && aurl; i++) {
        const r = await fetch(aurl);
        const j = await r.json();
        if (j.error) break;
        for (const as of (j.data || [])) {
          const ob = map[as.campaign_id];
          if (!ob) continue;
          const dest = String(as.destination_type || "").toUpperCase();
          const opt = String(as.optimization_goal || "").toUpperCase();
          // Mensagem: otimiza pra conversas ou destino de mensagens — mas venda (OUTCOME_SALES) continua venda mesmo via WhatsApp
          if (ob.tipo !== "conversao" && (opt === "CONVERSATIONS" || /MESSENGER|WHATSAPP|INSTAGRAM_DIRECT|MESSAGING/.test(dest))) {
            map[as.campaign_id] = { ...ob, tipo: "mensagens", rotulo: "Mensagens", metrica: "custo por conversa iniciada" };
          } else if (opt === "THRUPLAY" || opt === "TWO_SECOND_CONTINUOUS_VIDEO_VIEWS") {
            map[as.campaign_id] = { ...ob, tipo: "video", rotulo: "Vídeo / Distribuição", metrica: "custo por ThruPlay/view, CPM" };
          } else if (opt === "REACH" && ob.tipo === "engajamento") {
            map[as.campaign_id] = { ...ob, tipo: "alcance", rotulo: "Alcance / Distribuição", metrica: "CPM, alcance, frequência" };
          }
        }
        aurl = j.paging?.next || null;
      }
    } catch (_e) { /* sem local de conversão: mantém o objetivo puro */ }
    return map;
  }
  // miniaturas SÓ dos ad_ids informados (batch /?ids=, sem paginar a conta inteira)
  async function fetchThumbsByIds(adIds: string[]) {
    const map: Record<string, string> = {};
    for (let i = 0; i < adIds.length; i += 50) {
      const chunk = adIds.slice(i, i + 50);
      const r = await fetch(`${base}/?ids=${chunk.join(",")}&fields=creative{thumbnail_url,image_url}&access_token=${token}`);
      const j = await r.json();
      if (j.error) continue;
      for (const id of chunk) {
        const cr = j[id]?.creative;
        const t = cr?.thumbnail_url || cr?.image_url;
        if (t) map[id] = t;
      }
    }
    return map;
  }
  // Pega UM tipo canonico (o primeiro presente na ordem de prioridade), evitando somar tipos sobrepostos (que contam em dobro, como o Meta faz).
  function pickOne(arr: any[], types: string[]) {
    if (!Array.isArray(arr)) return 0;
    for (const ty of types) { const hit = arr.find((x) => x.action_type === ty); if (hit) return parseFloat(hit.value || "0"); }
    return 0;
  }
  function shape(row: any) {
    const purchases = pickOne(row.actions, ["omni_purchase", "offsite_conversion.fb_pixel_purchase", "purchase"]);
    const revenue = pickOne(row.action_values, ["omni_purchase", "offsite_conversion.fb_pixel_purchase", "purchase"]);
    const roas = Array.isArray(row.purchase_roas) && row.purchase_roas.length ? parseFloat(row.purchase_roas[0].value || "0") : (parseFloat(row.spend || "0") ? revenue / parseFloat(row.spend) : 0);
    return {
      campaign: row.campaign_name || null, campaignId: row.campaign_id || null,
      spend: parseFloat(row.spend || "0"), impressions: parseInt(row.impressions || "0"), clicks: parseInt(row.clicks || "0"),
      ctr: parseFloat(row.ctr || "0"), cpc: parseFloat(row.cpc || "0"), cpm: parseFloat(row.cpm || "0"),
      reach: parseInt(row.reach || "0"), frequency: parseFloat(row.frequency || "0"),
      purchases, revenue, roas,
      leads: pickOne(row.actions, ["offsite_conversion.fb_pixel_lead", "onsite_conversion.lead_grouped", "leadgen_grouped", "lead"]),
      addToCart: pickOne(row.actions, ["omni_add_to_cart", "offsite_conversion.fb_pixel_add_to_cart", "add_to_cart"]),
      initiateCheckout: pickOne(row.actions, ["omni_initiated_checkout", "offsite_conversion.fb_pixel_initiate_checkout", "initiate_checkout"]),
      conversas: pickOne(row.actions, ["onsite_conversion.messaging_conversation_started_7d", "messaging_conversation_started_7d", "onsite_conversion.total_messaging_connection"]),
      videoViews: pickOne(row.actions, ["video_view"]),
      engajamentos: pickOne(row.actions, ["post_engagement"]),
    };
  }
  const totAgg: any = { spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0, videoViews: 0, engajamentos: 0 };
  const byCamp: Record<string, any> = {};
  const ads: any[] = [];
  const wantObj = m.byAd || m.byCampaign;
  // Contas em PARALELO, e dentro de cada conta as chamadas (conta/objetivos/anuncios/campanhas/thumbs) tambem em paralelo.
  // Status da conta (restrita/desativada NAO gera erro na API de insights — precisa checar explicitamente)
  async function fetchAccountStatus(acct: string): Promise<string | null> {
    try {
      const r = await fetch(`${base}/act_${acct}?fields=account_status,disable_reason&access_token=${token}`);
      const j = await r.json();
      if (j.error) return null; // erro de chamada ja e tratado pelo catch das insights
      const st = Number(j.account_status);
      const dr = Number(j.disable_reason || 0);
      const drTxt = dr === 3 ? " (motivo: pagamento/risco)" : dr === 1 || dr === 5 ? " (motivo: política de anúncios)" : dr ? ` (código do Meta: ${dr})` : "";
      if (st === 1) return null;
      if (st === 3) return "Conta RESTRITA por pagamento — o último pagamento não foi processado; os anúncios estão parados até regularizar.";
      if (st === 2) return "Conta DESATIVADA pelo Meta" + drTxt + ".";
      if (st === 9) return "Conta em período de carência de pagamento — regularize pra não parar os anúncios.";
      if (st === 100 || st === 101) return "Conta encerrada/em encerramento no Meta.";
      if (st === 7 || st === 8) return "Conta pendente de análise/acerto no Meta.";
      return `Conta com status atípico no Meta (código ${st}).`;
    } catch (_e) { return null; }
  }
  const perAccount = await Promise.all(accounts.map(async (acc) => {
    const statusIssue = await fetchAccountStatus(acc.id);
    try {
      const [accountRows, acctDaily, objByCampId, adRows, campRows, campDedup] = await Promise.all([
        fetchInsights(acc.id, "account"),
        m.daily ? fetchInsights(acc.id, "account", "&time_increment=1") : Promise.resolve([] as any[]),
        wantObj ? fetchObjectives(acc.id) : Promise.resolve({} as Record<string, any>),
        m.byAd ? fetchInsights(acc.id, "ad") : Promise.resolve([] as any[]),
        m.byCampaign ? fetchInsights(acc.id, "campaign", m.daily ? "&time_increment=1" : "") : Promise.resolve([] as any[]),
        // ALCANCE nao e somavel: a busca diaria (time_increment=1) soma a mesma pessoa a cada dia.
        // Buscamos tambem SEM quebra diaria pra ter o reach/frequencia DEDUPLICADO por campanha no periodo.
        (m.byCampaign && m.daily) ? fetchInsights(acc.id, "campaign", "") : Promise.resolve([] as any[]),
      ]);
      return { acc, accountRows, acctDaily, objByCampId, adRows, campRows, campDedup, error: statusIssue as string | null };
    } catch (e) {
      // conta com erro NAO derruba as outras: devolve vazia + motivo (front mostra o disclaimer)
      return { acc, accountRows: [] as any[], acctDaily: [] as any[], objByCampId: {} as Record<string, any>, adRows: [] as any[], campRows: [] as any[], campDedup: [] as any[], error: statusIssue || (e as any)?.message || String(e) };
    }
  }));
  const accountErrors = perAccount.filter((p) => p.error).map((p) => ({ id: p.acc.id, name: p.acc.name || p.acc.id, error: p.error }));
  const totRecByDate: Record<string, any> = {};
  for (const { acc, accountRows, acctDaily, objByCampId, adRows, campRows, campDedup } of perAccount) {
    for (const row of acctDaily) {
      const s = shape(row); const k = row.date_start;
      if (!totRecByDate[k]) totRecByDate[k] = { date: k, sales: 0, spend: 0, revenue: 0, clicks: 0, impressions: 0, reach: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0, addToCart: 0, checkout: 0 };
      const rec = totRecByDate[k];
      rec.sales += Math.round(s.purchases); rec.spend += s.spend; rec.revenue += s.revenue; rec.clicks += s.clicks; rec.impressions += s.impressions;
      rec.reach += s.reach; rec.leads += s.leads; rec.conversas += s.conversas; rec.videoViews += s.videoViews; rec.engajamentos += s.engajamentos; rec.addToCart += s.addToCart; rec.checkout += s.initiateCheckout;
    }
    const at = accountRows.length ? shape(accountRows[0]) : shape({});
    totAgg.spend += at.spend; totAgg.impressions += at.impressions; totAgg.clicks += at.clicks; totAgg.reach += at.reach;
    totAgg.revenue += at.revenue; totAgg.purchases += at.purchases; totAgg.leads += at.leads; totAgg.addToCart += at.addToCart; totAgg.initiateCheckout += at.initiateCheckout;
    totAgg.conversas += at.conversas; totAgg.videoViews += at.videoViews; totAgg.engajamentos += at.engajamentos;
    for (const row of adRows) {
      const s = shape(row);
      ads.push({
        adId: row.ad_id, adName: row.ad_name || "(sem nome)", campaign: row.campaign_name || "", campaignId: row.campaign_id || null, adset: row.adset_name || "", adsetId: row.adset_id || null,
        account: acc.name || acc.id, thumbnail: null,
        objetivo: objByCampId[row.campaign_id] || metaObjetivo(""),
        spend: s.spend, impressions: s.impressions, clicks: s.clicks, reach: s.reach, frequency: s.frequency,
        ctr: s.ctr, cpc: s.cpc, cpm: s.cpm, purchases: s.purchases, revenue: s.revenue, roas: s.roas,
        leads: s.leads, addToCart: s.addToCart, initiateCheckout: s.initiateCheckout,
        conversas: s.conversas, videoViews: s.videoViews, engajamentos: s.engajamentos,
        cpa: s.purchases ? s.spend / s.purchases : 0,
      });
    }
    for (const row of campRows) {
      const label = row.campaign_name || "Meta Ads";
      const s = shape(row);
      if (!byCamp[label]) byCamp[label] = { campaign: label, account: acc.name || acc.id, objetivo: objByCampId[row.campaign_id] || metaObjetivo(""), spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, records: [] };
      const c = byCamp[label];
      // NAO soma reach aqui: metricas aditivas (spend/impressoes/etc) somam por dia; reach vem do dedup abaixo
      c.spend += s.spend; c.impressions += s.impressions; c.clicks += s.clicks;
      c.revenue += s.revenue; c.purchases += s.purchases; c.leads += s.leads; c.addToCart += s.addToCart; c.initiateCheckout += s.initiateCheckout;
      if (m.daily) c.records.push({ date: row.date_start, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: s.reach, leads: s.leads, conversas: s.conversas, videoViews: s.videoViews, engajamentos: s.engajamentos });
    }
    // reach/frequencia DEDUPLICADO por campanha no periodo (fonte nao-diaria; quando m.daily=false, campRows ja e o dedup)
    const dedupSrc = (m.daily ? campDedup : campRows) as any[];
    for (const row of dedupSrc) {
      const label = row.campaign_name || "Meta Ads";
      const s = shape(row);
      if (byCamp[label]) byCamp[label].reach += s.reach; // soma so entre contas (audiencias distintas), nunca entre dias/anuncios
    }
  }
  const total = {
    ...totAgg,
    ctr: totAgg.impressions ? (totAgg.clicks / totAgg.impressions) * 100 : 0,
    cpc: totAgg.clicks ? totAgg.spend / totAgg.clicks : 0,
    cpm: totAgg.impressions ? (totAgg.spend / totAgg.impressions) * 1000 : 0,
    roas: totAgg.spend ? totAgg.revenue / totAgg.spend : 0,
    records: Object.values(totRecByDate).sort((a: any, b: any) => a.date < b.date ? -1 : 1),
  };
  const campaigns = Object.values(byCamp).map((c: any) => {
    c.ctr = c.impressions ? (c.clicks / c.impressions) * 100 : 0;
    c.cpc = c.clicks ? c.spend / c.clicks : 0;
    c.cpm = c.impressions ? (c.spend / c.impressions) * 1000 : 0;
    c.roas = c.spend ? c.revenue / c.spend : 0;
    c.frequency = c.reach ? c.impressions / c.reach : 0; // impressoes(periodo) / alcance dedup
    return c;
  }).sort((a: any, b: any) => b.spend - a.spend);
  ads.sort((a: any, b: any) => b.spend - a.spend);
  if (m.byAd && ads.length) {
    const topIds = ads.slice(0, 20).map((a: any) => a.adId).filter(Boolean);
    const thumbs = await fetchThumbsByIds(topIds);
    for (const a of ads) if (thumbs[a.adId]) a.thumbnail = thumbs[a.adId];
  }
  return { total, campaigns, ads, accounts, accountErrors, period: m.since && m.until ? { since: m.since, until: m.until } : { datePreset: m.datePreset || "last_30d" } };
}

// Saldo pré-pago da conta (pix/boleto): funding_source_details traz o saldo disponível
async function metaFunding(m: any) {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const accounts = (Array.isArray(m.accounts) ? m.accounts : []).map((a: any) => ({ id: String(a.id).replace(/^act_/, ""), name: a.name || "" }));
  if (!accounts.length) throw new Error("accounts obrigatorio");
  const base = "https://graph.facebook.com/v21.0";
  const out: any[] = [];
  await Promise.all(accounts.map(async (acc: any) => {
    try {
      const r = await fetch(`${base}/act_${acc.id}?fields=name,account_status,funding_source_details&access_token=${token}`);
      const j = await r.json();
      if (j.error) { out.push({ id: acc.id, name: acc.name, error: j.error.message }); return; }
      const ds = j.funding_source_details || {};
      const disp = String(ds.display_string || "");
      let saldo: number | null = null;
      const mm = disp.match(/R\$\s?([\d\.]+,\d{2}|[\d\.]+)/);
      if (mm) saldo = parseFloat(mm[1].replace(/\./g, "").replace(",", "."));
      out.push({ id: acc.id, name: j.name || acc.name, display: disp, saldo, tipo: ds.type ?? null, status: j.account_status });
    } catch (e) { out.push({ id: acc.id, name: acc.name, error: (e as any)?.message || String(e) }); }
  }));
  return { accounts: out };
}

async function metaListAccounts() {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const out: any[] = [];
  let url = `https://graph.facebook.com/v21.0/me/adaccounts?fields=name,account_id,account_status,currency&limit=200&access_token=${token}`;
  for (let i = 0; i < 10 && url; i++) {
    const r = await fetch(url); const j = await r.json();
    if (j.error) throw new Error(j.error.message);
    out.push(...(j.data || []));
    url = j.paging?.next || "";
  }
  return out.map((a) => ({ id: a.account_id, name: a.name, status: a.account_status, currency: a.currency }));
}

// Lista entidades acionaveis (campanhas, conjuntos, anuncios) com id/status/orcamento atuais.
async function metaEntities(m: any) {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  let accounts: { id: string; name: string }[] = [];
  if (Array.isArray(m.accounts) && m.accounts.length) accounts = m.accounts.map((a: any) => ({ id: String(a.id).replace(/^act_/, ""), name: a.name || "" }));
  else if (Array.isArray(m.accountIds) && m.accountIds.length) accounts = m.accountIds.map((id: any) => ({ id: String(id).replace(/^act_/, ""), name: "" }));
  else if (m.accountId) accounts = [{ id: String(m.accountId).replace(/^act_/, ""), name: "" }];
  if (!accounts.length) throw new Error("accountId(s) obrigatorio");
  const base = "https://graph.facebook.com/v21.0";
  async function pageAll(path: string) {
    const out: any[] = [];
    let url: string | null = `${base}/${path}${path.includes("?") ? "&" : "?"}limit=200&access_token=${token}`;
    for (let i = 0; i < 15 && url; i++) {
      const r = await fetch(url); const j = await r.json();
      if (j.error) throw new Error(j.error.message);
      out.push(...(j.data || []));
      url = j.paging?.next || null;
    }
    return out;
  }
  const lightStatus = Array.isArray(m.adIds) && m.adIds.length; // modo leve: só status dos anúncios da tela
  const campaigns: any[] = [], adsets: any[] = [], ads: any[] = [];
  for (const acc of accounts) {
    const cs = await pageAll(`act_${acc.id}/campaigns?fields=id,name,status,effective_status,daily_budget,lifetime_budget,objective`);
    const objById: Record<string, any> = {};
    for (const c of cs) { const ob = metaObjetivo(c.objective); objById[c.id] = ob; campaigns.push({ id: c.id, nome: c.name, status: c.status, entrega: c.effective_status, orcamentoDiario: c.daily_budget ? +c.daily_budget / 100 : null, objetivo: ob, conta: acc.name || acc.id }); }
    const as = await pageAll(`act_${acc.id}/adsets?fields=id,name,status,effective_status,daily_budget,campaign_id`);
    for (const s of as) adsets.push({ id: s.id, nome: s.name, status: s.status, entrega: s.effective_status, orcamentoDiario: s.daily_budget ? +s.daily_budget / 100 : null, campanhaId: s.campaign_id, conta: acc.name || acc.id });
    if (!lightStatus) {
      const ds = await pageAll(`act_${acc.id}/ads?fields=id,name,status,effective_status,campaign_id,adset_id`);
      for (const d of ds) ads.push({ id: d.id, nome: d.name, status: d.status, entrega: d.effective_status, campanhaId: d.campaign_id, conjuntoId: d.adset_id, objetivo: objById[d.campaign_id] || metaObjetivo(""), conta: acc.name || acc.id });
    }
  }
  if (lightStatus) {
    const ids: string[] = m.adIds.filter(Boolean);
    for (let i = 0; i < ids.length; i += 50) {
      const chunk = ids.slice(i, i + 50);
      const r = await fetch(`${base}/?ids=${chunk.join(",")}&fields=id,status,effective_status&access_token=${token}`);
      const j = await r.json();
      if (j.error) continue;
      for (const id of chunk) if (j[id]) ads.push({ id, status: j[id].status, entrega: j[id].effective_status });
    }
  }
  return { campaigns, adsets, ads };
}

// Executa acoes de escrita no Meta (pausar/reativar/orcamento/duplicar). Requer escopo ads_management no token.
async function metaAction(m: any) {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const base = "https://graph.facebook.com/v21.0";
  const id = String(m.id || "");
  if (!id) throw new Error("id obrigatorio");
  async function post(path: string, params: Record<string, string>) {
    const bodyp = new URLSearchParams({ ...params, access_token: token });
    const r = await fetch(`${base}/${path}`, { method: "POST", body: bodyp });
    const j = await r.json();
    if (j.error) throw new Error(j.error.message);
    return j;
  }
  async function getField(objId: string, field: string) {
    const r = await fetch(`${base}/${objId}?fields=${field}&access_token=${token}`);
    const j = await r.json();
    if (j.error) throw new Error(j.error.message);
    return j;
  }
  if (m.action === "pause" || m.action === "activate") {
    const status = m.action === "pause" ? "PAUSED" : "ACTIVE";
    await post(id, { status });
    return { ok: true, detail: `${m.action === "pause" ? "Pausado" : "Reativado"}: ${m.nome || id}` };
  }
  if (m.action === "budget") {
    let cents: number;
    if (m.novoOrcamentoDiario != null) cents = Math.round(Number(m.novoOrcamentoDiario) * 100);
    else if (m.percentual != null) {
      const cur = await getField(id, "daily_budget");
      const curCents = Number(cur.daily_budget || 0);
      if (!curCents) throw new Error("Objeto sem orcamento diario (pode ser CBO no nivel da campanha ou orcamento vitalicio). Ajuste no nivel certo.");
      cents = Math.round(curCents * (1 + Number(m.percentual) / 100));
    } else throw new Error("Informe percentual ou novoOrcamentoDiario");
    if (cents < 100) throw new Error("Orcamento diario minimo ~R$1,00");
    await post(id, { daily_budget: String(cents) });
    return { ok: true, detail: `Orcamento diario ajustado para R$${(cents / 100).toFixed(2)}: ${m.nome || id}` };
  }
  if (m.action === "duplicate") {
    const j = await post(`${id}/copies`, { deep_copy: "true", status_option: "PAUSED" });
    return { ok: true, detail: `Campanha duplicada (copia PAUSADA): ${m.nome || id}`, copiedId: j.copied_campaign_id || j.id || null };
  }
  throw new Error("action invalida");
}

// Performance por segmentação (sexo / plataforma / posicionamento) — nível de conta, agregada entre contas
async function metaBreakdowns(m: any) {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const accounts = (Array.isArray(m.accounts) ? m.accounts : []).map((a: any) => String(a.id).replace(/^act_/, ""));
  if (!accounts.length) throw new Error("accounts obrigatorio");
  const base = "https://graph.facebook.com/v21.0";
  let range = "";
  if (m.since && m.until) range = `&time_range=${encodeURIComponent(JSON.stringify({ since: m.since, until: m.until }))}`;
  else range = `&date_preset=${m.datePreset || "last_30d"}`;
  const fields = "spend,impressions,clicks,actions,action_values";
  const pick = (arr: any[], types: string[]) => { if (!Array.isArray(arr)) return 0; for (const ty of types) { const hit = arr.find((x) => x.action_type === ty); if (hit) return parseFloat(hit.value || "0"); } return 0; };
  // posicionamento precisa vir pareado com publisher_platform (senão o Meta devolve vazio)
  const DIMS: Record<string, { bk: string; key: (row: any) => string }> = {
    sexo: { bk: "gender", key: (r) => String(r.gender || "desconhecido") },
    plataforma: { bk: "publisher_platform", key: (r) => String(r.publisher_platform || "desconhecido") },
    posicionamento: { bk: "publisher_platform,platform_position", key: (r) => `${r.publisher_platform || ""} · ${String(r.platform_position || "desconhecido").replace(/_/g, " ")}`.replace(/^ · /, "") },
  };
  const out: Record<string, Record<string, any>> = { sexo: {}, plataforma: {}, posicionamento: {} };
  await Promise.all(accounts.flatMap((acct: string) => Object.entries(DIMS).map(async ([dim, cfg]) => {
    try {
      let url: string | null = `${base}/act_${acct}/insights?level=account&fields=${fields}&breakdowns=${cfg.bk}${range}&limit=200&access_token=${token}`;
      for (let i = 0; i < 5 && url; i++) {
        const r = await fetch(url); const j = await r.json();
        if (j.error) break;
        for (const row of (j.data || [])) {
          const key = cfg.key(row);
          if (!out[dim][key]) out[dim][key] = { key, spend: 0, impressions: 0, clicks: 0, purchases: 0, revenue: 0, leads: 0, conversas: 0 };
          const o = out[dim][key];
          o.spend += parseFloat(row.spend || "0"); o.impressions += parseInt(row.impressions || "0"); o.clicks += parseInt(row.clicks || "0");
          o.purchases += pick(row.actions, ["omni_purchase", "offsite_conversion.fb_pixel_purchase", "purchase"]);
          o.revenue += pick(row.action_values, ["omni_purchase", "offsite_conversion.fb_pixel_purchase", "purchase"]);
          o.leads += pick(row.actions, ["offsite_conversion.fb_pixel_lead", "onsite_conversion.lead_grouped", "leadgen_grouped", "lead"]);
          o.conversas += pick(row.actions, ["onsite_conversion.messaging_conversation_started_7d", "messaging_conversation_started_7d"]);
        }
        url = j.paging?.next || null;
      }
    } catch (_e) { /* dim indisponível: segue com as outras */ }
  })));
  const sorted = (o: Record<string, any>) => Object.values(o).sort((a: any, b: any) => b.spend - a.spend);
  return { sexo: sorted(out.sexo), plataforma: sorted(out.plataforma), posicionamento: sorted(out.posicionamento) };
}

/* ================= GOOGLE ADS ================= */
const GADS_VER = "v23";

async function googleAdsAccessToken() {
  const clientId = Deno.env.get("GOOGLE_ADS_CLIENT_ID");
  const clientSecret = Deno.env.get("GOOGLE_ADS_CLIENT_SECRET");
  const refreshToken = Deno.env.get("GOOGLE_ADS_REFRESH_TOKEN");
  if (!clientId || !clientSecret || !refreshToken) throw new Error("Credenciais do Google Ads nao configuradas nos secrets (GOOGLE_ADS_CLIENT_ID/SECRET/REFRESH_TOKEN)");
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken, grant_type: "refresh_token" }),
  });
  const j = await r.json();
  if (!r.ok || !j.access_token) throw new Error("Google Ads OAuth: " + (j.error_description || j.error || "falha ao renovar token"));
  return j.access_token as string;
}

// searchStream paginado (search comum) na conta cid
async function gadsSearch(cid: string, query: string, accessToken: string) {
  const devToken = Deno.env.get("GOOGLE_ADS_DEV_TOKEN");
  const mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  if (!devToken) throw new Error("GOOGLE_ADS_DEV_TOKEN nao configurada nos secrets");
  const out: any[] = [];
  let pageToken: string | undefined;
  for (let i = 0; i < 20; i++) {
    const r = await fetch(`https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/googleAds:search`, {
      method: "POST",
      headers: { "Authorization": `Bearer ${accessToken}`, "developer-token": devToken, "login-customer-id": mcc, "Content-Type": "application/json" },
      body: JSON.stringify({ query, ...(pageToken ? { pageToken } : {}) }),
    });
    const j = await r.json();
    if (!r.ok) {
      const msg = j?.error?.details?.[0]?.errors?.[0]?.message || j?.error?.message || `HTTP ${r.status}`;
      throw new Error(msg);
    }
    out.push(...(j.results || []));
    pageToken = j.nextPageToken;
    if (!pageToken) break;
  }
  return out;
}

// Lista as contas-cliente sob a MCC
async function googleListAccounts() {
  const mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  if (!mcc) throw new Error("GOOGLE_ADS_MCC_ID nao configurada nos secrets");
  const token = await googleAdsAccessToken();
  const rows = await gadsSearch(mcc, `SELECT customer_client.id, customer_client.descriptive_name, customer_client.status, customer_client.manager, customer_client.currency_code FROM customer_client WHERE customer_client.hidden = FALSE`, token);
  return rows
    .map((r: any) => r.customerClient)
    .filter((c: any) => c && !c.manager)
    .map((c: any) => ({ id: String(c.id), name: c.descriptiveName || String(c.id), status: c.status, currency: c.currencyCode }));
}

// Objetivo "equivalente" pelo tipo de canal da campanha (pra aba Campanhas avaliar pela metrica certa)
function googleObjetivo(channelType: string) {
  const t = String(channelType || "").toUpperCase();
  const map: Record<string, { tipo: string; rotulo: string; metrica: string }> = {
    SEARCH: { tipo: "conversao", rotulo: "Google · Pesquisa", metrica: "ROAS, CPA, conversões" },
    PERFORMANCE_MAX: { tipo: "conversao", rotulo: "Google · Performance Max", metrica: "ROAS, CPA, conversões" },
    SHOPPING: { tipo: "conversao", rotulo: "Google · Shopping", metrica: "ROAS, CPA, conversões" },
    DISPLAY: { tipo: "alcance", rotulo: "Google · Display", metrica: "CPM, alcance, cliques" },
    VIDEO: { tipo: "video", rotulo: "Google · Vídeo (YouTube)", metrica: "custo por view, CPM" },
    DEMAND_GEN: { tipo: "engajamento", rotulo: "Google · Demand Gen", metrica: "CTR, custo por clique/engajamento" },
    DISCOVERY: { tipo: "engajamento", rotulo: "Google · Discovery", metrica: "CTR, CPC" },
    LOCAL: { tipo: "conversao", rotulo: "Google · Local", metrica: "conversões, CPA" },
    HOTEL: { tipo: "conversao", rotulo: "Google · Hotel", metrica: "ROAS, CPA" },
  };
  return { codigo: t || null, ...(map[t] || { tipo: "conversao", rotulo: "Google Ads", metrica: "conversões, CPA, ROAS" }) };
}

// v23 renomeou video_views -> video_trueview_views (e customer não aceita métricas de vídeo/engajamento)
const GADS_METRICS = "metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.conversions, metrics.conversions_value";
const GADS_METRICS_FULL = GADS_METRICS + ", metrics.video_trueview_views, metrics.engagements";
function gadsShape(m: any) {
  const spend = (Number(m?.costMicros) || 0) / 1e6;
  const impressions = Number(m?.impressions) || 0;
  const clicks = Number(m?.clicks) || 0;
  const purchases = Number(m?.conversions) || 0;
  const revenue = Number(m?.conversionsValue) || 0;
  return {
    spend, impressions, clicks,
    ctr: impressions ? (clicks / impressions) * 100 : 0,
    cpc: clicks ? spend / clicks : 0,
    cpm: impressions ? (spend / impressions) * 1000 : 0,
    reach: 0, frequency: 0,
    purchases, revenue, roas: spend ? revenue / spend : 0,
    leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0,
    videoViews: Number(m?.videoTrueviewViews ?? m?.videoViews) || 0,
    engajamentos: Number(m?.engagements) || 0,
  };
}

// Insights do Google Ads no MESMO formato do metaAdsInsights ({total, campaigns, ads, accounts, accountErrors, period})
async function googleAdsInsights(g: any) {
  let accounts: { id: string; name: string }[] = [];
  if (Array.isArray(g.accounts) && g.accounts.length) accounts = g.accounts.map((a: any) => ({ id: String(a.id).replace(/-/g, ""), name: a.name || "" }));
  else if (Array.isArray(g.accountIds) && g.accountIds.length) accounts = g.accountIds.map((id: any) => ({ id: String(id).replace(/-/g, ""), name: "" }));
  if (!accounts.length) throw new Error("accountId(s) do Google obrigatorio");
  const since = g.since, until = g.until;
  if (!since || !until) throw new Error("since e until obrigatorios (YYYY-MM-DD)");
  const range = `segments.date BETWEEN '${String(since).slice(0, 10)}' AND '${String(until).slice(0, 10)}'`;
  const token = await googleAdsAccessToken();

  const totAgg: any = { spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0, videoViews: 0, engajamentos: 0 };
  const totRecByDate: Record<string, any> = {};
  const byCamp: Record<string, any> = {};
  const ads: any[] = [];

  const perAccount = await Promise.all(accounts.map(async (acc) => {
    try {
      const [accountRows, acctDaily, campRows, adRows] = await Promise.all([
        gadsSearch(acc.id, `SELECT ${GADS_METRICS} FROM customer WHERE ${range}`, token),
        g.daily ? gadsSearch(acc.id, `SELECT segments.date, ${GADS_METRICS} FROM customer WHERE ${range}`, token) : Promise.resolve([] as any[]),
        g.byCampaign ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, campaign.advertising_channel_type${g.daily ? ", segments.date" : ""}, ${GADS_METRICS_FULL} FROM campaign WHERE ${range}`, token) : Promise.resolve([] as any[]),
        g.byAd ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, campaign.advertising_channel_type, ad_group.id, ad_group.name, ad_group_ad.ad.id, ad_group_ad.ad.name, ad_group_ad.ad.type, ${GADS_METRICS_FULL} FROM ad_group_ad WHERE ${range} AND metrics.cost_micros > 0`, token) : Promise.resolve([] as any[]),
      ]);
      return { acc, accountRows, acctDaily, campRows, adRows, error: null as string | null };
    } catch (e) {
      return { acc, accountRows: [] as any[], acctDaily: [] as any[], campRows: [] as any[], adRows: [] as any[], error: (e as any)?.message || String(e) };
    }
  }));
  const accountErrors = perAccount.filter((p) => p.error).map((p) => ({ id: p.acc.id, name: p.acc.name || p.acc.id, error: p.error }));

  for (const { acc, accountRows, acctDaily, campRows, adRows } of perAccount) {
    for (const row of accountRows) {
      const s = gadsShape(row.metrics);
      totAgg.spend += s.spend; totAgg.impressions += s.impressions; totAgg.clicks += s.clicks;
      totAgg.revenue += s.revenue; totAgg.purchases += s.purchases; totAgg.videoViews += s.videoViews; totAgg.engajamentos += s.engajamentos;
    }
    for (const row of acctDaily) {
      const s = gadsShape(row.metrics); const k = row.segments?.date;
      if (!k) continue;
      if (!totRecByDate[k]) totRecByDate[k] = { date: k, sales: 0, spend: 0, revenue: 0, clicks: 0, impressions: 0, reach: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0, addToCart: 0, checkout: 0 };
      const rec = totRecByDate[k];
      rec.sales += Math.round(s.purchases); rec.spend += s.spend; rec.revenue += s.revenue; rec.clicks += s.clicks; rec.impressions += s.impressions; rec.videoViews += s.videoViews; rec.engajamentos += s.engajamentos;
    }
    for (const row of campRows) {
      const label = row.campaign?.name || "Google Ads";
      const s = gadsShape(row.metrics);
      if (!byCamp[label]) byCamp[label] = { campaign: label, campaignId: row.campaign?.id ? String(row.campaign.id) : null, account: acc.name || acc.id, objetivo: googleObjetivo(row.campaign?.advertisingChannelType), _google: true, spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, records: [] };
      const c = byCamp[label];
      c.spend += s.spend; c.impressions += s.impressions; c.clicks += s.clicks;
      c.revenue += s.revenue; c.purchases += s.purchases;
      if (g.daily && row.segments?.date) c.records.push({ date: row.segments.date, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: 0, leads: 0, conversas: 0, videoViews: s.videoViews, engajamentos: s.engajamentos });
    }
    for (const row of adRows) {
      const s = gadsShape(row.metrics);
      const ad = row.adGroupAd?.ad || {};
      const adName = ad.name || (ad.type ? String(ad.type).replace(/_/g, " ").toLowerCase() : "anúncio") + " #" + (ad.id || "");
      ads.push({
        adId: ad.id ? "g" + ad.id : null, adName, campaign: row.campaign?.name || "", campaignId: row.campaign?.id ? String(row.campaign.id) : null,
        adset: row.adGroup?.name || "", adsetId: row.adGroup?.id ? String(row.adGroup.id) : null,
        account: acc.name || acc.id, thumbnail: null, _google: true,
        objetivo: googleObjetivo(row.campaign?.advertisingChannelType),
        spend: s.spend, impressions: s.impressions, clicks: s.clicks, reach: 0, frequency: 0,
        ctr: s.ctr, cpc: s.cpc, cpm: s.cpm, purchases: s.purchases, revenue: s.revenue, roas: s.roas,
        leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0, videoViews: s.videoViews, engajamentos: s.engajamentos,
        cpa: s.purchases ? s.spend / s.purchases : 0,
      });
    }
  }
  const total = {
    ...totAgg,
    ctr: totAgg.impressions ? (totAgg.clicks / totAgg.impressions) * 100 : 0,
    cpc: totAgg.clicks ? totAgg.spend / totAgg.clicks : 0,
    cpm: totAgg.impressions ? (totAgg.spend / totAgg.impressions) * 1000 : 0,
    roas: totAgg.spend ? totAgg.revenue / totAgg.spend : 0,
    records: Object.values(totRecByDate).sort((a: any, b: any) => a.date < b.date ? -1 : 1),
  };
  const campaigns = Object.values(byCamp).map((c: any) => {
    c.ctr = c.impressions ? (c.clicks / c.impressions) * 100 : 0;
    c.cpc = c.clicks ? c.spend / c.clicks : 0;
    c.cpm = c.impressions ? (c.spend / c.impressions) * 1000 : 0;
    c.roas = c.spend ? c.revenue / c.spend : 0;
    return c;
  }).sort((a: any, b: any) => b.spend - a.spend);
  ads.sort((a: any, b: any) => b.spend - a.spend);
  return { total, campaigns, ads, accounts, accountErrors, period: { since, until } };
}

// Detalhes específicos do Google: conversões por ação, palavras-chave e termos de busca (agregados entre contas)
async function googleBreakdowns(g: any) {
  let accounts: string[] = [];
  if (Array.isArray(g.accounts) && g.accounts.length) accounts = g.accounts.map((a: any) => String(a.id).replace(/-/g, ""));
  if (!accounts.length) throw new Error("accounts obrigatorio");
  const since = String(g.since || "").slice(0, 10), until = String(g.until || "").slice(0, 10);
  if (!since || !until) throw new Error("since e until obrigatorios");
  const range = `segments.date BETWEEN '${since}' AND '${until}'`;
  const token = await googleAdsAccessToken();
  const M = "metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.conversions, metrics.conversions_value";
  const conv: Record<string, any> = {}, kw: Record<string, any> = {}, st: Record<string, any> = {};
  const addRow = (map: Record<string, any>, key: string, m: any) => {
    if (!map[key]) map[key] = { key, spend: 0, impressions: 0, clicks: 0, conversions: 0, value: 0 };
    const o = map[key];
    o.spend += (Number(m?.costMicros) || 0) / 1e6; o.impressions += Number(m?.impressions) || 0; o.clicks += Number(m?.clicks) || 0;
    o.conversions += Number(m?.conversions) || 0; o.value += Number(m?.conversionsValue) || 0;
  };
  const errors: string[] = [];
  await Promise.all(accounts.flatMap((cid) => [
    gadsSearch(cid, `SELECT segments.conversion_action_name, metrics.conversions, metrics.conversions_value FROM campaign WHERE ${range} AND metrics.conversions > 0`, token)
      .then((rows) => rows.forEach((r: any) => addRow(conv, r.segments?.conversionActionName || "—", r.metrics)))
      .catch((e) => errors.push("conversões: " + e.message)),
    gadsSearch(cid, `SELECT ad_group_criterion.keyword.text, ${M} FROM keyword_view WHERE ${range} AND metrics.impressions > 0 ORDER BY metrics.cost_micros DESC LIMIT 200`, token)
      .then((rows) => rows.forEach((r: any) => addRow(kw, r.adGroupCriterion?.keyword?.text || "—", r.metrics)))
      .catch((e) => errors.push("keywords: " + e.message)),
    gadsSearch(cid, `SELECT search_term_view.search_term, ${M} FROM search_term_view WHERE ${range} ORDER BY metrics.cost_micros DESC LIMIT 200`, token)
      .then((rows) => rows.forEach((r: any) => addRow(st, r.searchTermView?.searchTerm || "—", r.metrics)))
      .catch((e) => errors.push("termos: " + e.message)),
  ]));
  const sorted = (o: Record<string, any>, by = "spend") => Object.values(o).sort((a: any, b: any) => b[by] - a[by]).slice(0, 100);
  return { conversoes: sorted(conv, "conversions"), keywords: sorted(kw), termos: sorted(st), errors: errors.length ? errors : undefined };
}

/* ================= RD STATION ================= */
const _SB_URL = Deno.env.get("SUPABASE_URL") || "";
const _SB_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
async function sbGet(table: string, query: string) {
  const r = await fetch(`${_SB_URL}/rest/v1/${table}?${query}`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
  return r.ok ? await r.json() : [];
}
// Token de acesso do RD daquele cliente (refresh_token do cliente + credenciais do App na conta)
async function rdAccessToken(clientId: string) {
  const acc = await sbGet("account_config", "id=eq.main&select=data");
  const app = (acc[0]?.data || {}).rd_station || {};
  if (!app.client_id || !app.client_secret) throw new Error("Credenciais do App do RD não configuradas (aba Configurações).");
  const cli = await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=rd_config`);
  const rt = (cli[0]?.rd_config || {}).refresh_token;
  if (!rt) throw new Error("Este cliente ainda não conectou o RD Station.");
  const r = await fetch("https://api.rd.services/auth/token", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ client_id: app.client_id, client_secret: app.client_secret, refresh_token: rt }),
  });
  const j = await r.json();
  if (!j.access_token) throw new Error("Falha ao renovar token do RD: " + (j.error_description || j.error || r.status));
  return j.access_token as string;
}
// Catálogo de eventos/conversões do RD (landing pages, popups etc. com contagem no período)
async function rdCatalog(m: any) {
  const clientId = String(m.clientId || "");
  if (!clientId) throw new Error("clientId obrigatório");
  const at = await rdAccessToken(clientId);
  const until = m.until || new Date().toISOString().slice(0, 10);
  const since = m.since || new Date(Date.now() - 730 * 864e5).toISOString().slice(0, 10); // 2 anos: pega o catálogo todo
  const r = await fetch(`https://api.rd.services/platform/analytics/conversions?start_date=${since}&end_date=${until}`, { headers: { Authorization: `Bearer ${at}` } });
  const j = await r.json();
  if (!r.ok) throw new Error("RD: " + (j.error_description || j.error || `HTTP ${r.status}`));
  const list = (j.conversions || []).map((x: any) => ({
    identifier: x.asset_identifier, type: x.assets_type,
    conversions: Number(x.conversion_count) || 0, visits: Number(x.visits_count) || 0,
  })).filter((x: any) => x.identifier);
  list.sort((a: any, b: any) => b.conversions - a.conversions);
  return { period: { since, until }, events: list };
}

// Nuvemshop · pedidos do cliente no período → agregado diário (data, contagem, faturamento) por status.
async function nuvemshopOrders(m: any) {
  const clientId = String(m.clientId || "");
  if (!clientId) throw new Error("clientId obrigatório");
  const cli = await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=nuvemshop_config`);
  const cfg = (cli[0]?.nuvemshop_config || {});
  if (!cfg.access_token || !cfg.store_id) throw new Error("Este cliente não conectou a Nuvemshop.");
  const acc = await sbGet("account_config", "id=eq.main&select=data");
  const ua = ((acc[0]?.data || {}).nuvemshop || {}).ua || "Central de Gestao (contato@gtmarketing.com.br)";
  const since = m.since || new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10);
  const until = m.until || new Date().toISOString().slice(0, 10);
  const base = `https://api.tiendanube.com/2025-03/${cfg.store_id}`;
  const H = { "Authentication": `bearer ${cfg.access_token}`, "User-Agent": ua, "Content-Type": "application/json" };
  const agg: Record<string, any> = {};
  let totalCount = 0, paidCount = 0, paidRevenue = 0;
  let page = 1;
  for (let i = 0; i < 20; i++) {
    const u = `${base}/orders?created_at_min=${since}T00:00:00-03:00&created_at_max=${until}T23:59:59-03:00&per_page=200&page=${page}&fields=id,total,created_at,completed_at,payment_status,status`;
    const r = await fetch(u, { headers: H });
    if (!r.ok) { const t = await r.text(); throw new Error(`Nuvemshop: HTTP ${r.status} ${t.slice(0, 120)}`); }
    const rows = await r.json();
    if (!Array.isArray(rows) || !rows.length) break;
    for (const o of rows) {
      const date = String(o.completed_at || o.created_at || "").slice(0, 10);
      const paid = o.payment_status === "paid" || o.payment_status === "authorized";
      const status = paid ? "Aprovado" : (o.payment_status === "voided" || o.status === "cancelled" ? "Cancelado" : "Aguardando");
      totalCount++;
      const key = date + "|" + status;
      if (!agg[key]) agg[key] = { date, status, count: 0, total: 0 };
      agg[key].count++; agg[key].total += parseFloat(o.total || "0");
      if (paid) { paidCount++; paidRevenue += parseFloat(o.total || "0"); }
    }
    if (rows.length < 200) break;
    page++;
  }
  return { rows: Object.values(agg), paidCount, paidRevenue, totalCount, period: { since, until } };
}

// Raio-X · leitura da IA: resumo executivo + por que cada criativo ganha/perde (Vision) + o que separa os lados.
async function raioxAI(m: any) {
  const sys = `Você é uma gestora de tráfego senior (nível Pedro Sobral) analisando o Raio-X de um cliente. Recebe: KPIs do período, o objetivo/modelo, o health score por dimensão, e os criativos vencedores e perdedores (com IMAGEM de cada um e a métrica do objetivo). Escreva uma análise afiada, específica e acionável, em português.
Responda SOMENTE um JSON válido, sem markdown, no formato:
{"resumo":"2-4 frases: diagnóstico do período (o que foi bem/mal, e por quê), citando números reais","separa":"1-3 frases: o que separa os vencedores dos perdedores (padrão de hook/oferta/estética/CTA)","ads":[{"adId":"<id>","why":"1 frase curta: por que ESTE criativo ganha ou perde — olhe a imagem (hook, oferta, estética, clareza, CTA)"}],"proximos_passos":["2-4 ações concretas priorizadas"]}
Regras: baseie-se nos números e nas imagens; nunca invente dados; avalie cada criativo pela métrica do objetivo dele.`;
  const content: any[] = [{ type: "text", text: `Cliente: ${m.clientName}\nPeríodo: ${m.periodo}\nModelo/objetivo: ${m.objetivo}\nKPIs: ${JSON.stringify(m.kpis)}\nHealth score: ${m.score}/100 (${m.classificacao}) — dimensões: ${JSON.stringify(m.dims)}\n${m.dna ? "DNA do cliente: " + JSON.stringify(m.dna).slice(0, 2000) : ""}\n\nCriativos abaixo (imagem + dados):` }];
  for (const a of (m.ads || []).slice(0, 10)) {
    if (a.thumbnail) content.push({ type: "image_url", image_url: { url: String(a.thumbnail) } });
    content.push({ type: "text", text: `^ adId=${a.adId} · ${a.adName} · lado=${a.lado} · ${a.metric}=${a.valor} · invest=${a.spend}` });
  }
  const json = await callOpenAI({ model: "gpt-4o", messages: [{ role: "system", content: sys }, { role: "user", content }], response_format: { type: "json_object" }, max_tokens: 1800, temperature: 0.5 });
  return JSON.parse(json.choices?.[0]?.message?.content || "{}");
}

// Análise de site: PageSpeed Insights (carregamento/acessibilidade/SEO) + leitura de UX/navegação por IA.
async function siteAudit(m: any) {
  const url = String(m.url || "").trim();
  if (!url) throw new Error("url obrigatória");
  const key = Deno.env.get("GOOGLE_PSI_KEY") || "";
  let psi: any = null, psiErr: string | null = null;
  try {
    const u = `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encodeURIComponent(url)}&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo${key ? `&key=${key}` : ""}`;
    const r = await fetch(u);
    const j = await r.json();
    if (j.error) psiErr = j.error.message || `HTTP ${r.status}`;
    else {
      const cat = j.lighthouseResult?.categories || {};
      const au = j.lighthouseResult?.audits || {};
      const pct = (x: any) => x?.score != null ? Math.round(x.score * 100) : null;
      psi = {
        scores: { performance: pct(cat.performance), accessibility: pct(cat.accessibility), bestPractices: pct(cat["best-practices"]), seo: pct(cat.seo) },
        metrics: {
          lcp: au["largest-contentful-paint"]?.displayValue, cls: au["cumulative-layout-shift"]?.displayValue,
          tbt: au["total-blocking-time"]?.displayValue, fcp: au["first-contentful-paint"]?.displayValue, si: au["speed-index"]?.displayValue,
        },
        opportunities: Object.values(au).filter((a: any) => a?.details?.type === "opportunity" && a.score != null && a.score < 0.9)
          .sort((a: any, b: any) => (b.details?.overallSavingsMs || 0) - (a.details?.overallSavingsMs || 0))
          .slice(0, 6).map((a: any) => ({ title: a.title, savingsMs: Math.round(a.details?.overallSavingsMs || 0) })),
      };
    }
  } catch (e) { psiErr = (e as any)?.message || String(e); }
  let ai: any = null;
  try {
    const text = await fetchUrlText(url);
    const sys = `Você é especialista em CRO/UX e otimização de conversão. A partir do conteúdo/estrutura da página, avalie: usabilidade, navegação, clareza da oferta, força do CTA, prova social, confiança e mobile. Responda SOMENTE JSON: {"resumo":"2-3 frases","pontos_fortes":["..."],"melhorias":["ações concretas priorizadas"]}. Português, sem inventar o que não está na página.`;
    const j = await callOpenAI({ messages: [{ role: "system", content: sys }, { role: "user", content: `URL: ${url}\n\nConteúdo:\n${text.slice(0, 14000)}` }], response_format: { type: "json_object" }, max_tokens: 900, temperature: 0.5 });
    ai = JSON.parse(j.choices?.[0]?.message?.content || "{}");
  } catch (_e) { /* sem IA: segue só com PSI */ }
  return { url, psi, psiErr, ai };
}

/* ===== WhatsApp (uazapi) — polling, envio, status ===== */
async function sbPost(table: string, row: Record<string, unknown>) {
  await fetch(`${_SB_URL}/rest/v1/${table}`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" }, body: JSON.stringify(row) });
}
// insert que CHECA o resultado (pra confirmar de verdade que gravou)
async function sbInsertOk(table: string, row: Record<string, unknown>): Promise<{ ok: boolean; err: string }> {
  const r = await fetch(`${_SB_URL}/rest/v1/${table}`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" }, body: JSON.stringify(row) });
  if (r.ok) return { ok: true, err: "" };
  const t = await r.text().catch(() => "");
  return { ok: false, err: (t || `HTTP ${r.status}`).slice(0, 160) };
}
async function sbPatchD(table: string, query: string, row: Record<string, unknown>) {
  await fetch(`${_SB_URL}/rest/v1/${table}?${query}`, { method: "PATCH", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" }, body: JSON.stringify(row) });
}
function _wuid() { return crypto.randomUUID().replace(/-/g, "").slice(0, 20); }
function waTs(v: any): string { const n = Number(v) || 0; if (!n) return new Date().toISOString(); return new Date(n > 1e12 ? n : n * 1000).toISOString(); }
function waText(m: any): string { const c = m.content || {}; return m.text || c.text || c.conversation || c.extendedTextMessage?.text || c.imageMessage?.caption || c.videoMessage?.caption || ""; }
function waOrigin(m: any): { type: string; data: Record<string, unknown> } | null {
  const c = m.content || {}; const ci = c.contextInfo || c.extendedTextMessage?.contextInfo || m.contextInfo || {};
  const ad = ci.externalAdReply || c.externalAdReply || null;
  if (ad && (ad.sourceId || ad.sourceUrl || ad.ctwaClid || ad.title)) return { type: "anuncio", data: { source_id: ad.sourceId || "", source_type: ad.sourceType || "", source_url: ad.sourceUrl || "", ctwa_clid: ad.ctwaClid || ci.ctwaClid || "", title: ad.title || "", body: ad.body || "", thumbnail: ad.thumbnailUrl || ad.thumbnail || "" } };
  if (m.track_source || m.track_id) return { type: m.track_source === "ad" ? "anuncio" : "utm", data: { track_source: m.track_source || "", track_id: m.track_id || "" } };
  return null;
}
async function waCall(host: string, token: string, path: string, method = "GET", payload?: any) {
  const r = await fetch(host.replace(/\/$/, "") + path, { method, headers: { token, "Content-Type": "application/json" }, body: payload ? JSON.stringify(payload) : undefined });
  const t = await r.text(); let j: any; try { j = JSON.parse(t); } catch { j = t; }
  return { status: r.status, j };
}
// Resolve o ad_id do CTWA (source_id da conversa) em nomes: campanha › conjunto › anúncio (Graph API)
async function waResolveAd(adId: string): Promise<Record<string, string> | null> {
  const token = Deno.env.get("META_USER_TOKEN"); if (!token || !adId) return null;
  try {
    const r = await fetch(`https://graph.facebook.com/v21.0/${encodeURIComponent(adId)}?fields=name,adset{name},campaign{name}&access_token=${token}`);
    const j = await r.json(); if (!j || j.error) return null;
    return { ad: j.name || "", adset: (j.adset && j.adset.name) || "", campaign: (j.campaign && j.campaign.name) || "" };
  } catch { return null; }
}
async function waUzConfig() { const acc = await sbGet("account_config", "id=eq.main&select=data"); const uz = (acc[0]?.data || {}).uazapi || {}; if (!uz.server || !uz.admin_token) throw new Error("uazapi não configurado (aba Configurações → WhatsApp)."); return uz; }
const CRM_DEFAULT_FIELDS = [
  { key: "nome", label: "Nome", type: "texto", hint: "Nome próprio que o lead usou ao se apresentar" },
  { key: "email", label: "Email", type: "texto", hint: "" },
  { key: "produto", label: "Produto/Serviço de Interesse", type: "texto", hint: "O que o lead quer comprar ou contratar" },
  { key: "valor", label: "Valor", type: "valor", hint: "Valor em R$ mencionado na negociação" },
];
const CRM_DEFAULT_STAGES = [
  { key: "sem", label: "Sem etapa", desc: "" }, { key: "novo", label: "Novo", desc: "Contato acabou de chegar; só a primeira mensagem." },
  { key: "mql", label: "MQL", desc: "Demonstrou interesse inicial, mas ainda sem informações suficientes pra ser oportunidade forte." },
  { key: "sql", label: "SQL", desc: "Pediu algo específico / qualificado / com intenção clara." }, { key: "comprou", label: "Comprou", desc: "Pagamento confirmado / fechou." },
  { key: "posvenda", label: "Pós-Venda", desc: "Já é cliente; comunicação pós-compra." }, { key: "perdido", label: "Perdido", desc: "Desistiu / sem interesse." },
];
// IA lê a conversa: extrai os campos configurados + CLASSIFICA a etapa do funil com um nível de confiança.
// autoApply: aplica a etapa automaticamente se a confiança >= mínimo configurado.
async function waExtract(convId: string, autoApply = false) {
  const cv = (await sbGet("wa_conversations", `id=eq.${encodeURIComponent(convId)}&select=id,name,fields,stage&limit=1`))[0];
  if (!cv) throw new Error("Conversa não encontrada.");
  const msgs = await sbGet("wa_messages", `conversation_id=eq.${encodeURIComponent(convId)}&order=ts.asc&select=direction,text&limit=200`);
  const transcript = (msgs || []).filter((m: any) => m.text).map((m: any) => `${m.direction === "in" ? "LEAD" : "ATENDENTE"}: ${m.text}`).join("\n").slice(0, 6000);
  if (!transcript) return { fields: cv.fields || {}, stage: "", confidence: 0, stageWhy: "", applied: false };
  const data = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  const fields = data.crm_fields || CRM_DEFAULT_FIELDS;
  const stages = (Array.isArray(data.crm_stages) && data.crm_stages.length) ? data.crm_stages : CRM_DEFAULT_STAGES;
  const minConf = Number(data.crm_min_confidence != null ? data.crm_min_confidence : 70);
  const spec = fields.map((f: any) => `- ${f.key} (${f.label}${f.type ? ", tipo " + f.type : ""})${f.hint ? ": " + f.hint : ""}`).join("\n");
  const stageSpec = stages.filter((s: any) => s.key !== "sem").map((s: any) => `- ${s.key} = ${s.label}${s.desc ? ": " + s.desc : ""}`).join("\n");
  const keys = stages.map((s: any) => s.key).join(", ");
  const sys = "Você é um SDR que lê uma conversa de WhatsApp entre o LEAD e o ATENDENTE. Faça duas coisas: (1) extraia os campos do lead — só o que aparece claramente, NÃO invente; para tipo 'valor' devolva só o número; (2) CLASSIFIQUE a etapa do funil usando as descrições dadas, e dê um 'confidence' de 0 a 100 (quão certo você está). Responda SOMENTE JSON.";
  const content = `CAMPOS A EXTRAIR:\n${spec}\n\nETAPAS DO FUNIL (escolha UMA key):\n${stageSpec}\n\nCONVERSA:\n${transcript}\n\nResponda JSON: {"fields":{"<key>":"<valor>"}, "stage":"<key entre: ${keys}>", "confidence":<0-100>, "stageWhy":"<motivo curto>"}`;
  const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content }], response_format: { type: "json_object" }, max_tokens: 700, temperature: 0.2 });
  let parsed: any = {}; try { parsed = JSON.parse(j.choices[0].message.content || "{}"); } catch { parsed = {}; }
  const outFields = { ...(cv.fields || {}), ...(parsed.fields || {}) };
  const stage = (parsed.stage && stages.some((s: any) => s.key === parsed.stage)) ? parsed.stage : "";
  const conf = Math.max(0, Math.min(100, Math.round(Number(parsed.confidence) || 0)));
  const patch: Record<string, unknown> = { fields: outFields, ai_stage: stage, ai_conf: conf, ai_why: parsed.stageWhy || "", ai_at: new Date().toISOString() };
  const applied = !!(autoApply && stage && conf >= minConf && stage !== cv.stage);
  if (applied) patch.stage = stage;
  await sbPatchD("wa_conversations", `id=eq.${encodeURIComponent(convId)}`, patch);
  if (applied) { const stObj = stages.find((s: any) => s.key === stage); if (stObj && stObj.event) { try { await waCapi(convId, stObj.event); } catch (_e) {} } }
  return { fields: outFields, stage, confidence: conf, stageWhy: parsed.stageWhy || "", applied, minConf };
}
// ---- CAPI (Conversions API): manda o evento da etapa pro Meta, atribuindo ao anúncio via ctwa_clid ----
const _pixelCache: Record<string, string | null> = {};
async function _sha256hex(s: string) { const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)); return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join(""); }
async function clientPixelId(clientId: string): Promise<string | null> {
  if (_pixelCache[clientId] !== undefined) return _pixelCache[clientId];
  const token = Deno.env.get("META_USER_TOKEN");
  const cli = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=meta_account_id`))[0];
  const acct = String(cli?.meta_account_id || "").split(",")[0].replace(/^act_/, "").trim();
  let pid: string | null = null;
  if (acct && token) { try { const r = await fetch(`https://graph.facebook.com/v21.0/act_${acct}/adspixels?fields=id&limit=1&access_token=${token}`); const j = await r.json(); pid = (j.data && j.data[0] && j.data[0].id) || null; } catch { pid = null; } }
  _pixelCache[clientId] = pid; return pid;
}
async function waCapi(convId: string, eventName: string) {
  const cv = (await sbGet("wa_conversations", `id=eq.${encodeURIComponent(convId)}&select=id,client_id,chat_id,origin`))[0];
  if (!cv) throw new Error("Conversa não encontrada.");
  const token = Deno.env.get("META_USER_TOKEN");
  const pid = cv.client_id ? await clientPixelId(cv.client_id) : null;
  const logId = _wuid();
  if (!pid || !token) { await sbPost("capi_events", { id: logId, client_id: cv.client_id, conversation_id: convId, event_name: eventName, status: "failed", error: "Pixel do cliente ou token do Meta ausente." }); return { ok: false, error: "pixel/token" }; }
  const o = cv.origin || {}; const phone = String(cv.chat_id || "").replace(/[^0-9]/g, "");
  const user_data: any = {}; if (phone) user_data.ph = await _sha256hex(phone); if (o.ctwa_clid) user_data.ctwa_clid = o.ctwa_clid;
  const ev: any = { event_name: eventName, event_time: Math.floor(Date.now() / 1000), action_source: o.ctwa_clid ? "business_messaging" : "website", user_data };
  if (o.ctwa_clid) ev.messaging_channel = "whatsapp";
  let status = "success", error = "", resp: any = null;
  try {
    const r = await fetch(`https://graph.facebook.com/v21.0/${pid}/events?access_token=${token}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ data: [ev] }) });
    resp = await r.json(); if (!r.ok || resp.error) { status = "failed"; error = (resp.error && resp.error.message) || ("HTTP " + r.status); }
  } catch (e) { status = "failed"; error = String(e); }
  await sbPost("capi_events", { id: logId, client_id: cv.client_id, conversation_id: convId, event_name: eventName, status, error, response: resp });
  return { ok: status === "success", status, error };
}
// ===== AndréIA no WhatsApp (grupo): entende a mensagem, analisa o cliente, cria tarefa (com confirmação) =====
function _waResumoMeta(t: any) { return { gasto: Math.round(t.spend), impressoes: t.impressions, cliques: t.clicks, ctr: +(t.ctr || 0).toFixed(2), cpc: +(t.cpc || 0).toFixed(2), cpm: +(t.cpm || 0).toFixed(2), alcance: t.reach, conversas: t.conversas, custoPorConversa: t.conversas ? +(t.spend / t.conversas).toFixed(2) : null, leads: t.leads, compras: Math.round(t.purchases || 0), roas: +(t.roas || 0).toFixed(2) }; }
const _snapCache: Record<string, { t: number; v: any }> = {};
async function waAgentSnapshot(clientId: string) {
  const cached = _snapCache[clientId]; if (cached && Date.now() - cached.t < 180000) return cached.v;
  const v = await _waBuildSnapshot(clientId);
  if (v) _snapCache[clientId] = { t: Date.now(), v };
  return v;
}
async function _waBuildSnapshot(clientId: string) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,benchmark_metas,meta_account_id,conversion_source`))[0];
  if (!c) return null;
  const out: any = { nome: c.name };
  const ids = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) { out.aviso = "Cliente sem conta Meta vinculada no sistema."; return out; }
  const accounts = ids.map((id: string) => ({ id, name: id }));
  const iso = (d: number) => new Date(Date.now() - d * 864e5).toISOString().slice(0, 10);
  const hoje = new Date().toISOString().slice(0, 10);
  const [r7, r30, ent] = await Promise.all([
    metaAdsInsights({ accounts, since: iso(7), until: hoje }).catch(() => null),
    metaAdsInsights({ accounts, since: iso(30), until: hoje, byCampaign: true, byAd: true }).catch(() => null),
    metaEntities({ accounts }).catch(() => null),
  ]);
  if (r7 && r7.total) out.ultimos7dias = _waResumoMeta(r7.total);
  if (r30 && r30.total) out.ultimos30dias = _waResumoMeta(r30.total);
  const bm = c.benchmark_metas || {}; const t = (r30 && r30.total) || null;
  if (t && Object.values(bm).some((v) => v != null)) {
    const cpConv = t.conversas ? t.spend / t.conversas : null, cpl = t.leads ? t.spend / t.leads : null, cpa = t.purchases ? t.spend / t.purchases : null;
    const cmp = (meta: any, atual: any, menorMelhor: boolean) => (meta == null || atual == null) ? null : { meta, atual: +atual.toFixed(2), atingida: menorMelhor ? atual <= meta : atual >= meta };
    const mvr: any = { ctr: cmp(bm.ctr, t.ctr, false), cpc: cmp(bm.cpc, t.cpc, true), cpm: cmp(bm.cpm, t.cpm, true), roas: cmp(bm.roas, t.roas, false), custoConversa: cmp(bm.custoConversa, cpConv, true), cpl: cmp(bm.cpl, cpl, true), cpa: cmp(bm.cpa, cpa, true) };
    Object.keys(mvr).forEach((k) => { if (!mvr[k]) delete mvr[k]; }); if (Object.keys(mvr).length) out.metasVsReal_30d = mvr;
  }
  if (r30 && r30.campaigns && r30.campaigns.length) out.campanhasComGasto_30d = r30.campaigns.slice(0, 25).map((x: any) => ({ nome: x.campaign, gasto: Math.round(x.spend), objetivo: (x.objetivo && x.objetivo.rotulo) || "", ctr: +(x.ctr || 0).toFixed(2), conversas: x.conversas || undefined, leads: x.leads || undefined, compras: x.purchases ? Math.round(x.purchases) : undefined }));
  if (ent) {
    const ativa = (x: any) => x.status === "ACTIVE" || x.entrega === "ACTIVE";
    out.campanhasAtivasAgora = (ent.campaigns || []).filter(ativa).slice(0, 30).map((x: any) => ({ nome: x.nome, objetivo: (x.objetivo && x.objetivo.rotulo) || "", orcamentoDiario: x.orcamentoDiario || undefined }));
    out.conjuntosAtivosComOrcamento = (ent.adsets || []).filter((x: any) => ativa(x) && x.orcamentoDiario).slice(0, 25).map((x: any) => ({ nome: x.nome, orcamentoDiario: x.orcamentoDiario }));
  }
  if (r30 && r30.ads && r30.ads.length) out.topAnuncios_30d = r30.ads.slice(0, 12).map((a: any) => ({ nome: a.adName, campanha: a.campaign, gasto: Math.round(a.spend), objetivo: (a.objetivo && a.objetivo.rotulo) || "", ctr: +(a.ctr || 0).toFixed(2), cpc: +(a.cpc || 0).toFixed(2), conversas: a.conversas || undefined, leads: a.leads || undefined }));
  if (!out.ultimos30dias && !out.campanhasComGasto_30d && !out.campanhasAtivasAgora) out.aviso = "Não consegui puxar os dados do Meta agora (token/conta).";
  return out;
}
async function waAgentLLM(text: string, history: any[], clientId: string | null, clients: any[]) {
  let snap = null; if (clientId) snap = await waAgentSnapshot(clientId);
  const names = clients.map((c) => c.name).slice(0, 250).join(" | ");
  const sys = `Você é a AndréIA, gestora de tráfego sênior, num grupo de WhatsApp com a EQUIPE da agência. Fale como gente: CURTO, direto, natural. Clientes: ${names}.

REGRA DE OURO — responda EXATAMENTE o que foi pedido, nada além:
- O SNAPSHOT é só seu conhecimento de fundo. NUNCA o recite/despeje. NÃO liste métricas a não ser que a pessoa PEÇA explicitamente análise/números/resultado/"como está".
- Se a pessoa pede uma AÇÃO (criar tarefa, pausar campanha, orçamento, lançamento...): responda em 1 linha confirmando SÓ a ação e pedindo SIM. NADA de métricas.
- Se pede análise/resultado: aí sim use os números do snapshot (curto, só o que importa).
- Se for conversa/dúvida: responda normal, curto.
- Identifique o cliente e devolva o nome EXATO em "client"; se não der, deixe vazio e pergunte. Um nome de PESSOA na frase (ex: "para o Dionathan") é o responsável da tarefa, não o cliente.

AÇÕES (execução real; sempre confirme com SIM antes — resuma no "reply" e preencha "action"). Tipos:
  · criar_tarefa: {"tipo":"criar_tarefa","nome":"<título>","obs":"<detalhe>"}
  · pausar_campanha / reativar_campanha: {"tipo":"pausar_campanha","campanha":"<nome exato da campanha>"}
  · orcamento: {"tipo":"orcamento","campanha":"<nome>","novoValor":<novo orçamento diário em R$, número>}
  · duplicar_campanha: {"tipo":"duplicar_campanha","campanha":"<nome>"}
  · criar_lancamento (financeiro): {"tipo":"criar_lancamento","natureza":"receita"|"despesa","descricao":"<ex: Fee mensal>","valor":<número>,"vencimento":"AAAA-MM-DD"}
  · dar_baixa (marcar lançamento como pago): {"tipo":"dar_baixa","descricao":"<parte da descrição do lançamento>"}
  Use o nome EXATO da campanha (do snapshot). NUNCA execute sem confirmação — e ao pedir confirmação SEMPRE diga o CLIENTE a que se refere (ex: "Crio a tarefa 'montar criativo' pro cliente MFlorImoveis (resp. Dionathan). Confirma?"). NUNCA diga que criou/fez algo por conta própria — quem executa é o sistema DEPOIS do SIM.
Responda SOMENTE JSON: {"client":"<nome|vazio>","reply":"<texto curto>","action":{...}|null}`;
  const ctx = snap ? ("SNAPSHOT COMPLETO de " + snap.nome + " (ultimos7dias, ultimos30dias, metasVsReal_30d, campanhasComGasto_30d, campanhasAtivasAgora com orçamento, conjuntosAtivosComOrcamento, topAnuncios_30d): " + JSON.stringify(snap).slice(0, 7000)) : "(nenhum cliente selecionado ainda)";
  const msgs = [{ role: "system", content: sys }, ...history.slice(-10).map((h: any) => ({ role: h.role === "assistant" ? "assistant" : "user", content: h.text })), { role: "user", content: ctx + "\n\nEQUIPE: " + text }];
  const j = await callOpenAI({ model: "gpt-4o-mini", messages: msgs, response_format: { type: "json_object" }, max_tokens: 700, temperature: 0.4 });
  try { return JSON.parse(j.choices[0].message.content || "{}"); } catch { return { reply: "Não entendi, pode repetir?", client: "", action: null }; }
}
async function waResolveCampaign(clientId: string, nome: string) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=meta_account_id`))[0];
  const ids = String(c?.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) return null;
  const ent = await metaEntities({ accounts: ids.map((id: string) => ({ id, name: id })) }).catch(() => null);
  if (!ent) return null;
  const q = String(nome || "").toLowerCase().trim(); const cs = ent.campaigns || [];
  return cs.find((c: any) => c.nome.toLowerCase() === q) || cs.find((c: any) => c.nome.toLowerCase().includes(q)) || cs.find((c: any) => q && q.includes(c.nome.toLowerCase())) || null;
}
async function _waClientNome(cid: string | null) { if (!cid) return ""; const c = (await sbGet("clients", `id=eq.${encodeURIComponent(cid)}&select=name`))[0]; return c?.name || ""; }
async function waAgentExec(pending: any, clientId: string | null) {
  const cid = pending.client_id || clientId;
  try {
    if (pending.tipo === "criar_tarefa") {
      const nome = pending.nome || "Tarefa (via AndréIA)";
      const res = await sbInsertOk("tasks", { id: _wuid(), name: nome, client: cid || null, owner: "eu", status: "todo", prio: pending.prio || "media", notes: pending.obs || "", urgent: false });
      if (!res.ok) return "❌ Não consegui salvar a tarefa: " + res.err;
      const cn = await _waClientNome(cid);
      return `✅ Tarefa criada${cn ? ` pro cliente ${cn}` : ""}: ${nome}`;
    }
    if (pending.tipo === "pausar_campanha" || pending.tipo === "reativar_campanha") {
      if (!cid) return "De qual cliente é a campanha?";
      const camp = await waResolveCampaign(cid, pending.campanha); if (!camp) return `Não achei a campanha "${pending.campanha || ""}".`;
      await metaAction({ action: pending.tipo === "pausar_campanha" ? "pause" : "activate", id: camp.id, nome: camp.nome });
      return (pending.tipo === "pausar_campanha" ? "⏸ Pausei" : "▶ Reativei") + ": " + camp.nome;
    }
    if (pending.tipo === "orcamento") {
      if (!cid) return "De qual cliente é a campanha?";
      const camp = await waResolveCampaign(cid, pending.campanha); if (!camp) return `Não achei a campanha "${pending.campanha || ""}".`;
      if (!camp.orcamentoDiario) return `A "${camp.nome}" não tem orçamento no nível da campanha (deve estar no conjunto). Quer que eu crie uma tarefa pra ajustar?`;
      await metaAction({ action: "budget", id: camp.id, nome: camp.nome, novoOrcamentoDiario: pending.novoValor });
      return `💰 Orçamento de "${camp.nome}" ajustado pra R$${Number(pending.novoValor).toFixed(2)}/dia.`;
    }
    if (pending.tipo === "duplicar_campanha") {
      if (!cid) return "De qual cliente é a campanha?";
      const camp = await waResolveCampaign(cid, pending.campanha); if (!camp) return `Não achei a campanha "${pending.campanha || ""}".`;
      await metaAction({ action: "duplicate", id: camp.id, nome: camp.nome });
      return `⧉ Dupliquei "${camp.nome}" (a cópia fica PAUSADA pra você revisar).`;
    }
    if (pending.tipo === "criar_lancamento") {
      const res = await sbInsertOk("finance", { id: _wuid(), type: pending.natureza === "despesa" ? "despesa" : "receita", client: cid || null, description: pending.descricao || "Lançamento (via AndréIA)", val: Number(pending.valor) || 0, due: pending.vencimento || new Date().toISOString().slice(0, 10), status: "pendente", auto: false });
      if (!res.ok) return "❌ Não consegui salvar o lançamento: " + res.err;
      const cn = await _waClientNome(cid);
      return `🧾 Lançamento criado${cn ? ` (cliente ${cn})` : ""}: ${pending.natureza === "despesa" ? "despesa" : "receita"} R$${(Number(pending.valor) || 0).toFixed(2)} — ${pending.descricao || ""} (venc. ${pending.vencimento || "hoje"}).`;
    }
    if (pending.tipo === "dar_baixa") {
      if (!cid) return "De qual cliente é o lançamento?";
      const term = String(pending.descricao || "").trim();
      const rows = await sbGet("finance", `client=eq.${encodeURIComponent(cid)}&status=eq.pendente${term ? `&description=ilike.*${encodeURIComponent(term)}*` : ""}&select=id,description,val&limit=6`);
      if (!rows.length) return `Não achei lançamento pendente${term ? ` com "${term}"` : ""} desse cliente.`;
      if (rows.length > 1) return `Achei ${rows.length} pendentes parecidos — seja mais específico: ${rows.map((r: any) => r.description).join(" / ")}`;
      await sbPatchD("finance", `id=eq.${encodeURIComponent(rows[0].id)}`, { status: "pago" });
      return `✅ Baixa dada: ${rows[0].description} (R$${Number(rows[0].val).toFixed(2)}).`;
    }
  } catch (e) { return "❌ Não consegui executar: " + String((e as any)?.message || e); }
  return "Feito 👍";
}
function _fmtR(v: number) { return "R$" + (Math.round((v || 0) * 100) / 100).toLocaleString("pt-BR", { minimumFractionDigits: 2 }); }
// métrica do OBJETIVO do cliente (deriva do que teve resultado): venda>lead>conversa>tráfego
function _objMetric(t: any, google: boolean) {
  const spend = t.spend || 0;
  if ((t.purchases || 0) > 0) { const roas = t.roas != null ? t.roas : (spend ? (t.revenue || 0) / spend : 0); return `Compras ${Math.round(t.purchases)} · ROAS ${(roas || 0).toFixed(2)}`; }
  if (!google && (t.leads || 0) > 0) return `Leads ${t.leads} · CPL ${_fmtR(t.leads ? spend / t.leads : 0)}`;
  if (!google && (t.conversas || 0) > 0) return `Conversas ${t.conversas} · Custo/conversa ${_fmtR(t.conversas ? spend / t.conversas : 0)}`;
  const ctr = t.ctr != null ? t.ctr : (t.impressions ? (t.clicks / t.impressions * 100) : 0);
  const cpc = t.cpc != null ? t.cpc : (t.clicks ? spend / t.clicks : 0);
  return `Cliques ${t.clicks || 0} · CTR ${(ctr || 0).toFixed(2)}% · CPC ${_fmtR(cpc)}`;
}
// Resumo de TODOS os clientes no período — por cliente, Meta/Google separados + total; pula quem não teve gasto. Retorna mensagens (chunked).
async function waAgentAllClientsSummary(days: number): Promise<string[]> {
  const since = new Date(Date.now() - days * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
  const clients = await sbGet("clients", "select=id,name,meta_account_id,google_account_id,status&limit=500");
  const active = clients.filter((c: any) => c.status !== "Encerrado" && (String(c.meta_account_id || "").trim() || String(c.google_account_id || "").trim()));
  const results: any[] = [];
  for (let i = 0; i < active.length; i += 8) {
    const ch = active.slice(i, i + 8);
    const rs = await Promise.all(ch.map(async (c: any) => {
      const mIds = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
      const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
      const [m, g] = await Promise.all([
        mIds.length ? metaAdsInsights({ accounts: mIds.map((id: string) => ({ id, name: id })), since, until }).catch(() => null) : Promise.resolve(null),
        gIds.length ? googleAdsInsights({ accounts: gIds.map((id: string) => ({ id, name: id })), since, until }).catch(() => null) : Promise.resolve(null),
      ]);
      const mt = (m && m.total && (m.total.spend || 0) > 0) ? m.total : null;
      const gt = (g && g.total && (g.total.spend || 0) > 0) ? g.total : null;
      return { nome: c.name, meta: mt, google: gt };
    }));
    results.push(...rs);
  }
  results.sort((a, b) => ((b.meta?.spend || 0) + (b.google?.spend || 0)) - ((a.meta?.spend || 0) + (a.google?.spend || 0)));
  const blocks: string[] = [];
  for (const r of results) {
    if (!r.meta && !r.google) continue;
    let b = `*${r.nome}*`;
    if (r.meta) b += `\n📘 Meta — Gasto ${_fmtR(r.meta.spend)} · ${_objMetric(r.meta, false)}`;
    if (r.google) b += `\n🔎 Google — Gasto ${_fmtR(r.google.spend)} · ${_objMetric(r.google, true)}`;
    if (r.meta && r.google) b += `\n➕ Total — Gasto ${_fmtR((r.meta.spend || 0) + (r.google.spend || 0))}`;
    blocks.push(b);
  }
  if (!blocks.length) return [`Nenhum cliente com investimento nos últimos ${days} dias.`];
  const msgs: string[] = []; let cur = `📊 Resumo dos últimos ${days} dias por cliente:\n`;
  for (const b of blocks) { if ((cur + "\n\n" + b).length > 3000) { msgs.push(cur); cur = b; } else cur += "\n\n" + b; }
  if (cur.trim()) msgs.push(cur);
  return msgs;
}
async function waAgentHandle(w: any) {
  const data = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  const cfg = data.andreia_wa || {};
  if (!cfg.instance_id || cfg.instance_id !== w.instanceId) return { skip: true };
  if (cfg.group_jid && w.chatid !== cfg.group_jid) return { skip: true };
  const allowed = (cfg.allowed || []).map((x: string) => String(x).replace(/[^0-9]/g, ""));
  const sender = String(w.sender || "").replace(/[^0-9]/g, "");
  if (allowed.length && !allowed.includes(sender)) return { skip: true };
  const inst = (await sbGet("wa_instances", `id=eq.${encodeURIComponent(w.instanceId)}&select=uaz_host,uaz_token`))[0];
  if (!inst) return { skip: true };
  const dest = w.chatid;
  const send = (t: string) => waCall(inst.uaz_host, inst.uaz_token, "/send/text", "POST", { number: dest, text: t });
  const skey = sender || w.chatid;
  let sess = (await sbGet("wa_agent_sessions", `phone=eq.${encodeURIComponent(skey)}&select=*`))[0];
  if (sess && sess.last_msgid === w.msgid) return { dup: true };
  const text = (w.text || "").trim(); if (!text) return { skip: true };
  const saveSess = async (patch: any) => { const row = { phone: skey, updated_at: new Date().toISOString(), ...patch }; if (sess) await sbPatchD("wa_agent_sessions", `phone=eq.${encodeURIComponent(skey)}`, row); else await sbPost("wa_agent_sessions", row); };
  if (sess && sess.pending) {
    const tl = text.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[\s.!,]+/g, " ").trim();
    const isYes = /^(sim|s|ss|isso|isso mesmo|pode|pode sim|confirmo|confirmado|confirmar|ok|okay|blz|beleza|claro|positivo|certo|faz|fazer|manda|manda ver|vai|bora|com certeza|👍|✅)$/.test(tl);
    const isNo = /^(nao|n|cancela|cancelar|deixa|esquece|para|negativo|nem|melhor nao)$/.test(tl);
    if (isYes) { const msg = await waAgentExec(sess.pending, sess.client_id); await saveSess({ pending: null, last_msgid: w.msgid }); await send(msg); return { ok: true }; }
    if (isNo) { await saveSess({ pending: null, last_msgid: w.msgid }); await send("Ok, cancelei 👍"); return { ok: true }; }
  }
  // pedido de RESUMO GERAL de todos os clientes → monta determinístico (sem alucinar placeholders)
  const low = text.toLowerCase();
  if (/\b(cada cliente|todos os clientes|todos clientes|resumo geral|de todos|geral dos clientes|resumo dos clientes|panorama|de cada cliente)\b/.test(low) || (/resumo/.test(low) && /clientes/.test(low))) {
    const days = /\b90\b/.test(text) ? 90 : (/\b30\b/.test(text) ? 30 : 7);
    await send(`⏳ Montando o resumo de todos os clientes (${days} dias)… um instante.`);
    const msgs = await waAgentAllClientsSummary(days);
    for (const mm of msgs) await send(mm);
    await saveSess({ pending: null, last_msgid: w.msgid, history: [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: "[resumo geral de clientes enviado]" }].slice(-16) });
    return { ok: true };
  }
  const clients = await sbGet("clients", "select=id,name&limit=500");
  const out = await waAgentLLM(text, (sess && sess.history) || [], (sess && sess.client_id) || null, clients);
  let clientId = (sess && sess.client_id) || null;
  if (out.client) { const q = String(out.client).toLowerCase(); const m = clients.find((c: any) => c.name.toLowerCase() === q) || clients.find((c: any) => c.name.toLowerCase().includes(q)); if (m) clientId = m.id; }
  let reply = out.reply || "Ok.";
  const pending = out.action ? { ...out.action, client_id: clientId } : null;
  // garante que a confirmação SEMPRE mostra o cliente (o modelo às vezes esquece)
  if (pending && clientId) {
    const cn = (clients.find((c: any) => c.id === clientId) || {}).name || "";
    if (cn && !reply.toLowerCase().includes(cn.toLowerCase())) reply = `📌 Cliente: ${cn}\n` + reply;
  }
  const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: reply }].slice(-16);
  await saveSess({ client_id: clientId, history: hist, pending, last_msgid: w.msgid });
  await send(reply);
  return { ok: true };
}
async function waHandler(w: any) {
  if (w.op === "extract") return await waExtract(w.convId);
  if (w.op === "capi") return await waCapi(w.convId, w.event);
  // criar instância nova (número da agência ou de um cliente) — não precisa de instanceId
  if (w.op === "create") {
    const uz = await waUzConfig();
    const name = String(w.name || ("num-" + _wuid())).replace(/[^a-zA-Z0-9-]/g, "-").slice(0, 40);
    const init = await fetch(uz.server.replace(/\/$/, "") + "/instance/init", { method: "POST", headers: { admintoken: uz.admin_token, "Content-Type": "application/json" }, body: JSON.stringify({ name }) });
    const ij = await init.json().catch(() => ({}));
    const itoken = ij.token || (ij.instance && ij.instance.token);
    if (!itoken) throw new Error("uazapi não devolveu token da instância: " + JSON.stringify(ij).slice(0, 160));
    const id = _wuid();
    await sbPost("wa_instances", { id, client_id: w.clientId || null, name, uaz_token: itoken, uaz_host: uz.server, status: "connecting" });
    const hook = `${_SB_URL}/functions/v1/tracking/wa/webhook/${id}`;
    // AndréIA precisa RECEBER mensagens de grupo; o CRM não (isGroupYes)
    const excl = w.includeGroups ? ["wasSentByApi"] : ["wasSentByApi", "isGroupYes"];
    try { await waCall(uz.server, itoken, "/webhook", "POST", { enabled: true, url: hook, events: ["messages", "connection"], excludeMessages: excl }); } catch (_e) {}
    return { id };
  }
  const inst = (await sbGet("wa_instances", `id=eq.${encodeURIComponent(w.instanceId)}&select=id,client_id,uaz_host,uaz_token,phone`))[0];
  if (!inst) throw new Error("Instância WhatsApp não encontrada.");
  const host = inst.uaz_host, token = inst.uaz_token, clientId = inst.client_id || null;
  const clientFilter = clientId ? "eq." + encodeURIComponent(clientId) : "is.null";
  if (w.op === "status") {
    const { j } = await waCall(host, token, "/instance/status"); const ins = (j && j.instance) || {};
    if (ins.status) { const patch: Record<string, unknown> = { status: ins.status, updated_at: new Date().toISOString() }; if (ins.owner) patch.phone = String(ins.owner).replace(/@.*$/, ""); if (ins.status === "connected") patch.connected_at = new Date().toISOString(); await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(inst.id)}`, patch); }
    return { status: ins.status || "unknown", phone: ins.owner || inst.phone || "", instance: ins };
  }
  if (w.op === "qr") {
    const { j } = await waCall(host, token, "/instance/connect", "POST", w.phone ? { phone: String(w.phone).replace(/[^0-9]/g, "") } : {});
    const ins = (j && j.instance) ? j.instance : (j || {});
    return { qrcode: ins.qrcode || "", paircode: ins.paircode || "", status: ins.status || "connecting" };
  }
  if (w.op === "groups") {
    const { j } = await waCall(host, token, "/group/list");
    const gs = (j && j.groups) || [];
    return { groups: gs.map((g: any) => ({ jid: g.JID || g.jid || "", name: g.Name || g.name || g.JID || "" })).filter((g: any) => g.jid) };
  }
  if (w.op === "remove") {
    const uz = await waUzConfig().catch(() => null);
    try { await waCall(host, token, "/instance/logout", "POST", {}); } catch (_e) {}
    if (uz) try { await fetch(host.replace(/\/$/, "") + "/instance", { method: "DELETE", headers: { admintoken: uz.admin_token, token } }); } catch (_e) {}
    await fetch(`${_SB_URL}/rest/v1/wa_instances?id=eq.${encodeURIComponent(inst.id)}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
    return { removed: true };
  }
  if (w.op === "send") {
    const number = String(w.number).replace(/[^0-9]/g, "");
    const { status, j } = await waCall(host, token, "/send/text", "POST", { number, text: w.text });
    if (status >= 200 && status < 300) {
      const ts = new Date().toISOString();
      const conv = (await sbGet("wa_conversations", `client_id=${clientFilter}&chat_id=eq.${number}&select=id&limit=1`))[0];
      let convId = conv?.id;
      if (!convId) { convId = _wuid(); await sbPost("wa_conversations", { id: convId, client_id: clientId, chat_id: number, name: number, last_text: w.text, last_at: ts, origin_type: "organico" }); }
      else await sbPatchD("wa_conversations", `id=eq.${convId}`, { last_text: w.text, last_at: ts });
      await sbPost("wa_messages", { id: _wuid(), client_id: clientId, conversation_id: convId, chat_id: number, wa_msgid: String((j && (j.id || j.messageid)) || _wuid()), direction: "out", msg_type: "text", text: w.text, ts, raw: j });
    }
    return { ok: status >= 200 && status < 300, status, resp: j };
  }
  if (w.op === "poll") {
    const { j } = await waCall(host, token, "/message/find", "POST", { limit: w.limit || 60 });
    const msgs: any[] = (j && j.messages) || [];
    const oneToOne = msgs.filter((m) => !(m.isGroup || String(m.chatid || "").endsWith("@g.us")));
    const ids = oneToOne.map((m) => String(m.messageid || m.id || "")).filter(Boolean);
    const known = new Set<string>();
    if (ids.length) { const ex = await sbGet("wa_messages", `wa_msgid=in.(${ids.map((x) => encodeURIComponent(x)).join(",")})&select=wa_msgid`); (ex || []).forEach((r: any) => known.add(r.wa_msgid)); }
    const adCache: Record<string, Record<string, string> | null> = {};
    const newInbound = new Set<string>();
    let added = 0;
    for (const m of oneToOne) {
      const msgid = String(m.messageid || m.id || ""); if (!msgid || known.has(msgid)) continue;
      const phone = String(m.chatid || m.sender_pn || m.sender || "").replace(/@.*$/, "").replace(/[^0-9]/g, ""); if (!phone) continue;
      const fromMe = !!m.fromMe; const text = waText(m); const ts = waTs(m.messageTimestamp);
      const existing = (await sbGet("wa_conversations", `client_id=${clientFilter}&chat_id=eq.${phone}&select=id,origin_type&limit=1`))[0];
      let convId = existing?.id; const origin = fromMe ? null : waOrigin(m);
      if (origin && origin.type === "anuncio" && origin.data.source_id && !origin.data.campaign) {
        const key = String(origin.data.source_id);
        if (adCache[key] === undefined) adCache[key] = await waResolveAd(key);
        if (adCache[key]) Object.assign(origin.data, adCache[key]);
      }
      if (!convId) { convId = _wuid(); await sbPost("wa_conversations", { id: convId, client_id: clientId, chat_id: phone, name: m.senderName || phone, last_text: text, last_at: ts, unread: fromMe ? 0 : 1, origin_type: origin ? origin.type : "organico", origin: origin ? origin.data : null }); }
      else { const patch: Record<string, unknown> = { last_text: text, last_at: ts }; if (!fromMe) patch.unread = 1; if (origin && (!existing.origin_type || existing.origin_type === "organico")) { patch.origin_type = origin.type; patch.origin = origin.data; } await sbPatchD("wa_conversations", `id=eq.${convId}`, patch); }
      await sbPost("wa_messages", { id: _wuid(), client_id: clientId, conversation_id: convId, chat_id: phone, wa_msgid: msgid, direction: fromMe ? "out" : "in", msg_type: m.messageType || "text", text, ts, raw: m });
      if (!fromMe) newInbound.add(convId);
      added++;
    }
    // classificação AUTOMÁTICA por IA das conversas que receberam nova mensagem do lead (limita p/ controlar custo)
    let classified = 0;
    if (newInbound.size) { for (const cid of [...newInbound].slice(0, 6)) { try { await waExtract(cid, true); classified++; } catch (_e) {} } }
    return { added, scanned: msgs.length, classified };
  }
  if (w.op === "resolveOrigins") {
    const convs = await sbGet("wa_conversations", `client_id=${clientFilter}&origin_type=eq.anuncio&select=id,origin`);
    const cache: Record<string, Record<string, string> | null> = {}; let done = 0;
    for (const cv of (convs || [])) {
      const o = cv.origin || {}; if (o.campaign || !o.source_id) continue;
      const key = String(o.source_id); if (cache[key] === undefined) cache[key] = await waResolveAd(key);
      if (cache[key]) { await sbPatchD("wa_conversations", `id=eq.${cv.id}`, { origin: { ...o, ...cache[key] } }); done++; }
    }
    return { resolved: done };
  }
  throw new Error("op inválida");
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  try {
    const body = await req.json();
    if (body.wa) {
      const r = await waHandler(body.wa);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.waAgent) {
      const r = await waAgentHandle(body.waAgent);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    const spreadsheetId = body.spreadsheetId;
    const tabs = body.tabs;
    const orders = body.orders;
    const analysis = body.analysis;
    const agent = body.agent;

    if (body.nuvemshopOrders) {
      const r = await nuvemshopOrders(body.nuvemshopOrders);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.raioxAI) {
      const r = await raioxAI(body.raioxAI);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.siteAudit) {
      const r = await siteAudit(body.siteAudit);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.rdCatalog) {
      const r = await rdCatalog(body.rdCatalog);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaBreakdowns) {
      const r = await metaBreakdowns(body.metaBreakdowns);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleBreakdowns) {
      const r = await googleBreakdowns(body.googleBreakdowns);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleAds) {
      const r = await googleAdsInsights(body.googleAds);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleAccounts) {
      const r = await googleListAccounts();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAds) {
      const r = await metaAdsInsights(body.metaAds);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAccounts) {
      const r = await metaListAccounts();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaFunding) {
      const r = await metaFunding(body.metaFunding);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaEntities) {
      const r = await metaEntities(body.metaEntities);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAction) {
      const r = await metaAction(body.metaAction);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.dnaExtract) {
      let text = body.dnaExtract.text || "";
      if (!text && body.dnaExtract.url) text = await fetchUrlText(body.dnaExtract.url);
      const r = await extractDna(text, body.dnaExtract.direcionamento || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.dnaRefine) {
      const r = await refineDna(body.dnaRefine.dna || {}, body.dnaRefine.instrucao || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (agent) {
      const r = await runAgent(agent);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (analysis) {
      const analysisText = await generateAnalysis(analysis, body.chat, body.styleExamples);
      return new Response(JSON.stringify({ data: { analysisText } }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if ((!spreadsheetId || !Array.isArray(tabs)) && !orders) {
      return new Response(JSON.stringify({ error: "spreadsheetId e tabs (array), ou orders, sao obrigatorios" }), { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }

    const keyJson = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_KEY");
    if (!keyJson) throw new Error("GOOGLE_SERVICE_ACCOUNT_KEY nao configurada");
    const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(keyJson), scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"] });
    const sheets = google.sheets({ version: "v4", auth });

    const result: Record<string, unknown> = {};
    for (const tab of tabs || []) {
      if (isBlockedTab(tab)) { result[tab] = { error: "Aba bloqueada por conter possivel dado pessoal (nome contem termo restrito)" }; continue; }
      try {
        const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tab}'!A1:Z5000` });
        result[tab] = res.data.values || [];
      } catch (e) { result[tab] = { error: (e as Error).message }; }
    }

    if (orders && orders.tab) {
      const ids: string[] = orders.spreadsheetIds || (orders.spreadsheetId ? [orders.spreadsheetId] : []);
      if (ids.length) {
        try { result[orders.tab] = await aggregateOrdersTabs(sheets, ids, orders.tab); }
        catch (e) { result[orders.tab] = { error: (e as Error).message }; }
      }
    }

    return new Response(JSON.stringify({ data: result }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
  } catch (e) {
    return new Response(JSON.stringify({ error: (e as Error).message }), { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } });
  }
});
