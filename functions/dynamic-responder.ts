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
      const [accountRows, acctDaily, objByCampId, adRows, campRows] = await Promise.all([
        fetchInsights(acc.id, "account"),
        m.daily ? fetchInsights(acc.id, "account", "&time_increment=1") : Promise.resolve([] as any[]),
        wantObj ? fetchObjectives(acc.id) : Promise.resolve({} as Record<string, any>),
        m.byAd ? fetchInsights(acc.id, "ad") : Promise.resolve([] as any[]),
        m.byCampaign ? fetchInsights(acc.id, "campaign", m.daily ? "&time_increment=1" : "") : Promise.resolve([] as any[]),
      ]);
      return { acc, accountRows, acctDaily, objByCampId, adRows, campRows, error: statusIssue as string | null };
    } catch (e) {
      // conta com erro NAO derruba as outras: devolve vazia + motivo (front mostra o disclaimer)
      return { acc, accountRows: [] as any[], acctDaily: [] as any[], objByCampId: {} as Record<string, any>, adRows: [] as any[], campRows: [] as any[], error: statusIssue || (e as any)?.message || String(e) };
    }
  }));
  const accountErrors = perAccount.filter((p) => p.error).map((p) => ({ id: p.acc.id, name: p.acc.name || p.acc.id, error: p.error }));
  const totRecByDate: Record<string, any> = {};
  for (const { acc, accountRows, acctDaily, objByCampId, adRows, campRows } of perAccount) {
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
      c.spend += s.spend; c.impressions += s.impressions; c.clicks += s.clicks; c.reach += s.reach;
      c.revenue += s.revenue; c.purchases += s.purchases; c.leads += s.leads; c.addToCart += s.addToCart; c.initiateCheckout += s.initiateCheckout;
      if (m.daily) c.records.push({ date: row.date_start, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: s.reach, leads: s.leads, conversas: s.conversas, videoViews: s.videoViews, engajamentos: s.engajamentos });
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

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  try {
    const body = await req.json();
    const spreadsheetId = body.spreadsheetId;
    const tabs = body.tabs;
    const orders = body.orders;
    const analysis = body.analysis;
    const agent = body.agent;

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
