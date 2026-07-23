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

// Clona a ESTRUTURA de uma campanha (campanha + conjuntos, PAUSADOS) pra OUTRA conta de anúncio (outro cliente).
// Não copia criativos/anúncios (são amarrados à conta de origem). Mapeia o pixel pro do destino quando possível.
async function metaCloneCampaign(m: any) {
  const token = Deno.env.get("META_USER_TOKEN"); if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const base = "https://graph.facebook.com/v21.0";
  const src = String(m.sourceCampaignId || ""); const tgt = String(m.targetAccountId || "").replace(/^act_/, "");
  if (!src || !tgt) throw new Error("sourceCampaignId e targetAccountId obrigatórios");
  const get = async (id: string, fields: string) => { const r = await fetch(`${base}/${id}?fields=${fields}&access_token=${token}`); const j = await r.json(); if (j.error) throw new Error(j.error.message); return j; };
  const post = async (path: string, params: Record<string, string>) => { const r = await fetch(`${base}/${path}`, { method: "POST", body: new URLSearchParams({ ...params, access_token: token }) }); const j = await r.json(); if (j.error) throw new Error(j.error.message); return j; };
  const camp = await get(src, "name,objective,special_ad_categories,buying_type,bid_strategy,daily_budget,lifetime_budget");
  const campParams: Record<string, string> = { name: (camp.name || "Campanha") + " (clone)", objective: camp.objective, status: "PAUSED", special_ad_categories: JSON.stringify(camp.special_ad_categories || []) };
  if (camp.buying_type) campParams.buying_type = camp.buying_type;
  if (camp.bid_strategy) campParams.bid_strategy = camp.bid_strategy;
  if (camp.daily_budget) campParams.daily_budget = String(camp.daily_budget);
  else if (camp.lifetime_budget) campParams.lifetime_budget = String(camp.lifetime_budget);
  let newCamp: any;
  try { newCamp = await post(`act_${tgt}/campaigns`, campParams); } catch (e) { throw new Error("Não consegui criar a campanha no cliente de destino: " + String((e as any)?.message || e)); }
  const newCampId = newCamp.id;
  let tgtPixel: string | null = null; try { const px = await get(`act_${tgt}/adspixels`, "id"); tgtPixel = (px.data && px.data[0] && px.data[0].id) || null; } catch (_e) { /* */ }
  const asRes = await get(`${src}/adsets`, "name,optimization_goal,billing_event,daily_budget,lifetime_budget,bid_amount,targeting,end_time,destination_type,promoted_object").catch(() => ({ data: [] }));
  const adsets = asRes.data || [];
  const criados: any[] = [], falhas: any[] = [];
  for (const as of adsets) {
    const p: Record<string, string> = { name: as.name || "Conjunto", campaign_id: newCampId, status: "PAUSED", billing_event: as.billing_event || "IMPRESSIONS", optimization_goal: as.optimization_goal || "REACH" };
    if (as.daily_budget) p.daily_budget = String(as.daily_budget); else if (as.lifetime_budget) { p.lifetime_budget = String(as.lifetime_budget); if (as.end_time) p.end_time = as.end_time; }
    if (as.bid_amount) p.bid_amount = String(as.bid_amount);
    if (as.targeting) { const t = { ...as.targeting }; delete t.custom_audiences; delete t.excluded_custom_audiences; p.targeting = JSON.stringify(t); }
    if (as.destination_type) p.destination_type = as.destination_type;
    let fallback = false;
    if (as.promoted_object) {
      const po = as.promoted_object;
      if (po.pixel_id && tgtPixel) { const npo: any = { pixel_id: tgtPixel }; if (po.custom_event_type) npo.custom_event_type = po.custom_event_type; p.promoted_object = JSON.stringify(npo); }
      else { fallback = true; p.optimization_goal = "LINK_CLICKS"; p.billing_event = "IMPRESSIONS"; }
    }
    try { const na = await post(`act_${tgt}/adsets`, p); criados.push({ nome: as.name, id: na.id, fallback }); }
    catch (e) { falhas.push({ nome: as.name, erro: String((e as any)?.message || e).slice(0, 160) }); }
  }
  return { ok: true, campanhaId: newCampId, campanhaNome: campParams.name, criados, falhas, pixelDestino: tgtPixel };
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
        g.byCampaign ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, campaign.advertising_channel_type, campaign_budget.amount_micros, campaign_budget.resource_name${g.daily ? ", segments.date" : ""}, ${GADS_METRICS_FULL} FROM campaign WHERE ${range}`, token) : Promise.resolve([] as any[]),
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
      if (!byCamp[label]) byCamp[label] = { campaign: label, campaignId: row.campaign?.id ? String(row.campaign.id) : null, account: acc.name || acc.id, accountId: acc.id, objetivo: googleObjetivo(row.campaign?.advertisingChannelType), _google: true, orcamentoDiario: row.campaignBudget?.amountMicros ? +row.campaignBudget.amountMicros / 1e6 : null, budgetResource: row.campaignBudget?.resourceName || null, spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, records: [] };
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
  // Campanhas com gasto que NÃO produzem linhas de anúncio (Performance Max, Shopping, Demand Gen — não têm ad_group_ad):
  // sintetiza uma linha em nível de campanha pra elas aparecerem na árvore de campanhas.
  if (g.byAd && g.byCampaign) {
    const comAd = new Set(ads.map((a: any) => a.campaignId).filter(Boolean));
    for (const c of Object.values(byCamp) as any[]) {
      if ((c.spend || 0) <= 0) continue;
      if (c.campaignId && comAd.has(c.campaignId)) continue;
      ads.push({
        adId: c.campaignId ? "gc" + c.campaignId : null, adName: (c.objetivo && c.objetivo.rotulo) || "Campanha",
        campaign: c.campaign, campaignId: c.campaignId, adset: "", adsetId: null,
        account: c.account, thumbnail: null, _google: true, _campaignLevel: true, objetivo: c.objetivo,
        spend: c.spend, impressions: c.impressions, clicks: c.clicks, reach: 0, frequency: 0,
        ctr: c.impressions ? (c.clicks / c.impressions) * 100 : 0, cpc: c.clicks ? c.spend / c.clicks : 0, cpm: c.impressions ? (c.spend / c.impressions) * 1000 : 0,
        purchases: c.purchases, revenue: c.revenue, roas: c.spend ? c.revenue / c.spend : 0,
        leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0, videoViews: c.videoViews || 0, engajamentos: c.engajamentos || 0,
        cpa: c.purchases ? c.spend / c.purchases : 0,
      });
    }
  }
  ads.sort((a: any, b: any) => b.spend - a.spend);
  return { total, campaigns, ads, accounts, accountErrors, period: { since, until } };
}

// Ajusta o orçamento diário de uma campanha do Google (mutate no campaign_budget).
async function googleUpdateBudget(m: any) {
  const cid = String(m.accountId || "").replace(/-/g, ""); const res = m.budgetResource; const novo = Number(m.novoValor);
  if (!cid || !res || !(novo > 0)) throw new Error("accountId, budgetResource e novoValor obrigatórios");
  const token = await googleAdsAccessToken();
  const devToken = Deno.env.get("GOOGLE_ADS_DEV_TOKEN"); const mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  const body = { operations: [{ updateMask: "amount_micros", update: { resourceName: res, amountMicros: String(Math.round(novo * 1e6)) } }] };
  const r = await fetch(`https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/campaignBudgets:mutate`, { method: "POST", headers: { "Authorization": `Bearer ${token}`, "developer-token": devToken!, "login-customer-id": mcc, "Content-Type": "application/json" }, body: JSON.stringify(body) });
  const j = await r.json();
  if (j.error) throw new Error(j?.error?.details?.[0]?.errors?.[0]?.message || j.error.message || "erro no Google Ads");
  return { ok: true, detail: `Orçamento diário ajustado para R$${novo.toFixed(2)}` };
}
// Sugestão de LIMPEZA de termos de busca (palavras-chave negativas) com base no DNA do cliente.
async function googleTermCleanup(m: any) {
  const termos: any[] = (m.termos || []).slice(0, 80);
  if (!termos.length) return { negativar: [], observacao: "Sem termos de busca no período." };
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(m.clientId || "")}&select=name,dna,seg`))[0];
  const dna = (c && c.dna) || {};
  const ctx = {
    marca: dna?.identidade?.marca || c?.name || "",
    posicionamento: dna?.identidade?.posicionamento || "",
    sobre: dna?.identidade?.sobre || "",
    segmento: c?.seg || "",
    produtos: (dna?.produtos || []).map((p: any) => p.nome).filter(Boolean).slice(0, 15),
    personas: (dna?.personas || []).map((p: any) => p.titulo).filter(Boolean).slice(0, 8),
    palavrasProibidas: (dna?.diretrizes?.palavrasProibidas || []).slice(0, 30),
    palavrasRessoam: (dna?.diretrizes?.palavrasRessoam || []).slice(0, 30),
  };
  const lista = termos.map((t: any) => ({ termo: t.key, gasto: Math.round(t.spend || 0), cliques: Math.round(t.clicks || 0), conversoes: +(t.conversions || 0).toFixed(1) }));
  const sys = `Você é especialista em Google Ads e gestão de palavras-chave NEGATIVAS. Recebe o DNA do cliente (o que vende, personas, palavras que ressoam e proibidas) e a lista de TERMOS DE BUSCA reais que dispararam os anúncios. Sua tarefa: identificar termos IRRELEVANTES / fora do público / que não têm a ver com o que o cliente vende (candidatos a palavra-chave NEGATIVA), para limpar o tráfego. Seja criterioso: só marque como negativar se realmente foge do negócio/persona (ex: busca por concorrente, produto que não vende, intenção errada, gratuito quando é pago, localidade errada). Termos com CONVERSÃO geralmente NÃO devem ser negativados. Responda SOMENTE JSON: {"negativar":[{"termo":"...","motivo":"curto"}],"observacao":"1 frase geral"}.`;
  const user = `DNA do cliente:\n${JSON.stringify(ctx)}\n\nTermos de busca (com gasto/cliques/conversões):\n${JSON.stringify(lista)}`;
  try {
    const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 1500, temperature: 0.3 });
    const parsed = JSON.parse(j.choices[0].message.content || "{}");
    return { negativar: (parsed.negativar || []).slice(0, 60), observacao: parsed.observacao || "", cliente: c?.name || "", temDna: !!(dna && Object.keys(dna).length) };
  } catch (e) { return { erro: String((e as any)?.message || e) }; }
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
  { key: "sem", label: "Sem etapa", desc: "" },
  { key: "novo", label: "Lead novo", desc: "Contato inicial. Só mandou a mensagem automática/genérica vinda do anúncio (ex: 'quero informações', 'saber mais') e ainda NÃO deu sinal real de qualificação. Permanece aqui até responder com interesse comercial concreto." },
  { key: "mql", label: "MQL", event: "Lead", desc: "Marketing Qualified Lead. Demonstrou interesse REAL no produto/serviço: perguntou sobre preço, disponibilidade, como funciona, pediu informações específicas — qualquer sinal de interesse comercial. NÃO classificar como MQL contatos que enviaram apenas a mensagem automática/genérica do anúncio, mesmo que contenha palavras como 'interesse', 'informações' ou 'saber mais'." },
  { key: "sql", label: "SQL", event: "QualifiedLead", desc: "Sales Qualified Lead. Definiu o que quer e está pronto para proposta: especificou produto, data, quantidade, pediu orçamento formal, quer agendar, ou está negociando condições de pagamento." },
  { key: "comprou", label: "Comprou", event: "Purchase", desc: "Pagamento ou contratação confirmada. Pix/cartão/link pago, agendamento confirmado com pagamento, contrato assinado. Precisa de confirmação EXPLÍCITA de fechamento." },
  { key: "posvenda", label: "Pós-Venda", desc: "Já é cliente; comunicação pós-compra (suporte, onboarding, recompra)." },
  { key: "perdido", label: "Perdido", desc: "Desistiu, sumiu ou disse que não tem interesse." },
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
// resumo já FILTRADO pela métrica do objetivo (venda→ROAS/CPA; leads→CPL; mensagem→custo por conversa; senão tráfego)
function _waResumoMeta(t: any) {
  const spend = t.spend || 0;
  const objetivo = (t.purchases || 0) > 0 ? "venda" : ((t.leads || 0) > 0 ? "leads" : ((t.conversas || 0) > 0 ? "mensagens" : "trafego"));
  const out: any = { objetivo, gasto: Math.round(spend), impressoes: t.impressions, cliques: t.clicks, ctr: +(t.ctr || 0).toFixed(2), cpc: +(t.cpc || 0).toFixed(2), cpm: +(t.cpm || 0).toFixed(2), alcance: t.reach };
  if (objetivo === "venda") { out.compras = Math.round(t.purchases || 0); out.roas = +(t.roas || 0).toFixed(2); out.cpa = t.purchases ? +(spend / t.purchases).toFixed(2) : null; }
  else if (objetivo === "leads") { out.leads = t.leads; out.cpl = t.leads ? +(spend / t.leads).toFixed(2) : null; }
  else if (objetivo === "mensagens") { out.conversas = t.conversas; out.custoPorConversa = t.conversas ? +(spend / t.conversas).toFixed(2) : null; }
  return out;
}
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
- REUNIÃO ≠ TAREFA: se pedirem "reuniões/agenda/compromissos/calls", use a ferramenta *reunioes* e liste SÓ reuniões — NUNCA misture tarefas operacionais. Se não houver reunião, diga que não há reunião no período (não caia pra tarefas).
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
      const urgente = !!pending.urgente, revisao = !!pending.revisao;
      const res = await sbInsertOk("tasks", { id: _wuid(), name: nome, client: cid || null, owner: pending._owner || "eu", status: "todo", prio: urgente ? "alta" : (pending.prio || "media"), notes: pending.obs || "", due: pending._due || null, urgent: urgente, review_requested: revisao, reviewer: revisao ? "eu" : null });
      if (!res.ok) return "❌ Não consegui salvar a tarefa: " + res.err;
      const cn = await _waClientNome(cid);
      return `✅ Tarefa criada${cn ? ` pro cliente ${cn}` : ""}${pending.responsavel ? ` · resp. ${pending.responsavel}` : ""}${pending._due ? ` · ${pending._due}` : ""}${urgente ? " · 🔴 URGENTE" : ""}${revisao ? " · 🔎 revisão solicitada" : ""}: ${nome}`;
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
    if (pending.tipo === "criar_reuniao") {
      const summary = pending.nome || "Reunião"; const date = pending._due; const time = pending.hora || "";
      if (!date) return "Pra qual dia é a reunião?";
      const r = await fetch(`${_SB_URL}/functions/v1/tracking/calendar/create`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ summary, date, time, clientId: cid || null }) });
      const d = await r.json();
      if (d.error === "reconnect") return "⚠ Preciso de permissão de edição na Google Agenda — reconecte o Google no sistema (⚙️ Configurações → Google Agenda).";
      if (d.error) return "❌ Não consegui criar a reunião: " + d.error;
      const cn = await _waClientNome(cid);
      return `📅 Reunião criada na agenda${cn ? ` (cliente ${cn})` : ""}: ${summary}${time ? ` às ${time}` : ""} — ${date.split("-").reverse().join("/")}.`;
    }
    if (pending.tipo === "cancelar_reuniao") {
      const term = String(pending.nome || "").trim();
      const q = ["id=like.cal*", "status=neq.done", "select=id,name,due,notes", "order=due.asc", "limit=8"];
      if (pending._due) q.push(`due=eq.${pending._due}`);
      if (term) q.push(`name=ilike.*${encodeURIComponent(term)}*`);
      const rows = await sbGet("tasks", q.join("&"));
      if (!rows.length) return `Não achei essa reunião${term ? ` ("${term}")` : ""}${pending._due ? ` em ${pending._due}` : ""}.`;
      if (rows.length > 1) return `Achei ${rows.length} reuniões parecidas — qual? ${rows.map((r: any) => `${r.name}${r.due ? ` (${r.due.split("-").reverse().join("/")})` : ""}`).join(" / ")}`;
      const r = await fetch(`${_SB_URL}/functions/v1/tracking/calendar/delete`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ taskId: rows[0].id }) });
      const d = await r.json();
      if (d.error === "reconnect") return "⚠ Preciso de permissão de edição na Google Agenda — reconecte o Google no sistema.";
      if (d.error) return "❌ Não consegui excluir: " + d.error;
      return `🗑 Reunião cancelada: ${rows[0].name}.`;
    }
  } catch (e) { return "❌ Não consegui executar: " + String((e as any)?.message || e); }
  return "Feito 👍";
}
function _fmtR(v: number) { return "R$" + (Math.round((v || 0) * 100) / 100).toLocaleString("pt-BR", { minimumFractionDigits: 2 }); }
// Objetivo DOMINANTE (por gasto) das campanhas de um resultado de insights (metaAdsInsights/googleAdsInsights com byCampaign)
function _domObj(r: any): string | null {
  if (!r || !r.campaigns || !r.campaigns.length) return null;
  const byTipo: Record<string, number> = {};
  r.campaigns.forEach((c: any) => { const tp = (c.objetivo && c.objetivo.tipo) || "outro"; byTipo[tp] = (byTipo[tp] || 0) + (c.spend || 0); });
  let best: string | null = null, bestv = -1;
  for (const k in byTipo) { if (byTipo[k] > bestv) { bestv = byTipo[k]; best = k; } }
  return best;
}
// métrica do OBJETIVO. Se `obj` vier (tipo dominante das campanhas), segue ELE; senão cai no heurístico por presença de valor.
function _objLabel(o?: string | null) { return ({ conversao: "venda", app: "venda", leads: "leads", mensagens: "mensagens", video: "vídeo", alcance: "alcance", distribuicao: "alcance", engajamento: "engajamento", trafego: "tráfego" } as Record<string, string>)[o || ""] || "tráfego"; }
function _objMetric(t: any, google: boolean, obj?: string | null) {
  const spend = t.spend || 0;
  const isVenda = obj === "conversao" || obj === "app" || (!obj && (t.purchases || 0) > 0);
  const isLead = obj === "leads" || (!obj && !google && (t.leads || 0) > 0);
  const isMsg = obj === "mensagens" || (!obj && !google && (t.conversas || 0) > 0);
  const isVideo = obj === "video";
  const isAlcance = obj === "alcance" || obj === "distribuicao";
  const isEngaj = obj === "engajamento";
  if (isVenda) { const roas = t.roas != null ? t.roas : (spend ? (t.revenue || 0) / spend : 0); return `Compras ${Math.round(t.purchases || 0)} · ROAS ${(roas || 0).toFixed(2)}`; }
  if (isLead) return `Leads ${t.leads || 0} · CPL ${_fmtR(t.leads ? spend / t.leads : 0)}`;
  if (isMsg) return `Conversas ${t.conversas || 0} · Custo/conversa ${_fmtR(t.conversas ? spend / t.conversas : 0)}`;
  if (isVideo) return `Views ${Math.round(t.videoViews || 0)} · Custo/view ${_fmtR(t.videoViews ? spend / t.videoViews : 0)}`;
  if (isAlcance) { const cpm = t.cpm != null ? t.cpm : (t.impressions ? spend / t.impressions * 1000 : 0); const freq = t.reach ? t.impressions / t.reach : 0; return `Alcance ${Math.round(t.reach || 0).toLocaleString("pt-BR")} · CPM ${_fmtR(cpm)}${freq ? ` · Freq ${freq.toFixed(2)}` : ""}`; }
  if (isEngaj) return `Engajamentos ${Math.round(t.engajamentos || 0).toLocaleString("pt-BR")} · Custo ${_fmtR(t.engajamentos ? spend / t.engajamentos : 0)}`;
  const ctr = t.ctr != null ? t.ctr : (t.impressions ? (t.clicks / t.impressions * 100) : 0);
  const cpc = t.cpc != null ? t.cpc : (t.clicks ? spend / t.clicks : 0);
  return `Cliques ${t.clicks || 0} · CTR ${(ctr || 0).toFixed(2)}% · CPC ${_fmtR(cpc)}`;
}
// Restrição de conta de anúncio (espelha metaAcctStatusText/googleAcctStatusText do front)
function _metaRestr(st: any) { st = Number(st); if (st === 1 || st === 201) return null; if (st === 2) return "desativada pelo Meta"; if (st === 3) return "restrita por pagamento"; if (st === 9) return "em carência de pagamento"; if (st === 8) return "pendente de acerto de pagamento"; if (st === 7) return "em análise de risco/política"; if (st === 100) return "em encerramento"; if (st === 101 || st === 202) return "encerrada"; return `status atípico (código ${st})`; }
function _googleRestr(st: any) { st = String(st || "").toUpperCase(); if (st === "" || st === "ENABLED" || st === "UNSPECIFIED") return null; if (st === "SUSPENDED") return "suspensa pelo Google"; if (st === "CANCELED") return "cancelada"; if (st === "CLOSED") return "encerrada"; return `status atípico (${st})`; }
async function _waAccountRestrictions() {
  const [ma, ga] = await Promise.all([metaListAccounts().catch(() => []), googleListAccounts().catch(() => [])]);
  const m: Record<string, any> = {}; (ma || []).forEach((a: any) => { const t = _metaRestr(a.status); if (t) m[String(a.id).replace(/^act_/, "")] = { canal: "Meta", txt: t }; });
  const g: Record<string, any> = {}; (ga || []).forEach((a: any) => { const t = _googleRestr(a.status); if (t) g[String(a.id).replace(/-/g, "")] = { canal: "Google", txt: t }; });
  return { m, g };
}
function _clientRestrictions(c: any, restr: any): any[] {
  const out: any[] = [];
  String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean).forEach((id: string) => { const r = restr.m[id.replace(/^act_/, "")]; if (r) out.push(r); });
  String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean).forEach((id: string) => { const r = restr.g[id.replace(/-/g, "")]; if (r) out.push(r); });
  return out;
}
// ===== REGRA DURA: venda/faturamento vêm da PLANILHA (aba VENDAS do canal), não do pixel. Espelha o dashboard. =====
const REPORT_SYN: Record<string, string[]> = { date: ["data", "day"], sales: ["venda", "vendas"], revenue: ["faturamento"] };
function _normH(h: any) { return String(h || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim(); }
function _numBR(v: any) { if (v == null) return 0; let s = String(v).replace(/R\$/gi, "").trim(); if (!s || s === "-") return 0; s = s.replace(/\s/g, ""); if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", "."); else if (s.includes(",")) s = s.replace(",", "."); const n = parseFloat(s); return isNaN(n) ? 0 : n; }
function _dateFlex(v: any) { if (!v) return null; const s = String(v).trim(); let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`; m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/); if (m) return `${m[3]}-${m[2]}-${m[1]}`; return null; }
function _chanPlat(tab: string) { const t = String(tab || "").replace(/^GERAL\s+/i, "").trim().toUpperCase(); const m = t.match(/^[A-ZÀ-Ú0-9]+(?:-[A-ZÀ-Ú0-9]+)?/); return m ? m[0] : t; }
let _sheetsClient: any = null;
function _getSheets() { if (_sheetsClient) return _sheetsClient; const keyJson = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_KEY"); if (!keyJson) return null; const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(keyJson), scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"] }); _sheetsClient = google.sheets({ version: "v4", auth }); return _sheetsClient; }
// Venda REAL do canal (source: 'meta'|'google') vinda da planilha do cliente, no período. null = cliente não usa planilha p/ esse canal.
async function _waSheetSales(c: any, source: string, since: string, until: string): Promise<{ sales: number; revenue: number } | null> {
  const cs = c.conversion_source || ""; if (cs === "meta" || cs === "none") return null;
  const url = c.report_sheet_url || ""; if (!url) return null;
  const mm = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/); const sid = mm ? mm[1] : null; if (!sid) return null;
  const tabs = String(c.report_tabs || "").split(",").map((s: string) => s.trim()).filter(Boolean); if (!tabs.length) return null;
  const want = source === "google" ? "GOOGLE" : "META";
  const matchTabs = tabs.filter((t: string) => _chanPlat(t) === want); if (!matchTabs.length) return null;
  const sheets = _getSheets(); if (!sheets) return null;
  let sales = 0, revenue = 0, any = false;
  for (const tab of matchTabs) {
    try {
      const res = await sheets.spreadsheets.values.get({ spreadsheetId: sid, range: `'${tab}'!A1:Z5000` });
      const rows = res.data.values || []; if (rows.length < 2) continue;
      const header = rows[0].map(_normH); const idx: Record<string, number> = {};
      for (const f in REPORT_SYN) { const i = header.findIndex((h: string) => REPORT_SYN[f].includes(h)); if (i !== -1) idx[f] = i; }
      if (idx.sales == null && idx.revenue == null) continue; any = true;
      for (const row of rows.slice(1)) { const dt = idx.date != null ? _dateFlex(row[idx.date]) : null; if (!dt || dt < since || dt > until) continue; if (idx.sales != null) sales += _numBR(row[idx.sales]); if (idx.revenue != null) revenue += _numBR(row[idx.revenue]); }
    } catch (_e) { /* aba sem acesso/erro: ignora */ }
  }
  return any ? { sales, revenue } : null;
}
function _applySheet(t: any, sheet: any) { if (!t || !sheet) return t; return { ...t, purchases: sheet.sales, revenue: sheet.revenue, roas: t.spend ? sheet.revenue / t.spend : 0 }; }
const ESCOPO_LABEL: Record<string, string> = { padrao: "clientes com investimento", todos: "todos os clientes", ativos: "clientes ativos", ativos_sem_restricao: "ativos sem restrição", rodaram: "só os que rodaram", com_restricao: "com restrição de conta" };
// KPIs COMPLETOS de um canal (relatório "completo") — 1 por linha. A métrica de RESULTADO segue o OBJETIVO (obj), não a presença de valor.
function _waKpiFull(t: any, google: boolean, obj?: string | null): string[] {
  const L: string[] = [];
  L.push(`Gasto: ${_fmtR(t.spend || 0)}`);
  L.push(`Impressões: ${Math.round(t.impressions || 0).toLocaleString("pt-BR")}`);
  if (!google && (t.reach || 0) > 0) L.push(`Alcance: ${Math.round(t.reach).toLocaleString("pt-BR")}`);
  L.push(`Cliques: ${Math.round(t.clicks || 0).toLocaleString("pt-BR")}`);
  const ctr = t.ctr != null ? t.ctr : (t.impressions ? t.clicks / t.impressions * 100 : 0);
  const cpc = t.cpc != null ? t.cpc : (t.clicks ? t.spend / t.clicks : 0);
  const cpm = t.cpm != null ? t.cpm : (t.impressions ? t.spend / t.impressions * 1000 : 0);
  L.push(`CTR: ${(ctr || 0).toFixed(2)}%`); L.push(`CPC: ${_fmtR(cpc)}`); L.push(`CPM: ${_fmtR(cpm)}`);
  const isVenda = obj === "conversao" || obj === "app" || (!obj && (t.purchases || 0) > 0);
  const isLead = obj === "leads" || (!obj && !google && (t.leads || 0) > 0);
  const isMsg = obj === "mensagens" || (!obj && !google && (t.conversas || 0) > 0);
  const isVideo = obj === "video" || (!obj && (t.videoViews || 0) > 0);
  const isAlcance = obj === "alcance" || obj === "distribuicao";
  const isEngaj = obj === "engajamento";
  if (isVenda) { const roas = t.roas != null ? t.roas : (t.spend ? (t.revenue || 0) / t.spend : 0); L.push(`Compras: ${Math.round(t.purchases || 0)}`); L.push(`ROAS: ${(roas || 0).toFixed(2)}x`); L.push(`CPA: ${t.purchases ? _fmtR(t.spend / t.purchases) : "—"}`); if (t.revenue) L.push(`Receita: ${_fmtR(t.revenue)}`); }
  else if (isLead) { L.push(`Leads: ${t.leads || 0}`); L.push(`CPL: ${t.leads ? _fmtR(t.spend / t.leads) : "—"}`); }
  else if (isMsg) { L.push(`Conversas: ${t.conversas || 0}`); L.push(`Custo/conversa: ${t.conversas ? _fmtR(t.spend / t.conversas) : "—"}`); }
  else if (isVideo) { L.push(`Visualizações: ${Math.round(t.videoViews || 0).toLocaleString("pt-BR")}`); }
  else if (isAlcance) { if (!google && t.reach && t.impressions) L.push(`Frequência: ${(t.impressions / t.reach).toFixed(2)}`); }
  else if (isEngaj) { L.push(`Engajamentos: ${Math.round(t.engajamentos || 0).toLocaleString("pt-BR")}`); }
  return L;
}
// Playbook de inteligência: princípios embutidos + base de conhecimento da agência (agent_knowledge global).
const WA_PLAYBOOK_BASE = `PLAYBOOK DE INTELIGÊNCIA — como AVALIAR e ORIENTAR (siga sempre):
- Avalie SEMPRE pela métrica do OBJETIVO do canal. Nunca julgue por venda/ROAS quem não é venda.
- Custo por lead/conversa alto (ou CPA alto) NÃO é automaticamente ruim: pode ser um lead mais QUALIFICADO. Antes de dizer "reduzir custo", ORIENTE a VERIFICAR A QUALIFICAÇÃO dos leads/conversas (se estão virando reunião/venda no CRM). Se estão qualificados e fechando, o custo pode estar saudável.
- Lead barato porém desqualificado é PIOR que lead caro que fecha. Olhe qualificação antes de olhar custo.
- CTR/CPC baixos em campanha de ALCANCE/reconhecimento não são problema — o objetivo é impressão/alcance/frequência.
- Não recomende pausar/escalar/otimizar com base numa métrica isolada; baseie-se no resultado do objetivo.
- Oriente o próximo passo concreto (ex: "checar a qualificação das X conversas no CRM antes de mexer no orçamento").`;
let _waPbCache: string | null = null, _waPbT = 0;
async function _waPlaybook(): Promise<string> {
  if (_waPbCache && Date.now() - _waPbT < 300000) return _waPbCache;
  let extra = "";
  try { const rows = await sbGet("agent_knowledge", "select=title,text,client_id&order=created_at.desc&limit=20"); const g = (rows || []).filter((r: any) => !r.client_id).slice(0, 3); if (g.length) extra = "\n\nMÉTODOS DA AGÊNCIA (base de conhecimento):\n" + g.map((k: any) => `- ${k.title}: ${String(k.text || "").slice(0, 2500)}`).join("\n"); } catch (_e) { /* */ }
  _waPbCache = WA_PLAYBOOK_BASE + extra; _waPbT = Date.now(); return _waPbCache;
}
// Análise do gestor POR CANAL, cada um julgado pelo SEU objetivo. Retorna cliente -> texto (1 linha por canal).
async function _waAnalises(items: any[]): Promise<Record<string, string>> {
  try {
    const data = items.map((r: any) => {
      const canais: any[] = [];
      if (r.meta && (r.meta.spend || 0) > 0) canais.push({ canal: "Meta", objetivo: _objLabel(r.objMeta), metricas: _objMetric(r.meta, false, r.objMeta), gasto: Math.round(r.meta.spend) });
      if (r.google && (r.google.spend || 0) > 0) canais.push({ canal: "Google", objetivo: _objLabel(r.objGoogle), metricas: _objMetric(r.google, true, r.objGoogle), gasto: Math.round(r.google.spend) });
      return { cliente: r.nome, canais };
    });
    const pb = await _waPlaybook();
    const sys = `${pb}\n\nVocê é a AndréIA, gestora de tráfego sênior. Analise CADA CLIENTE e, dentro dele, CADA CANAL SEPARADAMENTE — julgando pelo OBJETIVO daquele canal (o campo "objetivo"), SEMPRE seguindo o playbook acima (ex: custo alto → orientar a verificar qualificação, não só "reduzir custo"):
- venda: avalie ROAS/CPA/faturamento.
- leads: avalie quantidade de leads e CPL.
- mensagens: avalie conversas e custo por conversa.
- tráfego: avalie cliques, CPC e CTR.
- alcance: avalie alcance, CPM e frequência — NÃO fale de cliques/CTR nem de conversões/vendas.
- engajamento: avalie engajamentos e custo por engajamento.
- vídeo: avalie visualizações e custo por view.
REGRAS: nunca cite venda/conversão/ROAS se o objetivo não for venda. Nunca cite CTR/cliques se o objetivo for alcance. Uma frase curta (máx ~16 palavras) POR CANAL, dizendo o que está bom/ruim e o próximo passo. Se houver 2 canais, dê uma frase pra cada.
Responda em JSON: {"analises":[{"cliente":"nome exato","linhas":[{"canal":"Meta","texto":"..."},{"canal":"Google","texto":"..."}]}]}`;
    const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content: JSON.stringify(data) }], response_format: { type: "json_object" }, max_tokens: 1600, temperature: 0.4 });
    const parsed = JSON.parse(j.choices[0].message.content || "{}");
    const map: Record<string, string> = {};
    (parsed.analises || []).forEach((a: any) => {
      if (!a || !a.cliente) return;
      if (Array.isArray(a.linhas) && a.linhas.length) map[a.cliente] = a.linhas.filter((l: any) => l && l.texto).map((l: any) => `• *${l.canal}:* ${l.texto}`).join("\n");
      else if (a.texto) map[a.cliente] = a.texto;
    });
    return map;
  } catch (_e) { return {}; }
}
// Resumo de clientes no período (por cliente, Meta/Google separados). `escopo` filtra QUEM entra; `nivel` = resumido|completo. Retorna mensagens (chunked).
async function waAgentAllClientsSummary(days: number, escopo = "padrao", nivel = "resumido"): Promise<string[]> {
  const since = new Date(Date.now() - days * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
  const clients = await sbGet("clients", "select=id,name,meta_account_id,google_account_id,status,conversion_source,report_sheet_url,report_tabs&limit=500");
  const withAcct = clients.filter((c: any) => String(c.meta_account_id || "").trim() || String(c.google_account_id || "").trim());
  const needRestr = escopo === "ativos_sem_restricao" || escopo === "com_restricao";
  const restr = needRestr ? await _waAccountRestrictions() : { m: {}, g: {} };
  const isAtivo = (c: any) => c.status === "Ativo";
  let base = withAcct.filter((c: any) => c.status !== "Encerrado");
  if (escopo === "ativos" || escopo === "ativos_sem_restricao" || escopo === "rodaram") base = base.filter(isAtivo);
  if (escopo === "ativos_sem_restricao") base = base.filter((c: any) => _clientRestrictions(c, restr).length === 0);
  if (escopo === "com_restricao") base = base.filter((c: any) => _clientRestrictions(c, restr).length > 0);
  const showNon = escopo === "todos" || escopo === "ativos" || escopo === "ativos_sem_restricao";
  const results: any[] = [];
  for (let i = 0; i < base.length; i += 8) {
    const ch = base.slice(i, i + 8);
    const rs = await Promise.all(ch.map(async (c: any) => {
      const mIds = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
      const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
      const [m, g] = await Promise.all([
        mIds.length ? metaAdsInsights({ accounts: mIds.map((id: string) => ({ id, name: id })), since, until, byCampaign: true }).catch(() => null) : Promise.resolve(null),
        gIds.length ? googleAdsInsights({ accounts: gIds.map((id: string) => ({ id, name: id })), since, until, byCampaign: true }).catch(() => null) : Promise.resolve(null),
      ]);
      let mt = (m && m.total) || null, gt = (g && g.total) || null;
      let objMeta = _domObj(m), objGoogle = _domObj(g);
      // REGRA: venda vem da planilha (se o canal tem aba VENDAS). Sobrepõe o pixel e força a métrica de venda.
      if (mt) { const sh = await _waSheetSales(c, "meta", since, until); if (sh) { mt = _applySheet(mt, sh); objMeta = "conversao"; } }
      if (gt) { const sh = await _waSheetSales(c, "google", since, until); if (sh) { gt = _applySheet(gt, sh); objGoogle = "conversao"; } }
      return { nome: c.name, meta: mt, google: gt, objMeta, objGoogle, restr: escopo === "com_restricao" ? _clientRestrictions(c, restr) : [] };
    }));
    results.push(...rs);
  }
  results.sort((a, b) => ((b.meta?.spend || 0) + (b.google?.spend || 0)) - ((a.meta?.spend || 0) + (a.google?.spend || 0)));
  const completo = nivel === "completo" && escopo !== "com_restricao";
  const analises = completo ? await _waAnalises(results.filter((r: any) => (r.meta?.spend || 0) > 0 || (r.google?.spend || 0) > 0)) : {};
  const blocks: string[] = [];
  for (const r of results) {
    const mS = r.meta?.spend || 0, gS = r.google?.spend || 0, ran = mS > 0 || gS > 0;
    if (escopo === "com_restricao") {
      const rt = r.restr.map((x: any) => `${x.canal}: ${x.txt}`).join(" · ");
      let b = `*${r.nome}* — 🚫 ${rt}`;
      if (ran) { if (mS > 0) b += `\n📘 Meta — Gasto ${_fmtR(mS)} · ${_objMetric(r.meta, false, r.objMeta)}`; if (gS > 0) b += `\n🔎 Google — Gasto ${_fmtR(gS)} · ${_objMetric(r.google, true, r.objGoogle)}`; }
      else b += `\n⏸ não rodou no período`;
      blocks.push(b); continue;
    }
    if (!ran) { if (!showNon) continue; blocks.push(`*${r.nome}* — ⏸ não rodou no período`); continue; }
    let b = completo ? `━━━━━━━━━━━━━━━\n📊 *${r.nome.toUpperCase()}*` : `*${r.nome}*`;
    if (completo) {
      if (mS > 0) b += `\n\n📘 *META*\n${_waKpiFull(r.meta, false, r.objMeta).map((l: string) => `• ${l}`).join("\n")}`;
      if (gS > 0) b += `\n\n🔎 *GOOGLE*\n${_waKpiFull(r.google, true, r.objGoogle).map((l: string) => `• ${l}`).join("\n")}`;
      if (mS > 0 && gS > 0) b += `\n\n💰 *Total investido:* ${_fmtR(mS + gS)}`;
      if (analises[r.nome]) b += `\n\n💬 *Análise*\n${analises[r.nome]}`;
    } else {
      if (mS > 0) b += `\n📘 Meta — Gasto ${_fmtR(mS)} · ${_objMetric(r.meta, false, r.objMeta)}`;
      if (gS > 0) b += `\n🔎 Google — Gasto ${_fmtR(gS)} · ${_objMetric(r.google, true, r.objGoogle)}`;
      if (mS > 0 && gS > 0) b += `\n➕ Total — Gasto ${_fmtR(mS + gS)}`;
    }
    blocks.push(b);
  }
  const escLbl = ESCOPO_LABEL[escopo] || ESCOPO_LABEL.padrao;
  if (!blocks.length) return [escopo === "com_restricao" ? `✅ Nenhum cliente com restrição de conta.` : `Nenhum cliente (${escLbl}) com dados nos últimos ${days} dias.`];
  const cab = escopo === "com_restricao" ? `🚫 *Clientes com restrição de conta*\n_últimos ${days} dias_\n${WA_DIV}\n` : `📊 *Resumo ${completo ? "completo " : ""}— últimos ${days} dias*\n_${escLbl}_\n${WA_DIV}\n`;
  const msgs: string[] = []; let cur = cab;
  for (const b of blocks) { if ((cur + "\n\n" + b).length > 3000) { msgs.push(cur); cur = b; } else cur += "\n\n" + b; }
  if (cur.trim()) msgs.push(cur);
  return msgs;
}
// Tabelas que a AndréIA pode CONSULTAR (só leitura) + o que cada uma guarda
const WA_TABLES: Record<string, string> = {
  clients: "clientes (id, name, seg, status, fee, billing, category, meta_account_id, google_account_id, conversion_source, day)",
  finance: "lançamentos financeiros — aba Financeiro (type=receita|despesa, client=id do cliente, description, val, due=YYYY-MM-DD, status=pendente|pago, category, creditor)",
  tasks: "tarefas (name, client=id, owner, status=todo|doing|done, due, prio, notes)",
  wa_conversations: "conversas do CRM WhatsApp (client_id, chat_id=telefone, name, stage, origin_type, origin jsonb, fields jsonb, last_at, last_text)",
  wa_messages: "mensagens do WhatsApp (conversation_id, chat_id, direction=in|out, text, ts)",
  rd_conversions: "conversões RD Station (client, email, source, medium, campaign, converted_at)",
  order_aggregates: "pedidos por dia (client_id, date, status, count, total)",
  capi_events: "eventos CAPI enviados pro Meta (client_id, event_name, status, error, created_at)",
  track_events: "eventos do pixel de rastreamento (client_id, type)",
  track_links: "links rastreáveis (client_id, slug, kind)",
  report_analysis: "análises de relatório salvas (client_id, month, text)",
  creditors: "credores/fornecedores (id, name)",
  wallet: "carteira (client, type, description, val, date)",
  checkout_events: "checkouts (client_id, event_date)",
  notifications: "notificações internas da equipe (to_team, task_name, comment_text, read, type, created_at)",
};
const WA_TOOLS = [
  { type: "function", function: { name: "consultar_banco", description: "Consulta SOMENTE LEITURA de qualquer tabela do sistema pra buscar dados reais (cliente, financeiro, tarefas, conversas do CRM, RD, pedidos etc). SEMPRE use antes de responder sobre dados guardados.", parameters: { type: "object", properties: { tabela: { type: "string", enum: Object.keys(WA_TABLES) }, colunas: { type: "string", description: "colunas separadas por vírgula ou '*'" }, filtro: { type: "string", description: "filtro no formato PostgREST, ex: 'client=eq.<id>&status=eq.pendente'; datas: 'due=gte.2026-07-01&due=lte.2026-07-31'; texto: 'description=ilike.*fee*'. Vazio = sem filtro." }, ordenar: { type: "string", description: "ex: 'created_at.desc' ou 'due.asc'" }, limite: { type: "integer" } }, required: ["tabela"] } } },
  { type: "function", function: { name: "meta_insights", description: "Métricas de Meta Ads AO VIVO de UM cliente no período: 'total' (consolidado, já filtrado pela métrica do objetivo) e 'campanhas' (CADA campanha com gasto, orçamento diário e os KPIs do objetivo DELA). Use pra 'como tá o cliente X', 'campanhas do X', detalhes por campanha.", parameters: { type: "object", properties: { cliente: { type: "string", description: "nome do cliente" }, dias: { type: "integer", description: "7, 30 ou 90 (padrão 7)" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "google_insights", description: "Métricas de Google Ads AO VIVO de UM cliente no período.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "google_keywords", description: "Palavras-chave e termos de busca do Google Ads de UM cliente no período (por palavra: gasto, cliques, conversões, CPC). USE isto quando perguntarem 'como está cada palavra-chave', keywords, termos de busca ou o que as pessoas pesquisaram.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer", description: "7, 30 ou 90 (padrão 7)" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "resumo_todos_clientes", description: "Resumo de TODOS os clientes no período (gasto + métrica do objetivo, Meta/Google separados). Use quando pedirem panorama/todos os clientes.", parameters: { type: "object", properties: { dias: { type: "integer" } } } } },
  { type: "function", function: { name: "relatorio_cliente", description: "Gera um RELATÓRIO VISUAL e limpo de UM cliente PRONTO PRA ENVIAR AO CLIENTE (investimento, resultados pelo objetivo, alcance e uma análise). Use quando pedirem 'relatório do [cliente]', 'manda o relatório pro cliente', 'relatório pra enviar'. Envie o campo 'relatorio' EXATAMENTE como vier.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer", description: "7, 30 ou 90 (padrão 7)" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "reunioes", description: "REUNIÕES/compromissos da AGENDA (Google Agenda), que é DIFERENTE de tarefa operacional. USE isto quando perguntarem sobre reuniões, agenda, compromissos, calls. NÃO liste tarefas comuns aqui.", parameters: { type: "object", properties: { quando: { type: "string", description: "'hoje', 'amanha', 'semana' ou vazio (padrão hoje)" }, data: { type: "string", description: "data específica AAAA-MM-DD (opcional)" } } } } },
  { type: "function", function: { name: "financeiro", description: "Consulta financeira com TOTAL e itens já com o nome do cliente resolvido e a soma correta. USE ISSO pra qualquer pergunta de dinheiro (a receber, a pagar, recebido, pago, fluxo do mês).", parameters: { type: "object", properties: { tipo: { type: "string", enum: ["receita", "despesa"] }, status: { type: "string", enum: ["pendente", "pago"] }, mes: { type: "string", description: "AAAA-MM, ex: 2026-07" }, cliente: { type: "string" } } } } },
  { type: "function", function: { name: "preparar_acao", description: "Prepara uma AÇÃO de alto impacto pra CONFIRMAÇÃO (NÃO executa agora — o sistema pede SIM). Para criar_tarefa, o RESPONSÁVEL (quem faz) e o QUANDO (data) são obrigatórios — se o usuário não disser, PERGUNTE antes.", parameters: { type: "object", properties: { tipo: { type: "string", enum: ["criar_tarefa", "criar_reuniao", "cancelar_reuniao", "pausar_campanha", "reativar_campanha", "orcamento", "duplicar_campanha", "criar_lancamento", "dar_baixa"] }, cliente: { type: "string" }, nome: { type: "string", description: "título da tarefa OU da reunião (pra cancelar_reuniao, o título/pedaço do nome da reunião a cancelar). NÃO inclua 'urgente' nem 'revisão' no título — use os campos próprios." }, responsavel: { type: "string", description: "nome de quem vai fazer a tarefa (membro da equipe)" }, quando: { type: "string", description: "data em AAAA-MM-DD (calcule 'amanhã', 'sexta' etc. a partir de hoje) — usada por tarefa e reunião" }, hora: { type: "string", description: "horário da reunião em HH:MM (opcional)" }, urgente: { type: "boolean", description: "true se a tarefa foi pedida como URGENTE — marca a flag de urgência (NÃO escreva 'urgente' no título/obs)" }, revisao: { type: "boolean", description: "true se pediram para SOLICITAR REVISÃO da tarefa — marca a flag de revisão (NÃO escreva 'revisão' no título/obs)" }, obs: { type: "string" }, campanha: { type: "string" }, novoValor: { type: "number" }, natureza: { type: "string", enum: ["receita", "despesa"] }, descricao: { type: "string" }, valor: { type: "number" }, vencimento: { type: "string" } }, required: ["tipo"] } } },
];
const WA_MENU_TEXT = `🤖 *AndréIA — o que posso fazer aqui no grupo:*

📊 *Análise*
• _Como tá o [cliente] nos últimos 7 dias?_
• _Detalhes das campanhas do [cliente]_
• _Como está cada palavra-chave do [cliente]?_ (Google)
• _Relatório do [cliente] pra enviar_ (layout pronto pro cliente)
• _Resumo de todos os clientes_
• _Quem precisa de atenção?_
• _Saúde da carteira_
• _Recomendações da semana_

🗓 *Agenda*
• _Quais reuniões tenho hoje?_ / _amanhã?_
• _Minhas reuniões da semana_
• _Marca reunião com [cliente] sexta às 15h_ (peço confirmação)
• _Cancela a reunião [nome]_ (peço confirmação)

💰 *Financeiro*
• _Quanto temos a receber esse mês?_ / _a pagar?_
• _Cria um lançamento…_ / _Dá baixa em…_

✅ *Operacional*
• _Pendências operacionais_ / _Tarefas em aberto do [cliente]_
• _Cria uma tarefa pro [responsável] em [cliente] pra [data]: …_

⚙️ *Campanhas* (peço confirmação antes)
• _Pausa / reativa / duplica a campanha [nome]_
• _Sobe o orçamento da [nome] pra R$ X_

É só mandar em linguagem natural. 💬 Mande *menu* pra ver isso de novo.`;
function _waResolveClient(nomeOuId: string, clients: any[]) { if (!nomeOuId) return null; const q = String(nomeOuId).toLowerCase().trim(); return clients.find((c) => c.id === nomeOuId) || clients.find((c) => c.name.toLowerCase() === q) || clients.find((c) => c.name.toLowerCase().includes(q)) || null; }
let _waCliMap: Record<string, string> | null = null, _waCliMapT = 0;
async function _waClientsMap(): Promise<Record<string, string>> {
  if (_waCliMap && Date.now() - _waCliMapT < 300000) return _waCliMap;
  const cs = await sbGet("clients", "select=id,name&limit=1000"); const m: Record<string, string> = {};
  cs.forEach((c: any) => { m[c.id] = c.name; }); _waCliMap = m; _waCliMapT = Date.now(); return m;
}
async function waQueryTable(args: any) {
  const t = args.tabela; if (!WA_TABLES[t]) return { erro: "tabela não permitida" };
  const p = ["select=" + encodeURIComponent(args.colunas && String(args.colunas).trim() ? args.colunas : "*")];
  if (args.filtro && String(args.filtro).trim()) p.push(String(args.filtro).trim());
  if (args.ordenar) p.push("order=" + encodeURIComponent(args.ordenar));
  p.push("limit=" + Math.min(Number(args.limite) || 30, 100));
  try {
    const rows = await sbGet(t, p.join("&"));
    if (rows.length && (rows[0].client !== undefined || rows[0].client_id !== undefined)) { const map = await _waClientsMap(); rows.forEach((r: any) => { const cid = r.client || r.client_id; if (cid && map[cid]) r.cliente_nome = map[cid]; }); }
    return { linhas: rows, total: rows.length };
  } catch (e) { return { erro: String((e as any)?.message || e) }; }
}
// Reuniões da agenda (Google Agenda) — tarefas sincronizadas (id 'cal*', nota "Reunião (Google Agenda)"). Diferente de tarefa operacional.
async function waReunioes(args: any) {
  const now = new Date(Date.now() - 3 * 3600e3); const ymd = (d: Date) => d.toISOString().slice(0, 10);
  let since: string, until: string;
  const q = String(args?.quando || "").toLowerCase();
  if (args?.data) { since = until = String(args.data).slice(0, 10); }
  else if (q.includes("amanh")) { const d = new Date(now.getTime() + 864e5); since = until = ymd(d); }
  else if (q.includes("semana")) { since = ymd(now); until = ymd(new Date(now.getTime() + 7 * 864e5)); }
  else { since = until = ymd(now); }
  const rows = await sbGet("tasks", `id=like.cal*&status=neq.done&due=gte.${since}&due=lte.${until}&select=name,client,due,notes,link&order=due.asc&limit=100`);
  const map = await _waClientsMap();
  const reunioes = rows.map((r: any) => {
    const hm = (String(r.notes || "").match(/(\d{2}:\d{2})/) || [])[1] || "";
    const link = (String(r.notes || "").match(/https?:\/\/\S+/) || [])[0] || r.link || "";
    return { titulo: r.name, data: r.due, hora: hm, cliente: (r.client && map[r.client]) ? map[r.client] : null, link: link || null };
  });
  return { de: since, ate: until, quantidade: reunioes.length, reunioes };
}
// Financeiro determinístico: total + itens (com nome do cliente já resolvido). Evita o modelo errar nome/soma.
async function waFinanceiro(args: any) {
  const p = ["select=type,status,client,description,val,due", "limit=1000"];
  if (args.tipo) p.push("type=eq." + args.tipo);
  if (args.status) p.push("status=eq." + args.status);
  if (args.mes) p.push("due=like." + String(args.mes) + "*");
  if (args.cliente) { const clients = await sbGet("clients", "select=id,name&limit=1000"); const rc = _waResolveClient(args.cliente, clients); if (!rc) return { erro: `cliente "${args.cliente}" não encontrado` }; p.push("client=eq." + rc.id); }
  const rows = await sbGet("finance", p.join("&"));
  const map = await _waClientsMap();
  const itens = rows.map((r: any) => ({ cliente: map[r.client] || "(sem cliente)", descricao: r.description, valor: Number(r.val) || 0, vencimento: r.due, status: r.status, tipo: r.type }));
  const total = Math.round(itens.reduce((s: number, x: any) => s + x.valor, 0) * 100) / 100;
  return { total, quantidade: itens.length, itens };
}
// KPI de UMA campanha, pela métrica do objetivo DELA (venda→compras/ROAS; leads→CPL; mensagem→custo/conversa; senão tráfego)
function _waCampKpi(c: any) {
  const spend = c.spend || 0; const tipo = (c.objetivo && c.objetivo.tipo) || "";
  const base: any = { gasto: Math.round(spend), ctr: +(c.ctr || 0).toFixed(2), cpc: +(c.cpc || 0).toFixed(2) };
  if (tipo === "conversao" || tipo === "app") { base.compras = Math.round(c.purchases || 0); base.roas = c.roas != null ? +c.roas.toFixed(2) : (spend ? +((c.revenue || 0) / spend).toFixed(2) : 0); base.cpa = c.purchases ? +(spend / c.purchases).toFixed(2) : null; }
  else if (tipo === "leads") { base.leads = c.leads || 0; base.cpl = c.leads ? +(spend / c.leads).toFixed(2) : null; }
  else if (tipo === "mensagens") { base.conversas = c.conversas || 0; base.custoPorConversa = c.conversas ? +(spend / c.conversas).toFixed(2) : null; }
  else if (tipo === "video") { base.videoViews = c.videoViews || 0; }
  else { base.cliques = c.clicks || 0; }
  return base;
}
// Resumo de UM cliente com KPIs POR CAMPANHA (cada uma pela métrica do objetivo dela) + total filtrado, pro período pedido
async function waMetaResumo(clientId: string, dias: number) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,meta_account_id`))[0];
  if (!c) return { erro: "cliente não encontrado" };
  const ids = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) return { cliente: c.name, aviso: "cliente sem Meta Ads vinculado" };
  const accounts = ids.map((id: string) => ({ id, name: id }));
  const since = new Date(Date.now() - dias * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
  const [r, ent] = await Promise.all([
    metaAdsInsights({ accounts, since, until, byCampaign: true }).catch(() => null),
    metaEntities({ accounts }).catch(() => null),
  ]);
  const total = (r && r.total) ? _waResumoMeta(r.total) : null;
  const budgetByName: Record<string, number> = {};
  if (ent) (ent.campaigns || []).forEach((x: any) => { if (x.status === "ACTIVE" || x.entrega === "ACTIVE") budgetByName[x.nome] = x.orcamentoDiario; });
  const campanhas = ((r && r.campaigns) || []).filter((x: any) => (x.spend || 0) > 0).slice(0, 20).map((x: any) => ({ nome: x.campaign, objetivo: (x.objetivo && x.objetivo.rotulo) || "", orcamentoDiario: budgetByName[x.campaign] || undefined, ..._waCampKpi(x) }));
  return { cliente: c.name, dias, total, campanhas };
}
// Relatório VISUAL de UM cliente, pronto pra enviar pro cliente (layout limpo pro WhatsApp).
function _fmtN(v: number) { return Math.round(v || 0).toLocaleString("pt-BR"); }
function _brDia(iso: string) { const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}` : iso; }
async function _waAnaliseCliente(nome: string, mt: any, gt: any, objM?: string | null, objG?: string | null) {
  try {
    const data: any = { cliente: nome };
    if (mt) data.meta = { objetivo: _objLabel(objM), metricas: _objMetric(mt, false, objM) };
    if (gt) data.google = { objetivo: _objLabel(objG), metricas: _objMetric(gt, true, objG) };
    const pb = await _waPlaybook();
    const sys = `${pb}\n\nVocê é a AndréIA, gestora de tráfego da GT Marketing, escrevendo PARA O CLIENTE. Em 1 ou 2 frases curtas, profissionais e claras (sem jargão técnico, sem 'pausar/escalar/otimizar'), resuma o desempenho de forma honesta e positiva, SEMPRE pelo OBJETIVO de cada canal (o campo 'objetivo') e seguindo o playbook acima. Se houver dois canais com objetivos diferentes, comente cada um pelo seu objetivo. NUNCA cite venda/conversão se o objetivo não for venda; NUNCA cite CTR/cliques se o objetivo for alcance. Não use markdown nem emojis.`;
    const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content: JSON.stringify(data) }], max_tokens: 180, temperature: 0.5 });
    return (j.choices[0].message.content || "").trim();
  } catch (_e) { return ""; }
}
async function waRelatorioCliente(c: any, dias: number, comAnalise = true): Promise<string> {
  const since = new Date(Date.now() - dias * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
  const mIds = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  const [m, g] = await Promise.all([
    mIds.length ? metaAdsInsights({ accounts: mIds.map((id: string) => ({ id, name: id })), since, until, byCampaign: true }).catch(() => null) : Promise.resolve(null),
    gIds.length ? googleAdsInsights({ accounts: gIds.map((id: string) => ({ id, name: id })), since, until, byCampaign: true }).catch(() => null) : Promise.resolve(null),
  ]);
  let mt = m && m.total && (m.total.spend || 0) > 0 ? m.total : null;
  let gt = g && g.total && (g.total.spend || 0) > 0 ? g.total : null;
  let objM = _domObj(m), objG = _domObj(g);
  // REGRA: venda/faturamento da PLANILHA (aba VENDAS do canal) sobrepõem o pixel.
  if (mt) { const sh = await _waSheetSales(c, "meta", since, until); if (sh) { mt = _applySheet(mt, sh); objM = "conversao"; } }
  if (gt) { const sh = await _waSheetSales(c, "google", since, until); if (sh) { gt = _applySheet(gt, sh); objG = "conversao"; } }
  // objetivo dominante do cliente (canal que mais gastou decide)
  const obj = ((mt?.spend || 0) >= (gt?.spend || 0)) ? (objM || objG) : (objG || objM);
  const DIV = "━━━━━━━━━━━━━━━";
  if (!mt && !gt) return `📊 *RELATÓRIO — ${c.name}*\n${DIV}\nSem investimento no período (${_brDia(since)}–${_brDia(until)}).`;
  const tot = { spend: 0, impressions: 0, clicks: 0, reach: 0, purchases: 0, revenue: 0, leads: 0, conversas: 0, videoViews: 0 };
  [mt, gt].forEach((t: any) => { if (!t) return; tot.spend += t.spend || 0; tot.impressions += t.impressions || 0; tot.clicks += t.clicks || 0; tot.reach += (t.reach || 0); tot.purchases += t.purchases || 0; tot.revenue += t.revenue || 0; tot.leads += t.leads || 0; tot.conversas += t.conversas || 0; tot.videoViews += t.videoViews || 0; });
  let s = `📊 *RELATÓRIO DE PERFORMANCE*\n👤 *${c.name}*\n📅 ${_brDia(since)} a ${_brDia(until)} (${dias} dias)\n${DIV}\n`;
  // Investimento
  s += `\n💰 *Investimento*\n`;
  if (mt && gt) { s += `• Meta: ${_fmtR(mt.spend)}\n• Google: ${_fmtR(gt.spend)}\n• *Total: ${_fmtR(tot.spend)}*\n`; }
  else s += `• *Total: ${_fmtR(tot.spend)}*\n`;
  // Resultados — SEGUEM O OBJETIVO do cliente (não a presença de valor). Fallback = presença quando não há objetivo.
  const res: string[] = [];
  const isVenda = obj === "conversao" || obj === "app" || (!obj && tot.purchases > 0);
  const isLead = obj === "leads" || (!obj && tot.leads > 0);
  const isMsg = obj === "mensagens" || (!obj && tot.conversas > 0);
  const isVideo = obj === "video" || (!obj && tot.videoViews > 0);
  if (isVenda) { const roas = tot.spend ? tot.revenue / tot.spend : 0; res.push(`• Vendas: *${_fmtN(tot.purchases)}*`); if (tot.revenue) res.push(`• Faturamento: *${_fmtR(tot.revenue)}*`); res.push(`• ROAS: *${roas.toFixed(2)}x*`); res.push(`• Custo por venda: ${tot.purchases ? _fmtR(tot.spend / tot.purchases) : "—"}`); }
  else if (isLead) { res.push(`• Leads: *${_fmtN(tot.leads)}*`); res.push(`• Custo por lead: ${tot.leads ? _fmtR(tot.spend / tot.leads) : "—"}`); }
  else if (isMsg) { res.push(`• Conversas: *${_fmtN(tot.conversas)}*`); res.push(`• Custo por conversa: ${tot.conversas ? _fmtR(tot.spend / tot.conversas) : "—"}`); }
  else if (isVideo) { res.push(`• Visualizações: *${_fmtN(tot.videoViews)}*`); res.push(`• Custo por view: ${tot.videoViews ? _fmtR(tot.spend / tot.videoViews) : "—"}`); }
  if (res.length) s += `\n🎯 *Resultados*\n${res.join("\n")}\n`;
  // Alcance / tráfego
  const ctr = tot.impressions ? tot.clicks / tot.impressions * 100 : 0;
  s += `\n📈 *Alcance*\n• Impressões: ${_fmtN(tot.impressions)}\n`;
  if (tot.reach > 0) s += `• Pessoas alcançadas: ${_fmtN(tot.reach)}\n`;
  s += `• Cliques: ${_fmtN(tot.clicks)} · CTR ${ctr.toFixed(2)}%\n`;
  // Análise
  if (comAnalise) { const a = await _waAnaliseCliente(c.name, mt, gt, objM, objG); if (a) s += `\n💬 ${a}\n`; }
  s += `${DIV}\n_GT Marketing • Gestão de Tráfego_`;
  return s;
}
async function waExecTool(name: string, args: any, clients: any[]) {
  if (name === "consultar_banco") return await waQueryTable(args);
  if (name === "relatorio_cliente") { const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" }; const rep = await waRelatorioCliente(c, Number(args.dias) || 7); return { _cid: c.id, relatorio: rep, instrucao: "Envie o campo 'relatorio' EXATAMENTE como está, sem reescrever nem resumir." }; }
  if (name === "financeiro") return await waFinanceiro(args);
  if (name === "reunioes") return await waReunioes(args);
  if (name === "resumo_todos_clientes") { const msgs = await waAgentAllClientsSummary(Number(args.dias) || 7); return { texto: msgs.join("\n\n") }; }
  if (name === "meta_insights") { const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" }; const r = await waMetaResumo(c.id, Number(args.dias) || 7); return { _cid: c.id, ...r }; }
  if (name === "google_insights") {
    const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" };
    const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean); if (!gIds.length) return { cliente: c.name, aviso: "cliente sem Google Ads vinculado" };
    const d = Number(args.dias) || 30; const since = new Date(Date.now() - d * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
    const r = await googleAdsInsights({ accounts: gIds.map((id: string) => ({ id, name: id })), since, until }).catch((e: any) => ({ erro: String(e?.message || e) }));
    return { cliente: c.name, _cid: c.id, dias: d, total: r && (r as any).total };
  }
  if (name === "google_keywords") {
    const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" };
    const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean); if (!gIds.length) return { cliente: c.name, aviso: "cliente sem Google Ads vinculado" };
    const dd = Number(args.dias) || 7; const since = new Date(Date.now() - dd * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
    const r: any = await googleBreakdowns({ accounts: gIds.map((id: string) => ({ id })), since, until }).catch((e: any) => ({ erro: String(e?.message || e) }));
    if (r.erro) return { cliente: c.name, dias: dd, erro: r.erro };
    const fmt = (arr: any[]) => (arr || []).filter((k: any) => (k.spend || 0) > 0 || (k.clicks || 0) > 0).slice(0, 15).map((k: any) => ({ termo: k.key, gasto: Math.round(k.spend), cliques: k.clicks, conversoes: +(k.conversions || 0).toFixed(1), cpc: k.clicks ? +(k.spend / k.clicks).toFixed(2) : 0 }));
    const palavrasChave = fmt(r.keywords), termosDeBusca = fmt(r.termos);
    if (!palavrasChave.length && !termosDeBusca.length) return { cliente: c.name, dias: dd, aviso: "sem palavras-chave com dados no período (o cliente pode não estar rodando Rede de Pesquisa, ou não teve impressões)" };
    return { cliente: c.name, dias: dd, palavrasChave, termosDeBusca };
  }
  return { erro: "ferramenta desconhecida" };
}
function _waConfirmText(p: any, clients: any[]) {
  const cn = (clients.find((c) => c.id === p.client_id) || {}).name || ""; const cli = cn ? `📌 Cliente: ${cn}\n` : "";
  if (p.tipo === "criar_tarefa") return `${cli}Crio a tarefa "${p.nome || ""}"${p.obs ? ` (${p.obs})` : ""}${p.responsavel ? ` — resp. *${p.responsavel}*` : ""}${p._due ? ` — pra *${p._due}*` : ""}${p.urgente ? " — 🔴 *URGENTE*" : ""}${p.revisao ? " — 🔎 *com revisão*" : ""}. Confirma? (responda SIM)`;
  if (p.tipo === "criar_reuniao") return `${cli}Crio a reunião "${p.nome || ""}" na Google Agenda${p._due ? ` em *${p._due.split("-").reverse().join("/")}*` : ""}${p.hora ? ` às *${p.hora}*` : ""}. Confirma? (responda SIM)`;
  if (p.tipo === "cancelar_reuniao") return `${cli}Cancelar a reunião "${p.nome || ""}"${p._due ? ` de *${p._due.split("-").reverse().join("/")}*` : ""} (apaga também na Google Agenda). Confirma? (responda SIM)`;
  if (p.tipo === "pausar_campanha") return `${cli}Pausar a campanha "${p.campanha || ""}". Confirma?`;
  if (p.tipo === "reativar_campanha") return `${cli}Reativar a campanha "${p.campanha || ""}". Confirma?`;
  if (p.tipo === "orcamento") return `${cli}Ajustar o orçamento diário da "${p.campanha || ""}" pra R$${p.novoValor}. Confirma?`;
  if (p.tipo === "duplicar_campanha") return `${cli}Duplicar a campanha "${p.campanha || ""}" (cópia pausada). Confirma?`;
  if (p.tipo === "criar_lancamento") return `${cli}Criar lançamento ${p.natureza || "receita"} de R$${p.valor} — ${p.descricao || ""} (venc. ${p.vencimento || "hoje"}). Confirma?`;
  if (p.tipo === "dar_baixa") return `${cli}Dar baixa (marcar como pago) no lançamento "${p.descricao || ""}". Confirma?`;
  return `${cli}Confirma a ação?`;
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
  // menu de comandos
  if (/^(menu|ajuda|comandos|opções|opcoes|\?|o que voce faz|o que você faz|help)[.!?]*$/i.test(text.trim())) {
    await send(WA_MENU_TEXT);
    await saveSess({ last_msgid: w.msgid, pending: null, history: [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: "[menu enviado]" }].slice(-16) });
    return { ok: true };
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
  // relatório de UM cliente pra enviar — atalho (garante o layout limpo verbatim)
  if (/\brelat[óo]rio/.test(low) && !/clientes/.test(low)) {
    const clientsR = await sbGet("clients", "select=id,name,meta_account_id,google_account_id,conversion_source,report_sheet_url,report_tabs&limit=1000");
    const hit = clientsR.filter((c: any) => c.name && String(c.name).length >= 4 && low.includes(String(c.name).toLowerCase()));
    if (hit.length === 1) {
      const days = /\b90\b/.test(text) ? 90 : (/\b30\b/.test(text) ? 30 : 7);
      await send(`⏳ Montando o relatório de ${hit[0].name} (${days} dias)…`);
      const rep = await waRelatorioCliente(hit[0], days);
      await send(rep);
      await saveSess({ client_id: hit[0].id, pending: null, last_msgid: w.msgid, history: [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: "[relatório do cliente enviado]" }].slice(-16) });
      return { ok: true };
    }
  }
  // ===== Agente com FERRAMENTAS: consulta qualquer banco do sistema + Meta/Google ao vivo =====
  const clients = await sbGet("clients", "select=id,name,meta_account_id,google_account_id,conversion_source,report_sheet_url,report_tabs&limit=500");
  const nomes = clients.slice(0, 150).map((c: any) => c.name).join(" | ");
  const pb = await _waPlaybook();
  const sys = `${pb}

Você é a AndréIA, gestora de tráfego E financeiro, num grupo de WhatsApp com a equipe da agência. Fale CURTO, direto e natural (é WhatsApp). Ao AVALIAR/RECOMENDAR, siga sempre o PLAYBOOK acima (ex: custo por lead/conversa alto → orientar a verificar a QUALIFICAÇÃO antes de mandar reduzir custo).
- Você CONSULTA os dados reais do sistema com as ferramentas: consultar_banco (qualquer tabela: financeiro, tarefas, CRM, RD, pedidos, clientes…), meta_insights e google_insights (métricas ao vivo), resumo_todos_clientes. SEMPRE busque o dado real antes de responder — NUNCA invente número nem use placeholders (X, Y, Z). Se não houver dado, diga que não há.
- Traga SÓ o que tem dado, e a métrica do OBJETIVO do cliente. O snapshot já traz o campo 'objetivo' e só as métricas certas dele: venda→compras/ROAS/CPA; leads→leads/CPL; mensagens→conversas/custo por conversa; tráfego→cliques/CTR/CPC. NUNCA misture (ex: cliente de VENDA não mostra "custo por conversa").
- Formato WhatsApp: NÃO use markdown de título (nada de ### ou **). Negrito é com UM asterisco (*assim*). Listas com "• ". Seja enxuta.
- ATALHOS que a equipe pode pedir: "quem precisa de atenção?" → use resumo_todos_clientes e destaque os clientes abaixo da meta, com gasto sem resultado, ou parados; "saúde da carteira" → visão geral (gasto total do período, quantos performando/abaixo, e financeiro a receber/pagar via a ferramenta financeiro); "pendências operacionais" → tarefas em aberto (consultar_banco tabela tasks, filtro status=neq.done, ordena por due); "recomendações da semana" → 2-3 ações priorizadas (o que pausar/escalar/ajustar) com base nos dados. Sempre com dado real, curto.
- Ao pedirem detalhes/campanhas de um cliente, use meta_insights e liste CADA campanha do array 'campanhas' com os KPIs que a ferramenta já trouxe pra ela (gasto, orçamento e a métrica do objetivo dela). Use SOMENTE os campos que vieram — NÃO invente nem puxe métrica de fora (ex: não some conversas num total de venda). O consolidado é o campo 'total'.
- (assist.): vendas/compras/ROAS/CPA POR CAMPANHA vêm do GERENCIADOR (pixel), não da planilha — ao mostrá-las escreva "(assist.)" ao lado do número (ex: "Compras 12 (assist.) · ROAS 3,1 (assist.)"), porque a venda REAL da agência vem da planilha e o pixel não divide venda por campanha. No consolidado do cliente que usa planilha, a venda é a real (sem "(assist.)").
- Para AÇÕES (criar tarefa, criar/cancelar reunião na agenda, pausar/reativar/duplicar campanha, orçamento, criar lançamento, dar baixa) use preparar_acao. Reunião: passe o título em 'nome', o dia em 'quando' (AAAA-MM-DD) e o horário em 'hora' (HH:MM) se disser. Cliente é opcional em reunião. — o sistema pede confirmação (SIM) e executa. NUNCA diga que já executou por conta própria. Se a mensagem citar um cliente ("no cliente X", "pro X"), passe o nome EXATO em 'cliente' — NUNCA reaproveite o cliente de mensagens anteriores quando a atual cita outro. Se o cliente não existir, o sistema avisa.
- CRIAR TAREFA: o RESPONSÁVEL (quem vai fazer) e a DATA são OBRIGATÓRIOS. Se o usuário não informar os dois, PERGUNTE (não invente responsável nem data, não assuma você mesma). Passe 'responsavel' (nome da pessoa) e 'quando' já como data ISO AAAA-MM-DD — calcule "amanhã", "hoje", "sexta" a partir de hoje. Se a pessoa disser que é URGENTE, passe urgente=true (marca a flag de urgência — NÃO escreva "urgente" no título/obs). Se pedir para SOLICITAR REVISÃO, passe revisao=true (marca a flag de revisão — NÃO escreva "revisão" no título/obs).
- DINHEIRO/FINANCEIRO: para QUALQUER pergunta de valores (a receber, a pagar, recebido, pago, fluxo do mês) use a ferramenta **financeiro** — ela já devolve o TOTAL correto e os ITENS com o nome certo do cliente. "a receber" = {tipo:'receita',status:'pendente'}; "a pagar" = {tipo:'despesa',status:'pendente'}; "este mês" = mes:'${new Date().toISOString().slice(0, 7)}'. NUNCA some você mesma nem adivinhe o nome do cliente — use os campos 'total' e 'itens' que a ferramenta retorna, exatamente.
- Datas: hoje é ${new Date().toISOString().slice(0, 10)}. Ao filtrar por um cliente específico use o id dele (está na lista abaixo entre colchetes, ou consulte a tabela clients). Clientes: ${clients.slice(0, 150).map((c: any) => `${c.name}[${c.id}]`).join(" | ")}.`;
  const hist0 = ((sess && sess.history) || []).slice(-8).map((h: any) => ({ role: h.role === "assistant" ? "assistant" : "user", content: h.text }));
  const messages: any[] = [{ role: "system", content: sys }, ...hist0, { role: "user", content: text }];
  let clientId = (sess && sess.client_id) || null;
  for (let it = 0; it < 6; it++) {
    const j = await callOpenAI({ model: "gpt-4o-mini", messages, tools: WA_TOOLS, tool_choice: "auto", max_tokens: 900, temperature: 0.3 });
    const msg = j.choices[0].message;
    if (msg.tool_calls && msg.tool_calls.length) {
      messages.push(msg);
      let acted = false;
      for (const tc of msg.tool_calls) {
        let args: any = {}; try { args = JSON.parse(tc.function.arguments || "{}"); } catch { args = {}; }
        if (tc.function.name === "preparar_acao") {
          let cid = clientId; let reply = "";
          if (args.cliente) {
            const rc = _waResolveClient(args.cliente, clients);
            if (rc) cid = rc.id;
            else reply = `🤔 Não achei o cliente "${args.cliente}" no sistema. Confere o nome pra mim? (se quiser, peço a lista de clientes)`;
          }
          const semCliente = args.tipo === "criar_reuniao" || args.tipo === "cancelar_reuniao";
          if (!reply && !cid && !semCliente) reply = "De qual cliente é essa ação? Me diz o nome do cliente.";
          // REUNIÃO: exige data pra criar; cliente é opcional
          if (!reply && args.tipo === "criar_reuniao") {
            const dataOk = args.quando && /^\d{4}-\d{2}-\d{2}$/.test(String(args.quando));
            if (!dataOk) reply = "Pra qual *dia* é a reunião? (e o horário, se tiver)";
            else args._due = args.quando;
          }
          if (!reply && args.tipo === "cancelar_reuniao" && args.quando && /^\d{4}-\d{2}-\d{2}$/.test(String(args.quando))) args._due = args.quando;
          // TAREFA: exige responsável (da equipe) + data — senão PERGUNTA
          if (!reply && args.tipo === "criar_tarefa") {
            const team = await sbGet("team", "select=id,name");
            let ownerId: string | null = null;
            if (args.responsavel) { const q = String(args.responsavel).toLowerCase(); const tm = team.find((t: any) => t.name.toLowerCase() === q) || team.find((t: any) => t.name.toLowerCase().includes(q)); if (tm) ownerId = tm.id; }
            const dataOk = args.quando && /^\d{4}-\d{2}-\d{2}$/.test(String(args.quando));
            if (args.responsavel && !ownerId) reply = `Não achei "${args.responsavel}" na equipe. Os responsáveis são: ${team.map((t: any) => t.name).join(", ")}. Pra quem é a tarefa?`;
            else {
              const faltam: string[] = [];
              if (!ownerId) faltam.push("o *responsável* (quem vai fazer)");
              if (!dataOk) faltam.push("a *data* (pra quando)");
              if (faltam.length) reply = `Pra criar a tarefa, faltou ${faltam.join(" e ")}. Pode me informar?`;
              else { args._owner = ownerId; args._due = args.quando; }
            }
          }
          if (!reply) { const pending = { ...args, client_id: cid }; reply = _waConfirmText(pending, clients); const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: reply }].slice(-16); await saveSess({ client_id: cid, pending, last_msgid: w.msgid, history: hist }); }
          else { const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: reply }].slice(-16); await saveSess({ pending: null, last_msgid: w.msgid, history: hist }); }
          await send(reply); acted = true; break;
        }
        const result = await waExecTool(tc.function.name, args, clients);
        if (result && (result as any)._cid) clientId = (result as any)._cid;
        messages.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify(result).slice(0, 7000) });
      }
      if (acted) return { ok: true };
      continue;
    }
    const reply = msg.content || "Ok.";
    const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: reply }].slice(-16);
    await saveSess({ client_id: clientId, pending: null, last_msgid: w.msgid, history: hist });
    await send(reply); return { ok: true };
  }
  await send("Me embananei aqui 😅 pode reformular?");
  return { ok: true };
}
async function waHandler(w: any) {
  if (w.op === "extract") return await waExtract(w.convId, w.autoApply !== false);
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

// ===== AndréIA — Automações / Central de notificações =====
function _brDate(s: any) { const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}` : (s || ""); }
function _spNow() { return new Date(Date.now() - 3 * 3600e3); } // America/Sao_Paulo (UTC-3)
function _mesAtual() { return _spNow().toISOString().slice(0, 7); }
// Padrão visual dos avisos/relatórios no WhatsApp: cabeçalho + divisória + bullets + negrito nos destaques.
const WA_DIV = "━━━━━━━━━━━━━━━";
function _waFmtFinanceiro(res: any, titulo: string) {
  if (!res || res.erro) return `${titulo}\n${WA_DIV}\n${(res && res.erro) || "sem dados"}`;
  if (!res.itens || !res.itens.length) return `${titulo}\n${WA_DIV}\nNada pendente 🎉`;
  const linhas = res.itens.slice().sort((a: any, b: any) => String(a.vencimento || "").localeCompare(String(b.vencimento || "")))
    .map((i: any) => `• *${i.cliente}* — ${_fmtR(i.valor)}  _(venc. ${_brDate(i.vencimento)})_`).join("\n");
  return `${titulo}\n${WA_DIV}\n${linhas}\n\n💰 *Total: ${_fmtR(res.total)}*`;
}
async function waPendenciasText() {
  const rows = await sbGet("tasks", "status=neq.done&select=name,client,owner,due,prio&order=due.asc&limit=40");
  if (!rows.length) return `✅ *Tarefas em aberto*\n${WA_DIV}\nNenhuma tarefa pendente 🎉`;
  const map = await _waClientsMap();
  const team = await sbGet("team", "select=id,name"); const tm: Record<string, string> = {}; team.forEach((t: any) => { tm[t.id] = t.name; });
  const linhas = rows.slice(0, 25).map((r: any) => `• *${r.name}*${r.client && map[r.client] ? ` — ${map[r.client]}` : ""}${r.owner && tm[r.owner] ? `\n   👤 ${tm[r.owner]}` : ""}${r.due ? `${r.owner && tm[r.owner] ? " · " : "\n   "}📅 ${_brDate(r.due)}` : ""}`).join("\n");
  return `✅ *Tarefas em aberto* (${rows.length})\n${WA_DIV}\n${linhas}`;
}
// Runner só-leitura (sem ações) pra gerar texto de análise/recomendações
async function waAgentOneShot(prompt: string): Promise<string> {
  const clients = await sbGet("clients", "select=id,name,meta_account_id,google_account_id,conversion_source,report_sheet_url,report_tabs&limit=1000");
  const nomes = clients.slice(0, 200).map((c: any) => c.name).join(" | ");
  const pb = await _waPlaybook();
  const sys = `${pb}\n\nVocê é a AndréIA, gestora de tráfego da GT Marketing, mandando um aviso automático no grupo de WhatsApp da equipe. Consulte os dados REAIS com as ferramentas antes de afirmar qualquer número. Analise cada cliente pelo OBJETIVO dele (venda→ROAS/CPA; leads→CPL; mensagens→custo por conversa; tráfego→CPC; alcance→alcance/CPM) — nunca mostre ROAS pra quem não é venda, e siga o PLAYBOOK acima (custo alto pede verificar QUALIFICAÇÃO, não só reduzir custo). Hoje é ${_spNow().toISOString().slice(0, 10)}. Clientes: ${nomes}.

FORMATAÇÃO (padrão dos avisos — siga SEMPRE, deixe visualmente limpo e organizado):
- 1ª linha: título com emoji + *negrito*. Logo abaixo, uma linha divisória: ${WA_DIV}
- Um bloco por cliente/item: nome em *negrito*, e cada informação numa linha própria começando com "• ".
- Números-chave (ROAS, gasto, CPL, custo) em *negrito*.
- Uma linha em branco entre um cliente/bloco e outro.
- Negrito é *asterisco simples*; itálico _underline_. NUNCA use # nem ** nem tabelas.`;
  const tools = WA_TOOLS.filter((t: any) => t.function.name !== "preparar_acao");
  const messages: any[] = [{ role: "system", content: sys }, { role: "user", content: prompt }];
  for (let it = 0; it < 5; it++) {
    const j = await callOpenAI({ model: "gpt-4o-mini", messages, tools, tool_choice: "auto", max_tokens: 900, temperature: 0.3 });
    const msg = j.choices[0].message;
    if (msg.tool_calls && msg.tool_calls.length) {
      messages.push(msg);
      for (const tc of msg.tool_calls) { let a: any = {}; try { a = JSON.parse(tc.function.arguments || "{}"); } catch { /* */ } const res = await waExecTool(tc.function.name, a, clients); messages.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify(res).slice(0, 7000) }); }
      continue;
    }
    return msg.content || "";
  }
  return "";
}
async function waAutoText(tipo: string, escopo = "padrao", prompt = "", nivel = "resumido"): Promise<string[]> {
  const escNota = escopo && escopo !== "padrao" ? ` Considere apenas os clientes do escopo "${ESCOPO_LABEL[escopo] || escopo}".` : "";
  if (tipo === "custom") { const p = String(prompt || "").trim(); if (!p) return []; const t = await waAgentOneShot(p + escNota); return t ? [t] : ["Não consegui montar esse aviso agora."]; }
  if (tipo === "resumo7") return await waAgentAllClientsSummary(7, escopo, nivel);
  if (tipo === "resumo30") return await waAgentAllClientsSummary(30, escopo, nivel);
  if (tipo === "restricoes") return await waAgentAllClientsSummary(7, "com_restricao");
  if (tipo === "receber") return [_waFmtFinanceiro(await waFinanceiro({ tipo: "receita", status: "pendente", mes: _mesAtual() }), "💰 *A receber este mês*")];
  if (tipo === "pagar") return [_waFmtFinanceiro(await waFinanceiro({ tipo: "despesa", status: "pendente", mes: _mesAtual() }), "💸 *A pagar este mês*")];
  if (tipo === "pendencias") return [await waPendenciasText()];
  if (tipo === "atencao") return [await waAgentOneShot(`Quem precisa de atenção hoje? Analise os clientes ativos (use resumo_todos_clientes e, se precisar, meta_insights por cliente) e destaque só os que estão abaixo da meta, gastando sem resultado, ou parados. Se estiver tudo bem, diga que está tudo em ordem. Curto.${escNota}`)];
  if (tipo === "recomendacoes") return [await waAgentOneShot(`Recomendações da semana: com base nos dados reais dos clientes ativos, liste 2 a 3 ações priorizadas (o que pausar, escalar ou ajustar), citando o cliente. Curto e prático.${escNota}`)];
  return [];
}
async function _andreiaGroupInst() {
  const data = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  const aw = data.andreia_wa || {};
  if (!aw.instance_id || !aw.group_jid) return { erro: "grupo da AndréIA não configurado" };
  const inst = (await sbGet("wa_instances", `id=eq.${encodeURIComponent(aw.instance_id)}&select=uaz_host,uaz_token`))[0];
  if (!inst) return { erro: "instância da AndréIA não encontrada" };
  return { inst, group: aw.group_jid };
}
async function _sendGroup(g: any, msgs: string[]) {
  for (const mm of msgs) { if (mm && String(mm).trim()) await waCall(g.inst.uaz_host, g.inst.uaz_token, "/send/text", "POST", { number: g.group, text: String(mm) }); }
}
// Access token do Google Agenda a partir do refresh_token guardado (mesmos secrets do OAuth).
async function _googleCalToken(): Promise<string | null> {
  const data = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  const rt = data.google_cal?.refresh_token; if (!rt) return null;
  const cid = Deno.env.get("GOOGLE_CAL_CLIENT_ID") || Deno.env.get("GOOGLE_ADS_CLIENT_ID") || "";
  const cs = Deno.env.get("GOOGLE_CAL_CLIENT_SECRET") || Deno.env.get("GOOGLE_ADS_CLIENT_SECRET") || "";
  try {
    const r = await fetch("https://oauth2.googleapis.com/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: cid, client_secret: cs, refresh_token: rt, grant_type: "refresh_token" }) });
    const j = await r.json(); return j.access_token || null;
  } catch (_e) { return null; }
}
// Lembra X min antes de cada reunião do Google Agenda. Roda a cada ~5 min (cron). Dedup via wa_reminded.
async function waMeetingRemindersTick() {
  const autos = await sbGet("andreia_automations", "enabled=eq.true&tipo=eq.lembrete_reuniao&select=*");
  if (!autos.length) return { skip: "nenhum lembrete de reunião ativo" };
  const g: any = await _andreiaGroupInst(); if (g.erro) return { skip: g.erro };
  const tok = await _googleCalToken(); if (!tok) return { skip: "Google Agenda não conectado" };
  const now = Date.now();
  const maxAnt = Math.max(...autos.map((a: any) => Number(a.antecedencia) || 15));
  const p = new URLSearchParams({ singleEvents: "true", orderBy: "startTime", maxResults: "50", timeMin: new Date(now - 60000).toISOString(), timeMax: new Date(now + (maxAnt + 6) * 60000).toISOString() });
  const r = await fetch(`https://www.googleapis.com/calendar/v3/calendars/primary/events?${p}`, { headers: { Authorization: `Bearer ${tok}` } });
  const gj = await r.json(); if (gj.error) return { skip: "calendar: " + (gj.error.message || "") };
  const items = (gj.items || []).filter((e: any) => e.status !== "cancelled" && e.start && e.start.dateTime);
  let sent = 0;
  for (const ev of items) {
    const minsUntil = (new Date(ev.start.dateTime).getTime() - now) / 60000;
    const match = autos.find((a: any) => { const ant = Number(a.antecedencia) || 15; return minsUntil <= ant + 0.5 && minsUntil > ant - 5.5; });
    if (!match) continue;
    if ((await sbGet("wa_reminded", `event_id=eq.${encodeURIComponent(ev.id)}&select=event_id&limit=1`)).length) continue;
    const hm = String(ev.start.dateTime).slice(11, 16);
    const meet = ev.hangoutLink || (ev.conferenceData?.entryPoints || []).map((x: any) => x.uri).find(Boolean) || "";
    const txt = `⏰ *Lembrete de reunião*\n${WA_DIV}\nComeça em ~${Math.max(1, Math.round(minsUntil))} min (${hm})\n*${ev.summary || "Reunião"}*${ev.location ? `\n📍 ${ev.location}` : ""}${meet ? `\n🔗 ${meet}` : ""}`;
    try { await waCall(g.inst.uaz_host, g.inst.uaz_token, "/send/text", "POST", { number: g.group, text: txt }); await sbPost("wa_reminded", { event_id: ev.id, reminded_at: new Date().toISOString() }); sent++; } catch (_e) { /* */ }
  }
  return { sent };
}
async function waAutomationRunNow(id: string) {
  const a = (await sbGet("andreia_automations", `id=eq.${encodeURIComponent(id)}&select=*`))[0];
  if (!a) return { erro: "automação não encontrada" };
  const g = await _andreiaGroupInst(); if ((g as any).erro) return g;
  const msgs = await waAutoText(a.tipo, a.escopo || "padrao", a.prompt || "", a.nivel || "resumido"); if (!msgs.length) return { erro: "tipo desconhecido ou aviso vazio" };
  await _sendGroup(g, msgs);
  await sbPatchD("andreia_automations", `id=eq.${encodeURIComponent(id)}`, { last_run: _spNow().toISOString().slice(0, 10) });
  return { ok: true, enviados: msgs.length };
}
async function waAutomationsTick() {
  const g = await _andreiaGroupInst(); if ((g as any).erro) return { skip: (g as any).erro };
  const now = _spNow(); const day = now.getUTCDay();
  const hhmm = String(now.getUTCHours()).padStart(2, "0") + ":" + String(now.getUTCMinutes()).padStart(2, "0");
  const today = now.toISOString().slice(0, 10);
  const autos = await sbGet("andreia_automations", "enabled=eq.true&select=*");
  let ran = 0; const feitas: string[] = [];
  for (const a of autos) {
    if (a.last_run === today) continue;
    const dias = Array.isArray(a.dias) ? a.dias.map((x: any) => String(x)) : ["todos"];
    const diaOk = dias.includes("todos") || (dias.includes("uteis") && day >= 1 && day <= 5) || dias.includes(String(day));
    if (!diaOk) continue;
    if ((a.hora || "08:00") > hhmm) continue; // ainda não chegou a hora hoje
    try { const msgs = await waAutoText(a.tipo, a.escopo || "padrao", a.prompt || "", a.nivel || "resumido"); await _sendGroup(g, msgs); await sbPatchD("andreia_automations", `id=eq.${encodeURIComponent(a.id)}`, { last_run: today }); ran++; feitas.push(a.titulo || a.tipo); } catch (e) { /* segue as demais */ }
  }
  return { ran, feitas, hhmm, day, today };
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
    if (body.automationsTick) {
      const r = await waAutomationsTick();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.reminderTick) {
      const r = await waMeetingRemindersTick();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.automationRunNow) {
      const r = await waAutomationRunNow(body.automationRunNow.id);
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
    if (body.googleTermCleanup) {
      const r = await googleTermCleanup(body.googleTermCleanup);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleBudget) {
      const r = await googleUpdateBudget(body.googleBudget);
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
    if (body.metaCloneCampaign) {
      const r = await metaCloneCampaign(body.metaCloneCampaign);
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
