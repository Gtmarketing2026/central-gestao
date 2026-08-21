import { google } from "npm:googleapis@144";
import { PDFDocument, StandardFonts, rgb } from "npm:pdf-lib@1.17.1";

// SEGURANÇA DE DEPLOY: manter verify_jwt=false nesta função. O projeto usa as novas
// chaves publicáveis/assinaturas do Supabase; a validação antiga do gateway pode rejeitar
// sessões AAL2 válidas. Toda rota é autenticada dentro do handler por accessControl,
// accountConfigAccess ou _guardUserRequest (incluindo perfil, AAL2, permissões e cliente).

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

// IA: o Google expõe um endpoint compatível com o da OpenAI, então trocar de provedor não mexe no
// resto do código — só no nome do modelo e em dois parâmetros. Quem é o principal: ver _iaProvider.
const _GEMINI_MODEL: Record<string, string> = { "gpt-4o": "gemini-flash-latest", "gpt-4o-mini": "gemini-3.5-flash-lite" };
// se o modelo estiver sobrecarregado ("high demand"), tenta o próximo da fila
const _GEMINI_FALLBACK: Record<string, string[]> = {
  "gemini-flash-latest": ["gemini-3.5-flash-lite", "gemini-3.5-flash"],
  "gemini-3.5-flash-lite": ["gemini-flash-latest", "gemini-3.5-flash"],
  "gemini-3.5-flash": ["gemini-flash-latest", "gemini-3.5-flash-lite"],
};
function _iaProviderOpenAI() {
  const o = Deno.env.get("OPENAI_API_KEY");
  return o ? { key: o, url: "https://api.openai.com/v1/chat/completions", nome: "OpenAI", mapa: null as any } : null;
}
function _iaProviderGemini() {
  const g = Deno.env.get("GEMINI_API_KEY");
  return g ? { key: g, url: "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions", nome: "Gemini", mapa: _GEMINI_MODEL } : null;
}
/* Principal = OpenAI, decidido em 18/08/2026: a chave do Gemini era do plano gratuito e a cota
   estourava a cada dois dias, derrubando TODA a IA de uma vez (DNA, CRM, leitura de print e a
   AndréIA do WhatsApp). O Gemini fica como reserva. Para trocar de novo não precisa de deploy:
   basta o secret IA_PRINCIPAL = "gemini". */
function _iaPreferido() { return String(Deno.env.get("IA_PRINCIPAL") || "openai").toLowerCase(); }
function _iaProvider() {
  const querGemini = _iaPreferido() === "gemini";
  const p = querGemini ? (_iaProviderGemini() || _iaProviderOpenAI()) : (_iaProviderOpenAI() || _iaProviderGemini());
  if (p) return p;
  throw new Error("Nenhuma chave de IA configurada (OPENAI_API_KEY ou GEMINI_API_KEY)");
}
// o outro provedor, seja qual for o principal — é ele que assume quando o principal fica sem cota
function _iaProviderReserva(atual: any) {
  const outro = atual?.mapa ? _iaProviderOpenAI() : _iaProviderGemini();
  return outro && outro.nome !== atual?.nome ? outro : null;
}
// Cota estourada derrubava TODA a IA do sistema de uma vez (DNA, Qualidade do Atendimento, resumos do CRM),
// e a tela so dizia "non-2xx". Aqui o erro vira uma frase que diz o que fazer.
function _iaErroHumano(msg: string, provedor: string) {
  const m = String(msg || "");
  if (/quota|RESOURCE_EXHAUSTED|billing|rate.?limit|429/i.test(m)) {
    // "sem crédito" e "rápido demais" pedem ações opostas; sem separar, a pessoa espera à toa ou
    // recarrega sem precisar. O trecho cru vai junto porque é ele que diz qual dos dois é.
    const semCredito = /insufficient_quota|billing|exceeded your current quota|free_tier/i.test(m);
    const acao = semCredito
      ? "A chave está sem crédito/cota do plano — é preciso ativar ou recarregar o faturamento dela."
      : "Foram pedidos demais em pouco tempo; costuma liberar sozinho em alguns minutos.";
    return `A IA atingiu o limite da chave do ${provedor}. ${acao} Nada foi perdido. (${m.slice(0, 140).replace(/\s+/g, " ")})`;
  }
  if (/high demand|overload|unavailable|503/i.test(m)) return `A IA do ${provedor} está sobrecarregada agora. Tente de novo em alguns minutos.`;
  return m || `Erro na API da ${provedor}`;
}
/* Nem toda chamada de IA passa por callOpenAI: transcricao de audio e leitura de video vao direto
   pro provedor. Sem registrar, elas somem da tela de custo — e o Whisper e cobrado por MINUTO, nao
   por token. Por isso o audio grava a DURACAO em input_units e o preco dele em ai_model_prices e
   "USD por 1 milhao de segundos" (US$ 100 = US$ 0,006/min), o que faz a mesma formula servir. */
async function _regUsoIa(service: string, action: string, model: string, inUnits = 0, outUnits = 0, clientId: string | null = null) {
  try {
    await sbPost("system_usage_events", { client_id: clientId, service_key: service, action: String(action).slice(0, 80),
      input_units: Math.round(inUnits) || 0, output_units: Math.round(outUnits) || 0, quantity: 1, meta: { model } });
  } catch (_e) { /* telemetria nunca interrompe a IA */ }
}
async function callOpenAI(body: any) {
  const p = _iaProvider();
  const telemetry = body?._telemetry || {};
  const modelo = body.model || "gpt-4o-mini";
  const payload: any = { temperature: 0.6, max_tokens: 1000, ...body };
  delete payload._telemetry;
  delete payload.model;
  const maxOriginal = payload.max_tokens || 1000;
  // o nome do modelo muda com o provedor; e o Gemini precisa de ajustes que a OpenAI recusa
  const modeloDe = (prov: any) => prov?.mapa ? (prov.mapa[modelo] || "gemini-3.5-flash") : modelo;
  const tentar = async (mod: string, prov = p) => {
    const corpo: any = { ...payload, model: mod };
    if (prov.mapa) { // Gemini gasta parte do orçamento "pensando": dá folga e pede raciocínio curto
      corpo.max_tokens = Math.max(1200, maxOriginal * 3);
      corpo.reasoning_effort = "low";
    } else { // a OpenAI rejeita reasoning_effort nos modelos comuns (400) — sem isso a troca falha calada
      delete corpo.reasoning_effort; corpo.max_tokens = maxOriginal;
    }
    const r = await fetch(prov.url, { method: "POST", headers: { "Authorization": `Bearer ${prov.key}`, "Content-Type": "application/json" }, body: JSON.stringify(corpo) });
    const j = await r.json();
    const jj = Array.isArray(j) ? j[0] : j;
    return { ok: r.ok && !jj?.error, json: jj, erro: jj?.error?.message || (r.ok ? "" : `HTTP ${r.status}`) };
  };
  let t = await tentar(modeloDe(p));
  let usado = p; // provedor que REALMENTE respondeu — é ele que vai pra telemetria de custo
  let erroPlanoB = ""; // por que a reserva não salvou — sem isso o erro culpa só o principal
  const semCota = (e: string) => /high demand|overload|unavailable|RESOURCE_EXHAUSTED|quota|503|429|insufficient_quota|billing/i.test(e || "");
  // modelo sobrecarregado: no Gemini vale tentar os irmãos antes de desistir
  if (!t.ok && p.mapa && semCota(t.erro)) {
    for (const alt of (_GEMINI_FALLBACK[modeloDe(p)] || [])) {
      t = await tentar(alt);
      if (t.ok) break;
    }
  }
  /* A cota é da CHAVE, não do modelo: estourou, o provedor inteiro para e o sistema fica sem IA.
     Aqui o OUTRO provedor assume, seja ele qual for — a troca funciona nos dois sentidos. */
  if (!t.ok && semCota(t.erro)) {
    const alt = _iaProviderReserva(p);
    if (!alt) erroPlanoB = `não há chave de reserva configurada (${p.mapa ? "OPENAI_API_KEY" : "GEMINI_API_KEY"})`;
    else {
      const t2 = await tentar(modeloDe(alt), alt);
      if (t2.ok) { t = t2; usado = alt; } else erroPlanoB = `${alt.nome}: ${t2.erro || "erro desconhecido"}`;
    }
  }
  if (!t.ok) {
    // erro cru dos dois provedores no log: é o que permite saber se é falta de crédito ou pico
    console.error("[IA] principal", p.nome, "→", String(t.erro || "").slice(0, 300), "| reserva →", erroPlanoB.slice(0, 300));
    throw new Error(_iaErroHumano(t.erro, p.nome) + (erroPlanoB ? ` [reserva também falhou — ${erroPlanoB}]` : ""));
  }
  try {
    const u = t.json?.usage || {};
    await sbPost("system_usage_events", { client_id: telemetry.clientId || null, service_key: usado.nome.toLowerCase(), action: String(telemetry.action || "ai_request").slice(0, 80), input_units: Number(u.prompt_tokens || u.input_tokens || 0), output_units: Number(u.completion_tokens || u.output_tokens || 0), quantity: 1, meta: { model: t.json?.model || modeloDe(usado) } });
  } catch (_e) { /* telemetria nunca interrompe a IA */ }
  return t.json;
}

const DNA_SHAPE = `{
  "identidade": {"marca": "", "promessa": "", "posicionamento": "", "tom": "", "sobre": ""},
  "produtos": [{"nome": "", "dorQueResolve": "", "desejo": "", "personaAlvo": ""}],
  "personas": [{"titulo": "", "descricao": "", "transformacao": "de [estado atual] para [estado desejado]", "estadoAtual": "", "dores": [""], "desejos": [""], "tensoes": "", "crencas": ""}],
  "objecoes": [{"objecao": "", "resposta": "", "personaAlvo": ""}],
  "diretrizes": {"tom": "", "palavrasRessoam": [""], "palavrasProibidas": [""], "abordagens": [""], "beneficios": [""], "sempreFim": [""]}
}`;
const DNA_SYSTEM = `Voce e uma estrategista de marketing e copywriting senior. A partir do material fornecido sobre um cliente/negocio (briefing, site, questionario, material institucional), voce monta o "DNA" do cliente: identidade da marca, produtos/servicos, personas detalhadas, objecoes de compra (com a quebra de cada uma) e diretrizes de copy.

Responda SOMENTE com um JSON valido no formato exato abaixo (sem markdown, sem comentarios, em portugues do Brasil):
${DNA_SHAPE}

Regras:
- Preencha com base no material; nao invente fatos concretos (nomes, precos), mas PODE inferir dores/desejos/tom coerentes com o segmento.
- personas: crie de 2 a 4 personas ricas. Cada uma com titulo curto e descritivo, descricao de 1-2 frases, transformacao (de X para Y), estadoAtual (o "antes" concreto), 4-6 dores e 4-6 desejos especificos, tensoes recorrentes e crencas/mitos a quebrar.
- produtos: liste os produtos/servicos identificados; se so houver um negocio, crie 1-3 entradas. personaAlvo deve referenciar o titulo de uma das personas.
- objecoes: 4-8 objecoes REAIS de compra desse publico, escritas na voz do cliente ("Nao tenho tempo", "Ta caro", "Sera que funciona pra mim?", "Ja tentei e nao deu certo", "Vou pensar", "Prefiro o concorrente X"). Para cada uma, "resposta" = a quebra em 1-2 frases, com argumento concreto do proprio material (garantia, prova, formato, suporte, prazo) — e nunca uma promessa que o material nao sustente. personaAlvo referencia o titulo de uma persona, ou fica vazio quando a objecao vale para todas.
- diretrizes: tom de comunicacao, 6-12 palavras que ressoam, 4-8 palavras proibidas, 4-6 abordagens de copy, 4-6 beneficios principais, e 1-3 frases para "sempre no fim da copy" (CTA/assinatura).
- Se um campo nao tiver base, deixe string vazia ou array vazio, nunca invente dado factual.`;

async function extractDna(text: string, direcionamento: string) {
  let user = `Material do cliente:\n${String(text || "").slice(0, 24000)}`;
  if (direcionamento) user += `\n\nDirecionamento do gestor (leve em conta): ${direcionamento}`;
  // 3500 nao cabia mais depois que "objecoes" entrou no formato: o JSON vinha cortado e o parse estourava
  // com erro generico. Com folga + mensagem propria, o gestor sabe o que aconteceu.
  const json = await callOpenAI({ messages: [{ role: "system", content: DNA_SYSTEM }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 5000, temperature: 0.7 });
  const content = json.choices?.[0]?.message?.content || "{}";
  try { return JSON.parse(content); } catch { throw new Error("a IA devolveu um DNA incompleto (resposta cortada). Tente de novo, ou gere por PDF/texto em vez de automático."); }
}

async function refineDna(dna: any, instrucao: string) {
  const sys = `Voce edita o DNA de um cliente (JSON). Aplique a instrucao do gestor ao DNA atual e devolva o DNA COMPLETO atualizado, no MESMO formato JSON, sem markdown. Formato:\n${DNA_SHAPE}\nMantenha tudo que nao foi pedido para mudar. Portugues do Brasil.`;
  const user = `DNA atual:\n${JSON.stringify(dna)}\n\nInstrucao: ${instrucao}`;
  const json = await callOpenAI({ messages: [{ role: "system", content: sys }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 5000, temperature: 0.5 });
  const content = json.choices?.[0]?.message?.content || "{}";
  try { return JSON.parse(content); } catch { throw new Error("a IA devolveu um DNA incompleto (resposta cortada). Refaça o pedido em partes menores."); }
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
  messages.push({ role: "user", content: `Dados do mes para o cliente "${m.clientName}", referente a ${m.mesLabel}:\n${JSON.stringify(m)}\n\nGere a analise gerencial mensal.` });
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

async function _andreiaUnifiedContext(clientId: string | null, surface: string) {
  const cid = String(clientId || "").trim();
  let clientBlock = "", knowledgeBlock = "", memoryBlock = "";
  try {
    if (cid) {
      const c = (await sbGet("clients", `id=eq.${encodeURIComponent(cid)}&select=id,name,seg,dna&limit=1`))[0];
      if (c) clientBlock = `\nCLIENTE ATUAL (uso exclusivo neste escopo):\n${JSON.stringify({ nome: c.name, segmento: c.seg || "", dna: c.dna || {} }).slice(0, 12000)}`;
    }
    const docs = await sbGet("agent_knowledge", `select=title,text,client_id&order=created_at.desc&limit=40`);
    const allowed = (docs || []).filter((d: any) => !String(d.client_id || "").trim() || (cid && String(d.client_id) === cid));
    const globalDocs = allowed.filter((d: any) => !String(d.client_id || "").trim()).slice(0, 10);
    const clientDocs = cid ? allowed.filter((d: any) => String(d.client_id) === cid).slice(0, 5) : [];
    const picked = [...globalDocs, ...clientDocs];
    if (picked.length) knowledgeBlock = "\nCONHECIMENTO COMPARTILHADO AUTORIZADO:\n" + picked.map((d: any) => `--- ${d.title || "Material"} [${d.client_id ? "cliente" : "agência"}] ---\n${String(d.text || "").slice(0, 5000)}`).join("\n\n");
  } catch (_e) { /* o núcleo continua disponível mesmo se a base estiver temporariamente indisponível */ }
  try {
    const mem = await sbGet("andreia_memory", `active=eq.true&select=client_id,kind,content,source,created_at&order=created_at.desc&limit=80`);
    const allowed = (mem || []).filter((x: any) => !x.client_id || (cid && String(x.client_id) === cid)).slice(0, 20);
    if (allowed.length) memoryBlock = "\nMEMÓRIA SELETIVA (decisões e preferências explícitas; não é transcrição):\n" + allowed.map((x: any) => `- [${x.kind || "decisão"}] ${x.content}`).join("\n");
  } catch (_e) { /* tabela pode ainda não existir durante a primeira publicação */ }
  return `
===== NÚCLEO ÚNICO DA ANDRÉIA =====
Você é UMA única inteligência em todas as interfaces da Central de Gestão. A interface atual é: ${surface}.
- Preserve continuidade de raciocínio, critérios e linguagem entre Sistema, CRM, Analytics e WhatsApp.
- Conhecimento global vale para toda a agência; DNA e conhecimento de cliente só podem ser usados no próprio cliente.
- O escopo técnico atual é ${cid ? `o cliente de id ${cid}` : "a carteira, sem cliente específico"}. Não troque de cliente por inferência.
- Nunca misture, compare nominalmente ou revele dados de outro cliente sem pedido explícito e autorização da interface.
- Dados operacionais e métricas devem ser consultados nas ferramentas/fontes reais; conhecimento ensina COMO analisar, nunca substitui dados.
- Ações com efeito no sistema devem ser preparadas e confirmadas. Nunca afirme que executou antes da confirmação.
- Diferencie fato, inferência e hipótese. Se faltar evidência, diga o que precisa ser verificado.
- Considere objetivos alternativos legítimos: vendas, leads, mensagens, recrutamento, suporte, distribuição e reconhecimento.
${clientBlock}${knowledgeBlock}${memoryBlock}`;
}

async function _andreiaMaybeRemember(clientId: string | null, surface: string, userText: any) {
  const raw = String(userText || "").trim();
  if (raw.length < 8 || raw.length > 4000) return;
  // Só memoriza instruções/decisões explícitas. Perguntas e conversas comuns nunca viram memória.
  const trigger = /\b(lembre|memorize|decidimos|definimos|a partir de agora|prefiro|não use|nunca use|sempre use|regra(?: geral)?|considere como regra|fica definido)\b/i;
  if (!trigger.test(raw)) return;
  const content = _crmAiMaskText(raw).replace(/\s+/g, " ").trim().slice(0, 700);
  if (content.length < 8) return;
  const cid = String(clientId || "").trim() || null;
  try {
    const recent = await sbGet("andreia_memory", `active=eq.true&select=id,client_id,content&order=created_at.desc&limit=100`);
    if ((recent || []).some((x: any) => String(x.client_id || "") === String(cid || "") && String(x.content || "").toLowerCase() === content.toLowerCase())) return;
    await sbPost("andreia_memory", { id: _wuid(), client_id: cid, scope: cid ? "client" : "global", kind: /prefiro|não use|nunca use|sempre use/i.test(raw) ? "preference" : "decision", content, source: String(surface || "Sistema").slice(0, 80), active: true });
  } catch (_e) { /* memória não pode bloquear a resposta */ }
}

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

DOCUMENTOS: Voce PODE montar documentos (relatorios, propostas, briefings, planos de acao, resumos executivos). Quando pedirem um documento/relatorio/PDF/Word, NUNCA diga que nao consegue gerar arquivos — escreva o CONTEUDO COMPLETO e bem formatado em markdown (titulos com #, ## e ###, listas, **negrito**, e tabelas em markdown com | quando fizer sentido) direto na resposta. Ao terminar, NAO escreva instrucao de download nem mencione botao/PDF/Word — o botao de baixar ja aparece sozinho na interface e essa frase suja o documento gerado. Termine no conteudo. O sistema converte sua resposta no layout da agencia (temas GT) automaticamente. Estruture como documento de verdade: titulo, secoes, e quando for relatorio de cliente siga a logica dos nossos templates (visao geral -> resultados por objetivo -> funil -> recomendacoes/proximos passos).

Baseie-se SOMENTE nos dados do snapshot (KPIs do relatorio do cliente, canais, funil, pedidos). Nunca invente numeros; se faltar um dado, diga que nao esta disponivel. Seja direta, especifica e priorize; nada de conselho generico de manual.

Metodo de analise:
1. Leia os KPIs: investimento, CTR, CPC, CPM, ROAS, ticket medio, pedidos e o funil (impressoes -> cliques -> checkout -> venda).
2. Diagnostique ONDE esta o gargalo:
   - CTR baixo (feed abaixo de 1%, search abaixo de 2%): sinal para investigar criativo, oferta, segmentacao e objetivo da campanha. Recomende novos angulos de criativo, hook nos primeiros 3s, revisar a estrategia de publico e a headline, sem afirmar que publico ou nicho estao errados.
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

RESUMO/RELATORIO DE RESULTADOS POR CAMPANHA (regra dura — erro real ja aconteceu aqui):
Quando pedirem resumo, relatorio ou resultados "por campanha" ou "por anuncio", NUNCA liste os campos crus de cada campanha em bullets sem analise (ex: so "Investimento: R$X / Impressoes: Y / Cliques: Z" um embaixo do outro) — isso NAO e um relatorio, e um dump de dados, e o gestor ja reclamou disso.
Para cada campanha/conjunto, estruture assim:
1. Nome da campanha + objetivo dela (use o campo 'objetivo' do snapshot).
2. SO os KPIs que aquele objetivo pede (use o playbook da regra acima — trafego mostra CPC/CTR/cliques, engajamento mostra custo por engajamento, etc.) — nao despeje TODAS as metricas disponiveis so porque existem no snapshot.
3. Uma leitura curta em texto (1-2 linhas): a campanha esta indo bem, mal ou merece atencao, e POR QUE — comparando com o benchmark do cliente ou de mercado quando der.
Feche com uma sintese geral (nao so soma de numeros: o que se destacou, o que precisa de acao) e recomendacao priorizada quando fizer sentido.
Cada cliente tem objetivo diferente, entao o conjunto de KPIs muda de cliente pra cliente — nao existe um template fixo de campos.
Se voce nao tiver certeza de quais KPIs o gestor quer ver naquele relatorio especifico, NAO trave esperando resposta: entregue o relatorio completo com os KPIs que voce mapeou pelo objetivo (do jeito que o gestor pediu, completo, sem faltar campanha) e feche com uma linha tipo "Mapeei os KPIs pelo objetivo de cada campanha — sentiu falta de algum? me fala que eu refaco com ele incluido."

PERIODOS E BENCHMARK:
- Se o snapshot tiver 'periodos' (ultimos7dias, ultimos30dias), use esses numeros quando perguntarem sobre 7 ou 30 dias.
- Se tiver 'benchmarkProprioCliente' (variacao % dos ultimos 30d vs os 30d anteriores), use como BENCHMARK DO PROPRIO CLIENTE — diga o que melhorou/piorou vs o historico dele (ex: "CTR subiu 12%, CPC caiu 8% vs o mes anterior").
- Se tiver 'metasCliente' (metas definidas pelo gestor + status atingida/abaixo vs ultimos 30d), esse e o BENCHMARK-ALVO OFICIAL do cliente: priorize ele. Diga claramente o que bateu a meta e o que ficou abaixo, com o numero da meta e o real.
- COMPARACAO DE MERCADO (referencias gerais do Meta Ads, use como parametro aproximado, nunca como verdade absoluta e sempre considerando o nicho): CTR no feed bom > 1% (otimo > 2%); CPC saudavel geralmente < R$2 (varia muito por nicho); frequencia > 3-4 no periodo indica saturacao; taxa de conversao de LP e-commerce tipica 1-3%; checkout->compra saudavel 30-50%; ROAS bom depende da margem, mas < 1 e prejuizo e > 2-3 costuma ser saudavel em e-commerce. Ao comparar com mercado, diga "acima/abaixo da media de mercado" com o numero.

Se o snapshot tiver 'dnaCliente' (identidade, produtos, personas com dores/desejos, diretrizes de copy), USE como base ao sugerir angulos de criativo, headlines, copies e publico — respeite o tom, as palavras que ressoam e evite as proibidas.
REGRA DE CONTEXTO E OBJETIVO: NUNCA afirme que "o publico esta errado", "o nicho esta incorreto" ou equivalente. Um publico, procura ou conversa fora do objetivo comercial mais comum pode pertencer a outra campanha legitima (por exemplo recrutamento/recebimento de curriculos, suporte, distribuicao ou reconhecimento). Trate isso como HIPOTESE a investigar: identifique campanha, objetivo configurado, criativo, palavra-chave e pagina de destino antes de concluir. Se esses dados nao estiverem disponiveis, diga que falta confirmar o objetivo e proponha a verificacao; nao classifique automaticamente como erro de trafego.

Voce tambem pode EXECUTAR acoes quando o gestor pedir explicitamente: criar/concluir tarefas E acoes reais no Meta Ads (pausar_meta, reativar_meta, ajustar_orcamento, duplicar_campanha). Para as acoes do Meta, use SEMPRE o 'id' e o 'nivel' que estao na lista 'metaEntidades' do snapshot (campanhas, conjuntos e anuncios com id, status e orcamento atuais) — nunca invente ids. O sistema mostra um card de confirmacao antes de executar; entao apenas PROPONHA a acao chamando a funcao e explique o porque em texto; nunca afirme que ja executou. So proponha acao no Meta quando o gestor pedir ou quando os dados claramente justificarem (ex: anuncio com gasto alto e 0 compras -> propor pausar). Seu valor principal continua sendo a analise tecnica.

===== LEITURA DE PRINT / IMAGEM (REGRA DURA) =====
Quando o gestor anexa um print de painel (e-commerce, plataforma, gerenciador), aquele print e a FONTE DA VERDADE daqueles numeros.
- TRANSCREVA o que esta escrito, exatamente como esta. Nao arredonde, nao converta, nao "melhore".
- NUNCA calcule nem estime uma variacao percentual que nao esteja escrita na imagem. Se o print nao mostra comparativo, escreva "nao informado" no lugar do percentual. E melhor faltar dado do que ter numero inventado — numero inventado vai pro relatorio do cliente e destroi a confianca.
- Se o print mostra o percentual, use o valor E o sinal exatos (+ ou -) que aparecem la.
- Nao misture a fonte: numero que veio do print nao pode ser somado nem comparado com numero do snapshot do sistema sem dizer que sao fontes diferentes.
- Se o gestor disser que voce errou um numero do print, RELEIA a imagem e corrija com o que esta escrito; nao repita o valor anterior nem invente outro.
- Se a imagem estiver ilegivel ou cortada no ponto que interessa, diga isso e peca um print melhor. Nunca preencha a lacuna com estimativa.
- Periodo de comparacao: use exatamente o que o gestor pediu (ex: mesmo periodo do mes anterior). Nao troque por ano anterior nem por "periodo anterior" generico por conta propria.

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

  /* CACHE DE PROMPT (OpenAI): o desconto só vale para o PREFIXO idêntico entre chamadas. O bloco de
     instruções acima não tem nenhuma interpolação — é igual em toda pergunta — então ele fica sozinho
     na primeira mensagem. Tudo que muda por cliente (contexto e base de conhecimento) vai numa
     SEGUNDA mensagem. Antes estava tudo concatenado, e um único caractere diferente jogava o prompt
     inteiro fora do cache. */
  let contexto = await _andreiaUnifiedContext(a.clientId || null, a.surface || "Sistema");
  /* Com imagem anexada o corpo do pedido ja carrega centenas de KB em base64. Somar a isso a base de
     conhecimento inteira e o snapshot cru estourava o limite e a resposta voltava 500 seca — sem erro
     na tela e sem rastro no log. Com print, o que importa e a imagem: o resto entra reduzido. */
  const _temImagem = Array.isArray(a.anexos) && a.anexos.some((x: any) => x?.tipo === "imagem" && x?.dataUrl);
  if (Array.isArray(a.knowledge) && a.knowledge.length) {
    /* Custo: a base de conhecimento inteira ia junto em TODA pergunta e, com o laço de ferramentas,
       era reenviada a cada volta. 14k caracteres por fonte é material de leitura, não de contexto —
       5k já carrega o método sem pagar o livro toda vez. */
    const fontes = a.knowledge.slice(0, _temImagem ? 4 : 9);
    const limite = _temImagem ? 3000 : 5000;
    contexto += `\n\n===== BASE DE CONHECIMENTO (JARVIS) =====\nEstes sao os metodos e frameworks dos gestores que a agencia treinou em voce (Pedro Sobral e outros). Eles sao a SUA forma de pensar: aplique estes principios, benchmarks e mentalidade em TODA analise e recomendacao, citando o raciocinio quando util. Nao os ignore.\n` +
      fontes.map((k: any, i: number) => `--- Fonte ${i + 1}: ${k.title || "material"} ---\n${String(k.text || "").slice(0, limite)}`).join("\n\n");
  }

  const messages: any[] = [{ role: "system", content: system }];
  if (contexto.trim()) messages.push({ role: "system", content: contexto });
  // sem indentação de propósito: JSON.stringify(x,null,2) enche o prompt de espaços que a IA cobra
  const _snap = JSON.stringify(a.snapshot);
  messages.push({ role: "user", content: `Snapshot atual (dados reais do sistema):\n${_temImagem ? _snap.slice(0, 40000) : _snap}` });
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
    const grandes: string[] = [];
    for (const ax of a.anexos.slice(0, 4)) {
      if (ax.tipo === "imagem" && ax.dataUrl) {
        // imagem gigante derruba a função inteira; melhor recusar UMA e responder do que falhar tudo
        if (String(ax.dataUrl).length > 4_000_000) { grandes.push(String(ax.nome || "imagem")); continue; }
        if (imgs.length < 3) imgs.push({ type: "image_url", image_url: { url: String(ax.dataUrl) } });
      } else if (ax.texto) messages.push({ role: "user", content: `Anexo "${ax.nome || "arquivo"}" (texto extraído):\n${String(ax.texto).slice(0, _temImagem ? 6000 : 15000)}` });
    }
    if (imgs.length) {
      messages.push({ role: "user", content: [{ type: "text", text: "Imagem(ns) anexada(s) pelo gestor. Leia os números EXATAMENTE como estão escritos; onde não houver percentual na imagem, escreva \"não informado\" — nunca calcule nem estime:" }, ...imgs] });
    }
    if (grandes.length) messages.push({ role: "user", content: `(Não consegui ler ${grandes.join(", ")}: arquivo grande demais. Avise o gestor para reenviar um print menor ou recortado só na parte que interessa.)` });
  }

  const actionNames = new Set(AGENT_TOOLS.map((t: any) => t.function.name));
  const readTools = WA_TOOLS.filter((t: any) => t.function.name !== "preparar_acao" && !actionNames.has(t.function.name));
  const allTools = [...readTools, ...AGENT_TOOLS];
  const clients = await sbGet("clients", "select=id,name,meta_account_id,google_account_id,conversion_source,report_sheet_url,report_tabs&limit=500");
  const actions: any[] = [];
  let answer = "";
  // 4 voltas em vez de 6: cada volta reenvia a conversa inteira, então o teto alto sai caro no pior caso
  for (let it = 0; it < 4; it++) {
    const json = await callOpenAI({ model: "gpt-4o", messages, tools: allTools, tool_choice: "auto", max_tokens: 2000, temperature: 0.4, _telemetry: { clientId: a.clientId || null, action: "andreia_system" } });
    const msg = json.choices?.[0]?.message || {};
    if (!Array.isArray(msg.tool_calls) || !msg.tool_calls.length) { answer = String(msg.content || ""); break; }
    messages.push(msg);
    for (const tc of msg.tool_calls) {
      let args: any = {}; try { args = JSON.parse(tc.function.arguments || "{}"); } catch (_e) { /* ignora */ }
      if (actionNames.has(tc.function.name)) {
        actions.push({ name: tc.function.name, args });
        messages.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify({ preparado: true, confirmacao_necessaria: true, instrucao: "Explique brevemente a ação preparada e aguarde confirmação no card." }) });
      } else {
        const result = await waExecTool(tc.function.name, args, clients);
        // o resultado da ferramenta volta em TODAS as chamadas seguintes do laço: cortar aqui economiza várias vezes
        messages.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify(result).slice(0, 6000) });
      }
    }
  }
  await _andreiaMaybeRemember(a.clientId || null, a.surface || "Sistema", a.question);
  return { answer, actions };
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

// Lista as Paginas do Facebook (com Instagram Business vinculado, quando tiver) que o META_USER_TOKEN enxerga.
// Usado pro seletor de conexao no cadastro do cliente e como diagnostico geral.
async function instagramListAccounts() {
  const token = await _metaUserToken();
  if (!token) return { ok: false, erro: "META_USER_TOKEN nao configurada" };
  const r = await fetch(`https://graph.facebook.com/v21.0/me/accounts?fields=name,instagram_business_account{id,username,followers_count}&limit=100&access_token=${token}`);
  const j = await r.json();
  if (j.error) return { ok: false, erro: j.error.message, code: j.error.code, type: j.error.type };
  const paginas = (j.data || []).map((p: any) => ({ pagina: p.name, pageId: p.id, instagram: p.instagram_business_account ? { id: p.instagram_business_account.id, username: p.instagram_business_account.username, seguidores: p.instagram_business_account.followers_count } : null }));
  return { ok: true, totalPaginas: paginas.length, comInstagram: paginas.filter((p: any) => p.instagram).length, paginas };
}
async function _instagramDiag() { return await instagramListAccounts(); }
function _igNormName(s: string) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }
// Tenta casar sozinho o(s) Instagram do cliente pelo nome, entre as Paginas que a agencia ja enxerga
// (sem pedir autorizacao nenhuma) - chamado automaticamente quando o cadastro tem Meta Ads mas ainda
// nao tem Instagram. Um cliente pode ter mais de um perfil (achado real: "Curso Fernanda Pessoa" tem 2
// Paginas diferentes) - conecta TODAS as correspondencias com confianca, nao so uma. So nao auto-conecta
// quando o nome bate mas ja tinha decisao manual anterior removendo aquele perfil especifico.
async function instagramAutoMatch(input: any) {
  const { clientId } = input;
  if (!clientId) throw new Error("clientId obrigatório.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=id,name,instagram_accounts,instagram_accounts_excluded`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const atual: any[] = Array.isArray(c.instagram_accounts) ? c.instagram_accounts : [];
  const excluidos = new Set((Array.isArray(c.instagram_accounts_excluded) ? c.instagram_accounts_excluded : []).map(String));
  const list = await instagramListAccounts();
  if (!list.ok) return { erro: list.erro };
  const comIg = (list.paginas || []).filter((p: any) => p.instagram);
  const alvo = _igNormName(c.name);
  const jaTemIds = new Set(atual.map((a: any) => a.id));
  const candidatos = comIg.filter((p: any) => {
    if (jaTemIds.has(p.instagram.id) || excluidos.has(String(p.instagram.id))) return false;
    const nome = _igNormName(p.pagina);
    const uname = _igNormName(p.instagram.username || "").replace(/ /g, "");
    return nome === alvo || (alvo.length > 3 && (nome.includes(alvo) || alvo.includes(nome))) || uname === alvo.replace(/ /g, "");
  });
  if (!candidatos.length) return { conectado: false, novos: 0, total: atual.length };
  const novos = candidatos.map((m: any) => ({ id: m.instagram.id, username: m.instagram.username || "", pagina: m.pagina }));
  const merged = [...atual, ...novos];
  await sbPatchD("clients", `id=eq.${encodeURIComponent(clientId)}`, { instagram_accounts: merged });
  return { conectado: true, automatico: true, novos: novos.length, total: merged.length, adicionados: novos };
}
// Posts organicos + metricas reais do(s) Instagram Business do cliente (ultimos N dias) - agrega todos
// os perfis conectados, marcando de qual username cada post veio. Alimenta a Curadoria de Conteudo do
// Briefing Criativo e, depois, a aba Social. "Percentual de alcance de nao-seguidores" e "retencao por
// trecho" (25/50/75/100%) NAO estao disponiveis nessa API pra posts organicos - nunca estimados, so
// ficam de fora dos criterios (ver briefingCuradoria).
async function instagramOrganicContent(input: any) {
  const { clientId, days, instagramId } = input;
  if (!clientId) throw new Error("clientId obrigatório.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=instagram_accounts,name`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const contas: any[] = (Array.isArray(c.instagram_accounts) ? c.instagram_accounts : []).filter((a: any) => !instagramId || a.id === instagramId);
  if (!contas.length) throw new Error("Esse cliente não tem Instagram conectado. Conecte em Configurações do cliente.");
  const token = await _metaUserToken();
  if (!token) throw new Error("META_USER_TOKEN não configurada.");
  const since = Math.floor(Date.now() / 1000) - (Number(days) || 90) * 86400;
  const fields = "id,caption,media_type,media_product_type,media_url,thumbnail_url,permalink,timestamp,like_count,comments_count";
  const porConta = await Promise.all(contas.map(async (conta: any) => {
    let url: string | null = `https://graph.facebook.com/v21.0/${conta.id}/media?fields=${fields}&limit=50&access_token=${token}`;
    const posts: any[] = [];
    for (let i = 0; i < 6 && url; i++) {
      const r: any = await fetch(url);
      const j: any = await r.json();
      if (j.error) return { conta, erro: j.error.message, posts: [] as any[] };
      let hitOld = false;
      for (const m of (j.data || [])) {
        const ts = Math.floor(new Date(m.timestamp).getTime() / 1000);
        if (ts < since) { hitOld = true; break; }
        posts.push(m);
      }
      url = hitOld ? null : (j.paging?.next || null);
    }
    const withInsights = await Promise.all(posts.map(async (m: any) => {
      const isVideo = m.media_type === "VIDEO" || m.media_product_type === "REELS";
      const metrics = isVideo ? "reach,saved,shares,total_interactions,plays" : "reach,saved,shares,total_interactions";
      const ins: Record<string, number> = {};
      try {
        const r = await fetch(`https://graph.facebook.com/v21.0/${m.id}/insights?metric=${metrics}&access_token=${token}`);
        const j = await r.json();
        if (!j.error) for (const d of (j.data || [])) ins[d.name] = d.values?.[0]?.value ?? d.total_value?.value ?? 0;
      } catch (_e) { /* segue sem insights desse post */ }
      const likes = m.like_count || 0, comments = m.comments_count || 0;
      const reach = ins.reach || 0, saved = ins.saved || 0, shares = ins.shares || 0;
      const eng = reach ? +(((likes + comments + saved + shares) / reach) * 100).toFixed(2) : null;
      return {
        id: m.id, contaId: conta.id, username: conta.username, caption: m.caption || "", tipo: m.media_type, permalink: m.permalink, midia: m.media_url || m.thumbnail_url,
        data: m.timestamp, likes, comments, reach: reach || null, saved: saved || null, shares: shares || null,
        views: ins.plays || null, eng,
      };
    }));
    return { conta, erro: null as string | null, posts: withInsights };
  }));
  const erros = porConta.filter((p) => p.erro).map((p) => `@${p.conta.username || p.conta.id}: ${p.erro}`);
  const todos = porConta.flatMap((p) => p.posts);
  todos.sort((a, b) => (b.eng ?? -1) - (a.eng ?? -1));
  return { clientId, cliente: c.name, perfis: contas.map((a: any) => a.username || a.id), total: todos.length, posts: todos, erros };
}
// Etapa 2 do Briefing Criativo: cura o conteudo organico buscado acima, aponta o que pode virar
// recorte em vez de producao nova. Bloqueios de licenca/qualidade (audio de biblioteca, marca dagua,
// rosto de terceiro, resolucao) exigem OLHAR o video/imagem - a API nao devolve isso, entao o prompt
// pede pra IA avaliar pela legenda/contexto disponivel e sinalizar "precisa conferencia humana" quando
// nao da pra saber, em vez de arriscar (nunca afirma um bloqueio que nao consegue checar).
async function briefingCuradoria(input: any) {
  const { briefingId, clientId, angulo, funil, canais, produto } = input;
  const { cliente, posts } = await instagramOrganicContent({ clientId, days: 90 });
  if (!posts.length) return { leitura: "Sem posts orgânicos no período, ou o Instagram desse cliente ainda não está conectado.", candidatos: [] };
  const conteudosTxt = posts.slice(0, 30).map((p: any) =>
    `${p.id} | ${p.tipo} | ${String(p.data).slice(0, 10)} | alcance: ${p.reach ?? "—"} | salvamentos: ${p.saved ?? "—"} | compartilhamentos: ${p.shares ?? "—"} | curtidas: ${p.likes} | comentários: ${p.comments} | engajamento: ${p.eng ?? "—"}% | legenda: ${String(p.caption || "").slice(0, 220).replace(/\n/g, " ")}`
  ).join("\n");
  const prompt = `Voce e curador de conteudo para trafego pago. Avalie o material organico abaixo e diga o que pode virar criativo pago por recorte, em vez de producao nova.

Cliente: ${cliente}
Angulo desejado: ${angulo || "não informado"}
Funil desejado: ${funil || "não informado"}
${produto ? `Produto/programa especifico: ${produto} - priorize posts cuja legenda fale claramente disso.` : ""}
Canais: ${(canais && canais.length ? canais.join(", ") : "Meta")}

Conteudos disponiveis (Instagram, ultimos 90 dias, metricas reais):
${conteudosTxt}

Criterios de pontuacao de 0 a 100, nesta ordem de peso (so com o que estiver disponivel nos dados - NAO temos percentual de alcance de nao-seguidores nem retencao por trecho de video, nao invente esses numeros):
1. Compartilhamentos sobre alcance. Indica ressonancia.
2. Salvamentos sobre alcance. Indica utilidade percebida, bom sinal de meio.
3. Engajamento geral (curtidas+comentarios+salvamentos+compartilhamentos sobre alcance).
4. Coerencia da legenda com o angulo/funil desejado.

Bloqueios que impedem o uso, quando der pra perceber pela legenda/contexto (senao, marque "precisa conferência humana" em vez de arriscar):
- Audio de biblioteca do Instagram. A licenca organica nao cobre uso comercial.
- Marca dagua de outra plataforma.
- Rosto de terceiro sem cessao de imagem, que exige conferencia humana.
- Oferta ou preco desatualizado na peca.

${REGRAS_DE_LINGUAGEM}

Traga no maximo 5 candidatos, do maior para o menor score. Item com score abaixo de 60 nao entra. Descreva o trecho/direcao de corte com base na legenda e no tipo de midia.

Responda APENAS com JSON valido, sem markdown, sem preambulo:
{"leitura":"<uma frase sobre o inventario disponivel>",
 "candidatos":[{"id":"","titulo":"","score":0,"funil":"","corte":"","aproveitar":"","complemento":"","bloqueios":""}]}`;
  const parsed = await _callOpenAIJson([{ role: "user", content: prompt }]);
  if (briefingId) {
    await sbPost("briefing_curadoria", { id: _wuid(), briefing_id: briefingId, leitura: parsed.leitura || "", candidatos_json: parsed.candidatos || [], gerado_em: new Date().toISOString() });
  }
  return parsed;
}

async function metaAdsInsights(m: any) {
  const token = await _metaUserToken();
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
  const fields = "spend,impressions,clicks,ctr,cpc,cpm,reach,frequency,actions,action_values,purchase_roas,video_thruplay_watched_actions,video_30_sec_watched_actions";

  async function fetchInsights(acct: string, level: string, extra = "") {
    let lvlFields = "";
    if (level === "campaign") lvlFields = ",campaign_name,campaign_id";
    else if (level === "adset") lvlFields = ",campaign_name,campaign_id,adset_name,adset_id";
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
          } else if (opt === "VISIT_INSTAGRAM_PROFILE") {
            map[as.campaign_id] = { ...ob, tipo: "perfil", rotulo: "Visitas ao perfil (Instagram)", metrica: "seguidores ganhos, custo por seguidor" };
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
      // ThruPlay é o "Resultado" das campanhas de vídeo — vem em campo PRÓPRIO (não no array actions). Fallback: 3s (video_view) / 30s.
      videoViews: Number((row.video_thruplay_watched_actions && row.video_thruplay_watched_actions[0] && row.video_thruplay_watched_actions[0].value) || 0) || pickOne(row.actions, ["video_view"]) || Number((row.video_30_sec_watched_actions && row.video_30_sec_watched_actions[0] && row.video_30_sec_watched_actions[0].value) || 0),
      engajamentos: pickOne(row.actions, ["post_engagement"]),
      // seguidores ganhos (campanha de visita ao perfil do Instagram) — nomes candidatos em ordem de prioridade;
      // se o Meta nao reportar follow nessa conta, fica 0 e o card mostra o resultado sem inventar numero
      seguidores: pickOne(row.actions, ["onsite_conversion.ig_follow", "ig_follow", "follow", "onsite_conversion.follow"]),
    };
  }
  const totAgg: any = { spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0, videoViews: 0, engajamentos: 0 };
  const byCamp: Record<string, any> = {};
  const byAdset: Record<string, any> = {};
  const byAdDaily: Record<string, any> = {}; // anuncio x dia (banco de dados de midia — schema `midia`) - so populado quando byAd+daily juntos
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
    let statusIssue: string | null = null;
    try {
      const [status, accountRows, acctDaily, objByCampId, adRows, campRows, campDedup, adsetRows, adDailyRows] = await Promise.all([
        fetchAccountStatus(acc.id),
        fetchInsights(acc.id, "account"),
        m.daily ? fetchInsights(acc.id, "account", "&time_increment=1") : Promise.resolve([] as any[]),
        wantObj ? fetchObjectives(acc.id) : Promise.resolve({} as Record<string, any>),
        m.byAd ? fetchInsights(acc.id, "ad") : Promise.resolve([] as any[]),
        m.byCampaign ? fetchInsights(acc.id, "campaign", m.daily ? "&time_increment=1" : "") : Promise.resolve([] as any[]),
        // ALCANCE nao e somavel: a busca diaria (time_increment=1) soma a mesma pessoa a cada dia.
        // Buscamos tambem SEM quebra diaria pra ter o reach/frequencia DEDUPLICADO por campanha no periodo.
        (m.byCampaign && m.daily) ? fetchInsights(acc.id, "campaign", "") : Promise.resolve([] as any[]),
        m.byAdset ? fetchInsights(acc.id, "adset", m.daily ? "&time_increment=1" : "") : Promise.resolve([] as any[]),
        // anúncio x dia é muito volumoso. Só busca quando solicitado explicitamente;
        // Dashboard/Relatório precisam do total por anúncio, não da duplicação anúncio × dia.
        m.byAdDaily ? fetchInsights(acc.id, "ad", "&time_increment=1") : Promise.resolve([] as any[]),
      ]);
      statusIssue = status;
      return { acc, accountRows, acctDaily, objByCampId, adRows, campRows, campDedup, adsetRows, adDailyRows, error: statusIssue as string | null };
    } catch (e) {
      // conta com erro NAO derruba as outras: devolve vazia + motivo (front mostra o disclaimer)
      return { acc, accountRows: [] as any[], acctDaily: [] as any[], objByCampId: {} as Record<string, any>, adRows: [] as any[], campRows: [] as any[], campDedup: [] as any[], adsetRows: [] as any[], adDailyRows: [] as any[], error: statusIssue || (e as any)?.message || String(e) };
    }
  }));
  const accountErrors = perAccount.filter((p) => p.error).map((p) => ({ id: p.acc.id, name: p.acc.name || p.acc.id, error: p.error }));
  const totRecByDate: Record<string, any> = {};
  for (const { acc, accountRows, acctDaily, objByCampId, adRows, campRows, campDedup, adsetRows, adDailyRows } of perAccount) {
    for (const row of acctDaily) {
      const s = shape(row); const k = row.date_start;
      if (!totRecByDate[k]) totRecByDate[k] = { date: k, sales: 0, spend: 0, revenue: 0, clicks: 0, impressions: 0, reach: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0, addToCart: 0, checkout: 0 };
      const rec = totRecByDate[k];
      rec.sales += Math.round(s.purchases); rec.spend += s.spend; rec.revenue += s.revenue; rec.clicks += s.clicks; rec.impressions += s.impressions;
      rec.reach += s.reach; rec.leads += s.leads; rec.conversas += s.conversas; rec.videoViews += s.videoViews; rec.engajamentos += s.engajamentos; rec.addToCart += s.addToCart; rec.checkout += s.initiateCheckout;
    }
    // accountRows (level=account, sem time_increment) normalmente vem em 1 linha só, mas o Meta pode
    // devolver mais de uma (ex: conta que mudou de configuracao de atribuicao no meio do periodo) — usar so
    // accountRows[0] descartava o resto silenciosamente e subcontava o investimento. Soma todas as linhas.
    for (const row of accountRows) {
      const at = shape(row);
      totAgg.spend += at.spend; totAgg.impressions += at.impressions; totAgg.clicks += at.clicks; totAgg.reach += at.reach;
      totAgg.revenue += at.revenue; totAgg.purchases += at.purchases; totAgg.leads += at.leads; totAgg.addToCart += at.addToCart; totAgg.initiateCheckout += at.initiateCheckout;
      totAgg.conversas += at.conversas; totAgg.videoViews += at.videoViews; totAgg.engajamentos += at.engajamentos;
    }
    for (const row of adRows) {
      const s = shape(row);
      ads.push({
        adId: row.ad_id, adName: row.ad_name || "(sem nome)", campaign: row.campaign_name || "", campaignId: row.campaign_id || null, adset: row.adset_name || "", adsetId: row.adset_id || null,
        account: acc.name || acc.id, accountId: acc.id, thumbnail: null,
        objetivo: objByCampId[row.campaign_id] || metaObjetivo(""),
        spend: s.spend, impressions: s.impressions, clicks: s.clicks, reach: s.reach, frequency: s.frequency,
        ctr: s.ctr, cpc: s.cpc, cpm: s.cpm, purchases: s.purchases, revenue: s.revenue, roas: s.roas,
        leads: s.leads, addToCart: s.addToCart, initiateCheckout: s.initiateCheckout,
        conversas: s.conversas, videoViews: s.videoViews, engajamentos: s.engajamentos, seguidores: (s as any).seguidores || 0,
        cpa: s.purchases ? s.spend / s.purchases : 0,
      });
    }
    // anuncio x dia (banco de dados de midia) — mesma logica do adRows, so que 1 linha por dia em vez de total do periodo
    for (const row of adDailyRows) {
      const s = shape(row);
      const adId = row.ad_id;
      if (!adId || !row.date_start) continue;
      if (!byAdDaily[adId]) byAdDaily[adId] = {
        adId, adName: row.ad_name || "(sem nome)", campaign: row.campaign_name || "", campaignId: row.campaign_id || null,
        adset: row.adset_name || "", adsetId: row.adset_id || null, account: acc.name || acc.id, accountId: acc.id,
        objetivo: objByCampId[row.campaign_id] || metaObjetivo(""), records: [],
      };
      byAdDaily[adId].records.push({
        date: row.date_start, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions,
        reach: s.reach, frequency: s.frequency, leads: s.leads, conversas: s.conversas, videoViews: s.videoViews, engajamentos: s.engajamentos,
      });
    }
    for (const row of campRows) {
      const label = row.campaign_name || "Meta Ads";
      const s = shape(row);
      if (!byCamp[label]) byCamp[label] = { campaign: label, campaignId: row.campaign_id || null, account: acc.name || acc.id, objetivo: objByCampId[row.campaign_id] || metaObjetivo(""), spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0, videoViews: 0, engajamentos: 0, records: [] };
      const c = byCamp[label];
      // NAO soma reach aqui: metricas aditivas (spend/impressoes/etc) somam por dia; reach vem do dedup abaixo
      c.spend += s.spend; c.impressions += s.impressions; c.clicks += s.clicks;
      c.revenue += s.revenue; c.purchases += s.purchases; c.leads += s.leads; c.addToCart += s.addToCart; c.initiateCheckout += s.initiateCheckout;
      c.conversas += s.conversas; c.videoViews += s.videoViews; c.engajamentos += s.engajamentos;
      if (m.daily) c.records.push({ date: row.date_start, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: s.reach, leads: s.leads, conversas: s.conversas, videoViews: s.videoViews, engajamentos: s.engajamentos });
    }
    // reach/frequencia DEDUPLICADO por campanha no periodo (fonte nao-diaria; quando m.daily=false, campRows ja e o dedup)
    const dedupSrc = (m.daily ? campDedup : campRows) as any[];
    for (const row of dedupSrc) {
      const label = row.campaign_name || "Meta Ads";
      const s = shape(row);
      if (byCamp[label]) byCamp[label].reach += s.reach; // soma so entre contas (audiencias distintas), nunca entre dias/anuncios
    }
    for (const row of adsetRows) {
      const label = (row.campaign_name || "Meta Ads") + " › " + (row.adset_name || "Conjunto");
      const s = shape(row);
      if (!byAdset[label]) byAdset[label] = { campaign: row.campaign_name || "", adset: row.adset_name || "", spend: 0, records: [] };
      const c = byAdset[label];
      c.spend += s.spend;
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
    c.frequency = c.reach ? c.impressions / c.reach : 0; // impressoes(periodo) / alcance dedup
    return c;
  }).sort((a: any, b: any) => b.spend - a.spend);
  ads.sort((a: any, b: any) => b.spend - a.spend);
  if (m.byAd && ads.length) {
    const topIds = ads.slice(0, 20).map((a: any) => a.adId).filter(Boolean);
    const thumbs = await fetchThumbsByIds(topIds);
    for (const a of ads) if (thumbs[a.adId]) a.thumbnail = thumbs[a.adId];
  }
  return { total, campaigns, adsets: Object.values(byAdset), adsDaily: Object.values(byAdDaily), ads, accounts, accountErrors, period: m.since && m.until ? { since: m.since, until: m.until } : { datePreset: m.datePreset || "last_30d" } };
}

// Saldo pré-pago da conta (pix/boleto): funding_source_details traz o saldo disponível
async function metaFunding(m: any) {
  const token = await _metaUserToken();
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

// Varre os criativos das contas do cliente e extrai domínios de site + números de WhatsApp usados nos anúncios.
function _bumpC(o: Record<string, number>, k: any) { k = String(k || "").trim(); if (!k) return; o[k] = (o[k] || 0) + 1; }
async function metaClientInfo(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const accs = (Array.isArray(m.accountIds) ? m.accountIds : []).map((x: any) => String(x).replace(/[^0-9]/g, "")).filter(Boolean);
  const siteCount: Record<string, number> = {}, waCount: Record<string, number> = {};
  const grabUrl = (u: any) => { const s = String(u || ""); const wa = s.match(/wa\.me\/(\d{8,15})|api\.whatsapp\.com\/send\/?\?phone=(\d{8,15})|whatsapp:.*?(\d{8,15})/i); if (wa) { _bumpC(waCount, wa[1] || wa[2] || wa[3]); return; } try { const h = new URL(s).hostname.replace(/^www\./, ""); if (h && !/facebook|instagram|fb\.me|fb\.com|whatsapp|l\.facebook/i.test(h)) _bumpC(siteCount, h); } catch { /* */ } };
  for (const acc of accs) {
    try {
      const r = await fetch(`https://graph.facebook.com/v21.0/act_${acc}/adcreatives?fields=link_url,object_story_spec&limit=250&access_token=${token}`);
      const j = await r.json();
      (j.data || []).forEach((cr: any) => {
        if (cr.link_url) grabUrl(cr.link_url);
        const ld = cr.object_story_spec && cr.object_story_spec.link_data;
        if (ld) { if (ld.link) grabUrl(ld.link); const cta = ld.call_to_action && ld.call_to_action.value; if (cta) { if (cta.link) grabUrl(cta.link); if (cta.app_link) grabUrl(cta.app_link); if (cta.whatsapp_number) _bumpC(waCount, String(cta.whatsapp_number).replace(/[^0-9]/g, "")); } }
      });
    } catch { /* */ }
  }
  const sites = Object.entries(siteCount).sort((a, b) => b[1] - a[1]).map((x) => x[0]).slice(0, 6);
  const whatsapps = Object.entries(waCount).sort((a, b) => b[1] - a[1]).map((x) => x[0]).slice(0, 6);
  return { sites, whatsapps };
}
// Lista os pixels das contas de anúncio do cliente (pra puxar automático no cadastro).
async function metaListPixels(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const accs = (Array.isArray(m.accountIds) ? m.accountIds : []).map((x: any) => String(x).replace(/[^0-9]/g, "")).filter(Boolean);
  const out: any[] = []; const seen = new Set<string>();
  for (const acc of accs) {
    try { const r = await fetch(`https://graph.facebook.com/v21.0/act_${acc}/adspixels?fields=id,name&limit=50&access_token=${token}`); const j = await r.json(); (j.data || []).forEach((p: any) => { if (p.id && !seen.has(p.id)) { seen.add(p.id); out.push({ id: p.id, name: p.name || p.id }); } }); } catch { /* */ }
  }
  return out;
}
async function metaListAccounts() {
  const token = await _metaUserToken();
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
  const token = await _metaUserToken();
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
  const token = await _metaUserToken();
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
  if (m.action === "rename") {
    const novo = String(m.novoNome || "").trim();
    if (novo.length < 2) throw new Error("Informe o novo nome (mínimo 2 caracteres).");
    await post(id, { name: novo });
    return { ok: true, detail: `Renomeado para "${novo}"` };
  }
  throw new Error("action invalida");
}
// Google Ads: pausar/reativar/renomear campanha. O sistema só sabia mexer em orçamento e palavras-chave —
// o resto obrigava a abrir o Gerenciador.
async function googleCampaignAction(m: any) {
  const cid = String(m.accountId || "").replace(/-/g, ""), campId = String(m.campaignId || "");
  if (!cid || !campId) throw new Error("Conta e campanha obrigatórias.");
  const token = await googleAdsAccessToken();
  const devToken = Deno.env.get("GOOGLE_ADS_DEV_TOKEN"), mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  const update: any = { resourceName: `customers/${cid}/campaigns/${campId}` };
  const mask: string[] = [];
  if (m.action === "pause" || m.action === "activate") { update.status = m.action === "pause" ? "PAUSED" : "ENABLED"; mask.push("status"); }
  else if (m.action === "rename") {
    const novo = String(m.novoNome || "").trim();
    if (novo.length < 2) throw new Error("Informe o novo nome (mínimo 2 caracteres).");
    update.name = novo; mask.push("name");
  } else throw new Error("action inválida para Google (use pause, activate ou rename).");
  const r = await fetch(`https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/campaigns:mutate`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "developer-token": devToken!, "login-customer-id": mcc, "Content-Type": "application/json" },
    body: JSON.stringify({ operations: [{ update, updateMask: mask.join(",") }] }),
  });
  const j = await r.json();
  if (j.error || !r.ok) throw new Error(j?.error?.details?.[0]?.errors?.[0]?.message || j?.error?.message || `HTTP ${r.status}`);
  const acao = m.action === "pause" ? "Pausada" : m.action === "activate" ? "Reativada" : `Renomeada para "${m.novoNome}"`;
  return { ok: true, detail: `${acao}: ${m.nome || campId}` };
}

// Clona a ESTRUTURA de uma campanha (campanha + conjuntos, PAUSADOS) pra OUTRA conta de anúncio (outro cliente).
// Não copia criativos/anúncios (são amarrados à conta de origem). Mapeia o pixel pro do destino quando possível.
async function metaCloneCampaign(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
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

// ===== PÚBLICOS (Custom/Saved Audiences do Meta) =====
function _audKind(subtype: string) { const s = String(subtype || "").toUpperCase(); if (s === "LOOKALIKE") return "lookalike"; if (s === "CUSTOM") return "custom"; if (["WEBSITE", "ENGAGEMENT", "VIDEO", "APP", "IG_BUSINESS", "OFFLINE_CONVERSION", "PAGE", "CLAIM", "BAG_OF_ACCOUNTS"].includes(s)) return "engagement"; return "custom"; }
// Lista os públicos (custom + salvos/interesse) das contas do cliente.
async function metaAudiences(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const accounts = (Array.isArray(m.accounts) ? m.accounts : []).map((a: any) => ({ id: String(a.id).replace(/^act_/, ""), name: a.name || "" }));
  if (!accounts.length) throw new Error("accounts obrigatorio");
  const out: any[] = []; const errors: string[] = [];
  const getJson = async (url: string) => { try { const r = await fetch(url); return await r.json(); } catch (e) { return { error: { message: String((e as any)?.message || e) } }; } };
  // custom + saved de TODAS as contas em paralelo
  await Promise.all(accounts.flatMap((acc: any) => [
    getJson(`${base}/act_${acc.id}/customaudiences?fields=id,name,subtype,approximate_count_lower_bound,approximate_count_upper_bound,retention_days,operation_status,description&limit=1000&access_token=${token}`).then((j: any) => {
      if (j && j.error) errors.push(`${acc.name || acc.id}: ${j.error.message}`);
      else if (j) (j.data || []).forEach((a: any) => out.push({ id: a.id, name: a.name || "(sem nome)", subtype: a.subtype || "", lo: a.approximate_count_lower_bound ?? null, hi: a.approximate_count_upper_bound ?? null, retention: a.retention_days ?? null, status: (a.operation_status && a.operation_status.description) || "", account: acc.name || acc.id, kind: _audKind(a.subtype) }));
    }),
    getJson(`${base}/act_${acc.id}/saved_audiences?fields=id,name,approximate_count_lower_bound,approximate_count_upper_bound&limit=500&access_token=${token}`).then((j: any) => {
      if (j && !j.error) (j.data || []).forEach((a: any) => out.push({ id: a.id, name: a.name || "(sem nome)", subtype: "SAVED", lo: a.approximate_count_lower_bound ?? null, hi: a.approximate_count_upper_bound ?? null, account: acc.name || acc.id, kind: "interesse" }));
    }),
  ]));
  return { audiences: out, errors: errors.length ? errors : undefined };
}
// Fontes pra criar públicos de engajamento: pixels, páginas, contas IG, vídeos, formulários.
async function metaAudienceSources(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const accounts = (Array.isArray(m.accounts) ? m.accounts : []).map((a: any) => ({ id: String(a.id).replace(/^act_/, ""), name: a.name || "" }));
  const out: any = { pixels: [], pages: [], igs: [], videos: [], forms: [] };
  const getJson = async (url: string) => { try { const r = await fetch(url); return await r.json(); } catch { return null; } };
  // Fase 1 — pixels + páginas + IG de TODAS as contas, tudo em paralelo
  await Promise.all(accounts.flatMap((acc: any) => [
    getJson(`${base}/act_${acc.id}/adspixels?fields=id,name&limit=50&access_token=${token}`).then((j: any) => { if (j && !j.error) (j.data || []).forEach((p: any) => out.pixels.push({ id: p.id, name: p.name || p.id, account: acc.id })); }),
    getJson(`${base}/act_${acc.id}/promote_pages?fields=id,name&limit=50&access_token=${token}`).then((j: any) => { if (j && !j.error) (j.data || []).forEach((p: any) => { if (!out.pages.some((x: any) => x.id === p.id)) out.pages.push({ id: p.id, name: p.name || p.id, account: acc.id }); }); }),
    getJson(`${base}/act_${acc.id}/instagram_accounts?fields=id,username,name&limit=50&access_token=${token}`).then((j: any) => { if (j && !j.error) (j.data || []).forEach((ig: any) => { if (ig.id && !out.igs.some((x: any) => x.id === ig.id)) out.igs.push({ id: ig.id, name: "@" + (ig.username || ig.name || ig.id), account: acc.id }); }); }),
  ]));
  // Algumas contas não expõem o Instagram em /instagram_accounts, mas expõem a Página promovida.
  await Promise.all(out.pages.map((pg: any) => getJson(`${base}/${pg.id}?fields=instagram_business_account{id,username,name}&access_token=${token}`).then((j: any) => {
    const ig = j && !j.error ? j.instagram_business_account : null;
    if (ig?.id && !out.igs.some((x: any) => String(x.id) === String(ig.id))) out.igs.push({ id: ig.id, name: "@" + (ig.username || ig.name || ig.id), account: pg.account, pageId: pg.id, pageName: pg.name || "" });
  })));
  // vídeos/formulários NÃO entram aqui (são pesados) — carregam sob demanda via metaAudienceMedia. Retorna já as páginas/IGs pra isso.
  return out;
}
// Carrega SOB DEMANDA (lazy) os vídeos (FB+IG) e formulários das páginas/IGs — chamado só quando o gestor abre a fonte Vídeo/Formulários.
async function metaAudienceMedia(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const pages = Array.isArray(m.pages) ? m.pages : [];
  const igs = Array.isArray(m.igs) ? m.igs : [];
  const out: any = { videos: [], forms: [] };
  const getJson = async (url: string) => { try { const r = await fetch(url); return await r.json(); } catch { return null; } };
  await Promise.all([
    ...pages.slice(0, 20).flatMap((pg: any) => [
      getJson(`${base}/${pg.id}/videos?fields=id,title,length,picture,permalink_url&limit=25&access_token=${token}`).then((j: any) => { if (j && !j.error) (j.data || []).forEach((v: any) => out.videos.push({ id: v.id, name: v.title || v.id, page: pg.id, account: pg.account, thumb: v.picture || "", url: v.permalink_url ? (String(v.permalink_url).startsWith("http") ? v.permalink_url : "https://www.facebook.com" + v.permalink_url) : ("https://www.facebook.com/" + v.id), len: v.length ? Math.round(v.length) : null, platform: "fb" })); }),
      getJson(`${base}/${pg.id}/leadgen_forms?fields=id,name&limit=50&access_token=${token}`).then((j: any) => { if (j && !j.error) (j.data || []).forEach((f: any) => out.forms.push({ id: f.id, name: f.name || f.id, page: pg.id, account: pg.account })); }),
    ]),
    ...igs.slice(0, 10).map((ig: any) => getJson(`${base}/${ig.id}/media?fields=id,caption,media_type,media_product_type,thumbnail_url,permalink&limit=50&access_token=${token}`).then((j: any) => { if (j && !j.error) (j.data || []).filter((x: any) => /VIDEO|REEL/i.test(x.media_type || "") || /REEL/i.test(x.media_product_type || "")).forEach((v: any) => out.videos.push({ id: v.id, name: v.caption ? String(v.caption).slice(0, 45) : (v.media_product_type === "REELS" ? "Reel" : "Vídeo IG"), page: ig.id, account: ig.account, thumb: v.thumbnail_url || "", url: v.permalink || "", platform: "ig" })); })),
  ]);
  return out;
}
// Cria em massa públicos de ENGAJAMENTO (site/pixel, página FB, IG, vídeo, formulário). Retorna resultado por item.
async function metaCreateAudiences(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const items = Array.isArray(m.audiences) ? m.audiences : [];
  if (!items.length) throw new Error("audiences obrigatório");
  const post = async (path: string, params: Record<string, string>) => { const r = await fetch(`${base}/${path}`, { method: "POST", body: new URLSearchParams({ ...params, access_token: token }) }); const j = await r.json(); if (j.error) throw new Error(j.error.message); return j; };
  const typeMap: Record<string, string> = { pixel: "pixel", facebook: "page", instagram: "ig_business", video: "video", form: "page" };
  const results: any[] = [];
  for (const a of items) {
    try {
      const acct = String(a.accountId || m.accountId || "").replace(/^act_/, ""); // cada público é criado na conta da sua fonte
      if (!acct) throw new Error("conta da fonte não definida");
      const secs = (Number(a.retentionDays) || 30) * 86400;
      let params: Record<string, string>;
      if (a.sourceType === "video") {
        // VÍDEO usa o formato LEGADO: lista de {event_name, object_id(vídeo), context_id(página)}
        const VID_EV: Record<string, string> = { P25: "video_view_25_percent", P50: "video_view_50_percent", P75: "video_view_75_percent", P100: "video_view_100_percent" };
        const ev = VID_EV[a.event] || "video_view_25_percent";
        const vids = Array.isArray(a.videos) ? a.videos : [];
        const legacy = vids.map((v: any) => ({ event_name: ev, object_id: String(v.id), context_id: String(v.page || "") })).filter((x: any) => x.object_id && x.context_id);
        if (!legacy.length) throw new Error("nenhum vídeo (ou página do vídeo) selecionado");
        params = { name: String(a.name).slice(0, 80), subtype: "ENGAGEMENT", retention_days: String(a.retentionDays || 30), rule: JSON.stringify(legacy) };
      } else {
        const rule: any = { event_sources: [{ type: typeMap[a.sourceType] || "page", id: String(a.sourceId) }], retention_seconds: secs };
        if (a.event && a.event !== "ALL") rule.filter = { operator: "and", filters: [{ field: "event", operator: "eq", value: a.event }] };
        const subtype = a.sourceType === "pixel" ? "WEBSITE" : "ENGAGEMENT";
        params = { name: String(a.name).slice(0, 80), subtype, retention_days: String(a.retentionDays || 30), rule: JSON.stringify({ inclusions: { operator: "or", rules: [rule] } }) };
        if (a.sourceType === "pixel") params.prefill = "true";
      }
      const j = await post(`act_${acct}/customaudiences`, params);
      results.push({ name: a.name, ok: true, id: j.id });
    } catch (e) { results.push({ name: a.name, ok: false, erro: String((e as any)?.message || e).slice(0, 200) }); }
  }
  return { results, criados: results.filter((r) => r.ok).length, falhas: results.filter((r) => !r.ok).length };
}
// Cria público CUSTOM a partir de lista (e-mails/telefones JÁ hasheados SHA-256 no front — LGPD).
async function metaCreateCustomList(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const acct = String(m.accountId || "").replace(/^act_/, "");
  const name = String(m.name || "Lista importada").slice(0, 80);
  const users: string[] = Array.isArray(m.usersHashed) ? m.usersHashed : [];
  const schema = /phone/i.test(m.schema || "") ? "PHONE" : "EMAIL";
  if (!acct || !users.length) throw new Error("accountId e usersHashed obrigatórios");
  const post = async (path: string, params: Record<string, string>) => { const r = await fetch(`${base}/${path}`, { method: "POST", body: new URLSearchParams({ ...params, access_token: token }) }); const j = await r.json(); if (j.error) throw new Error(j.error.message); return j; };
  const ca = await post(`act_${acct}/customaudiences`, { name, subtype: "CUSTOM", customer_file_source: "USER_PROVIDED_ONLY", description: String(m.description || "Importada via Central de Gestão").slice(0, 100) });
  let added = 0; const chunks: string[][] = [];
  for (let i = 0; i < users.length; i += 5000) chunks.push(users.slice(i, i + 5000));
  for (const ch of chunks) { try { await post(`${ca.id}/users`, { payload: JSON.stringify({ schema, data: ch }) }); added += ch.length; } catch (_e) { /* segue */ } }
  return { ok: true, id: ca.id, name, added, total: users.length };
}
// Cria público SALVO (interesse + geo + demografia).
async function metaCreateSavedAudience(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const acct = String(m.accountId || "").replace(/^act_/, "");
  const name = String(m.name || "Público de interesse").slice(0, 80);
  if (!acct || !m.targeting) throw new Error("accountId e targeting obrigatórios");
  const r = await fetch(`${base}/act_${acct}/saved_audiences`, { method: "POST", body: new URLSearchParams({ name, targeting: JSON.stringify(m.targeting), access_token: token }) });
  const j = await r.json(); if (j.error) throw new Error(j.error.message);
  return { ok: true, id: j.id, name };
}
// Busca de interesses/comportamentos pra segmentação (autocomplete do público de interesse).
async function metaTargetingSearch(m: any) {
  const token = await _metaUserToken(); if (!token) throw new Error("META_USER_TOKEN nao configurada");
  const base = "https://graph.facebook.com/v21.0";
  const q = String(m.q || "").trim(); if (!q) return { results: [] };
  const cls = m.type === "geo" ? "adgeolocations" : "adinterests";
  const url = m.type === "geo"
    ? `${base}/search?type=adgeolocation&location_types=["city","region","country"]&q=${encodeURIComponent(q)}&limit=15&access_token=${token}`
    : `${base}/search?type=adinterest&q=${encodeURIComponent(q)}&limit=20&access_token=${token}`;
  const r = await fetch(url); const j = await r.json(); if (j.error) throw new Error(j.error.message);
  return { results: (j.data || []).map((x: any) => ({ id: x.id || x.key, name: x.name, type: x.type, path: x.path, audience: x.audience_size_lower_bound || x.audience_size, key: x.key })) };
}

// Performance por segmentação (sexo / plataforma / posicionamento) — nível de conta, agregada entre contas
async function metaBreakdowns(m: any) {
  const token = await _metaUserToken();
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

// Sincroniza ativos Instagram pelo vínculo REAL da conta de anúncios Meta escolhida no cliente.
// É mais confiável que casar nomes. Nunca remove nem substitui escolha manual; só acrescenta ativos
// visíveis naquela(s) conta(s), respeitando a lista de exclusão feita pelo gestor.
async function metaAssetsSync(input: any = {}) {
  const dryRun = !!input.dryRun, only = Array.isArray(input.clientIds) ? new Set(input.clientIds.map(String)) : null;
  const clients = (await _sbAll("clients", "status=neq.Encerrado&meta_account_id=not.is.null&select=id,name,meta_account_id,instagram_accounts,instagram_accounts_excluded")).filter((c: any) => !only || only.has(String(c.id)));
  let clientsUpdated = 0, assetsAdded = 0; const results: any[] = [];
  for (const c of clients) {
    const ids = String(c.meta_account_id || "").split(",").map((x) => x.trim().replace(/^act_/, "")).filter(Boolean);
    if (!ids.length) continue;
    try {
      const src = await metaAudienceSources({ accounts: ids.map((id) => ({ id })) });
      const current: any[] = Array.isArray(c.instagram_accounts) ? c.instagram_accounts : [], known = new Set(current.map((x: any) => String(x.id)));
      const excluded = new Set((Array.isArray(c.instagram_accounts_excluded) ? c.instagram_accounts_excluded : []).map(String));
      const additions = (src.igs || []).filter((ig: any) => ig.id && !known.has(String(ig.id)) && !excluded.has(String(ig.id))).map((ig: any) => ({ id: String(ig.id), username: String(ig.name || "").replace(/^@/, ""), pagina: "", meta_account_id: String(ig.account || ""), auto_synced: true }));
      if (additions.length) {
        if (!dryRun) await sbPatchD("clients", `id=eq.${encodeURIComponent(c.id)}`, { instagram_accounts: [...current, ...additions] });
        clientsUpdated++; assetsAdded += additions.length;
      }
      results.push({ clientId: c.id, cliente: c.name, contasMeta: ids.length, encontrados: (src.igs || []).length, novos: additions.length, ativos: additions.map((x: any) => `@${x.username || x.id}`) });
    } catch (e) { results.push({ clientId: c.id, cliente: c.name, erro: String((e as any)?.message || e).slice(0, 180) }); }
  }
  return { dryRun, clientsChecked: clients.length, clientsUpdated, assetsAdded, results };
}

async function googleAudiences(m: any) {
  const ids = (Array.isArray(m.accounts) ? m.accounts : []).map((a: any) => String(a.id || a).replace(/\D/g, "")).filter(Boolean);
  if (!ids.length) throw new Error("Conta Google Ads obrigatória.");
  const token = await googleAdsAccessToken(); const audiences: any[] = [], errors: string[] = [];
  for (const accountId of ids) {
    try {
      const [lists, custom] = await Promise.all([
        gadsSearch(accountId, `SELECT user_list.id, user_list.name, user_list.description, user_list.type, user_list.membership_status, user_list.membership_life_span, user_list.size_for_display, user_list.size_for_search, user_list.eligible_for_display, user_list.eligible_for_search, user_list.read_only, user_list.access_reason, user_list.resource_name FROM user_list WHERE user_list.membership_status != 'CLOSED'`, token),
        gadsSearch(accountId, `SELECT custom_audience.id, custom_audience.name, custom_audience.description, custom_audience.type, custom_audience.status, custom_audience.resource_name, custom_audience.members FROM custom_audience WHERE custom_audience.status != 'REMOVED'`, token),
      ]);
      lists.forEach((r: any) => { const x = r.userList || {}; audiences.push({ accountId, resourceName: x.resourceName, id: String(x.id || ""), name: x.name || "Sem nome", description: x.description || "", type: x.type || "USER_LIST", source: "data", retention: Number(x.membershipLifeSpan) || null, sizeDisplay: Number(x.sizeForDisplay) || 0, sizeSearch: Number(x.sizeForSearch) || 0, eligibleDisplay: !!x.eligibleForDisplay, eligibleSearch: !!x.eligibleForSearch, readOnly: !!x.readOnly, accessReason: x.accessReason || "" }); });
      custom.forEach((r: any) => { const x = r.customAudience || {}; audiences.push({ accountId, resourceName: x.resourceName, id: String(x.id || ""), name: x.name || "Sem nome", description: x.description || "", type: x.type || "CUSTOM", source: "custom", members: x.members || [], eligibleDisplay: true, eligibleSearch: false }); });
    } catch (e) { errors.push(`${accountId}: ${(e as any)?.message || e}`); }
  }
  return { audiences, errors };
}
async function googleCreateCustomAudience(m: any) {
  const cid = String(m.accountId || "").replace(/\D/g, ""), name = String(m.name || "").trim();
  const terms = (Array.isArray(m.terms) ? m.terms : []).map((x: any) => String(x).trim()).filter(Boolean).slice(0, 1000);
  const urls = (Array.isArray(m.urls) ? m.urls : []).map((x: any) => String(x).trim()).filter(Boolean).slice(0, 1000);
  if (!cid || !name || (!terms.length && !urls.length)) throw new Error("Conta, nome e termos ou URLs são obrigatórios.");
  const token = await googleAdsAccessToken(), devToken = Deno.env.get("GOOGLE_ADS_DEV_TOKEN"), mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  const members = [...terms.map((keyword: string) => ({ keyword })), ...urls.map((url: string) => ({ url }))];
  const r = await fetch(`https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/customAudiences:mutate`, { method: "POST", headers: { Authorization: `Bearer ${token}`, "developer-token": devToken!, "login-customer-id": mcc, "Content-Type": "application/json" }, body: JSON.stringify({ operations: [{ create: { name, description: String(m.description || "Criado pela Central de Gestão"), type: "SEARCH", status: "ENABLED", members } }] }) });
  const j = await r.json(); if (!r.ok || j.error) throw new Error(j?.error?.details?.[0]?.errors?.[0]?.message || j?.error?.message || "Erro ao criar segmento no Google Ads.");
  return { ok: true, name, resourceName: j.results?.[0]?.resourceName || "", members: members.length };
}

// Classifica uma AÇÃO de conversão do Google (form, WhatsApp, ligação, compra...) no balde certo.
// É o evento que o gestor marca dentro do Google — cada campanha tem o seu.
function _gConvBucket(name: string, category: string): "purchases" | "conversas" | "leads" | "outros" {
  const nm = String(name || "").toLowerCase();
  const cat = String(category || "").toUpperCase();
  if (/whats|wpp|\bzap\b|mensag|message|\bchat\b|conversa|direct|\bdm\b/.test(nm)) return "conversas";
  if (cat === "PURCHASE" || /compra|purchase|venda|\bsale\b|checkout|pedido|receita|revenue|e-?commerce/.test(nm)) return "purchases";
  // Acoes que o Google conta como conversao mas NAO sao contato de lead — contavam como lead e inflavam o CPL
  // (numa conta real, 124 dos 675 "leads" eram visita a loja e view/inscricao do YouTube).
  if (/store visit|visita.*loja|follow-?on view|visualiza.*subsequen|channel subscri|inscri.*canal|\bsubscription/.test(nm)) return "outros";
  return "leads"; // form, contato, orçamento, cadastro, ligação, agendamento, etc.
}
// Objetivo por canal, mas com o TIPO vindo do evento de conversão dominante (quando houver)
function googleObjetivoConv(channelType: string, buckets?: { purchases: number; leads: number; conversas: number } | null) {
  const base = googleObjetivo(channelType);
  if (!buckets) return base;
  const { purchases, leads, conversas } = buckets;
  const tot = purchases + leads + conversas;
  if (tot <= 0) return base;
  let tipo = "leads"; let max = leads;
  if (conversas > max) { tipo = "mensagens"; max = conversas; }
  if (purchases > max) { tipo = "conversao"; max = purchases; }
  const rot: Record<string, string> = { conversao: "Vendas/Conversão", mensagens: "Mensagens (WhatsApp)", leads: "Leads" };
  return { ...base, tipo, rotulo: `${base.rotulo} · ${rot[tipo] || tipo}` };
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
  // Mesmo total, agora somando as CAMPANHAS — e o que o Gerenciador mostra na linha "Total: conta" (conferido
  // numa conta real: campanhas 94.765.087 impressoes = tela do Google; recurso `customer` dava 96.126.960, ~1,4%
  // a mais de entrega solta que nao esta em nenhuma campanha atual). O painel usa este; o da conta vai em `_conta`.
  const campAgg: any = { spend: 0, impressions: 0, clicks: 0, revenue: 0, purchases: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0 };
  const campRecByDate: Record<string, any> = {};
  const byCamp: Record<string, any> = {};
  const byAdset: Record<string, any> = {};
  const byAdDaily: Record<string, any> = {}; // anuncio x dia (banco de dados de midia — schema `midia`) - so populado quando byAd+daily juntos
  const ads: any[] = [];

  const perAccount = await Promise.all(accounts.map(async (acc) => {
    try {
      const [accountRows, acctDaily, campRows, adRows, campConvRows, adConvRows, adsetRows, adDailyRows] = await Promise.all([
        gadsSearch(acc.id, `SELECT ${GADS_METRICS} FROM customer WHERE ${range}`, token),
        g.daily ? gadsSearch(acc.id, `SELECT segments.date, ${GADS_METRICS} FROM customer WHERE ${range}`, token) : Promise.resolve([] as any[]),
        // Colunas de CONFIGURAÇÃO da campanha, as mesmas que o Gerenciador mostra (status, tipo, estratégia de
        // lance, pontuação de otimização, datas) + parcela de impressões perdida por orçamento/classificação.
        g.byCampaign ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, campaign.advertising_channel_type, campaign.status, campaign.bidding_strategy_type, campaign.optimization_score, campaign_budget.amount_micros, campaign_budget.resource_name${g.daily ? ", segments.date" : ""}, ${GADS_METRICS_FULL}, metrics.search_impression_share, metrics.search_budget_lost_impression_share, metrics.search_rank_lost_impression_share FROM campaign WHERE ${range}`, token) : Promise.resolve([] as any[]),
        g.byAd ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, campaign.advertising_channel_type, ad_group.id, ad_group.name, ad_group_ad.ad.id, ad_group_ad.ad.name, ad_group_ad.ad.type, ${GADS_METRICS_FULL} FROM ad_group_ad WHERE ${range} AND metrics.cost_micros > 0`, token) : Promise.resolve([] as any[]),
        // quebra das conversões por AÇÃO (form, WhatsApp, ligação, compra...) — por campanha e por anúncio
        (g.byCampaign || g.byAd) ? gadsSearch(acc.id, `SELECT campaign.id, segments.conversion_action_name, segments.conversion_action_category, metrics.conversions FROM campaign WHERE ${range} AND metrics.conversions > 0`, token).catch(() => [] as any[]) : Promise.resolve([] as any[]),
        g.byAd ? gadsSearch(acc.id, `SELECT ad_group_ad.ad.id, segments.conversion_action_name, segments.conversion_action_category, metrics.conversions FROM ad_group_ad WHERE ${range} AND metrics.conversions > 0`, token).catch(() => [] as any[]) : Promise.resolve([] as any[]),
        g.byAdset ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, ad_group.id, ad_group.name${g.daily ? ", segments.date" : ""}, ${GADS_METRICS_FULL} FROM ad_group WHERE ${range}`, token) : Promise.resolve([] as any[]),
        // anuncio x dia — so quando pedido explicitamente (banco de dados de midia); nenhum outro caller usa byAd+daily juntos hoje
        (g.byAd && g.daily) ? gadsSearch(acc.id, `SELECT campaign.id, campaign.name, ad_group.id, ad_group.name, ad_group_ad.ad.id, ad_group_ad.ad.name, ad_group_ad.ad.type, segments.date, ${GADS_METRICS_FULL} FROM ad_group_ad WHERE ${range} AND metrics.cost_micros > 0`, token) : Promise.resolve([] as any[]),
      ]);
      return { acc, accountRows, acctDaily, campRows, adRows, campConvRows, adConvRows, adsetRows, adDailyRows, error: null as string | null };
    } catch (e) {
      return { acc, accountRows: [] as any[], acctDaily: [] as any[], campRows: [] as any[], adRows: [] as any[], campConvRows: [] as any[], adConvRows: [] as any[], adsetRows: [] as any[], adDailyRows: [] as any[], error: (e as any)?.message || String(e) };
    }
  }));
  const accountErrors = perAccount.filter((p) => p.error).map((p) => ({ id: p.acc.id, name: p.acc.name || p.acc.id, error: p.error }));

  // Mapas de conversão por AÇÃO → baldes (purchases/leads/conversas) + detalhamento por ação, por campanha e por anúncio
  const convByCamp: Record<string, { purchases: number; leads: number; conversas: number; outros: number; acts: Record<string, number> }> = {};
  const convByAd: Record<string, { purchases: number; leads: number; conversas: number; outros: number; acts: Record<string, number> }> = {};
  const _accConv = (map: any, key: string, name: string, cat: string, v: number) => {
    if (!key || !(v > 0)) return;
    const b = map[key] || (map[key] = { purchases: 0, leads: 0, conversas: 0, outros: 0, acts: {} });
    b[_gConvBucket(name, cat)] += v; const an = name || "Conversão"; b.acts[an] = (b.acts[an] || 0) + v;
  };
  for (const { campConvRows, adConvRows } of perAccount) {
    for (const row of (campConvRows || [])) _accConv(convByCamp, row.campaign?.id ? String(row.campaign.id) : "", row.segments?.conversionActionName, row.segments?.conversionActionCategory, Number(row.metrics?.conversions) || 0);
    for (const row of (adConvRows || [])) _accConv(convByAd, row.adGroupAd?.ad?.id ? String(row.adGroupAd.ad.id) : "", row.segments?.conversionActionName, row.segments?.conversionActionCategory, Number(row.metrics?.conversions) || 0);
  }
  const _actList = (acts: Record<string, number>) => Object.entries(acts).map(([name, count]) => ({ name, count: Math.round(count), bucket: _gConvBucket(name, "") })).sort((a, b) => b.count - a.count);

  for (const { acc, accountRows, acctDaily, campRows, adRows, adsetRows, adDailyRows } of perAccount) {
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
      if (!byCamp[label]) {
        const _cb = row.campaign?.id ? convByCamp[String(row.campaign.id)] : null;
        // leads/conversas vem da quebra por ACAO de conversao (total do periodo, nao por linha): setados na criacao
        // e nunca somados de novo, senao cada dia da campanha multiplicaria o mesmo numero.
        byCamp[label] = { campaign: label, campaignId: row.campaign?.id ? String(row.campaign.id) : null, account: acc.name || acc.id, accountId: acc.id, objetivo: googleObjetivoConv(row.campaign?.advertisingChannelType, _cb ? { purchases: _cb.purchases, leads: _cb.leads, conversas: _cb.conversas } : null), _google: true, orcamentoDiario: row.campaignBudget?.amountMicros ? +row.campaignBudget.amountMicros / 1e6 : null, budgetResource: row.campaignBudget?.resourceName || null,
          // config da campanha (mesmas colunas do Gerenciador) — vem igual em toda linha da campanha, então grava na criação
          status: row.campaign?.status || null, tipoCampanha: row.campaign?.advertisingChannelType || null,
          estrategiaLance: row.campaign?.biddingStrategyType || null,
          pontuacaoOtimizacao: row.campaign?.optimizationScore != null ? +(Number(row.campaign.optimizationScore) * 100).toFixed(1) : null,
          impShare: null as number | null, impPerdidaOrcamento: null as number | null, impPerdidaRank: null as number | null,
          spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: _cb ? _cb.leads : 0, addToCart: 0, initiateCheckout: 0, conversas: _cb ? _cb.conversas : 0, videoViews: 0, engajamentos: 0, records: [] };
        campAgg.leads += _cb ? _cb.leads : 0; campAgg.conversas += _cb ? _cb.conversas : 0;
      }
      const c = byCamp[label];
      // parcela de impressões é percentual: pega o maior valor visto no período (não soma)
      const _pct = (v: any) => v == null ? null : +(Number(v) * 100).toFixed(1);
      const _mx = (a: any, b: any) => b == null ? a : (a == null ? b : Math.max(a, b));
      c.impShare = _mx(c.impShare, _pct(row.metrics?.searchImpressionShare));
      c.impPerdidaOrcamento = _mx(c.impPerdidaOrcamento, _pct(row.metrics?.searchBudgetLostImpressionShare));
      c.impPerdidaRank = _mx(c.impPerdidaRank, _pct(row.metrics?.searchRankLostImpressionShare));
      c.spend += s.spend; c.impressions += s.impressions; c.clicks += s.clicks;
      c.revenue += s.revenue; c.purchases += s.purchases; c.videoViews += s.videoViews; c.engajamentos += s.engajamentos;
      campAgg.spend += s.spend; campAgg.impressions += s.impressions; campAgg.clicks += s.clicks;
      campAgg.revenue += s.revenue; campAgg.purchases += s.purchases; campAgg.videoViews += s.videoViews; campAgg.engajamentos += s.engajamentos;
      if (g.daily && row.segments?.date) {
        c.records.push({ date: row.segments.date, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: 0, leads: 0, conversas: 0, videoViews: s.videoViews, engajamentos: s.engajamentos });
        const k = row.segments.date;
        const rec = campRecByDate[k] || (campRecByDate[k] = { date: k, sales: 0, spend: 0, revenue: 0, clicks: 0, impressions: 0, reach: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0, addToCart: 0, checkout: 0 });
        rec.sales += Math.round(s.purchases); rec.spend += s.spend; rec.revenue += s.revenue; rec.clicks += s.clicks; rec.impressions += s.impressions; rec.videoViews += s.videoViews; rec.engajamentos += s.engajamentos;
      }
    }
    for (const row of adRows) {
      const s = gadsShape(row.metrics);
      const ad = row.adGroupAd?.ad || {};
      const adName = ad.name || (ad.type ? String(ad.type).replace(/_/g, " ").toLowerCase() : "anúncio") + " #" + (ad.id || "");
      const campId = row.campaign?.id ? String(row.campaign.id) : "";
      const ab = convByAd[ad.id ? String(ad.id) : ""] || null; // quebra por ação deste anúncio
      ads.push({
        adId: ad.id ? "g" + ad.id : null, adName, campaign: row.campaign?.name || "", campaignId: campId || null,
        adset: row.adGroup?.name || "", adsetId: row.adGroup?.id ? String(row.adGroup.id) : null,
        account: acc.name || acc.id, thumbnail: null, _google: true,
        objetivo: googleObjetivoConv(row.campaign?.advertisingChannelType, ab ? { purchases: ab.purchases, leads: ab.leads, conversas: ab.conversas } : null),
        spend: s.spend, impressions: s.impressions, clicks: s.clicks, reach: 0, frequency: 0,
        ctr: s.ctr, cpc: s.cpc, cpm: s.cpm, purchases: s.purchases, revenue: s.revenue, roas: s.roas,
        // leads/conversas saem da quebra por acao (Lead, Clicks to call, Conversation started...). Ficavam fixos em
        // 0 aqui, e como o KPI de Leads do painel e montado a partir destas linhas, o Google nunca mostrava lead.
        leads: ab ? ab.leads : 0, addToCart: 0, initiateCheckout: 0, conversas: ab ? ab.conversas : 0, videoViews: s.videoViews, engajamentos: s.engajamentos,
        convActions: ab ? _actList(ab.acts) : undefined,
        cpa: s.purchases ? s.spend / s.purchases : 0,
      });
    }
    for (const row of (adsetRows || [])) {
      const label = (row.campaign?.name || "Google Ads") + " › " + (row.adGroup?.name || "Grupo de anúncios");
      const s = gadsShape(row.metrics);
      if (!byAdset[label]) byAdset[label] = { campaign: row.campaign?.name || "", campaignId: row.campaign?.id ? String(row.campaign.id) : null, adset: row.adGroup?.name || "", spend: 0, records: [] };
      const c = byAdset[label];
      c.spend += s.spend;
      if (g.daily && row.segments?.date) c.records.push({ date: row.segments.date, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: 0, leads: 0, conversas: 0, videoViews: s.videoViews, engajamentos: s.engajamentos });
    }
    // anuncio x dia (banco de dados de midia) — mesma logica do adRows, so que 1 linha por dia em vez de total do periodo
    for (const row of (adDailyRows || [])) {
      const s = gadsShape(row.metrics);
      const ad = row.adGroupAd?.ad || {};
      const adId = ad.id ? String(ad.id) : "";
      if (!adId || !row.segments?.date) continue;
      const adName = ad.name || (ad.type ? String(ad.type).replace(/_/g, " ").toLowerCase() : "anúncio") + " #" + adId;
      const campId = row.campaign?.id ? String(row.campaign.id) : "";
      if (!byAdDaily[adId]) byAdDaily[adId] = {
        adId: "g" + adId, adName, campaign: row.campaign?.name || "", campaignId: campId || null,
        adset: row.adGroup?.name || "", adsetId: row.adGroup?.id ? String(row.adGroup.id) : null,
        account: acc.name || acc.id, accountId: acc.id,
        objetivo: googleObjetivoConv(row.campaign?.advertisingChannelType, null), records: [],
      };
      // leads/conversas ficam 0 aqui (mesma decisao ja tomada pro adsetRows/campRows diario): a quebra por acao de conversao
      // (convByAd) so vem em total-de-periodo, nao por dia — atribuir por dia exigiria uma chamada GAQL nova, fora do escopo desta fase.
      byAdDaily[adId].records.push({ date: row.segments.date, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions, reach: 0, frequency: 0, leads: 0, conversas: 0, videoViews: s.videoViews, engajamentos: s.engajamentos });
    }
  }
  // Total do canal = soma das campanhas quando pedimos a quebra por campanha (e o que o Gerenciador mostra);
  // sem byCampaign, cai pro total do recurso `customer`.
  const usaCamp = !!g.byCampaign && campAgg.impressions > 0;
  const base: any = usaCamp ? campAgg : totAgg;
  const recs = (usaCamp && Object.keys(campRecByDate).length) ? campRecByDate : totRecByDate;
  const total = {
    ...totAgg, ...base,
    ctr: base.impressions ? (base.clicks / base.impressions) * 100 : 0,
    cpc: base.clicks ? base.spend / base.clicks : 0,
    cpm: base.impressions ? (base.spend / base.impressions) * 1000 : 0,
    roas: base.spend ? base.revenue / base.spend : 0,
    records: Object.values(recs).sort((a: any, b: any) => a.date < b.date ? -1 : 1),
    // total do recurso `customer`: inclui entrega que nao esta presa a nenhuma campanha atual. So pra conferencia.
    _conta: { spend: totAgg.spend, impressions: totAgg.impressions, clicks: totAgg.clicks, purchases: totAgg.purchases, revenue: totAgg.revenue },
  };
  const campaigns = Object.values(byCamp).map((c: any) => {
    c.ctr = c.impressions ? (c.clicks / c.impressions) * 100 : 0;
    c.cpc = c.clicks ? c.spend / c.clicks : 0;
    c.cpm = c.impressions ? (c.spend / c.impressions) * 1000 : 0;
    c.roas = c.spend ? c.revenue / c.spend : 0;
    c.cpa = c.purchases ? c.spend / c.purchases : 0;
    // anexa a QUEBRA por ação de conversão (form/WhatsApp/compra...) — número total de conversões fica em purchases
    const cb = c.campaignId ? convByCamp[c.campaignId] : null;
    if (cb) c.convActions = _actList(cb.acts);
    return c;
  }).sort((a: any, b: any) => b.spend - a.spend);
  // quebra total do canal por ação (pros relatórios/resumo)
  { const totActs: Record<string, number> = {};
    for (const k of Object.keys(convByCamp)) for (const [n, v] of Object.entries(convByCamp[k].acts)) totActs[n] = (totActs[n] || 0) + v;
    if (Object.keys(totActs).length) (total as any).convActions = _actList(totActs);
  }
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
        purchases: c.purchases || 0, revenue: c.revenue, roas: c.spend ? c.revenue / c.spend : 0,
        leads: c.leads || 0, addToCart: 0, initiateCheckout: 0, conversas: c.conversas || 0, videoViews: c.videoViews || 0, engajamentos: c.engajamentos || 0,
        convActions: c.convActions, cpa: c.purchases ? c.spend / c.purchases : 0,
      });
    }
  }
  // MESMA sintese acima, agora pro nivel "conjunto" (alimenta channel_metrics_daily/Banco de Dados) — sem isso, uma
  // campanha Performance Max/Demand Gen com gasto real simplesmente nao aparecia em nenhum grafico por dia.
  if (g.byAdset && g.byCampaign) {
    const comAdset = new Set(Object.values(byAdset).map((a: any) => a.campaignId).filter(Boolean));
    for (const c of Object.values(byCamp) as any[]) {
      if ((c.spend || 0) <= 0) continue;
      if (c.campaignId && comAdset.has(c.campaignId)) continue;
      const label = c.campaign + " › (campanha inteira — Performance Max/Demand Gen)";
      byAdset[label] = { campaign: c.campaign, campaignId: c.campaignId, adset: "(campanha inteira)", spend: c.spend, records: g.daily ? c.records : [] };
    }
  }
  // MESMA sintese, agora pro nivel "anuncio x dia" (banco de dados de midia, schema `midia`) — mesmo motivo.
  if (g.byAd && g.daily && g.byCampaign) {
    const comAdDaily = new Set(Object.values(byAdDaily).map((a: any) => a.campaignId).filter(Boolean));
    for (const c of Object.values(byCamp) as any[]) {
      if ((c.spend || 0) <= 0) continue;
      if (c.campaignId && comAdDaily.has(c.campaignId)) continue;
      const adId = "pmax_" + c.campaignId; // sintetico e estavel: mesmo id em toda sincronizacao, nao duplica no upsert
      byAdDaily[adId] = {
        adId, adName: "(campanha inteira — Performance Max/Demand Gen)", campaign: c.campaign, campaignId: c.campaignId,
        adset: "(campanha inteira)", adsetId: "grp_" + adId, account: c.account, accountId: c.accountId,
        objetivo: c.objetivo, records: c.records || [],
      };
    }
  }
  // COBERTURA PARCIAL: campanha de Video (YouTube) atribui so uma fatia do custo no nivel de anuncio —
  // ex: R$516 na campanha, R$19 somando os anuncios. As sinteses acima so cobrem campanha SEM nenhuma linha;
  // aqui entra a linha "(restante no nivel da campanha)" com a DIFERENCA, pra soma sempre bater com o gerenciador.
  if (g.byAd && g.byCampaign) {
    const sums: Record<string, any> = {};
    for (const a of ads) { if (!a.campaignId) continue; const s = sums[a.campaignId] ||= { spend: 0, impressions: 0, clicks: 0, purchases: 0, revenue: 0, videoViews: 0, engajamentos: 0 }; s.spend += a.spend || 0; s.impressions += a.impressions || 0; s.clicks += a.clicks || 0; s.purchases += a.purchases || 0; s.revenue += a.revenue || 0; s.videoViews += a.videoViews || 0; s.engajamentos += a.engajamentos || 0; }
    for (const c of Object.values(byCamp) as any[]) {
      if (!(c.spend > 0) || !c.campaignId) continue;
      const s = sums[c.campaignId]; if (!s || !(s.spend > 0)) continue; // sem nenhuma linha: a sintese de campanha inteira ja cobriu
      const diff = c.spend - s.spend;
      if (diff <= Math.max(0.05, c.spend * 0.005)) continue; // diferenca de arredondamento: ignora
      const dImp = Math.max(0, (c.impressions || 0) - s.impressions), dClk = Math.max(0, (c.clicks || 0) - s.clicks);
      ads.push({
        adId: "gcr" + c.campaignId, adName: "(restante no nível da campanha — não atribuído por anúncio)",
        campaign: c.campaign, campaignId: c.campaignId, adset: "", adsetId: null,
        account: c.account, thumbnail: null, _google: true, _campaignLevel: true, objetivo: c.objetivo,
        spend: +diff.toFixed(2), impressions: dImp, clicks: dClk, reach: 0, frequency: 0,
        ctr: dImp ? (dClk / dImp) * 100 : 0, cpc: dClk ? diff / dClk : 0, cpm: dImp ? (diff / dImp) * 1000 : 0,
        purchases: Math.max(0, (c.purchases || 0) - s.purchases), revenue: Math.max(0, +(((c.revenue || 0) - s.revenue)).toFixed(2)),
        roas: 0, leads: 0, addToCart: 0, initiateCheckout: 0, conversas: 0,
        videoViews: Math.max(0, (c.videoViews || 0) - s.videoViews), engajamentos: Math.max(0, (c.engajamentos || 0) - s.engajamentos),
        convActions: undefined, cpa: 0,
      });
    }
  }
  // Mesmo remendo pro nivel conjunto x dia (alimenta channel_metrics_daily): remainder POR DATA.
  if (g.byAdset && g.byCampaign && g.daily) {
    const sumByCampDate: Record<string, Record<string, any>> = {};
    for (const as of Object.values(byAdset) as any[]) {
      if (!as.campaignId) continue;
      const m = sumByCampDate[as.campaignId] ||= {};
      for (const r of (as.records || [])) { const d = m[r.date] ||= { spend: 0, impressions: 0, clicks: 0, sales: 0, revenue: 0, videoViews: 0, engajamentos: 0 }; d.spend += r.spend || 0; d.impressions += r.impressions || 0; d.clicks += r.clicks || 0; d.sales += r.sales || 0; d.revenue += r.revenue || 0; d.videoViews += r.videoViews || 0; d.engajamentos += r.engajamentos || 0; }
    }
    for (const c of Object.values(byCamp) as any[]) {
      if (!(c.spend > 0) || !c.campaignId) continue;
      const m = sumByCampDate[c.campaignId]; if (!m) continue; // sem nenhuma linha de conjunto: sintese de campanha inteira ja cobriu
      const recs: any[] = []; let extra = 0;
      for (const r of (c.records || [])) {
        const s = m[r.date] || { spend: 0, impressions: 0, clicks: 0, sales: 0, revenue: 0, videoViews: 0, engajamentos: 0 };
        const dspend = (r.spend || 0) - s.spend;
        if (dspend > 0.01) { recs.push({ date: r.date, spend: +dspend.toFixed(2), sales: Math.max(0, (r.sales || 0) - s.sales), revenue: Math.max(0, +(((r.revenue || 0) - s.revenue)).toFixed(2)), clicks: Math.max(0, (r.clicks || 0) - s.clicks), impressions: Math.max(0, (r.impressions || 0) - s.impressions), reach: 0, leads: 0, conversas: 0, videoViews: Math.max(0, (r.videoViews || 0) - s.videoViews), engajamentos: Math.max(0, (r.engajamentos || 0) - s.engajamentos) }); extra += dspend; }
      }
      if (recs.length && extra > Math.max(0.05, c.spend * 0.005)) {
        const label = c.campaign + " › (restante no nível da campanha)";
        if (!byAdset[label]) byAdset[label] = { campaign: c.campaign, campaignId: c.campaignId, adset: "(restante no nível da campanha)", spend: +extra.toFixed(2), records: recs };
      }
    }
  }
  // E pro nivel anuncio x dia (schema midia): remainder POR DATA com id sintetico estavel (nao duplica no upsert).
  if (g.byAd && g.daily && g.byCampaign) {
    const spendByCampDate: Record<string, Record<string, number>> = {};
    for (const ad of Object.values(byAdDaily) as any[]) {
      if (!ad.campaignId) continue;
      const m = spendByCampDate[ad.campaignId] ||= {};
      for (const r of (ad.records || [])) m[r.date] = (m[r.date] || 0) + (r.spend || 0);
    }
    for (const c of Object.values(byCamp) as any[]) {
      if (!(c.spend > 0) || !c.campaignId) continue;
      const m = spendByCampDate[c.campaignId]; if (!m) continue; // sem nenhuma linha: sintese pmax_ ja cobriu
      const recs: any[] = [];
      for (const r of (c.records || [])) {
        const dspend = (r.spend || 0) - (m[r.date] || 0);
        if (dspend > 0.01) recs.push({ date: r.date, spend: +dspend.toFixed(2), sales: 0, revenue: 0, clicks: 0, impressions: 0, reach: 0, frequency: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0 });
      }
      if (recs.length) {
        const adId = "resto_" + c.campaignId;
        byAdDaily[adId] = { adId, adName: "(restante no nível da campanha)", campaign: c.campaign, campaignId: c.campaignId, adset: "(restante no nível da campanha)", adsetId: "grp_" + adId, account: c.account, accountId: c.accountId, objetivo: c.objetivo, records: recs };
      }
    }
  }
  ads.sort((a: any, b: any) => b.spend - a.spend);
  return { total, campaigns, adsets: Object.values(byAdset), adsDaily: Object.values(byAdDaily), ads, accounts, accountErrors, period: { since, until } };
}

// ===== TikTok Ads =====
// Conexão é POR CLIENTE (cada um conecta a própria conta, igual RD Station/Nuvemshop) — não tem MCC como Meta/Google.
// O access_token do TikTok não expira sozinho (fica válido até o usuário revogar), então não precisa de refresh_token.
// ⚠️ Fase 1 (sem credenciais reais testadas ainda): shape da API (endpoints/campos) seguido pela documentação oficial —
// pode precisar de ajuste fino assim que a 1ª conta real conectar (ex.: nome exato de algum campo do relatório).
async function tiktokListAccounts(clientId: string) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=tiktok_config&limit=1`))[0];
  const cfg = c?.tiktok_config || {};
  if (!cfg.access_token) return { error: "Cliente não conectou o TikTok Ads ainda." };
  const acc = await sbGet("account_config", "id=eq.main&select=data");
  const tk = (acc[0]?.data || {}).tiktok_ads || {};
  if (!tk.app_id || !tk.secret) return { error: "Faltam as credenciais do App do TikTok em Configurações." };
  // /oauth2/advertiser/get/ devolve as contas de anúncio que o cliente autorizou pro nosso app (escopo "Ad Account Management").
  const p = new URLSearchParams({ app_id: tk.app_id, secret: tk.secret });
  const r = await fetch(`https://business-api.tiktok.com/open_api/v1.3/oauth2/advertiser/get/?${p}`, { headers: { "Access-Token": cfg.access_token } });
  const j = await r.json();
  if (j.code !== 0) return { error: "TikTok: " + (j.message || "erro ao listar contas") };
  const list = (j?.data?.list || []).map((a: any) => ({ id: a.advertiser_id, name: a.advertiser_name || a.advertiser_id }));
  return { accounts: list };
}
async function tiktokAdsInsights(m: any) {
  const clientId = m.clientId;
  if (!clientId) throw new Error("clientId obrigatório");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=tiktok_config&limit=1`))[0];
  const cfg = c?.tiktok_config || {};
  if (!cfg.access_token) throw new Error("Cliente não conectou o TikTok Ads.");
  const advertiserId = m.advertiserId || cfg.advertiser_id || (cfg.advertiser_ids || [])[0];
  if (!advertiserId) throw new Error("Nenhuma conta de anúncio TikTok selecionada pra esse cliente.");
  const since = m.since, until = m.until;
  if (!since || !until) throw new Error("since e until obrigatórios (YYYY-MM-DD)");
  const body = {
    advertiser_id: advertiserId, report_type: "BASIC", data_level: "AUCTION_CAMPAIGN",
    dimensions: ["campaign_id"], start_date: since, end_date: until,
    metrics: ["campaign_name", "spend", "impressions", "clicks", "conversion", "ctr", "cpc", "cpm"],
    page_size: 200,
  };
  const r = await fetch("https://business-api.tiktok.com/open_api/v1.3/report/integrated/get/", {
    method: "POST", headers: { "Access-Token": cfg.access_token, "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  const j = await r.json();
  if (j.code !== 0) throw new Error("TikTok Ads: " + (j.message || "erro na API"));
  const rows = j?.data?.list || [];
  const totAgg = { spend: 0, impressions: 0, clicks: 0, purchases: 0 };
  const campaigns = rows.map((row: any) => {
    const d = row.metrics || {};
    const spend = parseFloat(d.spend || "0"), impressions = parseInt(d.impressions || "0", 10) || 0, clicks = parseInt(d.clicks || "0", 10) || 0, conv = parseFloat(d.conversion || "0");
    totAgg.spend += spend; totAgg.impressions += impressions; totAgg.clicks += clicks; totAgg.purchases += conv;
    return { campaign: d.campaign_name || row.dimensions?.campaign_id || "TikTok Ads", campaignId: row.dimensions?.campaign_id || null, spend, impressions, clicks, purchases: conv, ctr: parseFloat(d.ctr || "0"), cpc: parseFloat(d.cpc || "0"), cpm: parseFloat(d.cpm || "0"), _tiktok: true };
  });
  const total = { ...totAgg, ctr: totAgg.impressions ? (totAgg.clicks / totAgg.impressions) * 100 : 0, cpc: totAgg.clicks ? totAgg.spend / totAgg.clicks : 0, cpm: totAgg.impressions ? (totAgg.spend / totAgg.impressions) * 1000 : 0 };
  return { total, campaigns, _tiktok: true, period: { since, until } };
}

// ===== Pinterest Ads =====
// Access token do Pinterest expira — sempre renova pelo refresh_token antes de chamar a API (evita guardar estado de expiração).
async function pinterestAccessToken(clientId: string): Promise<{ token: string; cfg: any }> {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=pinterest_config&limit=1`))[0];
  const cfg = c?.pinterest_config || {};
  if (!cfg.refresh_token) throw new Error("Cliente não conectou o Pinterest Ads.");
  const acc = await sbGet("account_config", "id=eq.main&select=data");
  const pt = (acc[0]?.data || {}).pinterest_ads || {};
  if (!pt.client_id || !pt.client_secret) throw new Error("Faltam as credenciais do App do Pinterest em Configurações.");
  const basic = btoa(`${pt.client_id}:${pt.client_secret}`);
  const r = await fetch("https://api.pinterest.com/v5/oauth/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", Authorization: `Basic ${basic}` }, body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: cfg.refresh_token }) });
  const j = await r.json();
  if (!r.ok || !j.access_token) throw new Error("Pinterest: falha ao renovar token (" + (j.message || r.status) + ")");
  return { token: j.access_token, cfg };
}
async function pinterestListAccounts(clientId: string) {
  try {
    const { token } = await pinterestAccessToken(clientId);
    const r = await fetch("https://api.pinterest.com/v5/ad_accounts", { headers: { Authorization: `Bearer ${token}` } });
    const j = await r.json();
    if (!r.ok) return { error: "Pinterest: " + (j.message || `HTTP ${r.status}`) };
    const list = (j.items || []).map((a: any) => ({ id: a.id, name: a.name || a.id }));
    return { accounts: list };
  } catch (e) { return { error: String((e as any)?.message || e) }; }
}
async function pinterestAdsInsights(m: any) {
  const clientId = m.clientId;
  if (!clientId) throw new Error("clientId obrigatório");
  const { token, cfg } = await pinterestAccessToken(clientId);
  const adAccountId = m.adAccountId || cfg.ad_account_id;
  if (!adAccountId) throw new Error("Nenhuma conta de anúncio Pinterest selecionada pra esse cliente.");
  const since = m.since, until = m.until;
  if (!since || !until) throw new Error("since e until obrigatórios (YYYY-MM-DD)");
  const cols = "SPEND_IN_DOLLAR,IMPRESSION_1,CLICKTHROUGH_1,TOTAL_CONVERSIONS,CTR,CPC_IN_DOLLAR,ECPM_IN_DOLLAR";
  const p = new URLSearchParams({ start_date: since, end_date: until, columns: cols, granularity: "TOTAL", level: "CAMPAIGN" });
  const r = await fetch(`https://api.pinterest.com/v5/ad_accounts/${encodeURIComponent(adAccountId)}/campaigns/analytics?${p}`, { headers: { Authorization: `Bearer ${token}` } });
  const j = await r.json();
  if (!r.ok) throw new Error("Pinterest Ads: " + (j.message || `HTTP ${r.status}`));
  const rows = Array.isArray(j) ? j : [];
  const totAgg = { spend: 0, impressions: 0, clicks: 0, purchases: 0 };
  const campaigns = rows.map((row: any) => {
    const spend = Number(row.SPEND_IN_DOLLAR || 0), impressions = Number(row.IMPRESSION_1 || 0), clicks = Number(row.CLICKTHROUGH_1 || 0), conv = Number(row.TOTAL_CONVERSIONS || 0);
    totAgg.spend += spend; totAgg.impressions += impressions; totAgg.clicks += clicks; totAgg.purchases += conv;
    return { campaign: row.CAMPAIGN_ID || row.campaign_id || "Pinterest Ads", campaignId: row.CAMPAIGN_ID || row.campaign_id || null, spend, impressions, clicks, purchases: conv, ctr: Number(row.CTR || 0), cpc: Number(row.CPC_IN_DOLLAR || 0), cpm: Number(row.ECPM_IN_DOLLAR || 0), _pinterest: true };
  });
  const total = { ...totAgg, ctr: totAgg.impressions ? (totAgg.clicks / totAgg.impressions) * 100 : 0, cpc: totAgg.clicks ? totAgg.spend / totAgg.clicks : 0, cpm: totAgg.impressions ? (totAgg.spend / totAgg.impressions) * 1000 : 0 };
  return { total, campaigns, _pinterest: true, period: { since, until } };
}

// ===== Banco de dados de midia (Fase 2, schema `midia`) — grava nivel-ANUNCIO x dia, em paralelo ao channel_metrics_daily.
// Reaproveita a MESMA chamada metaAdsInsights/googleAdsInsights (com byAd+daily) feita pelo channelMetricsCollect logo
// abaixo — nao dobra chamada de API. Ver docs/spec-banco-dados-midia.md e migrations/banco_dados_midia.sql.
// Falha aqui NUNCA derruba a gravacao do channel_metrics_daily (Power BI atual) — so acumula em errors.
async function _midiaWriteAdsDaily(clientId: string, platform: "meta" | "google", contaExternaId: string, contaNome: string, adsDaily: any[]): Promise<number> {
  if (!adsDaily || !adsDaily.length) return 0;
  const origemDeteccao = platform === "meta" ? "meta_actions" : "google_conversion_action";
  // Google nao fornece reach/frequencia na API (ver Etapa 4 da spec) — fica NULL sempre, nunca 0. Meta fornece de verdade,
  // entao usa ?? (so vira NULL se o campo realmente nao vier), preservando um alcance=0 legitimo do Meta quando acontecer.
  const semReachFreq = platform === "google";

  const contaRet = await _midiaUpsert("dim_conta", [{ client_id: clientId, plataforma_id: platform, conta_externa_id: contaExternaId, nome: contaNome || contaExternaId }], "client_id,plataforma_id,conta_externa_id");
  const contaId = contaRet[0]?.id;
  if (!contaId) return 0;

  const campsByExt: Record<string, any> = {};
  for (const a of adsDaily) {
    if (!a.campaignId) continue;
    const k = String(a.campaignId);
    if (!campsByExt[k]) campsByExt[k] = {
      conta_id: contaId, campanha_externa_id: k, nome: a.campaign || "",
      objetivo_bruto: (a.objetivo && (a.objetivo.codigo || a.objetivo.tipo)) || null,
      objetivo_tipo: (a.objetivo && a.objetivo.tipo) || null, objetivo_rotulo: (a.objetivo && a.objetivo.rotulo) || null,
    };
  }
  const campRet = await _midiaUpsert("dim_campanha", Object.values(campsByExt), "conta_id,campanha_externa_id");
  const campIdByExt: Record<string, string> = {};
  for (const c of campRet) campIdByExt[c.campanha_externa_id] = c.id;

  const gruposByExt: Record<string, any> = {};
  for (const a of adsDaily) {
    if (!a.adsetId || !a.campaignId) continue;
    const campanhaId = campIdByExt[String(a.campaignId)];
    if (!campanhaId) continue;
    const key = campanhaId + "|" + a.adsetId;
    if (!gruposByExt[key]) gruposByExt[key] = { campanha_id: campanhaId, grupo_externo_id: String(a.adsetId), nome: a.adset || "" };
  }
  const grupoRet = await _midiaUpsert("dim_grupo", Object.values(gruposByExt), "campanha_id,grupo_externo_id");
  const grupoIdByKey: Record<string, string> = {};
  for (const g of grupoRet) grupoIdByKey[g.campanha_id + "|" + g.grupo_externo_id] = g.id;

  const anunciosByExt: Record<string, any> = {};
  const anuncioIdKeyOf = (a: any) => {
    if (!a.adId || !a.adsetId || !a.campaignId) return null;
    const campanhaId = campIdByExt[String(a.campaignId)];
    if (!campanhaId) return null;
    const grupoId = grupoIdByKey[campanhaId + "|" + a.adsetId];
    if (!grupoId) return null;
    return { grupoId, key: grupoId + "|" + a.adId };
  };
  for (const a of adsDaily) {
    const r = anuncioIdKeyOf(a);
    if (!r) continue;
    if (!anunciosByExt[r.key]) anunciosByExt[r.key] = { grupo_id: r.grupoId, anuncio_externo_id: String(a.adId), nome: a.adName || "" };
  }
  const anuncioRet = await _midiaUpsert("dim_anuncio", Object.values(anunciosByExt), "grupo_id,anuncio_externo_id");
  const anuncioIdByKey: Record<string, string> = {};
  for (const ad of anuncioRet) anuncioIdByKey[ad.grupo_id + "|" + ad.anuncio_externo_id] = ad.id;

  const perfRows: any[] = [], resRows: any[] = [], videoRows: any[] = [], engRows: any[] = [];
  for (const a of adsDaily) {
    const r = anuncioIdKeyOf(a);
    if (!r) continue;
    const anuncioId = anuncioIdByKey[r.key];
    if (!anuncioId) continue;
    for (const rec of (a.records || [])) {
      if (!rec.date) continue;
      perfRows.push({
        anuncio_id: anuncioId, data: rec.date, moeda: "BRL",
        investimento: rec.spend || 0,
        impressoes: rec.impressions ?? null,
        cliques: rec.clicks ?? null,
        alcance: semReachFreq ? null : (rec.reach ?? null),
        frequencia: semReachFreq ? null : (rec.frequency ?? null),
        atualizado_em: new Date().toISOString(),
      });
      const pushResultado = (tipo: string, qtd: number) => { if (qtd > 0) resRows.push({ anuncio_id: anuncioId, data: rec.date, tipo_resultado: tipo, quantidade: qtd, valor: tipo === "compra" ? (rec.revenue ?? null) : null, origem_deteccao: origemDeteccao, moeda: "BRL" }); };
      pushResultado("compra", rec.sales || 0);
      pushResultado("lead", rec.leads || 0);
      pushResultado("conversa", rec.conversas || 0);
      if ((rec.videoViews || 0) > 0) videoRows.push({ anuncio_id: anuncioId, data: rec.date, visualizacoes: rec.videoViews });
      if ((rec.engajamentos || 0) > 0) engRows.push({ anuncio_id: anuncioId, data: rec.date, engajamentos_total: rec.engajamentos });
    }
  }
  await Promise.all([
    _midiaUpsert("fact_performance", perfRows, "anuncio_id,data"),
    _midiaUpsert("fact_resultado", resRows, "anuncio_id,data,tipo_resultado"),
    _midiaUpsert("fact_video", videoRows, "anuncio_id,data"),
    _midiaUpsert("fact_engajamento", engRows, "anuncio_id,data"),
  ]);
  return perfRows.length;
}

// ===== Banco de dados de midia (Fase 3, schema `midia`) — snapshot DIARIO de seguidores do Instagram.
// Gap real: instagramListAccounts() so busca ao vivo, nunca guardou historico (por isso "seguidores ganhos
// no periodo" nao existia). 1 chamada de API pra TODOS os perfis da agencia de uma vez (me/accounts ja
// retorna todo mundo) - nao faz 1 chamada por cliente. Ver docs/spec-banco-dados-midia.md.
async function instagramFollowersSnapshot(input: any = {}) {
  const list = await instagramListAccounts();
  if (!(list as any).ok) throw new Error((list as any).erro || "Falha ao listar contas do Instagram");
  const seguidoresPorId: Record<string, number> = {};
  for (const p of ((list as any).paginas || [])) if (p.instagram && p.instagram.id) seguidoresPorId[p.instagram.id] = Number(p.instagram.seguidores || 0);

  const clientesAll = await _sbAll("clients", "status=neq.Encerrado&select=id,name,instagram_accounts");
  const clientes = Array.isArray(input.clientIds) && input.clientIds.length ? clientesAll.filter((c: any) => input.clientIds.includes(c.id)) : clientesAll;
  const hoje = new Date().toISOString().slice(0, 10);
  let gravados = 0; const semDado: string[] = [];
  for (const c of clientes) {
    const contas: any[] = Array.isArray(c.instagram_accounts) ? c.instagram_accounts : [];
    for (const ig of contas) {
      if (!ig.id) continue;
      const seguidores = seguidoresPorId[ig.id];
      if (seguidores == null) { semDado.push(`${c.name} (${ig.username || ig.id})`); continue; }
      try {
        const contaRet = await _midiaUpsert("dim_conta", [{ client_id: c.id, plataforma_id: "instagram_organico", conta_externa_id: String(ig.id), nome: ig.username || c.name }], "client_id,plataforma_id,conta_externa_id");
        const contaId = contaRet[0]?.id;
        if (!contaId) continue;
        await _midiaUpsert("fact_seguidores_snapshot", [{ conta_id: contaId, data: hoje, total_seguidores: seguidores }], "conta_id,data");
        gravados++;
      } catch (e) { semDado.push(`${c.name} (${ig.username || ig.id}): ${String((e as any)?.message || e)}`); }
    }
  }
  return { gravados, semDado, totalPerfisAgencia: Object.keys(seguidoresPorId).length };
}

// ===== Banco de dados de midia (Fase 4, schema `midia`) — historiza conteudo organico do Instagram.
// Reaproveita instagramOrganicContent() (mesma chamada de API que a Curadoria de Conteudo ja usa) - so
// grava tambem em dim_conteudo_organico + fact_conteudo_organico_metricas (1 snapshot por dia de coleta;
// nao reescreve o post, so acumula historico de metricas ao longo do tempo). Ver docs/spec-banco-dados-midia.md.
async function instagramOrganicSnapshot(m: any) {
  const dias = Number(m && m.days) || 90;
  const clientesAll = await _sbAll("clients", "status=neq.Encerrado&select=id,name,instagram_accounts");
  const alvos = (Array.isArray(m?.clientIds) && m.clientIds.length)
    ? clientesAll.filter((c: any) => m.clientIds.includes(c.id))
    : clientesAll.filter((c: any) => Array.isArray(c.instagram_accounts) && c.instagram_accounts.length);
  const hoje = new Date().toISOString().slice(0, 10);
  let gravados = 0; const errors: any[] = [];
  for (const c of alvos) {
    try {
      const r = await instagramOrganicContent({ clientId: c.id, days: dias });
      const posts = ((r as any).posts || []) as any[];
      if (!posts.length) continue;

      const contasByExt: Record<string, any> = {};
      for (const ig of (c.instagram_accounts || [])) contasByExt[String(ig.id)] = ig;
      const contaIdsUsados = [...new Set(posts.map((p: any) => p.contaId).filter(Boolean))] as string[];
      const contaIdMap: Record<string, string> = {};
      for (const igId of contaIdsUsados) {
        const ig = contasByExt[String(igId)];
        const ret = await _midiaUpsert("dim_conta", [{ client_id: c.id, plataforma_id: "instagram_organico", conta_externa_id: String(igId), nome: (ig && ig.username) || String(igId) }], "client_id,plataforma_id,conta_externa_id");
        if (ret[0]?.id) contaIdMap[String(igId)] = ret[0].id;
      }

      const conteudoByExt: Record<string, any> = {};
      for (const p of posts) {
        if (!p.contaId || !contaIdMap[String(p.contaId)]) continue;
        conteudoByExt[p.id] = { conta_id: contaIdMap[String(p.contaId)], post_externo_id: String(p.id), tipo_midia: p.tipo || null, permalink: p.permalink || null, legenda: p.caption || null, publicado_em: p.data || null, thumbnail_url: p.midia || null };
      }
      const conteudoRet = Object.values(conteudoByExt).length ? await _midiaUpsert("dim_conteudo_organico", Object.values(conteudoByExt), "conta_id,post_externo_id") : [];
      const conteudoIdByExt: Record<string, string> = {};
      for (const cc of conteudoRet) conteudoIdByExt[cc.post_externo_id] = cc.id;

      const metricRows = posts.filter((p: any) => conteudoIdByExt[p.id]).map((p: any) => ({
        conteudo_id: conteudoIdByExt[p.id], data_coleta: hoje,
        curtidas: p.likes ?? null, comentarios: p.comments ?? null, compartilhamentos: p.shares ?? null, salvamentos: p.saved ?? null,
        alcance: p.reach ?? null, visualizacoes: p.views ?? null,
      }));
      if (metricRows.length) await _midiaUpsert("fact_conteudo_organico_metricas", metricRows, "conteudo_id,data_coleta");
      gravados += metricRows.length;
    } catch (e) { errors.push({ client: c.name, error: String((e as any)?.message || e) }); }
  }
  return { gravados, clientesProcessados: alvos.length, errors };
}

// ===== Banco de Dados (histórico diário por canal) — alimenta a aba "Banco de Dados", export CSV e a conexão do Power BI =====
// Roda pra todos os clientes ativos com pelo menos um canal conectado (ou só os clientIds informados, pra permitir
// reconstrução do histórico em lotes menores sem estourar o tempo/CPU da function). Usa daily:true nas mesmas funções
// de insights já existentes (metaAdsInsights/googleAdsInsights) — não duplica lógica de fetch, só grava o resultado.
async function channelMetricsCollect(m: any) {
  const since = m.since, until = m.until;
  if (!since || !until) throw new Error("since e until obrigatórios (YYYY-MM-DD)");
  const ON_CONFLICT = "client_id,channel,date,source_medium,campaign,adset,ad_content";
  const clientsAll = await _sbAll("clients", "select=id,name,status,meta_account_id,google_account_id,ga4_property_id");
  const targets = (Array.isArray(m.clientIds) && m.clientIds.length)
    ? clientsAll.filter((c: any) => m.clientIds.includes(c.id))
    : clientsAll.filter((c: any) => c.status !== "Encerrado" && (c.meta_account_id || c.google_account_id || c.ga4_property_id));
  let saved = 0; const errors: any[] = [];
  // Meta/Google: 1 linha por dia × CONJUNTO/GRUPO DE ANÚNCIOS (não por ad — evita explosão de volume). GA4 não tem essa
  // função porque já vem quebrado por dia×origem/mídia×campanha×conteúdo direto do runReport (ver ga4DailyBySource).
  const toRows = (clientId: string, channel: string, campaign: string, adset: string, records: any[]) => records.filter((r: any) => r.date).map((rec: any) => {
    const impressions = rec.impressions || 0, clicks = rec.clicks || 0, spend = rec.spend || 0;
    return {
      id: `${clientId}_${channel}_${rec.date}_${campaign}_${adset}`.slice(0, 300), client_id: clientId, channel, date: rec.date, source_medium: "", campaign, adset, ad_content: "",
      spend, impressions, clicks, ctr: impressions ? (clicks / impressions) * 100 : 0, cpm: impressions ? (spend / impressions) * 1000 : 0, reach: rec.reach || 0,
      purchases: rec.sales || 0, revenue: rec.revenue || 0, leads: rec.leads || 0, conversas: rec.conversas || 0,
      video_views: rec.videoViews || 0, engajamentos: rec.engajamentos || 0, updated_at: new Date().toISOString(),
    };
  });
  for (const c of targets) {
    const mIds = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
    const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
    if (mIds.length) {
      try {
        const r = await metaAdsInsights({ accounts: mIds.map((id: string) => ({ id, name: id })), since, until, daily: true, byCampaign: true, byAdset: true, byAd: true });
        const rows = ((r as any).adsets || []).flatMap((as: any) => toRows(c.id, "meta", as.campaign || "Meta Ads", as.adset || "", as.records || []));
        if (rows.length) { await _sbUpsert("channel_metrics_daily", rows, ON_CONFLICT); saved += rows.length; }
        try {
          const byAcct: Record<string, any[]> = {};
          for (const a of ((r as any).adsDaily || [])) (byAcct[a.accountId] || (byAcct[a.accountId] = [])).push(a);
          for (const [acctId, list] of Object.entries(byAcct)) await _midiaWriteAdsDaily(c.id, "meta", acctId, c.name, list);
        } catch (e2) { errors.push({ client: c.name, channel: "meta-midia", error: String((e2 as any)?.message || e2) }); }
      } catch (e) { errors.push({ client: c.name, channel: "meta", error: String((e as any)?.message || e) }); }
    }
    if (gIds.length) {
      try {
        const r = await googleAdsInsights({ accounts: gIds.map((id: string) => ({ id, name: id })), since, until, daily: true, byCampaign: true, byAdset: true, byAd: true });
        const rows = ((r as any).adsets || []).flatMap((as: any) => toRows(c.id, "google", as.campaign || "Google Ads", as.adset || "", as.records || []));
        if (rows.length) { await _sbUpsert("channel_metrics_daily", rows, ON_CONFLICT); saved += rows.length; }
        try {
          const byAcct: Record<string, any[]> = {};
          for (const a of ((r as any).adsDaily || [])) (byAcct[a.accountId] || (byAcct[a.accountId] = [])).push(a);
          for (const [acctId, list] of Object.entries(byAcct)) await _midiaWriteAdsDaily(c.id, "google", acctId, c.name, list);
        } catch (e2) { errors.push({ client: c.name, channel: "google-midia", error: String((e2 as any)?.message || e2) }); }
      } catch (e) { errors.push({ client: c.name, channel: "google", error: String((e as any)?.message || e) }); }
    }
    if (c.ga4_property_id) {
      try {
        const recs = await ga4DailyBySource({ propertyId: c.ga4_property_id, since, until });
        const rows = recs.map((rec: any) => ({
          id: `${c.id}_ga4_${rec.date}_${rec.sourceMedium}_${rec.campaign}_${rec.adContent}`.slice(0, 300), client_id: c.id, channel: "ga4", date: rec.date,
          source_medium: rec.sourceMedium, campaign: rec.campaign || "", adset: "", ad_content: rec.adContent || "",
          spend: 0, impressions: 0, clicks: 0, reach: 0, purchases: rec.purchases, revenue: rec.revenue,
          leads: 0, conversas: 0, video_views: 0, engajamentos: 0, updated_at: new Date().toISOString(),
        }));
        if (rows.length) { await _sbUpsert("channel_metrics_daily", rows, ON_CONFLICT); saved += rows.length; }
        // banco de dados de midia (Fase 5): mesma chamada ga4DailyBySource, so grava tambem em midia.fact_analytics_ga4.
        // sessionSourceMedium vem combinado ("google / cpc") - separa em origem/midia_texto pra poder filtrar cada um.
        try {
          const ga4Rows = recs.map((rec: any) => {
            const [origem, midiaTexto] = String(rec.sourceMedium || "").split(" / ");
            return {
              client_id: c.id, propriedade_id: String(c.ga4_property_id), data: rec.date,
              origem: origem || "", midia_texto: midiaTexto || "", campanha_texto: rec.campaign || "", conteudo_texto: rec.adContent || "",
              compras: rec.purchases || 0, receita: rec.revenue || 0, moeda: "BRL",
            };
          });
          if (ga4Rows.length) await _midiaUpsert("fact_analytics_ga4", ga4Rows, "client_id,propriedade_id,data,origem,midia_texto,campanha_texto,conteudo_texto");
        } catch (e2) { errors.push({ client: c.name, channel: "ga4-midia", error: String((e2 as any)?.message || e2) }); }
      } catch (e) { errors.push({ client: c.name, channel: "ga4", error: String((e as any)?.message || e) }); }
    }
  }
  return { saved, clientesProcessados: targets.length, errors };
}

/* ===== AGENTE DE BRIEFING CRIATIVO (Etapa 1: análise dos criativos por funil) =====
   Le criativos por anuncio (Meta/Google, ja existentes em metaAdsInsights/googleAdsInsights), resolve
   o funil de cada um e chama a IA pra devolver a leitura por funil (melhores/piores, pontos positivos
   e negativos). Etapas 2-4 (curadoria, geracao de fichas, chat de ajustes) ainda nao implementadas. */
const REGRAS_DE_LINGUAGEM = "Regras de linguagem obrigatorias: nunca use os termos qualificado, lead quente, intencao de compra ou engajado. " +
  "Nunca afirme causalidade a partir de correlacao. Nao use adjetivo sem numero que o sustente. " +
  "Escreva em portugues do Brasil, tom tecnico e direto, sem linguagem de venda.";

// Ordem de resolucao 1/2 do funil (campo proprio nao existe no sistema; aqui e so o parsing por nomenclatura).
// Achado real analisando nomes de campanha de varios clientes: campanhas [FUNIL] usam FRIO/QUENTE/REMARKETING
// como 3 campanhas separadas dentro do mesmo conjunto de conversao — mapeia bem pra Topo/Meio/Fundo.
function _briefingResolveFunil(campanha: string, adset: string): string | null {
  const s = `${campanha || ""} ${adset || ""}`.toUpperCase();
  if (/MISTO|FRIO\s*\/\s*QUENTE|QUENTE\s*\/\s*FRIO/.test(s)) return null; // combinado de proposito, ambiguo
  if (/\bBOFU\b|\bFUNDO\b|REMARKETING|RETARGETING|CARRINHO/.test(s)) return "Fundo";
  if (/\bTOFU\b|\bTOPO\b|\bFRIO\b|ALCANCE|RECONHECIMENTO|REC\.?\s*MARCA|PROSPEC/.test(s)) return "Topo";
  if (/\bMOFU\b|\bMEIO\b|\bQUENTE\b|ENGAJAMENTO|TR[AÁ]FEGO|TRAFEGO/.test(s)) return "Meio";
  return null;
}
// Funil V2: o OBJETIVO da campanha e a ancora principal (a gente sabe com certeza a etapa dele);
// o nome so confirma ou desempata quando o objetivo nao resolve. Pedido da gestora: nomenclatura
// nem sempre identifica a etapa, entao o objetivo manda e o nome cruza.
const _FUNIL_POR_OBJETIVO: Record<string, string> = {
  alcance: "Topo", video: "Topo", perfil: "Topo",
  trafego: "Meio", engajamento: "Meio",
  leads: "Fundo", mensagens: "Fundo", conversao: "Fundo", app: "Fundo",
};
function _briefingResolveFunilV2(objetivo: any, campanha: string, conjunto: string): { funil: string | null; origem: string } {
  const porNome = _briefingResolveFunil(campanha, conjunto);
  const porObjetivo = _FUNIL_POR_OBJETIVO[String(objetivo?.tipo || "").toLowerCase()] || null;
  if (porObjetivo && porNome === porObjetivo) return { funil: porObjetivo, origem: "objetivo+nome" };
  if (porObjetivo && porNome) return { funil: porObjetivo, origem: "objetivo (nome diverge)" };
  if (porObjetivo) return { funil: porObjetivo, origem: "objetivo" };
  if (porNome) return { funil: porNome, origem: "nome" };
  return { funil: null, origem: "indefinido" };
}
// Metrica de "resultado" certa pro objetivo do anuncio (mesma logica de classifyAdByObjective/raioXMetricRows do front, replicada aqui pro server).
function _briefingResultado(a: any): { label: string; valor: number; custo: number | null } {
  const tipo = (a.objetivo && a.objetivo.tipo) || "conversao";
  const spend = a.spend || 0;
  if (tipo === "leads") return { label: "leads", valor: a.leads || 0, custo: a.leads ? spend / a.leads : null };
  if (tipo === "trafego") return { label: "cliques", valor: a.clicks || 0, custo: a.clicks ? spend / a.clicks : null };
  if (tipo === "engajamento") return { label: "engajamentos", valor: a.engajamentos || 0, custo: a.engajamentos ? spend / a.engajamentos : null };
  if (tipo === "mensagens") return { label: "conversas", valor: a.conversas || 0, custo: a.conversas ? spend / a.conversas : null };
  if (tipo === "video") return { label: "views (ThruPlay)", valor: a.videoViews || 0, custo: a.videoViews ? spend / a.videoViews : null };
  if (tipo === "perfil") return { label: "seguidores", valor: a.seguidores || 0, custo: a.seguidores ? spend / a.seguidores : null };
  if (tipo === "alcance") return { label: "alcance", valor: a.reach || 0, custo: a.reach ? spend / a.reach : null };
  return { label: "compras", valor: a.purchases || 0, custo: a.purchases ? spend / a.purchases : null }; // conversao/app/default
}
// Piso de elegibilidade por funil (sugerido, ajustavel — ver spec). Abaixo do piso o anuncio entra como "em leitura".
function _briefingElegivel(funil: string | null, a: any): boolean {
  if (funil === "Topo") return (a.impressions || 0) >= 2000;
  if (funil === "Meio") return (a.impressions || 0) >= 1500;
  if (funil === "Fundo") return (a.spend || 0) >= 50;
  return (a.impressions || 0) >= 2000; // funil nao resolvido ainda: usa o piso mais permissivo, LLM decide o funil depois
}
async function _briefingMetaThumbs(adIds: string[]) {
  const out: Record<string, { thumb?: string; ig?: string }> = {}, token = await _metaUserToken();
  if (!token) return out;
  // Em lotes de 50 E EM PARALELO (6 por vez): num cliente com 1.600 anuncios no periodo isso e a
  // diferenca entre 33 chamadas em fila e 6 rodadas. Enfileirado, so as miniaturas ja comiam meio
  // minuto do orcamento da funcao.
  const lotes: string[][] = [];
  for (let i = 0; i < adIds.length; i += 50) { const ids = adIds.slice(i, i + 50).filter(Boolean); if (ids.length) lotes.push(ids); }
  for (let i = 0; i < lotes.length; i += 6) {
    await Promise.all(lotes.slice(i, i + 6).map(async (ids) => {
    try {
      // image_url primeiro (resolucao cheia); thumbnail em 512px pro fallback (video/carrossel so tem thumbnail).
      // O modificador de tamanho vai GRUDADO no campo expandido — creative.thumbnail_width(512){...} — como query
      // param solto ele e ignorado e o Graph devolve o padrao de 64px (borrao no zoom).
      // instagram_permalink_url = link do proprio post/anuncio no Instagram (quando existe).
      const r = await fetch(`https://graph.facebook.com/v21.0/?ids=${ids.join(",")}&fields=creative.thumbnail_width(512).thumbnail_height(512){thumbnail_url,image_url,instagram_permalink_url}&access_token=${token}`), j = await r.json();
      for (const id of ids) { const cr = j[id]?.creative; if (!cr) continue; const thumb = cr.image_url || cr.thumbnail_url; out[id] = { thumb: thumb || undefined, ig: cr.instagram_permalink_url || undefined }; }
    } catch (_e) { /* card continua com link e KPIs */ }
    }));
  }
  return out;
}
async function _briefingCriativos(clientId: string, since: string, until: string) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=id,name,meta_account_id,google_account_id`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const mIds = String(c.meta_account_id || "").split(",").map((s) => s.trim()).filter(Boolean);
  const gIds = String(c.google_account_id || "").split(",").map((s) => s.trim()).filter(Boolean);
  if (!mIds.length && !gIds.length) throw new Error("Cliente sem conta Meta ou Google conectada.");
  const [mRes, gRes] = await Promise.all([
    mIds.length ? metaAdsInsights({ accountIds: mIds, since, until, byAd: true }).catch((e: any) => ({ ads: [], error: String(e?.message || e) })) : Promise.resolve({ ads: [] as any[] }),
    gIds.length ? googleAdsInsights({ accountIds: gIds, since, until, byAd: true }).catch((e: any) => ({ ads: [], error: String(e?.message || e) })) : Promise.resolve({ ads: [] as any[] }),
  ]);
  const rawAds = [...((mRes as any).ads || []), ...((gRes as any).ads || [])];
  let inferidos = 0;
  const criativos = rawAds.map((a: any, i: number) => {
    const codigo = "AD" + String(i + 1).padStart(2, "0");
    const fr = _briefingResolveFunilV2(a.objetivo, a.campaign, a.adset);
    if (fr.origem === "nome" || fr.origem === "indefinido") inferidos++; // sem ancora de objetivo = menos confiavel
    const elegivel = _briefingElegivel(fr.funil, a);
    const r = _briefingResultado(a);
    return {
      codigo, adId: a.adId, nome: a.adName || a.campaign || codigo, canal: a._google ? "Google" : "Meta", campanha: a.campaign || "", conjunto: a.adset || "", objetivo: a.objetivo?.rotulo || a.objetivo?.tipo || "",
      funil: fr.funil, funilOrigem: fr.origem, elegivel, thumbnail: a.thumbnail || null,
      link: a._google
        ? `https://ads.google.com/aw/ads?campaignId=${encodeURIComponent(a.campaignId || "")}&adGroupId=${encodeURIComponent(a.adsetId || "")}`
        : `https://adsmanager.facebook.com/adsmanager/manage/ads?act=${encodeURIComponent(a.accountId || "")}&selected_ad_ids=${encodeURIComponent(a.adId || "")}`,
      spend: Math.round((a.spend || 0) * 100) / 100, impressions: a.impressions || 0, clicks: a.clicks || 0,
      reach: a.reach || 0, frequency: +(a.frequency || 0).toFixed(2), ctr: +(a.ctr || 0).toFixed(2), cpm: +(a.cpm || 0).toFixed(2),
      resultadoLabel: r.label, resultadoValor: Math.round(r.valor), custoPorResultado: r.custo != null ? Math.round(r.custo * 100) / 100 : null,
      compras: Math.round(a.purchases || 0), receita: Math.round((a.revenue || 0) * 100) / 100, leads: Math.round(a.leads || 0),
      seguidores: Math.round(a.seguidores || 0), videoViews: a.videoViews || 0,
    };
  });
  // Resolve todos os anúncios Meta relevantes: além da miniatura, traz o permalink do
  // post no Instagram quando o criativo foi publicado por lá. Sem permalink, o card
  // mantém o link seguro para o anúncio no Gerenciador.
  const metaToResolve = criativos.filter((x: any) => x.canal === "Meta" && (x.elegivel || x.spend > 0) && x.adId).map((x: any) => x.adId);
  if (metaToResolve.length) { const extra = await _briefingMetaThumbs(metaToResolve); for (const x of criativos) { const e = extra[x.adId]; if (!e) continue; if (!x.thumbnail && e.thumb) x.thumbnail = e.thumb; if (e.ig) (x as any).igUrl = e.ig; } }
  const total = criativos.length;
  const pctInferido = total ? Math.round((inferidos / total) * 1000) / 10 : 0;
  return { criativos, total, pctInferido, erros: [(mRes as any).error, (gRes as any).error].filter(Boolean) };
}
/* Busca os criativos guardando o resultado por 30 minutos.
   Buscar 79 dias de 7 contas Meta passa de 80 segundos. Fazer isso E a leitura da IA na mesma chamada
   estourava o limite da Edge Function (a tela mostrava so "non-2xx"). Agora a tela chama primeiro
   briefingPreparar (que enche este cache) e depois a analise, que le daqui em menos de um segundo. */
const BRIEFING_CACHE_MIN = 30;
async function _briefingCriativosCache(clientId: string, since: string, until: string, forcar = false) {
  if (!forcar) {
    const hit = (await sbGet("briefing_criativos_cache", `client_id=eq.${encodeURIComponent(clientId)}&since=eq.${since}&until=eq.${until}&select=payload,gerado_em&limit=1`))[0];
    if (hit && (Date.now() - new Date(hit.gerado_em).getTime()) < BRIEFING_CACHE_MIN * 60000) {
      return { ...hit.payload, doCache: true, geradoEm: hit.gerado_em };
    }
  }
  const r = await _briefingCriativos(clientId, since, until);
  await _sbUpsert("briefing_criativos_cache", [{ client_id: clientId, since, until, payload: r, gerado_em: new Date().toISOString() }], "client_id,since,until");
  return { ...r, doCache: false };
}
/* Prepara (e guarda) os criativos do periodo. Devolve so a contagem — a lista inteira de um cliente
   grande passa de meio mega e a tela nao precisa dela nesta etapa. */
async function briefingPreparar(input: any) {
  const { clientId, since, until, forcar } = input;
  if (!clientId || !since || !until) throw new Error("clientId, since e until são obrigatórios.");
  const t0 = Date.now();
  const { criativos, total, pctInferido, erros, doCache } = await _briefingCriativosCache(clientId, since, until, !!forcar);
  const porFunil: Record<string, number> = {};
  for (const c of criativos) if (c.elegivel) porFunil[c.funil || "não identificado"] = (porFunil[c.funil || "não identificado"] || 0) + 1;
  return { total, elegiveis: criativos.filter((c: any) => c.elegivel).length, porFunil, pctInferido, erros, doCache, segundos: Math.round((Date.now() - t0) / 1000) };
}
/* Quem vai pra dentro do prompt. A IA escolhe no maximo 2 melhores e 2 piores por funil — mandar 600
   criativos pra isso e caro, lento e nao melhora a leitura. Levamos os extremos de cada funil (os mais
   baratos e os mais caros por resultado), que e exatamente o que ela precisa comparar. O que ficou de
   fora e DEVOLVIDO na resposta e aparece na tela: corte silencioso viraria "analisei tudo" mentiroso. */
function _briefingAmostraParaIA(elegiveis: any[], porFunil = 40) {
  const grupos: Record<string, any[]> = {};
  for (const c of elegiveis) (grupos[c.funil || "não identificado"] ||= []).push(c);
  const amostra: any[] = []; let fora = 0;
  const custo = (x: any) => (x.custoPorResultado == null ? Number.MAX_SAFE_INTEGER : x.custoPorResultado);
  for (const f of Object.keys(grupos)) {
    const arr = grupos[f];
    if (arr.length <= porFunil) { amostra.push(...arr); continue; }
    const ord = [...arr].sort((a, b) => custo(a) - custo(b));
    const metade = Math.floor(porFunil / 2);
    amostra.push(...ord.slice(0, metade), ...ord.slice(-(porFunil - metade)));
    fora += arr.length - porFunil;
  }
  return { amostra, fora };
}
function _briefingDadosTxt(criativos: any[]) {
  return criativos.filter((c) => c.elegivel).map((c) =>
    `${c.codigo} | funil: ${c.funil || "não identificado"} | canal: ${c.canal} | campanha: ${c.campanha} | conjunto: ${c.conjunto} | ` +
    `invest: R$${c.spend} | impressões: ${c.impressions} | alcance: ${c.reach} | freq: ${c.frequency} | cliques: ${c.clicks} | ` +
    `CTR: ${c.ctr}% | CPM: R$${c.cpm} | ${c.resultadoLabel}: ${c.resultadoValor} | custo/${c.resultadoLabel}: ${c.custoPorResultado != null ? "R$" + c.custoPorResultado : "—"}`
  ).join("\n");
}
async function _callOpenAIJson(messages: any[]): Promise<any> {
  const json = await callOpenAI({ model: "gpt-4o", messages, max_tokens: 4000, temperature: 0.4 });
  const raw = String(json.choices?.[0]?.message?.content || "");
  const limpo = raw.replace(/```json/g, "").replace(/```/g, "").trim();
  try { return JSON.parse(limpo); } catch (_e) {
    // uma tentativa de reparo (JSON truncado/malformado) antes de falhar explicito
    const repair = await callOpenAI({ model: "gpt-4o", messages: [...messages, { role: "assistant", content: raw }, { role: "user", content: "Sua resposta anterior não é um JSON válido (provavelmente truncou). Responda de novo APENAS com o JSON completo e válido, sem markdown, mais curto se precisar." }], max_tokens: 4000, temperature: 0.2 });
    const raw2 = String(repair.choices?.[0]?.message?.content || "").replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(raw2); // se falhar de novo, sobe o erro — falha explicita, nunca renderiza parcial
  }
}
// Sugere Objetivo + Angulo pro briefing a partir do funil escolhido (e do DNA do cliente, quando tiver).
async function briefingSugerirCampos(input: any) {
  const { clientId, funil, produto } = input;
  if (!clientId) throw new Error("clientId obrigatório.");
  if (!funil) throw new Error("Escolha o funil desejado primeiro.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,seg,dna`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const dnaTxt = c.dna ? JSON.stringify(c.dna).slice(0, 4000) : "";
  const prompt = `Voce e estrategista de trafego pago. Sugira o OBJETIVO e o ANGULO de uma campanha/peca publicitaria pro cliente abaixo, pro funil indicado.

Cliente: ${c.name}
Segmento: ${c.seg || "não informado"}
Funil: ${funil}
${produto ? `Produto/programa especifico: ${produto}` : ""}
${dnaTxt ? `DNA do cliente (identidade, produtos, personas, diretrizes):\n${dnaTxt}` : "Sem DNA cadastrado - baseie-se so no nome/segmento, sem inventar produto ou oferta especifica."}

Regra por funil: Topo = atrair quem ainda nao conhece (dor ou desejo amplo, sem pedir decisao); Meio = considerar/comparar (prova, diferencial); Fundo = decisao/urgencia (oferta, condicao, prova social forte).

${REGRAS_DE_LINGUAGEM}

Responda APENAS com JSON valido, sem markdown: {"objetivo":"<1 frase curta, o que a campanha precisa entregar>","angulo":"<1 frase curta, a dor ou desejo a trabalhar>"}`;
  return await _callOpenAIJson([{ role: "user", content: prompt }]);
}
// Ranking de criativos SEM IA: devolve os anuncios do periodo classificados por funil (objetivo como ancora,
// nome como confirmacao) com KPIs e thumbnail — o front ordena e renderiza. A leitura de pontos fortes/fracos
// e um segundo passo SOB DEMANDA por criativo (briefingCreativoAnalise), pra nao poluir nem gastar IA a toa.
async function briefingRanking(input: any) {
  const { clientId, since, until, funilDesejado, produto } = input;
  if (!clientId || !since || !until) throw new Error("clientId, since e until são obrigatórios.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const { criativos, total, pctInferido, erros } = await _briefingCriativosCache(clientId, since, until);
  let lista = funilDesejado ? criativos.filter((x: any) => x.funil === funilDesejado) : criativos;
  if (produto) {
    const alvo = String(produto).toLowerCase();
    lista = lista.filter((x: any) => x.campanha.toLowerCase().includes(alvo) || x.conjunto.toLowerCase().includes(alvo));
  }
  return { criativos: lista, total, pctInferido, erros };
}
// Pontos fortes/fracos de UM criativo, sob demanda (botao no card). Compara com os pares do mesmo funil que o
// front ja tem em maos; quando ha thumbnail, manda a imagem junto pra leitura de design de verdade.
async function briefingCreativoAnalise(input: any) {
  const { clientId, criativo, pares } = input;
  if (!clientId || !criativo) throw new Error("clientId e criativo são obrigatórios.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,seg`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const linha = (x: any) => `${x.codigo} | ${x.nome} | campanha: ${x.campanha} | invest: R$${x.spend} | ${x.resultadoLabel}: ${x.resultadoValor} | custo/${x.resultadoLabel}: ${x.custoPorResultado != null ? "R$" + x.custoPorResultado : "—"} | CTR: ${x.ctr}% | CPM: R$${x.cpm}`;
  const prompt = `Voce e analista de criativos de trafego pago. Analise UM criativo especifico e devolva pontos fortes e fracos como instrucao util pra quem vai PRODUZIR a proxima peca (design/roteiro/edicao) — nao como descricao de metrica.

Cliente: ${c.name}${c.seg ? ` (${c.seg})` : ""}
Funil deste criativo: ${criativo.funil || "não identificado"} (classificado pelo ${criativo.funilOrigem || "?"})

CRIATIVO ANALISADO:
${linha(criativo)}

PARES DO MESMO FUNIL (contexto de comparação):
${(Array.isArray(pares) ? pares : []).slice(0, 8).map(linha).join("\n") || "(sem pares no período)"}

${criativo.thumbnail ? "A imagem do criativo esta anexada — analise TAMBEM o design (hierarquia visual, legibilidade, gancho, CTA visivel)." : "Sem imagem disponivel — analise apenas pelos numeros e pelos nomes, sem inventar atributo visual."}

${REGRAS_DE_LINGUAGEM}

Responda APENAS com JSON valido, sem markdown:
{"veredito":"<1 frase: por que este criativo esta acima ou abaixo dos pares>","positivos":["<max 3, frases curtas>"],"negativos":["<max 3, frases curtas>"]}`;
  let messages: any[] = [{ role: "user", content: prompt }];
  if (criativo.thumbnail) messages = [{ role: "user", content: [{ type: "text", text: prompt }, { type: "image_url", image_url: { url: criativo.thumbnail } }] }];
  try {
    return await _callOpenAIJson(messages);
  } catch (e) {
    if (!criativo.thumbnail) throw e;
    return await _callOpenAIJson([{ role: "user", content: prompt }]); // imagem inacessivel (URL do FB expira): repete so com texto
  }
}
async function briefingAnalise(input: any) {
  const { clientId, since, until, objetivo, criadoPor, funilDesejado, produto } = input;
  if (!clientId || !since || !until) throw new Error("clientId, since e until são obrigatórios.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name`))[0];
  if (!c) throw new Error("Cliente não encontrado.");
  const { criativos: todosCriativos, total, pctInferido, erros } = await _briefingCriativosCache(clientId, since, until);
  // se um funil especifico foi pedido, a analise (e o gasto com IA) fica so nele - nao processa os outros a toa
  let criativos = funilDesejado ? todosCriativos.filter((x) => x.funil === funilDesejado) : todosCriativos;
  // produto/segmento e texto livre (nao catalogo fixo) - filtra pelo nome de campanha/conjunto conter o termo.
  if (produto) {
    const alvo = String(produto).toLowerCase();
    criativos = criativos.filter((x) => x.campanha.toLowerCase().includes(alvo) || x.conjunto.toLowerCase().includes(alvo));
  }
  const elegiveis = criativos.filter((x) => x.elegivel);
  if (!elegiveis.length) return { erro: `Nenhum criativo elegível${funilDesejado ? ` do funil ${funilDesejado}` : ""}${produto ? ` com "${produto}" no nome da campanha/conjunto` : ""} no período. Amplie o período ou ajuste os filtros.`, erros };
  const thumbsByCodigo: Record<string, string> = {}, criativosByCodigo: Record<string, any> = {};
  for (const x of criativos) {
    if (x.thumbnail) thumbsByCodigo[x.codigo] = x.thumbnail;
    criativosByCodigo[x.codigo] = x;
  }
  const { amostra, fora } = _briefingAmostraParaIA(elegiveis);
  const dadosTxt = _briefingDadosTxt(amostra);
  const prompt = `Voce e analista de criativos de trafego pago. Analise os criativos que rodaram e devolva a leitura para o time de producao, que precisa entender o que funcionou antes de criar peca nova (direcao de DESIGN E VIDEO - a leitura de investimento/orcamento e feita em outro lugar, nao repita numero de gasto na resposta).

Cliente: ${c.name}
Periodo: ${since} a ${until}
Objetivo da campanha: ${objetivo || "não informado"}
${produto ? `Produto/segmento: SOMENTE "${produto}" - os dados abaixo ja vem filtrados por esse termo no nome da campanha/conjunto.` : ""}
${funilDesejado ? `Funil pedido: SOMENTE ${funilDesejado} - os dados abaixo ja vem filtrados so desse funil.` : ""}

Criativos que rodaram:
${dadosTxt}

${funilDesejado ? `Toda a analise e desse UNICO funil (${funilDesejado}).` : `Organize a analise POR FUNIL: topo, meio e fundo. Inclua apenas os funis presentes nos dados.`}
Se o funil nao estiver explicito (aparece como "não identificado"), infira pela campanha ou pelo nome do anuncio e acrescente (inferido) ao lado do codigo.

Compare cada criativo apenas dentro do proprio funil. Custo por resultado de topo e de fundo nao sao grandezas equivalentes.

Para cada criativo, traga pontos positivos e pontos negativos, mesmo nos melhores e nos piores: o melhor criativo tem algo a corrigir e o pior quase sempre tem algo aproveitavel. Escreva os pontos como instrucao util para quem vai PRODUZIR a peca (design/roteiro/edicao), nao como descricao de metrica de trafego.

${REGRAS_DE_LINGUAGEM}

Limites: no maximo 2 melhores e 2 piores por funil, no maximo 2 pontos positivos e 2 negativos por criativo, frases curtas de uma linha. No maximo 3 padroes no total. Padrao apoiado em menos de 3 pecas deve ter hipotese true. O campo "metricas" deve ser so uma pista curta e nao-financeira de qual sinal destacou essa peca (ex: "maior CTR do funil", "menor custo por resultado do grupo") - sem valores de investimento/gasto.

Responda APENAS com JSON valido, sem markdown, sem preambulo:
{"leitura":"<um paragrafo curto sobre o periodo inteiro>",
 "funis":[{"funil":"Topo","leitura":"<uma frase sobre este funil>","total_pecas":0,
   "melhores":[{"codigo":"","metricas":"","positivos":[""],"negativos":[""]}],
   "piores":[{"codigo":"","metricas":"","positivos":[""],"negativos":[""]}]}],
 "padroes":[{"afirmacao":"","evidencia":"","pecas":0,"hipotese":false}]}`;
  const parsed = await _callOpenAIJson([{ role: "user", content: prompt }]);

  const briefingId = _wuid();
  await sbPost("briefing", {
    id: briefingId, client_id: clientId, criado_por: criadoPor || null,
    periodo_inicio: since, periodo_fim: until, objetivo: objetivo || "", funil: funilDesejado || "", produto: produto || "",
    status: "pronto", criado_em: new Date().toISOString(), atualizado_em: new Date().toISOString(),
  });
  const analiseId = _wuid();
  await sbPost("briefing_analise", {
    id: analiseId, briefing_id: briefingId,
    leitura: parsed.leitura || "", funis_json: parsed.funis || [], padroes_json: parsed.padroes || [],
    criativos_analisados: elegiveis.length, criativos_em_leitura: total - elegiveis.length,
    funil_inferido_pct: pctInferido, criativos_json: criativosByCodigo, gerado_em: new Date().toISOString(),
  });
  return { briefingId, analiseId, ...parsed, thumbsByCodigo, criativosByCodigo, criativos_analisados: elegiveis.length, criativos_em_leitura: total - elegiveis.length, criativos_fora_da_leitura: fora, funil_inferido_pct: pctInferido, avisoConfiabilidade: pctInferido > 30, erros };
}
// Aprova o briefing (congela e "envia pra fila de producao"). So muda status - fichas ja foram geradas antes.
async function briefingAprovar(input: any) {
  const { briefingId } = input;
  if (!briefingId) throw new Error("briefingId obrigatório.");
  await sbPatchD("briefing", `id=eq.${encodeURIComponent(briefingId)}`, { status: "aprovado", atualizado_em: new Date().toISOString() });
  return { ok: true };
}
// Historico de briefings (aprovados por padrao) pra reconsulta - lista + o pacote completo de 1 briefing.
async function briefingHistorico(input: any) {
  const { clientId, status } = input;
  let q = `select=id,client_id,periodo_inicio,periodo_fim,objetivo,funil,produto,status,criado_em&order=criado_em.desc&limit=100`;
  if (clientId) q += `&client_id=eq.${encodeURIComponent(clientId)}`;
  q += `&status=eq.${encodeURIComponent(status || "aprovado")}`;
  const rows = await sbGet("briefing", q);
  const clientIds = [...new Set(rows.map((r: any) => r.client_id))];
  const clientes = clientIds.length ? await _sbAll("clients", `id=in.(${clientIds.map((id: any) => encodeURIComponent(id)).join(",")})&select=id,name`) : [];
  const nomeById: Record<string, string> = {}; for (const c of clientes) nomeById[c.id] = c.name;
  return { briefings: rows.map((r: any) => ({ ...r, clientName: nomeById[r.client_id] || r.client_id })) };
}
async function briefingCompleto(input: any) {
  const { briefingId } = input;
  if (!briefingId) throw new Error("briefingId obrigatório.");
  const [briefing, analise, curadoria, fichas] = await Promise.all([
    sbGet("briefing", `id=eq.${encodeURIComponent(briefingId)}&select=*`),
    sbGet("briefing_analise", `briefing_id=eq.${encodeURIComponent(briefingId)}&select=*&order=gerado_em.desc&limit=1`),
    sbGet("briefing_curadoria", `briefing_id=eq.${encodeURIComponent(briefingId)}&select=*&order=gerado_em.desc&limit=1`),
    sbGet("briefing_ficha", `briefing_id=eq.${encodeURIComponent(briefingId)}&select=*&order=ordem.asc`),
  ]);
  if (!briefing[0]) throw new Error("Briefing não encontrado.");
  return { briefing: briefing[0], analise: analise[0] || null, curadoria: curadoria[0] || null, fichas: fichas || [] };
}

async function briefingGerarFichas(input: any) {
  const briefingId = String(input.briefingId || "");
  if (!briefingId) throw new Error("briefingId obrigatório.");
  const b = (await sbGet("briefing", `id=eq.${encodeURIComponent(briefingId)}&select=*`))[0];
  if (!b) throw new Error("Briefing não encontrado.");
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(b.client_id)}&select=name,dna,seg`))[0] || {};
  const a = (await sbGet("briefing_analise", `briefing_id=eq.${encodeURIComponent(briefingId)}&select=*&order=gerado_em.desc&limit=1`))[0];
  if (!a) throw new Error("A análise de desempenho precisa ser gerada antes das fichas.");
  const curadoria = (await sbGet("briefing_curadoria", `briefing_id=eq.${encodeURIComponent(briefingId)}&select=leitura,candidatos_json&order=gerado_em.desc&limit=1`))[0] || null;

  const funis = (Array.isArray(input.funis) ? input.funis : []).filter((x: any) => ["Topo", "Meio", "Fundo"].includes(x));
  const canais = (Array.isArray(input.canais) ? input.canais : []).filter(Boolean);
  const formatos = (Array.isArray(input.formatos) ? input.formatos : []).filter(Boolean);
  const variacoes = Math.min(12, Math.max(1, Number(input.variacoes) || 4));
  if (!funis.length) throw new Error("Selecione pelo menos uma etapa do funil.");
  if (!canais.length) throw new Error("Selecione pelo menos um canal.");
  if (!formatos.length) throw new Error("Selecione pelo menos um formato.");

  const dna = c.dna || {};
  const contexto = {
    cliente: c.name || "", segmento: c.seg || "", objetivo: input.objetivo || b.objetivo || "",
    publico: input.publico || "", angulo: input.angulo || "", promessa: input.promessa || "",
    referencia: input.referencia || "", funis, canais, formatos, variacoes, prazo: input.prazo || null,
    marca: dna?.identidade || {}, produtos: (dna?.produtos || []).slice(0, 12), personas: (dna?.personas || []).slice(0, 8),
    objecoes: (dna?.objecoes || []).slice(0, 10), // o que trava a compra + a quebra: material direto pra copy
    diretrizes: dna?.diretrizes || {},
  };
  const prompt = `Você é diretor de criação para mídia paga. Transforme a análise de performance em fichas objetivas de solicitação de criativos para o time de produção.

CONTEXTO DO PEDIDO:
${JSON.stringify(contexto)}

ANÁLISE DE PERFORMANCE:
${JSON.stringify({ leitura: a.leitura, funis: a.funis_json, padroes: a.padroes_json })}

CURADORIA ORGÂNICA DISPONÍVEL (quando existir, pode gerar ficha com rota "recorte"):
${JSON.stringify(curadoria)}

Crie exatamente ${variacoes} fichas no total, distribuídas entre os funis selecionados. Cada ficha deve ter uma hipótese criativa diferente e aproveitar evidências da análise, sem inventar resultados. Adapte formato, linguagem, CTA e métrica esperada à etapa do funil e ao canal. Em Topo priorize atenção e consumo; em Meio, consideração e prova; em Fundo, ação e conversão. Se uma informação não foi fornecida, escreva uma orientação segura e marcável para validação, sem inventar oferta, preço ou garantia.

${REGRAS_DE_LINGUAGEM}

Responda APENAS com JSON válido:
{"fichas":[{"titulo":"","prioridade":"P1|P2|P3","rota":"nova","funil":"Topo|Meio|Fundo","canal":"","formato":"","objetivo":"","referencia":"","publico":"","angulo":"","promessa":"","roteiro":{"gancho":"","desenvolvimento":[""],"cta":""},"copy":{"texto_principal":"","titulo":"","descricao":""},"especificacoes":"","obrigatorio":[""],"proibido":[""]}]}`;
  const parsed = await _callOpenAIJson([{ role: "user", content: prompt }]);
  const fichas = (Array.isArray(parsed.fichas) ? parsed.fichas : []).slice(0, variacoes);
  if (!fichas.length) throw new Error("A IA não devolveu fichas válidas. Tente novamente.");

  const now = new Date().toISOString();
  await sbPatchD("briefing", `id=eq.${encodeURIComponent(briefingId)}`, {
    objetivo: contexto.objetivo, publico: contexto.publico, angulo: contexto.angulo, promessa: contexto.promessa,
    funil: funis.join(", "), formatos_json: formatos, canais_json: canais, variacoes,
    referencia: contexto.referencia, prazo: contexto.prazo, status: "pronto", atualizado_em: now,
  });
  const rows = fichas.map((f: any, i: number) => ({
    id: _wuid(), briefing_id: briefingId, codigo: `CR-${String(i + 1).padStart(2, "0")}`,
    titulo: f.titulo || `Criativo ${i + 1}`, prioridade: ["P1", "P2", "P3"].includes(f.prioridade) ? f.prioridade : "P2",
    rota: f.rota === "recorte" ? "recorte" : "nova", funil: f.funil || funis[i % funis.length],
    canal: f.canal || canais[i % canais.length], formato: f.formato || formatos[i % formatos.length],
    objetivo: f.objetivo || contexto.objetivo, referencia: f.referencia || contexto.referencia,
    publico: f.publico || contexto.publico, angulo: f.angulo || contexto.angulo, promessa: f.promessa || contexto.promessa,
    roteiro_json: f.roteiro || {}, copy_json: f.copy || {}, especificacoes: f.especificacoes || "",
    obrigatorio_json: Array.isArray(f.obrigatorio) ? f.obrigatorio : [], proibido_json: Array.isArray(f.proibido) ? f.proibido : [],
    prazo: contexto.prazo, ordem: i + 1, status: "fila",
  }));
  await sbPost("briefing_ficha", rows as any);
  return { briefingId, fichas: rows };
}

// Pausa/reativa palavra(s)-chave no Google (mutate status de ad_group_criterion). Pode vir 1 ou várias instâncias do mesmo texto.
async function googleKeywordAction(m: any) {
  const action = m.action === "enable" ? "enable" : "pause";
  const status = action === "pause" ? "PAUSED" : "ENABLED";
  const refs: any[] = Array.isArray(m.resources) && m.resources.length ? m.resources : (m.resourceName ? [{ accountId: m.accountId, resourceName: m.resourceName }] : []);
  if (!refs.length) throw new Error("nenhuma palavra-chave informada");
  const token = await googleAdsAccessToken();
  const devToken = Deno.env.get("GOOGLE_ADS_DEV_TOKEN"); const mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  const byAcc: Record<string, string[]> = {};
  refs.forEach((r) => { const cid = String(r.accountId || "").replace(/-/g, ""); if (cid && r.resourceName) (byAcc[cid] = byAcc[cid] || []).push(r.resourceName); });
  let n = 0;
  for (const cid of Object.keys(byAcc)) {
    const body = { operations: byAcc[cid].map((rn) => ({ updateMask: "status", update: { resourceName: rn, status } })) };
    const r = await fetch(`https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/adGroupCriteria:mutate`, { method: "POST", headers: { "Authorization": `Bearer ${token}`, "developer-token": devToken!, "login-customer-id": mcc, "Content-Type": "application/json" }, body: JSON.stringify(body) });
    const j = await r.json();
    if (j.error) throw new Error(j?.error?.details?.[0]?.errors?.[0]?.message || j.error.message || "erro no Google Ads");
    n += byAcc[cid].length;
  }
  return { ok: true, detail: `${n} palavra(s)-chave ${action === "pause" ? "pausada(s)" : "reativada(s)"}` };
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
// Incluir termo como palavra-chave (no conjunto) ou negativá-lo (na campanha) direto no Google Ads.
async function googleTermAction(m: any) {
  const action = String(m.action || ""); const termo = String(m.termo || "").trim();
  const cid = String(m.accountId || "").replace(/-/g, "");
  if (!cid || !termo) throw new Error("conta e termo obrigatórios");
  const mt = /exact/i.test(m.matchType) ? "EXACT" : /broad/i.test(m.matchType) ? "BROAD" : "PHRASE";
  const token = await googleAdsAccessToken();
  const devToken = Deno.env.get("GOOGLE_ADS_DEV_TOKEN"); const mcc = String(Deno.env.get("GOOGLE_ADS_MCC_ID") || "").replace(/-/g, "");
  const H = { "Authorization": `Bearer ${token}`, "developer-token": devToken!, "login-customer-id": mcc, "Content-Type": "application/json" };
  let url = "", body: any = {};
  if (action === "negative") {
    const campId = String(m.campaignId || "").replace(/[^0-9]/g, "");
    if (!campId) throw new Error("campanha não identificada para negativar este termo");
    url = `https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/campaignCriteria:mutate`;
    body = { operations: [{ create: { campaign: `customers/${cid}/campaigns/${campId}`, negative: true, keyword: { text: termo, matchType: mt } } }] };
  } else {
    const agId = String(m.adGroupId || "").replace(/[^0-9]/g, "");
    if (!agId) throw new Error("conjunto (ad group) não identificado para incluir este termo");
    url = `https://googleads.googleapis.com/${GADS_VER}/customers/${cid}/adGroupCriteria:mutate`;
    body = { operations: [{ create: { adGroup: `customers/${cid}/adGroups/${agId}`, status: "ENABLED", keyword: { text: termo, matchType: mt } } }] };
  }
  const r = await fetch(url, { method: "POST", headers: H, body: JSON.stringify(body) });
  const j = await r.json();
  if (j.error) throw new Error(j?.error?.details?.[0]?.errors?.[0]?.message || j.error.message || "erro no Google Ads");
  return { ok: true, detail: action === "negative" ? `🚫 "${termo}" adicionada como NEGATIVA (${mt}).` : `➕ "${termo}" incluída como palavra-chave (${mt}).` };
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
  const nota = String(m.nota || "").trim();
  const notaTxt = nota ? `\n\nORIENTAÇÃO DO GESTOR (siga à risca, tem prioridade sobre o resto): ${nota.slice(0, 500)}` : "";
  const user = `DNA do cliente:\n${JSON.stringify(ctx)}\n\nTermos de busca (com gasto/cliques/conversões):\n${JSON.stringify(lista)}${notaTxt}`;
  try {
    const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 1500, temperature: 0.3 });
    const parsed = JSON.parse(j.choices[0].message.content || "{}");
    return { negativar: (parsed.negativar || []).slice(0, 60), observacao: parsed.observacao || "", cliente: c?.name || "", temDna: !!(dna && Object.keys(dna).length) };
  } catch (e) { return { erro: String((e as any)?.message || e) }; }
}
// Lista campanhas › conjuntos (ad groups) das contas Google do cliente — pra escolher onde incluir a palavra-chave.
async function googleAdGroups(m: any) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(m.clientId || "")}&select=google_account_id`))[0];
  const ids = String(c?.google_account_id || "").split(",").map((s: string) => s.trim().replace(/-/g, "").replace(/[^0-9]/g, "")).filter(Boolean);
  if (!ids.length) return { adgroups: [] };
  const token = await googleAdsAccessToken();
  const out: any[] = [];
  for (const acc of ids) {
    // só campanhas de PESQUISA aceitam palavras-chave (Performance Max/Shopping não têm ad_group de keyword); traz ativos E pausados (não removidos)
    const rows = await gadsSearch(acc, `SELECT campaign.id, campaign.name, campaign.status, campaign.advertising_channel_type, ad_group.id, ad_group.name, ad_group.status FROM ad_group WHERE ad_group.status != 'REMOVED' AND campaign.status != 'REMOVED' AND campaign.advertising_channel_type = 'SEARCH'`, token).catch(() => []);
    rows.forEach((r: any) => out.push({ accountId: acc, campaignId: String(r.campaign?.id || ""), campaignName: r.campaign?.name || "", campaignStatus: r.campaign?.status || "", adGroupId: String(r.adGroup?.id || ""), adGroupName: r.adGroup?.name || "", adGroupStatus: r.adGroup?.status || "" }));
  }
  return { adgroups: out };
}
// IA GARIMPO: a partir dos termos de busca reais + DNA, sugere NOVAS palavras-chave pra COMPRAR (oportunidades).
async function googleTermMining(m: any) {
  const termos: any[] = (m.termos || []).slice(0, 120);
  const keywords: string[] = (m.keywords || []).map((k: any) => String(k.key || k).toLowerCase());
  if (!termos.length) return { sugestoes: [], observacao: "Sem termos de busca no período." };
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(m.clientId || "")}&select=name,dna,seg`))[0];
  const dna = (c && c.dna) || {};
  const ctx = { marca: dna?.identidade?.marca || c?.name || "", segmento: c?.seg || "", produtos: (dna?.produtos || []).map((p: any) => p.nome).filter(Boolean).slice(0, 15), personas: (dna?.personas || []).map((p: any) => p.titulo).filter(Boolean).slice(0, 8), ressoam: (dna?.diretrizes?.palavrasRessoam || []).slice(0, 20) };
  // só termos que já mostram intenção de compra (clique e/ou conversão), que ainda NÃO são palavra-chave
  const cand = termos.filter((t: any) => (t.clicks || 0) >= 1 && !keywords.includes(String(t.key || "").toLowerCase())).map((t: any) => ({ termo: t.key, cliques: Math.round(t.clicks || 0), conversoes: +(+(t.conversions || 0)).toFixed(1), gasto: Math.round(t.spend || 0) }));
  const nota = String(m.nota || "").trim();
  const sys = `Você é especialista em Google Ads. Recebe o DNA do cliente e os TERMOS DE BUSCA reais que geraram cliques/conversões mas ainda NÃO são palavras-chave. Sua tarefa: garimpar OPORTUNIDADES — sugerir novas PALAVRAS-CHAVE pra COMPRAR (adicionar à conta), priorizando as com intenção de compra e alinhadas ao que o cliente vende. Para cada sugestão, escolha o tipo de correspondência ('phrase' na dúvida; 'exact' se for muito específica; 'broad' só se for ampla e segura) e explique curto por quê. Baseie CADA sugestão num termo real da lista (campo baseTermo = o termo exato de onde veio). NÃO invente termos sem base. Ignore termos irrelevantes/curiosos. Responda SOMENTE JSON: {"sugestoes":[{"palavra":"...","baseTermo":"...","match":"phrase|exact|broad","motivo":"curto"}],"observacao":"1 frase"}.`;
  const user = `DNA:\n${JSON.stringify(ctx)}\n\nTermos candidatos (cliques/conversões/gasto):\n${JSON.stringify(cand.slice(0, 80))}${nota ? `\n\nORIENTAÇÃO DO GESTOR (prioridade): ${nota.slice(0, 400)}` : ""}`;
  try {
    const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content: user }], response_format: { type: "json_object" }, max_tokens: 1500, temperature: 0.4 });
    const parsed = JSON.parse(j.choices[0].message.content || "{}");
    return { sugestoes: (parsed.sugestoes || []).slice(0, 50), observacao: parsed.observacao || "", cliente: c?.name || "", temDna: !!(dna && Object.keys(dna).length) };
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
  const contaKwSet = new Set<string>(); // TODAS as palavras-chave positivas da CONTA (todas as campanhas, sem filtro de impressão) — pro "na conta"
  const addRow = (map: Record<string, any>, key: string, m: any) => {
    if (!map[key]) map[key] = { key, spend: 0, impressions: 0, clicks: 0, conversions: 0, value: 0 };
    const o = map[key];
    o.spend += (Number(m?.costMicros) || 0) / 1e6; o.impressions += Number(m?.impressions) || 0; o.clicks += Number(m?.clicks) || 0;
    o.conversions += Number(m?.conversions) || 0; o.value += Number(m?.conversionsValue) || 0;
  };
  const errors: string[] = [];
  await Promise.all(accounts.flatMap((cid) => [
    // TODAS as palavras-chave positivas da conta (entidade, sem métricas/período) — pra saber se o termo já é keyword em QUALQUER campanha
    gadsSearch(cid, `SELECT ad_group_criterion.keyword.text FROM ad_group_criterion WHERE ad_group_criterion.type = 'KEYWORD' AND ad_group_criterion.negative = FALSE AND ad_group_criterion.status != 'REMOVED' LIMIT 10000`, token)
      .then((rows) => rows.forEach((r: any) => { const t = r.adGroupCriterion?.keyword?.text; if (t) contaKwSet.add(String(t).toLowerCase().trim()); }))
      .catch((e) => errors.push("keywords da conta: " + e.message)),
    gadsSearch(cid, `SELECT segments.conversion_action_name, metrics.conversions, metrics.conversions_value FROM campaign WHERE ${range} AND metrics.conversions > 0`, token)
      .then((rows) => rows.forEach((r: any) => addRow(conv, r.segments?.conversionActionName || "—", r.metrics)))
      .catch((e) => errors.push("conversões: " + e.message)),
    gadsSearch(cid, `SELECT ad_group_criterion.keyword.text, ad_group.id, ad_group_criterion.criterion_id, ad_group_criterion.status, campaign.name, ad_group.name, ${M} FROM keyword_view WHERE ${range} AND metrics.impressions > 0 ORDER BY metrics.cost_micros DESC LIMIT 200`, token)
      .then((rows) => rows.forEach((r: any) => {
        const text = r.adGroupCriterion?.keyword?.text || "—"; addRow(kw, text, r.metrics);
        const o = kw[text]; const agId = r.adGroup?.id, critId = r.adGroupCriterion?.criterionId; const st = r.adGroupCriterion?.status || "";
        if (agId && critId) { (o._refs = o._refs || []).push({ accountId: cid, resourceName: `customers/${cid}/adGroupCriteria/${agId}~${critId}`, status: st, campaign: r.campaign?.name || "", adGroup: r.adGroup?.name || "" }); }
        if (st === "ENABLED") o._anyEnabled = true;
      }))
      .catch((e) => errors.push("keywords: " + e.message)),
    gadsSearch(cid, `SELECT search_term_view.search_term, campaign.id, campaign.name, ad_group.id, ad_group.name, ${M} FROM search_term_view WHERE ${range} ORDER BY metrics.cost_micros DESC LIMIT 200`, token)
      .then((rows) => rows.forEach((r: any) => {
        const key = r.searchTermView?.searchTerm || "—"; addRow(st, key, r.metrics);
        const o = st[key]; const cost = (Number(r.metrics?.costMicros) || 0) / 1e6;
        // guarda a campanha/conjunto onde o termo mais gastou (alvo pra incluir/negativar)
        if (!o._top || cost > o._top.cost) o._top = { cost, accountId: cid, campaignId: r.campaign?.id ? String(r.campaign.id) : null, campaignName: r.campaign?.name || "", adGroupId: r.adGroup?.id ? String(r.adGroup.id) : null, adGroupName: r.adGroup?.name || "" };
      }))
      .catch((e) => errors.push("termos: " + e.message)),
  ]));
  const sorted = (o: Record<string, any>, by = "spend") => Object.values(o).sort((a: any, b: any) => b[by] - a[by]).slice(0, 100);
  return { conversoes: sorted(conv, "conversions"), keywords: sorted(kw), termos: sorted(st), contaKeywords: [...contaKwSet], errors: errors.length ? errors : undefined };
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

// Diagnóstico · leitura da IA: resumo executivo + por que cada criativo ganha/perde (Vision) + o que separa os lados.
async function raioxAI(m: any) {
  const sys = `Você é uma gestora de tráfego senior (nível Pedro Sobral) fazendo o Diagnóstico de um cliente. Recebe: KPIs do período, o objetivo/modelo, o health score por dimensão, e os criativos vencedores e perdedores (com IMAGEM de cada um e a métrica do objetivo). Escreva uma análise afiada, específica e acionável, em português.
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
// Junta material do cliente automaticamente (site + copy dos anúncios do Meta + termos de busca do Google) pra montar o DNA.
// Nao deixa uma fonte lenta derrubar a coleta inteira: passou do tempo, segue sem ela.
function _comLimite<T>(p: Promise<T>, ms: number, fallback: T): Promise<T> {
  let t: number | undefined;
  return Promise.race([
    p.catch(() => fallback).then((v) => { if (t) clearTimeout(t); return v; }),
    new Promise<T>((r) => { t = setTimeout(() => r(fallback), ms); }),
  ]);
}
// A coleta era SEQUENCIAL (site → 3 contas Meta uma a uma → termos do Google). Em cliente grande
// (Curso Fernanda Pessoa: 7 contas Meta + a maior conta Google) isso somava ~90s e, com o modelo em cima,
// estourava o teto de 150s da Edge Function (504). Agora tudo em paralelo e com teto por fonte.
async function _dnaGatherFromAccount(clientId: string): Promise<string> {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,seg,site_url,meta_account_id,google_account_id,gsc_site_url`))[0];
  if (!c) return "";
  const accs = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean).slice(0, 3);
  const gAccs = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  const token = accs.length ? await _metaUserToken() : "";
  const metaCopies = async () => {
    const copies = new Set<string>();
    await Promise.all(accs.map(async (acc) => {
      try {
        const r = await fetch(`https://graph.facebook.com/v21.0/act_${acc}/adcreatives?fields=title,body,object_story_spec{link_data{message,name,description}}&limit=150&access_token=${token}`);
        const j = await r.json();
        (j.data || []).forEach((cr: any) => { const ld = cr.object_story_spec && cr.object_story_spec.link_data; [cr.title, cr.body, ld && ld.message, ld && ld.name, ld && ld.description].forEach((x: any) => { if (x && String(x).trim().length > 3) copies.add(String(x).trim()); }); });
      } catch { /* uma conta fora nao invalida as outras */ }
    }));
    return [...copies];
  };
  const googleTermos = async () => {
    const since = new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
    const br: any = await googleBreakdowns({ accounts: gAccs.map((id: string) => ({ id })), since, until });
    return ((br && br.termos) || []).slice(0, 40).map((t: any) => t.key);
  };
  // Instagram orgânico: as legendas dos posts são a fonte mais fiel do jeito que a marca fala no dia a dia
  // (o anúncio é escrito pela agência; o feed é a voz do cliente).
  const instaLegendas = async () => {
    const r = await fetch(`${_SB_URL}/rest/v1/dim_conteudo_organico?select=legenda,publicado_em,dim_conta!inner(client_id)&dim_conta.client_id=eq.${encodeURIComponent(clientId)}&order=publicado_em.desc&limit=60`, {
      headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Accept-Profile": "midia" },
    });
    if (!r.ok) return [] as string[];
    return ((await r.json()) || []).map((x: any) => String(x.legenda || "").trim()).filter((t: string) => t.length > 25);
  };
  // Perguntas reais dos leads no WhatsApp: é de onde saem as objeções de verdade, com as palavras deles.
  // Só mensagens RECEBIDAS e já mascaradas (telefone/e-mail/documento) — mesma máscara usada nas análises do CRM.
  const perguntasLeads = async () => {
    const convs = await sbGet("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&select=id,name&order=last_at.desc&limit=40`);
    if (!convs.length) return [] as string[];
    const ids = convs.map((x: any) => x.id).slice(0, 40);
    const msgs = await sbGet("wa_messages", `conversation_id=in.(${ids.map((x: string) => encodeURIComponent(x)).join(",")})&direction=eq.in&select=text&order=ts.desc&limit=200`);
    const nomes = convs.map((x: any) => x.name || "");
    const vistos = new Set<string>();
    return (msgs || []).map((m: any) => _crmAiMaskText(m.text, nomes).trim())
      .filter((t: string) => { if (t.length < 15 || t.length > 300 || vistos.has(t)) return false; vistos.add(t); return true; }).slice(0, 45);
  };
  const gscTermos = async () => {
    const until = new Date().toISOString().slice(0, 10), since = new Date(Date.now() - 90 * 864e5).toISOString().slice(0, 10);
    const g: any = await gscReport({ siteUrl: c.gsc_site_url, since, until });
    return ((g && g.termos) || []).slice(0, 30).map((t: any) => t.chave).filter(Boolean);
  };
  const [site, copies, termos, legendas, perguntas, buscas] = await Promise.all([
    c.site_url ? _comLimite(fetchUrlText(c.site_url), 25000, "") : Promise.resolve(""),
    (token && accs.length) ? _comLimite(metaCopies(), 35000, [] as string[]) : Promise.resolve([] as string[]),
    gAccs.length ? _comLimite(googleTermos(), 35000, [] as string[]) : Promise.resolve([] as string[]),
    _comLimite(instaLegendas(), 20000, [] as string[]),
    _comLimite(perguntasLeads(), 20000, [] as string[]),
    c.gsc_site_url ? _comLimite(gscTermos(), 25000, [] as string[]) : Promise.resolve([] as string[]),
  ]);
  const parts: string[] = [`Negócio: ${c.name}. Segmento: ${c.seg || "-"}.`];
  const fontes: string[] = [];
  if (site) { parts.push("=== SITE DO CLIENTE ===\n" + site.slice(0, 6000)); fontes.push("site"); }
  if (legendas.length) { parts.push("=== POSTS DO INSTAGRAM (voz da marca no orgânico) ===\n" + legendas.join("\n---\n").slice(0, 7000)); fontes.push(`${legendas.length} posts do Instagram`); }
  if (copies.length) { parts.push("=== COPY DOS ANÚNCIOS (Meta) ===\n" + copies.slice(0, 70).join("\n")); fontes.push(`${Math.min(copies.length, 70)} copies do Meta`); }
  if (perguntas.length) { parts.push("=== PERGUNTAS/DÚVIDAS REAIS DE LEADS NO WHATSAPP (use pra objeções, nas palavras deles) ===\n" + perguntas.join("\n").slice(0, 4000)); fontes.push(`${perguntas.length} perguntas de leads no WhatsApp`); }
  if (termos.length) { parts.push("=== TERMOS DE BUSCA (Google Ads) ===\n" + termos.join(", ")); fontes.push(`${termos.length} termos do Google Ads`); }
  if (buscas.length) { parts.push("=== BUSCAS QUE LEVAM AO SITE (Search Console) ===\n" + buscas.join(", ")); fontes.push(`${buscas.length} buscas do Search Console`); }
  _dnaFontes = fontes; // devolvido junto com o DNA pra tela mostrar de onde a IA tirou
  return parts.join("\n\n");
}
let _dnaFontes: string[] = [];
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
// Casa o [#ref] injetado pelo link rastreável → origem completa (campanha › grupo › palavra-chave / gclid).
async function _waRefOrigin(text: string): Promise<{ type: string; data: Record<string, unknown> } | null> {
  const mm = String(text || "").match(/\[#([a-z0-9]{6,10})\]/i); if (!mm) return null;
  const row = (await sbGet("wa_ref_origins", `ref=eq.${encodeURIComponent(mm[1])}&select=origin&limit=1`))[0];
  const o = row && row.origin; if (!o) return null;
  const paid = o.channel === "google" || o.channel === "meta" || /cpc|paid|ad/i.test(String(o.medium || ""));
  return { type: paid ? "anuncio" : "utm", data: { channel: o.channel || "", track_source: o.track_source || o.channel || "", campaign: o.campaign || "", adset: o.adgroup || "", ad: o.keyword || "", keyword: o.keyword || "", adgroup: o.adgroup || "", gclid: o.gclid || "", fbclid: o.fbclid || "", medium: o.medium || "" } };
}
async function waCall(host: string, token: string, path: string, method = "GET", payload?: any) {
  const r = await fetch(host.replace(/\/$/, "") + path, { method, headers: { token, "Content-Type": "application/json" }, body: payload ? JSON.stringify(payload) : undefined });
  const t = await r.text(); let j: any; try { j = JSON.parse(t); } catch { j = t; }
  return { status: r.status, j };
}
// Resolve o ad_id do CTWA (source_id da conversa) em nomes: campanha › conjunto › anúncio (Graph API)
async function waResolveAd(adId: string): Promise<Record<string, string> | null> {
  const token = await _metaUserToken(); if (!token || !adId) return null;
  try {
    const r = await fetch(`https://graph.facebook.com/v21.0/${encodeURIComponent(adId)}?fields=name,adset{name},campaign{name},account_id&access_token=${token}`);
    const j = await r.json(); if (!j || j.error) return null;
    return { ad: j.name || "", adset: (j.adset && j.adset.name) || "", campaign: (j.campaign && j.campaign.name) || "", account: String(j.account_id || "") };
  } catch { return null; }
}
// Fallback (via C): resolve o anúncio pelo TÍTULO/corpo do criativo, procurando nos ad accounts do cliente.
function _adNorm(s: any) { return String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim(); }
async function waResolveAdByTitle(title: string, accountIds: string[], body?: string): Promise<Record<string, string> | null> {
  const token = await _metaUserToken(); const target = _adNorm(title); const bTarget = _adNorm(body); if (!token || (!target && !bTarget) || !accountIds.length) return null;
  for (const acc of accountIds) {
    try {
      // inclui object_story_spec e asset_feed_spec: em criativo DINÂMICO o title/body de topo vêm VAZIOS e o texto real fica nesses campos
      let url: string | null = `https://graph.facebook.com/v21.0/act_${acc}/ads?fields=name,adset{name},campaign{name},creative{title,body,object_story_spec,asset_feed_spec}&limit=200&access_token=${token}`;
      let pages = 0;
      while (url && pages < 6) {
        pages++;
        const r = await fetch(url); const j = await r.json(); if (j.error || !Array.isArray(j.data)) break;
        for (const ad of j.data) {
          const cr = ad.creative || {};
          const texts: string[] = [];
          const push = (s: any) => { const n = _adNorm(s); if (n) texts.push(n); };
          push(cr.title); push(cr.body);
          const oss = cr.object_story_spec;
          if (oss) { const ld = oss.link_data || {}; push(ld.message); push(ld.name); push(ld.description); const vd = oss.video_data || {}; push(vd.message); push(vd.title); }
          const afs = cr.asset_feed_spec;
          if (afs) { (afs.bodies || []).forEach((b: any) => push(b.text)); (afs.titles || []).forEach((t: any) => push(t.text)); (afs.descriptions || []).forEach((d: any) => push(d.text)); }
          // corpo (copy) do CTWA casa com algum texto do criativo — mais confiável que o título (headline curto)
          const bodyHit = bTarget && bTarget.length > 15 && texts.some((t) => t.length > 15 && (t === bTarget || t.includes(bTarget) || bTarget.includes(t)));
          const titleHit = target && texts.some((t) => t === target || (t.length > 8 && (t.includes(target) || target.includes(t))));
          if (bodyHit || titleHit) {
            return { ad: ad.name || "", adset: (ad.adset && ad.adset.name) || "", campaign: (ad.campaign && ad.campaign.name) || "", account: acc };
          }
        }
        url = (j.paging && j.paging.next) || null;
      }
    } catch { /* segue pra próxima conta */ }
  }
  return null;
}
// Resolve os IDs do Google (ValueTrack {campaignid}/{adgroupid}) em NOMES, nas contas do cliente.
async function waResolveGoogleCampaign(campaignId: string, adgroupId: string, accountIds: string[]): Promise<{ campaign: string; adgroup: string } | null> {
  if (!accountIds.length || (!campaignId && !adgroupId)) return null;
  try {
    const token = await googleAdsAccessToken();
    for (const acc of accountIds.map((x) => String(x).replace(/-/g, ""))) {
      let campName = "", agName = "";
      if (campaignId) { const rows = await gadsSearch(acc, `SELECT campaign.name FROM campaign WHERE campaign.id = ${campaignId}`, token).catch(() => []); campName = rows[0]?.campaign?.name || ""; }
      if (adgroupId) { const rows = await gadsSearch(acc, `SELECT ad_group.name FROM ad_group WHERE ad_group.id = ${adgroupId}`, token).catch(() => []); agName = rows[0]?.adGroup?.name || ""; }
      if (campName || agName) return { campaign: campName, adgroup: agName };
    }
  } catch { /* */ }
  return null;
}
async function waUzConfig() { const acc = await sbGet("account_config", "id=eq.main&select=data"); const uz = (acc[0]?.data || {}).uazapi || {}; if (!uz.server || !uz.admin_token) throw new Error("uazapi não configurado (aba Configurações → WhatsApp)."); return uz; }
const CRM_DEFAULT_FIELDS = [
  { key: "nome", label: "Nome", type: "texto", hint: "Nome próprio que o lead usou ao se apresentar" },
  { key: "email", label: "Email", type: "texto", hint: "" },
  { key: "produto", label: "Produto/Serviço de Interesse", type: "texto", hint: "O que o lead quer comprar ou contratar" },
  { key: "valor", label: "Valor", type: "valor", hint: "Valor TOTAL do negócio em R$ — se for parcelado (entrada + parcelas), some tudo" },
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
// JORNADA do lead: registra CADA mudança de etapa (IA ou manual) com o porquê e a evidência — auditável pelo gestor.
async function waJourneyLog(o: { conversationId: string; clientId?: string | null; from?: string | null; to?: string | null; source: string; confidence?: number | null; why?: string; evidence?: string; actor?: string | null }) {
  try {
    await sbPost("wa_journey", {
      id: _wuid(), conversation_id: o.conversationId, client_id: o.clientId || null,
      from_stage: o.from || null, to_stage: o.to || null, source: o.source,
      confidence: o.confidence != null ? Math.round(o.confidence) : null,
      why: (o.why || "").slice(0, 900), evidence: (o.evidence || "").slice(0, 600), actor: o.actor || null,
      created_at: new Date().toISOString(),
    });
  } catch (_e) { /* jornada é log: nunca derruba o fluxo */ }
}
// autoApply: aplica a etapa automaticamente se a confiança >= mínimo configurado.
async function waExtract(convId: string, autoApply = false) {
  const cv = (await sbGet("wa_conversations", `id=eq.${encodeURIComponent(convId)}&select=id,name,fields,stage,client_id,origin_type,origin&limit=1`))[0];
  if (!cv) throw new Error("Conversa não encontrada.");
  const msgs = await sbGet("wa_messages", `conversation_id=eq.${encodeURIComponent(convId)}&order=ts.asc&select=direction,text&limit=200`);
  const hasInbound = (msgs || []).some((m: any) => m.direction === "in" && String(m.text || "").trim());
  if (!hasInbound) return { fields: cv.fields || {}, stage: cv.stage || "", confidence: 0, stageWhy: "Sem mensagem recebida do contato para classificar.", applied: false, skipped: true, skipReason: "sem_mensagem_recebida" };
  const transcript = (msgs || []).filter((m: any) => m.text).map((m: any) => `${m.direction === "in" ? "LEAD" : "ATENDENTE"}: ${m.text}`).join("\n").slice(0, 6000);
  if (!transcript) return { fields: cv.fields || {}, stage: "", confidence: 0, stageWhy: "", applied: false };
  // Cada cliente define seu próprio funil, campos e nível de confiança. Nunca reutilize
  // critérios de outro cliente, pois MQL/SQL dependem da operação comercial de cada negócio.
  const clCfg = cv.client_id ? (await sbGet("clients", `id=eq.${encodeURIComponent(cv.client_id)}&select=crm_config&limit=1`))[0]?.crm_config || {} : {};
  const fields = (Array.isArray(clCfg.fields) && clCfg.fields.length) ? clCfg.fields : CRM_DEFAULT_FIELDS;
  const stages = (Array.isArray(clCfg.stages) && clCfg.stages.length) ? clCfg.stages : CRM_DEFAULT_STAGES;
  const minConf = Number(clCfg.min_confidence != null ? clCfg.min_confidence : 70);
  const spec = fields.map((f: any) => `- ${f.key} (${f.label}${f.type ? ", tipo " + f.type : ""})${f.hint ? ": " + f.hint : ""}`).join("\n");
  const stageSpec = stages.filter((s: any) => s.key !== "sem").map((s: any) => `- ${s.key} = ${s.label}${s.desc ? ": " + s.desc : ""}`).join("\n");
  const keys = stages.map((s: any) => s.key).join(", ");
  // DNA do cliente pra avaliar relevância
  let dnaCtx = ""; let hasDna = false;
  if (cv.client_id) { const cl = (await sbGet("clients", `id=eq.${encodeURIComponent(cv.client_id)}&select=name,dna,seg`))[0]; const dna = cl?.dna || {}; const prods = (dna?.produtos || []).map((p: any) => p.nome).filter(Boolean); hasDna = !!(prods.length || dna?.identidade?.marca); dnaCtx = `\n\nNEGÓCIO DO CLIENTE (contexto): ${cl?.name || ""} · segmento ${cl?.seg || ""}. Vende: ${prods.join(", ") || dna?.identidade?.marca || "—"}. Personas: ${(dna?.personas || []).map((p: any) => p.titulo).filter(Boolean).join(", ") || "—"}.${(dna?.objecoes || []).length ? ` Objeções conhecidas e como responder: ${(dna.objecoes || []).slice(0, 8).map((o: any) => `"${o.objecao}" → ${o.resposta}`).join(" | ")}` : ""}`; }
  // O ANÚNCIO que o lead respondeu — quem pergunta sobre o que foi anunciado é RELEVANTE por definição
  let adCtx = ""; const _o = cv.origin || {};
  if (cv.origin_type === "anuncio" && (_o.title || _o.body || _o.campaign || _o.ad)) adCtx = `\n\nESTE LEAD VEIO DE UM ANÚNCIO. O que foi anunciado: "${String(_o.title || "").slice(0, 120)}${_o.body ? " — " + String(_o.body).slice(0, 200) : ""}"${_o.campaign ? ` (campanha: ${_o.campaign})` : ""}. Se o interesse do lead bate com esse anúncio, ele é RELEVANTE — o cliente está pagando justamente para atrair essas pessoas.`;
  const sys = `Você é um SDR que lê uma conversa de WhatsApp entre o LEAD e o ATENDENTE. Faça: (1) extraia os campos do lead — só o que aparece claramente, NÃO invente; para tipo 'valor' devolva o valor TOTAL do negócio, só o número: se o pagamento for parcelado (ex.: "entrada de 590 + 2 boletos de 590 cada"), SOME entrada e todas as parcelas pra chegar no total (590+590+590=1770) — NUNCA devolva só a entrada ou só uma parcela isolada; se depois a conversa trouxer um valor total mais atualizado/diferente, use o mais recente; (2) CLASSIFIQUE a etapa do funil usando as descrições ESPECÍFICAS DESTE CLIENTE, com 'confidence' 0-100; (3) 'numeroErrado'=true SÓ se ficar CLARO que o número está errado (a pessoa diz que não é quem procuramos, mandou errado, não conhece, pediu pra parar); (4) 'irrelevante'=true quando a conversa NÃO tem intenção de cumprir nenhum dos objetivos descritos no funil deste cliente, considerando DNA, histórico, anúncio e contexto. As descrições das ETAPAS DO FUNIL fornecidas na mensagem seguinte são a autoridade máxima para diferenciar Lead novo, MQL, SQL e demais etapas; nunca substitua esses critérios por uma definição genérica de mercado. Considere também objetivos alternativos legítimos, como recrutamento, currículos, suporte, distribuição e reconhecimento, quando estiverem previstos no anúncio, DNA ou nas regras do cliente. REGRAS CRÍTICAS de relevância: NUNCA marque irrelevante um contato que demonstra interesse no que foi ANUNCIADO. NUNCA use o segmento sozinho pra decidir. ${hasDna ? "" : "Não há dados suficientes do negócio, então "}na dúvida, ou faltando contexto, deixe irrelevante=FALSE. REGRA DE VENDA: só escolha uma etapa de compra/venda quando houver a confirmação exigida na descrição daquela etapa; menção à palavra venda, encaminhamento ao comercial ou envio de orçamento não comprovam fechamento. (5) DÊ UMA NOTA AO ATENDIMENTO da equipe nesta conversa, de 0 a 10 ('notaAtendimento'), com uma frase curta de motivo ('notaPorque'). Avalie a equipe, não o lead: velocidade percebida, entendimento da necessidade, clareza, personalização, tratamento de objeção e próximo passo combinado. Lead que escreveu e ficou sem resposta é nota BAIXA — é atendimento perdido. Use null na nota quando a conversa não for atendimento a cliente/lead (irrelevante) ou quando a equipe ainda não respondeu nada. Responda SOMENTE JSON.`;
  const content = `CAMPOS A EXTRAIR:\n${spec}\n\nETAPAS DO FUNIL (escolha UMA key):\n${stageSpec}${dnaCtx}${adCtx}\n\nCONVERSA:\n${transcript}\n\nResponda JSON: {"fields":{"<key>":"<valor>"}, "stage":"<key entre: ${keys}>", "confidence":<0-100>, "stageWhy":"<explique em 1-2 frases POR QUE essa etapa, citando o que o lead fez/disse>", "evidencia":"<a FRASE do lead (copiada da conversa) que mais prova essa classificação>", "numeroErrado":<bool>, "irrelevante":<bool>, "irrelevanteMotivo":"<curto se irrelevante>", "notaAtendimento":<0-10 ou null>, "notaPorque":"<uma frase curta sobre a condução do atendimento>"}`;
  const j = await callOpenAI({ model: "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content }], response_format: { type: "json_object" }, max_tokens: 800, temperature: 0.2, _telemetry: { clientId: cv.client_id || null, action: "crm_classification" } });
  let parsed: any = {}; try { parsed = JSON.parse(j.choices[0].message.content || "{}"); } catch { parsed = {}; }
  const jobInquiry = /curr[ií]cul|vaga de emprego|oportunidade de emprego|quero trabalhar|trabalhar com voc[eê]s|est[aã]o contratando|falar com (o )?rh/i.test(transcript);
  const allowsJobGoal = /recrut|curr[ií]cul|vaga|emprego|trabalhar|rh/i.test(`${stageSpec}\n${dnaCtx}\n${adCtx}`);
  if (jobInquiry && !allowsJobGoal) { parsed.irrelevante = true; parsed.irrelevanteMotivo = "currículo/procura de emprego"; parsed.stageWhy = "Contato sem intenção de compra e sem objetivo de recrutamento configurado para este cliente."; parsed.confidence = Math.max(95, Number(parsed.confidence) || 0); }
  const outFields = { ...(cv.fields || {}), ...(parsed.fields || {}) };
  /* Nota do atendimento em TODA conversa, e não só nas sorteadas pela avaliação em lote: sai na mesma
     chamada que já classifica o funil, então custa alguns tokens a mais e nada de chamada nova.
     Conversa que não é atendimento (irrelevante) não recebe nota — ver [[capa-relevancia-conversa]].
     A avaliação em lote continua mandando: ela é mais profunda e grava em fields.capa. */
  {
    const n = parsed.notaAtendimento;
    if (!parsed.irrelevante && n != null && n !== "" && !isNaN(Number(n))) {
      outFields.nota = { score: Math.max(0, Math.min(10, Math.round(Number(n)))), why: String(parsed.notaPorque || "").slice(0, 300), at: new Date().toISOString(), origem: "automatica" };
    } else if (parsed.irrelevante && outFields.nota) {
      delete outFields.nota;   // virou irrelevante depois de ter nota: a nota deixa de fazer sentido
    }
  }
  const lostStage = stages.find((s: any) => /perd|lost|desqual/i.test(String(s.key || "") + " " + String(s.label || "")));
  let stage = (parsed.stage && stages.some((s: any) => s.key === parsed.stage)) ? parsed.stage : "";
  if (parsed.irrelevante && lostStage) { stage = lostStage.key; parsed.confidence = Math.max(90, Number(parsed.confidence) || 0); parsed.stageWhy = parsed.stageWhy || `Contato fora da intenção de compra: ${parsed.irrelevanteMotivo || "assunto fora da oferta"}.`; }
  const conf = Math.max(0, Math.min(100, Math.round(Number(parsed.confidence) || 0)));
  const patch: Record<string, unknown> = { fields: outFields, ai_stage: stage, ai_conf: conf, ai_why: parsed.stageWhy || "", ai_at: new Date().toISOString(), num_errado: !!parsed.numeroErrado, irrelevante: !!parsed.irrelevante, irrelevante_motivo: parsed.irrelevante ? (parsed.irrelevanteMotivo || "") : null };
  const applied = !!(autoApply && stage && conf >= minConf && stage !== cv.stage);
  if (applied) patch.stage = stage;
  await sbPatchD("wa_conversations", `id=eq.${encodeURIComponent(convId)}`, patch);
  // JORNADA: registra quando a IA MUDA a etapa; e também quando ela quis mudar mas a confiança ficou abaixo do mínimo (sugestão)
  const _lbl = (k: string) => (stages.find((s: any) => s.key === k) || {}).label || k || "—";
  if (applied) {
    await waJourneyLog({ conversationId: convId, clientId: cv.client_id, from: cv.stage || null, to: stage, source: "ia", confidence: conf, why: parsed.stageWhy || "", evidence: parsed.evidencia || "", actor: "AndréIA" });
  } else if (stage && stage !== cv.stage && conf > 0) {
    await waJourneyLog({ conversationId: convId, clientId: cv.client_id, from: cv.stage || null, to: stage, source: "ia_sugestao", confidence: conf, why: `Sugerido (não aplicado: confiança ${conf}% < mínimo ${minConf}%). ${parsed.stageWhy || ""}`.trim(), evidence: parsed.evidencia || "", actor: "AndréIA" });
  }
  if (applied) { const stObj = stages.find((s: any) => s.key === stage); if (stObj && stObj.event) { try { await waCapi(convId, stObj.event); } catch (_e) {} } }
  return { fields: outFields, stage, confidence: conf, stageWhy: parsed.stageWhy || "", evidencia: parsed.evidencia || "", applied, minConf, stageLabel: _lbl(stage) };
}
// ---- CAPI (Conversions API): manda o evento da etapa pro Meta, atribuindo ao anúncio via ctwa_clid ----
const _pixelCache: Record<string, string | null> = {};
async function _sha256hex(s: string) { const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)); return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join(""); }
async function clientPixelId(clientId: string): Promise<string | null> {
  if (_pixelCache[clientId] !== undefined) return _pixelCache[clientId];
  const token = await _metaUserToken();
  const cli = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=meta_account_id,pixel_id`))[0];
  let pid: string | null = null;
  // 1) pixel manual do cadastro tem prioridade
  const manual = String(cli?.pixel_id || "").replace(/[^0-9]/g, "").trim();
  if (manual) { _pixelCache[clientId] = manual; return manual; }
  // 2) senão, pega o 1º pixel da conta de anúncio
  const acct = String(cli?.meta_account_id || "").split(",")[0].replace(/^act_/, "").trim();
  if (acct && token) { try { const r = await fetch(`https://graph.facebook.com/v21.0/act_${acct}/adspixels?fields=id&limit=1&access_token=${token}`); const j = await r.json(); pid = (j.data && j.data[0] && j.data[0].id) || null; } catch { pid = null; } }
  _pixelCache[clientId] = pid; return pid;
}
async function waCapi(convId: string, eventName: string) {
  const cv = (await sbGet("wa_conversations", `id=eq.${encodeURIComponent(convId)}&select=id,client_id,chat_id,origin`))[0];
  if (!cv) throw new Error("Conversa não encontrada.");
  const token = await _metaUserToken();
  const pid = cv.client_id ? await clientPixelId(cv.client_id) : null;
  const logId = _wuid();
  // Sem cliente vinculado (ex: número da agência) ou sem pixel → não dá pra atribuir; pula em silêncio (não é erro real, não loga falha).
  if (!pid || !token) { return { ok: false, skipped: true, error: "sem pixel/cliente" }; }
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
// Metas do cliente valendo pra um canal: {…geral, porCanal:{meta,google}} — o que o canal nao define cai pro Geral.
// Formato antigo (so os campos soltos) continua valendo como Geral.
const _BENCH_KEYS = ["ctr", "cpc", "cpm", "roas", "cpa", "cpl", "custoConversa", "custoView"];
function _benchFor(bm: any, canal: string) {
  if (!bm) return {};
  const ch = (canal === "meta" || canal === "google") ? ((bm.porCanal || {})[canal] || {}) : {};
  const out: any = {};
  for (const k of _BENCH_KEYS) { const v = (ch[k] != null && ch[k] !== "") ? ch[k] : bm[k]; if (v != null && v !== "" && +v !== 0) out[k] = +v; }
  return out;
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
  // r30 é só do Meta: usa as metas do canal Meta, caindo pro Geral campo a campo (mesma regra do painel)
  const bm = _benchFor(c.benchmark_metas, "meta"); const t = (r30 && r30.total) || null;
  if (t && Object.values(bm).some((v) => v != null)) {
    const cpConv = t.conversas ? t.spend / t.conversas : null, cpl = t.leads ? t.spend / t.leads : null, cpa = t.purchases ? t.spend / t.purchases : null;
    const cmp = (meta: any, atual: any, menorMelhor: boolean) => (meta == null || atual == null) ? null : { meta, atual: +atual.toFixed(2), atingida: menorMelhor ? atual <= meta : atual >= meta };
    const mvr: any = { ctr: cmp(bm.ctr, t.ctr, false), cpc: cmp(bm.cpc, t.cpc, true), cpm: cmp(bm.cpm, t.cpm, true), roas: cmp(bm.roas, t.roas, false), custoConversa: cmp(bm.custoConversa, cpConv, true), cpl: cmp(bm.cpl, cpl, true), cpa: cmp(bm.cpa, cpa, true) };
    Object.keys(mvr).forEach((k) => { if (!mvr[k]) delete mvr[k]; }); if (Object.keys(mvr).length) out.metasVsReal_30d = mvr;
  }
  // por campanha: linha de KPIs PRINCIPAIS já pronta — Gasto · Resultado · CPR (pelo objetivo da campanha)
  if (r30 && r30.campaigns && r30.campaigns.length) out.campanhasComGasto_30d = r30.campaigns.slice(0, 25).map((x: any) => ({ nome: x.campaign, objetivo: (x.objetivo && x.objetivo.rotulo) || "", kpi: `Gasto ${_fmtR(x.spend || 0)} · ${_objRC(x, false, x.objetivo && x.objetivo.tipo)}` }));
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
  · orcamento: {"tipo":"orcamento","campanha":"<nome>","novoValor":<novo orçamento diário em R$, número>,"conjunto":"<nome do conjunto, SÓ se a pessoa mencionar um específico - a campanha pode ter orçamento por CAMPANHA (CBO) ou por CONJUNTO, o sistema descobre sozinho qual é>"}
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
// Resolve o CONJUNTO (adset) de uma campanha p/ ajuste de orçamento quando ela nao tem CBO (orcamento no
// nivel da campanha) — precisa achar o conjunto certo. Devolve: null (sem conjunto nenhum), 1 objeto (achou
// exato ou so tem 1 conjunto na campanha - auto-seleciona) ou array (ambiguo, varios conjuntos, pergunta).
async function waResolveAdset(clientId: string, campanhaId: string, nomeConjunto?: string): Promise<any> {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=meta_account_id`))[0];
  const ids = String(c?.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) return null;
  const ent = await metaEntities({ accounts: ids.map((id: string) => ({ id, name: id })) }).catch(() => null);
  if (!ent) return null;
  const conjuntos = (ent.adsets || []).filter((a: any) => a.campanhaId === campanhaId);
  if (!conjuntos.length) return null;
  if (nomeConjunto) {
    const q = String(nomeConjunto).toLowerCase().trim();
    return conjuntos.find((a: any) => a.nome.toLowerCase() === q) || conjuntos.find((a: any) => a.nome.toLowerCase().includes(q)) || conjuntos.find((a: any) => q.includes(a.nome.toLowerCase())) || conjuntos;
  }
  return conjuntos.length === 1 ? conjuntos[0] : conjuntos;
}
// Lista as campanhas (nome/id/status) das contas Meta do cliente — pra oferecer opções quando a campanha não foi identificada.
async function waListCampaigns(clientId: string): Promise<any[]> {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=meta_account_id`))[0];
  const ids = String(c?.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) return [];
  const ent = await metaEntities({ accounts: ids.map((id: string) => ({ id, name: id })) }).catch(() => null);
  return (ent && ent.campaigns) || [];
}
const _waAcaoLabel: Record<string, string> = { pausar_campanha: "pausar", reativar_campanha: "reativar", orcamento: "ajustar o orçamento da", duplicar_campanha: "duplicar" };
// Texto perguntando QUAL campanha, já filtrando pelo estado que faz sentido (pausar→ativas, reativar→pausadas).
async function _waCampaignPickText(cid: string, tipo: string): Promise<string> {
  const list = await waListCampaigns(cid);
  const isActive = (c: any) => String(c.status || c.entrega || "").toUpperCase() === "ACTIVE";
  const filt = tipo === "reativar_campanha" ? list.filter((c: any) => !isActive(c)) : list.filter(isActive);
  const use = filt.length ? filt : list;
  const names = use.map((c: any) => c.nome).filter(Boolean);
  const rot = tipo === "reativar_campanha" ? "pausadas" : "ativas";
  const lab = _waAcaoLabel[tipo] || "ajustar";
  const nm = await _waClientNome(cid);
  return names.length
    ? `Qual campanha ${rot} do ${nm} você quer ${lab}? Me diz o nome exato:\n${names.slice(0, 25).map((n: string) => "• " + n).join("\n")}${names.length > 25 ? `\n…e mais ${names.length - 25}` : ""}`
    : `Não achei campanhas ${rot} nesse cliente pra ${lab}.`;
}
// ===== CLIENTE NOVO: contrato em PDF (template no account_config) + envio como documento no WhatsApp =====
const _EXT_N: Record<number, string> = { 1: "um", 2: "dois", 3: "três", 4: "quatro", 5: "cinco", 6: "seis", 7: "sete", 8: "oito", 9: "nove", 10: "dez", 11: "onze", 12: "doze" };
// valor por extenso em reais (até 999.999) — pro contrato ("R$ 800,00 (oitocentos reais)")
function _valorExtenso(v: number): string {
  const u = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"];
  const d = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];
  const c = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"];
  const trio = (n: number): string => {
    if (n === 0) return ""; if (n === 100) return "cem";
    const cc = Math.floor(n / 100), rest = n % 100, dd = Math.floor(rest / 10), uu = rest % 10;
    const parts: string[] = [];
    if (cc) parts.push(c[cc]);
    if (rest) { if (rest < 20) parts.push(u[rest]); else { parts.push(d[dd]); if (uu) parts.push(u[uu]); } }
    return parts.join(" e ");
  };
  const inteiro = Math.floor(Math.abs(v)), cents = Math.round((Math.abs(v) - inteiro) * 100);
  let s = "";
  const mil = Math.floor(inteiro / 1000), rest = inteiro % 1000;
  if (mil) s += (mil === 1 ? "mil" : trio(mil) + " mil") + (rest ? (rest < 100 || rest % 100 === 0 ? " e " : " ") : "");
  if (rest) s += trio(rest);
  if (!s) s = "zero";
  s += inteiro === 1 ? " real" : " reais";
  if (cents) s += " e " + trio(cents) + (cents === 1 ? " centavo" : " centavos");
  return s;
}
const _MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
// preenche o template do contrato (account_config id=contract_template) com os dados do cliente; faltantes viram linha em branco
async function _contratoTexto(f: any): Promise<{ text: string; faltando: string[] }> {
  const row = (await sbGet("account_config", "id=eq.contract_template&select=data"))[0];
  const tpl = row?.data?.template;
  if (!tpl) throw new Error("Modelo de contrato não encontrado no sistema (account_config/contract_template).");
  const BLANK = "______________________";
  const faltando: string[] = [];
  const get = (v: any, label: string) => { const s = String(v || "").trim(); if (!s) { faltando.push(label); return BLANK; } return s; };
  const isPF = /^(pf|f[íi]sica|pessoa f)/i.test(String(f.tipoPessoa || "")) || (String(f.docNumero || "").replace(/\D/g, "").length === 11);
  const valor = Number(f.valorMensal) || 0;
  const promo = Number(f.mesesPromo) || 0;
  const fid = Number(f.mesesFidelidade) || promo || 3;
  const vig = Number(f.mesesVigencia) || fid;
  let remun: string;
  if (valor > 0 && promo > 0) remun = `Condição promocional. Pela prestação dos serviços, a CONTRATANTE pagará à CONTRATADA o valor mensal promocional de ${_fmtR(valor).replace("R$", "R$ ")} (${_valorExtenso(valor)}) durante os ${promo} (${_EXT_N[promo] || promo}) primeiros meses de vigência (período promocional).`;
  else if (valor > 0) remun = `Pela prestação dos serviços, a CONTRATANTE pagará à CONTRATADA o valor mensal de ${_fmtR(valor).replace("R$", "R$ ")} (${_valorExtenso(valor)}), a título de honorários de gestão.`;
  else { remun = `Pela prestação dos serviços, a CONTRATANTE pagará à CONTRATADA o valor mensal de R$ ${BLANK}.`; faltando.push("honorário mensal"); }
  const now = _spNow();
  const dataExt = `${now.getUTCDate()} de ${_MESES_PT[now.getUTCMonth()]} de ${now.getUTCFullYear()}`;
  const map: Record<string, string> = {
    RAZAO_SOCIAL: get(f.razaoSocial || f.nome, "razão social/nome"),
    TIPO_PESSOA: isPF ? "pessoa física" : "pessoa jurídica de direito privado",
    DOC_TIPO: isPF ? "CPF" : "CNPJ",
    DOC_NUMERO: get(f.docNumero, isPF ? "CPF" : "CNPJ"),
    ENDERECO: get(f.endereco, "endereço"),
    EMAIL: get(f.email, "e-mail"),
    TELEFONE: get(f.telefone, "telefone"),
    REPRESENTANTE: get(f.representante, "representante"),
    CPF_REPRESENTANTE: get(f.cpfRepresentante, "CPF do representante"),
    TELEFONE_REPRESENTANTE: get(f.telefoneRepresentante || f.telefone, "telefone do representante"),
    CLAUSULA_REMUNERACAO: remun,
    MESES_FIDELIDADE: String(fid), MESES_FIDELIDADE_EXTENSO: _EXT_N[fid] || String(fid),
    MESES_VIGENCIA: String(vig), MESES_VIGENCIA_EXTENSO: _EXT_N[vig] || String(vig),
    DATA_EXTENSO: dataExt,
  };
  const text = String(tpl).replace(/\{\{(\w+)\}\}/g, (_m, k) => map[k] != null ? map[k] : "______");
  return { text, faltando: [...new Set(faltando)] };
}
// tira caracteres fora do WinAnsi (Helvetica não codifica emoji etc.)
function _winAnsi(s: string): string { return String(s || "").normalize("NFC").replace(/[•]/g, "•").replace(/[^\x20-\x7E\xA0-\xFF–—‘’“”•…]/g, ""); }
// gera o PDF do contrato na identidade GT (faixa escura + dourado, seções com barra dourada, rodapé de marca)
async function _contratoPdf(text: string): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const GOLD = rgb(0.784, 0.627, 0.306);     // #c8a04e — dourado GT
  const DARK = rgb(0.071, 0.082, 0.106);     // #12151b — fundo da marca
  const INK = rgb(0.104, 0.114, 0.141);      // texto
  const MUT = rgb(0.42, 0.45, 0.50);
  const HAIR = rgb(0.87, 0.88, 0.90);
  const W = 595.28, H = 841.89, M = 56, LW = W - M * 2;
  let page: any, y = 0, first = true;
  const header = () => {
    if (first) { // capa: faixa escura com a marca
      page.drawRectangle({ x: 0, y: H - 96, width: W, height: 96, color: DARK });
      page.drawText("GT MARKETING", { x: M, y: H - 52, size: 21, font: bold, color: GOLD });
      page.drawText(_winAnsi("G E S T A O   D E   T R A F E G O   D I G I T A L"), { x: M, y: H - 70, size: 7.2, font, color: rgb(0.62, 0.65, 0.70) });
      y = H - 96 - 34;
    } else { // demais páginas: linha fina dourada + marca discreta
      page.drawText("GT MARKETING", { x: M, y: H - 44, size: 8, font: bold, color: GOLD });
      page.drawLine({ start: { x: M, y: H - 52 }, end: { x: W - M, y: H - 52 }, thickness: 0.6, color: HAIR });
      y = H - 76;
    }
    first = false;
  };
  const newPage = () => { page = doc.addPage([W, H]); header(); };
  newPage();
  const wrap = (s: string, f: any, size: number, width: number): string[] => {
    const words = s.split(/\s+/); const out: string[] = []; let line = "";
    for (const wd of words) { const t = line ? line + " " + wd : wd; if (f.widthOfTextAtSize(t, size) > width && line) { out.push(line); line = wd; } else line = t; }
    if (line) out.push(line); return out;
  };
  const draw = (s: string, f: any, size: number, color: any, indent = 0, spacing = 4.5) => {
    for (const ln of wrap(s, f, size, LW - indent)) {
      if (y < M + 34) newPage();
      page.drawText(ln, { x: M + indent, y, size, font: f, color });
      y -= size + spacing;
    }
  };
  for (const raw of text.split("\n")) {
    const line = _winAnsi(raw.trim());
    if (!line) { y -= 6.5; continue; }
    if (raw.startsWith("# ")) { // título do contrato: centralizado + régua dourada
      y -= 6;
      const t = line.slice(2), size = 15;
      for (const ln of wrap(t, bold, size, LW)) { if (y < M + 60) newPage(); page.drawText(ln, { x: (W - bold.widthOfTextAtSize(ln, size)) / 2, y, size, font: bold, color: INK }); y -= size + 5; }
      page.drawLine({ start: { x: W / 2 - 40, y: y + 4 }, end: { x: W / 2 + 40, y: y + 4 }, thickness: 2, color: GOLD });
      y -= 18; continue;
    }
    if (raw.startsWith("## ")) { // seção: barra dourada + título
      y -= 12; if (y < M + 70) newPage();
      const t = line.slice(3), size = 10.5;
      page.drawRectangle({ x: M, y: y - 1.5, width: 3, height: size + 2, color: GOLD });
      for (const ln of wrap(t, bold, size, LW - 12)) { page.drawText(ln, { x: M + 11, y, size, font: bold, color: INK }); y -= size + 4; }
      page.drawLine({ start: { x: M, y: y + 5 }, end: { x: W - M, y: y + 5 }, thickness: 0.5, color: HAIR });
      y -= 10; continue;
    }
    if (line.startsWith("•")) { // bullet com ponto dourado
      const t = line.replace(/^•\s*/, "");
      if (y < M + 34) newPage();
      page.drawCircle({ x: M + 6, y: y + 3, size: 1.6, color: GOLD });
      draw(t, font, 10, INK, 16, 3.8); continue;
    }
    draw(line, font, 10, INK, 0, 3.8);
  }
  // rodapé de marca em todas as páginas
  const pages = doc.getPages();
  pages.forEach((p, i) => {
    p.drawLine({ start: { x: M, y: 44 }, end: { x: W - M, y: 44 }, thickness: 0.5, color: HAIR });
    p.drawText(_winAnsi("GT Marketing · Gestao de Trafego Digital"), { x: M, y: 31, size: 7.5, font, color: MUT });
    const pg = `${i + 1} / ${pages.length}`;
    p.drawText(pg, { x: W - M - font.widthOfTextAtSize(pg, 7.5), y: 31, size: 7.5, font, color: GOLD });
  });
  return await doc.save();
}
// sobe o PDF pro Storage (bucket PRIVADO docs; cria na primeira vez) e devolve um link assinado e
// temporario (1h) - so pra dar tempo da uazapi buscar o arquivo e anexar na mensagem do WhatsApp;
// nunca fica um link publico permanente igual a URL fixa (nome do cliente + timestamp e previsivel,
// virou vazamento real em 2026-08-06 quando o bucket estava marcado publico - ver rls-policy memo).
async function _uploadDoc(bytes: Uint8Array, path: string): Promise<string> {
  const H = { Authorization: `Bearer ${_SB_KEY}`, apikey: _SB_KEY };
  await fetch(`${_SB_URL}/storage/v1/bucket`, { method: "POST", headers: { ...H, "Content-Type": "application/json" }, body: JSON.stringify({ id: "docs", name: "docs", public: false }) }).catch(() => null); // já existe → 400, ok
  const up = await fetch(`${_SB_URL}/storage/v1/object/docs/${path}`, { method: "POST", headers: { ...H, "Content-Type": "application/pdf", "x-upsert": "true" }, body: bytes });
  if (!up.ok) throw new Error("upload do PDF falhou: " + (await up.text()).slice(0, 200));
  const signRes = await fetch(`${_SB_URL}/storage/v1/object/sign/docs/${path}`, { method: "POST", headers: { ...H, "Content-Type": "application/json" }, body: JSON.stringify({ expiresIn: 3600 }) });
  const signJson = await signRes.json().catch(() => ({}));
  if (!signRes.ok || !signJson.signedURL) throw new Error("gerar link assinado do PDF falhou: " + JSON.stringify(signJson).slice(0, 200));
  return `${_SB_URL}/storage/v1${signJson.signedURL}`;
}
// manda um DOCUMENTO no WhatsApp via uazapi (tenta /send/media type=document; fallback /send/document)
async function _waSendDoc(host: string, token: string, number: string, url: string, docName: string, caption = ""): Promise<boolean> {
  try { const r = await waCall(host, token, "/send/media", "POST", { number, type: "document", file: url, docName, caption }); if (r.status >= 200 && r.status < 300) return true; } catch (_e) { /* */ }
  try { const r = await waCall(host, token, "/send/document", "POST", { number, file: url, filename: docName, caption }); return r.status >= 200 && r.status < 300; } catch (_e) { return false; }
}
function _slugDoc(s: string) { return String(s || "cliente").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40) || "cliente"; }
// checklist de onboarding por canal (só cria as tarefas dos canais que o cliente vai trabalhar)
const _ONB_GERAL: [string, number][] = [["Reunião de kickoff / alinhamento inicial", 1], ["Criar grupo de WhatsApp com o cliente", 1], ["Coletar acessos (BM, contas de anúncio, site, redes)", 2], ["Briefing completo / montar DNA do cliente", 3], ["Instalar pixel / configurar rastreamento", 5]];
const _ONB_CANAL: Record<string, [string, number][]> = {
  meta: [["Configurar conta de anúncios Meta", 3], ["Estruturar campanhas Meta (funil + públicos)", 5], ["Subir campanhas Meta", 7]],
  google: [["Configurar conta Google Ads + ações de conversão", 3], ["Levantamento de palavras-chave (Google)", 5], ["Estruturar e subir campanhas Google", 7]],
  tiktok: [["Configurar conta TikTok Ads", 3], ["Estruturar e subir campanhas TikTok", 7]],
};
function _canaisNorm(canais: any): string[] {
  const arr = Array.isArray(canais) ? canais : String(canais || "").split(/[,e/&+]/);
  const out: string[] = [];
  for (const c of arr) { const s = String(c).toLowerCase(); if (/meta|face|insta/.test(s)) out.push("meta"); else if (/google|ads|pesquisa|youtube/.test(s)) out.push("google"); else if (/tik/.test(s)) out.push("tiktok"); }
  return [...new Set(out)];
}

// ESTADO do processo de cliente novo (guardado na sessão) — é o que impede a IA de repetir uma etapa já feita
const _FLOW_STEPS: Record<string, string> = { gerar_contrato: "contrato", cadastrar_cliente: "cadastro", criar_tarefas_onboarding: "tarefas", criar_lancamento: "financeiro" };
function _flowAfter(pending: any, flow: any) {
  const step = _FLOW_STEPS[pending?.tipo];
  if (!step) return flow || null;
  const nome = pending.nomeSistema || pending.nome || pending.razaoSocial || pending.cliente || (flow && flow.nome) || "";
  // contrato inicia (ou reinicia, se for outro cliente) o processo
  const mesmoCliente = flow && flow.nome && nome && String(flow.nome).toLowerCase() === String(nome).toLowerCase();
  const base = (pending.tipo === "gerar_contrato" && !mesmoCliente) ? { tipo: "cliente_novo", nome, feito: {}, desde: new Date().toISOString() } : (flow || { tipo: "cliente_novo", nome, feito: {}, desde: new Date().toISOString() });
  if (!base.nome && nome) base.nome = nome;
  base.feito = { ...(base.feito || {}), [step]: new Date().toISOString() };
  return base;
}
// texto injetado no prompt: o que já foi feito e qual é o próximo passo
function _flowPrompt(flow: any): string {
  if (!flow || !flow.feito) return "";
  const ordem = [["contrato", "gerar o contrato em PDF"], ["cadastro", "cadastrar o cliente no sistema"], ["tarefas", "criar as tarefas de onboarding"], ["financeiro", "criar o lançamento financeiro (fee)"]];
  const feitos = ordem.filter(([k]) => flow.feito[k]).map(([, l]) => l);
  const prox = ordem.find(([k]) => !flow.feito[k]);
  if (!feitos.length) return "";
  const hAgo = Math.round((Date.now() - new Date(flow.desde || Date.now()).getTime()) / 60000);
  return `\n\n⚠️ PROCESSO DE CLIENTE NOVO EM ANDAMENTO${flow.nome ? ` — cliente: ${flow.nome}` : ""} (iniciado há ${hAgo} min):
- JÁ CONCLUÍDO nesta conversa: ${feitos.join(", ")}. **NUNCA proponha nem repita essas etapas de novo** — já estão feitas.
- PRÓXIMO PASSO: ${prox ? prox[1] : "nada — o processo terminou; feche com um resumo do que foi feito"}.
Se o usuário responder algo que não é "sim/não" (ex: informar o nicho, a verba, os canais, ou dizer "isso você já fez"), NÃO volte pra etapa anterior: entenda o que ele disse, use como dado do próximo passo e siga em frente.`;
}
async function waAgentExec(pending: any, clientId: string | null, waCtx?: { host: string; token: string; number: string }) {
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
      if (camp.orcamentoDiario) {
        await metaAction({ action: "budget", id: camp.id, nome: camp.nome, novoOrcamentoDiario: pending.novoValor });
        return `💰 Orçamento de "${camp.nome}" ajustado pra R$${Number(pending.novoValor).toFixed(2)}/dia.`;
      }
      // sem CBO: o orçamento é por CONJUNTO — resolve o conjunto certo em vez de desistir.
      const conj = await waResolveAdset(cid, camp.id, pending.conjunto);
      if (!conj) return `A "${camp.nome}" não tem orçamento no nível da campanha, e não achei nenhum conjunto nela. Quer que eu crie uma tarefa pra ajustar?`;
      if (Array.isArray(conj)) {
        const nomes = conj.map((a: any) => a.nome).filter(Boolean);
        return `A "${camp.nome}" tem orçamento por CONJUNTO, não por campanha${pending.conjunto ? ` (não achei "${pending.conjunto}" exato)` : ""}. Qual conjunto?\n${nomes.slice(0, 20).map((n: string) => "• " + n).join("\n")}`;
      }
      if (!conj.orcamentoDiario) return `O conjunto "${conj.nome}" também não tem orçamento diário próprio (pode ser vitalício). Quer que eu crie uma tarefa pra ajustar?`;
      await metaAction({ action: "budget", id: conj.id, nome: conj.nome, novoOrcamentoDiario: pending.novoValor });
      return `💰 Orçamento do conjunto "${conj.nome}" (campanha "${camp.nome}") ajustado pra R$${Number(pending.novoValor).toFixed(2)}/dia.`;
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
    // ===== PROCESSO CLIENTE NOVO =====
    if (pending.tipo === "gerar_contrato") {
      const { text, faltando } = await _contratoTexto(pending);
      const bytes = await _contratoPdf(text);
      const fname = `contrato-${_slugDoc(pending.razaoSocial || pending.nome)}-${Date.now()}.pdf`;
      const url = await _uploadDoc(bytes, fname);
      let sentDoc = false;
      if (waCtx) sentDoc = await _waSendDoc(waCtx.host, waCtx.token, waCtx.number, url, `Contrato GT Marketing - ${pending.razaoSocial || pending.nome || "cliente"}.pdf`, "📄 Contrato pronto pra revisão e assinatura");
      const falta = faltando.length ? `\n⚠ Ficaram em branco (preenche antes de assinar): *${faltando.join(", ")}*.` : "";
      const link = sentDoc ? "" : `\n📎 Baixe aqui: ${url}`;
      const nomeCtr = String(pending.razaoSocial || pending.nome || "").trim();
      return `📄 *Contrato gerado!*${link}${falta}\n\n*Próximo passo:* quer que eu já *cadastre o cliente no sistema*?\n• *Nome no sistema:* uso "${nomeCtr}" (do contrato) ou prefere outro? (nome fantasia/social — é o que aparece no painel, relatórios e na minha conversa)\n• *Nicho*, *verba de mídia mensal* e *canais* (Meta/Google/TikTok) — se não tiver ainda, cadastro assim mesmo e você completa depois.`;
    }
    if (pending.tipo === "cadastrar_cliente") {
      // nome NO SISTEMA (fantasia/social) pode ser diferente do nome do contrato (razão social)
      const nome = String(pending.nomeSistema || pending.nome || pending.razaoSocial || "").trim();
      const razao = String(pending.razaoSocial || "").trim();
      if (!nome) return "Qual nome você quer usar pro cliente *no sistema*? (pode ser o nome fantasia/social — diferente da razão social do contrato)";
      const dup = await sbGet("clients", `name=ilike.${encodeURIComponent(nome)}&select=id&limit=1`);
      if (dup.length) return `Já existe um cliente "${nome}" no sistema — não dupliquei. Quer seguir pras *tarefas de onboarding* dele? Me confirma os canais (Meta/Google/TikTok) e o responsável.`;
      const canais = _canaisNorm(pending.canais);
      const chMap: Record<string, string> = { meta: "Meta Ads", google: "Google Ads", tiktok: "TikTok Ads" };
      // guarda a razão social nas observações quando o nome do sistema for diferente (não se perde o nome do contrato)
      const difRazao = razao && razao.toLowerCase() !== nome.toLowerCase();
      const row: Record<string, unknown> = { id: _wuid(), name: nome, seg: pending.nicho || "", status: "Ativo", fee: Number(pending.valorMensal || pending.fee) || 0, budget: Number(pending.verba) || 0, active_channels: canais.map((c) => chMap[c] || c), notes: "Criado via AndréIA (cliente novo)" + (difRazao ? ` · Razão social (contrato): ${razao}` : "") + (pending.docNumero ? ` · Doc: ${pending.docNumero}` : "") + (pending.nicho ? "" : " — completar nicho") + (Number(pending.verba) ? "" : " — completar verba de mídia") };
      const res = await sbInsertOk("clients", row);
      if (!res.ok) return "❌ Não consegui cadastrar: " + res.err;
      const faltou: string[] = [];
      if (!pending.nicho) faltou.push("nicho");
      if (!Number(pending.verba)) faltou.push("verba de mídia");
      if (!canais.length) faltou.push("canais");
      const faltaTxt = faltou.length ? `\n⚠ Faltou: *${faltou.join(", ")}* — completa depois no cadastro do sistema.` : "";
      return `✅ Cliente *${nome}* cadastrado no sistema${difRazao ? ` (razão social no contrato: ${razao} — guardei nas observações)` : ""}${pending.nicho ? ` · nicho: ${pending.nicho}` : ""}${Number(pending.verba) ? ` · verba: ${_fmtR(Number(pending.verba))}` : ""}${canais.length ? ` · canais: ${canais.map((c) => chMap[c]).join(" + ")}` : ""}.${faltaTxt}\n\n*Próximo passo:* quer que eu crie as *tarefas de onboarding*? Me confirma os *canais de mídia* que ele vai trabalhar (crio só as tarefas desses canais) e *quem é o responsável*.`;
    }
    if (pending.tipo === "criar_tarefas_onboarding") {
      if (!cid) {
        const nome = String(pending.cliente || pending.nomeSistema || pending.nome || "").trim();
        if (nome) { const hit = await sbGet("clients", `name=ilike.*${encodeURIComponent(nome)}*&select=id&limit=1`); if (hit.length) pending.client_id = hit[0].id; }
      }
      const cid2 = pending.client_id || cid;
      if (!cid2) return "De qual cliente são as tarefas de onboarding?";
      const canais = _canaisNorm(pending.canais);
      if (!canais.length) return "Quais canais esse cliente vai trabalhar? (Meta, Google, TikTok…) — crio só as tarefas desses canais.";
      const team = await sbGet("team", "select=id,name");
      let owner = "eu";
      if (pending.responsavel) { const q = String(pending.responsavel).toLowerCase(); const tm = team.find((t: any) => t.name.toLowerCase() === q) || team.find((t: any) => t.name.toLowerCase().includes(q)); if (tm) owner = tm.id; }
      const base = _spNow(); const dueDay = (d: number) => { const dt = new Date(base.getTime() + d * 864e5); return dt.toISOString().slice(0, 10); };
      const lista: [string, number][] = [..._ONB_GERAL, ...canais.flatMap((c) => _ONB_CANAL[c] || [])];
      let n = 0;
      for (const [nomeT, dias] of lista) { const r = await sbInsertOk("tasks", { id: _wuid(), name: nomeT, client: cid2, owner, status: "todo", prio: "media", notes: "Onboarding (via AndréIA)", due: dueDay(dias) }); if (r.ok) n++; }
      const cn = await _waClientNome(cid2);
      const chMap: Record<string, string> = { meta: "Meta", google: "Google", tiktok: "TikTok" };
      return `✅ *${n} tarefas de onboarding* criadas${cn ? ` pro cliente ${cn}` : ""} (${canais.map((c) => chMap[c]).join(" + ")}), distribuídas nos próximos 7 dias.\n\n*Pra fechar o processo:* quer que eu crie o *lançamento financeiro* (fee mensal) desse cliente? Me diz o valor e o vencimento — ou "não" pra encerrar.`;
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
  // Google: metrics.conversions é genérico (form/WhatsApp/ligação/compra/etc.) — não afirmar "Compras" nem ROAS sem confirmação de venda real (planilha).
  if (isVenda && google) return `Conversões ${Math.round(t.purchases || 0)} · Custo/conv ${_fmtR(t.purchases ? spend / t.purchases : 0)}`;
  if (isVenda) { const roas = t.roas != null ? t.roas : (spend ? (t.revenue || 0) / spend : 0); return `Compras ${Math.round(t.purchases || 0)} · ROAS ${(roas || 0).toFixed(2)}`; }
  // total do Google não separa leads/conversas por tipo (só o detalhamento por campanha/anúncio) — cai pro total bruto de conversões.
  if (isLead) { const cnt = t.leads || (google ? Math.round(t.purchases || 0) : 0); return `Leads ${cnt} · CPL ${_fmtR(cnt ? spend / cnt : 0)}`; }
  if (isMsg) { const cnt = t.conversas || (google ? Math.round(t.purchases || 0) : 0); return `Conversas ${cnt} · Custo/conversa ${_fmtR(cnt ? spend / cnt : 0)}`; }
  if (isVideo) return `Views ${Math.round(t.videoViews || 0)} · Custo/view ${_fmtR(t.videoViews ? spend / t.videoViews : 0)}`;
  if (isAlcance) { const cpm = t.cpm != null ? t.cpm : (t.impressions ? spend / t.impressions * 1000 : 0); const freq = t.reach ? t.impressions / t.reach : 0; return `Alcance ${Math.round(t.reach || 0).toLocaleString("pt-BR")} · CPM ${_fmtR(cpm)}${freq ? ` · Freq ${freq.toFixed(2)}` : ""}`; }
  if (isEngaj) return `Engajamentos ${Math.round(t.engajamentos || 0).toLocaleString("pt-BR")} · Custo ${_fmtR(t.engajamentos ? spend / t.engajamentos : 0)}`;
  const ctr = t.ctr != null ? t.ctr : (t.impressions ? (t.clicks / t.impressions * 100) : 0);
  const cpc = t.cpc != null ? t.cpc : (t.clicks ? spend / t.clicks : 0);
  return `Cliques ${t.clicks || 0} · CTR ${(ctr || 0).toFixed(2)}% · CPC ${_fmtR(cpc)}`;
}
// KPI enxuto por objetivo: só RESULTADO + CPR (custo por resultado) — pra linha "Gasto · Resultado · CPR" por campanha.
function _objRC(t: any, google: boolean, obj?: string | null): string {
  const spend = t.spend || 0;
  const isVenda = obj === "conversao" || obj === "app" || (!obj && (t.purchases || 0) > 0);
  const isLead = obj === "leads" || (!obj && !google && (t.leads || 0) > 0);
  const isMsg = obj === "mensagens" || (!obj && !google && (t.conversas || 0) > 0);
  const isVideo = obj === "video";
  const isAlcance = obj === "alcance" || obj === "distribuicao";
  const isEngaj = obj === "engajamento";
  const n = (v: number) => Math.round(v || 0).toLocaleString("pt-BR");
  // Google: metrics.conversions é genérico (form/WhatsApp/ligação/compra/etc.) — não afirmar "Compras" (só a planilha de VENDAS confirma venda real)
  if (isVenda && google) { const br = Array.isArray(t.convActions) && t.convActions.length ? ` (${t.convActions.slice(0, 4).map((a: any) => `${a.name}: ${n(a.count)}`).join(", ")})` : ""; return `Conversões ${n(t.purchases)} · Custo/conv. ${_fmtR(t.purchases ? spend / t.purchases : 0)}${br}`; }
  if (isVenda) return `Compras ${n(t.purchases)} · CPA ${_fmtR(t.purchases ? spend / t.purchases : 0)}`;
  if (isLead) { const cnt = t.leads || (google ? Math.round(t.purchases || 0) : 0); return `Leads ${n(cnt)} · CPL ${_fmtR(cnt ? spend / cnt : 0)}`; }
  if (isMsg) { const cnt = t.conversas || (google ? Math.round(t.purchases || 0) : 0); return `Conversas ${n(cnt)} · Custo/conversa ${_fmtR(cnt ? spend / cnt : 0)}`; }
  if (isVideo) return `Views ${n(t.videoViews)} · Custo/view ${_fmtR(t.videoViews ? spend / t.videoViews : 0)}`;
  if (isAlcance) { const cpm = t.cpm != null ? t.cpm : (t.impressions ? spend / t.impressions * 1000 : 0); return `Alcance ${n(t.reach)} · CPM ${_fmtR(cpm)}`; }
  if (isEngaj) return `Engajamentos ${n(t.engajamentos)} · Custo/eng ${_fmtR(t.engajamentos ? spend / t.engajamentos : 0)}`;
  const cpc = t.cpc != null ? t.cpc : (t.clicks ? spend / t.clicks : 0);
  return `Cliques ${n(t.clicks)} · CPC ${_fmtR(cpc)}`;
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
  try { const rows = await sbGet("agent_knowledge", "select=title,text,client_id&order=created_at.desc&limit=20"); const g = (rows || []).filter((r: any) => !r.client_id).slice(0, 10); if (g.length) extra = "\n\nMÉTODOS DA AGÊNCIA (base de conhecimento):\n" + g.map((k: any) => `- ${k.title}: ${String(k.text || "").slice(0, 2500)}`).join("\n"); } catch (_e) { /* */ }
  _waPbCache = WA_PLAYBOOK_BASE + extra; _waPbT = Date.now(); return _waPbCache;
}
// Análise do gestor POR CANAL, cada um julgado pelo SEU objetivo. Retorna cliente -> texto (1 linha por canal).
async function _waAnalises(items: any[]): Promise<Record<string, string>> {
  try {
    const data = items.map((r: any) => {
      const canais: any[] = [];
      if (r.meta && (r.meta.spend || 0) > 0) canais.push({ canal: "Meta", objetivo: _objLabel(r.objMeta), metricas: _objMetric(r.meta, false, r.objMeta), gasto: Math.round(r.meta.spend) });
      if (r.google && (r.google.spend || 0) > 0) canais.push({ canal: "Google", objetivo: _objLabel(r.objGoogle), metricas: _objMetric(r.google, true, r.objGoogle), gasto: Math.round(r.google.spend) });
      return { cliente: r.nome, canais, crm: r.crm && r.crm.total ? { leads: r.crm.total, qualificados: r.crm.qualificados, taxaQualificacao: r.crm.taxaQualificacao, vendas: r.crm.vendas } : null };
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
Se o cliente tiver 'crm' (funil de leads), adicione UMA frase sobre a QUALIFICAÇÃO: muitos leads com baixa taxa de qualificação → revisar segmentação/qualidade; alta qualificação → escalar; use o playbook (custo alto ≠ ruim se qualifica).
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
      const crm = await waCrmStats(c.id, Math.max(days, 30)).catch(() => null);
      return { nome: c.name, meta: mt, google: gt, objMeta, objGoogle, crm, restr: escopo === "com_restricao" ? _clientRestrictions(c, restr) : [] };
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
    const crmLine = (r.crm && r.crm.total) ? `\n🗣 CRM — ${r.crm.total} leads · ${r.crm.qualificados} qualif. (${r.crm.taxaQualificacao}%)${r.crm.vendas ? ` · ${r.crm.vendas} vendas` : ""}` : "";
    if (!ran) { if (!showNon) { if (crmLine) blocks.push(`*${r.nome}* — ⏸ sem tráfego${crmLine}`); continue; } blocks.push(`*${r.nome}* — ⏸ não rodou no período${crmLine}`); continue; }
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
    if (crmLine) b += (completo ? "\n" : "") + crmLine;
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
  // Aberto em 15/08: a AndréIA enxergava só 15 tabelas e ficava sem resposta sobre metade do sistema.
  // Fora daqui ficam SÓ credenciais e controle de acesso (secure_credentials, app_users, app_access_audit,
  // security_events, account_config) — dado de negócio ela vê tudo.
  team: "equipe da agência (id, name) — use pra resolver o responsável de tarefas",
  shows: "shows/eventos vendidos (client, name, date, val, status, payments jsonb) — aba Shows",
  event_projects: "projetos de evento/lançamento (client_id, nome, status)",
  event_editions: "edições de um evento (project_id, nome, data_inicio, data_fim, janela do evento)",
  event_snapshots: "números consolidados por edição de evento (edition_id, dados jsonb)",
  sales_snapshots: "venda consolidada por período/janela (client_id, date, revenue, orders) — aba Vendas",
  channel_metrics_daily: "métricas por canal e por DIA (client_id, channel=meta|google|ga4, date, spend, impressions, clicks, purchases, revenue, leads, campaign, adset, ad_content) — histórico do banco de mídia",
  calendar_events: "reuniões do Google Agenda já sincronizadas (client_id, title, start_at, end_at)",
  briefing: "briefings criativos (client_id, status, objetivo, funil, canais, created_at)",
  briefing_analise: "etapa 1 do briefing: leitura de performance (briefing_id, leitura, funis_json, padroes_json)",
  briefing_curadoria: "etapa 2: curadoria do orgânico que pode virar recorte (briefing_id, candidatos jsonb)",
  briefing_ficha: "etapa 3: fichas de criativo pra produção (briefing_id, titulo, roteiro, formato, status)",
  crm_capa_audits: "auditorias de Qualidade do Atendimento (client_id, stage, audited, average_score, aggregate jsonb, created_at)",
  crm_capa_cases: "casos avaliados na Qualidade do Atendimento (audit_id, conversation_id, score, diagnosis, break_point, recommended_message)",
  raiox: "Diagnóstico (Raio-X) salvo por cliente (client_id, since, until, data jsonb, generated_at)",
  lead_people: "pessoas/leads identificados na jornada (client_id, nome, email, telefone, first_seen, last_seen)",
  lead_touchpoints: "toques da jornada do lead: cada clique/visita/conversa até a venda (client_id, person_id, channel, source, campaign, ad, occurred_at, kind)",
  wa_journey: "log auditável de mudança de etapa no CRM (conversation_id, de, para, motivo, evidencia, quem, created_at)",
  journey_quality_events: "eventos de qualidade da jornada (client_id, tipo, detalhe, created_at)",
  creative_miner_items: "garimpo de criativos (client_id, plataforma, titulo, url, status)",
  report_templates: "modelos de relatório (id, name, objectiveType, metrics jsonb)",
  report_layouts: "layouts de relatório por cliente (client_id, layout jsonb)",
  agent_knowledge: "materiais de conhecimento do cliente que a AndréIA usa (client_id, title, content)",
  andreia_memory: "memória da AndréIA (client_id, tema, conteudo, created_at)",
  andreia_automations: "avisos automáticos configurados no grupo do WhatsApp (titulo, tipo, hora, enabled, last_run)",
  wa_instances: "instâncias de WhatsApp conectadas (client_id, name, phone, status, connected_at) — pra saber se o WhatsApp do cliente está no ar",
  system_usage_events: "consumo do sistema por serviço (service_key, action, input_units, output_units, occurred_at) — custo de IA e APIs",
  system_cost_services: "catálogo de custo dos serviços (service_key, nome, preço)",
};
const WA_TOOLS = [
  { type: "function", function: { name: "consultar_banco", description: "Consulta SOMENTE LEITURA de qualquer tabela do sistema pra buscar dados reais (cliente, financeiro, tarefas, conversas do CRM, RD, pedidos etc). SEMPRE use antes de responder sobre dados guardados.", parameters: { type: "object", properties: { tabela: { type: "string", enum: Object.keys(WA_TABLES) }, colunas: { type: "string", description: "colunas separadas por vírgula ou '*'" }, filtro: { type: "string", description: "filtro no formato PostgREST, ex: 'client=eq.<id>&status=eq.pendente'; datas: 'due=gte.2026-07-01&due=lte.2026-07-31'; texto: 'description=ilike.*fee*'. Vazio = sem filtro." }, ordenar: { type: "string", description: "ex: 'created_at.desc' ou 'due.asc'" }, limite: { type: "integer" } }, required: ["tabela"] } } },
  { type: "function", function: { name: "meta_insights", description: "Métricas de Meta Ads AO VIVO de UM cliente no período. Retorna consolidado + campanhas (KPIs pelo objetivo). Use nivel='campanha' (PADRÃO) pro resumo; só use nivel='conjunto' ou 'anuncio' quando pedirem EXPLICITAMENTE pra detalhar conjuntos/anúncios de uma campanha.", parameters: { type: "object", properties: { cliente: { type: "string", description: "nome do cliente" }, dias: { type: "integer", description: "7, 30 ou 90 (padrão 7)" }, nivel: { type: "string", enum: ["campanha", "conjunto", "anuncio"], description: "profundidade do detalhamento. PADRÃO 'campanha'. Só desça se o usuário pedir." } }, required: ["cliente"] } } },
  { type: "function", function: { name: "google_insights", description: "Métricas de Google Ads AO VIVO de UM cliente no período.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "meta_saldo", description: "Saldo pré-pago (pix/boleto) restante nas contas de anúncio Meta Ads de UM cliente. Use quando pedirem 'saldo', 'saldo restante', 'saldo da carteira' ou similar.", parameters: { type: "object", properties: { cliente: { type: "string" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "instagram_organico", description: "Posts do INSTAGRAM ORGÂNICO de UM cliente (aba Social), já ranqueados do melhor pro pior por engajamento e alcance, com legenda, formato (reel/carrossel/imagem), data, link e métricas. USE quando pedirem curadoria/seleção dos MELHORES CRIATIVOS ou POSTS, o que performou no feed, ideias do que impulsionar, melhor formato ou melhor dia de publicação. É o conteúdo do PERFIL (orgânico) — não confunda com anúncio, que vem de meta_insights.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer", description: "30, 60 ou 90 (padrão 90)" }, quantidade: { type: "integer", description: "quantos posts trazer no ranking (padrão 10)" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "google_keywords", description: "Palavras-chave e termos de busca do Google Ads de UM cliente no período (por palavra: gasto, cliques, conversões, CPC). USE isto quando perguntarem 'como está cada palavra-chave', keywords, termos de busca ou o que as pessoas pesquisaram.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer", description: "7, 30 ou 90 (padrão 7)" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "resumo_todos_clientes", description: "Resumo de TODOS os clientes no período (gasto + métrica do objetivo, Meta/Google separados). Use quando pedirem panorama/todos os clientes.", parameters: { type: "object", properties: { dias: { type: "integer" } } } } },
  { type: "function", function: { name: "relatorio_cliente", description: "Gera um RELATÓRIO VISUAL e limpo de UM cliente PRONTO PRA ENVIAR AO CLIENTE (investimento, resultados pelo objetivo, alcance e uma análise). Use quando pedirem 'relatório do [cliente]', 'manda o relatório pro cliente', 'relatório pra enviar'. Envie o campo 'relatorio' EXATAMENTE como vier.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer", description: "7, 30 ou 90 (padrão 7)" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "crm_funil", description: "Funil do CRM (WhatsApp) de UM cliente: leads por etapa (novo/MQL/SQL/comprou), taxa de qualificação, vendas e origem (anúncio×orgânico). USE quando pedirem sobre leads, funil, qualificação, atendimento ou conversas do cliente.", parameters: { type: "object", properties: { cliente: { type: "string" }, dias: { type: "integer", description: "padrão 30" } }, required: ["cliente"] } } },
  { type: "function", function: { name: "reunioes", description: "REUNIÕES/compromissos da AGENDA (Google Agenda), que é DIFERENTE de tarefa operacional. USE isto quando perguntarem sobre reuniões, agenda, compromissos, calls. NÃO liste tarefas comuns aqui.", parameters: { type: "object", properties: { quando: { type: "string", description: "'hoje', 'amanha', 'semana' ou vazio (padrão hoje)" }, data: { type: "string", description: "data específica AAAA-MM-DD (opcional)" } } } } },
  { type: "function", function: { name: "financeiro", description: "Consulta financeira com TOTAL e itens já com o nome do cliente resolvido e a soma correta. USE ISSO pra qualquer pergunta de dinheiro (a receber, a pagar, recebido, pago, fluxo do mês).", parameters: { type: "object", properties: { tipo: { type: "string", enum: ["receita", "despesa"] }, status: { type: "string", enum: ["pendente", "pago"] }, mes: { type: "string", description: "AAAA-MM, ex: 2026-07" }, cliente: { type: "string" } } } } },
  { type: "function", function: { name: "preparar_acao", description: "Prepara uma AÇÃO de alto impacto pra CONFIRMAÇÃO (NÃO executa agora — o sistema pede SIM). Para criar_tarefa, o RESPONSÁVEL (quem faz) e o QUANDO (data) são obrigatórios — se o usuário não disser, PERGUNTE antes.", parameters: { type: "object", properties: { tipo: { type: "string", enum: ["criar_tarefa", "criar_reuniao", "cancelar_reuniao", "pausar_campanha", "reativar_campanha", "orcamento", "duplicar_campanha", "criar_lancamento", "dar_baixa", "gerar_contrato", "cadastrar_cliente", "criar_tarefas_onboarding"] }, cliente: { type: "string" }, razaoSocial: { type: "string", description: "razão social/nome do CONTRATANTE — o nome JURÍDICO que vai no contrato (gerar_contrato)" }, nomeSistema: { type: "string", description: "nome do cliente COMO DEVE APARECER NO SISTEMA (fantasia/social) — pode ser diferente da razão social do contrato (cadastrar_cliente)" }, docNumero: { type: "string", description: "CNPJ ou CPF do contratante" }, endereco: { type: "string" }, email: { type: "string" }, telefone: { type: "string" }, representante: { type: "string" }, cpfRepresentante: { type: "string" }, telefoneRepresentante: { type: "string" }, valorMensal: { type: "number", description: "honorário mensal em R$ (contrato/fee)" }, mesesPromo: { type: "number", description: "meses de condição promocional (0 se não houver)" }, mesesFidelidade: { type: "number" }, nicho: { type: "string" }, verba: { type: "number", description: "verba de mídia mensal em R$" }, canais: { type: "array", items: { type: "string" }, description: "canais de mídia do cliente: meta, google, tiktok" }, nome: { type: "string", description: "título da tarefa OU da reunião (pra cancelar_reuniao, o título/pedaço do nome da reunião a cancelar). NÃO inclua 'urgente' nem 'revisão' no título — use os campos próprios." }, responsavel: { type: "string", description: "nome de quem vai fazer a tarefa (membro da equipe)" }, quando: { type: "string", description: "data em AAAA-MM-DD (calcule 'amanhã', 'sexta' etc. a partir de hoje) — usada por tarefa e reunião" }, hora: { type: "string", description: "horário da reunião em HH:MM (opcional)" }, urgente: { type: "boolean", description: "true se a tarefa foi pedida como URGENTE — marca a flag de urgência (NÃO escreva 'urgente' no título/obs)" }, revisao: { type: "boolean", description: "true se pediram para SOLICITAR REVISÃO da tarefa — marca a flag de revisão (NÃO escreva 'revisão' no título/obs)" }, obs: { type: "string" }, campanha: { type: "string" }, novoValor: { type: "number" }, natureza: { type: "string", enum: ["receita", "despesa"] }, descricao: { type: "string" }, valor: { type: "number" }, vencimento: { type: "string" } }, required: ["tipo"] } } },
];
const WA_MENU_TEXT = `🤖 *AndréIA — o que posso fazer aqui no grupo:*

📊 *Análise*
• _Como tá o [cliente] nos últimos 7 dias?_
• _Detalhes das campanhas do [cliente]_
• _Como está cada palavra-chave do [cliente]?_ (Google)
• _Relatório do [cliente] pra enviar_ (layout pronto pro cliente)
• _Funil do CRM do [cliente]?_ (leads, qualificação, vendas)
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

🆕 *Cliente novo*
• _Cliente novo: [nome + dados]_ → processo completo: contrato em PDF → cadastro no sistema → tarefas por canal → financeiro (confirmo cada etapa)

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
// Funil do CRM (WhatsApp) de UM cliente: leads por etapa, taxa de qualificação, vendas, origem. null se o cliente não tem CRM.
async function waCrmStats(clientId: string, dias = 30) {
  if (!clientId) return null;
  const since = new Date(Date.now() - dias * 864e5).toISOString();
  const rows = await sbGet("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&last_at=gte.${since}&select=stage,origin_type,num_errado,irrelevante&limit=5000`);
  if (!rows || !rows.length) return null;
  const c: any = { total: 0, sem: 0, novo: 0, mql: 0, sql: 0, comprou: 0, posvenda: 0, perdido: 0, anuncio: 0, organico: 0, numErrado: 0, irrelevante: 0 };
  rows.forEach((r: any) => { c.total++; const st = r.stage || "sem"; if (c[st] != null) c[st]++; else c.sem++; if (r.origin_type === "anuncio") c.anuncio++; else c.organico++; if (r.num_errado) c.numErrado++; if (r.irrelevante) c.irrelevante++; });
  const qualificados = c.mql + c.sql + c.comprou + c.posvenda;
  const vendas = c.comprou + c.posvenda;
  return { dias, total: c.total, etapas: { novo: c.novo, mql: c.mql, sql: c.sql, comprou: c.comprou, posvenda: c.posvenda, perdido: c.perdido, semEtapa: c.sem }, qualificados, vendas, taxaQualificacao: c.total ? +(qualificados / c.total * 100).toFixed(1) : 0, taxaConversao: c.total ? +(vendas / c.total * 100).toFixed(1) : 0, deAnuncio: c.anuncio, organico: c.organico, numeroErrado: c.numErrado, irrelevantes: c.irrelevante };
}

function _crmAiChannel(cv: any) {
  const o = cv.origin || {}, src = String(o.track_source || "").toLowerCase();
  if (o.channel === "google" || /google|gads|adwords/.test(src)) return "google";
  if (cv.origin_type === "anuncio") return "meta";
  if (!cv.origin_type || cv.origin_type === "organico") return "organico";
  return String(o.channel || src || "outro").toLowerCase();
}
function _crmAiSafeFields(fields: any) {
  const out: Record<string, string> = {};
  Object.entries(fields || {}).forEach(([k, v]) => {
    if (/nome|name|telefone|phone|whats|email|e-mail|cpf|cnpj|document|endereco|address/i.test(k)) return;
    const s = _crmAiMaskText(v).trim(); if (s) out[k] = s.slice(0, 240);
  });
  return out;
}
function _crmAiMaskText(v: any, knownNames: string[] = []) {
  let s = String(v || "")
    .replace(/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/g, "[email oculto]")
    .replace(/(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?9?\d{4}[-\s]?\d{4}/g, "[telefone oculto]")
    .replace(/\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b/g, "[documento oculto]")
    .replace(/\b\d{2}\.?\d{3}\.?\d{3}\/\d{4}-?\d{2}\b/g, "[documento oculto]")
    .replace(/\b\d{5}-?\d{3}\b/g, "[CEP oculto]")
    .replace(/\b(?:https?:\/\/|www\.)\S+/gi, "[link oculto]")
    .replace(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g, "[IP oculto]")
    .replace(/\b[A-Z0-9]{12,}\b/gi, "[identificador oculto]");
  for (const name of knownNames.filter((x) => String(x || "").trim().length >= 3)) {
    const escaped = String(name).trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    s = s.replace(new RegExp(escaped, "gi"), "[nome oculto]");
  }
  return s;
}

let _metaTokenCache = "", _metaTokenCacheAt = 0;
function _b64Bytes(a: Uint8Array) { let s = ""; for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode(...a.subarray(i, i + 0x8000)); return btoa(s); }
function _fromB64(s: string) { const b = atob(s), a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
async function _credentialKey() {
  const raw = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`central-gestao:${_SB_KEY}`));
  return await crypto.subtle.importKey("raw", raw, "AES-GCM", false, ["encrypt", "decrypt"]);
}
async function _encryptCredential(value: string) {
  const iv = crypto.getRandomValues(new Uint8Array(12)), key = await _credentialKey();
  const encrypted = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(value)));
  return `${_b64Bytes(iv)}.${_b64Bytes(encrypted)}`;
}
async function _decryptCredential(value: string) {
  const [iv64, data64] = String(value || "").split("."); if (!iv64 || !data64) return "";
  const key = await _credentialKey();
  return new TextDecoder().decode(await crypto.subtle.decrypt({ name: "AES-GCM", iv: _fromB64(iv64) }, key, _fromB64(data64)));
}
async function _metaUserToken() {
  if (_metaTokenCache && Date.now() - _metaTokenCacheAt < 300000) return _metaTokenCache;
  try {
    const r = await fetch(`${_SB_URL}/rest/v1/secure_credentials?id=eq.meta_user_token&select=secret_cipher&limit=1`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
    const rows = r.ok ? await r.json() : []; if (rows[0]?.secret_cipher) _metaTokenCache = await _decryptCredential(rows[0].secret_cipher);
  } catch (_e) { /* usa o secret legado */ }
  if (!_metaTokenCache) _metaTokenCache = Deno.env.get("META_USER_TOKEN") || "";
  _metaTokenCacheAt = Date.now(); return _metaTokenCache;
}
async function metaTokenUpdate(input: any, authorization: string) {
  if (!authorization) throw new Error("Sessão administrativa obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } });
  const user = ur.ok ? await ur.json() : null; const email = String(user?.email || "").toLowerCase();
  const ar = email ? await fetch(`${_SB_URL}/rest/v1/secure_credential_admins?email=eq.${encodeURIComponent(email)}&select=email&limit=1`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } }) : null;
  const admins = ar?.ok ? await ar.json() : [];
  if (!admins.length) throw new Error("Somente o administrador principal pode trocar credenciais globais.");
  const token = String(input?.token || "").trim(); if (token.length < 40) throw new Error("Token do Meta inválido ou incompleto.");
  const test = await fetch(`https://graph.facebook.com/v21.0/me/adaccounts?fields=id,name&limit=5&access_token=${encodeURIComponent(token)}`); const tj = await test.json();
  if (!test.ok || tj.error) throw new Error(`Meta: ${tj.error?.message || "token não autorizado"}`);
  const cipher = await _encryptCredential(token);
  const save = await fetch(`${_SB_URL}/rest/v1/secure_credentials?on_conflict=id`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify({ id: "meta_user_token", secret_cipher: cipher, updated_at: new Date().toISOString() }) });
  if (!save.ok) throw new Error("Não consegui guardar o token no cofre privado.");
  _metaTokenCache = token; _metaTokenCacheAt = Date.now();
  return { ok: true, contasTestadas: (tj.data || []).length, atualizadoEm: new Date().toISOString() };
}

async function _requireCredentialAdmin(authorization: string) {
  if (!authorization) throw new Error("Sessão administrativa obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } });
  const user = ur.ok ? await ur.json() : null; const email = String(user?.email || "").toLowerCase();
  const ar = email ? await fetch(`${_SB_URL}/rest/v1/secure_credential_admins?email=eq.${encodeURIComponent(email)}&select=email&limit=1`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } }) : null;
  const admins = ar?.ok ? await ar.json() : [];
  if (!admins.length) throw new Error("Somente o administrador principal pode alterar credenciais oficiais.");
  return email;
}
async function _saveSecureCredential(id: string, value: string) {
  const cipher = await _encryptCredential(value);
  const r = await fetch(`${_SB_URL}/rest/v1/secure_credentials?on_conflict=id`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify({ id, secret_cipher: cipher, updated_at: new Date().toISOString() }) });
  if (!r.ok) throw new Error("Não consegui guardar a credencial no cofre privado.");
}
async function _loadSecureCredential(id: string) {
  const rows = await sbGet("secure_credentials", `id=eq.${encodeURIComponent(id)}&select=secret_cipher&limit=1`);
  return rows[0]?.secret_cipher ? await _decryptCredential(rows[0].secret_cipher) : "";
}
async function apifyConfig(input: any, authorization: string) {
  await _requireCredentialAdmin(authorization);
  const op = String(input?.op || "status");
  const id = "apify_creative_miner_token";
  if (op === "status") {
    const token = await _loadSecureCredential(id);
    if (!token) return { configured: false };
    try {
      const r = await fetch("https://api.apify.com/v2/users/me", { headers: { Authorization: `Bearer ${token}` } });
      const j = await r.json();
      return { configured: r.ok, valid: r.ok, username: j?.data?.username || j?.data?.email || "", error: r.ok ? "" : (j?.error?.message || `HTTP ${r.status}`) };
    } catch (e) { return { configured: true, valid: false, error: String((e as any)?.message || e) }; }
  }
  if (op === "remove") {
    await fetch(`${_SB_URL}/rest/v1/secure_credentials?id=eq.${id}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
    return { ok: true, configured: false };
  }
  const token = String(input?.token || "").trim();
  if (token.length < 20) throw new Error("Token do Apify inválido ou incompleto.");
  const test = await fetch("https://api.apify.com/v2/users/me", { headers: { Authorization: `Bearer ${token}` } });
  const tj = await test.json();
  if (!test.ok || tj?.error) throw new Error(`Apify: ${tj?.error?.message || "token não autorizado"}`);
  await _saveSecureCredential(id, token);
  return { ok: true, configured: true, username: tj?.data?.username || tj?.data?.email || "" };
}
async function sbDeleteD(table: string, query: string) {
  const r = await fetch(`${_SB_URL}/rest/v1/${table}?${query}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
  if (!r.ok) throw new Error(`DB delete ${table}: ${await r.text()}`);
}

// ===== YouTube de eventos =====================================================
// Cada cliente tem o proprio refresh token no cofre. As consultas abaixo usam
// somente escopos de leitura e sempre filtram um unico video do canal conectado.
async function _youtubeClientToken(clientId: string): Promise<string> {
  const refresh = await _loadSecureCredential(`youtube_refresh_token:${clientId}`);
  const cid = Deno.env.get("GOOGLE_OAUTH_CLIENT_ID") || "";
  const secret = Deno.env.get("GOOGLE_OAUTH_CLIENT_SECRET") || "";
  if (!refresh) throw new Error("Conecte o YouTube deste cliente antes de abrir o Evento ao Vivo.");
  if (!cid || !secret) throw new Error("OAuth do Google ainda nao esta configurado no servidor.");
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: cid, client_secret: secret, refresh_token: refresh, grant_type: "refresh_token" }),
  });
  const j = await r.json();
  if (!r.ok || !j.access_token) throw new Error(`YouTube: ${j.error_description || j.error || "nao foi possivel renovar a conexao"}`);
  return j.access_token;
}
async function _ytJson(url: string, token: string) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const j = await r.json();
  if (!r.ok || j?.error) throw new Error(`YouTube: ${j?.error?.message || `HTTP ${r.status}`}`);
  return j;
}
function _ytRows(j: any) {
  const heads = (j?.columnHeaders || []).map((x: any) => x.name);
  return (j?.rows || []).map((row: any[]) => Object.fromEntries(heads.map((h: string, i: number) => [h, row[i]])));
}
async function _ytReport(token: string, input: { since: string; until: string; metrics: string; dimensions?: string; filters?: string; sort?: string; maxResults?: number }) {
  const q: any = { ids: "channel==MINE", startDate: input.since, endDate: input.until, metrics: input.metrics };
  if (input.dimensions) q.dimensions = input.dimensions;
  if (input.filters) q.filters = input.filters;
  if (input.sort) q.sort = input.sort;
  if (input.maxResults) q.maxResults = String(input.maxResults);
  return _ytRows(await _ytJson(`https://youtubeanalytics.googleapis.com/v2/reports?${new URLSearchParams(q)}`, token));
}
function _ytIsoSeconds(s: string) {
  const m = String(s || "").match(/^P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?$/);
  return m ? (Number(m[1]) * 86400 + Number(m[2]) * 3600 + Number(m[3]) * 60 + Number(m[4])) : 0;
}
async function _youtubeVideos(clientId: string) {
  const token = await _youtubeClientToken(clientId);
  const ch = await _ytJson("https://www.googleapis.com/youtube/v3/channels?part=snippet,contentDetails&mine=true", token);
  const channel = ch.items?.[0];
  if (!channel) throw new Error("A conta conectada nao possui um canal do YouTube.");
  const uploads = channel.contentDetails?.relatedPlaylists?.uploads;
  if (!uploads) throw new Error("Nao consegui localizar os videos enviados deste canal.");
  const pl = await _ytJson(`https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails&playlistId=${encodeURIComponent(uploads)}&maxResults=50`, token);
  const ids = (pl.items || []).map((x: any) => x.contentDetails?.videoId).filter(Boolean);
  if (!ids.length) return { channel: { id: channel.id, title: channel.snippet?.title || "YouTube" }, videos: [] };
  const vd = await _ytJson(`https://www.googleapis.com/youtube/v3/videos?part=snippet,contentDetails,statistics,liveStreamingDetails&id=${encodeURIComponent(ids.join(","))}`, token);
  const videos = (vd.items || []).map((v: any) => ({
    id: v.id, title: v.snippet?.title || v.id, publishedAt: v.snippet?.publishedAt || "",
    thumbnail: v.snippet?.thumbnails?.high?.url || v.snippet?.thumbnails?.medium?.url || v.snippet?.thumbnails?.default?.url || "",
    durationSeconds: _ytIsoSeconds(v.contentDetails?.duration || ""), viewsPublic: Number(v.statistics?.viewCount) || 0,
    likesPublic: Number(v.statistics?.likeCount) || 0, commentsPublic: Number(v.statistics?.commentCount) || 0,
    actualStartTime: v.liveStreamingDetails?.actualStartTime || "", actualEndTime: v.liveStreamingDetails?.actualEndTime || "",
    scheduledStartTime: v.liveStreamingDetails?.scheduledStartTime || "", isLiveEvent: !!v.liveStreamingDetails,
    url: `https://www.youtube.com/watch?v=${v.id}`,
  })).sort((a: any, b: any) => String(b.actualStartTime || b.publishedAt).localeCompare(String(a.actualStartTime || a.publishedAt)));
  return { channel: { id: channel.id, title: channel.snippet?.title || "YouTube", thumbnail: channel.snippet?.thumbnails?.default?.url || "" }, videos };
}
async function _youtubeLiveReport(input: any) {
  const clientId = String(input.clientId || ""), videoId = String(input.videoId || "").trim();
  const since = String(input.since || ""), until = String(input.until || "");
  if (!clientId || !videoId || !since || !until) throw new Error("Cliente, video e periodo sao obrigatorios.");
  const token = await _youtubeClientToken(clientId), filter = `video==${videoId}`, warnings: string[] = [];
  const safe = async (label: string, fn: () => Promise<any[]>) => { try { return await fn(); } catch (e) { warnings.push(`${label}: ${String((e as any)?.message || e)}`); return []; } };
  const videoJ = await _ytJson(`https://www.googleapis.com/youtube/v3/videos?part=snippet,contentDetails,statistics,liveStreamingDetails&id=${encodeURIComponent(videoId)}`, token);
  const v = videoJ.items?.[0]; if (!v) throw new Error("Video nao encontrado no canal conectado.");
  const video = { id: v.id, title: v.snippet?.title || v.id, thumbnail: v.snippet?.thumbnails?.high?.url || v.snippet?.thumbnails?.medium?.url || "", durationSeconds: _ytIsoSeconds(v.contentDetails?.duration || ""), actualStartTime: v.liveStreamingDetails?.actualStartTime || "", actualEndTime: v.liveStreamingDetails?.actualEndTime || "", url: `https://www.youtube.com/watch?v=${v.id}` };
  const summaryMetrics = "views,estimatedMinutesWatched,averageViewDuration,averageViewPercentage,subscribersGained,subscribersLost,likes,comments,shares";
  const summaryRows = await safe("indicadores", () => _ytReport(token, { since, until, metrics: summaryMetrics, filters: filter }));
  const concurrentRows = await safe("audiencia simultanea", () => _ytReport(token, { since, until, metrics: "averageConcurrentViewers,peakConcurrentViewers", filters: filter }));
  const daily = await safe("evolucao diaria", () => _ytReport(token, { since, until, metrics: summaryMetrics, dimensions: "day", filters: filter, sort: "day" }));
  const traffic = await safe("fontes de trafego", () => _ytReport(token, { since, until, metrics: "views,estimatedMinutesWatched", dimensions: "insightTrafficSourceType", filters: filter, sort: "-views", maxResults: 25 }));
  const external = await safe("sites externos", () => _ytReport(token, { since, until, metrics: "views,estimatedMinutesWatched", dimensions: "insightTrafficSourceDetail", filters: `${filter};insightTrafficSourceType==EXT_URL`, sort: "-views", maxResults: 25 }));
  const search = await safe("termos de busca", () => _ytReport(token, { since, until, metrics: "views,estimatedMinutesWatched", dimensions: "insightTrafficSourceDetail", filters: `${filter};insightTrafficSourceType==YT_SEARCH`, sort: "-views", maxResults: 25 }));
  const devices = await safe("dispositivos", () => _ytReport(token, { since, until, metrics: "views,estimatedMinutesWatched", dimensions: "deviceType", filters: filter, sort: "-estimatedMinutesWatched" }));
  const retention = await safe("retencao", () => _ytReport(token, { since, until, metrics: "audienceWatchRatio,relativeRetentionPerformance", dimensions: "elapsedVideoTimeRatio", filters: filter }));
  // comportamento da audiencia (casual / novo / recorrente) — dimensao propria do YouTube Analytics
  const audience = await safe("comportamento da audiencia", () => _ytReport(token, { since, until, metrics: "views,estimatedMinutesWatched", dimensions: "audienceType", filters: filter, sort: "-views" }));
  const summary: any = { ...(summaryRows[0] || {}), ...(concurrentRows[0] || {}) };
  summary.reactions = (Number(summary.likes) || 0) + (Number(summary.comments) || 0) + (Number(summary.shares) || 0);
  summary.watchHours = (Number(summary.estimatedMinutesWatched) || 0) / 60;
  const at30 = retention.reduce((best: any, x: any) => Math.abs(Number(x.elapsedVideoTimeRatio) - .3) < Math.abs(Number(best?.elapsedVideoTimeRatio ?? 99) - .3) ? x : best, null);
  summary.retention30 = at30 ? Number(at30.audienceWatchRatio) * 100 : null;
  const compare: any[] = [];
  for (const date of [input.compareDate1, input.compareDate2].filter(Boolean)) {
    const a = await safe(`comparativo ${date}`, () => _ytReport(token, { since, until: String(date), metrics: summaryMetrics, filters: filter }));
    const c = await safe(`simultaneos ${date}`, () => _ytReport(token, { since, until: String(date), metrics: "averageConcurrentViewers,peakConcurrentViewers", filters: filter }));
    const row: any = { date, ...(a[0] || {}), ...(c[0] || {}) };
    row.reactions = (Number(row.likes) || 0) + (Number(row.comments) || 0) + (Number(row.shares) || 0); row.watchHours = (Number(row.estimatedMinutesWatched) || 0) / 60;
    compare.push(row);
  }
  return { video, period: { since, until }, summary, daily, traffic, external, search, devices, retention, audience, compare, warnings, unavailable: { impressions: "A consulta imediata do YouTube Analytics nao fornece impressoes de thumbnail; esse dado exige o fluxo de relatorios em lote.", liveChatMessages: "A API oficial nao disponibiliza o historico completo do chat depois que a live termina.", uniqueViewers: "Pode nao ser disponibilizado para todas as contas e combinacoes de dimensoes." } };
}
function _minerPick(row: any, keys: string[]) { for (const k of keys) if (row?.[k] != null && row[k] !== "") return row[k]; return null; }
// Ranqueia os candidatos capturados pelo critério escolhido na tela, e devolve o texto do disclaimer
// que aparece no card — sem isso a pessoa não sabe POR QUE aquele post específico entrou na biblioteca.
function _minerScore(criterio: string, m: any): number {
  const likes = Number(m.likes) || 0, comments = Number(m.comments) || 0, views = Number(m.views) || 0, reach = Number(m.reach) || 0, saved = Number(m.saved) || 0, shares = Number(m.shares) || 0;
  if (criterio === "curtidas") return likes;
  if (criterio === "comentarios") return comments;
  if (criterio === "visualizacoes") return views || likes;
  if (m.eng != null) return Number(m.eng) || 0;
  const base = reach || views || likes || 1;
  return ((likes + comments + saved + shares) / base) * 100;
}
function _minerReason(criterio: string, m: any, score: number): string {
  const num = (n: number) => Math.round(n || 0).toLocaleString("pt-BR");
  if (criterio === "recentes") return `Selecionado por ser um dos mais recentes${m.data || m.timestamp ? ` (${new Date(m.data || m.timestamp).toLocaleDateString("pt-BR")})` : ""}.`;
  if (criterio === "curtidas") return `Selecionado por mais curtidas (${num(m.likes)}).`;
  if (criterio === "comentarios") return `Selecionado por mais comentários (${num(m.comments)}).`;
  if (criterio === "visualizacoes") return `Selecionado por mais visualizações (${num(m.views || m.likes)}).`;
  return `Selecionado por melhor engajamento (${score.toFixed(1).replace(".", ",")}% sobre ${m.reach ? "alcance" : "visualizações/curtidas"}).`;
}
async function _minerAnalyzePayload(item: any, client: any, funilDesejado?: string) {
  const funilLinha = funilDesejado ? `\nO cliente está minerando pensando em conteúdo de ${funilDesejado} de funil — ao montar "ideias_estaticas", priorize ideias que sirvam pra essa etapa e deixe isso explícito no campo "por_que" de cada uma.` : "";
  const instruction = `Você é o núcleo semântico do Minerador de Criativos. Analise esta referência para ${client?.name || "o cliente"}. DNA: ${JSON.stringify(client?.dna || {}).slice(0, 8000)}. Legenda: ${String(item.caption || "").slice(0, 10000)}.${funilLinha}
REGRA CENTRAL: entenda tudo o que o conteúdo FALA/ENSINA, e não apenas seus frames. Se ensina "10 tipos de criativos", extraia os dez tipos e suas explicações; não transforme o tema em "10 vídeos". Diferencie MEIO ORIGINAL, ASSUNTO REAL, CONHECIMENTO TRANSMITIDO e FORMA CRIATIVA. Frames são somente evidência visual. Gere aplicações originais para o cliente, sem copiar texto, marca ou identidade de terceiros.
Retorne somente JSON válido: {"assunto_real":"","tese_central":"","resumo_do_conteudo":"","itens_mencionados":[{"titulo":"","explicacao":"","exemplo":"","aplicacao_cliente":""}],"hook":{"texto":"","tipo":"","por_que_funciona":""},"promessa":"","estrutura_narrativa":[{"momento":"","funcao":"","conteudo":""}],"textos_na_tela":[],"prova_social":"","cta":"","etapa_funil":"","forma_criativa":{"ritmo":"","estetica":"","estrutura":""},"ideias_estaticas":[{"titulo":"","headline":"","conceito_visual":"","formato":"imagem|carrossel","funil":"","por_que":""}],"alertas_direitos":[]}`;
  const gem = Deno.env.get("GEMINI_API_KEY") || "";
  if (item.media_type === "video" && item.media_url && gem) {
    const vr = await fetch(String(item.media_url));
    if (vr.ok) {
      const bytes = new Uint8Array(await vr.arrayBuffer());
      if (bytes.length <= 24 * 1024 * 1024) {
        const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${gem}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ contents: [{ parts: [{ text: instruction }, { inline_data: { mime_type: vr.headers.get("content-type") || "video/mp4", data: _b64Bytes(bytes) } }] }], generationConfig: { responseMimeType: "application/json", temperature: 0.25, maxOutputTokens: 5000 } }) });
        const j = await r.json();
        const txt = j?.candidates?.[0]?.content?.parts?.map((p: any) => p.text || "").join("") || "";
        if (r.ok && txt) return JSON.parse(txt.replace(/^```json\s*|\s*```$/g, ""));
      }
    }
  }
  const content: any[] = [{ type: "text", text: instruction }];
  if (item.thumbnail_url || item.media_url) {
    try {
      const ir = await fetch(String(item.thumbnail_url || item.media_url), { headers: { "User-Agent": "Mozilla/5.0" } });
      if (ir.ok) {
        const b = new Uint8Array(await ir.arrayBuffer());
        if (b.length <= 8 * 1024 * 1024) content.push({ type: "image_url", image_url: { url: `data:${ir.headers.get("content-type") || "image/jpeg"};base64,${_b64Bytes(b)}` } });
      }
    } catch (_e) { /* legenda ainda permite análise parcial */ }
  }
  const j = await callOpenAI({ model: "gpt-4o", messages: [{ role: "user", content }], response_format: { type: "json_object" }, max_tokens: 2600, temperature: 0.3 });
  return JSON.parse(j.choices?.[0]?.message?.content || "{}");
}
async function _minerApifyCapture(url: string, rawLimit: number, sinceISO?: string) {
  const token = await _loadSecureCredential("apify_creative_miner_token");
  if (!token) throw new Error("Configure o token do Apify em Configurações → Integrações.");
  const endpoint = `https://api.apify.com/v2/acts/apify~instagram-scraper/run-sync-get-dataset-items?token=${encodeURIComponent(token)}&timeout=120&memory=1024`;
  const input: any = { directUrls: [url], resultsType: "posts", resultsLimit: Math.min(50, Math.max(1, rawLimit || 30)), searchType: "user" };
  if (sinceISO) input.onlyPostsNewerThan = sinceISO; // aceito pelo ator apify~instagram-scraper (data ou "1 month")
  const r = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input) });
  const rows = await r.json();
  if (!r.ok || !Array.isArray(rows)) throw new Error(`Apify: ${rows?.error?.message || `HTTP ${r.status}`}`);
  return rows;
}
// Perfis relacionados de verdade: pede ao Instagram (via Apify) os "perfis relacionados" que a própria
// plataforma sugere na página do perfil do cliente. Não inventa handle — se o Instagram não trouxer
// nada pra esse perfil agora, devolve lista vazia com aviso em vez de a IA chutar um nome.
async function _minerRelatedProfiles(username: string) {
  const token = await _loadSecureCredential("apify_creative_miner_token");
  if (!token) throw new Error("Configure o token do Apify em Configurações → Integrações.");
  const endpoint = `https://api.apify.com/v2/acts/apify~instagram-scraper/run-sync-get-dataset-items?token=${encodeURIComponent(token)}&timeout=90&memory=1024`;
  const input = { directUrls: [`https://www.instagram.com/${username}/`], resultsType: "details", resultsLimit: 1 };
  const r = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input) });
  const rows = await r.json();
  if (!r.ok || !Array.isArray(rows)) throw new Error(`Apify: ${rows?.error?.message || `HTTP ${r.status}`}`);
  const profile = rows[0] || {};
  const related = profile.relatedProfiles || profile.related_profiles || profile.chainingUsers || profile.suggestedUsers || [];
  return (Array.isArray(related) ? related : []).map((p: any) => {
    const uname = String(_minerPick(p, ["username", "value"]) || "").trim();
    return uname ? { username: uname, fullName: String(_minerPick(p, ["full_name", "fullName"]) || ""), private: !!(p.is_private ?? p.isPrivate), profilePic: String(_minerPick(p, ["profile_pic_url", "profilePicUrl"]) || ""), url: `https://www.instagram.com/${uname}/` } : null;
  }).filter(Boolean);
}
async function eventReports(input: any) {
  const op = String(input?.op || "list"), clientId = String(input?.clientId || "").trim();
  if (op === "youtubeVideos") {
    if (!clientId) throw new Error("Selecione um cliente.");
    return await _youtubeVideos(clientId);
  }
  if (op === "youtubeLive") {
    if (!clientId) throw new Error("Selecione um cliente.");
    return await _youtubeLiveReport(input);
  }
  if (op === "saveYoutubeSelection") {
    if (!clientId) throw new Error("Selecione um cliente.");
    const editionId = String(input.editionId || ""), videoId = String(input.videoId || "").trim();
    const edition = (await sbGet("event_editions", `id=eq.${encodeURIComponent(editionId)}&client_id=eq.${encodeURIComponent(clientId)}&select=config&limit=1`))[0];
    if (!edition) throw new Error("Evento nao encontrado para este cliente.");
    await sbPatchD("event_editions", `id=eq.${encodeURIComponent(editionId)}&client_id=eq.${encodeURIComponent(clientId)}`, { config: { ...(edition.config || {}), youtube_video_id: videoId, youtube_video_title: String(input.videoTitle || "").slice(0, 300) }, updated_at: new Date().toISOString() });
    return { ok: true };
  }
  if (op === "analyzeSection") {
    if (!clientId) throw new Error("Selecione um cliente.");
    const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,dna&limit=1`))[0];
    if (!c) throw new Error("Cliente nao encontrado.");
    const section = String(input.section || "secao do evento").slice(0, 100);
    const prompt = `Voce e AndreIA, consultora senior de eventos e YouTube. Analise SOMENTE os dados reais da secao \"${section}\" para o cliente ${c.name}. Diferencie fato, hipotese e dado indisponivel. Nao invente metricas e nao diga que publico ou nicho esta errado; proponha investigacoes quando houver mais de uma explicacao. Seja consultiva, direta e sem prazos. Retorne apenas JSON valido: {"resumo":"2 a 4 frases","destaques":["..."],"atencoes":[{"ponto":"...","evidencia":"...","investigacao":"..."}],"acoes":["passo pratico sem prazo"]}.\nDNA DO CLIENTE: ${JSON.stringify(c.dna || {}).slice(0, 7000)}\nDADOS DA SECAO: ${JSON.stringify(input.data || {}).slice(0, 30000)}`;
    return { analysis: await _callOpenAIJson([{ role: "user", content: prompt }]) };
  }
  if (op === "list") {
    const cq = clientId ? `client_id=eq.${encodeURIComponent(clientId)}&` : "";
    const projects = await sbGet("event_projects", `${cq}select=*&order=created_at.desc&limit=300`);
    const editions = await sbGet("event_editions", `${cq}select=*&order=event_date.desc&limit=500`);
    const editionIds = (editions || []).map((x: any) => x.id);
    let snapshots: any[] = [], versions: any[] = [];
    if (editionIds.length) {
      const ids = editionIds.map((x: string) => encodeURIComponent(x)).join(",");
      snapshots = await sbGet("event_snapshots", `edition_id=in.(${ids})&select=*&order=collected_at.desc&limit=1000`);
      versions = await sbGet("event_report_versions", `edition_id=in.(${ids})&select=*&order=version_no.desc&limit=1000`);
    }
    return { projects, editions, snapshots, versions };
  }
  if (!clientId) throw new Error("Selecione um cliente.");
  if (op === "create") {
    const name = String(input.name || "").trim(), editionName = String(input.editionName || "").trim();
    if (!name || !editionName) throw new Error("Informe o projeto e a edição.");
    const projectId = _wuid(), editionId = _wuid(), now = new Date().toISOString();
    await sbPost("event_projects", { id: projectId, client_id: clientId, name, description: String(input.description || ""), created_at: now, updated_at: now });
    await sbPost("event_editions", { id: editionId, project_id: projectId, client_id: clientId, name: editionName, year: Number(input.year) || null, capture_start: input.captureStart || null, capture_end: input.captureEnd || null, event_date: input.eventDate || null, sales_start: input.salesStart || null, sales_end: input.salesEnd || null, config: input.config || {}, created_at: now, updated_at: now });
    return { projectId, editionId };
  }
  if (op === "addEdition") {
    const projectId = String(input.projectId || ""), name = String(input.editionName || "").trim();
    if (!projectId || !name) throw new Error("Projeto e nome da edição são obrigatórios.");
    const id = _wuid(), now = new Date().toISOString();
    await sbPost("event_editions", { id, project_id: projectId, client_id: clientId, name, year: Number(input.year) || null, capture_start: input.captureStart || null, capture_end: input.captureEnd || null, event_date: input.eventDate || null, sales_start: input.salesStart || null, sales_end: input.salesEnd || null, config: input.config || {}, created_at: now, updated_at: now });
    return { editionId: id };
  }
  // Editar as datas da edição (captação, data do evento e JANELA DE VENDAS). Sem isso, uma edição criada com a
  // janela incompleta ficava sem conserto pela tela — e o card "Faturamento · janela do evento" nunca preenchia.
  if (op === "updateEdition") {
    const editionId = String(input.editionId || ""); if (!editionId) throw new Error("Edição obrigatória.");
    const campos: any = { updated_at: new Date().toISOString() };
    const mapa: Record<string, string> = { captureStart: "capture_start", captureEnd: "capture_end", eventDate: "event_date", salesStart: "sales_start", salesEnd: "sales_end" };
    for (const [de, para] of Object.entries(mapa)) if (de in input) campos[para] = input[de] || null;
    if (input.editionName) campos.name = String(input.editionName).trim();
    if (input.year != null) campos.year = Number(input.year) || null;
    const r = await sbPatchD("event_editions", `id=eq.${encodeURIComponent(editionId)}&client_id=eq.${encodeURIComponent(clientId)}`, campos);
    return { ok: true, editionId, atualizado: campos, _r: r || null };
  }
  if (op === "filterOptions") {
    const since = String(input.since || ""), until = String(input.until || ""); if (!since || !until) throw new Error("Período obrigatório.");
    const media = await _sbAll("channel_metrics_daily", `client_id=eq.${encodeURIComponent(clientId)}&date=gte.${since}&date=lte.${until}&select=channel,campaign,source_medium`);
    const rd = await _sbAll("rd_conversions", `client_id=eq.${encodeURIComponent(clientId)}&converted_at=gte.${since}T00:00:00Z&converted_at=lte.${until}T23:59:59Z&select=event_identifier`);
    // Campanha é filtro de MÍDIA. GA4 é a fonte independente de conversões/receita e
    // nunca deve aparecer aqui nem ser removido quando uma campanha paga é selecionada.
    const paidMedia = (media || []).filter((x: any) => String(x.channel || "").toLowerCase() !== "ga4");
    const campaigns = [...new Map(paidMedia.map((x: any) => { const name = String(x.campaign || "").trim(); return [`${x.channel}|${name}`, { key: `${x.channel}|${name}`, channel: x.channel, name }]; }).filter((x: any) => x[1].name)).values()].sort((a: any, b: any) => a.name.localeCompare(b.name));
    const rdEvents = [...new Set((rd || []).map((x: any) => String(x.event_identifier || "").trim()).filter(Boolean))].sort();
    return { campaigns, rdEvents, ga4Available: (media || []).some((x: any) => String(x.channel || "").toLowerCase() === "ga4") };
  }
  if (op === "snapshot") {
    const editionId = String(input.editionId || ""), since = String(input.since || ""), until = String(input.until || "");
    if (!editionId || !since || !until) throw new Error("Edição e período são obrigatórios.");
    let rows = await _sbAll("channel_metrics_daily", `client_id=eq.${encodeURIComponent(clientId)}&date=gte.${since}&date=lte.${until}&select=channel,source_medium,campaign,adset,ad_content,spend,impressions,clicks,reach,purchases,revenue,leads,conversas,video_views,engajamentos`);
    const selectedChannels = Array.isArray(input.channels) ? input.channels.map((x: any) => String(x).toLowerCase()) : [];
    if (selectedChannels.length) rows = rows.filter((r: any) => String(r.channel || "").toLowerCase() === "ga4" || selectedChannels.includes(String(r.channel || "").toLowerCase()));
    const selectedCampaigns = Array.isArray(input.campaigns) ? input.campaigns.map(String) : [];
    if (selectedCampaigns.length) { const set = new Set(selectedCampaigns); rows = rows.filter((r: any) => String(r.channel || "").toLowerCase() === "ga4" || set.has(`${r.channel}|${String(r.campaign || "").trim()}`)); }
    const empty = () => ({ spend: 0, impressions: 0, clicks: 0, reach: 0, purchases: 0, revenue: 0, leads: 0, conversas: 0, video_views: 0, engajamentos: 0 });
    const total: any = empty(), channels: Record<string, any> = {}, campaignMap: Record<string, any> = {};
    for (const r of rows || []) { const c = channels[r.channel] || (channels[r.channel] = empty()); for (const k of Object.keys(total)) { const v = Number(r[k]) || 0; c[k] += v; total[k] += v; }
      const ck = `${r.channel}|${r.campaign || r.source_medium || "(sem campanha)"}`; const cp = campaignMap[ck] || (campaignMap[ck] = { channel: r.channel, campaign: r.campaign || "", source_medium: r.source_medium || "", spend: 0, impressions: 0, reach: 0, clicks: 0, purchases: 0, revenue: 0, leads: 0, conversas: 0, video_views: 0 });
      for (const k of ["spend", "impressions", "reach", "clicks", "purchases", "revenue", "leads", "conversas", "video_views"]) cp[k] += Number(r[k]) || 0;
    }
    // Evita somar a mesma venda no gerenciador e no Analytics. Mídia fornece entrega/custo;
    // quando GA4 existe, ele é a fonte de verdade para vendas e faturamento do período.
    const paidTotal: any = empty();
    for (const [key, c] of Object.entries(channels)) if (String(key).toLowerCase() !== "ga4") for (const k of Object.keys(paidTotal)) paidTotal[k] += Number((c as any)[k]) || 0;
    const ga4 = channels.ga4;
    total.spend = paidTotal.spend; total.impressions = paidTotal.impressions; total.clicks = paidTotal.clicks; total.reach = paidTotal.reach;
    total.leads = paidTotal.leads; total.conversas = paidTotal.conversas; total.video_views = paidTotal.video_views; total.engajamentos = paidTotal.engajamentos;
    if (ga4 && (Number(ga4.purchases) || Number(ga4.revenue))) { total.purchases = Number(ga4.purchases) || 0; total.revenue = Number(ga4.revenue) || 0; }
    total.ctr = total.impressions ? total.clicks / total.impressions * 100 : 0; total.cpm = total.impressions ? total.spend / total.impressions * 1000 : 0; total.roas = total.spend ? total.revenue / total.spend : 0;
    Object.values(channels).forEach((c: any) => { c.ctr = c.impressions ? c.clicks / c.impressions * 100 : 0; c.cpm = c.impressions ? c.spend / c.impressions * 1000 : 0; c.roas = c.spend ? c.revenue / c.spend : 0; });
    let rd = await _sbAll("rd_conversions", `client_id=eq.${encodeURIComponent(clientId)}&converted_at=gte.${since}T00:00:00Z&converted_at=lte.${until}T23:59:59Z&select=event_identifier,source,medium,campaign`);
    const selectedRd = Array.isArray(input.rdEvents) ? input.rdEvents.map(String) : [];
    if (selectedRd.length) { const set = new Set(selectedRd); rd = rd.filter((r: any) => set.has(String(r.event_identifier || ""))); }
    const rdEvents: Record<string, number> = {}; for (const x of rd || []) { const k = String(x.event_identifier || "Conversão"); rdEvents[k] = (rdEvents[k] || 0) + 1; }
    const convs = await _sbAll("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&created_at=gte.${since}T00:00:00Z&created_at=lte.${until}T23:59:59Z&select=stage,origin_type`);
    const crmStages: Record<string, number> = {}; for (const x of convs || []) { const k = String(x.stage || "sem_etapa"); crmStages[k] = (crmStages[k] || 0) + 1; }
    const salesTp = await _sbAll("lead_touchpoints", `client_id=eq.${encodeURIComponent(clientId)}&kind=eq.purchase&ts=gte.${since}T00:00:00Z&ts=lte.${until}T23:59:59Z&select=value,channel,campaign,label`);
    const commerce = { purchases: (salesTp || []).length, revenue: (salesTp || []).reduce((a: number, x: any) => a + (Number(x.value) || 0), 0) };
    const campaigns = Object.values(campaignMap).map((x: any) => ({ ...x, ctr: x.impressions ? x.clicks / x.impressions * 100 : 0, cpa: (x.purchases || x.leads || x.conversas) ? x.spend / (x.purchases || x.leads || x.conversas) : 0 })).sort((a: any, b: any) => b.spend - a.spend).slice(0, 100);
    const id = _wuid(), collectedAt = new Date().toISOString(), metrics = { total, channels, campaigns, rd: { total: (rd || []).length, events: rdEvents }, crm: { total: (convs || []).length, stages: crmStages }, commerce };
    await sbPost("event_snapshots", { id, edition_id: editionId, client_id: clientId, period_start: since, period_end: until, metrics, sources: { channel_metrics_daily: true, journey_sales: !!salesTp.length, rd: !!rd.length, crm: !!convs.length, youtube: false, filters: { channels: selectedChannels, campaigns: selectedCampaigns, rd_events: selectedRd, since, until } }, collected_at: collectedAt });
    return { snapshot: { id, edition_id: editionId, client_id: clientId, period_start: since, period_end: until, metrics, collected_at: collectedAt } };
  }
  if (op === "saveVersion") {
    const editionId = String(input.editionId || ""); if (!editionId) throw new Error("Edição obrigatória.");
    const prev = await sbGet("event_report_versions", `edition_id=eq.${encodeURIComponent(editionId)}&select=version_no&order=version_no.desc&limit=1`);
    const versionNo = (Number(prev?.[0]?.version_no) || 0) + 1, id = _wuid();
    await sbPost("event_report_versions", { id, edition_id: editionId, snapshot_id: input.snapshotId || null, version_no: versionNo, status: input.status || "draft", title: String(input.title || `Relatório v${versionNo}`), instructions: String(input.instructions || ""), report_data: input.reportData || {}, created_by: input.createdBy || null });
    return { id, versionNo };
  }
  if (op === "analyze") {
    const editionId = String(input.editionId || ""), snapshotId = String(input.snapshotId || "");
    const edition = (await sbGet("event_editions", `id=eq.${encodeURIComponent(editionId)}&client_id=eq.${encodeURIComponent(clientId)}&select=*&limit=1`))[0];
    const snapshot = (await sbGet("event_snapshots", `id=eq.${encodeURIComponent(snapshotId)}&edition_id=eq.${encodeURIComponent(editionId)}&select=metrics,sources,period_start,period_end&limit=1`))[0];
    if (!edition || !snapshot) throw new Error("Atualize os dados antes de gerar a análise.");
    const prompt = `Você é AndréIA, consultora sênior de eventos de marketing. Analise somente os dados reais abaixo. Diferencie fatos, hipóteses e dados ausentes. Não diga que público/nicho está errado; quando houver dúvida, proponha investigação. Produza JSON em português: {"resumo":"3-5 frases","destaques":["..."],"gargalos":[{"ponto":"...","evidencia":"...","acao":"..."}],"canais":[{"canal":"...","leitura":"..."}],"criativos_campanhas":[{"nome":"...","leitura":"..."}],"projecoes":["..."],"proximos_passos":["..."]}. Não inclua prazos.\nPERÍODO: ${snapshot.period_start} a ${snapshot.period_end}\nORIENTAÇÃO DO GESTOR: ${String(input.instructions || "Sem orientação adicional").slice(0, 3000)}\nDADOS: ${JSON.stringify(snapshot.metrics).slice(0, 45000)}\nFONTES: ${JSON.stringify(snapshot.sources)}`;
    return { analysis: await _callOpenAIJson([{ role: "user", content: prompt }]) };
  }
  throw new Error("Operação de evento inválida.");
}

type AccessActor = { user: any; profile: any; aal: string };
function _jwtPayload(authorization: string) {
  try {
    const token = String(authorization || "").replace(/^Bearer\s+/i, "");
    const raw = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    return JSON.parse(atob(raw.padEnd(Math.ceil(raw.length / 4) * 4, "=")));
  } catch (_e) { return {}; }
}
async function _accessActor(authorization: string, requireAal2 = false): Promise<AccessActor> {
  if (!authorization) throw new Error("Sessão obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } });
  const user = ur.ok ? await ur.json() : null;
  if (!user?.id) throw new Error("Sessão inválida ou expirada.");
  const rows = await sbGet("app_users", `user_id=eq.${encodeURIComponent(user.id)}&select=*&limit=1`);
  const profile = rows[0];
  if (!profile) throw new Error("Este login ainda não foi cadastrado em Usuários e acessos.");
  if (profile.status === "inactive") throw new Error("Este usuário está desativado. Fale com o usuário master.");
  const aal = String(_jwtPayload(authorization)?.aal || "aal1");
  if (requireAal2 && profile.mfa_required !== false && aal !== "aal2") throw new Error("Confirme o código de autenticação em duas etapas para continuar.");
  return { user, profile, aal };
}
function _isMaster(a: AccessActor) { return a.profile?.role === "master" && a.profile?.protected === true; }
async function _auditAccess(actorId: string, targetId: string | null, action: string, before: any, after: any) {
  await fetch(`${_SB_URL}/rest/v1/app_access_audit`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: JSON.stringify({ actor_user_id: actorId, target_user_id: targetId, action, before_data: before || null, after_data: after || null }) });
}
async function _replaceUserClients(userId: string, clientIds: string[]) {
  await fetch(`${_SB_URL}/rest/v1/app_user_clients?user_id=eq.${encodeURIComponent(userId)}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
  const rows = [...new Set((clientIds || []).map(String).filter(Boolean))].map(client_id => ({ user_id: userId, client_id }));
  if (rows.length) {
    const r = await fetch(`${_SB_URL}/rest/v1/app_user_clients`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" }, body: JSON.stringify(rows) });
    if (!r.ok) throw new Error("Não consegui salvar os clientes permitidos.");
  }
}
function _safePermissions(v: any) {
  const menus = Array.isArray(v?.menus) ? [...new Set(v.menus.map(String))].slice(0, 100) : [];
  const actions = Array.isArray(v?.actions) ? [...new Set(v.actions.map(String))].slice(0, 100) : [];
  return { menus, actions };
}
async function accessControl(input: any, authorization: string) {
  const op = String(input?.op || "me");
  const actor = await _accessActor(authorization, op !== "me");
  if (op === "me") {
    const clients = await sbGet("app_user_clients", `user_id=eq.${encodeURIComponent(actor.user.id)}&select=client_id`);
    return { profile: actor.profile, clientIds: clients.map((x: any) => x.client_id), aal: actor.aal };
  }
  if (op === "activate_self") {
    if (actor.aal !== "aal2") throw new Error("Conclua o 2FA antes de ativar o acesso.");
    if (actor.profile.status === "invited") {
      await fetch(`${_SB_URL}/rest/v1/app_users?user_id=eq.${encodeURIComponent(actor.user.id)}`, { method: "PATCH", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: JSON.stringify({ status: "active", updated_at: new Date().toISOString() }) });
    }
    return { ok: true };
  }
  const canAdminUsers = _isMaster(actor) || actor.profile?.role === "admin";
  if (!canAdminUsers) throw new Error("Somente master ou administrador podem administrar usuários.");
  if (op === "list") {
    const users = await sbGet("app_users", "select=*&order=created_at.asc");
    const links = await sbGet("app_user_clients", "select=user_id,client_id");
    const ar = await fetch(`${_SB_URL}/auth/v1/admin/users?per_page=1000`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
    const authUsers = ar.ok ? ((await ar.json()).users || []) : [];
    return { users: users.map((u: any) => ({ ...u, clientIds: links.filter((x: any) => x.user_id === u.user_id).map((x: any) => x.client_id), mfaEnrolled: !!authUsers.find((x: any) => x.id === u.user_id)?.factors?.some((f: any) => f.status === "verified") })) };
  }
  if (op === "invite") {
    const email = String(input.email || "").trim().toLowerCase(), name = String(input.name || "").trim();
    if (!/^\S+@\S+\.\S+$/.test(email)) throw new Error("Informe um e-mail válido.");
    const role = _isMaster(actor) && input.role === "admin" ? "admin" : "gestor";
    const redirect = "https://app.gt-marketing.app.br";
    const ir = await fetch(`${_SB_URL}/auth/v1/invite?redirect_to=${encodeURIComponent(redirect)}`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: JSON.stringify({ email, data: { name } }) });
    const ij = await ir.json();
    if (!ir.ok && !/already|registered|exists/i.test(String(ij?.msg || ij?.message || ""))) throw new Error(ij?.msg || ij?.message || "Não consegui enviar o convite.");
    let userId = ij?.id;
    if (!userId) {
      const lr = await fetch(`${_SB_URL}/auth/v1/admin/users?per_page=1000`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
      const list = lr.ok ? ((await lr.json()).users || []) : [];
      userId = list.find((x: any) => String(x.email || "").toLowerCase() === email)?.id;
    }
    if (!userId) throw new Error("O login foi convidado, mas ainda não consegui vinculá-lo ao perfil. Tente atualizar em instantes.");
    const row = { user_id: userId, email, name: name || email.split("@")[0], role, status: "invited", all_clients: !!input.allClients, permissions: _safePermissions(input.permissions), mfa_required: true, protected: false, created_by: actor.user.id, updated_at: new Date().toISOString() };
    const sr = await fetch(`${_SB_URL}/rest/v1/app_users?on_conflict=user_id`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=representation" }, body: JSON.stringify(row) });
    if (!sr.ok) throw new Error("Convite enviado, mas não consegui salvar o perfil de acesso.");
    await _replaceUserClients(userId, input.clientIds || []);
    await _auditAccess(actor.user.id, userId, "user_invited", null, row);
    return { ok: true, invited: ir.ok, user: (await sr.json())[0] };
  }
  const userId = String(input.userId || "");
  const current = (await sbGet("app_users", `user_id=eq.${encodeURIComponent(userId)}&select=*&limit=1`))[0];
  if (!current) throw new Error("Usuário não encontrado.");
  if (current.protected) throw new Error("O usuário master protegido não pode ser alterado ou excluído.");
  if (!_isMaster(actor) && current.role !== "gestor") throw new Error("Administradores só podem gerenciar usuários gestores.");
  if (op === "reset_mfa") {
    const fr = await fetch(`${_SB_URL}/auth/v1/admin/users/${encodeURIComponent(userId)}/factors`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
    const fj = fr.ok ? await fr.json() : [];
    const factors = Array.isArray(fj) ? fj : (fj?.factors || []);
    let removed = 0;
    for (const f of factors) {
      const dr = await fetch(`${_SB_URL}/auth/v1/admin/users/${encodeURIComponent(userId)}/factors/${encodeURIComponent(f.id)}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
      if (dr.ok) removed++;
    }
    await _auditAccess(actor.user.id, userId, "mfa_reset", { factors: factors.length }, { removed });
    return { ok: true, removed };
  }
  if (op === "update") {
    const role = _isMaster(actor) && input.role === "admin" ? "admin" : "gestor";
    const status = ["active", "inactive", "invited"].includes(input.status) ? input.status : current.status;
    const patch = { name: String(input.name || current.name).trim(), role, status, all_clients: !!input.allClients, permissions: _safePermissions(input.permissions), mfa_required: true, updated_at: new Date().toISOString() };
    const r = await fetch(`${_SB_URL}/rest/v1/app_users?user_id=eq.${encodeURIComponent(userId)}`, { method: "PATCH", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "return=representation" }, body: JSON.stringify(patch) });
    if (!r.ok) throw new Error("Não consegui atualizar o usuário.");
    await _replaceUserClients(userId, input.clientIds || []);
    await _auditAccess(actor.user.id, userId, "user_updated", current, { ...current, ...patch });
    return { ok: true, user: (await r.json())[0] };
  }
  if (op === "deactivate") {
    await fetch(`${_SB_URL}/rest/v1/app_users?user_id=eq.${encodeURIComponent(userId)}`, { method: "PATCH", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: JSON.stringify({ status: "inactive", updated_at: new Date().toISOString() }) });
    await _auditAccess(actor.user.id, userId, "user_deactivated", current, { ...current, status: "inactive" });
    return { ok: true };
  }
  throw new Error("Operação de acesso não reconhecida.");
}

function _sanitizedAccountConfig(data: any) {
  return { agency_client_active: !!data?.agency_client_active, crm_show_agency: !!data?.crm_show_agency, bd_views: data?.bd_views || [], andreia_wa: data?.andreia_wa ? { group_jid: data.andreia_wa.group_jid || "", group_name: data.andreia_wa.group_name || "", allowed: data.andreia_wa.allowed || [] } : {} };
}
async function accountConfigAccess(input: any, authorization: string) {
  const actor = await _accessActor(authorization, true);
  const rows = await sbGet("account_config", "id=eq.main&select=data&limit=1"), current = rows[0]?.data || {};
  if (String(input?.op || "read") === "read") return { data: _isMaster(actor) ? current : _sanitizedAccountConfig(current), full: _isMaster(actor) };
  const patch = input?.patch && typeof input.patch === "object" ? input.patch : {};
  const allowed = _isMaster(actor) ? Object.keys(patch) : ["bd_views"].filter(k => Object.prototype.hasOwnProperty.call(patch, k));
  if (!allowed.length) throw new Error("Você não tem permissão para alterar esta configuração global.");
  const clean: any = {}; allowed.forEach(k => clean[k] = patch[k]);
  const next = { ...current, ...clean };
  const r = await fetch(`${_SB_URL}/rest/v1/account_config?on_conflict=id`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify({ id: "main", data: next, updated_at: new Date().toISOString() }) });
  if (!r.ok) throw new Error("Não consegui salvar a configuração.");
  await _auditAccess(actor.user.id, actor.user.id, "account_config_updated", null, { keys: allowed });
  return { ok: true, data: _isMaster(actor) ? next : _sanitizedAccountConfig(next) };
}
function _accessArray(profile: any, kind: "menus" | "actions") { const v = profile?.permissions?.[kind]; return Array.isArray(v) ? v.map(String) : []; }
function _actorHas(actor: AccessActor, kind: "menus" | "actions", code: string) { if (["master", "admin"].includes(actor.profile?.role)) return true; const a = _accessArray(actor.profile, kind); return a.includes("*") || a.includes(code); }
async function _actorClientIds(actor: AccessActor) { if (["master", "admin"].includes(actor.profile?.role) || actor.profile?.all_clients) return null; return (await sbGet("app_user_clients", `user_id=eq.${encodeURIComponent(actor.user.id)}&select=client_id`)).map((x: any) => String(x.client_id)); }
function _bodyClientIds(v: any, out = new Set<string>(), depth = 0) {
  if (!v || typeof v !== "object" || depth > 3) return out;
  for (const [k, x] of Object.entries(v)) {
    if ((k === "clientId" || k === "client_id") && typeof x === "string" && x) out.add(x);
    else if (typeof x === "object") _bodyClientIds(x, out, depth + 1);
  }
  return out;
}
function _bodyAccountIds(v: any) {
  const out = new Set<string>();
  const walk = (x: any, depth = 0) => { if (x == null || depth > 3) return; if (Array.isArray(x)) { x.forEach(y => walk(y, depth + 1)); return; } if (typeof x !== "object") return; for (const [k, y] of Object.entries(x)) { if (/^(accountId|account_id)$/.test(k) && y) out.add(String(y).replace(/\D/g, "")); else if (k === "accounts" && Array.isArray(y)) y.forEach((a: any) => out.add(String(a?.id || a).replace(/\D/g, ""))); else if (typeof y === "object") walk(y, depth + 1); } };
  walk(v); return [...out].filter(Boolean);
}
function _firstBodyKey(body: any) { return Object.keys(body || {}).find(k => body[k] !== undefined) || ""; }
function _menuForOperation(key: string) {
  if (/^crm|^wa/i.test(key)) return "crm";
  if (/^briefing|creativeMiner/i.test(key)) return "briefing";
  if (/^event/i.test(key)) return "eventos";
  if (/^journey|ga4|gsc/i.test(key)) return "jornada";
  if (/^meta|^google|^tiktok|^pinterest|channelMetrics/i.test(key)) return /Audience|CreateCustom/i.test(key) ? "publicos" : "campanhas";
  if (/^agent$/i.test(key)) return "iagestora";
  return "";
}
async function _guardUserRequest(body: any, authorization: string) {
  const jwt = _jwtPayload(authorization);
  if (jwt?.role === "service_role") return;
  // Chave nova do Supabase (sb_secret_...) nao e JWT — o decode acima da {} e a chamada INTERNA (crons via
  // tracking.ts) caia na validacao de sessao humana, derrubando todas as automacoes. Igualdade exata com a
  // chave do proprio servidor = mesma garantia do role service_role, sem chance de forjar.
  const rawTok = String(authorization || "").replace(/^Bearer\s+/i, "").trim();
  if (rawTok && rawTok === String(Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "")) return;
  const actor = await _accessActor(authorization, true);
  if (actor.profile.status !== "active") throw new Error("Conclua a ativação do seu acesso antes de usar o sistema.");
  if (["master", "admin"].includes(actor.profile.role)) return;
  const key = _firstBodyKey(body), menu = _menuForOperation(key);
  if (menu && !_actorHas(actor, "menus", menu)) throw new Error(`Seu usuário não tem acesso ao menu ${menu}.`);
  const writes: Record<string, string> = {
    metaAction: "campaign.manage", metaCloneCampaign: "campaign.manage", metaCreateAudiences: "campaign.manage", metaCreateCustomList: "campaign.manage", metaCreateSavedAudience: "campaign.manage",
    googleBudget: "campaign.manage", googleCampaignAction: "campaign.manage", googleTermAction: "campaign.manage", googleKeywordAction: "campaign.manage", googleCreateCustomAudience: "campaign.manage",
    crmAndreiaAction: "crm.write", crmCapaAudit: "crm.write", briefingAprovar: "data.write",
  };
  let required = writes[key] || "";
  if (key === "wa" && /send|remove|capi|import|reprocess|poll/i.test(String(body.wa?.op || ""))) required = "crm.write";
  if (required && !_actorHas(actor, "actions", required)) throw new Error("Seu usuário pode consultar, mas não tem permissão para executar esta alteração.");
  const allowed = await _actorClientIds(actor); if (allowed === null) return;
  const requested = [..._bodyClientIds(body)];
  if (requested.some(id => !allowed.includes(id))) throw new Error("Este usuário não tem acesso ao cliente solicitado.");
  const accountIds = _bodyAccountIds(body);
  if (accountIds.length) {
    const clients = await sbGet("clients", "select=id,meta_account_id,google_account_id");
    for (const aid of accountIds) {
      const owner = clients.find((c: any) => [c.meta_account_id, c.google_account_id].some((v: any) => String(v || "").split(/[,;\s]+/).map((z: string) => z.replace(/\D/g, "")).includes(aid)));
      if (!owner || !allowed.includes(String(owner.id))) throw new Error("A conta de anúncios solicitada não pertence a um cliente permitido para este usuário.");
    }
  }
  if (key === "agent" && !requested.length && !actor.profile.all_clients) throw new Error("Selecione um cliente permitido para conversar com a AndréIA.");
}

// Presets de período da tela viram "há quantos dias" (API oficial e Apify só filtram pra trás a
// partir de hoje — não existe "até" arbitrário sem reprocessar paginação inteira por data).
function _minerPeriodo(periodo: string, diasCustom?: number): { days: number; sinceISO: string } {
  const now = new Date();
  let days: number;
  if (periodo === "mes") days = Math.max(1, Math.ceil((now.getTime() - new Date(now.getFullYear(), now.getMonth(), 1).getTime()) / 86400000));
  else if (periodo === "ano") days = Math.max(1, Math.ceil((now.getTime() - new Date(now.getFullYear(), 0, 1).getTime()) / 86400000));
  else if (periodo === "30") days = 30;
  else if (periodo === "365") days = 365;
  else if (periodo === "custom") days = Math.min(365, Math.max(1, Number(diasCustom) || 90));
  else days = 90;
  return { days, sinceISO: new Date(now.getTime() - days * 86400000).toISOString().slice(0, 10) };
}
async function creativeMiner(input: any) {
  const op = String(input?.op || "list"), clientId = String(input?.clientId || "").trim();
  if (!clientId) throw new Error("Selecione um cliente.");
  if (op === "list") return { items: await sbGet("creative_miner_items", `client_id=eq.${encodeURIComponent(clientId)}&select=*&order=created_at.desc&limit=200`) };
  if (op === "remove") { await sbDeleteD("creative_miner_items", `id=eq.${encodeURIComponent(String(input.id || ""))}&client_id=eq.${encodeURIComponent(clientId)}`); return { ok: true }; }
  if (op === "suggest_profiles") {
    const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,instagram_accounts&limit=1`))[0];
    const contas: any[] = Array.isArray(c?.instagram_accounts) ? c.instagram_accounts : [];
    const conta = contas.find((a: any) => a.id === input.instagramId) || contas[0];
    if (!conta?.username) throw new Error("Esse cliente não tem Instagram conectado. Conecte em Configurações do cliente pra buscar perfis relacionados.");
    const suggestions = await _minerRelatedProfiles(conta.username);
    return { profile: conta.username, suggestions, aviso: suggestions.length ? "" : "O Instagram não trouxe perfis relacionados pra este perfil agora — tente de novo mais tarde." };
  }
  if (op === "capture_official") {
    const criterio = String(input.criterio || "engajamento");
    const funil = String(input.funil || "").trim();
    const { days } = _minerPeriodo(String(input.periodo || ""), Number(input.dias));
    const org = await instagramOrganicContent({ clientId, days });
    let posts = [...(org.posts || [])];
    posts = criterio === "recentes" ? posts.sort((a: any, b: any) => String(b.data || "").localeCompare(String(a.data || ""))) : posts.sort((a: any, b: any) => _minerScore(criterio, b) - _minerScore(criterio, a));
    const picked = posts.slice(0, Math.min(50, Number(input.limit) || 20));
    const rows = picked.map((p: any) => ({ id: `cm_${_hash36(clientId + "|" + p.permalink)}`, client_id: clientId, source_type: "instagram_official", source_url: p.permalink, profile: p.username || "", media_type: /VIDEO|REELS/i.test(p.tipo) ? "video" : "image", caption: p.caption || "", media_url: p.midia || null, thumbnail_url: p.midia || null, published_at: p.data || null, metrics: { likes: p.likes, comments: p.comments, reach: p.reach, saved: p.saved, shares: p.shares, views: p.views, engagement: p.eng }, selection_reason: _minerReason(criterio, p, _minerScore(criterio, p)), funnel_target: funil || null, status: "captured", updated_at: new Date().toISOString() }));
    if (rows.length) await _sbUpsert("creative_miner_items", rows, "client_id,source_url");
    return { captured: rows.length, items: rows };
  }
  if (op === "capture_external") {
    const url = String(input.url || "").trim(); if (!/^https:\/\/(www\.)?instagram\.com\//i.test(url)) throw new Error("Informe uma URL pública do Instagram.");
    const criterio = String(input.criterio || "engajamento");
    const funil = String(input.funil || "").trim();
    const limit = Math.min(30, Math.max(1, Number(input.limit) || 12));
    const { sinceISO } = _minerPeriodo(String(input.periodo || ""), Number(input.dias));
    const raw = await _minerApifyCapture(url, Math.max(limit, 30), sinceISO);
    const mapped = raw.map((x: any) => {
      const sourceUrl = String(_minerPick(x, ["url", "postUrl", "inputUrl", "shortCodeUrl"]) || url);
      const type = String(_minerPick(x, ["type", "productType", "mediaType"]) || "post").toLowerCase();
      const video = _minerPick(x, ["videoUrl", "video_url", "displayUrl"]), image = _minerPick(x, ["displayUrl", "imageUrl", "thumbnailUrl"]);
      const likes = Number(_minerPick(x, ["likesCount", "likes"]) || 0), comments = Number(_minerPick(x, ["commentsCount", "comments"]) || 0), views = Number(_minerPick(x, ["videoViewCount", "videoPlayCount", "views"]) || 0);
      return { sourceUrl, type, video, image, likes, comments, views, timestamp: _minerPick(x, ["timestamp", "takenAt", "date"]), profile: String(_minerPick(x, ["ownerUsername", "username", "ownerFullName"]) || ""), caption: String(_minerPick(x, ["caption", "text", "description"]) || "") };
    }).filter((x: any) => x.sourceUrl);
    const ranked = criterio === "recentes" ? mapped.sort((a: any, b: any) => String(b.timestamp || "").localeCompare(String(a.timestamp || ""))) : mapped.sort((a: any, b: any) => _minerScore(criterio, b) - _minerScore(criterio, a));
    const picked = ranked.slice(0, limit);
    const rows = picked.map((x: any) => ({ id: `cm_${_hash36(clientId + "|" + x.sourceUrl)}`, client_id: clientId, source_type: "apify", source_url: x.sourceUrl, profile: x.profile, media_type: /video|reel/.test(x.type) || !!x.video ? "video" : (/sidecar|carousel/.test(x.type) ? "carousel" : "image"), caption: x.caption, media_url: x.video || x.image || null, thumbnail_url: x.image || null, published_at: x.timestamp || null, metrics: { likes: x.likes, comments: x.comments, views: x.views }, selection_reason: _minerReason(criterio, x, _minerScore(criterio, x)), funnel_target: funil || null, status: "captured", updated_at: new Date().toISOString() }));
    if (rows.length) await _sbUpsert("creative_miner_items", rows, "client_id,source_url");
    return { captured: rows.length, items: rows };
  }
  const item = (await sbGet("creative_miner_items", `id=eq.${encodeURIComponent(String(input.id || ""))}&client_id=eq.${encodeURIComponent(clientId)}&select=*&limit=1`))[0];
  if (!item) throw new Error("Referência não encontrada.");
  if (op === "analyze") {
    const client = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,seg,dna&limit=1`))[0];
    const funil = String(input.funil || item.funnel_target || "").trim();
    const analysis = await _minerAnalyzePayload(item, client, funil);
    await sbPatchD("creative_miner_items", `id=eq.${encodeURIComponent(item.id)}`, { analysis, concepts: analysis.ideias_estaticas || [], status: "analyzed", updated_at: new Date().toISOString() });
    return { item: { ...item, analysis, concepts: analysis.ideias_estaticas || [], status: "analyzed" } };
  }
  throw new Error("Operação do Minerador não reconhecida.");
}
async function waCloudConfig(input: any, authorization: string) {
  await _requireCredentialAdmin(authorization);
  const op = String(input?.op || "status"), clientId = String(input?.clientId || "").trim();
  if (!clientId) throw new Error("Selecione o cliente.");
  if (op === "status") {
    const rows = await sbGet("wa_instances", `client_id=eq.${encodeURIComponent(clientId)}&provider=eq.cloud&select=id,client_id,name,phone,status,provider,waba_id,phone_number_id,meta_app_id,verified_name,quality_rating,updated_at&limit=1`);
    const x = rows[0] || null;
    return { configured: !!x?.phone_number_id, instance: x, webhookUrl: x ? `${_SB_URL}/functions/v1/tracking/wa/webhook/${x.id}` : "", verifyToken: x?.id || "" };
  }
  const phoneNumberId = String(input?.phoneNumberId || "").replace(/\D/g, ""), wabaId = String(input?.wabaId || "").replace(/\D/g, ""), appId = String(input?.appId || "").replace(/\D/g, "");
  const token = String(input?.token || "").trim(), appSecret = String(input?.appSecret || "").trim();
  if (!phoneNumberId || !appId || !token) throw new Error("Preencha Phone Number ID, App ID e token permanente.");
  const test = await fetch(`https://graph.facebook.com/v21.0/${phoneNumberId}?fields=id,display_phone_number,verified_name,quality_rating,code_verification_status,platform_type&access_token=${encodeURIComponent(token)}`);
  const tj = await test.json(); if (!test.ok || tj.error) throw new Error(`WhatsApp Meta: ${tj.error?.message || "não consegui validar o número"}`);
  let rows = await sbGet("wa_instances", `client_id=eq.${encodeURIComponent(clientId)}&provider=eq.cloud&select=id&limit=1`);
  const id = rows[0]?.id || `cloud_${crypto.randomUUID().replace(/-/g, "").slice(0, 20)}`;
  if (!appSecret && !rows[0]) throw new Error("Preencha o App Secret na primeira conexão.");
  await _saveSecureCredential(`wa_cloud_token:${id}`, token);
  if (appSecret) await _saveSecureCredential(`wa_cloud_app_secret:${id}`, appSecret);
  const phone = String(tj.display_phone_number || "").replace(/\D/g, "");
  const saveInst = await fetch(`${_SB_URL}/rest/v1/wa_instances?on_conflict=id`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify({ id, client_id: clientId, name: tj.verified_name || "WhatsApp Oficial", phone: phone || null, status: "connected", provider: "cloud", waba_id: wabaId || null, phone_number_id: phoneNumberId, meta_app_id: appId, verified_name: tj.verified_name || null, quality_rating: tj.quality_rating || null, connected_at: new Date().toISOString(), updated_at: new Date().toISOString() }) });
  if (!saveInst.ok) throw new Error("Não consegui salvar a conexão oficial.");
  return { ok: true, configured: true, instance: { id, phone, status: "connected", provider: "cloud", waba_id: wabaId, phone_number_id: phoneNumberId, meta_app_id: appId, verified_name: tj.verified_name || "", quality_rating: tj.quality_rating || "" }, webhookUrl: `${_SB_URL}/functions/v1/tracking/wa/webhook/${id}`, verifyToken: id };
}
async function _crmAndreiaPrepareAction(args: any, client: any, clients: any[]) {
  const p: any = { ...args, client_id: client.id };
  if (args.cliente) {
    const hit = _waResolveClient(String(args.cliente), clients);
    if (!hit) return { error: `Não achei o cliente "${args.cliente}" no sistema.` };
    p.client_id = hit.id;
  }
  if (p.tipo === "criar_tarefa") {
    const team = await sbGet("team", "select=id,name&limit=500");
    const q = String(p.responsavel || "").toLowerCase().trim();
    const owner = q ? (team.find((t: any) => t.name.toLowerCase() === q) || team.find((t: any) => t.name.toLowerCase().includes(q))) : null;
    const dueOk = /^\d{4}-\d{2}-\d{2}$/.test(String(p.quando || ""));
    const missing: string[] = [];
    if (!owner) missing.push(q ? `um responsável válido (${team.map((t: any) => t.name).join(", ")})` : "o responsável");
    if (!dueOk) missing.push("a data");
    if (missing.length) return { error: `Para criar a tarefa, preciso de ${missing.join(" e ")}.` };
    p._owner = owner.id; p.responsavel = owner.name; p._due = p.quando;
  }
  return { action: p, confirmation: _waConfirmText(p, clients) };
}
async function crmAndreia(input: any) {
  const clientId = String(input.clientId || ""), question = String(input.question || "").trim();
  if (!clientId || !question) throw new Error("Cliente e pergunta são obrigatórios.");
  const days = Math.min(180, Math.max(7, Number(input.days) || 30));
  const since = new Date(Date.now() - days * 864e5).toISOString();
  const clients = await sbGet("clients", "select=id,name,seg,dna,meta_account_id,google_account_id,conversion_source,report_sheet_url,report_tabs&limit=1000");
  const client = clients.find((c: any) => c.id === clientId);
  if (!client) throw new Error("Cliente não encontrado.");
  let convs = await sbGet("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&last_at=gte.${encodeURIComponent(since)}&select=id,stage,origin_type,origin,fields,last_at,last_text,num_errado,irrelevante&order=last_at.desc&limit=2000`);
  const f = input.filters || {}, channels = Array.isArray(f.channels) ? f.channels : [], campaigns = Array.isArray(f.campaigns) ? f.campaigns : [];
  convs = convs.filter((cv: any) => {
    if (channels.length && !channels.includes(_crmAiChannel(cv))) return false;
    if (campaigns.length && !campaigns.includes(cv.origin?.campaign || "")) return false;
    if (f.origin && (cv.origin_type || "organico") !== f.origin && !(String(f.origin).startsWith("utm:") && cv.origin_type === "utm" && String(cv.origin?.track_source || "").toLowerCase() === String(f.origin).slice(4))) return false;
    if (f.adset && cv.origin?.adset !== f.adset) return false;
    return true;
  });
  const ids = convs.map((x: any) => x.id);
  const journeys: any[] = [], messages: any[] = [];
  // Prioriza perdas e oportunidades avançadas; limita volume para manter a análise rápida e sem exposição desnecessária.
  const sampleIds = convs.filter((x: any) => /perd|compr|sql|mql|pos/i.test(String(x.stage || ""))).concat(convs).map((x: any) => x.id).filter((x: string, i: number, a: string[]) => a.indexOf(x) === i).slice(0, 450);
  for (let i = 0; i < ids.length; i += 150) {
    const part = ids.slice(i, i + 150); if (!part.length) continue;
    journeys.push(...await sbGet("wa_journey", `conversation_id=in.(${part.map((x: string) => encodeURIComponent(x)).join(",")})&select=conversation_id,from_stage,to_stage,why,source,created_at&order=created_at.asc&limit=3000`));
  }
  for (let i = 0; i < sampleIds.length; i += 100) {
    const part = sampleIds.slice(i, i + 100); if (!part.length) continue;
    messages.push(...await sbGet("wa_messages", `conversation_id=in.(${part.map((x: string) => encodeURIComponent(x)).join(",")})&direction=eq.in&select=conversation_id,text,ts&order=ts.asc&limit=2500`));
  }
  const stageCounts: Record<string, number> = {}, channelCounts: Record<string, number> = {}, campaignCounts: Record<string, number> = {};
  convs.forEach((cv: any) => { const st = cv.stage || "sem_etapa", ch = _crmAiChannel(cv), cp = cv.origin?.campaign || "sem campanha"; stageCounts[st] = (stageCounts[st] || 0) + 1; channelCounts[ch] = (channelCounts[ch] || 0) + 1; campaignCounts[cp] = (campaignCounts[cp] || 0) + 1; });
  const total = convs.length, qual = convs.filter((x: any) => /mql|sql|compr|pos/i.test(String(x.stage || ""))).length, sales = convs.filter((x: any) => /compr|pos/i.test(String(x.stage || ""))).length;
  // Funil por passagem, não apenas por status "perdido": quem está numa etapa avançada necessariamente passou pelas anteriores.
  const reached: Record<string, Set<string>> = { mql: new Set(), sql: new Set(), venda: new Set() };
  const markReached = (cid: string, stage: any) => { const s = String(stage || "").toLowerCase(); if (/mql|qualificad/.test(s)) reached.mql.add(cid); if (/sql|oportun|propost/.test(s)) { reached.mql.add(cid); reached.sql.add(cid); } if (/compr|vend|fech|client|pos/.test(s)) { reached.mql.add(cid); reached.sql.add(cid); reached.venda.add(cid); } };
  convs.forEach((cv: any) => markReached(cv.id, cv.stage));
  // Se a própria IA marcou venda e se corrigiu para uma etapa anterior em até 10 minutos,
  // ignora apenas esse par. Vendas legítimas anteriores/posteriores continuam preservadas.
  const invalidSaleRows = new Set<number>(), aiSaleState: Record<string, { at: number; row: number }> = {};
  journeys.forEach((j: any, i: number) => { const cid = j.conversation_id, at = new Date(j.created_at).getTime(), sale = /compr|vend|fech|client|pos/i.test(String(j.to_stage || "")); if (j.source === "ia" && sale) aiSaleState[cid] = { at, row: i }; else { const s = aiSaleState[cid]; if (s && j.source === "ia" && !sale && at - s.at <= 10 * 60000) { invalidSaleRows.add(s.row); invalidSaleRows.add(i); delete aiSaleState[cid]; } } });
  journeys.forEach((j: any, i: number) => { const bad = invalidSaleRows.has(i); if (!(bad && /compr|vend|fech|client|pos/i.test(String(j.from_stage || "")))) markReached(j.conversation_id, j.from_stage); if (!(bad && /compr|vend|fech|client|pos/i.test(String(j.to_stage || "")))) markReached(j.conversation_id, j.to_stage); });
  const pct = (n: number, d: number) => d ? +(n / d * 100).toFixed(1) : 0;
  const progression = {
    entrada: total,
    chegaram_mql: reached.mql.size,
    nao_qualificaram_como_mql: Math.max(0, total - reached.mql.size),
    taxa_entrada_para_mql_pct: pct(reached.mql.size, total),
    chegaram_sql: reached.sql.size,
    mql_que_ainda_nao_passaram_para_sql: Math.max(0, reached.mql.size - reached.sql.size),
    taxa_mql_para_sql_pct: pct(reached.sql.size, reached.mql.size),
    chegaram_venda: reached.venda.size,
    sql_que_ainda_nao_passaram_para_venda: Math.max(0, reached.sql.size - reached.venda.size),
    taxa_sql_para_venda_pct: pct(reached.venda.size, reached.sql.size),
  };
  // Funil separado por canal: evita que o resumo geral misture Meta, Google e orgânico
  // e permite comparar volume, etapa atual e qualificação de cada origem.
  const channelFunnel: Record<string, any> = {};
  convs.forEach((cv: any) => {
    const ch = _crmAiChannel(cv), st = String(cv.stage || "sem_etapa");
    const a = channelFunnel[ch] ||= { entrada: 0, etapas: {}, qualificados_atuais: 0, vendas_atuais: 0, chegaram_mql: 0, chegaram_sql: 0, chegaram_venda: 0 };
    a.entrada++; a.etapas[st] = (a.etapas[st] || 0) + 1;
    if (/mql|sql|compr|pos/i.test(st)) a.qualificados_atuais++;
    if (/compr|pos/i.test(st)) a.vendas_atuais++;
    if (reached.mql.has(cv.id)) a.chegaram_mql++;
    if (reached.sql.has(cv.id)) a.chegaram_sql++;
    if (reached.venda.has(cv.id)) a.chegaram_venda++;
  });
  Object.values(channelFunnel).forEach((a: any) => {
    a.nao_qualificaram_como_mql = Math.max(0, a.entrada - a.chegaram_mql);
    a.mql_que_ainda_nao_passaram_para_sql = Math.max(0, a.chegaram_mql - a.chegaram_sql);
    a.sql_que_ainda_nao_passaram_para_venda = Math.max(0, a.chegaram_sql - a.chegaram_venda);
    a.taxa_qualificacao_pct = pct(a.chegaram_mql, a.entrada);
    a.taxa_mql_para_sql_pct = pct(a.chegaram_sql, a.chegaram_mql);
    a.taxa_sql_para_venda_pct = pct(a.chegaram_venda, a.chegaram_sql);
  });
  // Origem que efetivamente levou contatos adiante no CRM: anúncio no Meta e palavra-chave no Google.
  const driverMap: Record<string, any> = {};
  convs.forEach((cv: any) => {
    const ch = _crmAiChannel(cv), o = cv.origin || {};
    if (ch !== "meta" && ch !== "google") return;
    const type = ch === "google" ? "palavra_chave" : "anuncio";
    const name = String(ch === "google" ? (o.keyword || o.term || o.utm_term || "(palavra-chave não identificada)") : (o.ad || o.title || "(anúncio não identificado)"));
    const key = `${ch}|${name}|${o.campaign || ""}|${o.adset || o.adgroup || ""}`;
    const a = driverMap[key] ||= { canal: ch, tipo: type, nome: name, campanha: o.campaign || "", conjunto_ou_grupo: o.adset || o.adgroup || "", entrada: 0, etapas_atuais: {}, chegaram_mql: 0, chegaram_sql: 0, chegaram_venda: 0 };
    const st = String(cv.stage || "sem_etapa"); a.entrada++; a.etapas_atuais[st] = (a.etapas_atuais[st] || 0) + 1;
    if (reached.mql.has(cv.id)) a.chegaram_mql++;
    if (reached.sql.has(cv.id)) a.chegaram_sql++;
    if (reached.venda.has(cv.id)) a.chegaram_venda++;
  });
  const conversionDrivers = Object.values(driverMap).sort((a: any, b: any) => (b.chegaram_venda - a.chegaram_venda) || (b.chegaram_sql - a.chegaram_sql) || (b.chegaram_mql - a.chegaram_mql) || (b.entrada - a.entrada)).slice(0, 100);
  const stalledByStage: Record<string, { total: number; acima_3_dias: number; acima_7_dias: number }> = {};
  convs.forEach((cv: any) => { const st = String(cv.stage || "sem_etapa"), age = Math.max(0, (Date.now() - new Date(cv.last_at).getTime()) / 864e5), a = stalledByStage[st] ||= { total: 0, acima_3_dias: 0, acima_7_dias: 0 }; a.total++; if (age >= 3) a.acima_3_dias++; if (age >= 7) a.acima_7_dias++; });
  const half = Date.now() - Math.floor(days / 2) * 864e5, recent = convs.filter((x: any) => new Date(x.last_at).getTime() >= half).length, prior = total - recent;
  const commercial = convs.slice(0, 700).map((cv: any) => ({ id: cv.id, etapa: cv.stage || "sem_etapa", canal: _crmAiChannel(cv), campanha: cv.origin?.campaign || "", conjunto: cv.origin?.adset || cv.origin?.adgroup || "", anuncio: cv.origin?.ad || "", palavra_chave: cv.origin?.keyword || cv.origin?.term || cv.origin?.utm_term || "", campos: _crmAiSafeFields(cv.fields), ultima_mensagem: _crmAiMaskText(cv.last_text).slice(0, 220), numero_errado: !!cv.num_errado, irrelevante: !!cv.irrelevante }));
  const msgBy: Record<string, string[]> = {}; messages.forEach((m: any) => { const t = _crmAiMaskText(m.text).trim(); if (t && t.length > 2) (msgBy[m.conversation_id] ||= []).push(t.slice(0, 320)); });
  const transcripts = commercial.filter((x: any) => msgBy[x.id]?.length).slice(0, 350).map((x: any) => ({ etapa: x.etapa, canal: x.canal, campanha: x.campanha, campos: x.campos, mensagens_lead: msgBy[x.id].slice(0, 10) }));
  const convById: Record<string, any> = {}; convs.forEach((cv: any) => { convById[cv.id] = cv; });
  const lossMoves = journeys.filter((j: any) => /perd|desqual|lost/i.test(String(j.to_stage || "")) || /perd|desist|sem retorno|preco|valor|data|lot|vaga|concorr/i.test(String(j.why || ""))).slice(-600).map((j: any) => { const cv = convById[j.conversation_id] || {}, o = cv.origin || {}; return { de: j.from_stage || "", para: j.to_stage || "", motivo: _crmAiMaskText(j.why).slice(0, 260), data: j.created_at, canal: _crmAiChannel(cv), campanha: o.campaign || "", anuncio: o.ad || "", palavra_chave: o.keyword || o.term || o.utm_term || "" }; });
  const base = { cliente: client.name, segmento: client.seg || "", periodo_dias: days, filtros: f, total, etapas_atuais: stageCounts, progressao_entre_etapas: progression, funil_e_qualificacao_por_canal: channelFunnel, conversoes_por_anuncio_ou_palavra_chave: conversionDrivers, parados_por_etapa_e_tempo: stalledByStage, canais: channelCounts, campanhas: campaignCounts, qualificados: qual, vendas: sales, taxa_qualificacao_pct: total ? +(qual / total * 100).toFixed(1) : 0, taxa_fechamento_pct: total ? +(sales / total * 100).toFixed(1) : 0, atividade_metade_recente: recent, atividade_metade_anterior: prior, movimentos_com_motivo_registrado: lossMoves, leads_comerciais: commercial.map(({ id: _id, ...x }: any) => x), conversas_amostra: transcripts };
  const history = (Array.isArray(input.history) ? input.history : []).slice(-8).map((m: any) => ({ role: m.role === "assistant" ? "assistant" : "user", content: String(m.text || "").slice(0, 1800) }));
  const playbook = await _waPlaybook();
  const unified = await _andreiaUnifiedContext(clientId, "CRM e Analytics");
  const sys = `${unified}

${playbook}

Você é a AndréIA, o cérebro único e conversacional da Central de Gestão. É a mesma inteligência usada no CRM, Analytics, sistema e WhatsApp. Neste momento você está dentro do CRM do cliente ${client.name}, mas pode conversar, consultar dados reais de outras áreas e preparar ações no sistema usando as ferramentas disponíveis.

ROTEAMENTO:
- Pergunta simples ou conversa: responda naturalmente e de forma direta; não force um relatório.
- Pergunta sobre dados: consulte as ferramentas quando o pacote CRM não for suficiente. Nunca invente.
- Pedido de relatório/análise: use o pacote CRM e a estrutura analítica abaixo.
- Pedido de ação: use preparar_acao. Não transforme a solicitação em diagnóstico e não diga que executou. O sistema mostrará uma confirmação antes de executar.
- Para criar tarefa, responsável e data são obrigatórios. Se faltar algo, pergunte apenas o que falta.
- Hoje é ${new Date().toISOString().slice(0, 10)}. Converta "hoje", "amanhã" e dias da semana para AAAA-MM-DD antes de preparar uma ação.
- O cliente atual é ${client.name}; use-o como padrão, salvo se a pessoa citar claramente outro cliente.

Ao analisar o CRM, respeite o cliente, período e filtros ativos. Além de analisar procura por produtos/serviços, perfil e qualidade dos contatos, motivos, passagem entre etapas, canais/campanhas, atendimento, gargalos, conversão e projeções, transforme o diagnóstico em orientação prática para melhorar o funil.

REGRAS:
- Nunca invente número, produto, motivo ou tendência. Diferencie claramente DADO, INFERÊNCIA e PROJEÇÃO.
- Para produtos mais procurados, agrupe sinônimos usando campos comerciais e mensagens dos leads e mostre quantidade apenas quando conseguir contar evidências; se for amostra, diga que é amostra.
- Analise gargalos por PASSAGEM ENTRE ETAPAS usando progressao_entre_etapas: não qualificou como MQL; chegou a MQL mas ainda não passou para SQL; chegou a SQL mas ainda não virou venda. Não exija status "perdido". Diga "ainda não avançou" quando não houver encerramento confirmado.
- Para motivos, use movimentos_com_motivo_registrado e a amostra de mensagens, agrupando assuntos semelhantes. Não trate ausência de motivo como motivo inventado.
- Responda em VISÃO CONSOLIDADA: volumes, percentuais, temas recorrentes e prioridade. Não conte a história de um lead, não cite produto específico de uma única conversa e não faça análise um a um. Tema com uma única ocorrência deve ser tratado como evidência insuficiente, salvo se a pessoa pedir exemplos.
- Seja CONSULTIVA: para cada gargalo relevante, explique (1) o que o dado comprova, (2) quais causas são hipóteses plausíveis, (3) como validar essas hipóteses e (4) o que fazer para melhorar.
- Separe responsabilidades: TRÁFEGO (segmentação, promessa, criativo, palavra-chave, campanha e qualidade do lead), COMERCIAL (tempo de resposta, abordagem, qualificação, follow-up, objeções e fechamento) e PROCESSO/CRM (etapas, campos, automação, registro de motivos e SLA). Não atribua culpa a uma área sem evidência.
- NUNCA afirme que "o público está errado", "o nicho está incorreto" ou equivalente. Contatos aparentemente fora do objetivo comercial principal podem vir de campanhas legítimas com outro objetivo, como recrutamento e recebimento de currículos, suporte, distribuição ou reconhecimento. Apresente apenas como hipótese de investigação e, antes de concluir, verifique objetivo da campanha, criativo, palavra-chave, página de destino e DNA/histórico do cliente. Se faltarem esses dados, recomende confirmar o objetivo; não marque como erro de tráfego.
- Recomendações devem ser específicas ao gargalo encontrado. Evite conselhos genéricos como "melhorar atendimento"; diga qual mudança testar, em qual etapa e qual indicador acompanhar.
- Projeções devem mostrar período, base usada, cálculo/premissa e faixa prudente; não prometa resultado.
- Considere etapa MQL/SQL/comprou conforme configurada no CRM. Não chame todo contato de qualificado.
- Quando a pergunta for "Resumo geral do CRM" ou pedir visão geral, use obrigatoriamente funil_e_qualificacao_por_canal. Separe TODOS os canais presentes (Meta, Google, orgânico e outros) e informe, para cada canal: entrada, quantidade em cada etapa atual, chegaram a MQL, chegaram a SQL, chegaram a venda e taxas de qualificação/passagem. Não misture os canais em um único total sem mostrar essa abertura.
- "Resumo geral do CRM" significa DIAGNÓSTICO COMPLETO, não apenas funil. Analise obrigatoriamente, quando houver dados: (1) volume e evolução da atividade; (2) mídia paga versus orgânico e demais canais; (3) procura por produtos/serviços e temas recorrentes; (4) qualidade e qualificação dos contatos; (5) números e passagem em todas as etapas; (6) gargalos e tempo parado; (7) motivos recorrentes de perda ou não avanço; (8) campanhas e, quando identificados, anúncios Meta ou palavras-chave Google que geraram avanço; (9) sinais sobre atendimento/comercial; (10) projeção prudente para o próximo período; (11) prioridades e plano consultivo. Diga claramente quando uma dimensão não puder ser analisada por falta de registro.
- Quando pedirem quais anúncios ou palavras-chave trouxeram conversões, use conversoes_por_anuncio_ou_palavra_chave: Meta deve ser detalhado por anúncio e Google por palavra-chave. Mostre entrada e avanço para MQL, SQL e venda, além da distribuição das etapas atuais. Não chame MQL ou SQL de venda e não atribua conversão a um item sem identificação.
- A abertura por canal também deve ser usada nos demais relatórios quando for pertinente: procura de produtos/serviços deve relacionar procura e canal; qualificação e gargalos devem comparar etapas e taxas por canal; perdas devem agrupar motivos por canal quando houver volume; projeções devem considerar a base de cada canal. Para Meta, detalhe anúncio quando a análise for de origem/desempenho; para Google, detalhe palavra-chave. Se o dado não estiver identificado, diga isso claramente em vez de omitir ou inventar.
- Se a pergunta não puder ser respondida com estes dados, diga exatamente qual dado está faltando e como a equipe deve registrá-lo.
- Não mencione IDs internos. Não exponha dados pessoais. Responda em português, de modo executivo, claro e acionável.

FORMATO PARA RELATÓRIOS E ANÁLISES (não use em conversa curta ou pedido de ação):
- Use Markdown limpo com títulos iniciados por ##. Nunca use asterisco simples para negrito, nunca use "---" entre parágrafos e nunca coloque vários dados na mesma linha.
- Comece por "## Resumo executivo" com no máximo 3 frases.
- Quando a pergunta envolver funil, use "## Funil no período" e uma lista: Entrada; Não qualificaram como MQL; MQL que ainda não chegaram a SQL; SQL que ainda não viraram venda. Mostre volume e percentual quando houver base.
- Sempre que apresentar o funil, coloque logo abaixo uma legenda curta chamada "**Como classificamos**": **MQL** = contato que demonstrou interesse real e fez pergunta concreta sobre o produto/serviço; **SQL** = contato que definiu melhor o que precisa e está pronto para orçamento, proposta, agendamento ou negociação. Use linguagem simples para o cliente e não transforme a legenda em uma seção longa.
- No resumo geral, inclua "## Canais e qualificação" em tabela compacta com uma linha por canal e SOMENTE estas medidas: Entrada, chegaram a MQL, chegaram a SQL, chegaram à venda e taxas de passagem/qualificação. NÃO inclua nesse resumo geral as colunas de estoque atual "Novo", "Sem etapa", "MQL atual", "SQL atual" e "Comprou atual". O detalhamento de todas as etapas atuais fica reservado ao relatório específico de qualificação ou quando a pessoa pedir explicitamente. Depois destaque diferenças relevantes entre os canais sem declarar causalidade sem evidência.
- No resumo geral completo, use nesta ordem apenas as seções que tenham dados: "## Resumo executivo", "## Visão geral do CRM", "## Canais e qualificação", "## Produtos e serviços procurados", "## Funil e gargalos", "## Perdas e não avanço", "## Campanhas, anúncios e palavras-chave", "## Projeção", "## Diagnóstico consultivo" e "## Plano de melhoria". Preserve a projeção sempre que houver base suficiente.
- Depois use somente as seções relevantes entre: "## Gargalos prioritários", "## Diagnóstico consultivo", "## Motivos recorrentes", "## Projeção" e "## Plano de melhoria".
- Em "## Diagnóstico consultivo", organize cada gargalo como **Dado**, **Hipótese** e **Como validar**.
- Em "## Plano de melhoria", entregue ações em ordem de prioridade e identifique no começo de cada bullet: **Tráfego**, **Comercial** ou **Processo/CRM**. Inclua ação, responsável sugerido e indicador de sucesso. Não inclua prazo.
- Cada bullet deve conter uma ideia e ter no máximo 2 frases. No máximo 5 bullets por seção. Não repita números.
- Destaque apenas números ou termos curtos com **negrito**; nunca deixe um parágrafo inteiro em negrito.
- Se o volume for pequeno, mostre um aviso curto em vez de conclusões extensas. No resumo geral completo, máximo 800 palavras; nas demais respostas, máximo 450 palavras.`;
  const messagesAi: any[] = [{ role: "system", content: sys }, ...history, { role: "user", content: `PERGUNTA: ${question}\n\nPACOTE CRM ATUAL:\n${JSON.stringify(base).slice(0, 110000)}` }];
  let answer = "", action: any = null;
  for (let it = 0; it < 5; it++) {
    const ai = await callOpenAI({ model: "gpt-4o", messages: messagesAi, tools: WA_TOOLS, tool_choice: "auto", max_tokens: 2200, temperature: 0.25 });
    const msg = ai.choices?.[0]?.message || {};
    if (!msg.tool_calls?.length) { answer = String(msg.content || "Não consegui responder agora."); break; }
    messagesAi.push(msg);
    for (const tc of msg.tool_calls) {
      let args: any = {}; try { args = JSON.parse(tc.function.arguments || "{}"); } catch (_e) { /* */ }
      if (tc.function.name === "preparar_acao") {
        const prepared = await _crmAndreiaPrepareAction(args, client, clients);
        if (prepared.error) messagesAi.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify({ erro: prepared.error, instrucao: "Peça somente os dados faltantes ao usuário." }) });
        else { action = prepared.action; answer = prepared.confirmation; messagesAi.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify({ preparado: true, confirmacao: prepared.confirmation }) }); }
      } else {
        const result = await waExecTool(tc.function.name, args, clients);
        messagesAi.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify(result).slice(0, 6000) });
      }
    }
    if (action) break;
  }
  await _andreiaMaybeRemember(clientId, "CRM e Analytics", question);
  return { answer: answer || "Não consegui responder agora.", action, scope: { cliente: client.name, dias: days, conversas: total, filtros: f }, suggestions: ["Resumo geral do CRM", "Resumo geral da procura de produtos e serviços", "Resumo da qualificação dos leads", "Onde os leads deixam de avançar entre MQL, SQL e venda?", "Faça uma projeção prudente para os próximos 30 dias."] };
}

async function crmCapaAudit(input: any) {
  /* Etapa e tempo parado deixam de ser obrigatorios: "" em stage = qualquer etapa, 0 em minHours =
     sem tempo minimo. Antes so dava pra avaliar conversa travada ha 24h numa etapa especifica, o que
     impedia analisar o atendimento de forma geral. O ?? preserva a string vazia, que || comeria. */
  const clientId = String(input.clientId || "");
  const stage = String(input.stage ?? "sql").toLowerCase();
  const qualquerEtapa = !stage;
  const sampleSize = Math.min(20, Math.max(5, Number(input.sampleSize) || 5)), days = Math.min(180, Math.max(7, Number(input.days) || 30));
  const minHours = Math.max(0, Number(input.minHours ?? 24) || 0), since = new Date(Date.now() - days * 864e5).toISOString(), cutoff = Date.now() - minHours * 36e5;
  if (!clientId) throw new Error("Cliente obrigatório.");
  const client = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=id,name,seg,dna&limit=1`))[0];
  if (!client) throw new Error("Cliente não encontrado.");
  let convs = await sbGet("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&last_at=gte.${encodeURIComponent(since)}&select=id,chat_id,name,stage,origin_type,origin,fields,last_at,last_text,num_errado,irrelevante&order=last_at.asc&limit=1000`);
  /* Nem toda conversa do WhatsApp é atendimento. Entravam papo com a própria equipe, fornecedor e
     assunto pessoal, todos ganhando nota e puxando a média. Duas camadas resolvem: aqui o que dá pra
     saber SEM gastar IA (grupo, número nosso), e depois a própria IA julga o contexto de cada conversa.
     Como parte vai ser descartada, buscamos candidatos a mais pra ainda entregar o tanto pedido. */
  const nossosNumeros = new Set((await sbGet("wa_instances", "select=phone")).map((x: any) => String(x.phone || "").replace(/\D/g, "")).filter((x: string) => x.length >= 10));
  const descartesFixos: any[] = [];
  let jaDescartadas = 0;
  const folga = Math.max(2, Math.ceil(sampleSize / 3));
  convs = convs.filter((c: any) => {
    const st = String(c.stage || "sem_etapa").toLowerCase();
    if (!qualquerEtapa && st !== stage) return false;
    if (c.num_errado || c.irrelevante) return false;
    if (c.fields?.capa_irrelevante) { jaDescartadas++; return false; }   // já julgada fora de atendimento antes
    if (minHours > 0 && new Date(c.last_at).getTime() > cutoff) return false;
    const chat = String(c.chat_id || "");
    if (/@g\.us$/i.test(chat)) { descartesFixos.push({ conv: c, motivo: "Grupo do WhatsApp — não é atendimento individual." }); return false; }
    if (nossosNumeros.has(chat.replace(/\D/g, ""))) { descartesFixos.push({ conv: c, motivo: "Conversa com um número da própria operação." }); return false; }
    return true;
  }).slice(0, sampleSize + folga);
  if (!convs.length) return { ok: true, cliente: client.name, stage, requested: sampleSize, audited: 0, descartadas: descartesFixos.map((d: any) => ({ leadName: d.conv.name || "Conversa", conversationId: d.conv.id, motivo: d.motivo })), jaDescartadas, answer: `## Qualidade do Atendimento\n\nNenhuma conversa ${qualquerEtapa ? "" : `em **${stage.toUpperCase()}** `}${minHours > 0 ? `parada há pelo menos ${minHours}` : "no período"} horas dentro dos últimos ${days} dias${descartesFixos.length || jaDescartadas ? ` (${descartesFixos.length + jaDescartadas} foram descartadas por não serem atendimento a cliente)` : ""}.` };
  const ids = convs.map((c: any) => c.id), allMsgs: any[] = [];
  for (let i = 0; i < ids.length; i += 10) {
    const part = ids.slice(i, i + 10);
    allMsgs.push(...await sbGet("wa_messages", `conversation_id=in.(${part.map((x: string) => encodeURIComponent(x)).join(",")})&select=conversation_id,direction,text,ts&order=ts.asc&limit=1200`));
  }
  const convName: Record<string, string> = {}; convs.forEach((c: any) => { convName[c.id] = c.name || ""; });
  const by: Record<string, any[]> = {}; allMsgs.forEach((m: any) => { const txt = _crmAiMaskText(m.text, [convName[m.conversation_id]]).trim(); if (txt) (by[m.conversation_id] ||= []).push({ quem: m.direction === "out" ? "equipe" : "lead", texto: txt.slice(0, 500), data: m.ts }); });
  /* total_mensagens e trecho_mostrado entram porque a IA só recebe as últimas 30 falas: sem isso ela
     julgava uma conversa de 215 mensagens como se fossem 30 e chamava de "atendimento interrompido"
     algo que já vinha de meses de relação. Contexto é o que separa a conversa ruim da conversa que
     nem era atendimento. */
  const refs: Record<string, any> = {}, cases = convs.map((cv: any, i: number) => {
    const ref = `C${i + 1}`; refs[ref] = cv;
    const todas = by[cv.id] || [], amostra = todas.slice(-30);
    return {
      ref, etapa: cv.stage, canal: _crmAiChannel(cv), campanha: cv.origin?.campaign || "", conjunto_ou_grupo: cv.origin?.adset || cv.origin?.adgroup || "", anuncio: cv.origin?.ad || "", palavra_chave: cv.origin?.keyword || cv.origin?.term || "",
      horas_sem_avancar: Math.round((Date.now() - new Date(cv.last_at).getTime()) / 36e5),
      veio_de_anuncio: cv.origin_type === "anuncio",
      total_mensagens: todas.length, mensagens_do_lead: todas.filter((m: any) => m.quem === "lead").length, mensagens_da_equipe: todas.filter((m: any) => m.quem === "equipe").length,
      trecho_mostrado: todas.length > amostra.length ? `últimas ${amostra.length} de ${todas.length} mensagens` : "conversa inteira",
      campos: _crmAiSafeFields(cv.fields), conversa: amostra,
    };
  });
  const playbook = await _waPlaybook(); const results: any[] = [];
  for (let i = 0; i < cases.length; i += 5) {
    const part = cases.slice(i, i + 5);
    const prompt = `${playbook}\n\nVocê é a AndréIA realizando uma AVALIAÇÃO DE QUALIDADE do atendimento comercial.\n\nPRIMEIRO PASSO, ANTES DE QUALQUER NOTA: decida se a conversa é mesmo um ATENDIMENTO A CLIENTE OU LEAD. Nem tudo que chega no WhatsApp do cliente é atendimento. Marque "relevante": false, explique em "motivo_irrelevancia" e NÃO dê nota (use nota 0) quando a conversa for: papo com a própria equipe, sócio ou colega de trabalho; fornecedor, parceiro, prestador ou cobrança; assunto pessoal, familiar ou social; grupo, lista de transmissão, robô, aviso automático ou spam; número errado; teste interno; ou conversa sem qualquer conteúdo comercial. Use o contexto inteiro para decidir — quem procurou quem, o assunto, o vocabulário, o histórico e os campos do CRM.\n\nAtenção ao contexto antes de julgar: "total_mensagens" diz o tamanho real da conversa e "trecho_mostrado" avisa quando você está vendo só o final. Conversa longa e antiga costuma ser relação em andamento, não atendimento interrompido — não chame de abandono o que é continuidade. Já lead que escreveu e ficou sem resposta É relevante e merece nota baixa: é atendimento perdido, não conversa irrelevante.\n\nPara as conversas relevantes, analise de forma rigorosa e consultiva. O objetivo não é culpar: é descobrir o que impediu o próximo passo e ensinar a melhor condução. Nunca invente falas ou fatos. Quotes devem ser trechos EXATOS presentes na conversa, já anonimizada. A nota 0–10 deve considerar velocidade percebida, descoberta da necessidade, clareza, personalização, tratamento de objeção, CTA/próximo passo e follow-up. Não declare que público/nicho está errado; objetivos alternativos são possíveis. A mensagem recomendada deve ser pronta para envio, específica ao contexto, sem promessas inventadas.\n\nRetorne SOMENTE JSON válido: {"casos":[{"ref":"C1","relevante":true,"motivo_irrelevancia":"","nota":0,"diagnostico":"","ponto_de_quebra":"","quotes":[{"quem":"lead|equipe","texto":"","evidencia":""}],"faltou":[""],"mensagem_recomendada":"","follow_up":"","acoes_trafego":[""],"acoes_comercial":[""],"acoes_processo":[""]}]}\n\nCASOS:\n${JSON.stringify(part)}`;
    const parsed = await _callOpenAIJson([{ role: "user", content: prompt }]);
    results.push(...(Array.isArray(parsed.casos) ? parsed.casos : []));
  }
  /* Só o que era atendimento de verdade entra na nota, nos padrões e na média. O resto não some: volta
     como "descartadas", com o motivo, pra ela conferir se a IA acertou o corte. */
  const relevantes = results.filter((x: any) => x && x.relevante !== false);
  const descartesIA = results.filter((x: any) => x && x.relevante === false);
  const descartadas = [
    ...descartesFixos.map((d: any) => ({ leadName: d.conv.name || "Conversa", conversationId: d.conv.id, motivo: d.motivo })),
    ...descartesIA.map((x: any) => ({ leadName: (refs[x.ref]?.name) || "Conversa", conversationId: refs[x.ref]?.id || null, motivo: String(x.motivo_irrelevancia || "Não é atendimento a cliente ou lead.").slice(0, 300) })),
  ];
  if (!relevantes.length) {
    return { ok: true, cliente: client.name, stage, requested: sampleSize, audited: 0, descartadas, jaDescartadas, answer: `## Qualidade do Atendimento\n\nNenhuma das ${results.length + descartesFixos.length} conversas do período era atendimento a cliente ou lead — todas foram descartadas (equipe, fornecedor, assunto pessoal, grupo ou número da própria operação). Aumente o período ou o número de conversas.` };
  }
  const aggregate = await _callOpenAIJson([{ role: "user", content: `Você é a AndréIA. Consolide esta avaliação de qualidade do atendimento sem identificar cliente ou pessoas e sem analisar um a um novamente. Agrupe padrões, quantifique ocorrências somente contando os casos fornecidos e proponha melhorias práticas. Sem prazo. Separe Tráfego, Comercial e Processo/CRM. Inclua treinamento recomendado e indicadores para acompanhar. Retorne SOMENTE JSON: {"resumo":"","padroes":[{"tema":"","ocorrencias":0,"impacto":""}],"trafego":[""],"comercial":[""],"processo":[""],"treinamento":[""],"indicadores":[""]}. CASOS: ${JSON.stringify(relevantes)}` }]);
  const auditedAt = new Date().toISOString(), auditId = _wuid();
  const scores = relevantes.map((x: any) => Math.max(0, Math.min(10, Number(x.nota) || 0)));
  const averageScore = scores.length ? Math.round(scores.reduce((a: number, b: number) => a + b, 0) / scores.length * 10) / 10 : 0;
  await sbPost("crm_capa_audits", { id: auditId, client_id: clientId, stage, days, min_hours: minHours, requested: sampleSize, audited: relevantes.length, average_score: averageScore, aggregate, created_by: input.createdBy || null, created_at: auditedAt });
  const caseRows: any[] = [];
  /* Descartada pela IA fica registrada com o motivo e recebe uma marca no próprio card: na próxima
     avaliação ela não volta pra fila e não gasta IA de novo. A marca é nossa (fields.capa_irrelevante),
     não mexe em nada do cliente, e ela pode desfazer pelo CRM se a IA tiver errado. */
  for (const item of descartesIA) {
    const cv = refs[item.ref]; if (!cv) continue;
    const motivo = String(item.motivo_irrelevancia || "Não é atendimento a cliente ou lead.").slice(0, 300);
    await sbPatchD("wa_conversations", `id=eq.${encodeURIComponent(cv.id)}`, { fields: { ...(cv.fields || {}), capa_irrelevante: { motivo, at: auditedAt, audit_id: auditId } } });
    caseRows.push({ id: _wuid(), audit_id: auditId, client_id: clientId, conversation_id: cv.id, stage, channel: _crmAiChannel(cv), score: 0, relevant: false, irrelevance_reason: motivo, diagnosis: "", break_point: "", themes: [], recommended_message: "", follow_up: "", traffic_actions: [], commercial_actions: [], process_actions: [], created_at: auditedAt });
  }
  for (const item of relevantes) {
    const cv = refs[item.ref]; if (!cv) continue;
    const score = Math.max(0, Math.min(10, Number(item.nota) || 0));
    const fields = { ...(cv.fields || {}), capa: { audited_at: auditedAt, audit_id: auditId, stage, score } };
    const tags = Array.isArray(fields.tags) ? fields.tags : [];
    if (!tags.includes("Qualidade avaliada")) fields.tags = [...tags, "Qualidade avaliada"];
    await sbPatchD("wa_conversations", `id=eq.${encodeURIComponent(cv.id)}`, { fields });
    item.conversationId = cv.id; item.leadName = cv.name || "Conversa";
    caseRows.push({ id: _wuid(), audit_id: auditId, client_id: clientId, conversation_id: cv.id, stage, channel: _crmAiChannel(cv), score, relevant: true, diagnosis: item.diagnostico || "", break_point: item.ponto_de_quebra || "", themes: item.faltou || [], recommended_message: item.mensagem_recomendada || "", follow_up: item.follow_up || "", traffic_actions: item.acoes_trafego || [], commercial_actions: item.acoes_comercial || [], process_actions: item.acoes_processo || [], created_at: auditedAt });
  }
  if (caseRows.length) await sbPost("crm_capa_cases", caseRows as any);
  return { ok: true, auditId, cliente: client.name, stage, requested: sampleSize, audited: relevantes.length, averageScore, minHours, days, cases: relevantes, descartadas, jaDescartadas, aggregate, auditedAt };
}

async function crmCapaDashboard(input: any) {
  const clientId = String(input.clientId || ""), days = Math.min(365, Math.max(7, Number(input.days) || 90));
  if (!clientId) throw new Error("Cliente obrigatório.");
  const since = new Date(Date.now() - days * 864e5).toISOString();
  const [audits, cases] = await Promise.all([
    sbGet("crm_capa_audits", `client_id=eq.${encodeURIComponent(clientId)}&created_at=gte.${encodeURIComponent(since)}&select=id,stage,audited,average_score,aggregate,created_at&order=created_at.asc&limit=500`),
    // relevant=is.true: conversa descartada (equipe, fornecedor, pessoal) fica gravada pra consulta,
    // mas não entra em média, tema nem fila de retomada — senão ela puxa o histórico inteiro pra baixo.
    sbGet("crm_capa_cases", `client_id=eq.${encodeURIComponent(clientId)}&created_at=gte.${encodeURIComponent(since)}&relevant=is.true&select=id,audit_id,conversation_id,stage,channel,score,diagnosis,break_point,themes,recommended_message,created_at&order=created_at.desc&limit=2000`),
  ]);
  const byStage: Record<string, any> = {}, byChannel: Record<string, any> = {}, themes: Record<string, number> = {};
  for (const c of cases) {
    const add = (map: any, key: string) => { const x = map[key] ||= { key, cases: 0, total: 0, low: 0 }; x.cases++; x.total += Number(c.score) || 0; if ((Number(c.score) || 0) < 6) x.low++; };
    add(byStage, c.stage || "sem_etapa"); add(byChannel, c.channel || "organico");
    for (const t of (Array.isArray(c.themes) ? c.themes : [])) { const k = String(t || "").trim(); if (k) themes[k] = (themes[k] || 0) + 1; }
  }
  const finish = (m: any) => Object.values(m).map((x: any) => ({ ...x, average: x.cases ? Math.round(x.total / x.cases * 10) / 10 : 0 })).sort((a: any, b: any) => b.cases - a.cases);
  const weekly: Record<string, any> = {};
  for (const a of audits) { const d = new Date(a.created_at), day = d.getUTCDay() || 7; d.setUTCDate(d.getUTCDate() - day + 1); const key = d.toISOString().slice(0, 10); const w = weekly[key] ||= { week: key, audits: 0, cases: 0, weighted: 0 }; w.audits++; w.cases += Number(a.audited) || 0; w.weighted += (Number(a.average_score) || 0) * (Number(a.audited) || 0); }
  const trend = Object.values(weekly).map((w: any) => ({ ...w, average: w.cases ? Math.round(w.weighted / w.cases * 10) / 10 : 0 }));
  const queue = cases.filter((c: any) => Number(c.score) < 6).slice(0, 30);
  return { ok: true, days, audits: audits.length, cases: cases.length, averageScore: cases.length ? Math.round(cases.reduce((s: number, c: any) => s + (Number(c.score) || 0), 0) / cases.length * 10) / 10 : 0, trend, byStage: finish(byStage), byChannel: finish(byChannel), themes: Object.entries(themes).map(([theme, count]) => ({ theme, count })).sort((a: any, b: any) => b.count - a.count).slice(0, 12), queue };
}

async function crmAndreiaAction(input: any) {
  const action = input?.action || {}, clientId = String(input?.clientId || action.client_id || "");
  if (!action.tipo) throw new Error("Ação inválida.");
  const allowed = new Set(["criar_tarefa", "criar_reuniao", "cancelar_reuniao", "pausar_campanha", "reativar_campanha", "orcamento", "duplicar_campanha", "criar_lancamento", "dar_baixa"]);
  if (!allowed.has(String(action.tipo))) throw new Error("Esta ação não está liberada no CRM.");
  return { message: await waAgentExec(action, clientId || null) };
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
  // KPIs PRINCIPAIS da campanha, já prontos pelo OBJETIVO dela: Gasto · Resultado · CPR (nunca força CTR/CPC fora de tráfego)
  const tipo = (c.objetivo && c.objetivo.tipo) || "";
  return { kpi: `Gasto ${_fmtR(c.spend || 0)} · ${_objRC(c, false, tipo)}` };
}
// Resumo de UM cliente com KPIs POR CAMPANHA (cada uma pela métrica do objetivo dela) + total filtrado, pro período pedido
// Saldo pré-pago (pix/boleto) das contas Meta do cliente — pro gestor não precisar abrir o gerenciador só pra isso.
async function waMetaSaldo(clientId: string) {
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,meta_account_id`))[0];
  if (!c) return { erro: "cliente não encontrado" };
  const ids = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) return { cliente: c.name, aviso: "cliente sem Meta Ads vinculado" };
  const r = await metaFunding({ accounts: ids.map((id: string) => ({ id, name: id })) }).catch((e: any) => ({ accounts: [], erro: e?.message || String(e) }));
  const contas = (r.accounts || []).map((a: any) => ({ conta: a.name, saldo: a.saldo, tipo: a.tipo, erro: a.error }));
  const comSaldo = contas.filter((a: any) => a.saldo != null);
  return { cliente: c.name, contas, saldoTotal: comSaldo.length ? comSaldo.reduce((s: number, a: any) => s + a.saldo, 0) : null };
}
async function waMetaResumo(clientId: string, dias: number, nivel = "campanha") {
  const detalhar = nivel === "conjunto" || nivel === "anuncio"; // conjunto/anúncio só quando pedido
  const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,meta_account_id`))[0];
  if (!c) return { erro: "cliente não encontrado" };
  const ids = String(c.meta_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean);
  if (!ids.length) return { cliente: c.name, aviso: "cliente sem Meta Ads vinculado" };
  const accounts = ids.map((id: string) => ({ id, name: id }));
  const since = new Date(Date.now() - dias * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
  const [r, ent] = await Promise.all([
    metaAdsInsights({ accounts, since, until, byCampaign: true, byAd: detalhar }).catch(() => null),
    metaEntities({ accounts }).catch(() => null),
  ]);
  const gasto = (r && r.total && r.total.spend) || 0;
  const objM = _domObj(r);
  if (gasto <= 0) return { cliente: c.name, dias, objetivo: _objLabel(objM), semGastoNoPeriodo: true, campanhas: [] };
  const budgetByName: Record<string, number> = {};
  if (ent) (ent.campaigns || []).forEach((x: any) => { if (x.status === "ACTIVE" || x.entrega === "ACTIVE") budgetByName[x.nome] = x.orcamentoDiario; });
  // hierarquia campanha › conjunto › anúncio (dos anúncios com gasto) — SÓ quando pedirem detalhe
  const hier: Record<string, any> = {};
  if (detalhar) ((r && r.ads) || []).filter((a: any) => (a.spend || 0) > 0).forEach((a: any) => {
    const cn = a.campaign || "—"; if (!hier[cn]) hier[cn] = {};
    const sn = a.adset || "—"; const S = hier[cn][sn] || (hier[cn][sn] = { nome: sn, spend: 0, clicks: 0, impressions: 0, reach: 0, purchases: 0, revenue: 0, leads: 0, conversas: 0, videoViews: 0, engajamentos: 0, ads: [] });
    ["spend", "clicks", "impressions", "reach", "purchases", "revenue", "leads", "conversas", "videoViews", "engajamentos"].forEach((k) => S[k] += (a[k] || 0));
    S.ads.push({ nome: a.adName || "(sem nome)", spend: a.spend || 0, clicks: a.clicks || 0, impressions: a.impressions || 0, reach: a.reach || 0, purchases: a.purchases || 0, revenue: a.revenue || 0, leads: a.leads || 0, conversas: a.conversas || 0, videoViews: a.videoViews || 0, engajamentos: a.engajamentos || 0 });
  });
  // métricas cruas: SEMPRE vão no payload (não custa nada, já vem da API) — o PROMPT decide se mostra
  // (resumo padrão = só o kpi; se o usuário pedir formato/métricas específicas, a IA usa estes campos).
  const _brutas = (x: any) => ({ impressoes: Math.round(x.impressions || 0), cliques: Math.round(x.clicks || 0), alcance: Math.round(x.reach || 0), ctrPct: x.impressions ? +((x.clicks / x.impressions) * 100).toFixed(2) : 0, cpm: x.impressions ? +((x.spend / x.impressions) * 1000).toFixed(2) : 0, cpc: x.clicks ? +(x.spend / x.clicks).toFixed(2) : 0 });
  const campanhas = ((r && r.campaigns) || []).filter((x: any) => (x.spend || 0) > 0).slice(0, 20).map((x: any) => {
    const tipo = (x.objetivo && x.objetivo.tipo) || "";
    const base: any = { nome: x.campaign, objetivo: (x.objetivo && x.objetivo.rotulo) || "", orcamentoDiario: budgetByName[x.campaign] || undefined, ..._waCampKpi(x), spend: x.spend || 0, ..._brutas(x) };
    if (!detalhar) return base; // resumo padrão = só campanha
    const H = hier[x.campaign] || {};
    base.conjuntos = Object.values(H).sort((a: any, b: any) => b.spend - a.spend).slice(0, 10).map((S: any) => { const conj: any = { nome: S.nome, kpi: `Gasto ${_fmtR(S.spend)} · ${_objRC(S, false, tipo)}`, spend: S.spend, ..._brutas(S) }; if (nivel === "anuncio") conj.anuncios = S.ads.sort((a: any, b: any) => b.spend - a.spend).slice(0, 10).map((ad: any) => ({ nome: ad.nome, kpi: `Gasto ${_fmtR(ad.spend)} · ${_objRC(ad, false, tipo)}`, spend: ad.spend, ..._brutas(ad) })); return conj; });
    return base;
  });
  // consolidado = kpi verbatim (Gasto · Resultado · CPR) + as métricas cruas do total, pro caso de formato customizado
  return { cliente: c.name, dias, objetivo: _objLabel(objM), kpi: `Gasto ${_fmtR(gasto)} · ${_objRC(r.total, false, objM)}`, gastoTotal: gasto, ..._brutas(r.total), campanhas };
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
  // CRM (WhatsApp) — se o cliente tem funil de atendimento
  const crm = await waCrmStats(c.id, Math.max(dias, 30)).catch(() => null);
  if (crm && crm.total) {
    s += `\n🗣 *Atendimento (CRM)*\n• Leads: *${_fmtN(crm.total)}*${crm.deAnuncio ? ` (${crm.deAnuncio} de anúncio)` : ""}\n• Qualificados (MQL+): *${crm.qualificados}* (${crm.taxaQualificacao}%)\n`;
    if (crm.vendas) s += `• Vendas: *${crm.vendas}* (${crm.taxaConversao}%)\n`;
  }
  // Análise
  if (comAnalise) { const a = await _waAnaliseCliente(c.name, mt, gt, objM, objG); if (a) s += `\n💬 ${a}\n`; }
  s += `${DIV}\n_GT Marketing • Gestão de Tráfego_`;
  return s;
}
async function waExecTool(name: string, args: any, clients: any[]) {
  if (name === "consultar_banco") return await waQueryTable(args);
  if (name === "relatorio_cliente") { const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" }; const rep = await waRelatorioCliente(c, Number(args.dias) || 7); return { _cid: c.id, relatorio: rep, instrucao: "Envie o campo 'relatorio' EXATAMENTE como está, sem reescrever nem resumir." }; }
  if (name === "financeiro") return await waFinanceiro(args);
  if (name === "crm_funil") { const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" }; const s = await waCrmStats(c.id, Number(args.dias) || 30); return s ? { cliente: c.name, ...s } : { cliente: c.name, aviso: "cliente sem CRM/leads no período" }; }
  if (name === "reunioes") return await waReunioes(args);
  if (name === "resumo_todos_clientes") { const msgs = await waAgentAllClientsSummary(Number(args.dias) || 7); return { texto: msgs.join("\n\n") }; }
  if (name === "meta_insights") { const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" }; const r = await waMetaResumo(c.id, Number(args.dias) || 7, args.nivel || "campanha"); return { _cid: c.id, ...r }; }
  if (name === "meta_saldo") { const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" }; const r = await waMetaSaldo(c.id); return { _cid: c.id, ...r }; }
  // Curadoria dos melhores posts do perfil (aba Social). A AndréIA não enxergava o orgânico: perguntavam
  // "quais os melhores criativos" e ela só tinha dado de anúncio pago pra responder.
  if (name === "instagram_organico") {
    const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" };
    const dias = Number(args.dias) || 90, qtd = Math.min(25, Math.max(3, Number(args.quantidade) || 10));
    const r: any = await instagramOrganicContent({ clientId: c.id, days: dias }).catch((e: any) => ({ erro: String(e?.message || e) }));
    if (r && r.erro) return { cliente: c.name, dias, erro: r.erro };
    const posts = (r.posts || []).filter((p: any) => p.eng != null || p.reach);
    if (!posts.length) return { cliente: c.name, dias, aviso: "sem posts orgânicos no período (ou o Instagram deste cliente não está conectado na Configuração do Cliente)" };
    const fmt = (p: any) => p.tipo === "VIDEO" ? "reel/vídeo" : p.tipo === "CAROUSEL_ALBUM" ? "carrossel" : "imagem";
    const medEng = +(posts.reduce((s: number, p: any) => s + (p.eng || 0), 0) / posts.length).toFixed(2);
    // melhor formato e melhor dia da semana: é o que embasa a recomendação de "o que impulsionar/repetir"
    const porFormato: Record<string, { n: number; eng: number }> = {}, porDia: Record<string, { n: number; eng: number }> = {};
    const DIAS = ["domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado"];
    for (const p of posts) {
      const f = fmt(p); (porFormato[f] ||= { n: 0, eng: 0 }); porFormato[f].n++; porFormato[f].eng += p.eng || 0;
      const d = DIAS[new Date(p.data).getDay()]; (porDia[d] ||= { n: 0, eng: 0 }); porDia[d].n++; porDia[d].eng += p.eng || 0;
    }
    const media = (o: Record<string, { n: number; eng: number }>) => Object.entries(o).map(([k, v]) => ({ chave: k, posts: v.n, engMedio: +(v.eng / v.n).toFixed(2) })).sort((a, b) => b.engMedio - a.engMedio);
    return {
      _cid: c.id, cliente: c.name, perfis: r.perfis, dias, totalPosts: posts.length, engajamentoMedioPct: medEng,
      nota: "ranking do ORGÂNICO (perfil), do melhor pro pior por engajamento. eng% = (curtidas+comentários+salvamentos+compartilhamentos) ÷ alcance. 'destaque' marca quem passou de 1,5x a média do período.",
      porFormato: media(porFormato), porDiaDaSemana: media(porDia),
      melhores: posts.slice(0, qtd).map((p: any, i: number) => ({
        posicao: i + 1, formato: fmt(p), data: String(p.data || "").slice(0, 10), engPct: p.eng, alcance: p.reach,
        curtidas: p.likes, comentarios: p.comments, salvamentos: p.saved, compartilhamentos: p.shares, views: p.views,
        destaque: !!(p.eng && medEng && p.eng >= medEng * 1.5), link: p.permalink,
        legenda: String(p.caption || "").replace(/\s+/g, " ").slice(0, 220),
      })),
      erros: r.erros && r.erros.length ? r.erros : undefined,
    };
  }
  if (name === "google_insights") {
    const c = _waResolveClient(args.cliente, clients); if (!c) return { erro: "cliente não encontrado" };
    const gIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim()).filter(Boolean); if (!gIds.length) return { cliente: c.name, aviso: "cliente sem Google Ads vinculado" };
    const d = Number(args.dias) || 30; const since = new Date(Date.now() - d * 864e5).toISOString().slice(0, 10), until = new Date().toISOString().slice(0, 10);
    const r: any = await googleAdsInsights({ accounts: gIds.map((id: string) => ({ id, name: id })), since, until, byCampaign: true }).catch((e: any) => ({ erro: String(e?.message || e) }));
    if (r && r.erro) return { cliente: c.name, dias: d, erro: r.erro };
    const total = r && r.total;
    // objetivo DOMINANTE do Google (por gasto) — pra mostrar o resultado certo (ex: Vídeo → views)
    let objG: string | null = null;
    if (r && r.campaigns && r.campaigns.length) { const bs: any = {}; r.campaigns.forEach((x: any) => { const tp = (x.objetivo && x.objetivo.tipo) || "conversao"; bs[tp] = (bs[tp] || 0) + (x.spend || 0); }); const dom = Object.entries(bs).sort((a: any, b: any) => b[1] - a[1])[0]; objG = dom ? dom[0] : null; }
    const gasto = (total && total.spend) || 0;
    if (gasto <= 0) return { cliente: c.name, _cid: c.id, dias: d, objetivo: objG, semGastoNoPeriodo: true, campanhas: [] };
    const kpi = `Gasto ${_fmtR(gasto)} · ${_objRC(total, true, objG)}`;
    const campanhas = ((r && r.campaigns) || []).filter((x: any) => (x.spend || 0) > 0).slice(0, 20).map((x: any) => ({ nome: x.campaign, objetivo: (x.objetivo && x.objetivo.rotulo) || "", kpi: `Gasto ${_fmtR(x.spend || 0)} · ${_objRC(x, true, x.objetivo && x.objetivo.tipo)}` }));
    return { cliente: c.name, _cid: c.id, dias: d, objetivo: objG, kpi, campanhas };
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
  if (p.tipo === "orcamento") return `${cli}Ajustar o orçamento diário d${p.conjunto ? `o conjunto "${p.conjunto}" (campanha "${p.campanha || ""}")` : `a campanha "${p.campanha || ""}"`} pra R$${p.novoValor}. Confirma?`;
  if (p.tipo === "duplicar_campanha") return `${cli}Duplicar a campanha "${p.campanha || ""}" (cópia pausada). Confirma?`;
  if (p.tipo === "criar_lancamento") return `${cli}Criar lançamento ${p.natureza || "receita"} de R$${p.valor} — ${p.descricao || ""} (venc. ${p.vencimento || "hoje"}). Confirma?`;
  if (p.tipo === "dar_baixa") return `${cli}Dar baixa (marcar como pago) no lançamento "${p.descricao || ""}". Confirma?`;
  if (p.tipo === "gerar_contrato") { const v = Number(p.valorMensal) || 0; const promo = Number(p.mesesPromo) || 0; return `📄 Gerar o *CONTRATO* em PDF pra *${p.razaoSocial || p.nome || "?"}*${v ? ` — honorário ${_fmtR(v)}/mês` : ""}${promo ? ` (promocional nos ${promo} primeiros meses)` : ""}${p.mesesFidelidade ? ` · fidelidade ${p.mesesFidelidade} meses` : ""}. Te mando o PDF aqui. Confirma?`; }
  if (p.tipo === "cadastrar_cliente") { const canais = _canaisNorm(p.canais); const nm = String(p.nomeSistema || p.nome || p.razaoSocial || "?").trim(); const rz = String(p.razaoSocial || "").trim(); const dif = rz && rz.toLowerCase() !== nm.toLowerCase(); return `🗂 Cadastrar no sistema com o nome *${nm}*${dif ? ` _(razão social do contrato: ${rz})_` : ""}${p.nicho ? ` · nicho: ${p.nicho}` : ""}${Number(p.verba) ? ` · verba: ${_fmtR(Number(p.verba))}` : ""}${Number(p.valorMensal || p.fee) ? ` · fee: ${_fmtR(Number(p.valorMensal || p.fee))}` : ""}${canais.length ? ` · canais: ${canais.join(" + ")}` : ""}. Confirma?`; }
  if (p.tipo === "criar_tarefas_onboarding") { const canais = _canaisNorm(p.canais); return `📋 Criar as *tarefas de onboarding*${cli ? ` de ${cli.replace("[", "").replace("] ", "")}` : ""} pros canais *${canais.length ? canais.join(" + ") : "?"}*${p.responsavel ? ` (resp. ${p.responsavel})` : ""} — checklist geral + as atividades específicas de cada canal, distribuídas nos próximos 7 dias. Confirma?`; }
  return `${cli}Confirma a ação?`;
}
// baixa uma URL e devolve base64 (para anexar PDF à IA)
async function waFetchB64(url: string): Promise<string> {
  try {
    const r = await fetch(url); if (!r.ok) return "";
    const buf = new Uint8Array(await r.arrayBuffer());
    let bin = ""; const chunk = 0x8000;
    for (let i = 0; i < buf.length; i += chunk) bin += String.fromCharCode.apply(null, buf.subarray(i, i + chunk) as any);
    return btoa(bin);
  } catch { return ""; }
}
// transcreve um áudio (URL) via Whisper
async function waTranscribe(url: string): Promise<string> {
  const gem = Deno.env.get("GEMINI_API_KEY");
  try {
    const a = await fetch(url); if (!a.ok) return "";
    const blob = await a.blob();
    if (gem) { // Gemini transcreve o áudio direto (o Whisper é exclusivo da OpenAI)
      const buf = new Uint8Array(await blob.arrayBuffer());
      let bin = ""; for (let i = 0; i < buf.length; i += 8192) bin += String.fromCharCode(...buf.subarray(i, i + 8192));
      const b64 = btoa(bin);
      const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=${gem}`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contents: [{ parts: [{ text: "Transcreva este áudio em português, apenas o texto falado." }, { inline_data: { mime_type: blob.type || "audio/ogg", data: b64 } }] }] }),
      });
      const j = await r.json();
      const u = j?.usageMetadata || {};
      await _regUsoIa("gemini", "transcricao_audio", "gemini-3.5-flash", u.promptTokenCount || 0, u.candidatesTokenCount || 0);
      return String(j?.candidates?.[0]?.content?.parts?.[0]?.text || "").trim();
    }
    const key = Deno.env.get("OPENAI_API_KEY"); if (!key) return "";
    const fd = new FormData(); fd.append("file", blob, "audio.mp3"); fd.append("model", "whisper-1");
    const r = await fetch("https://api.openai.com/v1/audio/transcriptions", { method: "POST", headers: { Authorization: `Bearer ${key}` }, body: fd });
    const j = await r.json();
    // sem duracao na resposta, estima pelo tamanho do arquivo (audio de voz ~2 KB/s)
    await _regUsoIa("openai", "transcricao_audio", "whisper-1", Math.max(1, Math.round((blob.size || 0) / 2048)), 0);
    return (j && j.text) ? String(j.text).trim() : "";
  } catch { return ""; }
}
// Extrai o conteúdo útil de um anexo (imagem/PDF/texto) como texto — pra orientar a IA de palavras-chave.
async function _extractAttachmentText(dataUrl: string, mime: string, name: string): Promise<string> {
  const key = Deno.env.get("OPENAI_API_KEY"); if (!key || !dataUrl) return "";
  const m = String(mime || "").toLowerCase();
  // texto puro (txt/csv): decodifica direto, sem IA
  if (m.startsWith("text/") || /\.(txt|csv|md)$/i.test(name)) {
    try { const r = await fetch(dataUrl); const t = await r.text(); return t.slice(0, 6000); } catch { return ""; }
  }
  const sys = "Extraia e liste, em português, o conteúdo útil deste anexo para orientar uma IA de palavras-chave do Google Ads (termos, serviços, produtos, público, regiões, o que incluir ou evitar). Seja objetivo, só o conteúdo relevante, sem preâmbulo.";
  const parts: any[] = [{ type: "text", text: "Anexo: " + (name || "arquivo") }];
  if (m.startsWith("image/")) parts.push({ type: "image_url", image_url: { url: dataUrl } });
  else parts.push({ type: "file", file: { filename: name || "documento.pdf", file_data: dataUrl } }); // PDF e afins
  try {
    const j = await callOpenAI({ model: "gpt-4o", messages: [{ role: "system", content: sys }, { role: "user", content: parts }], max_tokens: 700, temperature: 0.2 });
    return (j.choices?.[0]?.message?.content || "").trim();
  } catch (e) { return "erro ao ler anexo: " + ((e as any)?.message || e); }
}
// baixa (descriptografa) a mídia de uma mensagem no uazapi → { url, mime }
async function waMediaUrl(host: string, token: string, msgid: string): Promise<{ url: string; mime: string } | null> {
  try {
    const d = await waCall(host, token, "/message/download", "POST", { id: msgid });
    const b = (d && d.j) || {}; const url = b.fileURL || b.url; if (!url) return null;
    return { url: String(url), mime: String(b.mimetype || "") };
  } catch { return null; }
}
// Avisa no proprio grupo que a mensagem nao pode ser processada, dizendo o que fazer.
async function _waAgentAvisaErro(w: any, motivo: string) {
  const inst = (await sbGet("wa_instances", `id=eq.${encodeURIComponent(w.instanceId)}&select=uaz_host,uaz_token`))[0];
  if (!inst || !w.chatid) return;
  const cota = /limite do plano|cota|quota|RESOURCE_EXHAUSTED|429/i.test(motivo);
  const texto = cota
    ? `⚠️ Não consegui processar sua mensagem agora: *a IA atingiu o limite do plano*.\n\nSua mensagem não foi perdida — é só reenviar depois. Se estiver acontecendo direto, a chave da IA precisa de faturamento ativo ou de uma chave reserva.`
    : `⚠️ Não consegui processar sua mensagem agora.\n\n_${String(motivo).slice(0, 220)}_\n\nPode reenviar em instantes.`;
  await waCall(inst.uaz_host, inst.uaz_token, "/send/text", "POST", { number: w.chatid, text: texto });
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
  // ===== Mídia: áudio (transcreve), imagem (visão) e PDF (leitura) =====
  const visionParts: any[] = [];
  if (w.mtype && w.msgid) {
    const media = await waMediaUrl(inst.uaz_host, inst.uaz_token, w.msgid);
    if (media && media.url) {
      const mime = media.mime || w.mime || "";
      if (w.mtype === "audio") {
        const tr = await waTranscribe(media.url);
        if (tr) w.text = ((w.text || "") + " " + tr).trim();
        else { await send("🎧 Recebi seu áudio, mas não consegui transcrever agora. Pode mandar por texto?"); return { ok: true }; }
      } else if (w.mtype === "image") {
        visionParts.push({ type: "image_url", image_url: { url: media.url } });
        if (!(w.text || "").trim()) w.text = "O usuário enviou uma imagem. Analise o que há nela e responda de forma útil.";
      } else if (w.mtype === "document") {
        const b64 = await waFetchB64(media.url);
        if (b64) visionParts.push({ type: "file", file: { filename: w.fname || "documento.pdf", file_data: `data:${mime || "application/pdf"};base64,${b64}` } });
        else { await send(`📄 Recebi o documento${w.fname ? ` "${w.fname}"` : ""}, mas não consegui abrir agora.`); return { ok: true }; }
        if (!(w.text || "").trim()) w.text = `O usuário enviou um documento${w.fname ? ` ("${w.fname}")` : ""}. Leia e responda/resuma o que for útil.`;
      }
    }
  }
  const text = (w.text || "").trim(); if (!text && !visionParts.length) return { skip: true };
  const saveSess = async (patch: any) => { const row = { phone: skey, updated_at: new Date().toISOString(), ...patch }; if (sess) await sbPatchD("wa_agent_sessions", `phone=eq.${encodeURIComponent(skey)}`, row); else await sbPost("wa_agent_sessions", row); };
  // Espera pelo NOME da campanha (depois de "qual campanha?") — determinístico, não depende do modelo
  if (sess && sess.pending && sess.pending._awaitCampaign) {
    const tl0 = text.toLowerCase().replace(/[\s.!,]+/g, " ").trim();
    if (/^(nao|n|cancela|cancelar|deixa|esquece|para|negativo)$/.test(tl0)) { await saveSess({ pending: null, last_msgid: w.msgid }); await send("Ok, cancelei 👍"); return { ok: true }; }
    // ESCAPE: a espera só vale por 10 min e só pra mensagem curta que pareça nome de campanha.
    // Sem isso a AndréIA ficava presa e lia QUALQUER pedido novo como "nome da campanha".
    const _velho = sess.updated_at ? (Date.now() - new Date(sess.updated_at).getTime()) > 10 * 60000 : false;
    const _pedidoNovo = text.length > 60 || /^(me d|qual|quais|como|quanto|mostra|manda|gera|cria|resumo|relat|kpi|analis|status|preciso|quero)/i.test(text.trim());
    if (_velho || _pedidoNovo) { await saveSess({ pending: null }); sess.pending = null; } // solta e segue pro fluxo normal
  }
  if (sess && sess.pending && sess.pending._awaitCampaign) {
    const tl0 = text.toLowerCase().replace(/[\s.!,]+/g, " ").trim();
    const cidp = sess.pending.client_id || null;
    const camp = cidp ? await waResolveCampaign(cidp, text) : null;
    if (camp) {
      const pending = { ...sess.pending, campanha: camp.nome }; delete pending._awaitCampaign;
      const cls = await sbGet("clients", "select=id,name&limit=1000");
      const confirm = _waConfirmText(pending, cls);
      await saveSess({ pending, last_msgid: w.msgid });
      await send(confirm); return { ok: true };
    }
    await saveSess({ last_msgid: w.msgid });
    await send(`Não achei a campanha "${text}" nesse cliente. Copia o nome EXATO de uma da lista, por favor.`);
    return { ok: true };
  }
  if (sess && sess.pending && !sess.pending._awaitCampaign) {
    const tl = text.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[\s.!,]+/g, " ").trim();
    const isYes = /^(sim|s|ss|isso|isso mesmo|pode|pode sim|confirmo|confirmado|confirmar|ok|okay|blz|beleza|claro|positivo|certo|faz|fazer|manda|manda ver|vai|bora|com certeza|👍|✅)$/.test(tl);
    const isNo = /^(nao|n|cancela|cancelar|deixa|esquece|para|negativo|nem|melhor nao)$/.test(tl);
    if (isYes) {
      const p = sess.pending;
      const msg = await waAgentExec(p, sess.client_id, { host: inst.uaz_host, token: inst.uaz_token, number: dest });
      // grava no HISTÓRICO o que foi executado (senão a IA "esquece" e repete a etapa) + atualiza o estado do processo
      const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: `[EXECUTADO: ${p.tipo}] ${msg}` }].slice(-16);
      await saveSess({ pending: null, last_msgid: w.msgid, history: hist, flow: _flowAfter(p, (sess && sess.flow) || null) });
      await send(msg); return { ok: true };
    }
    if (isNo) { const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: "[CANCELADO pelo usuário]" }].slice(-16); await saveSess({ pending: null, last_msgid: w.msgid, history: hist }); await send("Ok, cancelei 👍"); return { ok: true }; }
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
  const unified = await _andreiaUnifiedContext((sess && sess.client_id) || null, "WhatsApp da equipe");
  const sys = `Você é a AndréIA, gestora de tráfego E financeiro, num grupo de WhatsApp com a equipe da agência. Fale CURTO, direto e natural (é WhatsApp). Ao AVALIAR/RECOMENDAR, siga sempre o PLAYBOOK acima (ex: custo por lead/conversa alto → orientar a verificar a QUALIFICAÇÃO antes de mandar reduzir custo).
- Você CONSULTA os dados reais do sistema com as ferramentas: consultar_banco (qualquer tabela: financeiro, tarefas, CRM, RD, pedidos, clientes…), meta_insights e google_insights (métricas ao vivo), resumo_todos_clientes. SEMPRE busque o dado real antes de responder — NUNCA invente número nem use placeholders (X, Y, Z). Se não houver dado, diga que não há.
- Traga SÓ o que tem dado, e a métrica do OBJETIVO do cliente. O snapshot já traz o campo 'objetivo' e só as métricas certas dele: venda→compras/ROAS/CPA; leads→leads/CPL; mensagens→conversas/custo por conversa; tráfego→cliques/CTR/CPC. NUNCA misture (ex: cliente de VENDA não mostra "custo por conversa").
- Formato WhatsApp: NÃO use markdown de título (nada de ### ou **). Negrito é com UM asterisco (*assim*). Listas com "• ". Seja enxuta.
- ATALHOS que a equipe pode pedir: "quem precisa de atenção?" → use resumo_todos_clientes e destaque os clientes abaixo da meta, com gasto sem resultado, ou parados; "saúde da carteira" → visão geral (gasto total do período, quantos performando/abaixo, e financeiro a receber/pagar via a ferramenta financeiro); "pendências operacionais" → tarefas em aberto (consultar_banco tabela tasks, filtro status=neq.done, ordena por due); "recomendações da semana" → 2-3 ações priorizadas (o que pausar/escalar/ajustar) com base nos dados. Sempre com dado real, curto.
- RESUMO PADRÃO de um cliente (meta_insights/google_insights, quando a pessoa só pede "resumo"/"como está"/"resultados"): por canal mostre SÓ a linha do campo **'kpi' VERBATIM** do consolidado (Gasto · Resultado · CPR pelo objetivo do canal) e, abaixo, cada campanha com o SEU 'kpi' verbatim. **NUNCA** liste métricas soltas (Impressões, CTR, CPC, CPM, Alcance, Conversas, Compras, Leads) — só Gasto · Resultado · CPR. Mostre APENAS campanhas e canais que TIVERAM GASTO no período; se vier 'semGastoNoPeriodo', diga só que o canal não teve gasto no período (não invente 0s). Se tiver 'orcamentoDiario' pode citar junto.
- FORMATO CUSTOMIZADO/ESPECÍFICO: se a pessoa pedir um formato próprio (ex.: "nesse formato:", colar um exemplo, ou pedir métricas específicas por nome — impressões, alcance, cliques, CTR, CPM, saldo) NÃO se limite ao 'kpi' verbatim: cada campanha/conjunto do payload de meta_insights já traz os campos crus (impressoes, cliques, alcance, ctrPct, cpm, cpc, spend) — USE-OS. Nunca diga que não tem o dado ou que "o sistema consolida só pelo objetivo" — isso não é mais verdade, os números crus vêm sempre no payload. Se o formato pedido tiver quebra geográfica/por região que a API não te dá pronta (ex: "Rio" vs "Niterói"), some os conjuntos/campanhas cujo NOME bata com cada região (o gestor já usa esse padrão de nomear campanhas por praça).
- SALDO DE CARTEIRA / saldo restante: use a tool meta_saldo (retorna 'saldoTotal' e o saldo por conta). NUNCA diga "verifique direto no gerenciador" — a informação está disponível pela tool.
- DETALHAR por conjunto/anúncio: no resumo padrão chame meta_insights SEM 'nivel' (ou nivel='campanha') — vem SÓ campanhas, NUNCA mostre conjuntos/anúncios. SÓ passe nivel='conjunto' (ou nivel='anuncio') quando o usuário pedir EXPLICITAMENTE ("detalhe/abre a campanha X", "por conjunto", "por anúncio", "quais anúncios"); aí cada campanha traz 'conjuntos' (com 'kpi', e 'anuncios' se nivel='anuncio') — liste-os com o 'kpi' verbatim de cada. Se o payload não trouxer 'conjuntos', é porque não foi pedido detalhe: fique no nível de campanha.
- google_insights (Google Ads): use o campo **'kpi' VERBATIM** do consolidado (Gasto · Resultado · CPR pelo objetivo do Google — ex: Vídeo→Views/Custo por view; Pesquisa→Compras/CPA; Display→Alcance/CPM) e, se listar campanhas, o 'kpi' de cada uma. NÃO despeje uma lista de Impressões/Cliques/CTR/CPC/Conversas/Compras/Leads todos juntos — mostre só o resultado do OBJETIVO daquele canal.
- (assist.): vendas/compras/ROAS/CPA POR CAMPANHA vêm do GERENCIADOR (pixel), não da planilha — ao mostrá-las escreva "(assist.)" ao lado do número (ex: "Compras 12 (assist.) · ROAS 3,1 (assist.)"), porque a venda REAL da agência vem da planilha e o pixel não divide venda por campanha. No consolidado do cliente que usa planilha, a venda é a real (sem "(assist.)").
- Para AÇÕES (criar tarefa, criar/cancelar reunião na agenda, pausar/reativar/duplicar campanha, orçamento, criar lançamento, dar baixa) use preparar_acao. Reunião: passe o título em 'nome', o dia em 'quando' (AAAA-MM-DD) e o horário em 'hora' (HH:MM) se disser. Cliente é opcional em reunião. — o sistema pede confirmação (SIM) e executa. NUNCA diga que já executou por conta própria. Se a mensagem citar um cliente ("no cliente X", "pro X"), passe o nome EXATO em 'cliente' — NUNCA reaproveite o cliente de mensagens anteriores quando a atual cita outro. Se o cliente não existir, o sistema avisa.
- CRIAR TAREFA: o RESPONSÁVEL (quem vai fazer) e a DATA são OBRIGATÓRIOS. Se o usuário não informar os dois, PERGUNTE (não invente responsável nem data, não assuma você mesma). Passe 'responsavel' (nome da pessoa) e 'quando' já como data ISO AAAA-MM-DD — calcule "amanhã", "hoje", "sexta" a partir de hoje. Urgência e revisão são OPCIONAIS: só marque urgente=true se a mensagem disser explicitamente que é urgente, e revisao=true só se pedir revisão. **NUNCA pergunte se é urgente ou se precisa de revisão** — se a pessoa não citou, crie a tarefa normalmente SEM essas flags e sem perguntar nada sobre isso. (E não escreva "urgente"/"revisão" no título/obs — são só flags.)
- 🆕 PROCESSO DE CLIENTE NOVO (inicialização de contrato — processo FECHADO, sempre nesta ordem, UMA etapa por vez, SEMPRE confirmando antes de cada uma):
  Quando disserem "cliente novo" (ou fechamos um cliente/contrato) + dados, PRIMEIRO pergunte: "Quer que eu inicie o processo de cliente novo? (contrato → cadastro → tarefas → financeiro)". Confirmado, siga:
  1) CONTRATO: junte os dados do CONTRATANTE (razão social/nome, CNPJ ou CPF, endereço, e-mail, telefone, representante e CPF dele) + condições (honorário mensal, meses promocionais se houver, fidelidade). Pergunte SÓ o que faltar em UMA mensagem; se a pessoa disser que não tem, siga assim mesmo (fica em branco no PDF). Aí chame preparar_acao tipo=gerar_contrato com os campos — o sistema gera o PDF e envia aqui.
  2) CADASTRO: depois do contrato, pergunte se quer cadastrar o cliente no sistema. SEMPRE pergunte primeiro **com que NOME ele deve aparecer no sistema** (nome fantasia/social) — pode ser DIFERENTE da razão social do contrato; ofereça o nome do contrato como padrão ("uso X ou prefere outro?"). Peça também nicho, verba de mídia e canais; sem esses dados, cadastre mesmo assim (o sistema avisa pra completar depois). preparar_acao tipo=cadastrar_cliente {nomeSistema (o nome escolhido pro sistema), razaoSocial (a do contrato, pra registro), nicho, verba, valorMensal, canais}.
  3) TAREFAS: pergunte se cria as tarefas de onboarding e CONFIRME os CANAIS que o cliente vai trabalhar (meta/google/tiktok) + o responsável — as tarefas criadas são SÓ dos canais confirmados. preparar_acao tipo=criar_tarefas_onboarding {cliente, canais, responsavel}.
  4) FINANCEIRO: por fim pergunte se quer criar o lançamento (fee mensal) — preparar_acao tipo=criar_lancamento. Se disser não, encerre com um resumo do que foi feito.
  Nunca pule etapa nem execute duas de uma vez; cada etapa passa pela confirmação (SIM) do sistema.
- DINHEIRO/FINANCEIRO: para QUALQUER pergunta de valores (a receber, a pagar, recebido, pago, fluxo do mês) use a ferramenta **financeiro** — ela já devolve o TOTAL correto e os ITENS com o nome certo do cliente. "a receber" = {tipo:'receita',status:'pendente'}; "a pagar" = {tipo:'despesa',status:'pendente'}; "este mês" = o mês atual, informado no contexto abaixo. NUNCA some você mesma nem adivinhe o nome do cliente — use os campos 'total' e 'itens' que a ferramenta retorna, exatamente.
- 🗣 VOCÊ É CONVERSACIONAL, não um robô de script: LEIA o histórico antes de responder. Mensagens marcadas "[EXECUTADO: x]" no histórico são coisas que VOCÊ JÁ FEZ — nunca proponha de novo. Se a pessoa responder algo fora do "sim/não" (informar um dado, corrigir você, dizer "isso você já fez", "já fizemos isso", "pula essa"), ENTENDA e siga em frente: reconheça em 1 linha e vá pro próximo passo pendente, sem repetir a confirmação anterior. Se estiver em dúvida sobre onde parou, PERGUNTE ("já gerei o contrato — quer que eu siga pro cadastro?") em vez de repetir a etapa.`;
  /* CACHE DE PROMPT: o desconto da OpenAI só vale pro PREFIXO idêntico entre chamadas. Por isso o
     bloco de instruções acima não tem nenhuma interpolação, e tudo que muda — contexto do cliente,
     playbook, data de hoje, lista de clientes e estado do fluxo — vai numa SEGUNDA mensagem. Antes o
     contexto vinha ANTES das instruções, o que jogava o prompt inteiro pra fora do cache. */
  const sysDin = `${unified}

${pb}

- Hoje é ${new Date().toISOString().slice(0, 10)} e o mês atual é ${new Date().toISOString().slice(0, 7)}.
- Ao filtrar por um cliente específico use o id dele (entre colchetes). Clientes: ${clients.slice(0, 150).map((c: any) => `${c.name}[${c.id}]`).join(" | ")}.${_flowPrompt(sess && sess.flow)}`;
  const hist0 = ((sess && sess.history) || []).slice(-8).map((h: any) => ({ role: h.role === "assistant" ? "assistant" : "user", content: h.text }));
  const userContent: any = visionParts.length ? [{ type: "text", text: text || "" }, ...visionParts] : text;
  const messages: any[] = [{ role: "system", content: sys }, { role: "system", content: sysDin }, ...hist0, { role: "user", content: userContent }];
  const agentModel = "gpt-4o"; // conversacional de verdade (segue o histórico e o estado do processo); também lê imagem/PDF
  let clientId = (sess && sess.client_id) || null;
  for (let it = 0; it < 6; it++) {
    const j = await callOpenAI({ model: agentModel, messages, tools: WA_TOOLS, tool_choice: "auto", max_tokens: 900, temperature: 0.3, _telemetry: { clientId, action: "andreia_whatsapp" } });
    const msg = j.choices[0].message;
    if (msg.tool_calls && msg.tool_calls.length) {
      messages.push(msg);
      let acted = false;
      for (const tc of msg.tool_calls) {
        let args: any = {}; try { args = JSON.parse(tc.function.arguments || "{}"); } catch { args = {}; }
        if (tc.function.name === "preparar_acao") {
          let cid = clientId; let reply = "";
          // TRAVA: etapa do processo de cliente novo já concluída → não repete; empurra pro próximo passo
          const _fl = (sess && sess.flow) || null; const _stp = _FLOW_STEPS[args.tipo];
          if (_fl && _fl.feito && _stp && _fl.feito[_stp]) {
            const ordem = [["contrato", "cadastrar o cliente no sistema"], ["cadastro", "criar as tarefas de onboarding"], ["tarefas", "criar o lançamento financeiro (fee)"], ["financeiro", ""]];
            const prox = (ordem.find(([k]) => k === _stp) || [])[1];
            const done: Record<string, string> = { contrato: "o contrato já foi gerado", cadastro: "o cliente já foi cadastrado", tarefas: "as tarefas já foram criadas", financeiro: "o lançamento já foi criado" };
            const msg = `Isso já está feito ✅ — ${done[_stp]}${_fl.nome ? ` (${_fl.nome})` : ""}.${prox ? ` Quer que eu siga pra ${prox}?` : " O processo de cliente novo está completo."}`;
            const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: msg }].slice(-16);
            await saveSess({ pending: null, last_msgid: w.msgid, history: hist });
            await send(msg); return { ok: true };
          }
          if (args.cliente) {
            const rc = _waResolveClient(args.cliente, clients);
            if (rc) cid = rc.id;
            else reply = `🤔 Não achei o cliente "${args.cliente}" no sistema. Confere o nome pra mim? (se quiser, peço a lista de clientes)`;
          }
          const semCliente = args.tipo === "criar_reuniao" || args.tipo === "cancelar_reuniao" || args.tipo === "gerar_contrato" || args.tipo === "cadastrar_cliente" || args.tipo === "criar_tarefas_onboarding";
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
          // CAMPANHA: resolve o nome EXATO antes de confirmar; se veio vazia/errada, lista as opções em vez de confirmar em branco
          let _awaitCamp = false;
          if (!reply && _waAcaoLabel[args.tipo] && cid) {
            const camp = (args.campanha && String(args.campanha).trim()) ? await waResolveCampaign(cid, args.campanha) : null;
            if (camp) args.campanha = camp.nome; // nome exato → confirmação mostra certo
            else { reply = await _waCampaignPickText(cid, args.tipo); _awaitCamp = true; } // guarda estado de espera pela campanha
          }
          if (!reply) { const pending = { ...args, client_id: cid }; reply = _waConfirmText(pending, clients); const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: reply }].slice(-16); await saveSess({ client_id: cid, pending, last_msgid: w.msgid, history: hist }); }
          else if (_awaitCamp) { const pending = { ...args, client_id: cid, _awaitCampaign: true }; const hist = [...((sess && sess.history) || []), { role: "user", text }, { role: "assistant", text: reply }].slice(-16); await saveSess({ client_id: cid, pending, last_msgid: w.msgid, history: hist }); }
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
    await _andreiaMaybeRemember(clientId, "WhatsApp da equipe", text);
    await send(reply); return { ok: true };
  }
  await send("Me embananei aqui 😅 pode reformular?");
  return { ok: true };
}
async function waHandler(w: any) {
  if (w.op === "extract") return await waExtract(w.convId, w.autoApply !== false);
  if (w.op === "capi") return await waCapi(w.convId, w.event);
  // CRM manual: cria apenas o vínculo interno necessário para importar e classificar
  // conversas, sem abrir uma instância no provedor de WhatsApp.
  if (w.op === "ensureManual") {
    const clientId = String(w.clientId || "").trim();
    if (!clientId) throw new Error("Selecione um cliente para ativar o CRM manual.");
    const client = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=id,name&limit=1`))[0];
    if (!client) throw new Error("Cliente não encontrado.");
    const current = await sbGet("wa_instances", `client_id=eq.${encodeURIComponent(clientId)}&select=id,client_id,name,status,provider,uaz_host,uaz_token,phone&order=created_at.asc&limit=20`);
    // Se o número já foi conectado, usa a conexão real. Caso contrário, reaproveita
    // o vínculo manual já existente para não fragmentar o histórico do cliente.
    const existing = current.find((x: any) => String(x.provider || "uazapi").toLowerCase() !== "manual") || current.find((x: any) => String(x.provider || "").toLowerCase() === "manual");
    if (existing) return { instance: existing, created: false, manual: String(existing.provider || "").toLowerCase() === "manual" };
    const id = _wuid();
    const row = { id, client_id: clientId, name: `crm-manual-${id.slice(0, 8)}`, uaz_token: "", uaz_host: "", phone: "", status: "manual", provider: "manual" };
    const saved = await sbInsertOk("wa_instances", row);
    if (!saved.ok) throw new Error("Não foi possível ativar o CRM manual: " + saved.err);
    return { instance: row, created: true, manual: true };
  }
  if (w.op === "revokeConsent") {
    const instanceId = String(w.instanceId || ""); if (!instanceId) throw new Error("instanceId obrigatório");
    await sbPatchD("wa_qr_consents", `instance_id=eq.${encodeURIComponent(instanceId)}&revoked_at=is.null`, { revoked_at: new Date().toISOString() });
    const ri = (await sbGet("wa_instances", `id=eq.${encodeURIComponent(instanceId)}&select=id,uaz_host,uaz_token`))[0]; if (!ri) throw new Error("Instância não encontrada.");
    try { await waCall(ri.uaz_host, ri.uaz_token, "/instance/logout", "POST", {}); } catch (_e) {}
    await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(instanceId)}`, { status: "disconnected", updated_at: new Date().toISOString() });
    return { revoked: true, disconnected: true };
  }
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
  const inst = (await sbGet("wa_instances", `id=eq.${encodeURIComponent(w.instanceId)}&select=id,client_id,uaz_host,uaz_token,phone,provider,status`))[0];
  if (!inst) throw new Error("Instância WhatsApp não encontrada.");
  const host = inst.uaz_host, token = inst.uaz_token, clientId = inst.client_id || null;
  const clientFilter = clientId ? "eq." + encodeURIComponent(clientId) : "is.null";
  /* Uma empresa pode ter MAIS DE UM numero. A conversa pertence ao numero por onde entrou: o mesmo lead
     falando nas duas linhas tem uma conversa em cada, e a resposta sai pela linha certa. Conversa antiga
     (de antes do multi-numero, sem instance_id) e adotada por este numero em vez de virar duplicata. */
  const _acharConversa = async (chat: string, campos = "id,name,origin_type") => {
    const base = `client_id=${clientFilter}&chat_id=eq.${encodeURIComponent(chat)}`;
    let r = (await sbGet("wa_conversations", `${base}&instance_id=eq.${encodeURIComponent(inst.id)}&select=${campos}&limit=1`))[0];
    if (!r) {
      const legado = (await sbGet("wa_conversations", `${base}&instance_id=is.null&select=${campos}&limit=1`))[0];
      if (legado) { await sbPatchD("wa_conversations", `id=eq.${legado.id}`, { instance_id: inst.id }); r = legado; }
    }
    return r;
  };
  if (w.op === "listLeadImports") {
    if (!clientId) return { imports: [] };
    const imports = await sbGet("crm_import_batches", `client_id=eq.${encodeURIComponent(clientId)}&select=*&order=created_at.desc&limit=50`);
    return { imports };
  }
  if (w.op === "deleteLeadImport") {
    if (!clientId || !w.batchId) throw new Error("Lote de importação inválido.");
    const batch = (await sbGet("crm_import_batches", `id=eq.${encodeURIComponent(w.batchId)}&client_id=eq.${encodeURIComponent(clientId)}&status=eq.active&select=*&limit=1`))[0];
    if (!batch) throw new Error("Arquivo importado não encontrado ou já excluído.");
    const convs = await sbGet("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&import_batch_id=eq.${encodeURIComponent(batch.id)}&select=id&limit=6000`);
    const ids = convs.map((x: any) => String(x.id)); const protectedIds = new Set<string>();
    for (let i = 0; i < ids.length; i += 200) { const chunk = ids.slice(i, i + 200); if (!chunk.length) continue; const msgs = await sbGet("wa_messages", `conversation_id=in.(${chunk.map((x: string) => encodeURIComponent(x)).join(",")})&select=conversation_id&limit=5000`); (msgs || []).forEach((x: any) => protectedIds.add(String(x.conversation_id))); }
    const deletable = ids.filter((id: string) => !protectedIds.has(id));
    for (let i = 0; i < deletable.length; i += 200) { const chunk = deletable.slice(i, i + 200); const r = await fetch(`${_SB_URL}/rest/v1/wa_conversations?id=in.(${chunk.map((x: string) => encodeURIComponent(x)).join(",")})`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, Prefer: "return=minimal" } }); if (!r.ok) throw new Error(`Falha ao remover leads do lote (${r.status}).`); }
    await sbPatchD("crm_import_batches", `id=eq.${encodeURIComponent(batch.id)}`, { status: "deleted", deleted_count: deletable.length, protected_count: protectedIds.size, deleted_at: new Date().toISOString() });
    return { deleted: deletable.length, protected: protectedIds.size };
  }
  if (w.op === "importLeads") {
    if (!clientId) throw new Error("Selecione o CRM de um cliente para importar leads.");
    const input = Array.isArray(w.leads) ? w.leads.slice(0, 5000) : [];
    if (!input.length) throw new Error("CSV sem leads válidos.");
    const cfg = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=crm_config&limit=1`))[0]?.crm_config || {};
    const stages = (Array.isArray(cfg.stages) && cfg.stages.length) ? cfg.stages : CRM_DEFAULT_STAGES;
    const normStage = (v: any) => { const q = String(v || "").trim().toLowerCase(); const found = stages.find((s: any) => String(s.key || "").toLowerCase() === q || String(s.label || "").trim().toLowerCase() === q); return found?.key || stages[0]?.key || "sem"; };
    const hash = (s: string) => { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); };
    const clean = input.map((x: any) => { const phone = String(x.phone || "").replace(/[^0-9]/g, ""); const email = String(x.email || "").trim().toLowerCase(); const chat = phone.length >= 8 ? phone : (email ? `email_${hash(email)}` : ""); let entered = new Date(x.date || ""), last = new Date(x.lastDate || x.date || ""); if (isNaN(entered.getTime())) entered = new Date(); if (isNaN(last.getTime())) last = entered; const originText = String(x.source || "").trim(), relevance = String(x.relevance || "").toLowerCase(); const fields: any = {}; if (email) fields.email = email; if (x.product) fields.produto = String(x.product).slice(0, 500); if (x.value) fields.valor = String(x.value).slice(0, 120); if (x.saleStatus) fields.status_venda = String(x.saleStatus).slice(0, 160); if (x.adUrl) fields.url_anuncio = String(x.adUrl).slice(0, 1000); return { chat, phone, email, name: String(x.name || phone || email || "Lead importado").slice(0, 160), stage: normStage(x.stage), entered: entered.toISOString(), at: last.toISOString(), fields, msgCount: Math.max(0, Number(x.messages) || 0), responseTime: Math.max(0, Number(x.responseTime) || 0), numErrado: /wrong|numero.*err|n[uú]mero.*err/i.test(relevance), irrelevante: /irrelevant|irrelevante/i.test(relevance), origin_type: /meta|google|ads|an[uú]ncio|tr[aá]fego|paid/i.test(originText) || x.campaign || x.ctwa ? "anuncio" : "organico", origin: { channel: originText.toLowerCase(), campaign: String(x.campaign || "").slice(0, 500), adset: String(x.adset || "").slice(0, 500), ad: String(x.ad || "").slice(0, 500), ctwa_clid: String(x.ctwa || "").slice(0, 1000), source_url: String(x.adUrl || "").slice(0, 1000), imported_from: "csv" } }; }).filter((x: any) => x.chat);
    const unique = new Map<string, any>(); clean.forEach((x: any) => { if (!unique.has(x.chat)) unique.set(x.chat, x); }); const rows = [...unique.values()];
    const existing = new Set<string>();
    /* "Ja existe" passa a ser por NUMERO: lead que ja esta na outra linha da mesma empresa nao impede a
       importacao nesta. Conversa sem numero (anterior ao multi-numero) continua contando como existente
       pra nao duplicar o que ja estava la. */
    for (let i = 0; i < rows.length; i += 200) { const ids = rows.slice(i, i + 200).map((x: any) => encodeURIComponent(x.chat)); const found = await sbGet("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&chat_id=in.(${ids.join(",")})&or=(instance_id.eq.${encodeURIComponent(inst.id)},instance_id.is.null)&select=chat_id`); (found || []).forEach((x: any) => existing.add(String(x.chat_id))); }
    const batchId = _wuid();
    const fresh = rows.filter((x: any) => !existing.has(x.chat)).map((x: any) => ({ id: _wuid(), client_id: clientId, chat_id: x.chat, instance_id: inst.id, name: x.name, stage: x.stage, fields: { ...x.fields, _csv_entrada: x.entered, _csv_mensagens: x.msgCount, _csv_tempo_resposta_s: x.responseTime }, last_text: "Lead importado via CSV", last_at: x.at, unread: 0, origin_type: x.origin_type, origin: x.origin, num_errado: x.numErrado, irrelevante: x.irrelevante, irrelevante_motivo: x.irrelevante ? "Marcado como irrelevante no CSV importado" : null, import_batch_id: batchId, import_source: "csv" }));
    for (let i = 0; i < fresh.length; i += 250) await sbPost("wa_conversations", fresh.slice(i, i + 250) as any);
    const duplicates = rows.length - fresh.length + (clean.length - rows.length), invalid = input.length - clean.length;
    await sbPost("crm_import_batches", { id: batchId, client_id: clientId, instance_id: inst.id, file_name: String(w.fileName || "Importação CSV").slice(0, 240), row_count: input.length, added_count: fresh.length, duplicate_count: duplicates, invalid_count: invalid, status: "active" });
    return { batchId, added: fresh.length, duplicates, invalid };
  }
  if (w.op === "importHistory") {
    const phone = String(w.phone || "").replace(/[^0-9]/g, "");
    const input = Array.isArray(w.messages) ? w.messages.slice(0, 10000) : [];
    if (!phone || phone.length < 8) throw new Error("Telefone do contato inválido.");
    if (!input.length) throw new Error("Arquivo sem mensagens reconhecidas.");
    const hash = (s: string) => { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); };
    const clean = input.map((x: any) => { const d = new Date(x.ts); return { ts: isNaN(d.getTime()) ? "" : d.toISOString(), text: String(x.text || "").slice(0, 12000), direction: x.direction === "out" ? "out" : "in", sender: String(x.sender || "") }; }).filter((x: any) => x.ts && x.text).sort((a: any, b: any) => a.ts.localeCompare(b.ts));
    if (!clean.length) throw new Error("Nenhuma mensagem válida no arquivo.");
    let conv = await _acharConversa(phone);
    let convId = conv?.id;
    const last = clean[clean.length - 1]; const contactName = String(w.contactName || phone).slice(0, 160);
    if (!convId) { convId = _wuid(); await sbPost("wa_conversations", { id: convId, client_id: clientId, chat_id: phone, instance_id: inst.id, name: contactName, last_text: last.text, last_at: last.ts, unread: 0, origin_type: "organico" }); }
    else { await sbPatchD("wa_conversations", `id=eq.${convId}`, { last_text: last.text, last_at: last.ts, ...(conv.name === phone && contactName !== phone ? { name: contactName } : {}) }); }
    const rows = clean.map((x: any) => { const mid = `hist_${hash(`${phone}|${x.ts}|${x.direction}|${x.text}`)}`; return { id: _wuid(), client_id: clientId, conversation_id: convId, chat_id: phone, instance_id: inst.id, wa_msgid: mid, direction: x.direction, msg_type: "text", text: x.text, ts: x.ts, raw: { imported: true, source: "whatsapp_export", sender: x.sender } }; });
    const known = new Set<string>();
    for (let i = 0; i < rows.length; i += 300) { const ids = rows.slice(i, i + 300).map((x: any) => x.wa_msgid); const ex = await sbGet("wa_messages", `wa_msgid=in.(${ids.map((x: string) => encodeURIComponent(x)).join(",")})&select=wa_msgid`); (ex || []).forEach((x: any) => known.add(x.wa_msgid)); }
    const fresh = rows.filter((x: any) => !known.has(x.wa_msgid));
    for (let i = 0; i < fresh.length; i += 300) await sbPost("wa_messages", fresh.slice(i, i + 300) as any);
    // Uma única classificação ao final: o histórico inteiro já está disponível para a IA.
    if (fresh.some((x: any) => x.direction === "in")) try { await waExtract(convId, true); } catch (_e) {}
    return { added: fresh.length, duplicates: rows.length - fresh.length, conversationId: convId };
  }
  if (String(inst.provider || "").toLowerCase() === "manual") {
    throw new Error("Este CRM está no modo manual. Importe um arquivo para incluir conversas; conecte o WhatsApp para sincronizar ou enviar mensagens.");
  }
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
  if (w.op === "disconnect") {
    try { await waCall(host, token, "/instance/logout", "POST", {}); } catch (_e) {}
    await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(inst.id)}`, { status: "disconnected", updated_at: new Date().toISOString() });
    return { disconnected: true };
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
      const conv = await _acharConversa(number, "id");
      let convId = conv?.id;
      if (!convId) { convId = _wuid(); await sbPost("wa_conversations", { id: convId, client_id: clientId, chat_id: number, instance_id: inst.id, name: number, last_text: w.text, last_at: ts, origin_type: "organico" }); }
      else await sbPatchD("wa_conversations", `id=eq.${convId}`, { last_text: w.text, last_at: ts });
      await sbPost("wa_messages", { id: _wuid(), client_id: clientId, conversation_id: convId, chat_id: number, instance_id: inst.id, wa_msgid: String((j && (j.id || j.messageid)) || _wuid()), direction: "out", msg_type: "text", text: w.text, ts, raw: j });
    }
    return { ok: status >= 200 && status < 300, status, resp: j };
  }
  if (w.op === "poll") {
    // sincronização (recupera histórico se o WhatsApp caiu / webhook falhou): sinceDays filtra o quanto puxar.
    // A API do provedor não tem filtro de data nativo — pedimos um lote grande e cortamos localmente pela timestamp
    // (a resposta vem ordenada da mais recente pra mais antiga, então paramos de processar assim que passar do corte).
    const sinceDays = Number(w.sinceDays) || 0;
    const sinceDate = /^\d{4}-\d{2}-\d{2}$/.test(String(w.sinceDate || "")) ? new Date(`${w.sinceDate}T00:00:00-03:00`).getTime() : 0;
    const untilDate = /^\d{4}-\d{2}-\d{2}$/.test(String(w.untilDate || "")) ? new Date(`${w.untilDate}T23:59:59-03:00`).getTime() : 0;
    const historical = !!(sinceDays || sinceDate || untilDate);
    const limit = historical ? Math.min(Math.max(Number(w.limit) || 3000, 500), 5000) : (w.limit || 60);
    const cutoff = sinceDate || (sinceDays ? Date.now() - sinceDays * 864e5 : 0);
    const { j } = await waCall(host, token, "/message/find", "POST", { limit });
    // a API devolve um array puro (não {messages:[...]}) — antes disso o parsing sempre resultava em [] e o poll não importava nada
    const msgsAll: any[] = Array.isArray(j) ? j : ((j && j.messages) || []);
    let msgs = msgsAll.filter((m) => !(m.isGroup || String(m.chatid || "").endsWith("@g.us")));
    if (cutoff) msgs = msgs.filter((m) => waTs(m.messageTimestamp) >= new Date(cutoff).toISOString());
    if (untilDate) msgs = msgs.filter((m) => waTs(m.messageTimestamp) <= new Date(untilDate).toISOString());
    const ids = msgs.map((m) => String(m.messageid || m.id || "")).filter(Boolean);
    const known = new Set<string>();
    for (let i = 0; i < ids.length; i += 300) { const chunk = ids.slice(i, i + 300); const ex = await sbGet("wa_messages", `wa_msgid=in.(${chunk.map((x) => encodeURIComponent(x)).join(",")})&select=wa_msgid`); (ex || []).forEach((r: any) => known.add(r.wa_msgid)); }
    const novas = msgs.filter((m) => { const id = String(m.messageid || m.id || ""); return id && !known.has(id); });
    // agrupa por telefone: 1 lookup/upsert de conversa por pessoa, não por mensagem (era o gargalo em lotes grandes)
    const byPhone: Record<string, any[]> = {};
    for (const m of novas) { const phone = String(m.chatid || m.sender_pn || m.sender || "").replace(/@.*$/, "").replace(/[^0-9]/g, ""); if (!phone) continue; (byPhone[phone] = byPhone[phone] || []).push(m); }
    const phones = Object.keys(byPhone);
    const existingMap: Record<string, any> = {};
    /* Traz instance_id junto: com duas linhas na empresa, a conversa DESTE numero manda; a conversa
       sem numero (anterior ao multi-numero) fica de reserva e e adotada mais abaixo. */
    for (let i = 0; i < phones.length; i += 200) { const chunk = phones.slice(i, i + 200); const ex = await sbGet("wa_conversations", `client_id=${clientFilter}&chat_id=in.(${chunk.map((x) => encodeURIComponent(x)).join(",")})&select=id,chat_id,origin_type,name,instance_id`); (ex || []).forEach((r: any) => { const atual = existingMap[r.chat_id]; if (!atual || (atual.instance_id !== inst.id && r.instance_id === inst.id)) existingMap[r.chat_id] = r; }); }
    const adCache: Record<string, Record<string, string> | null> = {};
    const newInbound = new Set<string>();
    const msgRows: any[] = [];
    let added = 0;
    for (const phone of phones) {
      const group = byPhone[phone].sort((a, b) => (a.messageTimestamp || 0) - (b.messageTimestamp || 0)); // mais antiga → mais nova
      let existing = existingMap[phone];
      // conversa do OUTRO numero da mesma empresa: nao e esta. Vira conversa nova, desta linha.
      if (existing && existing.instance_id && existing.instance_id !== inst.id) existing = undefined;
      let convId = existing?.id;
      const last = group[group.length - 1];
      const lastText = waText(last), lastTs = waTs(last.messageTimestamp);
      const anyInbound = group.some((m) => !m.fromMe);
      // origem: pega da 1ª mensagem inbound do lote (é a que carrega o clique/anúncio original)
      let origin: any = null;
      const firstInbound = group.find((m) => !m.fromMe);
      if (firstInbound) {
        origin = waOrigin(firstInbound);
        if (!origin) { const rf = await _waRefOrigin(waText(firstInbound)); if (rf) origin = rf; }
        if (!origin) { const ph = await waPhoneHistoryOrigin(clientId, phone); if (ph) origin = ph; }
        if (origin && origin.type === "anuncio" && origin.data.source_id && !origin.data.campaign) {
          const key = String(origin.data.source_id);
          if (adCache[key] === undefined) adCache[key] = await waResolveAd(key);
          if (adCache[key]) Object.assign(origin.data, adCache[key]);
        }
      }
      const nomeMsg = [...group].reverse().find((m) => m.senderName)?.senderName;
      if (!convId) { convId = _wuid(); await sbPost("wa_conversations", { id: convId, client_id: clientId, chat_id: phone, instance_id: inst.id, name: nomeMsg || phone, last_text: lastText, last_at: lastTs, unread: anyInbound ? 1 : 0, origin_type: origin ? origin.type : "organico", origin: origin ? origin.data : null }); }
      else { const patch: Record<string, unknown> = { last_text: lastText, last_at: lastTs }; if (anyInbound) patch.unread = 1; if (!existing.instance_id) patch.instance_id = inst.id; if (origin && (!existing.origin_type || existing.origin_type === "organico")) { patch.origin_type = origin.type; patch.origin = origin.data; } if (nomeMsg && (!existing.name || existing.name === phone || /^\d+$/.test(String(existing.name)))) patch.name = nomeMsg; await sbPatchD("wa_conversations", `id=eq.${convId}`, patch); }
      for (const m of group) { msgRows.push({ id: _wuid(), client_id: clientId, conversation_id: convId, chat_id: phone, instance_id: inst.id, wa_msgid: String(m.messageid || m.id || ""), direction: m.fromMe ? "out" : "in", msg_type: m.messageType || "text", text: waText(m), ts: waTs(m.messageTimestamp), raw: m }); added++; }
      if (anyInbound) newInbound.add(convId);
    }
    // grava as mensagens em lotes (era 1 insert por mensagem — lento demais pra sincronizar semanas de histórico)
    for (let i = 0; i < msgRows.length; i += 300) await sbPost("wa_messages", msgRows.slice(i, i + 300) as any);
    // classificação AUTOMÁTICA por IA das conversas que receberam nova mensagem do lead (limita p/ controlar custo)
    let classified = 0;
    if (newInbound.size) { for (const cid of [...newInbound].slice(0, 10)) { try { await waExtract(cid, true); classified++; } catch (_e) {} } }
    return { added, scanned: msgsAll.length, filtrados: msgs.length, classified, conversas: phones.length };
  }
  if (w.op === "resolveOrigins") {
    const convs = await sbGet("wa_conversations", `client_id=${clientFilter}&origin_type=eq.anuncio&select=id,origin`);
    // contas Meta pra busca por título (via C): do cliente da instância OU passadas pelo front (cliente ativo no CRM)
    const accIds: string[] = (Array.isArray(w.metaAccountIds) && w.metaAccountIds.length) ? w.metaAccountIds.map((x: any) => String(x).replace(/[^0-9]/g, "")).filter(Boolean) : [];
    const gAccIds: string[] = (Array.isArray(w.googleAccountIds) && w.googleAccountIds.length) ? w.googleAccountIds.map((x: any) => String(x).replace(/[^0-9]/g, "")).filter(Boolean) : [];
    const cache: Record<string, Record<string, string> | null> = {}; const titleCache: Record<string, Record<string, string> | null> = {}; let done = 0;
    for (const cv of (convs || [])) {
      const o = cv.origin || {};
      // Só considera resolvido quando campanha E grupo já deixaram de ser IDs do ValueTrack.
      // Antes, uma campanha com nome fazia o fluxo ignorar um adgroup ainda numérico.
      const rawGroup = String(o.adgroup || o.adset || "");
      const campaignResolved = !!o.campaign && !/^\d+$/.test(String(o.campaign));
      const groupResolved = !rawGroup || !/^\d+$/.test(rawGroup);
      if (campaignResolved && groupResolved) continue;
      let res: Record<string, string> | null = null;
      // (A) Meta: id do anúncio (source_id) → Graph direto
      if (o.source_id) { const key = String(o.source_id); if (cache[key] === undefined) cache[key] = await waResolveAd(key); res = cache[key]; }
      // (C) Meta: casar pelo TÍTULO do criativo nas contas do cliente
      if (!res && (o.title || o.body) && accIds.length) { const tk = String(o.title || "") + "␟" + String(o.body || ""); if (titleCache[tk] === undefined) titleCache[tk] = await waResolveAdByTitle(String(o.title || ""), accIds, String(o.body || "")); res = titleCache[tk]; }
      // Google: ValueTrack {campaignid}/{adgroupid} → nomes
      if (!res && o.channel === "google" && gAccIds.length) { const gid = /^\d+$/.test(rawGroup) ? rawGroup : ""; const gr = await waResolveGoogleCampaign(/^\d+$/.test(String(o.campaign || "")) ? String(o.campaign) : "", gid, gAccIds); if (gr) res = { campaign: gr.campaign || o.campaign || "", adset: gr.adgroup || o.adset || "", adgroup: gr.adgroup || o.adgroup || "" }; }
      if (res) { await sbPatchD("wa_conversations", `id=eq.${cv.id}`, { origin: { ...o, ...res } }); done++; }
    }
    return { resolved: done };
  }
  throw new Error("op inválida");
}

// Resolve origem (campanha › conjunto › anúncio) de anúncios de TODOS os clientes automaticamente,
// usando as contas Meta/Google de cada cliente. Chamado por cron — não depende de clique manual.
async function waResolveAllOrigins(): Promise<{ resolved: number; clients: number }> {
  // conversas de anúncio ainda não resolvidas (campanha vazia ou ainda em id numérico do ValueTrack)
  const convs = await sbGet("wa_conversations", "origin_type=eq.anuncio&select=id,client_id,origin&limit=2000");
  // ctwa_only (sem sourceId/título) não tem como resolver campanha pela API — não fica re-tentando
  const pend = (convs || []).filter((cv: any) => { const o = cv.origin || {}; if (o.ctwa_only && !o.source_id && !o.title) return false; const rawGroup = String(o.adgroup || o.adset || ""); return cv.client_id && (!o.campaign || /^\d+$/.test(String(o.campaign)) || /^\d+$/.test(rawGroup)); });
  if (!pend.length) return { resolved: 0, clients: 0 };
  // agrupa por cliente
  const byClient: Record<string, any[]> = {};
  for (const cv of pend) { (byClient[cv.client_id] = byClient[cv.client_id] || []).push(cv); }
  const cids = Object.keys(byClient);
  const clis = await sbGet("clients", `id=in.(${cids.map((x) => encodeURIComponent(x)).join(",")})&select=id,meta_account_id,google_account_id`);
  const cliMap: Record<string, any> = {}; (clis || []).forEach((c: any) => { cliMap[c.id] = c; });
  let done = 0;
  for (const cid of cids) {
    const c = cliMap[cid]; if (!c) continue;
    const accIds = String(c.meta_account_id || "").split(",").map((s: string) => s.trim().replace(/^act_/, "").replace(/[^0-9]/g, "")).filter(Boolean);
    const gAccIds = String(c.google_account_id || "").split(",").map((s: string) => s.trim().replace(/-/g, "").replace(/[^0-9]/g, "")).filter(Boolean);
    const cache: Record<string, Record<string, string> | null> = {}; const titleCache: Record<string, Record<string, string> | null> = {};
    for (const cv of byClient[cid]) {
      const o = cv.origin || {}; let res: Record<string, string> | null = null;
      if (o.source_id) { const key = String(o.source_id); if (cache[key] === undefined) cache[key] = await waResolveAd(key); res = cache[key]; }
      if (!res && (o.title || o.body) && accIds.length) { const tk = String(o.title || "") + "␟" + String(o.body || ""); if (titleCache[tk] === undefined) titleCache[tk] = await waResolveAdByTitle(String(o.title || ""), accIds, String(o.body || "")); res = titleCache[tk]; }
      if (!res && o.channel === "google" && gAccIds.length) { const rawGroup = String(o.adgroup || o.adset || ""); const gr = await waResolveGoogleCampaign(/^\d+$/.test(String(o.campaign || "")) ? String(o.campaign) : "", /^\d+$/.test(rawGroup) ? rawGroup : "", gAccIds); if (gr) res = { campaign: gr.campaign || o.campaign || "", adset: gr.adgroup || o.adset || "", adgroup: gr.adgroup || o.adgroup || "" }; }
      if (res) { await sbPatchD("wa_conversations", `id=eq.${cv.id}`, { origin: { ...o, ...res } }); done++; }
    }
  }
  return { resolved: done, clients: cids.length };
}


// Verifica se o pixel está mesmo instalado no site do cliente: procura no HTML e, se não achar,
// procura dentro do container do Google Tag Manager (instalação via GTM não aparece no HTML).
async function pixelCheck(m: any) {
  const clientId = String(m.clientId || "").trim();
  if (!clientId) throw new Error("clientId obrigatório");
  const cli = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=name,site_url`))[0] || {};
  const cfg = (await sbGet("tracking_config", `client_id=eq.${encodeURIComponent(clientId)}&select=token`))[0] || {};
  const token = cfg.token || "";
  let site = String(cli.site_url || "").trim();
  if (!site) return { ok: false, motivo: "sem_site", msg: "Cadastre o site do cliente (aba Relatórios) pra eu conseguir verificar." };
  if (!/^https?:\/\//i.test(site)) site = "https://" + site;
  if (!token) return { ok: false, motivo: "sem_token", msg: "Este cliente ainda não tem token de pixel." };
  const achou = (t: string) => t.includes(token) || t.includes("pixel.gt-marketing.app.br/pixel/script");
  let html = "";
  try {
    const r = await fetch(site, { headers: { "User-Agent": "Mozilla/5.0 (compatible; GTMarketingBot/1.0)" }, redirect: "follow" });
    html = await r.text();
  } catch (e) { return { ok: false, motivo: "site_off", msg: `Não consegui abrir ${site}: ${String((e as any).message || e).slice(0, 80)}` }; }
  if (achou(html)) return { ok: true, via: "html", site, msg: "Pixel encontrado direto no HTML do site." };
  // procura via Google Tag Manager
  const gtms = [...new Set((html.match(/GTM-[A-Z0-9]{4,}/g) || []))];
  for (const g of gtms.slice(0, 3)) {
    try {
      const rc = await fetch(`https://www.googletagmanager.com/gtm.js?id=${g}`);
      const js = await rc.text();
      if (achou(js)) return { ok: true, via: "gtm", gtm: g, site, msg: `Pixel encontrado no Google Tag Manager (${g}) — publicado e ativo.` };
    } catch (_e) { /* segue */ }
  }
  return { ok: false, motivo: "nao_encontrado", site, gtms,
    msg: gtms.length ? `Não achei o pixel no site nem no GTM (${gtms.join(", ")}). Se acabou de publicar no GTM, espere 1 min e verifique de novo.` : "Não achei o pixel no HTML do site. Cole o script antes do </head> ou instale via GTM." };
}
// ===== Google Analytics 4 + Search Console (mesma service account das planilhas) =====
// Camada AGREGADA: valida/complementa a jornada individual do nosso pixel (o GA4 não expõe usuário a usuário).
async function _gsaToken(scopes: string[]): Promise<string> {
  const keyJson = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_KEY");
  if (!keyJson) throw new Error("GOOGLE_SERVICE_ACCOUNT_KEY não configurada nos secrets");
  const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(keyJson), scopes });
  const client = await auth.getClient();
  const t: any = await client.getAccessToken();
  const tok = typeof t === "string" ? t : (t && t.token);
  if (!tok) throw new Error("não consegui autenticar a service account");
  return tok;
}
function _gsaEmail(): string { try { return JSON.parse(Deno.env.get("GOOGLE_SERVICE_ACCOUNT_KEY") || "{}").client_email || ""; } catch { return ""; } }

// ===== Login pessoal do Google (OAuth) — fallback quando a service account nao tem acesso a uma propriedade
// especifica (ex: GA4 de um cliente onde ninguem lembrou de compartilhar com a conta de servico, mas um membro
// da equipe ja tem acesso via o proprio login). O refresh_token fica no cofre (secure_credentials), igual ao
// token pessoal do Meta. E-mail/data de conexao ficam em account_config.data.google_oauth (sem segredo, so exibicao).
async function _googleUserRefreshToken(): Promise<string> {
  const r = await fetch(`${_SB_URL}/rest/v1/secure_credentials?id=eq.google_user_refresh_token&select=secret_cipher&limit=1`, { headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
  const rows = r.ok ? await r.json() : [];
  return rows[0]?.secret_cipher ? await _decryptCredential(rows[0].secret_cipher) : "";
}
async function _googleUserToken(scopes: string[]): Promise<string | null> {
  const clientId = Deno.env.get("GOOGLE_OAUTH_CLIENT_ID"), clientSecret = Deno.env.get("GOOGLE_OAUTH_CLIENT_SECRET");
  if (!clientId || !clientSecret) return null;
  const refresh = await _googleUserRefreshToken();
  if (!refresh) return null;
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, refresh_token: refresh, grant_type: "refresh_token" }),
  });
  const j = await r.json();
  return j.access_token || null;
}
// Tenta a service account primeiro; se o erro for de PERMISSAO e existir login pessoal conectado, tenta de novo com ele.
// fetchFn(token) faz a chamada e devolve o JSON já parseado (com .error quando falha, no formato do Google).
async function _googleTryTokens(scopes: string[], fetchFn: (token: string) => Promise<any>): Promise<any> {
  let j1: any = null, err1: unknown = null;
  try {
    const gsa = await _gsaToken(scopes);
    j1 = await fetchFn(gsa);
    if (!j1?.error) return j1;
    if (!/permission|PERMISSION|not have|403/i.test(j1.error.message || "")) return j1; // erro que nao e de permissao: nao adianta trocar de token
  } catch (e) { err1 = e; } // service account nem configurada/autenticou — ainda vale tentar o login pessoal
  const userTok = await _googleUserToken(scopes).catch(() => null);
  if (!userTok) { if (err1) throw err1; return j1; }
  const j2 = await fetchFn(userTok);
  if (!j2?.error) return j2;
  if (err1) throw err1; // login pessoal tambem falhou: mostra o erro original da service account, mais familiar
  return j1;
}
async function googleOAuthStart(authorization: string) {
  if (!authorization) throw new Error("Sessão obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } });
  if (!ur.ok) throw new Error("Sessão inválida.");
  const clientId = Deno.env.get("GOOGLE_OAUTH_CLIENT_ID");
  if (!clientId) throw new Error("GOOGLE_OAUTH_CLIENT_ID não configurada nos secrets — crie o OAuth Client no Google Cloud primeiro.");
  // state assinado (mesma criptografia do cofre — simetrica com a decriptação que o callback faz em tracking.ts)
  const state = await _encryptCredential(`google_oauth:${Date.now()}:${crypto.randomUUID()}`);
  const redirect = `${_SB_URL}/functions/v1/tracking/google/oauth/callback`;
  const scopes = ["https://www.googleapis.com/auth/analytics.readonly", "https://www.googleapis.com/auth/webmasters.readonly", "openid", "email"].join(" ");
  const url = `https://accounts.google.com/o/oauth2/v2/auth?${new URLSearchParams({ client_id: clientId, redirect_uri: redirect, response_type: "code", access_type: "offline", prompt: "consent", scope: scopes, state })}`;
  return { url };
}
async function googleOAuthDisconnect(authorization: string) {
  if (!authorization) throw new Error("Sessão obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } });
  if (!ur.ok) throw new Error("Sessão inválida.");
  await fetch(`${_SB_URL}/rest/v1/secure_credentials?id=eq.google_user_refresh_token`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}` } });
  const acc = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  delete acc.google_oauth;
  await sbPatchD("account_config", "id=eq.main", { data: acc });
  return { ok: true };
}
async function youtubeOAuth(input: any, authorization: string) {
  if (!authorization) throw new Error("Sessão obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } }); if (!ur.ok) throw new Error("Sessão inválida.");
  const clientId = String(input?.clientId || ""); if (!clientId) throw new Error("Cliente obrigatório.");
  if (input.op === "status") { const c = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=youtube_config&limit=1`))[0]; return { connected: !!c?.youtube_config?.channel_id, config: c?.youtube_config || {}, redirectUri: `${_SB_URL}/functions/v1/tracking/youtube/callback` }; }
  if (input.op === "disconnect") { await sbDeleteD("secure_credentials", `id=eq.${encodeURIComponent(`youtube_refresh_token:${clientId}`)}`); await sbPatchD("clients", `id=eq.${encodeURIComponent(clientId)}`, { youtube_config: {} }); return { ok: true }; }
  const cid = Deno.env.get("GOOGLE_OAUTH_CLIENT_ID"); if (!cid) throw new Error("GOOGLE_OAUTH_CLIENT_ID não configurado.");
  const state = await _encryptCredential(`youtube_oauth:${clientId}:${Date.now()}:${crypto.randomUUID()}`), redirect = `${_SB_URL}/functions/v1/tracking/youtube/callback`;
  const scope = ["https://www.googleapis.com/auth/youtube.readonly", "https://www.googleapis.com/auth/yt-analytics.readonly", "openid", "email"].join(" ");
  return { url: `https://accounts.google.com/o/oauth2/v2/auth?${new URLSearchParams({ client_id: cid, redirect_uri: redirect, response_type: "code", access_type: "offline", prompt: "consent", scope, state })}`, redirectUri: redirect };
}
function _normName(s: string): string { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim(); }
function _nameScore(a: string, b: string): number {
  const na = _normName(a), nb = _normName(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  if (na.includes(nb) || nb.includes(na)) return 0.85;
  const ta = new Set(na.split(" ")), tb = new Set(nb.split(" "));
  let common = 0; for (const t of ta) if (t.length > 2 && tb.has(t)) common++;
  return (common / (Math.max(ta.size, tb.size) || 1)) * 0.7;
}
// Lista as propriedades do GA4 que a conta de servico E o login pessoal conseguem ver (Admin API), casa por nome
// com os clientes que ainda nao tem ga4_property_id preenchido. NUNCA grava nada sozinho — so sugere; quem confirma
// o vinculo e o gestor, clicando (ver dados-clientes-inviolaveis: nao alteramos cadastro de cliente sem confirmacao).
async function googleGa4Discover(authorization: string) {
  if (!authorization) throw new Error("Sessão obrigatória.");
  const ur = await fetch(`${_SB_URL}/auth/v1/user`, { headers: { apikey: _SB_KEY, Authorization: authorization } });
  if (!ur.ok) throw new Error("Sessão inválida.");
  const scopes = ["https://www.googleapis.com/auth/analytics.readonly"];
  type Prop = { id: string; name: string; account: string; via: string };
  const props: Prop[] = []; const seen = new Set<string>();
  const pull = async (token: string | null, via: string) => {
    if (!token) return;
    try {
      const r = await fetch("https://analyticsadmin.googleapis.com/v1beta/accountSummaries?pageSize=200", { headers: { Authorization: `Bearer ${token}` } });
      const j = await r.json();
      for (const acc of (j.accountSummaries || [])) {
        for (const p of (acc.propertySummaries || [])) {
          const id = String(p.property || "").replace("properties/", "");
          if (!id || seen.has(id)) continue;
          seen.add(id);
          props.push({ id, name: p.displayName || id, account: acc.displayName || "", via });
        }
      }
    } catch (_e) { /* essa identidade pode nao ter Admin API habilitada — segue com a outra */ }
  };
  await pull(await _gsaToken(scopes).catch(() => null), "conta de serviço");
  await pull(await _googleUserToken(scopes).catch(() => null), "login pessoal");
  const clients = await _sbAll("clients", "select=id,name,ga4_property_id&status=neq.Encerrado");
  const semGa4 = clients.filter((c: any) => !c.ga4_property_id);
  const sugestoes: any[] = [];
  for (const c of semGa4) {
    let best: (Prop & { score: number }) | null = null;
    for (const p of props) {
      const score = Math.max(_nameScore(c.name, p.name), _nameScore(c.name, p.account));
      if (score >= 0.5 && (!best || score > best.score)) best = { ...p, score };
    }
    if (best) sugestoes.push({ clientId: c.id, clientName: c.name, propertyId: best.id, propertyName: best.name, account: best.account, via: best.via, confianca: Math.round(best.score * 100) });
  }
  const jaSugeridos = new Set(sugestoes.map((s) => s.clientId));
  const semAcesso = semGa4.filter((c: any) => !jaSugeridos.has(c.id)).map((c: any) => ({ id: c.id, name: c.name }));
  return { totalClientes: clients.length, jaConfigurados: clients.length - semGa4.length, propriedadesEncontradas: props.length, sugestoes: sugestoes.sort((a, b) => b.confianca - a.confianca), semAcesso };
}
async function ga4Report(m: any) {
  const prop = String(m.propertyId || "").replace(/[^0-9]/g, "");
  if (!prop) throw new Error("propertyId do GA4 obrigatório (só os números, ex: 123456789)");
  const run = async (dims: string[], mets: string[], limit = 50, orderMetric = "") => {
    const j = await _googleTryTokens(["https://www.googleapis.com/auth/analytics.readonly"], async (token) => {
      const r = await fetch(`https://analyticsdata.googleapis.com/v1beta/properties/${prop}:runReport`, {
        method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ dateRanges: [{ startDate: m.since, endDate: m.until }], dimensions: dims.map((name) => ({ name })), metrics: mets.map((name) => ({ name })), ...(orderMetric ? { orderBys: [{ metric: { metricName: orderMetric }, desc: true }] } : {}), limit }),
      });
      return r.json();
    });
    if (j.error) throw new Error(`GA4: ${j.error.message}${/permission|PERMISSION/i.test(j.error.message || "") ? ` — dê acesso de Leitor à ${_gsaEmail()} na propriedade ${prop} (ou conecte um login pessoal com acesso em Configurações)` : ""}`);
    return (j.rows || []).map((row: any) => {
      const o: any = {};
      (j.dimensionHeaders || []).forEach((h: any, i: number) => { o[h.name] = row.dimensionValues[i].value; });
      (j.metricHeaders || []).forEach((h: any, i: number) => { o[h.name] = Number(row.metricValues[i].value) || 0; });
      return o;
    });
  };
  const [canais, paginas, origens, campanhas, qualidadeCanais, qualidadePaginas, qualidadePalavras] = await Promise.all([
    run(["sessionDefaultChannelGroup"], ["sessions", "totalUsers", "conversions", "purchaseRevenue", "transactions", "engagedSessions", "averageSessionDuration", "userEngagementDuration", "screenPageViewsPerSession", "keyEvents"], 30, "sessions").catch((e) => ({ _err: String(e.message || e) } as any)),
    run(["pagePath"], ["screenPageViews", "totalUsers"], 25, "screenPageViews").catch(() => []),
    run(["sessionSource", "sessionMedium"], ["sessions", "conversions", "purchaseRevenue", "transactions"], 25, "sessions").catch(() => []),
    // campanha com RECEITA e COMPRAS — é o que liga o GA4 à jornada/atribuição
    run(["sessionCampaignName", "sessionSource", "sessionMedium"], ["sessions", "transactions", "purchaseRevenue", "conversions"], 40, "purchaseRevenue").catch(() => []),
    // Relatórios separados mantêm compatibilidade do GA4 e evitam que uma dimensão muito granular
    // altere os totais gerais. Estes três blocos alimentam a Qualidade da Navegação na Jornada.
    run(["sessionDefaultChannelGroup", "sessionSource", "sessionMedium", "sessionCampaignName"], ["sessions", "totalUsers", "engagedSessions", "engagementRate", "averageSessionDuration", "userEngagementDuration", "screenPageViewsPerSession", "keyEvents"], 800, "sessions").catch(() => []),
    run(["sessionSourceMedium", "pagePath", "pageTitle"], ["sessions", "totalUsers", "screenPageViews", "userEngagementDuration", "averageSessionDuration", "engagementRate", "keyEvents"], 300, "screenPageViews").catch(() => []),
    run(["sessionGoogleAdsKeyword", "sessionGoogleAdsQuery", "sessionGoogleAdsCampaignName"], ["sessions", "engagedSessions", "engagementRate", "averageSessionDuration", "userEngagementDuration", "screenPageViewsPerSession", "keyEvents"], 500, "sessions").catch(() => []),
  ]);
  if ((canais as any)._err) throw new Error((canais as any)._err);
  const tot = (arr: any[], k: string) => (Array.isArray(arr) ? arr : []).reduce((s: number, r: any) => s + (Number(r[k]) || 0), 0);
  const qChannel = (source: string, medium: string, group = "") => {
    const s = String(source || "").toLowerCase(), g = String(group || "").toLowerCase();
    if (/pinterest/.test(s)) return "pinterest";
    if (/google/.test(s)) return "google";
    if (/facebook|instagram|meta/.test(s)) return "meta";
    if (/tiktok/.test(s)) return "tiktok";
    if (/youtube/.test(s)) return "youtube";
    if (/direct/.test(s) || g.includes("direct")) return "direto";
    if (g.includes("organic")) return "orgânico";
    return g || s || String(medium || "").toLowerCase() || "outros";
  };
  const qRows = (Array.isArray(qualidadeCanais) ? qualidadeCanais : []).map((r: any) => ({
    channel: qChannel(r.sessionSource, r.sessionMedium, r.sessionDefaultChannelGroup), source: r.sessionSource || "", medium: r.sessionMedium || "", sourceMedium: [r.sessionSource, r.sessionMedium].filter(Boolean).join(" / "), campaign: r.sessionCampaignName || "",
    sessions: r.sessions || 0, users: r.totalUsers || 0, engagedSessions: r.engagedSessions || 0, engagementRate: r.engagementRate || 0, avgSeconds: r.averageSessionDuration || 0, activeSeconds: r.userEngagementDuration || 0, pagesPerSession: r.screenPageViewsPerSession || 0, conversions: r.keyEvents || 0,
  }));
  const qPages = (Array.isArray(qualidadePaginas) ? qualidadePaginas : []).map((r: any) => {
    const sm = String(r.sessionSourceMedium || ""), parts = sm.split(" / "), source = parts.shift() || "", medium = parts.join(" / ");
    return { channel: qChannel(source, medium), sourceMedium: sm, page: r.pagePath || "", title: r.pageTitle || "", sessions: r.sessions || 0, users: r.totalUsers || 0, views: r.screenPageViews || 0, avgSeconds: r.averageSessionDuration || 0, activeSeconds: r.userEngagementDuration || 0, engagementRate: r.engagementRate || 0, conversions: r.keyEvents || 0 };
  });
  const qKeywords = (Array.isArray(qualidadePalavras) ? qualidadePalavras : []).filter((r: any) => !/^\(not set\)$|^$/i.test(String(r.sessionGoogleAdsKeyword || r.sessionGoogleAdsQuery || ""))).map((r: any) => ({
    channel: "google", keyword: r.sessionGoogleAdsKeyword || "", query: r.sessionGoogleAdsQuery || "", campaign: r.sessionGoogleAdsCampaignName || "", sessions: r.sessions || 0, engagedSessions: r.engagedSessions || 0, engagementRate: r.engagementRate || 0, avgSeconds: r.averageSessionDuration || 0, activeSeconds: r.userEngagementDuration || 0, pagesPerSession: r.screenPageViewsPerSession || 0, conversions: r.keyEvents || 0,
  }));
  const totalSessions = tot(canais as any[], "sessions"), totalEngaged = tot(canais as any[], "engagedSessions");
  const weighted = (k: string) => totalSessions ? (Array.isArray(canais) ? canais : []).reduce((s: number, r: any) => s + (Number(r[k]) || 0) * (Number(r.sessions) || 0), 0) / totalSessions : 0;
  return { propertyId: prop, periodo: { since: m.since, until: m.until }, canais, paginas, origens, campanhas,
    quality: { source: "ga4", channels: qRows, pages: qPages, keywords: qKeywords, total: { sessions: totalSessions, users: tot(canais as any[], "totalUsers"), engagedSessions: totalEngaged, engagementRate: totalSessions ? totalEngaged / totalSessions : 0, avgSeconds: weighted("averageSessionDuration"), activeSeconds: tot(canais as any[], "userEngagementDuration"), pagesPerSession: weighted("screenPageViewsPerSession"), conversions: tot(canais as any[], "keyEvents") } },
    total: { receita: tot(canais as any[], "purchaseRevenue"), compras: tot(canais as any[], "transactions"), sessoes: totalSessions, conversoes: tot(canais as any[], "conversions") } };
}

// Fallback universal dos KPIs de qualidade, alimentado pelo pixel próprio depois do consentimento.
// É intencionalmente agregado em memória sobre a tabela compacta nova — nunca consulta o track_events histórico pesado.
async function journeyQualityPixel(m: any) {
  const clientId = String(m.clientId || ""), since = String(m.since || ""), until = String(m.until || "");
  if (!clientId || !since || !until) throw new Error("clientId, since e until são obrigatórios");
  const rows = await _sbAll("journey_quality_events", `client_id=eq.${encodeURIComponent(clientId)}&created_at=gte.${encodeURIComponent(since + "T00:00:00Z")}&created_at=lte.${encodeURIComponent(until + "T23:59:59Z")}&select=event_type,session_id,anon_id,channel,source,medium,campaign,term,content,page,title,active_seconds,created_at&order=created_at.asc`, 50000);
  const sessions: Record<string, any> = {};
  for (const r of rows) {
    const key = r.session_id || (r.anon_id ? `${r.anon_id}:${String(r.created_at).slice(0, 10)}` : ""); if (!key) continue;
    const s = sessions[key] || (sessions[key] = { key, channel: r.channel || "direto", source: r.source || "", medium: r.medium || "", campaign: r.campaign || "", term: r.term || "", activeSeconds: 0, pageviews: 0, pages: new Set<string>() });
    if (!s.campaign && r.campaign) s.campaign = r.campaign; if (!s.term && r.term) s.term = r.term;
    if (r.event_type === "pageview") s.pageviews++;
    if (r.page) s.pages.add(r.page);
    s.activeSeconds = Math.min(14400, s.activeSeconds + Math.max(0, Number(r.active_seconds) || 0));
  }
  const list = Object.values(sessions); const aggregate = (items: any[], keyFn: (s: any) => string) => {
    const out: Record<string, any> = {};
    for (const s of items) { const k = keyFn(s); if (!k) continue; const a = out[k] || (out[k] = { sessions: 0, engagedSessions: 0, activeSeconds: 0, pageviews: 0 }); a.sessions++; a.engagedSessions += s.activeSeconds >= 10 || s.pageviews >= 2 ? 1 : 0; a.activeSeconds += s.activeSeconds; a.pageviews += s.pageviews; }
    return out;
  };
  const byChannel = aggregate(list, (s) => [s.channel, s.source, s.medium, s.campaign].join("|"));
  const channels = Object.entries(byChannel).map(([k, a]: any) => { const [channel, source, medium, campaign] = k.split("|"); return { channel, source, medium, sourceMedium: [source, medium].filter(Boolean).join(" / "), campaign, sessions: a.sessions, users: a.sessions, engagedSessions: a.engagedSessions, engagementRate: a.sessions ? a.engagedSessions / a.sessions : 0, avgSeconds: a.sessions ? a.activeSeconds / a.sessions : 0, activeSeconds: a.activeSeconds, pagesPerSession: a.sessions ? a.pageviews / a.sessions : 0, conversions: 0 }; }).sort((a, b) => b.sessions - a.sessions);
  const pageAgg: Record<string, any> = {};
  for (const r of rows) { if (!r.page) continue; const k = [r.channel || "direto", r.source || "", r.medium || "", r.page, r.title || ""].join("|"); const a = pageAgg[k] || (pageAgg[k] = { views: 0, activeSeconds: 0, sessionKeys: new Set<string>() }); if (r.event_type === "pageview") a.views++; a.activeSeconds += Math.max(0, Number(r.active_seconds) || 0); if (r.session_id) a.sessionKeys.add(r.session_id); }
  const pages = Object.entries(pageAgg).map(([k, a]: any) => { const [channel, source, medium, page, title] = k.split("|"); const n = a.sessionKeys.size || 1; return { channel, sourceMedium: [source, medium].filter(Boolean).join(" / "), page, title, sessions: a.sessionKeys.size, users: a.sessionKeys.size, views: a.views, avgSeconds: a.activeSeconds / n, activeSeconds: a.activeSeconds, engagementRate: 0, conversions: 0 }; }).sort((a, b) => b.views - a.views || b.activeSeconds - a.activeSeconds);
  const kwAgg = aggregate(list.filter((s: any) => s.term), (s) => [s.term, s.campaign].join("|"));
  const keywords = Object.entries(kwAgg).map(([k, a]: any) => { const [keyword, campaign] = k.split("|"); return { channel: "google", keyword, query: "", campaign, sessions: a.sessions, engagedSessions: a.engagedSessions, engagementRate: a.sessions ? a.engagedSessions / a.sessions : 0, avgSeconds: a.sessions ? a.activeSeconds / a.sessions : 0, activeSeconds: a.activeSeconds, pagesPerSession: a.sessions ? a.pageviews / a.sessions : 0, conversions: 0 }; }).sort((a, b) => b.sessions - a.sessions);
  const totalSessions = list.length, activeSeconds = list.reduce((s: number, x: any) => s + x.activeSeconds, 0), engagedSessions = list.filter((s: any) => s.activeSeconds >= 10 || s.pageviews >= 2).length, pageviews = list.reduce((s: number, x: any) => s + x.pageviews, 0);
  return { source: "pixel", channels, pages, keywords, total: { sessions: totalSessions, users: new Set(list.map((s: any) => s.key.split(":")[0])).size, engagedSessions, engagementRate: totalSessions ? engagedSessions / totalSessions : 0, avgSeconds: totalSessions ? activeSeconds / totalSessions : 0, activeSeconds, pagesPerSession: totalSessions ? pageviews / totalSessions : 0, conversions: 0 }, truncated: rows.length >= 50000 };
}
// GA4 por DIA × ORIGEM/MÍDIA × CAMPANHA × CONTEÚDO DO ANÚNCIO — evento purchase (transações) e receita.
// Alimenta o Banco de Dados como canal "ga4". Retenção de dados do GA4 (2/14 meses) só afeta relatórios de
// usuário/evento individual — relatórios agregados (data + dimensões como essas) seguem consultáveis sempre.
// Receita = "Valor do evento" (eventValue) filtrado por eventName=purchase — bate com a Exploração que o gestor já usa
// no GA4 (não usar purchaseRevenue: é a receita "oficial" de e-commerce do GA4, calculada diferente e não confere).
async function ga4DailyBySource(m: any) {
  const prop = String(m.propertyId || "").replace(/[^0-9]/g, "");
  if (!prop) throw new Error("propertyId do GA4 obrigatório");
  const j = await _googleTryTokens(["https://www.googleapis.com/auth/analytics.readonly"], async (token) => {
    const r = await fetch(`https://analyticsdata.googleapis.com/v1beta/properties/${prop}:runReport`, {
      method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        dateRanges: [{ startDate: m.since, endDate: m.until }],
        dimensions: [{ name: "date" }, { name: "sessionSourceMedium" }, { name: "sessionCampaignName" }, { name: "sessionManualAdContent" }],
        metrics: [{ name: "eventCount" }, { name: "eventValue" }],
        dimensionFilter: { filter: { fieldName: "eventName", stringFilter: { value: "purchase" } } },
        limit: 100000,
      }),
    });
    return r.json();
  });
  if (j.error) throw new Error(`GA4: ${j.error.message}${/permission|PERMISSION/i.test(j.error.message || "") ? ` — dê acesso de Leitor à ${_gsaEmail()} na propriedade ${prop} (ou conecte um login pessoal com acesso em Configurações)` : ""}`);
  return (j.rows || []).map((row: any) => {
    const raw = row.dimensionValues[0].value || ""; // YYYYMMDD
    const date = raw.length === 8 ? `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}` : raw;
    return {
      date, sourceMedium: row.dimensionValues[1].value || "(not set)",
      campaign: row.dimensionValues[2].value || "", adContent: row.dimensionValues[3].value || "",
      purchases: Number(row.metricValues[0].value) || 0, revenue: Number(row.metricValues[1].value) || 0,
    };
  }).filter((r: any) => r.purchases > 0 || r.revenue > 0);
}
// Diagnóstico: quais propriedades do Search Console a service account enxerga
async function gscSites() {
  const j = await _googleTryTokens(["https://www.googleapis.com/auth/webmasters.readonly"], async (token) => {
    const r = await fetch("https://searchconsole.googleapis.com/webmasters/v3/sites", { headers: { Authorization: `Bearer ${token}` } });
    return r.json();
  });
  if (j.error) throw new Error(`Search Console: ${j.error.message}`);
  return { email: _gsaEmail(), sites: (j.siteEntry || []).map((s: any) => ({ site: s.siteUrl, permissao: s.permissionLevel })) };
}
async function _gscListSites(): Promise<string[]> {
  const j = await _googleTryTokens(["https://www.googleapis.com/auth/webmasters.readonly"], async (token) => {
    const r = await fetch("https://searchconsole.googleapis.com/webmasters/v3/sites", { headers: { Authorization: `Bearer ${token}` } });
    return r.json();
  });
  return ((j && j.siteEntry) || []).map((s: any) => s.siteUrl);
}
async function gscReport(m: any) {
  let site = String(m.siteUrl || "").trim();
  if (!site) throw new Error("siteUrl do Search Console obrigatório (ex: https://site.com.br/ ou sc-domain:site.com.br)");
  // o Search Console é exigente com o formato (http/https, www, barra final, sc-domain).
  // Se o que está no cadastro não for exatamente uma das propriedades, acha a do mesmo domínio.
  try {
    const lista = await _gscListSites();
    if (lista.length && !lista.includes(site)) {
      const dom = (u: string) => String(u).replace(/^sc-domain:/, "").replace(/^https?:\/\//, "").replace(/^www\./, "").replace(/\/+$/, "").toLowerCase();
      const alvo = dom(site);
      // entre as candidatas do mesmo domínio, a propriedade de DOMÍNIO (sc-domain:) é a mais completa —
      // cobre http/https/www/subdomínios. Só cai na de prefixo de URL se não houver domínio.
      const cands = lista.filter((s) => dom(s) === alvo).sort((a, b) => (b.startsWith("sc-domain:") ? 1 : 0) - (a.startsWith("sc-domain:") ? 1 : 0));
      const hit = cands[0];
      if (hit) site = hit;
      else throw new Error(`a service account não tem acesso a "${site}". Propriedades disponíveis: ${lista.join(", ") || "(nenhuma)"} — adicione ${_gsaEmail()} como usuário da propriedade certa no Search Console (ou conecte um login pessoal com acesso em Configurações).`);
    }
  } catch (e) { if (/service account não tem acesso/.test(String((e as any).message))) throw e; /* senão segue com o valor informado */ }
  const q = async (dimensions: string[], rowLimit = 25) => {
    const j = await _googleTryTokens(["https://www.googleapis.com/auth/webmasters.readonly"], async (token) => {
      const r = await fetch(`https://searchconsole.googleapis.com/webmasters/v3/sites/${encodeURIComponent(site)}/searchAnalytics/query`, {
        method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ startDate: m.since, endDate: m.until, dimensions, rowLimit }),
      });
      return r.json();
    });
    if (j.error) throw new Error(`Search Console: ${j.error.message}${/permission|not have|403/i.test(j.error.message || "") ? ` — adicione ${_gsaEmail()} como usuário da propriedade (ou conecte um login pessoal com acesso em Configurações)` : ""}`);
    return (j.rows || []).map((row: any) => ({ chave: (row.keys || []).join(" · "), cliques: row.clicks || 0, impressoes: row.impressions || 0, ctr: +((row.ctr || 0) * 100).toFixed(2), posicao: +(row.position || 0).toFixed(1) }));
  };
  const [termos, paginas] = await Promise.all([q(["query"], 30), q(["page"], 20).catch(() => [])]);
  return { site, periodo: { since: m.since, until: m.until }, termos, paginas };
}
// ===================== JORNADA DO LEAD =====================
// Junta TODAS as fontes numa linha do tempo por pessoa, casando identidades (telefone/email/anon/ctwa).
// Idempotente: cada toque é gravado com (ref_table, ref_id) único — rodar de novo não duplica.
function _hash36(s: string): string { let h = 5381; for (let i = 0; i < s.length; i++) { h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; } let h2 = 52711; for (let i = s.length - 1; i >= 0; i--) { h2 = ((h2 << 5) + h2 + s.charCodeAt(i)) >>> 0; } return h.toString(36) + h2.toString(36); }
const _digits = (v: any) => String(v || "").replace(/[^0-9]/g, "");
const _phoneKey = (v: any) => { const d = _digits(v); return d.length >= 8 ? d.slice(-8) : ""; }; // ignora DDI/9º dígito
const _emailKey = (v: any) => String(v || "").trim().toLowerCase();
// Fallback de origem por TELEFONE JÁ CONHECIDO: quando uma conversa nova chega sem CTWA nativo e sem [#ref] no
// texto, olha se esse telefone já apareceu em outro toque com canal/campanha real (RD Station, conversa
// anterior, pedido da planilha) via o grafo de identidade da Jornada (lead_identities/lead_touchpoints) — reusa
// essa origem em vez de marcar como orgânico. Marca matched_by:'phone_history' pra deixar claro que é INFERIDO,
// não um sinal direto desta conversa (nunca confundir com CTWA/ref, que são diretos).
async function waPhoneHistoryOrigin(clientId: string, phone: string): Promise<{ type: string; data: Record<string, unknown> } | null> {
  const pk = _phoneKey(phone); if (!pk || !clientId) return null;
  const ident = (await sbGet("lead_identities", `client_id=eq.${encodeURIComponent(clientId)}&kind=eq.phone&value=eq.${encodeURIComponent(pk)}&select=person_id&limit=1`))[0];
  if (!ident) return null;
  const toques = await sbGet("lead_touchpoints", `client_id=eq.${encodeURIComponent(clientId)}&person_id=eq.${encodeURIComponent(ident.person_id)}&channel=in.(meta,google)&select=channel,source,campaign,adset,ad,term,ts&order=ts.desc&limit=1`);
  const t = toques[0]; if (!t) return null;
  return { type: "anuncio", data: { channel: t.channel, campaign: t.campaign || "", adset: t.adset || "", ad: t.ad || "", keyword: t.term || "", matched_by: "phone_history", matched_at: t.ts } };
}
// de onde veio o toque (canal), a partir de utm/click ids/referrer
function _jChannel(source?: string, medium?: string, gclid?: string, fbclid?: string, referrer?: string, selfDom?: string): string {
  // o RD manda "categoria | detalhe" (ex.: "referência | linktr.ee", "social | link+da+bio+do+rd+station")
  const raw = String(source || "").toLowerCase().replace(/\+/g, " ");
  const parts = raw.split("|").map((x) => x.trim());
  const cat = parts[0] || "", det = parts.slice(1).join(" ") || "";
  const m = String(medium || "").toLowerCase();
  const all = `${cat} ${det} ${m}`;
  if (gclid || /google|gads|adwords|gclid/.test(all)) return "google";
  if (fbclid || /meta|facebook|\bfb\b|instagram|\big\b|fbclid/.test(all)) return "meta";
  if (/whats|wpp|zap/.test(all)) return "whatsapp";
  if (/e-?mail|newsletter|rd ?station|rdstation/.test(all)) return "email";
  // auto-referência: o próprio site do cliente não é um canal de aquisição
  if (selfDom && det && det.includes(selfDom)) return "direto";
  if (/linktr|link.?bio|beacons|\bbio\b/.test(all)) return "social";
  if (/social|tiktok|youtube|pinterest|linkedin|twitter|kwai/.test(all)) return "social";
  if (/direto|direct|\bnone\b/.test(all)) return "direto";
  if (/org[âa]nic|organic|busca/.test(all)) return "organico";
  if (/refer[êe]ncia|referral/.test(cat)) return "referral";
  if (/cpc|ppc|paid|ads|pago/.test(m)) return cat || "pago";
  if (cat && !/outros|not set|desconhecid/.test(cat)) return cat;
  if (referrer) return "referral";
  return "direto";
}
// ---- Vendas da planilha de pedidos, casadas por HASH (o e-mail/telefone nunca é gravado) ----
async function _sha(s: string): Promise<string> {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}
// data BR (dd/mm/aaaa) ou ISO → ISO
function _dtBR(v: any): string {
  const s = String(v || "").trim(); if (!s) return "";
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}T12:00:00Z`;
  const i = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (i) return `${i[0]}T12:00:00Z`;
  return "";
}
// status da planilha → venda confirmada ou pedido em aberto
function _ordStatus(v: any): "purchase" | "checkout" | "" {
  const s = String(v || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  if (!s) return "checkout";
  // As NEGATIVAS vêm primeiro de propósito: "Não Autorizada" contém "autoriz" e era classificada como VENDA
  // pela regra de baixo (numa base real, 13.774 linhas desse status seriam contadas como compra paga).
  // "Liberado Automaticamente" (1.462 registros na base da CFP) NÃO é venda — confirmado pela gestora.
  // exceção antes das negativas: no estorno parcial a maior parte do dinheiro ficou, e o "aguardando
  // estorno" ainda não voltou. Ambos contam como venda — igual à aba Negócio, pra não divergir.
  if (/estorn\w* parcial|aguardando estorno/.test(s)) return "purchase";
  if (/nao autoriz|nao aprovad|cancel|recus|estorn|expir|reembols|falh|negad|charge|blacklist|nao process|liberad/.test(s)) return "";
  // "Disponivel" e "Debitado" são venda paga do vocabulário antigo da plataforma (até set/2022):
  // naquele período NÃO existe nenhuma linha "Pago", e sem isso 9 meses inteiros de venda somem.
  // "Lead" é a fatura gerada no checkout — mesma pessoa/dia/produto/valor da linha "Disponivel",
  // então é pedido em aberto, nunca compra. "Finalizado" já cai no /finaliz/ abaixo (assinatura
  // recorrente concluída, cuja receita só existe naquela linha).
  if (/disponivel|debitad/.test(s)) return "purchase";
  if (/aprovad|pago|paga|conclu|complet|autoriz|captur|finaliz|entregue/.test(s)) return "purchase";
  return "checkout"; // pendente, aguardando, processando…
}
// Lê a aba de pedidos e devolve os toques de compra JÁ CASADOS por hash com as identidades do lead.
// Nada de e-mail/CPF/nome/CEP sai desta função — só pessoa, data, valor, produto e status.
async function ordersFromSheet(spreadsheetId: string, tab: string, hashToPerson: Record<string, string>, de = 2, blocos = 2) {
  const keyJson = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_KEY"); if (!keyJson) return { toques: [], lidas: 0, casadas: 0 };
  const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(keyJson), scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"] });
  const sheets = google.sheets({ version: "v4", auth });
  // lê o cabeçalho e depois pagina em blocos (planilha grande não cabe na memória de uma vez)
  const h0 = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tab}'!A1:Z1` });
  const head = (((h0.data.values || [])[0]) || []).map((h: any) => String(h || "").toLowerCase().trim());
  if (!head.length) return { toques: [], lidas: 0, casadas: 0 };
  const col = (...names: string[]) => { for (const n of names) { const i = head.findIndex((h) => h === n || h.includes(n)); if (i >= 0) return i; } return -1; };
  const cData = col("data"), cMail = col("email", "e-mail"), cCel = col("celular", "fone", "telefone"), cSt = col("status"), cTot = col("total", "valor"), cProd = col("produto"), cId = col("id");
  const toques: any[] = []; let casadas = 0, lidas = 0;
  const BLOCO = 3000; let ultima = de, fimDaPlanilha = false;
  for (let b = 0, ini = de; b < blocos; b++, ini += BLOCO) {
    ultima = ini;
    const fim = ini + BLOCO - 1;
    let res: any;
    try {
      res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `'${tab}'!A${ini}:Z${fim}` });
    } catch (e) {
      // Cursor parado UMA linha depois do fim da planilha: o Google recusa o range ("exceeds grid limits") e o
      // erro travava a sincronização inteira — nenhuma venda nova entrava na Jornada e o cursor nunca mais
      // avançava (a da Curso Fernanda Pessoa ficou parada de 03/08 até aqui). Fim de planilha não é falha.
      if (/exceeds grid limits/i.test(String((e as any)?.message || e))) { fimDaPlanilha = true; break; }
      throw e;
    }
    const rows: any[][] = res.data.values || [];
    if (!rows.length) { fimDaPlanilha = true; break; }
    lidas += rows.length;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r]; if (!row || !row.length) continue;
    const kind = _ordStatus(cSt >= 0 ? row[cSt] : "");
    if (!kind) continue; // cancelado/estornado: ignora
    // casa por hash — o valor cru some da memória logo em seguida
    let pid = "";
    if (cMail >= 0 && row[cMail]) pid = hashToPerson["email:" + String(row[cMail]).trim().toLowerCase()] || "";
    if (!pid && cCel >= 0 && row[cCel]) { const d = String(row[cCel]).replace(/[^0-9]/g, ""); if (d.length >= 8) pid = hashToPerson["phone:" + d.slice(-8)] || ""; }
    if (!pid) continue;
    casadas++;
    const ts = _dtBR(cData >= 0 ? row[cData] : ""); if (!ts) continue;
    toques.push({ _pid: pid, ts, kind, channel: "site", label: kind === "purchase" ? ("Compra" + (cProd >= 0 && row[cProd] ? " · " + String(row[cProd]).slice(0, 60) : "")) : ("Pedido " + String(cSt >= 0 ? row[cSt] : "em aberto").slice(0, 30)),
      value: cTot >= 0 ? (parseNumberBR(row[cTot]) || 0) : 0, ref_table: "sheet_order", ref_id: String((cId >= 0 && row[cId]) || (tab + "_" + (ini + r))) });
    }
    ultima = ini + rows.length;
    if (rows.length < BLOCO) { fimDaPlanilha = true; break; }
  }
  return { toques, lidas, casadas, proxima: fimDaPlanilha ? 0 : ultima, fim: fimDaPlanilha };
}
// Importa vendas enviadas de FORA (planilha local do cliente, via enviador agendado no computador dele).
// Mesma regra da leitura por Google Sheets: só entra quem casa com um lead conhecido, o ID do pedido evita
// duplicar, e reenviar a mesma linha só atualiza — de propósito, porque a base é corrigida retroativamente.
async function journeyOrdersImport(m: any) {
  const clientId = String(m.clientId || "").trim();
  if (!clientId) throw new Error("clientId obrigatório");
  const linhas: any[] = Array.isArray(m.rows) ? m.rows : [];
  if (!linhas.length) return { recebidas: 0, casadas: 0, gravadas: 0, compras: 0 };
  const ident = await _sbAll("lead_identities", `client_id=eq.${encodeURIComponent(clientId)}&kind=in.(email,phone)&select=person_id,kind,value`);
  const h2p: Record<string, string> = {};
  for (const i of (ident || [])) h2p[`${i.kind}:${i.value}`] = i.person_id;
  const toques: any[] = []; let casadas = 0;
  for (const row of linhas) {
    const kind = _ordStatus(row.status || "");
    if (!kind) continue;
    let pid = "";
    if (row.email) pid = h2p["email:" + String(row.email).trim().toLowerCase()] || "";
    if (!pid && row.celular) { const d = String(row.celular).replace(/[^0-9]/g, ""); if (d.length >= 8) pid = h2p["phone:" + d.slice(-8)] || ""; }
    if (!pid) continue;
    casadas++;
    const ts = _dtBR(row.data || ""); if (!ts) continue;
    const refId = String(row.id || "").trim(); if (!refId) continue;
    toques.push({ id: "t_" + _wuid(), client_id: clientId, person_id: pid, ts, kind, channel: "site",
      label: kind === "purchase" ? ("Compra" + (row.produto ? " · " + String(row.produto).slice(0, 60) : "")) : ("Pedido " + String(row.status || "em aberto").slice(0, 30)),
      value: parseNumberBR(row.total) || 0, ref_table: "sheet_order", ref_id: refId });
  }
  for (let i = 0; i < toques.length; i += 400) await _sbUpsert("lead_touchpoints", toques.slice(i, i + 400), "ref_table,ref_id");
  const compras = toques.filter((x) => x.kind === "purchase");
  return { recebidas: linhas.length, casadas, gravadas: toques.length, compras: compras.length, receita: compras.reduce((s, x) => s + (x.value || 0), 0), identidades: (ident || []).length };
}
// classifica o evento do RD: compra confirmada, pedido/checkout iniciado, negociação ou formulário comum
function _rdKind(ev?: string): string {
  const e = String(ev || "").toLowerCase();
  if (/compra|purchase|venda\b|pagamento|pago|aprovad|matricul/.test(e)) return "purchase";
  if (/checkout|carrinho|\bcart\b|pedido|\border\b/.test(e)) return "checkout";
  if (/negocia|deal|oportunidad/.test(e)) return "deal";
  return "form";
}
// grafo de identidade em memória durante a reconstrução (union-find simples)
function _identityGraph() {
  const parent: Record<string, string> = {}, info: Record<string, any> = {};
  const find = (k: string): string => { let r = k; while (parent[r] && parent[r] !== r) r = parent[r]; let c = k; while (parent[c] && parent[c] !== c) { const n = parent[c]; parent[c] = r; c = n; } return r; };
  const add = (k: string) => { if (!parent[k]) { parent[k] = k; info[k] = {}; } return find(k); };
  const union = (keys: string[]) => { const ks = keys.filter(Boolean).map(add); if (!ks.length) return ""; const root = find(ks[0]); ks.forEach((k) => { const r = find(k); if (r !== root) { parent[r] = root; info[root] = { ...(info[r] || {}), ...(info[root] || {}) }; } }); return root; };
  const setInfo = (k: string, patch: any) => { const r = find(add(k)); info[r] = { ...(info[r] || {}), ...Object.fromEntries(Object.entries(patch).filter(([, v]) => v)) }; };
  return { find, add, union, setInfo, parent, info, keys: () => Object.keys(parent) };
}
async function journeyRebuild(m: any) {
  const clientId = String(m.clientId || "").trim();
  if (!clientId) throw new Error("clientId obrigatório");
  const dias = Number(m.dias) || 365;
  const since = new Date(Date.now() - dias * 864e5).toISOString();
  // domínio do próprio cliente — pra não contar auto-referência como canal de aquisição
  const _cli = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=site_url&limit=1`))[0] || {};
  const selfDom = String(_cli.site_url || "").replace(/^https?:\/\//, "").replace(/^www\./, "").replace(/\/.*$/, "").toLowerCase() || "";
  const G = _identityGraph();
  const touches: any[] = [];
  const push = (t: any) => { if (t.ts && t.keys && t.keys.length) touches.push(t); };

  // ---- 1) RD Station: formulários/eventos (email + telefone = a identidade mais forte) ----
  const rd = await _sbAll("rd_conversions", `client_id=eq.${encodeURIComponent(clientId)}&converted_at=gte.${since}&select=id,event_identifier,email,name,phone,source,medium,campaign,content,term,converted_at&order=converted_at.asc`);
  for (const r of (rd || [])) {
    const ks: string[] = []; const ek = _emailKey(r.email), pk = _phoneKey(r.phone);
    if (ek) ks.push("email:" + ek); if (pk) ks.push("phone:" + pk);
    if (!ks.length) continue;
    const root = G.union(ks); G.setInfo(root, { name: r.name, email: r.email, phone: r.phone });
    push({ keys: ks, ts: r.converted_at, kind: _rdKind(r.event_identifier), channel: _jChannel(r.source, r.medium, "", "", "", selfDom), source: r.source, medium: r.medium, campaign: r.campaign, content: r.content, term: r.term, label: r.event_identifier || "conversão", ref_table: "rd_conversions", ref_id: r.id });
  }

  // ---- 2) WhatsApp: conversa (origem do anúncio) + 1º contato + trocas de etapa ----
  const convs = await _sbAll("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&select=id,chat_id,name,origin_type,origin,stage,created_at,last_at`);
  const convKeys: Record<string, string[]> = {};
  for (const cv of (convs || [])) {
    const pk = _phoneKey(cv.chat_id); if (!pk) continue;
    const ks = ["phone:" + pk]; const o = cv.origin || {};
    if (o.ctwa_clid) ks.push("ctwa:" + String(o.ctwa_clid).slice(0, 60));
    convKeys[cv.id] = ks;
    const root = G.union(ks); G.setInfo(root, { name: cv.name, phone: cv.chat_id, stage: cv.stage });
    // o clique no anúncio que originou a conversa
    if (cv.origin_type === "anuncio") {
      push({ keys: ks, ts: cv.created_at, kind: "ad_click", channel: o.channel === "google" ? "google" : "meta", source: o.channel || (o.platform || "meta"), medium: "cpc", campaign: o.campaign || "", adset: o.adset || "", ad: o.ad || o.title || "", term: o.keyword || "", label: "Clique no anúncio" + (o.platform ? ` (${o.platform})` : ""), ref_table: "wa_origin", ref_id: cv.id });
    } else if (cv.origin_type === "utm") {
      push({ keys: ks, ts: cv.created_at, kind: "link_click", channel: _jChannel(o.track_source || o.channel, o.medium, "", "", "", selfDom), source: o.track_source || o.channel || "", medium: o.medium || "", campaign: o.campaign || "", label: "Link rastreável", ref_table: "wa_origin", ref_id: cv.id });
    }
  }
  // 1ª mensagem do lead e 1ª resposta (dá pra medir tempo de resposta) — o chat completo fica no CRM
  const cvIds = Object.keys(convKeys);
  for (let i = 0; i < cvIds.length; i += 40) {
    const chunk = cvIds.slice(i, i + 40);
    const msgs = await _sbAll("wa_messages", `conversation_id=in.(${chunk.map((x) => encodeURIComponent(x)).join(",")})&select=id,conversation_id,direction,ts,text&order=ts.asc`, 8000);
    const seen: Record<string, any> = {};
    for (const g of (msgs || [])) {
      const k = g.conversation_id + "|" + g.direction; if (seen[k]) continue; seen[k] = 1;
      push({ keys: convKeys[g.conversation_id], ts: g.ts, kind: g.direction === "in" ? "message_in" : "message_out", channel: "whatsapp", label: g.direction === "in" ? "Iniciou conversa no WhatsApp" : "Primeira resposta da equipe", meta: { texto: String(g.text || "").slice(0, 160) }, ref_table: "wa_first_msg", ref_id: k });
    }
  }
  // trocas de etapa (MQL, SQL, comprou…)
  const jr = await _sbAll("wa_journey", `client_id=eq.${encodeURIComponent(clientId)}&select=id,conversation_id,to_stage,why,created_at&order=created_at.asc`);
  for (const j of (jr || [])) {
    const ks = convKeys[j.conversation_id]; if (!ks) continue;
    push({ keys: ks, ts: j.created_at, kind: /compr|vend|fech/i.test(String(j.to_stage)) ? "purchase" : "stage", channel: "whatsapp", label: String(j.to_stage || "").toUpperCase(), meta: { motivo: String(j.why || "").slice(0, 200) }, ref_table: "wa_journey", ref_id: j.id });
  }

  // ---- 3) Pixel do site: pageview / clique em link / clique no WhatsApp ----
  const ev = await _sbAll("track_events", `client_id=eq.${encodeURIComponent(clientId)}&type=in.(pageview,wpp_click,link_click,form)&created_at=gte.${since}&select=id,type,session_id,anon_id,utm_source,utm_medium,utm_campaign,utm_content,utm_term,fbclid,gclid,referrer,landing,link_slug,created_at&order=created_at.asc`);
  for (const e of (ev || [])) {
    const ks: string[] = [];
    if (e.anon_id) ks.push("anon:" + e.anon_id);
    if (e.gclid) ks.push("gclid:" + e.gclid);
    if (e.fbclid) ks.push("fbclid:" + e.fbclid);
    if (!ks.length) continue;
    G.union(ks);
    push({ keys: ks, ts: e.created_at, kind: e.type === "wpp_click" ? "wpp_click" : e.type === "link_click" ? "link_click" : "pageview", channel: _jChannel(e.utm_source, e.utm_medium, e.gclid, e.fbclid, e.referrer, selfDom), source: e.utm_source || "", medium: e.utm_medium || "", campaign: e.utm_campaign || "", content: e.utm_content || "", term: e.utm_term || "", page: e.landing || "", referrer: e.referrer || "", label: e.type === "wpp_click" ? "Clicou no WhatsApp do site" : e.type === "link_click" ? ("Link /" + (e.link_slug || "")) : "Visitou o site", ref_table: "track_events", ref_id: e.id });
  }

  // ---- 4) resolve as pessoas e grava ----
  // person_id DETERMINÍSTICO (hash da menor identidade do grupo): o mesmo lead recebe sempre o mesmo id
  // entre rebuilds, sem precisar carregar o mapa de identidades antigo na memória.
  const rootToPerson: Record<string, string> = {};
  const groups: Record<string, string[]> = {};
  for (const k of G.keys()) { const root = G.find(k); (groups[root] = groups[root] || []).push(k); }
  const people: Record<string, any> = {};
  for (const root of Object.keys(groups)) {
    const keys = groups[root].sort();
    const pid = "p_" + _hash36(clientId + "|" + keys[0]);
    rootToPerson[root] = pid;
    people[pid] = { id: pid, client_id: clientId, ...(G.info[root] || {}), keys };
  }
  const personOf = (keys: string[]) => { for (const k of keys) { const r = G.find(k); if (rootToPerson[r]) return rootToPerson[r]; } return ""; };

  // um passe só: resolve a pessoa de cada toque e acumula primeiro/último contato (evita varrer tudo por pessoa)
  const span: Record<string, { min: string; max: string }> = {};
  for (const t of touches) {
    const pid = t._pid || personOf(t.keys); t._pid = pid; if (!pid || !t.ts) continue;
    const s = span[pid] || (span[pid] = { min: t.ts, max: t.ts });
    if (t.ts < s.min) s.min = t.ts; if (t.ts > s.max) s.max = t.ts;
  }
  // pessoas
  const nowIso = new Date().toISOString();
  const peopleRows = Object.values(people).map((p: any) => ({ id: p.id, client_id: clientId, name: p.name || null, phone: p.phone ? String(p.phone).slice(0, 30) : null, email: p.email || null, stage: p.stage || null, first_seen: (span[p.id] || {}).min || null, last_seen: (span[p.id] || {}).max || null, updated_at: nowIso }));
  for (let i = 0; i < peopleRows.length; i += 500) await _sbUpsert("lead_people", peopleRows.slice(i, i + 500), "id");
  // identidades
  const identRows: any[] = [];
  Object.values(people).forEach((p: any) => p.keys.forEach((k: string) => { const ix = k.indexOf(":"); identRows.push({ id: "i_" + _wuid(), client_id: clientId, person_id: p.id, kind: k.slice(0, ix), value: k.slice(ix + 1) }); }));
  for (let i = 0; i < identRows.length; i += 500) await _sbUpsert("lead_identities", identRows.slice(i, i + 500), "client_id,kind,value");
  // toques — grava em lotes convertendo só o lote (não duplica o array inteiro na memória)
  let nT = 0;
  for (let i = 0; i < touches.length; i += 500) {
    const batch = touches.slice(i, i + 500).filter((t) => t._pid).map((t) => ({ id: "t_" + _wuid(), client_id: clientId, person_id: t._pid, ts: t.ts, kind: t.kind, channel: t.channel || null, source: t.source || null, medium: t.medium || null, campaign: t.campaign || null, adset: t.adset || null, ad: t.ad || null, term: t.term || null, content: t.content || null, page: t.page || null, referrer: t.referrer || null, label: t.label || null, value: t.value || 0, ref_table: t.ref_table, ref_id: String(t.ref_id), meta: t.meta || null }));
    if (batch.length) { await _sbUpsert("lead_touchpoints", batch, "ref_table,ref_id"); nT += batch.length; }
  }
  // pessoas que deixaram de existir (viraram outra depois de um merge) — some com elas
  try { await fetch(`${_SB_URL}/rest/v1/lead_people?client_id=eq.${encodeURIComponent(clientId)}&updated_at=lt.${nowIso}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, Prefer: "return=minimal" } }); } catch (_e) { /* segue */ }
  return { pessoas: peopleRows.length, identidades: identRows.length, toques: nT, fontes: { rd: (rd || []).length, whatsapp: (convs || []).length, etapas: (jr || []).length, pixel: (ev || []).length } };
}

// Vendas da planilha → toques de compra, casadas por HASH. Roda separado do rebuild (planilha grande).
async function journeyOrders(m: any) {
  const clientId = String(m.clientId || "").trim();
  if (!clientId) throw new Error("clientId obrigatório");
  const cli = (await sbGet("clients", `id=eq.${encodeURIComponent(clientId)}&select=report_orders_sheet_url,report_orders_tab`))[0] || {};
  const sid = (String(cli.report_orders_sheet_url || "").match(/\/d\/([a-zA-Z0-9-_]+)/) || [])[1];
  const tab = String(m.tab || cli.report_orders_tab || "").trim();
  if (!sid || !tab) return { erro: "cliente sem planilha de pedidos configurada" };
  // mapa hash→pessoa a partir das identidades já gravadas (e-mail/telefone viram hash em memória)
  const ident = await _sbAll("lead_identities", `client_id=eq.${encodeURIComponent(clientId)}&kind=in.(email,phone)&select=person_id,kind,value`);
  const h2p: Record<string, string> = {};
  for (const i of (ident || [])) h2p[`${i.kind}:${i.value}`] = i.person_id;
  const r = await ordersFromSheet(sid, tab, h2p, Number(m.de) || 2, Number(m.blocos) || 2);
  const rows = r.toques.map((t: any) => ({ id: "t_" + _wuid(), client_id: clientId, person_id: t._pid, ts: t.ts, kind: t.kind, channel: t.channel, label: t.label, value: t.value || 0, ref_table: t.ref_table, ref_id: t.ref_id }));
  for (let i = 0; i < rows.length; i += 400) await _sbUpsert("lead_touchpoints", rows.slice(i, i + 400), "ref_table,ref_id");
  const compras = rows.filter((x: any) => x.kind === "purchase");
  return { linhasLidas: r.lidas, casadasComLead: r.casadas, gravadas: rows.length, compras: compras.length, receita: compras.reduce((s: number, x: any) => s + (x.value || 0), 0), identidades: (ident || []).length, proxima: (r as any).proxima, fim: (r as any).fim };
}
// Tick diário: processa SÓ as linhas novas da planilha de pedidos de cada cliente (cursor em journey_sync).
// Retomável: se a planilha tiver muita linha nova, o cursor avança e o próximo tick continua de onde parou.
async function journeyOrdersTick(m: any) {
  const blocos = Number(m && m.blocos) || 2;   // ~6 mil linhas por cliente por rodada
  const clis = await _sbAll("clients", "status=eq.Ativo&select=id,name,report_orders_sheet_url,report_orders_tab");
  const sync = await _sbAll("journey_sync", "select=client_id,tab,last_row");
  const cur: Record<string, number> = {}; (sync || []).forEach((s: any) => { cur[`${s.client_id}|${s.tab}`] = s.last_row || 2; });
  const out: any[] = [];
  for (const c of (clis || [])) {
    const sid = (String(c.report_orders_sheet_url || "").match(/\/d\/([a-zA-Z0-9-_]+)/) || [])[1];
    const tab = String(c.report_orders_tab || "").trim();
    if (!sid || !tab) continue;
    const de = cur[`${c.id}|${tab}`] || 2;
    try {
      const r: any = await journeyOrders({ clientId: c.id, tab, de, blocos });
      const prox = r.proxima || de; // se acabou a planilha, mantém o cursor (as linhas novas entram depois dele)
      await _sbUpsert("journey_sync", [{ client_id: c.id, tab, last_row: r.fim ? Math.max(de, (de + (r.linhasLidas || 0))) : prox, last_run: new Date().toISOString(), resultado: { compras: r.compras, casadas: r.casadasComLead, lidas: r.linhasLidas, fim: !!r.fim } }], "client_id,tab");
      out.push({ cliente: c.name, de, lidas: r.linhasLidas, compras: r.compras, casadas: r.casadasComLead, fim: !!r.fim });
    } catch (e) { out.push({ cliente: c.name, erro: String((e as any).message || e).slice(0, 120) }); }
  }
  return { clientes: out.length, detalhe: out };
}
// Reprocessa a jornada de TODOS os clientes ativos (usado pelo botão "atualizar todos" e pelo cron)
async function journeyRebuildAll(m: any) {
  const dias = Number(m && m.dias) || 90;
  const clis = await _sbAll("clients", "status=eq.Ativo&select=id,name&order=name");
  const out: any[] = [];
  for (const c of (clis || [])) {
    try { const r = await journeyRebuild({ clientId: c.id, dias }); if (r.toques > 0) out.push({ cliente: c.name, ...r }); }
    catch (e) { out.push({ cliente: c.name, erro: String((e as any).message || e).slice(0, 120) }); }
  }
  return { clientes: out.length, detalhe: out };
}
// Diagnóstico de conexões: o que cada cliente já tem ligado e o que falta pra jornada completa
async function journeyStatus() {
  const clis = await _sbAll("clients", "status=eq.Ativo&select=id,name,meta_account_id,google_account_id,ga4_property_id,gsc_site_url,rd_config,site_url,whatsapp&order=name");
  const cfg = await _sbAll("tracking_config", "select=client_id,token,handle");
  const tokByCli: Record<string, any> = {}; (cfg || []).forEach((t: any) => { tokByCli[t.client_id] = t; });
  const ev = await _sbAll("track_events", "select=client_id");
  const evCount: Record<string, number> = {}; (ev || []).forEach((e: any) => { evCount[e.client_id] = (evCount[e.client_id] || 0) + 1; });
  const wa = await _sbAll("wa_instances", "select=client_id,status");
  const waByCli: Record<string, any> = {}; (wa || []).forEach((w: any) => { if (w.client_id) waByCli[w.client_id] = w; });
  const rd = await _sbAll("rd_conversions", "select=client_id");
  const rdCount: Record<string, number> = {}; (rd || []).forEach((r: any) => { rdCount[r.client_id] = (rdCount[r.client_id] || 0) + 1; });
  const tp = await _sbAll("lead_touchpoints", "select=client_id");
  const tpCount: Record<string, number> = {}; (tp || []).forEach((t: any) => { tpCount[t.client_id] = (tpCount[t.client_id] || 0) + 1; });
  return (clis || []).map((c: any) => ({
    id: c.id, nome: c.name,
    metaAds: !!c.meta_account_id, googleAds: !!c.google_account_id,
    ga4: !!c.ga4_property_id, searchConsole: !!c.gsc_site_url,
    pixelToken: tokByCli[c.id] ? tokByCli[c.id].token : "", pixelEventos: evCount[c.id] || 0,
    whatsapp: !!waByCli[c.id], rdEventos: rdCount[c.id] || 0,
    site: c.site_url || "", toques: tpCount[c.id] || 0,
  }));
}
// PostgREST devolve no máximo 1000 linhas por chamada — pagina até trazer tudo
async function _sbAll(table: string, query: string, max = 60000): Promise<any[]> {
  const out: any[] = [], page = 1000;
  for (let off = 0; off < max; off += page) {
    const rows = await sbGet(table, `${query}&offset=${off}&limit=${page}`);
    if (!rows || !rows.length) break;
    out.push(...rows);
    if (rows.length < page) break;
  }
  return out;
}
async function _sbUpsert(table: string, rows: any[], onConflict: string) {
  if (!rows.length) return;
  await fetch(`${_SB_URL}/rest/v1/${table}?on_conflict=${encodeURIComponent(onConflict)}`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify(rows) });
}
// ===== Banco de dados de midia (schema `midia`) — upsert com Content-Profile/Accept-Profile do PostgREST.
// Precisa do schema "midia" na lista de "Exposed schemas" do Supabase (Settings > API > Data API) — passo manual,
// nao dá pra fazer por SQL sem mexer no role authenticator (compartilhado com o app inteiro).
// return=representation pra trazer o id (uuid gerado) de volta e resolver as dimensoes sem round-trip de SELECT.
async function _midiaUpsert(table: string, rows: any[], onConflict: string): Promise<any[]> {
  if (!rows.length) return [];
  const r = await fetch(`${_SB_URL}/rest/v1/${table}?on_conflict=${encodeURIComponent(onConflict)}`, {
    method: "POST",
    headers: {
      apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json",
      "Content-Profile": "midia", "Accept-Profile": "midia",
      Prefer: "resolution=merge-duplicates,return=representation",
    },
    body: JSON.stringify(rows),
  });
  if (!r.ok) throw new Error(`midia.${table}: HTTP ${r.status} ${(await r.text()).slice(0, 300)}`);
  return await r.json();
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
- 1ª linha: título com emoji + *negrito*. Na 2ª linha, o período analisado em _itálico_ (ex: _últimos 7 dias (11/08 a 18/08)_) — sem isso quem lê compara com outro intervalo e acha que o número está errado. Logo abaixo, uma linha divisória: ${WA_DIV}
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
/* Janela dos avisos automáticos num lugar só. O aviso saía sem dizer o período, e quem lia comparava
   com o mês no painel e achava que o número estava errado. */
const _AVISO_DIAS = 7;
// data pronta em vez de pedir pra IA calcular: assim o período escrito no aviso é sempre o real
function _avisoPeriodo() {
  const br = (d: Date) => `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  const ate = new Date(), de = new Date(Date.now() - _AVISO_DIAS * 864e5);
  return `últimos ${_AVISO_DIAS} dias (${br(de)} a ${br(ate)})`;
}
async function waAutoText(tipo: string, escopo = "padrao", prompt = "", nivel = "resumido"): Promise<string[]> {
  const escNota = escopo && escopo !== "padrao" ? ` Considere apenas os clientes do escopo "${ESCOPO_LABEL[escopo] || escopo}".` : "";
  // Aviso que a IA não conseguiu montar (cota estourada, por exemplo) não vira mensagem no grupo: mandar
  // "Não consegui montar esse aviso agora" só gasta a atenção de quem lê, sem informar nada.
  if (tipo === "custom") { const p = String(prompt || "").trim(); if (!p) return []; const t = await waAgentOneShot(p + escNota); return t ? [t] : []; }
  if (tipo === "resumo7") return await waAgentAllClientsSummary(7, escopo, nivel);
  if (tipo === "resumo30") return await waAgentAllClientsSummary(30, escopo, nivel);
  if (tipo === "restricoes") return await waAgentAllClientsSummary(7, "com_restricao");
  if (tipo === "receber") return [_waFmtFinanceiro(await waFinanceiro({ tipo: "receita", status: "pendente", mes: _mesAtual() }), "💰 *A receber este mês*")];
  if (tipo === "pagar") return [_waFmtFinanceiro(await waFinanceiro({ tipo: "despesa", status: "pendente", mes: _mesAtual() }), "💸 *A pagar este mês*")];
  if (tipo === "pendencias") return [await waPendenciasText()];
  if (tipo === "atencao") return [await waAgentOneShot(`Quem precisa de atenção? Analise os clientes ativos (use resumo_todos_clientes e, se precisar, meta_insights por cliente) e destaque só os que estão abaixo da meta, gastando sem resultado, ou parados.

PERÍODO (obrigatório): chame resumo_todos_clientes com dias=${_AVISO_DIAS} e ESCREVA o período analisado logo abaixo do título, exatamente assim: "_${_avisoPeriodo()}_". Todo número do aviso tem que ser desse mesmo período — nunca misture com mês fechado nem com outro intervalo, senão o gestor compara com o painel e os valores não batem.

OBJETIVO DE CADA CLIENTE (regra dura): cada cliente tem um objetivo diferente (venda, leads, mensagens/WhatsApp, tráfego, visualizações, alcance...) e a métrica que o resumo traz já é a certa pra ele. NUNCA cite "sem compra"/"0 compras"/"0 conversões"/ROAS baixo como problema de quem não tem venda como objetivo. E mais: NÃO INCLUA na lista o cliente cujo objetivo não é conversão só porque ele não tem conversão — isso não é baixo resultado, é o objetivo dele. Ele só entra se a métrica DO OBJETIVO dele estiver ruim (ex: custo por visualização alto, CTR despencando, custo por conversa subindo) ou se estiver gastando e sem entrega nenhuma. Ao citar um cliente, diga o objetivo dele e a métrica desse objetivo, não a que falta.

Se estiver tudo bem, diga que está tudo em ordem. Curto.${escNota}`)];
  if (tipo === "recomendacoes") return [await waAgentOneShot(`Recomendações da semana: com base nos dados reais dos clientes ativos, liste 2 a 3 ações priorizadas (o que pausar, escalar ou ajustar), citando o cliente. Considere o objetivo de cada cliente (venda, leads, mensagens, tráfego...) — não recomende nada baseado em "sem venda"/ROAS de quem não tem objetivo de venda. Curto e prático.${escNota}`)];
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
// Monitor de conectividade: detecta quando um WhatsApp cai → notificação no sino + aviso da AndréIA no grupo.
function _fmtFone(p: any) { p = String(p || "").replace(/[^0-9]/g, ""); if (!p) return ""; const m = p.match(/^(\d{2})(\d{2})(\d{4,5})(\d{4})$/); return m ? `+${m[1]} (${m[2]}) ${m[3]}-${m[4]}` : p; }
async function waConnectivityCheck() {
  const insts = await sbGet("wa_instances", "select=id,name,uaz_host,uaz_token,status,phone,connected_at,health_fail_count,health_last_alert_at,health_last_ok_at,health_last_recovery_at");
  const team = await sbGet("team", "select=id");
  const g: any = await _andreiaGroupInst();
  const caidos: string[] = [], recuperados: string[] = [], oscilando: string[] = [], nuncaConectaram: string[] = [];
  for (const inst of (insts || [])) {
    if (!inst.uaz_host || !inst.uaz_token) continue;
    let cur: string | null = null;
    try { const { j } = await waCall(inst.uaz_host, inst.uaz_token, "/instance/status"); cur = (j && j.instance && j.instance.status) || null; } catch { continue; }
    if (!cur) continue;
    const was = inst.status;
    const now = new Date(), nowIso = now.toISOString();
    if (cur !== "connected") {
      const fails = (Number(inst.health_fail_count) || 0) + 1;
      await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(inst.id)}`, { status: cur, health_fail_count: fails, updated_at: nowIso });
      // Uma falha isolada costuma ser apenas oscilação de sessão/rede. Confirma queda somente na 2ª leitura consecutiva.
      if (fails < 2) { oscilando.push(inst.name || inst.phone || inst.id); continue; }
      // Instância que NUNCA ficou OK pro monitor não "caiu": é cadastro pendente de leitura do QR. A KWAN acumulou
      // 458 falhas seguidas, sem um único OK, e gerou 14 avisos de queda. Isso não é alerta, é tarefa de configuração.
      // (connected_at sozinho não serve: ele é gravado na tentativa de conexão, mesmo quando o QR nunca é lido.)
      if (!inst.health_last_ok_at && fails >= 10) { nuncaConectaram.push(inst.name || inst.phone || inst.id); continue; }
      const lastAlert = inst.health_last_alert_at ? new Date(inst.health_last_alert_at).getTime() : 0;
      // Cooldown evita repetir o mesmo alerta a cada execução do cron enquanto a sessão continua instável.
      if (lastAlert && now.getTime() - lastAlert < 6 * 3600e3) continue;
      const fone = _fmtFone(inst.phone); const quem = inst.name || fone || "instância";
      const title = `🔴 WhatsApp desconectado: ${quem}`;
      const detail = `O número ${fone || "(sem número)"} (${inst.name || "instância"}) caiu (${cur}). Reconecte em Configurações → WhatsApp pra não perder mensagens.`;
      // Enquanto a instância continua caída, o cooldown de 6h fazia o MESMO aviso voltar pro sino pra sempre
      // (a KWAN, caída desde 11/08, acumulou 14 avisos iguais e afogou o resto). Agora: um aviso por instância
      // até alguém ler. Depois de lido, se ainda estiver caída, volta a avisar — aí sim como lembrete útil.
      for (const t of (team || [])) {
        try {
          const jaTem = await sbGet("notifications", `to_team=eq.${encodeURIComponent(t.id)}&type=eq.wa_disconnect&read=eq.false&task_name=eq.${encodeURIComponent(title)}&select=id&limit=1`);
          if (jaTem.length) continue;
          await sbPost("notifications", { id: _wuid(), to_team: t.id, from_team: "sistema", task_id: null, task_name: title, comment_text: detail, read: false, type: "wa_disconnect" });
        } catch (_e) { /* */ }
      }
      if (!g.erro) { try { await waCall(g.inst.uaz_host, g.inst.uaz_token, "/send/text", "POST", { number: g.group, text: `🔴 *WhatsApp desconectado*\n${WA_DIV}\n*${quem}*${fone ? ` (${fone})` : ""} caiu.\nReconecte em Configurações → WhatsApp pra não perder mensagens. 📲` }); } catch (_e) { /* */ } }
      await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(inst.id)}`, { health_last_alert_at: nowIso });
      caidos.push(quem);
    } else {
      const hadConfirmedAlert = !!inst.health_last_alert_at && (!inst.health_last_recovery_at || new Date(inst.health_last_recovery_at).getTime() < new Date(inst.health_last_alert_at).getTime());
      await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(inst.id)}`, { status: "connected", health_fail_count: 0, health_last_ok_at: nowIso, ...(was !== "connected" ? { connected_at: nowIso } : {}) });
      if (hadConfirmedAlert) {
        const quem = inst.name || _fmtFone(inst.phone) || "instância";
        // Recupera o histórico recente que pode ter ficado sem webhook durante a queda.
        let recovered = 0; try { const sync: any = await waHandler({ op: "poll", instanceId: inst.id, sinceDays: 2, limit: 3000 }); recovered = Number(sync?.added) || 0; } catch (_e) { /* sincronização manual continua disponível */ }
        if (!g.erro) { try { await waCall(g.inst.uaz_host, g.inst.uaz_token, "/send/text", "POST", { number: g.group, text: `🟢 *WhatsApp restabelecido*\n${WA_DIV}\n*${quem}* voltou a conectar.${recovered ? `\nRecuperei *${recovered} mensagem(ns)* que estavam pendentes.` : "\nO histórico recente foi conferido automaticamente."}` }); } catch (_e) { /* */ } }
        await sbPatchD("wa_instances", `id=eq.${encodeURIComponent(inst.id)}`, { health_last_recovery_at: nowIso });
        recuperados.push(quem);
      }
    }
  }
  return { checked: (insts || []).length, caidos, recuperados, oscilando, nuncaConectaram };
}
// Poll server-side da(s) instancia(s) da AGENCIA (client_id null) que nao tem webhook proprio (o numero da
// agencia tem o webhook ocupado por outro sistema externo - ver memoria whatsapp-uazapi). Sem isso, a
// classificacao por IA dessas conversas so rodava quando alguem tinha a aba CRM aberta no navegador (o poll
// era so client-side, setInterval 20s). Exclui a instancia da AndreIA (grupo) - essa ja recebe tudo via
// webhook proprio (forward pro waAgentHandle), nao precisa de poll.
async function waAgencyPollTick() {
  const cfg = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  const andreiaInstId = (cfg.andreia_wa || {}).instance_id || null;
  // Instâncias da agência: a cada rodada (2 min). Instâncias de CLIENTE: a cada 10 min — elas recebem por
  // webhook (chega em ~6s), mas se o webhook falha o buraco só era tapado quando alguém abria o CRM na tela;
  // foi o que deixou um cliente com 3,5 dias de atraso. Aqui é a rede de segurança do lado do servidor.
  const comClientes = new Date().getUTCMinutes() % 10 === 0;
  const insts = await sbGet("wa_instances", `${comClientes ? "" : "client_id=is.null&"}status=eq.connected&select=id,name`);
  const alvos = (insts || []).filter((i: any) => i.id !== andreiaInstId);
  const resultados: any[] = [];
  for (const inst of alvos) {
    try { const r = await waHandler({ op: "poll", instanceId: inst.id }); resultados.push({ instancia: inst.name, ...r }); }
    catch (e) { resultados.push({ instancia: inst.name, erro: String((e as any)?.message || e) }); }
  }
  return { instanciasChecadas: alvos.length, resultados };
}
// Rotina de seguranca: roda public.security_audit() (RLS aberta pro anon, tabela sem RLS, view sem
// security_invoker) via cron diario. So avisa (sino + grupo da AndréIA) quando aparece um achado NOVO,
// pra nao repetir aviso todo dia do mesmo problema ja conhecido - compara com o que ficou salvo da ultima rodada.
async function securityAuditTick() {
  const r = await fetch(`${_SB_URL}/rest/v1/rpc/security_audit`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: "{}" });
  const rows: any[] = r.ok ? await r.json() : [];
  const since24 = new Date(Date.now() - 864e5).toISOString(), events = await sbGet("security_events", `created_at=gte.${encodeURIComponent(since24)}&select=kind&limit=2000`);
  if (events.length >= 20) {
    const by: Record<string, number> = {}; events.forEach((e: any) => { by[e.kind || "outro"] = (by[e.kind || "outro"] || 0) + 1; });
    rows.push({ check_key: `attack_spike:${new Date().toISOString().slice(0, 10)}`, severity: "HIGH", detail: `${events.length} tentativas bloqueadas nas últimas 24h (${Object.entries(by).map(([k,v]) => `${k}: ${v}`).join(", ")})` });
  }
  // Minimização LGPD: telemetria de segurança (já sem IP em texto) expira em 90 dias.
  const cutoff90 = new Date(Date.now() - 90 * 864e5).toISOString();
  try { await fetch(`${_SB_URL}/rest/v1/security_events?created_at=lt.${encodeURIComponent(cutoff90)}`, { method: "DELETE", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, Prefer: "return=minimal" } }); } catch { /* */ }
  const current = rows.map((x) => x.check_key);
  const cfg = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  const prev: string[] = cfg.security_audit_issues || [];
  const novos = rows.filter((x) => !prev.includes(x.check_key));
  const resolvidos = prev.filter((k) => !current.includes(k));
  if (novos.length) {
    const team = await sbGet("team", "select=id");
    const title = `🚨 ${novos.length} problema(s) de segurança encontrado(s) no banco`;
    const detail = novos.map((x) => `[${x.severity}] ${x.detail}`).join("\n");
    for (const t of (team || [])) { try { await sbPost("notifications", { id: _wuid(), to_team: t.id, from_team: "sistema", task_id: null, task_name: title, comment_text: detail.slice(0, 1500), read: false, type: "security_alert" }); } catch (_e) { /* */ } }
    const g: any = await _andreiaGroupInst();
    if (!g.erro) { try { await waCall(g.inst.uaz_host, g.inst.uaz_token, "/send/text", "POST", { number: g.group, text: `🚨 *Alerta de segurança*\n${WA_DIV}\n${novos.map((x) => `• ${x.detail}`).join("\n")}\n\nVerifique o quanto antes.` }); } catch (_e) { /* */ } }
  }
  await sbPatchD("account_config", "id=eq.main", { data: { ...cfg, security_audit_issues: current, security_audit_last_run: new Date().toISOString() } });
  return { total: current.length, novos: novos.length, resolvidos: resolvidos.length };
}
// ===== AGENTE DE SAUDE DO SISTEMA (diario, apos as sincronizacoes) =====
// Junta seguranca + instabilidade + erros numa fotografia unica: falhas de cron, taxa de erro das chamadas
// internas (teria pego o apagao do guard em minutos), frescor dos canais e contas de midia paradas.
// Resultado fica em account_config.data.health_report (painel geral le de la); sino so quando surge problema NOVO.
/* ===== MONITOR DE CRÉDITO DE IA =====
   Nem a OpenAI nem o Google expõem SALDO pela chave da API. O que dá pra fazer sem credencial nova:
   acompanhar o GASTO do mês contra um teto que a gestora define, e o consumo do dia contra o limite
   do plano gratuito do Gemini. Resolve o que importa — avisar ANTES de parar — em vez de descobrir
   quando a IA já morreu. Avisa uma vez por faixa (50/80/100%) por mês, no sino e no WhatsApp. */
async function iaCreditoTick() {
  const r = await fetch(`${_SB_URL}/rest/v1/rpc/ai_gasto_mes`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: "{}" });
  if (!r.ok) throw new Error("ai_gasto_mes falhou: HTTP " + r.status);
  const gastos: any[] = await r.json();
  const watch = await sbGet("ai_credit_watch", "select=*");
  const mesAtual = new Date().toISOString().slice(0, 7);
  const hoje = new Date().toISOString().slice(0, 10);
  const avisos: string[] = [];
  for (const w of watch) {
    const g = gastos.find((x: any) => x.provider === w.provider) || { usd: 0, requests: 0, requests_hoje: 0 };
    const ult = w.ultimo_aviso || {};
    const novo: any = { ...ult };
    // teto de gasto do mes
    const teto = Number(w.teto_usd_mes) || 0;
    if (teto > 0) {
      const pct = Number(g.usd) * 100 / teto;
      const faixas = (w.avisar_em || [50, 80, 100]).slice().sort((a: number, b: number) => b - a);
      const bateu = faixas.find((f: number) => pct >= f);
      if (bateu && ult[`gasto_${mesAtual}`] !== bateu) {
        novo[`gasto_${mesAtual}`] = bateu;
        avisos.push(bateu >= 100
          ? `🔴 *${w.provider.toUpperCase()}* estourou o teto do mês: US$ ${Number(g.usd).toFixed(2)} de US$ ${teto.toFixed(2)}. Se a chave ficar sem crédito, a IA para.`
          : `🟡 *${w.provider.toUpperCase()}* já usou ${Math.round(pct)}% do teto do mês: US$ ${Number(g.usd).toFixed(2)} de US$ ${teto.toFixed(2)}.`);
      }
    }
    // limite diario do plano gratuito
    const lim = Number(w.limite_dia_requests) || 0;
    if (lim > 0) {
      const pctD = Number(g.requests_hoje) * 100 / lim;
      if (pctD >= 80 && ult[`dia_${hoje}`] !== true) {
        novo[`dia_${hoje}`] = true;
        avisos.push(`🟡 *${w.provider.toUpperCase()}* já fez ${g.requests_hoje} das ${lim} chamadas do plano gratuito de hoje (${Math.round(pctD)}%). Passando disso, ele para até amanhã.`);
      }
    }
    if (JSON.stringify(novo) !== JSON.stringify(ult)) {
      await sbPatchD("ai_credit_watch", `provider=eq.${encodeURIComponent(w.provider)}`, { ultimo_aviso: novo, atualizado_em: new Date().toISOString() });
    }
  }
  if (!avisos.length) return { ok: true, avisos: 0, gastos };
  const texto = `💳 *Crédito de IA*\n${WA_DIV}\n\n${avisos.join("\n\n")}\n\n_Sem crédito nos dois provedores, o sistema fica sem IA: AndréIA, DNA, CRM e leitura de print._`;
  try { const g = await _andreiaGroupInst(); if (!(g as any).erro) await _sendGroup(g, [texto]); } catch (_e) { /* sino ainda vale */ }
  try {
    const team = await sbGet("team", "select=id,name");
    for (const m of team) {
      await sbPost("notifications", { id: "ntf" + Math.random().toString(36).slice(2, 11), to_team: m.id, from_team: team[0]?.id || m.id,
        task_id: null, task_name: "💳 Crédito de IA acabando", comment_text: avisos.join(" · ").replace(/\*/g, ""), read: false, type: "system" });
    }
  } catch (_e) { /* nao pode derrubar o tick */ }
  return { ok: true, avisos: avisos.length, gastos };
}
async function systemHealthTick() {
  const r = await fetch(`${_SB_URL}/rest/v1/rpc/system_health_snapshot`, { method: "POST", headers: { apikey: _SB_KEY, Authorization: `Bearer ${_SB_KEY}`, "Content-Type": "application/json" }, body: "{}" });
  if (!r.ok) throw new Error("system_health_snapshot falhou: HTTP " + r.status);
  const snap: any = await r.json();
  const hoje = new Date().toISOString().slice(0, 10);
  const ontem = new Date(Date.now() - 864e5).toISOString().slice(0, 10);
  type Check = { key: string; sev: "alta" | "media" | "baixa"; msg: string };
  const checks: Check[] = [];
  for (const f of (snap.cron_falhas_24h || [])) checks.push({ key: `cron:${f.job}`, sev: "alta", msg: `Automação "${f.job}" falhou ${f.falhas}x nas últimas 24h` });
  const errs = (snap.http_erros_24h || []).reduce((s: number, e: any) => s + (e.qtd || 0), 0);
  const tot = snap.http_total_24h || 0;
  if (tot >= 20 && errs / tot > 0.2) checks.push({ key: `http:taxa:${hoje}`, sev: "alta", msg: `${Math.round((errs / tot) * 100)}% das chamadas internas com erro nas últimas 24h (${errs} de ${tot}) — ${(snap.http_erros_24h || []).map((e: any) => `HTTP ${e.status}: ${e.qtd}`).join(", ")}` });
  else if (errs >= 30) checks.push({ key: `http:volume:${hoje}`, sev: "media", msg: `${errs} chamadas internas com erro nas últimas 24h (${(snap.http_erros_24h || []).map((e: any) => `HTTP ${e.status}: ${e.qtd}`).join(", ")})` });
  // O coletor (channel-metrics-daily) roda 04:15 UTC e grava o dia ANTERIOR. Checar contra "ontem" antes disso
  // acusava canal parado todo dia de madrugada (quem apertasse "Verificar agora" as 23h daqui via alerta falso).
  // Depois das 05:00 UTC a cobranca volta a ser de ontem; antes disso, so cobra a partir de anteontem.
  const anteontem = new Date(Date.now() - 2 * 864e5).toISOString().slice(0, 10);
  const limiteSync = new Date().getUTCHours() >= 5 ? ontem : anteontem;
  for (const c of (snap.sync_canais || [])) if (c.ultima_data && c.ultima_data < limiteSync) checks.push({ key: `sync:${c.canal}`, sev: "media", msg: `Canal ${c.canal} sem dado novo desde ${c.ultima_data} (a sincronização diária pode estar falhando)` });
  const paradas = snap.contas_midia_paradas || [];
  // so entram contas que investiram nos ultimos 30 dias (filtro no system_health_snapshot): conta
  // encerrada ha meses nunca mais tem dado novo e ficava listada pra sempre.
  if (paradas.length) checks.push({ key: `midia:paradas`, sev: "baixa", msg: `${paradas.length} conta(s) de mídia que estavam rodando sem dado novo há 3+ dias (pode ser pausa proposital): ${paradas.slice(0, 6).map((p: any) => `${p.cliente} (${p.plataforma}, último dado ${p.ultima_data}, ${_fmtR(Number(p.gasto_30d) || 0)} em 30d)`).join("; ")}${paradas.length > 6 ? "…" : ""}` });
  for (const s of (snap.seguranca_eventos_24h || [])) if (s.tipo === "unauthorized_internal_route" && s.qtd > 0) checks.push({ key: `sec:unauthorized`, sev: "media", msg: `${s.qtd} tentativa(s) de acesso a rota interna sem credencial nas últimas 24h` });
  // problemas da auditoria de seguranca do banco (ja rodada as 05:00) entram no mesmo painel
  const cfg = (await sbGet("account_config", "id=eq.main&select=data"))[0]?.data || {};
  for (const k of (cfg.security_audit_issues || [])) checks.push({ key: `audit:${k}`, sev: "alta", msg: `Auditoria de segurança do banco: ${k}` });
  const status = checks.some((c) => c.sev === "alta") ? "critico" : checks.length ? "atencao" : "ok";
  const prevKeys: string[] = (cfg.health_report?.checks || []).filter((c: any) => c.sev === "alta").map((c: any) => c.key);
  const novosAltos = checks.filter((c) => c.sev === "alta" && !prevKeys.includes(c.key));
  if (novosAltos.length) {
    const team = await sbGet("team", "select=id");
    const title = `🩺 Saúde do sistema: ${novosAltos.length} problema(s) novo(s)`;
    const detail = novosAltos.map((c) => `• ${c.msg}`).join("\n");
    for (const t of (team || [])) { try { await sbPost("notifications", { id: _wuid(), to_team: t.id, from_team: "sistema", task_id: null, task_name: title, comment_text: detail.slice(0, 1500), read: false, type: "health_alert" }); } catch (_e) { /* */ } }
    // Regra do grupo do WhatsApp: só falha grave, invasão e segurança — nunca entrega/rotina, que fica no sino.
    // Os achados "audit:*" JÁ foram enviados pela rotina de segurança 30 min antes; repetir aqui gerava dois
    // avisos com o mesmo conteúdo no grupo. No sino eles continuam, porque lá a visão é consolidada.
    const proWhats = novosAltos.filter((c) => !c.key.startsWith("audit:"));
    if (proWhats.length) try {
      const g: any = await _andreiaGroupInst();
      if (!g.erro) await waCall(g.inst.uaz_host, g.inst.uaz_token, "/send/text", "POST", { number: g.group, text: `🩺 *Falha grave no sistema*\n${WA_DIV}\n${proWhats.map((c) => `• ${c.msg}`).join("\n")}\n\nDetalhes no painel geral → card de saúde.` });
    } catch (_e) { /* aviso no sino ja foi */ }
  }
  await sbPatchD("account_config", "id=eq.main", { data: { ...cfg, health_report: { gerado_em: new Date().toISOString(), status, checks, resumo: { erros_http_24h: errs, chamadas_24h: tot, seguranca_24h: snap.seguranca_eventos_24h || [], sync_canais: snap.sync_canais || [] } } } });
  return { status, problemas: checks.length, novosAlertas: novosAltos.length };
}
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
    if (body.accessControl) {
      const r = await accessControl(body.accessControl, req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.accountConfigAccess) {
      const r = await accountConfigAccess(body.accountConfigAccess, req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    await _guardUserRequest(body, req.headers.get("Authorization") || "");
    if (body.wa) {
      const r = await waHandler(body.wa);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.waAgent) {
      /* Se a IA falha aqui, quem mandou a mensagem no grupo nao recebe NADA - foi o que aconteceu com
         "Criar tarefa..." quando a cota do Gemini estourou. Silencio e o pior retorno possivel: a
         pessoa acha que foi feito. Agora a falha vira resposta no proprio grupo. */
      try {
        const r = await waAgentHandle(body.waAgent);
        return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
      } catch (e) {
        const motivo = (e as Error)?.message || "erro desconhecido";
        console.error("[waAgent] falhou:", motivo);
        try { await _waAgentAvisaErro(body.waAgent, motivo); } catch (_e) { /* nem o aviso pode derrubar */ }
        return new Response(JSON.stringify({ data: { erro: motivo } }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
      }
    }
    if (body.iaCreditoTick) {
      const r = await iaCreditoTick();
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
    if (body.waConnCheck) {
      const r = await waConnectivityCheck();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.waAgencyPollTick) {
      const r = await waAgencyPollTick();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.securityAuditTick) {
      const r = await securityAuditTick();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.systemHealthTick) {
      const r = await systemHealthTick();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.crmAndreia) {
      const r = await crmAndreia(body.crmAndreia);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.crmCapaAudit) {
      const r = await crmCapaAudit(body.crmCapaAudit);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.crmCapaDashboard) {
      const r = await crmCapaDashboard(body.crmCapaDashboard);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.crmAndreiaAction) {
      const r = await crmAndreiaAction(body.crmAndreiaAction);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingSugerirCampos) {
      const r = await briefingSugerirCampos(body.briefingSugerirCampos);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingAnalise) {
      const r = await briefingAnalise(body.briefingAnalise);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingPreparar) {
      const r = await briefingPreparar(body.briefingPreparar);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingRanking) {
      const r = await briefingRanking(body.briefingRanking);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingCreativoAnalise) {
      const r = await briefingCreativoAnalise(body.briefingCreativoAnalise);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingGerarFichas) {
      const r = await briefingGerarFichas(body.briefingGerarFichas);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingAprovar) {
      const r = await briefingAprovar(body.briefingAprovar);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingHistorico) {
      const r = await briefingHistorico(body.briefingHistorico);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingCompleto) {
      const r = await briefingCompleto(body.briefingCompleto);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.instagramDiag) {
      const r = await _instagramDiag();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.instagramAutoMatch) {
      const r = await instagramAutoMatch(body.instagramAutoMatch);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.instagramListAccounts) {
      const r = await instagramListAccounts();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.instagramFollowersSnapshot) {
      const r = await instagramFollowersSnapshot(body.instagramFollowersSnapshot === true ? {} : body.instagramFollowersSnapshot);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.instagramOrganicSnapshot) {
      const r = await instagramOrganicSnapshot(body.instagramOrganicSnapshot);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.instagramOrganicContent) {
      const r = await instagramOrganicContent(body.instagramOrganicContent);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.briefingCuradoria) {
      const r = await briefingCuradoria(body.briefingCuradoria);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.creativeMiner) {
      const r = await creativeMiner(body.creativeMiner);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.eventReports) {
      const r = await eventReports(body.eventReports);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.resolveAllOrigins) {
      const r = await waResolveAllOrigins();
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
    if (body.googleCampaignAction) {
      const r = await googleCampaignAction(body.googleCampaignAction);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleTermAction) {
      const r = await googleTermAction(body.googleTermAction);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.ga4Report) {
      const r = await ga4Report(body.ga4Report);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyQualityPixel) {
      const r = await journeyQualityPixel(body.journeyQualityPixel);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.gscReport) {
      const r = await gscReport(body.gscReport);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.gscSites) {
      const r = await gscSites();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.gsaEmail) {
      return new Response(JSON.stringify({ data: { email: _gsaEmail() } }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    // Só os CABEÇALHOS de uma aba (linha 1). Nome de coluna não é dado pessoal — serve pra mapear a planilha
    // sem trazer nenhuma linha de dado. Por isso não passa pela trava de abas bloqueadas.
    if (body.sheetHeaders) {
      const keyJson = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_KEY");
      if (!keyJson) throw new Error("GOOGLE_SERVICE_ACCOUNT_KEY nao configurada");
      const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(keyJson), scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"] });
      const sheets = google.sheets({ version: "v4", auth });
      const sid = String(body.sheetHeaders.spreadsheetId || "");
      const out: any = {};
      if (!body.sheetHeaders.tabs) { // sem abas: lista os nomes das abas
        const meta = await sheets.spreadsheets.get({ spreadsheetId: sid });
        out._abas = (meta.data.sheets || []).map((s: any) => s.properties.title);
      } else {
        for (const tab of body.sheetHeaders.tabs) {
          try { const r = await sheets.spreadsheets.values.get({ spreadsheetId: sid, range: `'${tab}'!A1:AZ1` }); out[tab] = (r.data.values && r.data.values[0]) || []; }
          catch (e) { out[tab] = { error: (e as Error).message }; }
        }
      }
      return new Response(JSON.stringify({ data: out }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.pixelCheck) {
      const r = await pixelCheck(body.pixelCheck);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyOrdersTick) {
      const r = await journeyOrdersTick(body.journeyOrdersTick);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyOrdersImport) {
      const r = await journeyOrdersImport(body.journeyOrdersImport);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyOrders) {
      const r = await journeyOrders(body.journeyOrders);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyRebuildAll) {
      const r = await journeyRebuildAll(body.journeyRebuildAll);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyStatus) {
      const r = await journeyStatus();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.journeyRebuild) {
      const r = await journeyRebuild(body.journeyRebuild);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleKeywordAction) {
      const r = await googleKeywordAction(body.googleKeywordAction);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleTermMining) {
      const r = await googleTermMining(body.googleTermMining);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    // Gera o contrato (teste/uso direto pelo sistema): devolve a URL do PDF sem enviar no WhatsApp
    if (body.contratoPdf) {
      const { text, faltando } = await _contratoTexto(body.contratoPdf);
      const bytes = await _contratoPdf(text);
      const url = await _uploadDoc(bytes, `contrato-${_slugDoc(body.contratoPdf.razaoSocial || body.contratoPdf.nome)}-${Date.now()}.pdf`);
      return new Response(JSON.stringify({ data: { url, faltando } }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    // Áudio (gravado no navegador) → texto, pra ditar a orientação da IA (garimpo/limpeza)
    if (body.transcribe) {
      const t = await waTranscribe(String(body.transcribe.dataUrl || "")).catch(() => "");
      return new Response(JSON.stringify({ data: { text: t } }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    // Anexo (imagem/PDF/texto) → extrai o conteúdo relevante como texto, pra orientar a IA
    if (body.extractAttachment) {
      const t = await _extractAttachmentText(String(body.extractAttachment.dataUrl || ""), String(body.extractAttachment.mime || ""), String(body.extractAttachment.name || "")).catch((e) => "erro: " + (e?.message || e));
      return new Response(JSON.stringify({ data: { text: t } }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleAdGroups) {
      const r = await googleAdGroups(body.googleAdGroups);
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
    if (body.googleAudiences) {
      const r = await googleAudiences(body.googleAudiences);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleCreateCustomAudience) {
      const r = await googleCreateCustomAudience(body.googleCreateCustomAudience);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.tiktokAds) {
      const r = await tiktokAdsInsights(body.tiktokAds);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.tiktokAccounts) {
      const r = await tiktokListAccounts(String(body.tiktokAccounts.clientId || ""));
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.pinterestAds) {
      const r = await pinterestAdsInsights(body.pinterestAds);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.pinterestAccounts) {
      const r = await pinterestListAccounts(String(body.pinterestAccounts.clientId || ""));
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.channelMetricsCollect) {
      const r = await channelMetricsCollect(body.channelMetricsCollect);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAds) {
      const r = await metaAdsInsights(body.metaAds);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaTokenUpdate) {
      const r = await metaTokenUpdate(body.metaTokenUpdate, req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleOAuthStart) {
      const r = await googleOAuthStart(req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.youtubeOAuth) {
      const r = await youtubeOAuth(body.youtubeOAuth, req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleOAuthDisconnect) {
      const r = await googleOAuthDisconnect(req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.googleGa4Discover) {
      const r = await googleGa4Discover(req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.apifyConfig) {
      const r = await apifyConfig(body.apifyConfig, req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.waCloudConfig) {
      const r = await waCloudConfig(body.waCloudConfig, req.headers.get("Authorization") || "");
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAccounts) {
      const r = await metaListAccounts();
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAssetsSync) {
      const r = await metaAssetsSync(body.metaAssetsSync);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaPixels) {
      const r = await metaListPixels(body.metaPixels);
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
    if (body.metaAudiences) { const r = await metaAudiences(body.metaAudiences); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.metaAudienceSources) { const r = await metaAudienceSources(body.metaAudienceSources); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.metaAudienceMedia) { const r = await metaAudienceMedia(body.metaAudienceMedia); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.metaCreateAudiences) { const r = await metaCreateAudiences(body.metaCreateAudiences); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.metaCreateCustomList) { const r = await metaCreateCustomList(body.metaCreateCustomList); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.metaCreateSavedAudience) { const r = await metaCreateSavedAudience(body.metaCreateSavedAudience); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.metaTargetingSearch) { const r = await metaTargetingSearch(body.metaTargetingSearch); return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } }); }
    if (body.dnaExtract) {
      let text = body.dnaExtract.text || "";
      if (!text && body.dnaExtract.url) text = await fetchUrlText(body.dnaExtract.url);
      // coleta com problema (banco lento, conta fora do ar) nao pode virar 500 seco: cai na mensagem
      // explicativa logo abaixo, que diz o que fazer.
      if (!text && body.dnaExtract.clientId) text = await _dnaGatherFromAccount(body.dnaExtract.clientId).catch(() => "");
      if (!text || text.replace(/\s/g, "").length < 60) return new Response(JSON.stringify({ error: "Sem informação suficiente na conta pra montar o DNA. Preencha o site do cliente, tenha anúncios ativos, ou cole um material (PDF/texto)." }), { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } });
      const r = await extractDna(text, body.dnaExtract.direcionamento || "");
      if (body.dnaExtract.clientId && _dnaFontes.length) r._fontes = _dnaFontes;
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
    // Sem isto, uma falha some: a tela mostra "non-2xx" e os logs não guardam nada do motivo.
    console.error("[dynamic-responder] falhou:", (e as Error)?.message, (e as Error)?.stack?.slice(0, 600));
    return new Response(JSON.stringify({ error: (e as Error).message }), { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } });
  }
});
