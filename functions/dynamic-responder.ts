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

async function generateAnalysis(m: any, chat: any[], styleExamples: string[]) {
  let system = `Voce e uma gestora de trafego pago senior, especialista em performance (Meta Ads, Google Ads, funil de vendas e e-commerce). Escreve analises gerenciais mensais claras, diretas e acionaveis. Baseie-se SEMPRE nos numeros reais fornecidos, nunca invente dados. Responda apenas com o texto da analise, em portugues, sem markdown e sem titulos, em 2 a 4 paragrafos curtos.`;
  if (Array.isArray(styleExamples) && styleExamples.length) {
    system += `\n\nO gestor humano tem um estilo proprio de escrever. Imite o tom, o tamanho e a estrutura destes exemplos de analises anteriores dele:\n` + styleExamples.map((s, i) => `--- Exemplo ${i + 1} ---\n${s}`).join("\n\n");
  }
  const messages: any[] = [{ role: "system", content: system }];
  messages.push({ role: "user", content: `Dados do mes para o cliente "${m.clientName}", referente a ${m.mesLabel}:\n${JSON.stringify(m, null, 2)}\n\nGere a analise gerencial mensal.` });
  if (Array.isArray(chat)) for (const t of chat) messages.push({ role: t.role === "user" ? "user" : "assistant", content: String(t.text || "") });
  const json = await callOpenAI({ messages });
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
];

async function runAgent(a: any) {
  let system = `Voce e a AndreIA, gestora de trafego pago senior de uma agencia de performance. Domina Meta Ads, Google Ads, funil de vendas, CRO (otimizacao de paginas) e analise de dados, no nivel dos melhores gestores de trafego do Brasil.

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

Voce tambem pode EXECUTAR acoes quando o gestor pedir explicitamente: criar tarefas e concluir (dar baixa em) tarefas, chamando as funcoes disponiveis. O sistema mostra um card de confirmacao antes de executar, entao apenas proponha a acao chamando a funcao; nao afirme que ja executou. Mas seu valor principal e a analise tecnica de otimizacao.`;

  if (Array.isArray(a.knowledge) && a.knowledge.length) {
    system += `\n\nBASE DE CONHECIMENTO (metodos e frameworks de gestores de trafego que a agencia quer que voce siga). Use estes principios como referencia nas suas recomendacoes:\n` +
      a.knowledge.map((k: any, i: number) => `--- Fonte ${i + 1}: ${k.title || "material"} ---\n${String(k.text || "").slice(0, 6000)}`).join("\n\n");
  }

  const messages: any[] = [{ role: "system", content: system }];
  messages.push({ role: "user", content: `Snapshot atual (dados reais do sistema):\n${JSON.stringify(a.snapshot, null, 2)}` });
  if (Array.isArray(a.history)) for (const t of a.history) messages.push({ role: t.role === "user" ? "user" : "assistant", content: String(t.text || "") });

  const json = await callOpenAI({ messages, tools: AGENT_TOOLS, tool_choice: "auto" });
  const msg = json.choices?.[0]?.message || {};
  const actions: any[] = [];
  if (Array.isArray(msg.tool_calls)) {
    for (const tc of msg.tool_calls) {
      try { actions.push({ name: tc.function.name, args: JSON.parse(tc.function.arguments || "{}") }); } catch (_e) { /* ignora */ }
    }
  }
  return { answer: msg.content || "", actions };
}

async function metaAdsInsights(m: any) {
  const token = Deno.env.get("META_USER_TOKEN");
  if (!token) throw new Error("META_USER_TOKEN nao configurada nos secrets");
  const acct = String(m.accountId || "").replace(/^act_/, "");
  if (!acct) throw new Error("accountId obrigatorio");
  const ver = "v21.0";
  const base = `https://graph.facebook.com/${ver}`;
  // periodo: since/until (YYYY-MM-DD) ou date_preset (ex: last_30d)
  let range = "";
  if (m.since && m.until) range = `&time_range=${encodeURIComponent(JSON.stringify({ since: m.since, until: m.until }))}`;
  else range = `&date_preset=${m.datePreset || "last_30d"}`;
  const fields = "spend,impressions,clicks,ctr,cpc,cpm,reach,frequency,actions,action_values,purchase_roas";

  async function fetchInsights(level: string, extra = "") {
    const url = `${base}/act_${acct}/insights?level=${level}&fields=${fields}${level !== "account" ? ",campaign_name" : ""}${range}${extra}&limit=200&access_token=${token}`;
    const r = await fetch(url);
    const j = await r.json();
    if (j.error) throw new Error(j.error.message);
    return j.data || [];
  }
  function pickAction(arr: any[], types: string[]) {
    if (!Array.isArray(arr)) return 0;
    return arr.filter((x) => types.includes(x.action_type)).reduce((s, x) => s + parseFloat(x.value || "0"), 0);
  }
  function shape(row: any) {
    const purchases = pickAction(row.actions, ["purchase", "omni_purchase", "offsite_conversion.fb_pixel_purchase"]);
    const revenue = pickAction(row.action_values, ["purchase", "omni_purchase", "offsite_conversion.fb_pixel_purchase"]);
    const roas = Array.isArray(row.purchase_roas) && row.purchase_roas.length ? parseFloat(row.purchase_roas[0].value || "0") : (parseFloat(row.spend || "0") ? revenue / parseFloat(row.spend) : 0);
    return {
      campaign: row.campaign_name || null,
      spend: parseFloat(row.spend || "0"), impressions: parseInt(row.impressions || "0"), clicks: parseInt(row.clicks || "0"),
      ctr: parseFloat(row.ctr || "0"), cpc: parseFloat(row.cpc || "0"), cpm: parseFloat(row.cpm || "0"),
      reach: parseInt(row.reach || "0"), frequency: parseFloat(row.frequency || "0"),
      purchases, revenue, roas,
      leads: pickAction(row.actions, ["lead", "offsite_conversion.fb_pixel_lead"]),
      addToCart: pickAction(row.actions, ["add_to_cart", "offsite_conversion.fb_pixel_add_to_cart"]),
      initiateCheckout: pickAction(row.actions, ["initiate_checkout", "offsite_conversion.fb_pixel_initiate_checkout"]),
    };
  }
  const accountRows = await fetchInsights("account");
  const total = accountRows.length ? shape(accountRows[0]) : shape({});
  let campaigns: any[] = [];
  if (m.byCampaign && !m.daily) {
    const rows = await fetchInsights("campaign");
    campaigns = rows.map(shape).sort((a: any, b: any) => b.spend - a.spend);
  } else if (m.byCampaign && m.daily) {
    // detalhe diario por campanha: agrega e guarda records por dia (para a tabela semanal)
    const rows = await fetchInsights("campaign", "&time_increment=1");
    const byCamp: Record<string, any> = {};
    for (const row of rows) {
      const name = row.campaign_name || "Meta Ads";
      const s = shape(row);
      if (!byCamp[name]) byCamp[name] = { campaign: name, spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0, records: [] };
      const c = byCamp[name];
      c.spend += s.spend; c.impressions += s.impressions; c.clicks += s.clicks; c.reach += s.reach;
      c.revenue += s.revenue; c.purchases += s.purchases; c.leads += s.leads; c.addToCart += s.addToCart; c.initiateCheckout += s.initiateCheckout;
      c.records.push({ date: row.date_start, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions });
    }
    campaigns = Object.values(byCamp).map((c: any) => {
      c.ctr = c.impressions ? (c.clicks / c.impressions) * 100 : 0;
      c.cpc = c.clicks ? c.spend / c.clicks : 0;
      c.cpm = c.impressions ? (c.spend / c.impressions) * 1000 : 0;
      c.roas = c.spend ? c.revenue / c.spend : 0;
      return c;
    }).sort((a: any, b: any) => b.spend - a.spend);
  }
  return { total, campaigns, period: m.since && m.until ? { since: m.since, until: m.until } : { datePreset: m.datePreset || "last_30d" } };
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

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  try {
    const body = await req.json();
    const spreadsheetId = body.spreadsheetId;
    const tabs = body.tabs;
    const orders = body.orders;
    const analysis = body.analysis;
    const agent = body.agent;

    if (body.metaAds) {
      const r = await metaAdsInsights(body.metaAds);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAccounts) {
      const r = await metaListAccounts();
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
