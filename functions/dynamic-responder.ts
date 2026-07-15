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

Voce tambem pode EXECUTAR acoes quando o gestor pedir explicitamente: criar/concluir tarefas E acoes reais no Meta Ads (pausar_meta, reativar_meta, ajustar_orcamento, duplicar_campanha). Para as acoes do Meta, use SEMPRE o 'id' e o 'nivel' que estao na lista 'metaEntidades' do snapshot (campanhas, conjuntos e anuncios com id, status e orcamento atuais) — nunca invente ids. O sistema mostra um card de confirmacao antes de executar; entao apenas PROPONHA a acao chamando a funcao e explique o porque em texto; nunca afirme que ja executou. So proponha acao no Meta quando o gestor pedir ou quando os dados claramente justificarem (ex: anuncio com gasto alto e 0 compras -> propor pausar). Seu valor principal continua sendo a analise tecnica.`;

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
    else if (level === "ad") lvlFields = ",campaign_name,campaign_id,adset_name,ad_name,ad_id";
    let url: string | null = `${base}/act_${acct}/insights?level=${level}&fields=${fields}${lvlFields}${range}${extra}&limit=200&access_token=${token}`;
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
    return map;
  }
  // miniaturas dos criativos por ad_id (uma chamada por conta)
  async function fetchThumbs(acct: string) {
    const map: Record<string, string> = {};
    let url: string | null = `${base}/act_${acct}/ads?fields=id,creative{thumbnail_url,image_url}&limit=200&access_token=${token}`;
    for (let i = 0; i < 20 && url; i++) {
      const r = await fetch(url);
      const j = await r.json();
      if (j.error) break;
      for (const ad of (j.data || [])) {
        const t = ad.creative?.thumbnail_url || ad.creative?.image_url;
        if (t) map[ad.id] = t;
      }
      url = j.paging?.next || null;
    }
    return map;
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
      campaign: row.campaign_name || null, campaignId: row.campaign_id || null,
      spend: parseFloat(row.spend || "0"), impressions: parseInt(row.impressions || "0"), clicks: parseInt(row.clicks || "0"),
      ctr: parseFloat(row.ctr || "0"), cpc: parseFloat(row.cpc || "0"), cpm: parseFloat(row.cpm || "0"),
      reach: parseInt(row.reach || "0"), frequency: parseFloat(row.frequency || "0"),
      purchases, revenue, roas,
      leads: pickAction(row.actions, ["lead", "offsite_conversion.fb_pixel_lead"]),
      addToCart: pickAction(row.actions, ["add_to_cart", "offsite_conversion.fb_pixel_add_to_cart"]),
      initiateCheckout: pickAction(row.actions, ["initiate_checkout", "offsite_conversion.fb_pixel_initiate_checkout"]),
    };
  }
  const totAgg: any = { spend: 0, impressions: 0, clicks: 0, reach: 0, revenue: 0, purchases: 0, leads: 0, addToCart: 0, initiateCheckout: 0 };
  const byCamp: Record<string, any> = {};
  const ads: any[] = [];
  const wantObj = m.byAd || m.byCampaign;
  // Contas em PARALELO, e dentro de cada conta as chamadas (conta/objetivos/anuncios/campanhas/thumbs) tambem em paralelo.
  const perAccount = await Promise.all(accounts.map(async (acc) => {
    const [accountRows, objByCampId, adRows, campRows, thumbs] = await Promise.all([
      fetchInsights(acc.id, "account"),
      wantObj ? fetchObjectives(acc.id) : Promise.resolve({} as Record<string, any>),
      m.byAd ? fetchInsights(acc.id, "ad") : Promise.resolve([] as any[]),
      m.byCampaign ? fetchInsights(acc.id, "campaign", m.daily ? "&time_increment=1" : "") : Promise.resolve([] as any[]),
      m.byAd ? fetchThumbs(acc.id) : Promise.resolve({} as Record<string, string>),
    ]);
    return { acc, accountRows, objByCampId, adRows, campRows, thumbs };
  }));
  for (const { acc, accountRows, objByCampId, adRows, campRows, thumbs } of perAccount) {
    const at = accountRows.length ? shape(accountRows[0]) : shape({});
    totAgg.spend += at.spend; totAgg.impressions += at.impressions; totAgg.clicks += at.clicks; totAgg.reach += at.reach;
    totAgg.revenue += at.revenue; totAgg.purchases += at.purchases; totAgg.leads += at.leads; totAgg.addToCart += at.addToCart; totAgg.initiateCheckout += at.initiateCheckout;
    for (const row of adRows) {
      const s = shape(row);
      ads.push({
        adId: row.ad_id, adName: row.ad_name || "(sem nome)", campaign: row.campaign_name || "", adset: row.adset_name || "",
        account: acc.name || acc.id, thumbnail: thumbs[row.ad_id] || null,
        objetivo: objByCampId[row.campaign_id] || metaObjetivo(""),
        spend: s.spend, impressions: s.impressions, clicks: s.clicks, reach: s.reach, frequency: s.frequency,
        ctr: s.ctr, cpc: s.cpc, cpm: s.cpm, purchases: s.purchases, revenue: s.revenue, roas: s.roas,
        leads: s.leads, addToCart: s.addToCart, initiateCheckout: s.initiateCheckout,
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
      if (m.daily) c.records.push({ date: row.date_start, spend: s.spend, sales: s.purchases, revenue: s.revenue, clicks: s.clicks, impressions: s.impressions });
    }
  }
  const total = {
    ...totAgg,
    ctr: totAgg.impressions ? (totAgg.clicks / totAgg.impressions) * 100 : 0,
    cpc: totAgg.clicks ? totAgg.spend / totAgg.clicks : 0,
    cpm: totAgg.impressions ? (totAgg.spend / totAgg.impressions) * 1000 : 0,
    roas: totAgg.spend ? totAgg.revenue / totAgg.spend : 0,
  };
  const campaigns = Object.values(byCamp).map((c: any) => {
    c.ctr = c.impressions ? (c.clicks / c.impressions) * 100 : 0;
    c.cpc = c.clicks ? c.spend / c.clicks : 0;
    c.cpm = c.impressions ? (c.spend / c.impressions) * 1000 : 0;
    c.roas = c.spend ? c.revenue / c.spend : 0;
    return c;
  }).sort((a: any, b: any) => b.spend - a.spend);
  ads.sort((a: any, b: any) => b.spend - a.spend);
  return { total, campaigns, ads, accounts, period: m.since && m.until ? { since: m.since, until: m.until } : { datePreset: m.datePreset || "last_30d" } };
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
  const campaigns: any[] = [], adsets: any[] = [], ads: any[] = [];
  for (const acc of accounts) {
    const cs = await pageAll(`act_${acc.id}/campaigns?fields=id,name,status,effective_status,daily_budget,lifetime_budget,objective`);
    const objById: Record<string, any> = {};
    for (const c of cs) { const ob = metaObjetivo(c.objective); objById[c.id] = ob; campaigns.push({ id: c.id, nome: c.name, status: c.status, entrega: c.effective_status, orcamentoDiario: c.daily_budget ? +c.daily_budget / 100 : null, objetivo: ob, conta: acc.name || acc.id }); }
    const as = await pageAll(`act_${acc.id}/adsets?fields=id,name,status,effective_status,daily_budget,campaign_id`);
    for (const s of as) adsets.push({ id: s.id, nome: s.name, status: s.status, entrega: s.effective_status, orcamentoDiario: s.daily_budget ? +s.daily_budget / 100 : null, campanhaId: s.campaign_id, conta: acc.name || acc.id });
    const ds = await pageAll(`act_${acc.id}/ads?fields=id,name,status,effective_status,campaign_id,adset_id`);
    for (const d of ds) ads.push({ id: d.id, nome: d.name, status: d.status, entrega: d.effective_status, campanhaId: d.campaign_id, conjuntoId: d.adset_id, objetivo: objById[d.campaign_id] || metaObjetivo(""), conta: acc.name || acc.id });
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
    if (body.metaEntities) {
      const r = await metaEntities(body.metaEntities);
      return new Response(JSON.stringify({ data: r }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    if (body.metaAction) {
      const r = await metaAction(body.metaAction);
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
