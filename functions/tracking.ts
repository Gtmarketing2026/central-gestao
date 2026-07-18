// Função pública de rastreamento (verify_jwt = false).
// Rotas (sob /functions/v1/tracking):
//   GET  /pixel/script/<token>.js  -> serve o pixel JS (captura origem, pageview, cliques de WhatsApp)
//   POST /collect                  -> ingere evento do pixel
//   GET  /l/<client_id>/<slug>     -> link rastreável: registra clique e redireciona (bio/site/WhatsApp)
// Escreve no banco com a service_role (ignora RLS) — visitantes são anônimos.

const SB_URL = Deno.env.get("SUPABASE_URL")!;
const SB_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "content-type",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
};

async function sbInsert(table: string, row: Record<string, unknown>) {
  await fetch(`${SB_URL}/rest/v1/${table}`, {
    method: "POST",
    headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" },
    body: JSON.stringify(row),
  });
}
async function sbSelect(table: string, query: string) {
  const r = await fetch(`${SB_URL}/rest/v1/${table}?${query}`, {
    headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}` },
  });
  return r.ok ? await r.json() : [];
}
async function sbPatch(table: string, query: string, row: Record<string, unknown>) {
  await fetch(`${SB_URL}/rest/v1/${table}?${query}`, {
    method: "PATCH",
    headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, "Content-Type": "application/json", Prefer: "return=minimal" },
    body: JSON.stringify(row),
  });
}
function uid() { return crypto.randomUUID().replace(/-/g, "").slice(0, 20); }

// Webhook do RD Station: RD chama a cada conversão. Guarda o payload cru + campos extraídos.
async function handleRdWebhook(url: URL, req: Request) {
  const clientId = url.searchParams.get("client") || "";
  let body: any = {};
  try { body = await req.json(); } catch (_e) { /* ignora */ }
  // O RD manda {leads:[{...}]} ou {event_type,...,leads:[...]}; guardamos cada lead como uma linha.
  const leads: any[] = Array.isArray(body?.leads) ? body.leads : (body?.leads ? [body.leads] : [body]);
  const rows = leads.map((L: any) => {
    const conv = L?.last_conversion || (Array.isArray(L?.conversions) ? L.conversions[L.conversions.length - 1] : null) || {};
    const src = conv?.content || conv || {};
    return {
      id: uid(), client_id: clientId,
      event_identifier: clip(conv?.conversion_identifier || conv?.identifier || L?.conversion_identifier, 200),
      email: clip(L?.email, 200), name: clip(L?.name, 200),
      source: clip(src?.utm_source || L?.traffic_source || conv?.source, 120),
      medium: clip(src?.utm_medium || L?.traffic_medium, 120),
      campaign: clip(src?.utm_campaign || L?.traffic_campaign, 200),
      content: clip(src?.utm_content, 200), term: clip(src?.utm_term, 200),
      converted_at: conv?.created_at || conv?.event_timestamp || null,
      payload: L || body,
    };
  });
  for (const r of rows) { try { await sbInsert("rd_conversions", r); } catch (_e) { /* segue */ } }
  return new Response("ok", { headers: { ...cors, "Content-Type": "text/plain" } });
}

// OAuth do RD Station: recebe o ?code, troca por refresh_token e guarda em account_config.data.rd_station
async function handleRdCallback(url: URL) {
  // Supabase força text/plain nas Edge Functions — então retorno TEXTO PURO, sem tags e sem acento (evita mojibake).
  const strip = (s: string) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "");
  const page = (title: string, msg: string, ok: boolean) => new Response(
    (ok ? "[OK] " : "[!] ") + strip(title) + "\n\n" + strip(msg) + "\n\nPode fechar esta aba e voltar ao sistema.",
    { status: 200, headers: { "content-type": "text/plain; charset=utf-8" } });
  const code = url.searchParams.get("code");
  const clientId = url.searchParams.get("state"); // qual cliente está conectando o RD dele
  if (!code) return page("Autorização não concluída", "Não recebi o código do RD Station. Tente conectar de novo.", false);
  if (!clientId) return page("Cliente não identificado", "Faltou identificar o cliente. Refaça a conexão pelo cadastro do cliente.", false);
  // client_id/secret do APP (nível conta)
  const rows = await sbSelect("account_config", `id=eq.main&select=data`);
  const rd = (rows[0]?.data || {}).rd_station || {};
  if (!rd.client_id || !rd.client_secret) return page("Faltam credenciais do App", "Salve o Client ID e o Client Secret do App do RD na aba Configurações antes de conectar.", false);
  const redirect = `${url.origin}/functions/v1/tracking/rd/callback`;
  try {
    const r = await fetch("https://api.rd.services/auth/token", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ client_id: rd.client_id, client_secret: rd.client_secret, code, redirect_uri: redirect }),
    });
    const j = await r.json();
    if (!r.ok || !j.refresh_token) return page("Falha ao conectar", "O RD não devolveu o token. Motivo: " + (j.error_description || j.error || `HTTP ${r.status}`), false);
    // guarda o refresh_token NO CLIENTE (cada cliente tem o RD dele)
    const cRows = await sbSelect("clients", `id=eq.${encodeURIComponent(clientId)}&select=rd_config,name`);
    if (!cRows[0]) return page("Cliente não encontrado", "O cliente informado não existe mais.", false);
    const cfg = cRows[0].rd_config || {};
    const newCfg = { ...cfg, refresh_token: j.refresh_token, connected_at: new Date().toISOString() };
    await sbPatch("clients", `id=eq.${encodeURIComponent(clientId)}`, { rd_config: newCfg });
    return page("RD Station conectado!", `A conta RD de "${cRows[0].name || "cliente"}" foi conectada. Já podemos puxar os eventos e conversões dela.`, true);
  } catch (e) {
    return page("Erro ao conectar", String((e as any)?.message || e), false);
  }
}
function clip(s: unknown, n = 400) { const v = s == null ? null : String(s); return v ? v.slice(0, n) : null; }

// ---- pixel script (embute client_id + base) ----
function pixelScript(cid: string, base: string) {
  return `(function(){
"use strict";
var CID=${JSON.stringify(cid)},BASE=${JSON.stringify(base)};
function q(n){try{return new URLSearchParams(location.search).get(n)||''}catch(e){return ''}}
function ck(n,v,d){if(v===undefined){var m=document.cookie.match('(^|;)\\\\s*'+n+'\\\\s*=\\\\s*([^;]+)');return m?m.pop():''}var e=new Date();e.setTime(e.getTime()+(d||365)*864e5);document.cookie=n+'='+v+';path=/;expires='+e.toUTCString()+';SameSite=Lax'}
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,10)}
var anon=ck('_alc_a');if(!anon){anon=uid();ck('_alc_a',anon,365)}
var sess=sessionStorage.getItem('_alc_s');if(!sess){sess=uid();try{sessionStorage.setItem('_alc_s',sess)}catch(e){}}
var o={utm_source:q('utm_source'),utm_medium:q('utm_medium'),utm_campaign:q('utm_campaign'),utm_content:q('utm_content'),utm_term:q('utm_term'),fbclid:q('fbclid'),gclid:q('gclid')||q('gbraid')||q('wbraid')};
var has=Object.keys(o).some(function(k){return o[k]});
try{var st=JSON.parse(localStorage.getItem('_alc_o')||'null');if(has){localStorage.setItem('_alc_o',JSON.stringify(o))}else if(st){o=st}}catch(e){}
function send(t,x){var b={cid:CID,type:t,anon:anon,sess:sess,ref:document.referrer||'',landing:location.pathname+location.search,ua:navigator.userAgent};for(var k in o)b[k]=o[k];if(x)for(var j in x)b[j]=x[j];var s=JSON.stringify(b);try{if(navigator.sendBeacon){navigator.sendBeacon(BASE+'/collect',s);return}}catch(e){}try{fetch(BASE+'/collect',{method:'POST',body:s,keepalive:true,headers:{'Content-Type':'application/json'}}).catch(function(){})}catch(e){}}
send('pageview');
document.addEventListener('click',function(e){var a=e.target&&e.target.closest?e.target.closest('a'):null;if(!a||!a.href)return;if(/wa\\.me|api\\.whatsapp\\.com|whatsapp:/i.test(a.href)){send('wpp_click',{dest:a.href.slice(0,300)})}},true);
window.ALICIA={send:send,origin:o,anon:anon};
})();`;
}

async function handleCollect(req: Request) {
  let b: any = {};
  try { b = await req.json(); } catch (_e) { return new Response("bad", { status: 400, headers: cors }); }
  const cid = clip(b.cid, 40);
  if (!cid || !b.type) return new Response("bad", { status: 400, headers: cors });
  await sbInsert("track_events", {
    id: uid(), client_id: cid, type: clip(b.type, 20), session_id: clip(b.sess, 40), anon_id: clip(b.anon, 40),
    utm_source: clip(b.utm_source, 120), utm_medium: clip(b.utm_medium, 120), utm_campaign: clip(b.utm_campaign, 200),
    utm_content: clip(b.utm_content, 200), utm_term: clip(b.utm_term, 200),
    fbclid: clip(b.fbclid, 300), gclid: clip(b.gclid, 300),
    referrer: clip(b.ref, 300), landing: clip(b.landing, 300), user_agent: clip(b.ua, 300),
    meta: b.dest ? { dest: clip(b.dest, 300) } : null,
  });
  return new Response("ok", { headers: { ...cors, "Content-Type": "text/plain" } });
}

async function handleRedirect(cid: string, slug: string, url: URL, ref: string | null) {
  const rows = await sbSelect("track_links", `client_id=eq.${encodeURIComponent(cid)}&slug=eq.${encodeURIComponent(slug)}&select=destination,kind&limit=1`);
  const link = rows[0];
  if (!link) return new Response("Link não encontrado", { status: 404, headers: cors });
  const g = (n: string) => url.searchParams.get(n) || "";
  // registra o clique com a origem que veio na própria URL do link
  await sbInsert("track_events", {
    id: uid(), client_id: cid, type: "link_click", link_slug: slug,
    utm_source: clip(g("utm_source"), 120), utm_medium: clip(g("utm_medium"), 120), utm_campaign: clip(g("utm_campaign"), 200),
    utm_content: clip(g("utm_content"), 200), utm_term: clip(g("utm_term"), 200),
    fbclid: clip(g("fbclid"), 300), gclid: clip(g("gclid") || g("gbraid") || g("wbraid"), 300),
    referrer: clip(ref, 300), meta: { kind: link.kind },
  });
  // destino: repassa a origem adiante (o pixel do destino também captura)
  let dest = String(link.destination || "");
  const passthrough = url.search.replace(/^\?/, "");
  if (passthrough) dest += (dest.includes("?") ? "&" : "?") + passthrough;
  return new Response(null, { status: 302, headers: { ...cors, Location: dest } });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  const url = new URL(req.url);
  // path após o slug da função
  const p = url.pathname.replace(/^\/functions\/v1\/tracking/, "").replace(/^\/tracking/, "") || "/";
  const ref = req.headers.get("referer");

  // GET /pixel/script/<token>.js
  const mScript = p.match(/^\/pixel\/script\/([^/]+)\.js$/);
  if (mScript) {
    const token = mScript[1];
    const rows = await sbSelect("tracking_config", `token=eq.${encodeURIComponent(token)}&select=client_id&limit=1`);
    const cid = rows[0]?.client_id;
    const base = `${url.origin}/functions/v1/tracking`;
    const js = cid ? pixelScript(cid, base) : `console.warn("[Rastreamento] token inválido");`;
    return new Response(js, { headers: { ...cors, "Content-Type": "application/javascript; charset=utf-8", "Cache-Control": "public, max-age=300" } });
  }

  if (p === "/collect" && req.method === "POST") return handleCollect(req);

  if (p === "/rd/callback") return handleRdCallback(url);

  if (p === "/rd/webhook" && req.method === "POST") return handleRdWebhook(url, req);
  if (p === "/rd/webhook") return new Response("rd webhook ok", { headers: { ...cors, "Content-Type": "text/plain" } });

  // GET /l/<client_id>/<slug>
  const mLink = p.match(/^\/l\/([^/]+)\/([^/]+)$/);
  if (mLink) return handleRedirect(decodeURIComponent(mLink[1]), decodeURIComponent(mLink[2]), url, ref);

  return new Response("tracking ok", { headers: { ...cors, "Content-Type": "text/plain" } });
});
