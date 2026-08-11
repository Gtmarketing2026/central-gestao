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

// OAuth da Nuvemshop: recebe o ?code, troca por access_token + store_id e guarda no cliente (state = clientId).
async function handleNuvemshopCallback(url: URL) {
  const strip = (s: string) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "");
  const page = (t: string, m: string, ok: boolean) => new Response((ok ? "[OK] " : "[!] ") + strip(t) + "\n\n" + strip(m) + "\n\nPode fechar esta aba e voltar ao sistema.", { status: 200, headers: { "content-type": "text/plain; charset=utf-8" } });
  const code = url.searchParams.get("code");
  const clientId = url.searchParams.get("state");
  if (!code) return page("Autorizacao nao concluida", "Nao recebi o codigo da Nuvemshop.", false);
  if (!clientId) return page("Cliente nao identificado", "Refaca a conexao pelo cadastro do cliente.", false);
  const acc = await sbSelect("account_config", `id=eq.main&select=data`);
  const ns = (acc[0]?.data || {}).nuvemshop || {};
  if (!ns.client_id || !ns.client_secret) return page("Faltam credenciais do App", "Configure o App da Nuvemshop em Configuracoes.", false);
  try {
    const r = await fetch("https://www.tiendanube.com/apps/authorize/token", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ client_id: ns.client_id, client_secret: ns.client_secret, grant_type: "authorization_code", code }),
    });
    const j = await r.json();
    if (!r.ok || !j.access_token) return page("Falha ao conectar", "A Nuvemshop nao devolveu o token: " + (j.error_description || j.error || `HTTP ${r.status}`), false);
    const cRows = await sbSelect("clients", `id=eq.${encodeURIComponent(clientId)}&select=nuvemshop_config,name`);
    if (!cRows[0]) return page("Cliente nao encontrado", "O cliente informado nao existe mais.", false);
    const cfg = { ...(cRows[0].nuvemshop_config || {}), access_token: j.access_token, store_id: j.user_id, connected_at: new Date().toISOString() };
    await sbPatch("clients", `id=eq.${encodeURIComponent(clientId)}`, { nuvemshop_config: cfg });
    return page("Nuvemshop conectada!", `A loja de "${cRows[0].name || "cliente"}" foi conectada. Ja podemos puxar os pedidos.`, true);
  } catch (e) {
    return page("Erro ao conectar", String((e as any)?.message || e), false);
  }
}

// Webhook do RD Station: RD chama a cada conversão. Guarda o payload cru + campos extraídos.
// Normaliza um evento do RD em linha da tabela. Cobre os DOIS formatos:
//  a) NOVO (webhook de contato):   {contact:{name,email,mobile_phone,personal_phone,uuid,funnel:{origin}}, event_identifier, event_timestamp}
//  b) ANTIGO (leads[]):            {leads:[{email,name,last_conversion:{conversion_identifier,content:{utm_*}}}]}
export function rdRow(clientId: string, L: any, body?: any) {
  const B = body || L || {};
  const ct = L?.contact || B?.contact || null;
  const conv = L?.last_conversion || (Array.isArray(L?.conversions) ? L.conversions[L.conversions.length - 1] : null) || {};
  const src = conv?.content || conv || {};
  const digits = (v: any) => String(v || "").replace(/[^0-9]/g, "");
  const phoneRaw = ct?.mobile_phone || ct?.personal_phone || L?.mobile_phone || L?.personal_phone || L?.phone || "";
  const ph = digits(phoneRaw);
  return {
    id: uid(), client_id: clientId,
    event_identifier: clip(B?.event_identifier || conv?.conversion_identifier || conv?.identifier || L?.conversion_identifier, 200),
    email: clip((ct?.email || L?.email || "").toString().toLowerCase() || null, 200),
    name: clip(ct?.name || L?.name, 200),
    phone: ph ? clip(ph, 20) : null,
    source: clip(src?.utm_source || L?.traffic_source || conv?.source || (ct?.funnel?.origin && ct.funnel.origin !== "Desconhecido" ? ct.funnel.origin : null), 120),
    medium: clip(src?.utm_medium || L?.traffic_medium, 120),
    campaign: clip(src?.utm_campaign || L?.traffic_campaign, 200),
    content: clip(src?.utm_content, 200), term: clip(src?.utm_term, 200),
    converted_at: B?.event_timestamp || B?.timestamp || conv?.created_at || conv?.event_timestamp || null,
    payload: L || B,
  };
}
async function handleRdWebhook(url: URL, req: Request) {
  const clientId = url.searchParams.get("client") || "";
  let body: any = {};
  try { body = await req.json(); } catch (_e) { /* ignora */ }
  const leads: any[] = Array.isArray(body?.leads) ? body.leads : (body?.leads ? [body.leads] : [body]);
  const rows = leads.map((L: any) => rdRow(clientId, L, body));
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

// OAuth do TikTok Ads: recebe ?auth_code (TikTok não usa "code") + state=clientId, troca por access_token e guarda no cliente.
// TikTok não expira o access_token (fica válido até o usuário revogar) — não tem refresh_token nesse fluxo.
async function handleTikTokCallback(url: URL) {
  const strip = (s: string) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "");
  const page = (t: string, m: string, ok: boolean) => new Response((ok ? "[OK] " : "[!] ") + strip(t) + "\n\n" + strip(m) + "\n\nPode fechar esta aba e voltar ao sistema.", { status: 200, headers: { "content-type": "text/plain; charset=utf-8" } });
  const authCode = url.searchParams.get("auth_code") || url.searchParams.get("code");
  const clientId = url.searchParams.get("state");
  if (!authCode) return page("Autorizacao nao concluida", "Nao recebi o auth_code do TikTok.", false);
  if (!clientId) return page("Cliente nao identificado", "Refaca a conexao pelo cadastro do cliente.", false);
  const acc = await sbSelect("account_config", `id=eq.main&select=data`);
  const tk = (acc[0]?.data || {}).tiktok_ads || {};
  if (!tk.app_id || !tk.secret) return page("Faltam credenciais do App", "Configure o App ID e o Secret do TikTok em Configuracoes.", false);
  try {
    const r = await fetch("https://business-api.tiktok.com/open_api/v1.3/oauth2/access_token/", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ app_id: tk.app_id, secret: tk.secret, auth_code: authCode }),
    });
    const j = await r.json();
    const d = j?.data;
    if (!r.ok || !d?.access_token) return page("Falha ao conectar", "O TikTok nao devolveu o token: " + (j?.message || `HTTP ${r.status}`), false);
    const cRows = await sbSelect("clients", `id=eq.${encodeURIComponent(clientId)}&select=tiktok_config,name`);
    if (!cRows[0]) return page("Cliente nao encontrado", "O cliente informado nao existe mais.", false);
    const cfg = { ...(cRows[0].tiktok_config || {}), access_token: d.access_token, advertiser_ids: d.advertiser_ids || [], scope: d.scope || [], connected_at: new Date().toISOString() };
    await sbPatch("clients", `id=eq.${encodeURIComponent(clientId)}`, { tiktok_config: cfg });
    return page("TikTok Ads conectado!", `A conta TikTok de "${cRows[0].name || "cliente"}" foi conectada (${(d.advertiser_ids || []).length} conta(s) de anuncio encontrada(s)). Volte ao cadastro do cliente pra escolher qual conta usar.`, true);
  } catch (e) {
    return page("Erro ao conectar", String((e as any)?.message || e), false);
  }
}

// OAuth do Pinterest Ads: fluxo padrao com refresh_token (expira, precisa renovar periodicamente).
async function handlePinterestCallback(url: URL) {
  const strip = (s: string) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "");
  const page = (t: string, m: string, ok: boolean) => new Response((ok ? "[OK] " : "[!] ") + strip(t) + "\n\n" + strip(m) + "\n\nPode fechar esta aba e voltar ao sistema.", { status: 200, headers: { "content-type": "text/plain; charset=utf-8" } });
  const code = url.searchParams.get("code");
  const clientId = url.searchParams.get("state");
  const err = url.searchParams.get("error");
  if (err) return page("Conexao cancelada", "Voce recusou o acesso (" + err + ").", false);
  if (!code) return page("Autorizacao nao concluida", "Nao recebi o codigo do Pinterest.", false);
  if (!clientId) return page("Cliente nao identificado", "Refaca a conexao pelo cadastro do cliente.", false);
  const acc = await sbSelect("account_config", `id=eq.main&select=data`);
  const pt = (acc[0]?.data || {}).pinterest_ads || {};
  if (!pt.client_id || !pt.client_secret) return page("Faltam credenciais do App", "Configure o Client ID e o Secret do Pinterest em Configuracoes.", false);
  const redirect = `${url.origin}/functions/v1/tracking/pinterest/callback`;
  try {
    const basic = btoa(`${pt.client_id}:${pt.client_secret}`);
    const r = await fetch("https://api.pinterest.com/v5/oauth/token", {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", Authorization: `Basic ${basic}` },
      body: new URLSearchParams({ grant_type: "authorization_code", code, redirect_uri: redirect }),
    });
    const j = await r.json();
    if (!r.ok || !j.access_token) return page("Falha ao conectar", "O Pinterest nao devolveu o token: " + (j.message || j.error_description || `HTTP ${r.status}`), false);
    const cRows = await sbSelect("clients", `id=eq.${encodeURIComponent(clientId)}&select=pinterest_config,name`);
    if (!cRows[0]) return page("Cliente nao encontrado", "O cliente informado nao existe mais.", false);
    const cfg = { ...(cRows[0].pinterest_config || {}), access_token: j.access_token, refresh_token: j.refresh_token || null, connected_at: new Date().toISOString() };
    await sbPatch("clients", `id=eq.${encodeURIComponent(clientId)}`, { pinterest_config: cfg });
    return page("Pinterest Ads conectado!", `A conta Pinterest de "${cRows[0].name || "cliente"}" foi conectada. Volte ao cadastro do cliente pra escolher qual conta de anuncio usar.`, true);
  } catch (e) {
    return page("Erro ao conectar", String((e as any)?.message || e), false);
  }
}

// ---- pixel script (embute client_id + base + LGPD) ----
// showBanner=true (padrao, cliente sem banner de cookies proprio no site): mostra um banner leve ANTES de
// gravar cookie/mandar qualquer evento, guarda a decisao em localStorage (_alc_consent) e so rastreia se
// aceitar. showBanner=false (cliente marcou "ja tenho aviso de cookies proprio" em Configuracoes): roda igual
// a antes, sem banner nenhum - confia no aviso generico que o site ja tem (padrao normal pra script de terceiro).
function pixelScript(cid: string, base: string, showBanner: boolean) {
  return `(function(){
"use strict";
var CID=${JSON.stringify(cid)},_FB=${JSON.stringify(base)},SHOW_BANNER=${showBanner ? "true" : "false"};
var BASE=(function(){try{var s=document.currentScript;if(!s){var a=document.getElementsByTagName('script');s=a[a.length-1]}var u=(s&&s.src)||'';var m=u.replace(/\\/pixel\\/script\\/[^/]*$/,'');return (m&&m!==u)?m:_FB}catch(e){return _FB}})();
function q(n){try{return new URLSearchParams(location.search).get(n)||''}catch(e){return ''}}
function ck(n,v,d){if(v===undefined){var m=document.cookie.match('(^|;)\\\\s*'+n+'\\\\s*=\\\\s*([^;]+)');return m?m.pop():''}var e=new Date();e.setTime(e.getTime()+(d||365)*864e5);document.cookie=n+'='+v+';path=/;expires='+e.toUTCString()+';SameSite=Lax'}
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,10)}
window.ALICIA={send:function(){},origin:{},anon:''};
function startTracking(){
var anon=ck('_alc_a');if(!anon){anon=uid();ck('_alc_a',anon,365)}
var sess=sessionStorage.getItem('_alc_s');if(!sess){sess=uid();try{sessionStorage.setItem('_alc_s',sess)}catch(e){}}
var o={utm_source:q('utm_source'),utm_medium:q('utm_medium'),utm_campaign:q('utm_campaign'),utm_content:q('utm_content'),utm_term:q('utm_term'),fbclid:q('fbclid'),gclid:q('gclid')||q('gbraid')||q('wbraid'),campaignid:q('campaignid')||q('campaign_id'),adgroupid:q('adgroupid')||q('adset_id'),adid:q('adid')||q('ad_id'),keyword:q('keyword')||q('utm_term'),matchtype:q('matchtype'),placement:q('placement')||q('network')};
var has=Object.keys(o).some(function(k){return o[k]});
try{var st=JSON.parse(localStorage.getItem('_alc_o')||'null');if(has){localStorage.setItem('_alc_o',JSON.stringify(o))}else if(st){o=st}}catch(e){}
function send(t,x){var b={cid:CID,type:t,anon:anon,sess:sess,ref:document.referrer||'',landing:location.pathname+location.search,ua:navigator.userAgent};for(var k in o)b[k]=o[k];if(x)for(var j in x)b[j]=x[j];var s=JSON.stringify(b);try{if(navigator.sendBeacon){navigator.sendBeacon(BASE+'/collect',s);return}}catch(e){}try{fetch(BASE+'/collect',{method:'POST',body:s,keepalive:true,headers:{'Content-Type':'application/json'}}).catch(function(){})}catch(e){}}
send('pageview');
document.addEventListener('click',function(e){var a=e.target&&e.target.closest?e.target.closest('a'):null;if(!a||!a.href)return;if(/wa\\.me|api\\.whatsapp\\.com|whatsapp:/i.test(a.href)){send('wpp_click',{dest:a.href.slice(0,300)});var hasO=o.utm_campaign||o.gclid||o.fbclid||o.utm_source;if(hasO&&a.href.indexOf('[#')===-1){var rid=uid().slice(0,8);try{var u=new URL(a.href);var t=u.searchParams.get('text')||'';u.searchParams.set('text',(t?t+' ':'')+'[#'+rid+']');a.href=u.toString()}catch(_e){a.href=a.href+(a.href.indexOf('?')>-1?'&':'?')+'text='+encodeURIComponent('[#'+rid+']')}var pl=JSON.stringify({ref:rid,cid:CID,utm_source:o.utm_source,utm_medium:o.utm_medium,utm_campaign:o.utm_campaign,utm_content:o.utm_content,utm_term:o.utm_term,gclid:o.gclid,fbclid:o.fbclid});try{if(navigator.sendBeacon){navigator.sendBeacon(BASE+'/wa/ref',pl)}else{fetch(BASE+'/wa/ref',{method:'POST',headers:{'Content-Type':'application/json'},body:pl,keepalive:true}).catch(function(){})}}catch(_e2){}}}},true);
document.addEventListener('submit',function(e){var f=e.target;if(!f||f.tagName!=='FORM')return;var idv=(f.id||'')+' '+(f.getAttribute('name')||'')+' '+(f.className||'')+' '+(f.action||'');if(/search|busca|pesquisa/i.test(idv))return;send('form',{action:(f.action||'').slice(0,300),id:(f.id||'').slice(0,100)})},true);
window.ALICIA={send:send,origin:o,anon:anon};
}
function showConsentBanner(){
var d=document.createElement('div');
d.setAttribute('style','position:fixed;left:0;right:0;bottom:0;z-index:2147483647;background:#1a1a1a;color:#f2f2f2;font:13px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;padding:14px 16px;display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between;box-shadow:0 -2px 12px rgba(0,0,0,.25)');
var t=document.createElement('span');t.style.flex='1 1 260px';t.textContent='Este site usa cookies para entender de onde vêm os visitantes e melhorar sua experiência.';
var box=document.createElement('span');box.style.cssText='display:flex;gap:8px;flex-shrink:0';
function btn(label,primary){var b=document.createElement('button');b.type='button';b.textContent=label;b.setAttribute('style','cursor:pointer;border-radius:6px;padding:8px 14px;font-size:13px;border:1px solid '+(primary?'#4ade80':'#666')+';background:'+(primary?'#4ade80':'transparent')+';color:'+(primary?'#111':'#f2f2f2'));return b}
var accept=btn('Aceitar',true),decline=btn('Recusar',false);
accept.onclick=function(){try{localStorage.setItem('_alc_consent','accepted')}catch(e){}d.remove();startTracking()};
decline.onclick=function(){try{localStorage.setItem('_alc_consent','declined')}catch(e){}d.remove()};
box.appendChild(decline);box.appendChild(accept);d.appendChild(t);d.appendChild(box);
function mount(){document.body.appendChild(d)}
if(document.body)mount();else document.addEventListener('DOMContentLoaded',mount)
}
if(!SHOW_BANNER){startTracking()}
else{var consent=null;try{consent=localStorage.getItem('_alc_consent')}catch(e){}if(consent==='accepted'){startTracking()}else if(consent!=='declined'){showConsentBanner()}}
})();`;
}

// Pixel: clique no WhatsApp → grava a origem (utm/gclid) e devolve um ref pra injetar no texto.
async function handleWaRef(req: Request) {
  let b: any = {}; try { b = await req.json(); } catch (_e) { /* */ }
  const rid = (clip(b.ref, 10) || "").replace(/[^a-z0-9]/gi, "") || uid().slice(0, 8);
  const chan = _channelOf(b.utm_source || "", b.gclid || "", b.fbclid || "");
  await sbInsert("wa_ref_origins", { ref: rid, client_id: (await _resolveCid(b.cid || "")) || null, created_at: new Date().toISOString(), origin: { channel: chan, track_source: b.utm_source || chan, medium: b.utm_medium || "", campaign: b.utm_campaign || "", adgroup: b.utm_content || "", keyword: b.utm_term || "", gclid: b.gclid || "", fbclid: b.fbclid || "" } });
  return new Response(JSON.stringify({ ref: rid }), { headers: { ...cors, "Content-Type": "application/json" } });
}
// resolve o cid do pixel: aceita tanto o client_id quanto o TOKEN (snippet antigo/instalação manual)
async function _resolveCid(raw: string): Promise<string> {
  const v = clip(raw, 60);
  if (!v || !v.startsWith("tk_")) return v;
  const rows = await sbSelect("tracking_config", `token=eq.${encodeURIComponent(v)}&select=client_id&limit=1`);
  return rows[0]?.client_id || v;
}
async function handleCollect(req: Request) {
  let b: any = {};
  try { b = await req.json(); } catch (_e) { return new Response("bad", { status: 400, headers: cors }); }
  if (!b.cid || !b.type) return new Response("bad", { status: 400, headers: cors });
  const cid = await _resolveCid(b.cid);
  if (!cid) return new Response("bad", { status: 400, headers: cors });
  await sbInsert("track_events", {
    id: uid(), client_id: cid, type: clip(b.type, 20), session_id: clip(b.sess, 40), anon_id: clip(b.anon, 40),
    utm_source: clip(b.utm_source, 120), utm_medium: clip(b.utm_medium, 120), utm_campaign: clip(b.utm_campaign, 200),
    utm_content: clip(b.utm_content, 200), utm_term: clip(b.utm_term, 200),
    fbclid: clip(b.fbclid, 300), gclid: clip(b.gclid, 300),
    referrer: clip(b.ref, 300), landing: clip(b.landing, 300), user_agent: clip(b.ua, 300),
    meta: (b.dest || b.campaignid || b.adgroupid || b.adid || b.keyword || b.matchtype || b.placement) ? {
      dest: b.dest ? clip(b.dest, 300) : undefined,
      campaignid: clip(b.campaignid, 40), adgroupid: clip(b.adgroupid, 40), adid: clip(b.adid, 40),
      keyword: clip(b.keyword, 200), matchtype: clip(b.matchtype, 20), placement: clip(b.placement, 60),
    } : null,
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
  // Se o destino é WhatsApp e veio origem (utm/gclid/fbclid), grava um REF e injeta [#ref] no texto pré-preenchido,
  // pra casar o clique rastreado (campanha › grupo › palavra-chave) com a conversa que chega.
  const isWa = /wa\.me|api\.whatsapp\.com|whatsapp:/i.test(dest);
  const hasOrigin = !!(g("utm_campaign") || g("utm_source") || g("gclid") || g("gbraid") || g("wbraid") || g("fbclid"));
  if (isWa && hasOrigin) {
    const rid = uid().slice(0, 8);
    const chan = _channelOf(g("utm_source"), g("gclid") || g("gbraid") || g("wbraid"), g("fbclid"));
    await sbInsert("wa_ref_origins", { ref: rid, client_id: cid, created_at: new Date().toISOString(), origin: { channel: chan, track_source: g("utm_source") || chan, medium: g("utm_medium"), campaign: g("utm_campaign"), adgroup: g("utm_content"), keyword: g("utm_term"), gclid: g("gclid") || g("gbraid") || g("wbraid"), fbclid: g("fbclid"), link_slug: slug, kind: link.kind } });
    try { const u = new URL(dest); const t = u.searchParams.get("text") || ""; u.searchParams.set("text", (t ? t + " " : "") + "[#" + rid + "]"); dest = u.toString(); }
    catch (_e) { dest += (dest.includes("?") ? "&" : "?") + "text=" + encodeURIComponent("[#" + rid + "]"); }
  }
  return new Response(null, { status: 302, headers: { ...cors, Location: dest } });
}
function _channelOf(src: string, gclid: string, fbclid: string) { src = String(src || "").toLowerCase(); if (gclid || /google|gads|gclid|adwords/.test(src)) return "google"; if (fbclid || /meta|facebook|^fb$|instagram|^ig$/.test(src)) return "meta"; return src || "link"; }

// Extrai a atribuição de anúncio (Click-to-WhatsApp) de uma mensagem do uazapi
function waExtractOrigin(m: any): { type: string; data: Record<string, unknown> } | null {
  const c = m.content || {};
  const ci = c.contextInfo || c.extendedTextMessage?.contextInfo || m.contextInfo || {};
  const ad = ci.externalAdReply || c.externalAdReply || null;
  if (ad && (ad.sourceId || ad.sourceUrl || ad.ctwaClid || ad.title)) {
    return { type: "anuncio", data: {
      source_id: ad.sourceId || "", source_type: ad.sourceType || "", source_url: ad.sourceUrl || "",
      ctwa_clid: ad.ctwaClid || ci.ctwaClid || ci.ctwaPayload || "", title: ad.title || "", body: ad.body || "",
      thumbnail: ad.thumbnailUrl || ad.thumbnail || "", media_type: ad.mediaType || "",
      platform: ad.entryPointConversionApp || ci.entryPointConversionApp || "",
    } };
  }
  // CTWA sem externalAdReply: formato novo/opaco traz só sinais no contextInfo (ctwaPayload, conversionSource FB_Ads, entryPointConversionSource ctwa_ad).
  // É anúncio pago do Meta (mesmo sem o sourceId do anúncio) — não deixar cair como orgânico.
  const epSrc = String(ci.entryPointConversionSource || "").toLowerCase();
  const convSrc = String(ci.conversionSource || ci.entryPointConversionExternalSource || "");
  if (ci.ctwaPayload || ci.ctwaSignals || /ctwa|ad/.test(epSrc) || /fb_ads|ig_ads|instagram_ads|facebook_ads/i.test(convSrc)) {
    const app = String(ci.entryPointConversionApp || "").toLowerCase();
    return { type: "anuncio", data: {
      source_id: "", source_type: "ctwa_ad", source_url: "",
      ctwa_clid: ci.ctwaPayload || ci.ctwaClid || "", title: "", body: "",
      thumbnail: "", media_type: "",
      platform: app.includes("insta") ? "instagram" : app.includes("face") ? "facebook" : (app || (/fb_ads/i.test(convSrc) ? "facebook" : "")),
      conversion_source: convSrc || "", ctwa_only: true, // sem sourceId: sabemos que é Meta pago, mas não a campanha exata
    } };
  }
  // uazapi às vezes já traz origem em track_source/track_id (do rastreamento). NÃO usar m.source: é a plataforma do device (ios/web/android), não anúncio.
  if (m.track_source || m.track_id) {
    return { type: (m.track_source === "ad" ? "anuncio" : "utm"), data: { track_source: m.track_source || "", track_id: m.track_id || "" } };
  }
  return null;
}
// Casa o [#ref] injetado pelo link rastreável → origem completa (campanha › grupo › palavra-chave / gclid).
async function refOrigin(text: string): Promise<{ type: string; data: Record<string, unknown> } | null> {
  const mm = String(text || "").match(/\[#([a-z0-9]{6,10})\]/i); if (!mm) return null;
  const row = (await sbSelect("wa_ref_origins", `ref=eq.${encodeURIComponent(mm[1])}&select=origin&limit=1`))[0];
  const o = row && row.origin; if (!o) return null;
  const paid = o.channel === "google" || o.channel === "meta" || /cpc|paid|ad/i.test(String(o.medium || ""));
  return { type: paid ? "anuncio" : "utm", data: { channel: o.channel || "", track_source: o.track_source || o.channel || "", campaign: o.campaign || "", adset: o.adgroup || "", ad: o.keyword || "", keyword: o.keyword || "", adgroup: o.adgroup || "", gclid: o.gclid || "", fbclid: o.fbclid || "", medium: o.medium || "" } };
}
function _waMatchText(s: string) {
  return String(s || "").replace(/\[#(?:[a-z0-9]{6,10})\]/gi, "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
// Alguns fluxos do WhatsApp removem o [#ref] do texto pre-preenchido antes de entregar a mensagem.
// Fallback de origem por TELEFONE JA CONHECIDO: quando uma conversa nova chega sem CTWA nativo, sem [#ref] e
// sem match de texto+tempo, olha se esse telefone ja apareceu em outro toque com canal/campanha real (RD
// Station, conversa anterior, pedido da planilha) via o grafo de identidade da Jornada (lead_identities/
// lead_touchpoints) - reusa essa origem em vez de marcar como organico. matched_by:'phone_history' deixa claro
// que e INFERIDO, nao um sinal direto desta conversa. Espelha waPhoneHistoryOrigin do dynamic-responder.ts.
async function waPhoneHistoryOrigin(clientId: string | null, phone: string): Promise<{ type: string; data: Record<string, unknown> } | null> {
  if (!clientId) return null;
  const digits = String(phone || "").replace(/[^0-9]/g, "");
  const pk = digits.length >= 8 ? digits.slice(-8) : "";
  if (!pk) return null;
  const ident = (await sbSelect("lead_identities", `client_id=eq.${encodeURIComponent(clientId)}&kind=eq.phone&value=eq.${encodeURIComponent(pk)}&select=person_id&limit=1`))[0];
  if (!ident) return null;
  const toques = await sbSelect("lead_touchpoints", `client_id=eq.${encodeURIComponent(clientId)}&person_id=eq.${encodeURIComponent(ident.person_id)}&channel=in.(meta,google)&select=channel,campaign,adset,ad,term,ts&order=ts.desc&limit=1`);
  const t = toques[0]; if (!t) return null;
  return { type: "anuncio", data: { channel: t.channel, campaign: t.campaign || "", adset: t.adset || "", ad: t.ad || "", keyword: t.term || "", matched_by: "phone_history", matched_at: t.ts } };
}

// Fallback conservador: casa o texto recebido com o texto do botao clicado, dentro de 15 minutos.
// So atribui quando todos os cliques candidatos com o mesmo texto sao do mesmo canal pago e nenhum ja foi usado.
async function recentClickOrigin(clientId: string | null, text: string, ts: string): Promise<{ type: string; data: Record<string, unknown> } | null> {
  if (!clientId) return null;
  const norm = _waMatchText(text); if (norm.length < 12) return null;
  const end = new Date(ts); if (isNaN(end.getTime())) return null;
  const start = new Date(end.getTime() - 15 * 60_000).toISOString();
  const evs = await sbSelect("track_events", `client_id=eq.${encodeURIComponent(clientId)}&type=eq.wpp_click&created_at=gte.${encodeURIComponent(start)}&created_at=lte.${encodeURIComponent(end.toISOString())}&select=id,created_at,utm_source,utm_medium,utm_campaign,utm_content,utm_term,gclid,fbclid,meta&order=created_at.desc&limit=30`);
  const matches = evs.filter((e: any) => {
    let preset = ""; try { preset = new URL(String(e.meta?.dest || "")).searchParams.get("text") || ""; } catch (_e) { /* */ }
    return _waMatchText(preset) === norm;
  }).filter((e: any) => _channelOf(e.utm_source || "", e.gclid || "", e.fbclid || "") !== "link");
  if (!matches.length) return null;
  const channels = [...new Set(matches.map((e: any) => _channelOf(e.utm_source || "", e.gclid || "", e.fbclid || "")))];
  if (channels.length !== 1) return null;
  for (const e of matches) {
    const used = await sbSelect("wa_conversations", `client_id=eq.${encodeURIComponent(clientId)}&origin-%3E%3Etrack_event_id=eq.${encodeURIComponent(e.id)}&select=id&limit=1`);
    if (used.length) continue;
    const m = e.meta || {}, channel = channels[0];
    return { type: "anuncio", data: { channel, track_source: e.utm_source || channel, medium: e.utm_medium || "", campaign: e.utm_campaign || m.campaignid || "", adset: e.utm_content || m.adgroupid || "", ad: m.adid || e.utm_term || "", keyword: e.utm_term || m.keyword || "", adgroup: m.adgroupid || "", gclid: e.gclid || "", fbclid: e.fbclid || "", track_event_id: e.id, attribution_method: "click_text_time", confidence: "inferida" } };
  }
  return null;
}
function waMsgText(m: any): string {
  const c = m.content || {};
  return m.text || c.text || c.conversation || c.extendedTextMessage?.text || c.imageMessage?.caption || c.videoMessage?.caption || "";
}
// uazapi manda messageTimestamp em MILISSEGUNDOS (13 dígitos); alguns em segundos (10). Normaliza.
function waTs(v: any): string { const n = Number(v) || 0; if (!n) return new Date().toISOString(); return new Date(n > 1e12 ? n : n * 1000).toISOString(); }
// Recebe eventos do uazapi (verify_jwt=false). Guarda cru + normaliza conversa/mensagem/atribuição.

// Evolution API manda {event:"messages.upsert", instance, data:{key:{remoteJid,fromMe,id}, pushName, message:{...}, messageTimestamp, contextInfo}}
// Traduz pro mesmo formato que o resto do código já entende (uazapi/Baileys).
function evolutionToMsgs(body: any): any[] {
  const ev = String(body?.event || "").toLowerCase();
  if (ev && !ev.includes("messages")) return []; // só mensagens (connection.update é tratado antes)
  const arr = Array.isArray(body?.data) ? body.data : (body?.data ? [body.data] : []);
  return arr.map((d: any) => {
    const k = d.key || {};
    const msg = d.message || {};
    const ctx = msg.extendedTextMessage?.contextInfo || msg.imageMessage?.contextInfo || msg.videoMessage?.contextInfo || d.contextInfo || {};
    const jid = String(k.remoteJid || d.remoteJid || "");
    const txt = msg.conversation || msg.extendedTextMessage?.text || msg.imageMessage?.caption || msg.videoMessage?.caption || d.text || "";
    const mt = String(d.messageType || Object.keys(msg)[0] || "").toLowerCase();
    const tipo = /audio|ptt/.test(mt) ? "audio" : /image/.test(mt) ? "image" : /document/.test(mt) ? "document" : "";
    return {
      chatid: jid, sender_pn: jid, sender: jid,
      isGroup: jid.endsWith("@g.us"),
      fromMe: !!k.fromMe,
      wasSentByApi: !!d.wasSentByApi,
      messageid: k.id || d.id || "", id: k.id || d.id || "",
      text: txt,
      senderName: d.pushName || d.senderName || "",
      messageType: d.messageType || "",
      messageTimestamp: d.messageTimestamp || d.timestamp || Date.now(),
      mediaType: tipo,
      // o extrator de origem (CTWA) lê content.contextInfo — mantém o mesmo caminho
      content: { text: txt, contextInfo: ctx, ...msg },
      _evolution: true,
    };
  }).filter((m: any) => m.chatid);
}

// WhatsApp Cloud API (Meta oficial): {object:"whatsapp_business_account", entry:[{changes:[{field:"messages", value:{contacts,messages,metadata}}]}]}
// Traduz pro formato interno. O bônus: o objeto "referral" traz o AD_ID REAL do clique no anúncio (CTWA).
function cloudApiToMsgs(body: any): any[] {
  const out: any[] = [];
  for (const ent of (body?.entry || [])) {
    for (const ch of (ent.changes || [])) {
      const v = ch.value || {};
      if (!Array.isArray(v.messages)) continue;
      const nomePorWaid: Record<string, string> = {};
      (v.contacts || []).forEach((c: any) => { if (c.wa_id) nomePorWaid[c.wa_id] = c.profile?.name || ""; });
      for (const m of v.messages) {
        const from = String(m.from || "");
        const txt = m.text?.body || m.button?.text || m.interactive?.list_reply?.title || m.interactive?.button_reply?.title
          || m.image?.caption || m.video?.caption || m.document?.caption || "";
        const t = String(m.type || "").toLowerCase();
        const tipo = /audio|voice/.test(t) ? "audio" : /image|sticker/.test(t) ? "image" : /document/.test(t) ? "document" : "";
        // referral = clique no anúncio (CTWA) com o id do anúncio de verdade
        const rf = m.referral || null;
        const ctx = rf ? { externalAdReply: {
          sourceId: rf.source_id || "", sourceType: rf.source_type || "", sourceUrl: rf.source_url || "",
          ctwaClid: rf.ctwa_clid || "", title: rf.headline || "", body: rf.body || "",
          thumbnailUrl: rf.image_url || rf.thumbnail_url || "", mediaType: rf.media_type || "",
        } } : {};
        out.push({
          chatid: from, sender_pn: from, sender: from, isGroup: false,
          fromMe: false, wasSentByApi: false,
          messageid: m.id || "", id: m.id || "",
          text: txt, senderName: nomePorWaid[from] || "",
          messageType: m.type || "text",
          messageTimestamp: (Number(m.timestamp) || 0) * 1000,
          mediaType: tipo,
          content: { text: txt, contextInfo: ctx },
          _cloud: true, _phoneNumberId: v.metadata?.phone_number_id || "",
        });
      }
    }
  }
  return out;
}
async function handleWaWebhook(instId: string, req: Request): Promise<Response> {
  const ok = () => new Response("ok", { headers: { ...cors, "Content-Type": "text/plain" } });
  let body: any;
  try { body = await req.json(); } catch { return ok(); }
  const inst = (await sbSelect("wa_instances", `id=eq.${encodeURIComponent(instId)}&select=id,client_id,phone`))[0];
  if (!inst) return ok(); // instância desconhecida: ignora silenciosamente
  const clientId = inst.client_id || null;
  const evt = body.EventType || body.event || body.type || "";
  // conexão: atualiza status
  if (String(evt).toLowerCase().includes("connect") || body.connection || body.status) {
    const st = body.status || body.connection || (body.instance && body.instance.status);
    if (st) await sbPatch("wa_instances", `id=eq.${encodeURIComponent(instId)}`, { status: String(st), updated_at: new Date().toISOString() });
    return ok();
  }
  // mensagens: pode vir 1 ou várias
  let msgs: any[] = [];
  // WhatsApp Cloud API (Meta oficial)
  if (body.object === "whatsapp_business_account" && Array.isArray(body.entry)) msgs = cloudApiToMsgs(body);
  // Evolution API (provider=evolution): payload {event:"messages.upsert", data:{key,message,...}}
  else
  if (String(body.event || "").includes(".") && body.data && (body.data.key || (Array.isArray(body.data) && body.data[0]?.key))) msgs = evolutionToMsgs(body);
  else if (Array.isArray(body.messages)) msgs = body.messages;
  else if (body.message) msgs = [body.message];
  else if (body.data && Array.isArray(body.data)) msgs = body.data;
  else if (body.data && body.data.message) msgs = [body.data.message];
  else if (body.chatid || body.messageid) msgs = [body];
  // Formato desconhecido: guarda o payload pra eu conseguir escrever o tradutor (sem isso o webhook some em silêncio)
  if (!msgs.length) {
    try { await sbInsert("wa_webhook_log", { id: uid(), instance_id: instId, reconhecido: false, payload: body }); } catch (_e) { /* segue */ }
    return ok();
  }
  // Instância da AndréIA: encaminha as mensagens do grupo configurado pro cérebro dela (não vai pro CRM)
  const aw = (((await sbSelect("account_config", "id=eq.main&select=data"))[0]?.data) || {}).andreia_wa || {};
  if (aw.instance_id === instId) {
    for (const m of msgs) {
      const chatid = String(m.chatid || "");
      if (aw.group_jid && chatid !== aw.group_jid) continue;
      const sender = String(m.sender_pn || m.sender || "").replace(/@.*$/, "");
      // mídia: áudio (ptt/audio), imagem, documento (pdf) → a AndréIA baixa e processa
      const _mt = String(m.mediaType || "").toLowerCase();
      const _mime = String((m.content && m.content.mimetype) || "");
      let _mkind = "";
      if (_mt === "ptt" || _mt === "audio" || /audio/.test(_mime)) _mkind = "audio";
      else if (_mt === "image" || /image/.test(_mime)) _mkind = "image";
      else if (_mt === "document" || /pdf/.test(_mime)) _mkind = "document";
      // O webhook do uazapi JÁ exclui os envios da própria AndréIA (excludeMessages:["wasSentByApi"]),
      // então não há eco/loop. Processa tudo que chega — inclusive texto/mídia da dona pelo próprio número (fromMe),
      // pra o "sim"/comandos dela funcionarem. Garante só pulando qualquer envio via API que porventura chegue.
      if (m.wasSentByApi) continue;
      const payload = { waAgent: { instanceId: instId, chatid, sender, text: waMsgText(m), msgid: String(m.messageid || m.id || ""), mtype: _mkind, mime: _mime, fname: String((m.content && (m.content.fileName || m.content.title)) || "") } };
      try { await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify(payload) }); } catch (_e) {}
    }
    return ok();
  }
  let needResolve = false;
  const toClassify = new Set<string>();
  for (const m of msgs) {
    const isGroup = m.isGroup || String(m.chatid || "").endsWith("@g.us");
    if (isGroup) continue; // CRM é só 1:1
    const chatJid = String(m.chatid || m.sender_pn || m.sender || "");
    const phone = chatJid.replace(/@.*$/, "").replace(/[^0-9]/g, "");
    if (!phone) continue;
    const fromMe = !!m.fromMe;
    const text = waMsgText(m);
    const ts = waTs(m.messageTimestamp);
    const msgid = m.messageid || m.id || uid();
    // conversa (upsert por client_id + chat_id)
    const existing = (await sbSelect("wa_conversations", `client_id=${clientId ? "eq." + encodeURIComponent(clientId) : "is.null"}&chat_id=eq.${encodeURIComponent(phone)}&select=id,origin_type,name&limit=1`))[0];
    let convId = existing?.id;
    let origin = fromMe ? null : waExtractOrigin(m);
    if (!fromMe) { const rf = await refOrigin(text); if (rf) origin = rf; }
    if (!fromMe && !origin && (!existing || !existing.origin_type || existing.origin_type === "organico")) origin = await recentClickOrigin(clientId, text, ts);
    if (!fromMe && !origin && (!existing || !existing.origin_type || existing.origin_type === "organico")) origin = await waPhoneHistoryOrigin(clientId, phone);
    if (!convId) {
      convId = uid();
      // nome vem SÓ de mensagem do LEAD (inbound) — em msg enviada, senderName é o dono da instância (ficava tudo com o mesmo nome)
      const leadName = (!fromMe && (m.senderName || m.pushName)) ? (m.senderName || m.pushName) : phone;
      await sbInsert("wa_conversations", {
        id: convId, client_id: clientId, chat_id: phone, name: leadName,
        last_text: text, last_at: ts, unread: fromMe ? 0 : 1,
        origin_type: origin ? origin.type : "organico", origin: origin ? origin.data : null,
      });
      if (origin && origin.type === "anuncio") needResolve = true;
    } else {
      const patch: Record<string, unknown> = { last_text: text, last_at: ts };
      if (!fromMe) patch.unread = 1;
      // atualiza o nome só com mensagem do LEAD (inbound); nunca com o nome do dono da instância (msg enviada)
      const _curName = existing.name || "";
      if (!fromMe && m.senderName && m.senderName !== _curName) patch.name = m.senderName;
      // grava atribuição só se ainda não tiver (primeira toca ganha)
      if (origin && (!existing.origin_type || existing.origin_type === "organico")) { patch.origin_type = origin.type; patch.origin = origin.data; if (origin.type === "anuncio") needResolve = true; }
      await sbPatch("wa_conversations", `id=eq.${convId}`, patch);
    }
    await sbInsert("wa_messages", {
      id: uid(), client_id: clientId, conversation_id: convId, chat_id: phone, wa_msgid: String(msgid),
      direction: fromMe ? "out" : "in", msg_type: m.messageType || "text", text, ts, raw: m,
    });
    if (!fromMe && convId) toClassify.add(convId); // msg nova do lead → IA classifica sozinha (não depende do gestor abrir o CRM)
  }
  // classificação automática por IA das conversas que receberam msg do lead (independe do gestor estar online).
  // IMPORTANTE: sem waitUntil o worker morre ao responder "ok" e o fetch é abortado — a conversa ficava sem etapa.
  const bg: Promise<unknown>[] = [];
  for (const cid of [...toClassify].slice(0, 25)) {
    bg.push(fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ wa: { op: "extract", convId: cid, autoApply: true } }) }).catch(() => {}));
  }
  // chegou conversa de anúncio → resolve campanha›conjunto›anúncio na hora (não espera o cron de 20min)
  if (needResolve) {
    bg.push(fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ resolveAllOrigins: true }) }).catch(() => {}));
  }
  // segura o worker vivo até o trabalho de fundo terminar (senão o Deno aborta os fetches pendentes)
  if (bg.length) { try { (globalThis as any).EdgeRuntime?.waitUntil?.(Promise.all(bg)); } catch (_e) { /* ignora */ } }
  return ok();
}

// Página/endpoint público de conexão: devolve o QR + código de pareamento de uma instância (id não-adivinhável)
async function handleWaConnect(id: string, url: URL): Promise<Response> {
  const j = (o: unknown, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, "Content-Type": "application/json" } });
  const inst = (await sbSelect("wa_instances", `id=eq.${encodeURIComponent(id)}&select=id,name,uaz_host,uaz_token,status`))[0];
  if (!inst) return j({ error: "not_found" }, 404);
  const host = String(inst.uaz_host || "").replace(/\/$/, ""); const token = inst.uaz_token;
  const call = async (path: string, method = "GET", body?: any) => {
    try { const r = await fetch(host + path, { method, headers: { token, "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined }); return await r.json(); } catch { return {}; }
  };
  const phone = (url.searchParams.get("phone") || "").replace(/[^0-9]/g, "");
  let ins: any = ((await call("/instance/status")) || {}).instance || {};
  if (ins.status !== "connected" && (!ins.qrcode || phone)) {
    const conn = await call("/instance/connect", "POST", phone ? { phone } : {});
    ins = (conn && conn.instance) ? conn.instance : (conn || ins);
  }
  if (ins.status) { const patch: Record<string, unknown> = { status: ins.status, updated_at: new Date().toISOString() }; if (ins.owner) patch.phone = String(ins.owner).replace(/@.*$/, ""); if (ins.status === "connected") patch.connected_at = new Date().toISOString(); await sbPatch("wa_instances", `id=eq.${encodeURIComponent(id)}`, patch); }
  return j({ status: ins.status || "connecting", qrcode: ins.qrcode || "", paircode: ins.paircode || ins.pairCode || ins.code || ins.paircode_text || "", name: inst.name || "" });
}

// ---- Google Agenda (iCal) → tarefas: lê o link secreto, cria tarefa por reunião (dedup por UID) ----
function icsUnfold(t: string) { return t.replace(/\r\n[ \t]/g, "").replace(/\n[ \t]/g, ""); }
function icsParse(text: string) {
  const lines = icsUnfold(text).split(/\r?\n/); const evs: any[] = []; let cur: any = null;
  for (const line of lines) {
    if (line === "BEGIN:VEVENT") { cur = { attendees: 0 }; continue; }
    if (line === "END:VEVENT") { if (cur) evs.push(cur); cur = null; continue; }
    if (!cur) continue;
    const i = line.indexOf(":"); if (i < 0) continue;
    const keyfull = line.slice(0, i), val = line.slice(i + 1); const key = keyfull.split(";")[0].toUpperCase();
    if (key === "SUMMARY") cur.summary = val;
    else if (key === "UID") cur.uid = val;
    else if (key === "DTSTART") { cur.start = val; cur.allday = /VALUE=DATE(;|$|:)/i.test(keyfull) || /^\d{8}$/.test(val); }
    else if (key === "LOCATION") cur.location = val;
    else if (key === "DESCRIPTION") cur.description = val;
    else if (key === "STATUS") cur.status = val;
    else if (key === "ATTENDEE") cur.attendees++;
    else if (key === "X-GOOGLE-CONFERENCE" || key === "X-GOOGLE-HANGOUT") cur.conference = val;
  }
  return evs;
}
function icsDate(val: string) {
  const m = String(val).match(/^(\d{4})(\d{2})(\d{2})(T(\d{2})(\d{2})(\d{2}))?(Z)?/); if (!m) return null;
  const ymd = `${m[1]}-${m[2]}-${m[3]}`, hm = m[5] ? `${m[5]}:${m[6]}` : "";
  const iso = `${ymd}T${m[5] || "00"}:${m[6] || "00"}:00${m[8] ? "Z" : ""}`; const d = new Date(iso);
  return { ymd, hm, time: isNaN(d.getTime()) ? Date.parse(ymd) : d.getTime() };
}
// ===== Google Agenda via OAuth (Calendar API) — modo preferido; iCal segue como fallback =====
// Prefere um cliente OAuth dedicado (GOOGLE_CAL_*) se existir; senão reaproveita o do Google Ads.
const GOOGLE_CID = Deno.env.get("GOOGLE_CAL_CLIENT_ID") || Deno.env.get("GOOGLE_ADS_CLIENT_ID") || "";
const GOOGLE_CSEC = Deno.env.get("GOOGLE_CAL_CLIENT_SECRET") || Deno.env.get("GOOGLE_ADS_CLIENT_SECRET") || "";
function googleRedirectUri() { return `${SB_URL}/functions/v1/tracking/google/callback`; }
async function saveAcctData(patch: Record<string, unknown>) {
  const cur = ((await sbSelect("account_config", "id=eq.main&select=data"))[0]?.data || {}) as any;
  await sbPatch("account_config", "id=eq.main", { data: { ...cur, ...patch } });
}
function handleGoogleAuth(): Response {
  if (!GOOGLE_CID) return new Response("GOOGLE_ADS_CLIENT_ID não configurado nos secrets.", { status: 500, headers: cors });
  const p = new URLSearchParams({
    client_id: GOOGLE_CID, redirect_uri: googleRedirectUri(), response_type: "code",
    access_type: "offline", prompt: "consent", include_granted_scopes: "true",
    scope: "https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/userinfo.email",
  });
  return new Response(null, { status: 302, headers: { ...cors, Location: `https://accounts.google.com/o/oauth2/v2/auth?${p}` } });
}
function _oauthPage(title: string, msg: string) {
  return new Response(`<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><body style="font-family:system-ui,-apple-system,sans-serif;background:#0f0f12;color:#eee;display:flex;align-items:center;justify-content:center;height:100vh;margin:0"><div style="text-align:center;max-width:420px;padding:24px"><h2 style="margin:0 0 8px">${title}</h2><p style="color:#aaa;line-height:1.5">${msg}</p><p style="color:#666;font-size:13px;margin-top:18px">Pode fechar esta aba.</p></div><script>try{setTimeout(function(){window.close()},2500)}catch(e){}</script>`, { headers: { ...cors, "Content-Type": "text/html; charset=utf-8" } });
}
async function handleGoogleCallback(url: URL): Promise<Response> {
  const code = url.searchParams.get("code"), err = url.searchParams.get("error");
  if (err) return _oauthPage("Conexão cancelada", "Você recusou o acesso (" + err + ").");
  if (!code) return _oauthPage("Erro", "Não recebi o código do Google.");
  try {
    const tr = await fetch("https://oauth2.googleapis.com/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ code, client_id: GOOGLE_CID, client_secret: GOOGLE_CSEC, redirect_uri: googleRedirectUri(), grant_type: "authorization_code" }) });
    const tj = await tr.json();
    if (!tj.refresh_token) return _oauthPage("Quase lá", "O Google não devolveu autorização permanente. Vá em <b>myaccount.google.com/permissions</b>, remova o app e tente conectar de novo (precisa aparecer a tela de consentimento).");
    let email = "";
    try { const ur = await fetch("https://www.googleapis.com/oauth2/v2/userinfo", { headers: { Authorization: `Bearer ${tj.access_token}` } }); email = (await ur.json()).email || ""; } catch (_e) { /* */ }
    await saveAcctData({ google_cal: { refresh_token: tj.refresh_token, email, connected_at: new Date().toISOString() } });
    return _oauthPage("✅ Google conectado!", `Agenda de <b>${email || "sua conta"}</b> conectada. Suas reuniões vão virar tarefas automaticamente.`);
  } catch (e) { return _oauthPage("Erro", String(e)); }
}
async function googleCalAccessToken(): Promise<string | null> {
  const cfg = ((await sbSelect("account_config", "id=eq.main&select=data"))[0]?.data || {}) as any;
  const rt = cfg.google_cal?.refresh_token; if (!rt) return null;
  const r = await fetch("https://oauth2.googleapis.com/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: GOOGLE_CID, client_secret: GOOGLE_CSEC, refresh_token: rt, grant_type: "refresh_token" }) });
  const j = await r.json();
  return j.access_token || null;
}
function mapGEvent(ev: any) {
  const raw = ev.start?.dateTime || ev.start?.date || "";
  const d = raw ? { time: new Date(raw).getTime(), ymd: String(raw).slice(0, 10), hm: ev.start?.dateTime ? String(raw).slice(11, 16) : "" } : null;
  const conf = ev.hangoutLink || (ev.conferenceData?.entryPoints || []).map((e: any) => e.uri).find(Boolean) || "";
  return { uid: ev.id, summary: ev.summary || "", cancelled: ev.status === "cancelled", attendees: (ev.attendees || []).length, conference: conf, description: ev.description || "", location: ev.location || "", d };
}
async function createMeetingTasks(evs: any[], j: (o: unknown, s?: number) => Response) {
  const now = Date.now(), horizon = now + 60 * 864e5;
  const clients = await sbSelect("clients", "select=id,name&limit=1000");
  let created = 0, skipped = 0;
  for (const ev of evs) {
    if (!ev.uid || !ev.summary) continue;
    if (ev.cancelled) continue;
    const isMeeting = ev.attendees > 0 || ev.conference || /meet\.google|zoom\.us|teams\.microsoft|hangout|whereby|meet\.jit/i.test((ev.description || "") + " " + (ev.location || ""));
    if (!isMeeting) continue;
    const d = ev.d; if (!d) continue;
    if (d.time < now - 864e5 || d.time > horizon) continue;
    if ((await sbSelect("calendar_events", `uid=eq.${encodeURIComponent(ev.uid)}&select=uid&limit=1`)).length) { skipped++; continue; }
    let clientId: string | null = null; const sl = ev.summary.toLowerCase();
    const mc = clients.find((c: any) => c.name && c.name.length >= 4 && sl.includes(c.name.toLowerCase()));
    if (mc) clientId = mc.id;
    const meet = ev.conference || ((ev.description || "").match(/https?:\/\/(meet\.google|[^\s"']*zoom\.us|teams\.microsoft)[^\s"'<>]*/i) || [])[0] || "";
    const notes = `🗓 Reunião (Google Agenda)${d.hm ? ` · ${d.hm}` : ""}${meet ? `\n${meet}` : ""}${ev.location ? `\n📍 ${ev.location}` : ""}`;
    const tid = "cal" + uid();
    await sbInsert("tasks", { id: tid, name: ev.summary.slice(0, 200), client: clientId, owner: "eu", status: "todo", prio: "media", due: d.ymd, notes, urgent: false });
    await sbInsert("calendar_events", { uid: ev.uid, task_id: tid, summary: ev.summary.slice(0, 200), start: d.ymd });
    created++;
  }
  return j({ created, skipped });
}
async function handleCalendarSync(): Promise<Response> {
  const j = (o: unknown, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, "Content-Type": "application/json" } });
  const cfg = ((await sbSelect("account_config", "id=eq.main&select=data"))[0]?.data || {}) as any;
  // Modo preferido: OAuth (Calendar API)
  if (cfg.google_cal?.refresh_token) {
    const tok = await googleCalAccessToken();
    if (!tok) return j({ error: "Não consegui renovar o acesso ao Google — reconecte sua conta." });
    const now = Date.now(), horizon = now + 60 * 864e5;
    const p = new URLSearchParams({ singleEvents: "true", orderBy: "startTime", maxResults: "250", timeMin: new Date(now - 864e5).toISOString(), timeMax: new Date(horizon).toISOString() });
    const r = await fetch(`https://www.googleapis.com/calendar/v3/calendars/primary/events?${p}`, { headers: { Authorization: `Bearer ${tok}` } });
    const gj = await r.json();
    if (gj.error) return j({ error: "Google Calendar: " + (gj.error.message || JSON.stringify(gj.error)) });
    return await createMeetingTasks((gj.items || []).map(mapGEvent), j);
  }
  // Fallback: iCal
  const url = cfg.calendar_ics;
  if (!url) return j({ error: "Conecte sua conta Google (ou cole o link iCal)." });
  let text = ""; try { const r = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0", "Accept": "text/calendar,*/*" } }); text = await r.text(); } catch (e) { return j({ error: "Falha ao ler o iCal: " + String(e) }); }
  if (!/BEGIN:VCALENDAR/.test(text)) return j({ error: "O link não parece um iCal válido (comece pelo 'Endereço secreto no formato iCal')." });
  const evs = icsParse(text).map((e: any) => ({ uid: e.uid, summary: e.summary, cancelled: String(e.status || "").toUpperCase() === "CANCELLED", attendees: e.attendees, conference: e.conference, description: e.description, location: e.location, d: icsDate(e.start) }));
  return await createMeetingTasks(evs, j);
}
async function sbDelete(table: string, query: string) {
  await fetch(`${SB_URL}/rest/v1/${table}?${query}`, { method: "DELETE", headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}` } });
}
// Cria uma reunião na Google Agenda (escrita) + tarefa local (id cal*). body: {summary,date,time?,durationMin?,description?,clientId?}
async function handleCalCreate(req: Request): Promise<Response> {
  const j = (o: unknown, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, "Content-Type": "application/json" } });
  let b: any = {}; try { b = await req.json(); } catch (_e) { /* */ }
  const summary = String(b.summary || b.titulo || "").trim(); if (!summary) return j({ error: "título obrigatório" }, 400);
  const date = String(b.date || b.data || "").slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return j({ error: "data inválida (use AAAA-MM-DD)" }, 400);
  const time = String(b.time || b.hora || "").slice(0, 5);
  const dur = Number(b.durationMin || b.duracao || 60) || 60;
  const tok = await googleCalAccessToken(); if (!tok) return j({ error: "Google não conectado — conecte sua conta primeiro." }, 400);
  const tz = "America/Sao_Paulo";
  let ebody: any;
  if (/^\d{2}:\d{2}$/.test(time)) {
    const [hh, mm] = time.split(":").map(Number); const endM = hh * 60 + mm + dur;
    const eh = String(Math.floor(endM / 60) % 24).padStart(2, "0"), em = String(endM % 60).padStart(2, "0");
    ebody = { summary, description: b.description || b.obs || "", start: { dateTime: `${date}T${time}:00`, timeZone: tz }, end: { dateTime: `${date}T${eh}:${em}:00`, timeZone: tz } };
  } else {
    ebody = { summary, description: b.description || b.obs || "", start: { date }, end: { date } };
  }
  const r = await fetch("https://www.googleapis.com/calendar/v3/calendars/primary/events", { method: "POST", headers: { Authorization: `Bearer ${tok}`, "Content-Type": "application/json" }, body: JSON.stringify(ebody) });
  const ev = await r.json();
  if (ev.error) { if (r.status === 403 || r.status === 401) return j({ error: "reconnect", detail: "Precisa reconectar o Google com permissão de edição da agenda." }, 403); return j({ error: "Google: " + (ev.error.message || "falha ao criar") }, 400); }
  const clients = await sbSelect("clients", "select=id,name&limit=1000");
  let clientId = b.clientId || null;
  if (!clientId) { const sl = summary.toLowerCase(); const mc = clients.find((c: any) => c.name && c.name.length >= 4 && sl.includes(c.name.toLowerCase())); if (mc) clientId = mc.id; }
  const notes = `🗓 Reunião (Google Agenda)${time ? ` · ${time}` : ""}`;
  const tid = "cal" + uid();
  await sbInsert("tasks", { id: tid, name: summary.slice(0, 200), client: clientId, owner: "eu", status: "todo", prio: "media", due: date, notes, urgent: false });
  await sbInsert("calendar_events", { uid: ev.id, task_id: tid, summary: summary.slice(0, 200), start: date });
  return j({ ok: true, id: ev.id, taskId: tid, date, time });
}
// Exclui uma reunião: apaga no Google Agenda + tarefa/calendar_event local. body: {taskId?} ou {uid?}
async function handleCalDelete(req: Request): Promise<Response> {
  const j = (o: unknown, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, "Content-Type": "application/json" } });
  let b: any = {}; try { b = await req.json(); } catch (_e) { /* */ }
  let uidVal = String(b.uid || ""), tid = String(b.taskId || b.id || "");
  if (!uidVal && tid) { const row = (await sbSelect("calendar_events", `task_id=eq.${encodeURIComponent(tid)}&select=uid&limit=1`))[0]; if (row) uidVal = row.uid; }
  if (!tid && uidVal) { const row = (await sbSelect("calendar_events", `uid=eq.${encodeURIComponent(uidVal)}&select=task_id&limit=1`))[0]; if (row) tid = row.task_id; }
  if (!uidVal && !tid) return j({ error: "reunião não encontrada" }, 404);
  if (uidVal) {
    const tok = await googleCalAccessToken();
    if (tok) { const r = await fetch(`https://www.googleapis.com/calendar/v3/calendars/primary/events/${encodeURIComponent(uidVal)}`, { method: "DELETE", headers: { Authorization: `Bearer ${tok}` } }); if (r.status === 403 || r.status === 401) return j({ error: "reconnect", detail: "Precisa reconectar o Google com permissão de edição." }, 403); /* 404/410 = já não existe no Google, segue */ }
  }
  if (tid) await sbDelete("tasks", `id=eq.${encodeURIComponent(tid)}`);
  if (uidVal) await sbDelete("calendar_events", `uid=eq.${encodeURIComponent(uidVal)}`);
  return j({ ok: true });
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
    let showBanner = true;
    if (cid) { const cr = await sbSelect("clients", `id=eq.${encodeURIComponent(cid)}&select=pixel_banner_proprio&limit=1`); showBanner = !cr[0]?.pixel_banner_proprio; }
    const js = cid ? pixelScript(cid, base, showBanner) : `console.warn("[Rastreamento] token inválido");`;
    return new Response(js, { headers: { ...cors, "Content-Type": "application/javascript; charset=utf-8", "Cache-Control": "public, max-age=300" } });
  }

  if (p === "/collect" && req.method === "POST") return handleCollect(req);
  if (p === "/wa/ref" && req.method === "POST") return handleWaRef(req);

  if (p === "/rd/callback") return handleRdCallback(url);

  if (p === "/rd/webhook" && req.method === "POST") return handleRdWebhook(url, req);
  if (p === "/rd/webhook") return new Response("rd webhook ok", { headers: { ...cors, "Content-Type": "text/plain" } });

  if (p === "/nuvemshop/callback") return handleNuvemshopCallback(url);

  if (p === "/tiktok/callback") return handleTikTokCallback(url);
  if (p === "/pinterest/callback") return handlePinterestCallback(url);

  // GET /calendar/sync -> lê o Google Agenda (OAuth ou iCal) e cria tarefas das reuniões
  if (p === "/calendar/sync") return handleCalendarSync();
  if (p === "/calendar/create" && req.method === "POST") return handleCalCreate(req);
  if (p === "/calendar/delete" && req.method === "POST") return handleCalDelete(req);

  // OAuth do Google Agenda (Calendar API)
  if (p === "/google/auth") return handleGoogleAuth();
  if (p === "/google/callback") return handleGoogleCallback(url);
  if (p === "/google/status") { const cfg = ((await sbSelect("account_config", "id=eq.main&select=data"))[0]?.data || {}) as any; return new Response(JSON.stringify({ connected: !!cfg.google_cal?.refresh_token, email: cfg.google_cal?.email || "" }), { headers: { ...cors, "Content-Type": "application/json" } }); }
  if (p === "/google/disconnect") { await saveAcctData({ google_cal: null }); return new Response(JSON.stringify({ ok: true }), { headers: { ...cors, "Content-Type": "application/json" } }); }

  // /automations/tick -> dispara as automações da AndréIA que estão no horário (chamado pelo cron)
  if (p === "/automations/tick" || p === "/automations/reminders" || p === "/wa/connectivity" || p === "/wa/resolve-origins" || p === "/journey/orders-tick" || p === "/wa/agency-poll") {
    const payload = p === "/automations/reminders" ? { reminderTick: true } : (p === "/wa/connectivity" ? { waConnCheck: true } : (p === "/wa/resolve-origins" ? { resolveAllOrigins: true } : (p === "/journey/orders-tick" ? { journeyOrdersTick: { blocos: 2 } } : (p === "/wa/agency-poll" ? { waAgencyPollTick: true } : { automationsTick: true }))));
    try {
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // /metrics/tick -> salva o historico de channel_metrics_daily. Sem parametros: o dia de ontem (America/Sao_Paulo), pra
  // todos os clientes (chamado pelo cron diario). Aceita ?since=&until=&clientIds=a,b,c pra reconstruir o passado em lotes.
  if (p === "/metrics/tick") {
    let since = url.searchParams.get("since") || "", until = url.searchParams.get("until") || "";
    if (!since || !until) {
      const now = new Date(new Date().toLocaleString("en-US", { timeZone: "America/Sao_Paulo" }));
      now.setDate(now.getDate() - 1);
      since = until = now.toISOString().slice(0, 10);
    }
    const clientIdsParam = url.searchParams.get("clientIds");
    const clientIds = clientIdsParam ? clientIdsParam.split(",").map((s) => s.trim()).filter(Boolean) : undefined;
    try {
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ channelMetricsCollect: { since, until, clientIds } }) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // /instagram/followers-tick -> snapshot diario de seguidores no schema `midia` (Fase 3, chamado pelo cron diario).
  if (p === "/instagram/followers-tick") {
    try {
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ instagramFollowersSnapshot: true }) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // /instagram/organic-tick -> historiza conteudo organico no schema `midia` (Fase 4, chamado pelo cron diario).
  if (p === "/instagram/organic-tick") {
    try {
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ instagramOrganicSnapshot: { days: 30 } }) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // GET /instagram/diag -> diagnostico temporario: o META_USER_TOKEN ja enxerga Paginas/Instagram vinculado?
  if (p === "/instagram/diag") {
    try {
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ instagramDiag: true }) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // GET /instagram/organico?clientId=... -> teste manual dos posts organicos + metricas reais desse cliente
  if (p === "/instagram/organico") {
    try {
      const clientId = url.searchParams.get("clientId") || "";
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ instagramOrganicContent: { clientId, days: 90 } }) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // /security/audit -> roda public.security_audit() (RLS aberta pro anon, tabela sem RLS, view sem
  // security_invoker) e avisa a equipe (sino + grupo) so quando aparece achado novo. Chamado pelo cron diario.
  if (p === "/security/audit") {
    try {
      const r = await fetch(`${SB_URL}/functions/v1/dynamic-responder`, { method: "POST", headers: { Authorization: `Bearer ${SB_KEY}`, apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ securityAuditTick: true }) });
      const t = await r.text();
      return new Response(t, { status: r.status, headers: { ...cors, "Content-Type": "application/json" } });
    } catch (e) { return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...cors, "Content-Type": "application/json" } }); }
  }

  // GET /wa/connect/<instanceId> -> JSON com qrcode/paircode/status (usado pela página pública de conexão)
  const mWaC = p.match(/^\/wa\/connect\/([^/]+)$/);
  if (mWaC) return handleWaConnect(mWaC[1], url);

  // POST /wa/webhook/<instanceId>  -> ingere eventos do uazapi (mensagens/conexão) da instância
  const mWa = p.match(/^\/wa\/webhook\/([^/]+)$/);
  if (mWa) {
    // Verificação do webhook da Cloud API (Meta): GET com hub.challenge — o token de verificação é o id da instância
    if (req.method === "GET" && url.searchParams.get("hub.mode") === "subscribe") {
      const tk = url.searchParams.get("hub.verify_token") || "";
      const ch = url.searchParams.get("hub.challenge") || "";
      if (tk && tk === mWa[1]) return new Response(ch, { headers: { "Content-Type": "text/plain" } });
      return new Response("token invalido", { status: 403, headers: { "Content-Type": "text/plain" } });
    }
    if (req.method !== "POST") return new Response("wa webhook ok", { headers: { ...cors, "Content-Type": "text/plain" } });
    return handleWaWebhook(mWa[1], req);
  }

  // GET /l/<client_id>/<slug>
  const mLink = p.match(/^\/l\/([^/]+)\/([^/]+)$/);
  if (mLink) return handleRedirect(decodeURIComponent(mLink[1]), decodeURIComponent(mLink[2]), url, ref);

  // Link CURTO com o handle do cliente: GET /<handle> ou /<handle>/<slug>  (ex.: /motirobar/bio)
  // handle vazio → usa o slug "bio" como padrão. Resolve o handle → client_id e reaproveita o redirect rastreado.
  const RESERVED = new Set(["collect", "wa", "rd", "google", "nuvemshop", "tiktok", "pinterest", "calendar", "automations", "pixel", "l", "favicon.ico", "robots.txt"]);
  const mHandle = p.match(/^\/([a-z0-9][a-z0-9_-]{1,60})(?:\/([a-z0-9_-]{1,60}))?\/?$/i);
  if (mHandle && req.method === "GET" && !RESERVED.has(mHandle[1].toLowerCase())) {
    const handle = mHandle[1].toLowerCase();
    const slug = (mHandle[2] || "bio").toLowerCase();
    const cfg = await sbSelect("tracking_config", `handle=eq.${encodeURIComponent(handle)}&select=client_id&limit=1`);
    const cid = cfg[0]?.client_id;
    if (cid) return handleRedirect(String(cid), slug, url, ref);
  }

  return new Response("tracking ok", { headers: { ...cors, "Content-Type": "text/plain" } });
});
