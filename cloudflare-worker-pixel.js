// Cloudflare Worker — pixel/links de rastreio no domínio próprio (pixel.gt-marketing.app.br).
//
// GET /pixel/script/<token>.js e POST /collect são respondidos AQUI DIRETO, gravando via REST/Data API
// do Supabase (Prefer: return=minimal) — isso NÃO conta na cota de "Invocações de Funções de Borda" do
// Supabase, só a Edge Function conta. Antes, este Worker so repassava 1:1 pra Edge Function "tracking",
// e cada pageview de cada cliente virava 2 invocações (script + collect) — foi isso que estourou a cota
// (196% do limite, ver Uso do projeto). O resto das rotas (links /l/..., /wa/ref, /wa/webhook etc., baixo
// volume) continua indo pra Edge Function normalmente, sem mudança nenhuma.
//
// Deploy: Cloudflare -> Workers & Pages -> abrir o worker que já responde por pixel.gt-marketing.app.br
// -> Editar código -> colar isto -> Deploy.
// Secret necessário (Settings -> Variables and Secrets -> Add -> tipo "Secret", não "Text"):
//   SUPABASE_SERVICE_ROLE_KEY = a MESMA chave que a Edge Function "tracking" já usa hoje (cofre local).
//   Sem esse secret o Worker cai pro proxy antigo em tudo (nada quebra, só não corrige a cota).

const SUPABASE_URL = "https://mocrfqmdjwvyhqvdpimm.supabase.co";
const TRACKING_FN = SUPABASE_URL + "/functions/v1/tracking";
const REST = SUPABASE_URL + "/rest/v1";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "content-type",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
};

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
}
function clip(v, n) {
  if (v == null) return "";
  return String(v).slice(0, n);
}

// Copia 1:1 da função pixelScript(cid, base, showBanner) de functions/tracking.ts — qualquer mudança no
// comportamento do pixel (novos eventos, novos campos, banner) precisa ser replicada nos dois lugares.
// showBanner=true (padrao): mostra banner leve de cookies ANTES de rastrear, so roda se aceitar (LGPD).
// showBanner=false (cliente marcou "ja tem aviso proprio" em Configuracoes): roda sem banner, como antes.
function pixelScript(cid, base, showBanner) {
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

// token (tk_...) -> client_id, com cache de edge (Cache API) de 5min pra nao bater a REST API a cada
// pageview/evento — o mesmo cache serve tanto o /pixel/script quanto o /collect.
async function resolveCid(env, raw, ctx) {
  const v = clip(raw, 60);
  if (!v || !v.startsWith("tk_")) return v;
  const cache = caches.default;
  const cacheKey = new Request("https://cache.internal/tk-cid/" + v);
  const hit = await cache.match(cacheKey);
  if (hit) return await hit.text();
  const r = await fetch(`${REST}/tracking_config?token=eq.${encodeURIComponent(v)}&select=client_id&limit=1`, {
    headers: { apikey: env.SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}` },
  });
  const j = await r.json().catch(() => []);
  const cid = (j[0] && j[0].client_id) || "";
  ctx.waitUntil(cache.put(cacheKey, new Response(cid, { headers: { "Cache-Control": "max-age=300" } })));
  return cid;
}

// se o cliente ja tem aviso de cookies proprio no site (clients.pixel_banner_proprio), o nosso pixel NAO
// mostra banner - confia no aviso que ja existe. Cache de edge 10min (o campo quase nunca muda).
async function resolveShowBanner(env, cid, ctx) {
  const cache = caches.default;
  const cacheKey = new Request("https://cache.internal/banner-flag/" + cid);
  const hit = await cache.match(cacheKey);
  if (hit) return (await hit.text()) !== "0";
  const r = await fetch(`${REST}/clients?id=eq.${encodeURIComponent(cid)}&select=pixel_banner_proprio&limit=1`, {
    headers: { apikey: env.SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}` },
  });
  const j = await r.json().catch(() => []);
  const temProprio = !!(j[0] && j[0].pixel_banner_proprio);
  ctx.waitUntil(cache.put(cacheKey, new Response(temProprio ? "0" : "1", { headers: { "Cache-Control": "max-age=600" } })));
  return !temProprio;
}

async function insertTrackEvent(env, row) {
  await fetch(`${REST}/track_events`, {
    method: "POST",
    headers: {
      apikey: env.SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
      "Content-Type": "application/json", Prefer: "return=minimal",
    },
    body: JSON.stringify(row),
  });
}

async function proxyToTracking(request, url) {
  const dest = TRACKING_FN + url.pathname + url.search;
  const headers = new Headers(request.headers);
  headers.delete("host");
  const init = {
    method: request.method,
    headers,
    body: (request.method === "GET" || request.method === "HEAD") ? undefined : await request.arrayBuffer(),
    redirect: "manual", // preserva os redirects 3xx dos links /l/... (destino final wa.me etc.)
  };
  return fetch(dest, init);
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return new Response("ok", { headers: cors });

    // sem o secret configurado ainda: comportamento antigo (proxy 1:1), nada quebra.
    if (!env.SUPABASE_SERVICE_ROLE_KEY) return proxyToTracking(request, url);

    // GET /pixel/script/<token>.js — servido direto aqui, fora da cota de Edge Function
    const mScript = url.pathname.match(/^\/pixel\/script\/([^/]+)\.js$/);
    if (mScript && request.method === "GET") {
      const cid = await resolveCid(env, mScript[1], ctx);
      const showBanner = cid ? await resolveShowBanner(env, cid, ctx) : true;
      const js = cid ? pixelScript(cid, url.origin, showBanner) : `console.warn("[Rastreamento] token inválido");`;
      return new Response(js, { headers: { ...cors, "Content-Type": "application/javascript; charset=utf-8", "Cache-Control": "public, max-age=3600" } });
    }

    // POST /collect — grava direto no Postgres via REST (Data API), fora da cota de Edge Function
    if (url.pathname === "/collect" && request.method === "POST") {
      let b = {};
      try { b = await request.json(); } catch (_e) { return new Response("bad", { status: 400, headers: cors }); }
      if (!b.cid || !b.type) return new Response("bad", { status: 400, headers: cors });
      const cid = await resolveCid(env, b.cid, ctx);
      if (!cid) return new Response("bad", { status: 400, headers: cors });
      const meta = (b.dest || b.campaignid || b.adgroupid || b.adid || b.keyword || b.matchtype || b.placement) ? {
        dest: b.dest ? clip(b.dest, 300) : undefined,
        campaignid: clip(b.campaignid, 40), adgroupid: clip(b.adgroupid, 40), adid: clip(b.adid, 40),
        keyword: clip(b.keyword, 200), matchtype: clip(b.matchtype, 20), placement: clip(b.placement, 60),
      } : null;
      ctx.waitUntil(insertTrackEvent(env, {
        id: uid(), client_id: cid, type: clip(b.type, 20), session_id: clip(b.sess, 40), anon_id: clip(b.anon, 40),
        utm_source: clip(b.utm_source, 120), utm_medium: clip(b.utm_medium, 120), utm_campaign: clip(b.utm_campaign, 200),
        utm_content: clip(b.utm_content, 200), utm_term: clip(b.utm_term, 200),
        fbclid: clip(b.fbclid, 300), gclid: clip(b.gclid, 300),
        referrer: clip(b.ref, 300), landing: clip(b.landing, 300), user_agent: clip(b.ua, 300),
        meta,
      }));
      return new Response("ok", { headers: { ...cors, "Content-Type": "text/plain" } });
    }

    // tudo o resto (links /l/..., /wa/ref, /wa/webhook etc. — baixo volume) continua na Edge Function
    return proxyToTracking(request, url);
  },
};
