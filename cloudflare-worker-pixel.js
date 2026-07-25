// Cloudflare Worker — serve o pixel/links de rastreio no domínio próprio (pixel.gt-marketing.app.br)
// repassando pra Edge Function "tracking" do Supabase. Deploy: Cloudflare → Workers & Pages → Create Worker
// → cole este código → Deploy → no Worker, Settings → Domains & Routes → Add Custom Domain: pixel.gt-marketing.app.br
export default {
  async fetch(request) {
    const url = new URL(request.url);
    const TARGET = "https://mocrfqmdjwvyhqvdpimm.supabase.co/functions/v1/tracking";
    const dest = TARGET + url.pathname + url.search;
    const headers = new Headers(request.headers);
    headers.delete("host"); // deixa o fetch definir o Host correto do Supabase
    const init = {
      method: request.method,
      headers,
      body: (request.method === "GET" || request.method === "HEAD") ? undefined : await request.arrayBuffer(),
      redirect: "manual", // preserva os redirects 3xx dos links /l/... (destino final wa.me etc.)
    };
    return fetch(dest, init);
  },
};
