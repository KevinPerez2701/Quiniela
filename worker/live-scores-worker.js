/**
 * Cloudflare Worker — caching proxy for football-data.org World Cup scores.
 *
 * Why a Worker:
 *   - football-data.org requires the API token in an `X-Auth-Token` header.
 *     Calling it straight from the browser would expose the token AND is blocked
 *     by CORS (the API only allows `http://localhost` as origin). This Worker
 *     holds the token as a secret and adds permissive CORS for your Pages site.
 *   - It caches the upstream response ~45s so many viewers share ONE upstream
 *     call, staying well under the free tier's 10 requests/minute.
 *
 * Setup:
 *   1. Create the Worker (dashboard or `wrangler init`), paste this code.
 *   2. Add the secret:  wrangler secret put FOOTBALL_DATA_TOKEN
 *      (or Dashboard > Worker > Settings > Variables and Secrets > add Secret).
 *   3. Deploy. Copy the https://<name>.<sub>.workers.dev URL into
 *      docs/index.html  ->  const LIVE_SCORES_WORKER_URL.
 *
 * It returns a slim JSON: { resultSet, matches: [{ utcDate, status, home, away, hs, as }] }
 */

const UPSTREAM = "https://api.football-data.org/v4/competitions/WC/matches";
const CACHE_SECONDS = 45;

const CORS = {
  "Access-Control-Allow-Origin": "*", // tighten to your Pages origin if you prefer
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

export default {
  async fetch(request, env, ctx) {
    if (request.method === "OPTIONS") return new Response(null, { headers: CORS });
    if (request.method !== "GET") {
      return jsonResponse({ error: "method_not_allowed" }, 405);
    }
    if (!env.FOOTBALL_DATA_TOKEN) {
      return jsonResponse({ error: "missing_secret", hint: "set FOOTBALL_DATA_TOKEN" }, 500);
    }

    const cache = caches.default;
    const cacheKey = new Request("https://wc-scores-cache/all", { method: "GET" });

    const cached = await cache.match(cacheKey);
    if (cached) return withCors(cached, "HIT");

    let upstream;
    try {
      upstream = await fetch(UPSTREAM, {
        headers: { "X-Auth-Token": env.FOOTBALL_DATA_TOKEN },
        cf: { cacheTtl: 0 },
      });
    } catch (e) {
      return jsonResponse({ error: "upstream_fetch_failed" }, 502);
    }
    if (!upstream.ok) {
      return jsonResponse({ error: "upstream_status", status: upstream.status }, 502);
    }

    const data = await upstream.json();
    const matches = (data.matches || []).map((m) => ({
      utcDate: m.utcDate,
      status: m.status,
      home: m.homeTeam && m.homeTeam.name,
      away: m.awayTeam && m.awayTeam.name,
      hs: m.score && m.score.fullTime ? m.score.fullTime.home : null,
      as: m.score && m.score.fullTime ? m.score.fullTime.away : null,
    }));

    const body = JSON.stringify({ resultSet: data.resultSet || null, matches });
    const resp = new Response(body, {
      headers: {
        ...CORS,
        "Content-Type": "application/json",
        "Cache-Control": `s-maxage=${CACHE_SECONDS}`,
        "X-Cache": "MISS",
      },
    });
    ctx.waitUntil(cache.put(cacheKey, resp.clone()));
    return resp;
  },
};

function withCors(response, cacheState) {
  const r = new Response(response.body, response);
  for (const [k, v] of Object.entries(CORS)) r.headers.set(k, v);
  if (cacheState) r.headers.set("X-Cache", cacheState);
  return r;
}

function jsonResponse(obj, status) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...CORS, "Content-Type": "application/json" },
  });
}
