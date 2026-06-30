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

// ── Reliable auto-scores trigger ──────────────────────────────────────────────
// GitHub throttles `*/5` scheduled workflows to ~30-90 min, so a finished match
// can wait too long to be applied. This Worker's Cron Trigger (set in Cloudflare
// to "*/5 * * * *") reliably pokes the auto-scores workflow via the GitHub API.
// Requires the Worker secret GH_DISPATCH_PAT (fine-grained PAT, Actions: Read &
// write on this repo). Set the Cron Trigger in: Worker > Settings > Triggers.
const GH_OWNER = "KevinPerez2701";
const GH_REPO = "Quiniela";
const GH_WORKFLOW = "auto-scores.yml";
const GH_REF = "main";

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
    const matches = (data.matches || []).map((m) => {
      const s = m.score || {};
      const reg = s.regularTime || {}, et = s.extraTime || {}, ft = s.fullTime || {}, pen = s.penalties || {};
      // Game result = regulation + extra time (NOT fullTime, which mixes in the
      // shootout on knockout matches). Plain matches only carry fullTime.
      let hs, as_;
      if (reg.home != null && reg.away != null) {
        hs = (reg.home || 0) + (et.home || 0);
        as_ = (reg.away || 0) + (et.away || 0);
      } else {
        hs = ft.home != null ? ft.home : null;
        as_ = ft.away != null ? ft.away : null;
      }
      // A finished shootout always has a winner, so equal penalties (e.g. 4-4)
      // are garbage from the simulated feed — don't surface them.
      const shootout = s.duration === "PENALTY_SHOOTOUT" && pen.home != null && pen.away != null && pen.home !== pen.away;
      return {
        utcDate: m.utcDate,
        status: m.status,
        home: m.homeTeam && m.homeTeam.name,
        away: m.awayTeam && m.awayTeam.name,
        hs,
        as: as_,
        ph: shootout ? pen.home : null,
        pa: shootout ? pen.away : null,
      };
    });

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

  // Cron Trigger: fires on the schedule configured in Cloudflare (set it to
  // "*/5 * * * *"). Pokes the GitHub auto-scores workflow so FINISHED results
  // get applied within minutes — but only when a match is actually live, about
  // to start, or recently finished (see maybeDispatchAutoScores), to avoid
  // dozens of no-op runs during the long stretches with no matches.
  async scheduled(event, env, ctx) {
    ctx.waitUntil(maybeDispatchAutoScores(env));
  },
};

// Gate: dispatch only when a result could actually land soon. Avoids no-op
// Actions runs without hardcoding match hours (which shift every day/phase).
// On a fetch error we dispatch anyway (fail-safe — never miss a result).
async function maybeDispatchAutoScores(env) {
  if (!env.GH_DISPATCH_PAT) {
    console.log("GH_DISPATCH_PAT not set — skipping auto-scores dispatch.");
    return;
  }
  let matches = null;
  try {
    matches = await getRawMatches(env);
  } catch (e) {
    console.log("match fetch failed; dispatching anyway:", e);
  }
  if (matches && !anyMatchActive(matches)) {
    console.log("No live/upcoming/recent WC match — skipping dispatch this tick.");
    return;
  }
  await dispatchAutoScores(env);
}

async function getRawMatches(env) {
  const r = await fetch(UPSTREAM, {
    headers: { "X-Auth-Token": env.FOOTBALL_DATA_TOKEN },
    cf: { cacheTtl: 30 },
  });
  if (!r.ok) throw new Error("upstream " + r.status);
  const data = await r.json();
  return data.matches || [];
}

function anyMatchActive(matches) {
  const now = Date.now();
  const MIN = 60 * 1000, HOUR = 3600 * 1000;
  for (const m of matches) {
    const s = m.status;
    if (s === "IN_PLAY" || s === "PAUSED" || s === "LIVE") return true;
    const ko = Date.parse(m.utcDate || "");
    if (isNaN(ko)) continue;
    if ((s === "TIMED" || s === "SCHEDULED") && ko - now > 0 && ko - now < 15 * MIN) return true;
    if (s === "FINISHED" && now - ko < 4 * HOUR) return true; // buffer for free-tier delay
  }
  return false;
}

async function dispatchAutoScores(env) {
  const url = `https://api.github.com/repos/${GH_OWNER}/${GH_REPO}/actions/workflows/${GH_WORKFLOW}/dispatches`;
  const r = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${env.GH_DISPATCH_PAT}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "quiniela-live-scores-worker",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ ref: GH_REF }),
  });
  if (!r.ok) {
    console.log(`auto-scores dispatch failed: ${r.status} ${await r.text()}`);
  }
}

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
