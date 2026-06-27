# Live-scores Worker (marcador en vivo)

Proxy mínimo en **Cloudflare Workers** (plan gratis) que sirve los marcadores del
Mundial desde football-data.org al dashboard, ocultando el token y resolviendo el
CORS. Es **solo visualización**: no toca `scores.json` ni los puntos. La
actualización de puntos la hace el GitHub Action `auto-scores.yml` (al `FINISHED`).

## Por qué hace falta

- football-data.org exige el token en el header `X-Auth-Token`. Llamarlo directo
  desde el navegador **expondría el token** y además el CORS de la API solo
  permite `http://localhost` (tu GitHub Pages quedaría bloqueado).
- El Worker guarda el token como *secret*, agrega CORS para tu sitio y **cachea**
  la respuesta ~45 s, de modo que muchos visitantes comparten **una** llamada y
  nunca se pasa del límite gratis (10 req/min).

## Despliegue (una vez)

Opción A — Dashboard (sin instalar nada):
1. Entra a <https://dash.cloudflare.com> → **Workers & Pages** → **Create** →
   **Create Worker**. Dale un nombre, p. ej. `wc-scores`. **Deploy**.
2. **Edit code** → pega el contenido de [`live-scores-worker.js`](./live-scores-worker.js) → **Deploy**.
3. **Settings → Variables and Secrets → Add** → tipo **Secret**,
   nombre `FOOTBALL_DATA_TOKEN`, valor = tu token de football-data.org → **Deploy**.
4. Copia la URL pública del Worker (`https://wc-scores.<tu-sub>.workers.dev`).

Opción B — Wrangler (CLI):
```bash
npm create cloudflare@latest wc-scores      # o: npx wrangler init wc-scores
# reemplaza src/index.js por live-scores-worker.js
npx wrangler secret put FOOTBALL_DATA_TOKEN # pega el token cuando lo pida
npx wrangler deploy
```

## Conectar con el dashboard

En [`docs/index.html`](../docs/index.html) pon la URL del Worker:
```js
const LIVE_SCORES_WORKER_URL = 'https://wc-scores.<tu-sub>.workers.dev';
```
Vacío (`''`) = función apagada (el sitio funciona igual, sin marcador en vivo).

## Probar

```bash
curl -s https://dark-tooth-2a90.kevinperez9a.workers.dev/ | head -c 


# { "resultSet": {...}, "matches": [ { "utcDate": "...", "status": "TIMED",
#   "home": "Panama", "away": "England", "hs": null, "as": null }, ... ] }
```
El header `X-Cache: HIT|MISS` indica si vino del caché del Worker.

## Nota sobre el plan gratis

En el plan **gratis** de football-data.org los marcadores en juego van con
**retraso** (no es tiempo real exacto). El *live* real arranca en el plan
*"Free w/ Livescores"* (~12 €/mes), **misma API y mismo código** — solo cambiarías
el token. Para los puntos (al `FINISHED`) el plan gratis es suficiente.
