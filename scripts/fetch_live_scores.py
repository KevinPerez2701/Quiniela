#!/usr/bin/env python3
"""
Auto-fills scores.json from football-data.org FINISHED World Cup matches.

Runs in CI on a schedule (see .github/workflows/auto-scores.yml). For every match
the API reports as FINISHED, it maps the official result onto the quiniela's
fixture match number and writes it into scores.json. Committing scores.json then
triggers the existing update-scores pipeline, which recalculates every player's
points exactly as if the admin had typed the result by hand.

Only FINISHED matches are written — never in-play scores — so the expensive
LibreOffice recalculation pipeline runs once per match, when it ends.

Policy — FILL_IF_EMPTY (default on):
  Only fills matches NOT already present in scores.json, so a manual edit made in
  the admin panel is never overwritten by the API. Set FILL_IF_EMPTY=0 to let the
  API's FINISHED result overwrite an existing value (API becomes source of truth).

Env vars:
  FOOTBALL_DATA_TOKEN  (required) football-data.org API token.
  FILL_IF_EMPTY        "0" to allow overwriting existing scores (default "1").
  DRY_RUN              "1" to print proposed changes without writing scores.json.

Team-name bridge: the fixture stores Spanish names, the API returns English ones,
so scripts/teams_es_en.json maps between them. Matches whose teams are still
placeholders (unresolved knockout slots) are skipped until both teams are real.
"""

import json
import os
import sys
import urllib.request
import urllib.error

ROOT        = os.path.join(os.path.dirname(__file__), "..")
DATA_PATH   = os.path.join(ROOT, "docs", "data.json")
SCORES_PATH = os.path.join(ROOT, "scores.json")
CONFIG_PATH = os.path.join(ROOT, "auto_config.json")
MAP_PATH    = os.path.join(os.path.dirname(__file__), "teams_es_en.json")

API_URL = "https://api.football-data.org/v4/competitions/WC/matches?status=FINISHED"

TOKEN         = os.environ.get("FOOTBALL_DATA_TOKEN", "").strip()
FILL_IF_EMPTY = os.environ.get("FILL_IF_EMPTY", "1") != "0"
DRY_RUN       = os.environ.get("DRY_RUN") == "1"


def fetch_finished():
    req = urllib.request.Request(API_URL, headers={"X-Auth-Token": TOKEN})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r).get("matches", [])


def auto_enabled():
    """Global kill-switch toggled from the admin panel (auto_config.json).
    Missing/invalid file is treated as enabled (the default)."""
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f).get("auto_enabled", True) is not False
    except (FileNotFoundError, ValueError):
        return True


def main():
    if not TOKEN:
        sys.exit("FOOTBALL_DATA_TOKEN is not set.")

    if not auto_enabled():
        print("Auto-update is OFF (auto_config.json) — admin is in manual mode. Skipping.")
        return

    es2en = json.load(open(MAP_PATH, encoding="utf-8"))
    en2es = {v: k for k, v in es2en.items()}

    fixture = json.load(open(DATA_PATH, encoding="utf-8")).get("fixture", [])
    # (home_es, away_es) -> match number. Group-stage pairings are unique and the
    # fixture uses the same home/away order as the API, so an exact ordered match
    # is safe. Placeholder knockout rows carry non-country names and never match.
    pair2num = {}
    for m in fixture:
        if m.get("home") and m.get("away") and m.get("num"):
            pair2num[(m["home"], m["away"])] = m["num"]

    scores = {}
    if os.path.exists(SCORES_PATH):
        with open(SCORES_PATH, encoding="utf-8") as f:
            scores = json.load(f)

    try:
        matches = fetch_finished()
    except urllib.error.HTTPError as e:
        sys.exit(f"football-data.org HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:200]}")

    changes = []
    skipped_unmapped = []
    for mt in matches:
        if mt.get("status") != "FINISHED":
            continue
        en_h = (mt.get("homeTeam") or {}).get("name")
        en_a = (mt.get("awayTeam") or {}).get("name")
        es_h, es_a = en2es.get(en_h), en2es.get(en_a)
        if not es_h or not es_a:
            skipped_unmapped.append(f"{en_h} vs {en_a}")
            continue
        num = pair2num.get((es_h, es_a))
        if not num:
            skipped_unmapped.append(f"{es_h} vs {es_a} (not in fixture)")
            continue
        ft = (mt.get("score") or {}).get("fullTime") or {}
        gh, ga = ft.get("home"), ft.get("away")
        if gh is None or ga is None:
            continue

        key = str(num)
        existing = scores.get(key)
        if FILL_IF_EMPTY and existing is not None:
            continue
        if existing and existing.get("goal_home") == gh and existing.get("goal_away") == ga:
            continue

        changes.append((int(num), es_h, es_a, int(gh), int(ga), existing))
        scores[key] = {"goal_home": int(gh), "goal_away": int(ga)}

    if skipped_unmapped:
        print(f"Skipped {len(skipped_unmapped)} unmapped/placeholder match(es): "
              f"{', '.join(skipped_unmapped[:8])}{' …' if len(skipped_unmapped) > 8 else ''}")

    if not changes:
        print("No new FINISHED results to apply — scores.json unchanged.")
        return

    print(f"{len(changes)} FINISHED match(es) to apply:")
    for num, h, a, gh, ga, existing in sorted(changes):
        was = f"  (was {existing['goal_home']}-{existing['goal_away']})" if existing else ""
        print(f"  #{num:>3}: {h} {gh}-{ga} {a}{was}")

    if DRY_RUN:
        print("DRY_RUN set — scores.json NOT written.")
        return

    ordered = dict(sorted(scores.items(), key=lambda kv: int(kv[0])))
    with open(SCORES_PATH, "w", encoding="utf-8") as f:
        json.dump(ordered, f, ensure_ascii=False, indent=2)
    print(f"scores.json updated ({len(ordered)} total match(es)).")


if __name__ == "__main__":
    main()
