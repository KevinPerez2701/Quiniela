#!/usr/bin/env python3
"""
Reads CLAS, Fixture, DailyPrediction and DailyClas from all Excel files
matching ADMINExcelMundial2026*.xlsx and writes a merged docs/data.json.

To support more than 25 players, simply add a second file named e.g.
  ADMINExcelMundial2026_2.xlsx
The script will discover it automatically, merge all players, and
re-rank the combined standings.
"""

import glob
import json
import os
import warnings
from datetime import datetime, timezone, timedelta

warnings.filterwarnings("ignore", category=UserWarning)
import openpyxl

REPO_ROOT   = os.path.join(os.path.dirname(__file__), "..")
OUTPUT_PATH = os.path.join(REPO_ROOT, "docs", "data.json")

MAX_VALUES = {
    "f_grupos":   432,
    "pos_grupos": 156,
    "eq_16":       96,
    "pt_16":      144,
    "eq_8":        64,
    "pt_8":        96,
    "eq_4":        40,
    "pt_4":        60,
    "eq_2":        28,
    "pt_2":        36,
    "eq_34":       10,
    "eq_final":    20,
    "pt_34":       18,
    "pt_final":    21,
    "honor":       71,
}

COLUMNS = [
    ("pos",         "Pos"),
    ("jugador",     "Jugador"),
    ("puntos",      "Puntos Totales"),
    ("f_grupos",    "F. Grupos"),
    ("pos_grupos",  "Pos. Grupos"),
    ("eq_16",       "Equipos 1/16"),
    ("pt_16",       "Partidos 1/16"),
    ("eq_8",        "Equipos 1/8"),
    ("pt_8",        "Partidos 1/8"),
    ("eq_4",        "Equipos 1/4"),
    ("pt_4",        "Partidos 1/4"),
    ("eq_2",        "Equipos 1/2"),
    ("pt_2",        "Partidos 1/2"),
    ("eq_34",       "Equipos 3-4"),
    ("eq_final",    "Equipos Final"),
    ("pt_34",       "Partido 3-4"),
    ("pt_final",    "Partido Final"),
    ("honor",       "Cuadro de Honor"),
]

PHASE_MAP = {
    range(1, 73):    "grupos",
    range(73, 89):   "dieciseisavos",
    range(89, 97):   "octavos",
    range(97, 101):  "cuartos",
    range(101, 103): "semis",
    range(103, 104): "tercero",
    range(104, 105): "final",
}

PHASE_LABELS = {
    "grupos":        "Fase de Grupos",
    "dieciseisavos": "1/16 de Final",
    "octavos":       "1/8 de Final",
    "cuartos":       "Cuartos de Final",
    "semis":         "Semifinales",
    "tercero":       "Tercer Puesto",
    "final":         "Final",
}


def to_num(val):
    if val is None or val == "-":
        return 0
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0


def fmt_date(dt):
    if dt is None:
        return None
    if hasattr(dt, "strftime"):
        return dt.strftime("%Y-%m-%d %H:%M")
    return str(dt)


def get_phase(num):
    if not isinstance(num, int):
        return "final"
    for rng, phase in PHASE_MAP.items():
        if num in rng:
            return phase
    return "final"


def assign_positions(players, sort_key="puntos"):
    """Sort players by sort_key descending and assign positions (ties share the same rank)."""
    players.sort(key=lambda p: p.get(sort_key, 0), reverse=True)
    pos = 1
    for i, p in enumerate(players):
        if i > 0 and p.get(sort_key, 0) < players[i - 1].get(sort_key, 0):
            pos = i + 1
        p["pos"] = pos
    return players


# ── CLAS ─────────────────────────────────────────────────────────────────────

def extract_clas_players(ws):
    """Return (title, [players]) from a CLAS worksheet."""
    rows = list(ws.iter_rows(values_only=True))
    title = rows[0][1] or "CLASIFICACIÓN MUNDIAL 2026"
    players = []
    for row in rows[4:]:
        if not isinstance(row[2], str) or row[2] == "-":
            continue
        players.append({
            "jugador":    row[2],
            "puntos":     to_num(row[3]),
            "f_grupos":   to_num(row[4]),
            "pos_grupos": to_num(row[5]),
            "eq_16":      to_num(row[6]),
            "pt_16":      to_num(row[7]),
            "eq_8":       to_num(row[8]),
            "pt_8":       to_num(row[9]),
            "eq_4":       to_num(row[10]),
            "pt_4":       to_num(row[11]),
            "eq_2":       to_num(row[12]),
            "pt_2":       to_num(row[13]),
            "eq_34":      to_num(row[14]),
            "eq_final":   to_num(row[15]),
            "pt_34":      to_num(row[16]),
            "pt_final":   to_num(row[17]),
            "honor":      to_num(row[18]),
        })
    return title, players


# ── FIXTURE ───────────────────────────────────────────────────────────────────

def extract_fixture(ws):
    matches = []
    current_group = None
    for row in ws.iter_rows(values_only=True):
        grp = row[35]
        if grp and isinstance(grp, str) and grp.startswith("Grupo"):
            current_group = grp

        date = row[23]
        home = row[26]
        away = row[31]
        num  = row[33]
        if not (date and hasattr(date, "strftime") and home and away):
            continue

        phase  = get_phase(num)
        goal_h = row[28]
        goal_a = row[29]
        matches.append({
            "num":         num,
            "date":        fmt_date(date),
            "matchday":    str(row[25]) if row[25] else None,
            "group":       current_group if phase == "grupos" else None,
            "home":        str(home),
            "away":        str(away),
            "goal_home":   int(goal_h) if isinstance(goal_h, (int, float)) else None,
            "goal_away":   int(goal_a) if isinstance(goal_a, (int, float)) else None,
            "phase":       phase,
            "phase_label": PHASE_LABELS.get(phase, phase),
        })

    matches.sort(key=lambda m: (m["date"] or "", m["num"] or 999))
    return matches


# ── DAILY PREDICTION ──────────────────────────────────────────────────────────

def extract_daily_prediction_players(ws):
    """Return (fecha, matches, [players]) from a DailyPrediction worksheet."""
    rows   = list(ws.iter_rows(values_only=True))
    day_raw = rows[0][7]
    day_str = fmt_date(day_raw).split(" ")[0] if day_raw else None
    matches = [str(v).strip().replace("\n", "") for v in rows[1] if v and str(v).strip()]
    players = []
    for row in rows[3:]:
        name = row[5]
        if not (name and isinstance(name, str)):
            continue
        preds = []
        for v in row[7:7 + len(matches)]:
            preds.append(None if (v is None or v == "-") else str(v).strip())
        players.append({"jugador": name, "predicciones": preds})
    return day_str, matches, players


# ── DAILY CLAS ────────────────────────────────────────────────────────────────

def extract_daily_clas_players(ws):
    """Return (fecha, matches, [players]) from a DailyClas worksheet."""
    rows    = list(ws.iter_rows(values_only=True))
    day_raw = rows[0][8]
    day_str = fmt_date(day_raw).split(" ")[0] if day_raw else None
    matches = [str(v).strip().replace("\n", "") for v in rows[1] if v and str(v).strip()]
    players = []
    for row in rows[3:]:
        name = row[5]
        if not (name and isinstance(name, str)):
            continue
        total = row[6]
        pts_partidos = []
        for v in row[7:7 + len(matches)]:
            if v is None or v == "-":
                pts_partidos.append(None)
            else:
                try:
                    pts_partidos.append(int(v))
                except (ValueError, TypeError):
                    pts_partidos.append(0)
        players.append({
            "jugador":         name,
            "puntos_dia":      to_num(total),
            "puntos_partidos": pts_partidos,
        })
    return day_str, matches, players


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    # Discover all Excel files (sorted so primary always comes first)
    pattern = os.path.join(REPO_ROOT, "ADMINExcelMundial2026*.xlsx")
    excel_files = sorted(glob.glob(pattern))

    if not excel_files:
        raise FileNotFoundError(f"No Excel files found matching {pattern}")

    # First file alphabetically is the primary (provides Fixture + match columns)
    primary_file = excel_files[0]
    print(f"Found {len(excel_files)} Excel file(s): {[os.path.basename(f) for f in excel_files]}")
    print(f"Primary file: {os.path.basename(primary_file)}")

    vet     = timezone(timedelta(hours=-4))
    now_vet = datetime.now(vet)

    # ── Merge CLAS, DailyPrediction and DailyClas from all files ─────────────
    all_clas_players       = []
    all_daily_pred_players = []
    all_daily_clas_players = []
    title = "CLASIFICACIÓN MUNDIAL 2026"
    fecha_pred = fecha_clas = None
    partidos_pred = partidos_clas = []

    for path in excel_files:
        wb = openpyxl.load_workbook(path, data_only=True)
        is_primary = (os.path.abspath(path) == os.path.abspath(primary_file))

        t, players = extract_clas_players(wb["CLAS"])
        if is_primary:
            title = t
        all_clas_players.extend(players)

        fp, mp, pp = extract_daily_prediction_players(wb["DailyPrediction"])
        if is_primary or not fecha_pred:
            fecha_pred   = fp
            partidos_pred = mp
        all_daily_pred_players.extend(pp)

        fc, mc, pc = extract_daily_clas_players(wb["DailyClas"])
        if is_primary or not fecha_clas:
            fecha_clas   = fc
            partidos_clas = mc
        all_daily_clas_players.extend(pc)

        # Fixture only from primary
        if is_primary:
            fixture = extract_fixture(wb["WORLDCUP"])

    # ── Re-rank merged standings ──────────────────────────────────────────────
    clas_players      = assign_positions(all_clas_players, "puntos")
    daily_clas_players = assign_positions(all_daily_clas_players, "puntos_dia")

    output = {
        "updated_at": now_vet.strftime("%d/%m/%Y %H:%M (hora Venezuela)"),
        "clas": {"title": title, "players": clas_players},
        "max_values":  MAX_VALUES,
        "columns":     COLUMNS,
        "fixture":     fixture,
        "daily_pred":  {"fecha": fecha_pred, "partidos": partidos_pred, "jugadores": all_daily_pred_players},
        "daily_clas":  {"fecha": fecha_clas, "partidos": partidos_clas, "jugadores": daily_clas_players},
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"Wrote {len(clas_players)} players, {len(fixture)} matches → {OUTPUT_PATH}")


if __name__ == "__main__":
    main()


import json
import os
import warnings
from datetime import datetime, timezone, timedelta

warnings.filterwarnings("ignore", category=UserWarning)
import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "..", "ADMINExcelMundial2026.xlsx")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "docs", "data.json")

MAX_VALUES = {
    "f_grupos":   432,
    "pos_grupos": 156,
    "eq_16":       96,
    "pt_16":      144,
    "eq_8":        64,
    "pt_8":        96,
    "eq_4":        40,
    "pt_4":        60,
    "eq_2":        28,
    "pt_2":        36,
    "eq_34":       10,
    "eq_final":    20,
    "pt_34":       18,
    "pt_final":    21,
    "honor":       71,
}

COLUMNS = [
    ("pos",         "Pos"),
    ("jugador",     "Jugador"),
    ("puntos",      "Puntos Totales"),
    ("f_grupos",    "F. Grupos"),
    ("pos_grupos",  "Pos. Grupos"),
    ("eq_16",       "Equipos 1/16"),
    ("pt_16",       "Partidos 1/16"),
    ("eq_8",        "Equipos 1/8"),
    ("pt_8",        "Partidos 1/8"),
    ("eq_4",        "Equipos 1/4"),
    ("pt_4",        "Partidos 1/4"),
    ("eq_2",        "Equipos 1/2"),
    ("pt_2",        "Partidos 1/2"),
    ("eq_34",       "Equipos 3-4"),
    ("eq_final",    "Equipos Final"),
    ("pt_34",       "Partido 3-4"),
    ("pt_final",    "Partido Final"),
    ("honor",       "Cuadro de Honor"),
]

PHASE_MAP = {
    range(1, 73):   "grupos",
    range(73, 89):  "dieciseisavos",
    range(89, 97):  "octavos",
    range(97, 101): "cuartos",
    range(101, 103):"semis",
    range(103, 104):"tercero",
    range(104, 105):"final",
}

PHASE_LABELS = {
    "grupos":          "Fase de Grupos",
    "dieciseisavos":   "1/16 de Final",
    "octavos":         "1/8 de Final (Cuartos)",
    "cuartos":         "Cuartos de Final",
    "semis":           "Semifinales",
    "tercero":         "Tercer Puesto",
    "final":           "Final",
}


def to_num(val):
    if val is None or val == "-":
        return 0
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0


def fmt_date(dt):
    if dt is None:
        return None
    if hasattr(dt, "strftime"):
        return dt.strftime("%Y-%m-%d %H:%M")
    return str(dt)


def get_phase(num):
    if not isinstance(num, int):
        return "final"
    for rng, phase in PHASE_MAP.items():
        if num in rng:
            return phase
    return "final"


# ── CLAS ────────────────────────────────────────────────────────────────────

def extract_clas(ws):
    rows = list(ws.iter_rows(values_only=True))
    title = rows[0][1] or "CLASIFICACIÓN MUNDIAL 2026"
    players = []
    for row in rows[4:]:
        if not isinstance(row[2], str) or row[2] == "-":
            continue
        players.append({
            "pos":        to_num(row[1]),
            "jugador":    row[2],
            "puntos":     to_num(row[3]),
            "f_grupos":   to_num(row[4]),
            "pos_grupos": to_num(row[5]),
            "eq_16":      to_num(row[6]),
            "pt_16":      to_num(row[7]),
            "eq_8":       to_num(row[8]),
            "pt_8":       to_num(row[9]),
            "eq_4":       to_num(row[10]),
            "pt_4":       to_num(row[11]),
            "eq_2":       to_num(row[12]),
            "pt_2":       to_num(row[13]),
            "eq_34":      to_num(row[14]),
            "eq_final":   to_num(row[15]),
            "pt_34":      to_num(row[16]),
            "pt_final":   to_num(row[17]),
            "honor":      to_num(row[18]),
        })
    return {"title": title, "players": players}


# ── FIXTURE ──────────────────────────────────────────────────────────────────

def extract_fixture(ws):
    matches = []
    current_group = None
    for row in ws.iter_rows(values_only=True):
        date = row[23]
        home = row[26]
        away = row[31]
        num  = row[33]
        grp  = row[35]

        if grp and isinstance(grp, str) and grp.startswith("Grupo"):
            current_group = grp

        if not (date and hasattr(date, "strftime") and home and away):
            continue

        phase = get_phase(num)
        goal_h = row[28]
        goal_a = row[29]

        matches.append({
            "num":       num,
            "date":      fmt_date(date),
            "matchday":  str(row[25]) if row[25] else None,
            "group":     current_group if phase == "grupos" else None,
            "home":      str(home),
            "away":      str(away),
            "goal_home": int(goal_h) if isinstance(goal_h, (int, float)) else None,
            "goal_away": int(goal_a) if isinstance(goal_a, (int, float)) else None,
            "phase":     phase,
            "phase_label": PHASE_LABELS.get(phase, phase),
        })

    matches.sort(key=lambda m: (m["date"] or "", m["num"] or 999))
    return matches


# ── DAILY PREDICTION ─────────────────────────────────────────────────────────

def extract_daily_prediction(ws):
    rows = list(ws.iter_rows(values_only=True))
    day_raw = rows[0][7]
    day_str = fmt_date(day_raw).split(" ")[0] if day_raw else None
    matches = [str(v).strip().replace("\n", "") for v in rows[1] if v and str(v).strip()]

    players = []
    for row in rows[3:]:
        name = row[5]
        if not (name and isinstance(name, str)):
            continue
        preds = []
        for v in row[7:7 + len(matches)]:
            if v is None or v == "-":
                preds.append(None)
            else:
                preds.append(str(v).strip())
        players.append({"jugador": name, "predicciones": preds})

    return {"fecha": day_str, "partidos": matches, "jugadores": players}


# ── DAILY CLAS ───────────────────────────────────────────────────────────────

def extract_daily_clas(ws):
    rows = list(ws.iter_rows(values_only=True))
    day_raw = rows[0][8]
    day_str = fmt_date(day_raw).split(" ")[0] if day_raw else None
    matches = [str(v).strip().replace("\n", "") for v in rows[1] if v and str(v).strip()]

    players = []
    for row in rows[3:]:
        name = row[5]
        if not (name and isinstance(name, str)):
            continue
        total = row[6]
        pts_partidos = []
        for v in row[7:7 + len(matches)]:
            if v is None or v == "-":
                pts_partidos.append(None)
            else:
                try:
                    pts_partidos.append(int(v))
                except (ValueError, TypeError):
                    pts_partidos.append(0)
        players.append({
            "pos":            to_num(row[4]),
            "jugador":        name,
            "puntos_dia":     to_num(total),
            "puntos_partidos": pts_partidos,
        })

    return {"fecha": day_str, "partidos": matches, "jugadores": players}


# ── STATS ─────────────────────────────────────────────────────────────────────

def extract_stats(ws):
    rows = list(ws.iter_rows(values_only=True))
    day_raw = rows[4][4]
    day_str = fmt_date(day_raw).split(" ")[0] if day_raw else None

    stat_indices = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
    stats = []
    for i in stat_indices:
        if i >= len(rows):
            continue
        row = rows[i]
        label = row[3]
        val   = row[4]
        pct   = row[5]
        if not (label and str(label).strip()):
            continue
        # Clean up date-embedded labels (e.g. "... 11/06/26 ...")
        label_clean = str(label).strip()
        val_clean = val if isinstance(val, (int, float)) else (0 if not val or str(val).startswith("#") else str(val))
        pct_clean = round(float(pct) * 100, 1) if isinstance(pct, float) else 0
        stats.append({"label": label_clean, "value": val_clean, "pct": pct_clean})

    return {"fecha": day_str, "stats": stats}


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    vet = timezone(timedelta(hours=-4))
    now_vet = datetime.now(vet)

    clas      = extract_clas(wb["CLAS"])
    fixture   = extract_fixture(wb["WORLDCUP"])
    daily_pred = extract_daily_prediction(wb["DailyPrediction"])
    daily_clas = extract_daily_clas(wb["DailyClas"])
    stats     = extract_stats(wb["Stats"])

    output = {
        "updated_at": now_vet.strftime("%d/%m/%Y %H:%M (hora Venezuela)"),
        "clas":        clas,
        "max_values":  MAX_VALUES,
        "columns":     COLUMNS,
        "fixture":     fixture,
        "daily_pred":  daily_pred,
        "daily_clas":  daily_clas,
        "stats":       stats,
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"Wrote {len(clas['players'])} players, {len(fixture)} matches → {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
