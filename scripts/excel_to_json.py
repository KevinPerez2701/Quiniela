#!/usr/bin/env python3
"""Reads the CLAS sheet from the Excel quiniela file and writes docs/data.json."""

import json
import os
import warnings
from datetime import datetime, timezone, timedelta

# Suppress openpyxl extension warnings
warnings.filterwarnings("ignore", category=UserWarning)

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "..", "ADMINExcelMundial2026.xlsx")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "docs", "data.json")

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


def to_num(val):
    if val is None or val == "-":
        return 0
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0


def main():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["CLAS"]

    rows = list(ws.iter_rows(values_only=True))

    # Row 0 (index): title
    title = rows[0][1] or "CLASIFICACIÓN MUNDIAL 2026"

    # Rows 4+ (index 4..): player data until the filler row
    keys = [col[0] for col in COLUMNS]
    players = []
    for row in rows[4:]:
        if row[0] is None or row[0] == 25 and row[2] == "-":
            continue
        # col B=index1 → pos, col C=index2 → jugador, col D=index3 → puntos ...
        # Actual layout from inspection: col[0]=seq, col[1]=pos, col[2]=jugador, col[3..19]=scores
        if not isinstance(row[2], str):
            continue
        player = {
            "pos":         to_num(row[1]),
            "jugador":     row[2],
            "puntos":      to_num(row[3]),
            "f_grupos":    to_num(row[4]),
            "pos_grupos":  to_num(row[5]),
            "eq_16":       to_num(row[6]),
            "pt_16":       to_num(row[7]),
            "eq_8":        to_num(row[8]),
            "pt_8":        to_num(row[9]),
            "eq_4":        to_num(row[10]),
            "pt_4":        to_num(row[11]),
            "eq_2":        to_num(row[12]),
            "pt_2":        to_num(row[13]),
            "eq_34":       to_num(row[14]),
            "eq_final":    to_num(row[15]),
            "pt_34":       to_num(row[16]),
            "pt_final":    to_num(row[17]),
            "honor":       to_num(row[18]),
        }
        players.append(player)

    vet = timezone(timedelta(hours=-4))
    now_vet = datetime.now(vet)

    output = {
        "title":      title,
        "updated_at": now_vet.strftime("%d/%m/%Y %H:%M (hora Venezuela)"),
        "max_values": MAX_VALUES,
        "columns":    COLUMNS,
        "players":    players,
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"Wrote {len(players)} players to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
