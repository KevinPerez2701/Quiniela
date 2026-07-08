#!/usr/bin/env python3
"""
Normalización de nombres de futbolistas del Cuadro de Honor (botas y balones).

Los picks de los jugadores son texto libre pegado en el Excel ("Mpappe",
"mbappe", "Francia - Kylian Mpappe", "Mbappe (Francia)"…) y el Excel los
compara contra el ganador real con COUNTIF, que ignora mayúsculas pero NO
acentos ni variantes. Este módulo lleva cualquier variante a una forma
canónica única (scripts/player_aliases.json) para que la comparación cierre.

La clave de comparación (canon_key) quita paréntesis "(País)", prefijos
"País - ", espacios repetidos, mayúsculas y acentos. Un valor sin alias
conocido pasa sin cambios (solo recortado).
"""

import json
import os
import re
import unicodedata

ALIASES_PATH = os.path.join(os.path.dirname(__file__), "player_aliases.json")


def canon_key(s):
    s = str(s)
    s = re.sub(r"\([^)]*\)", " ", s)      # "Mbappe (Francia)" -> "Mbappe"
    s = s.split(" - ")[-1]                 # "Francia - Kylian Mpappe" -> "Kylian Mpappe"
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return s


def load_alias_map(path=ALIASES_PATH):
    """clave canónica normalizada -> nombre canónico oficial."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    amap = {}
    for canonical, aliases in data.items():
        amap[canon_key(canonical)] = canonical
        for a in aliases:
            amap[canon_key(a)] = canonical
    return amap


def normalize_player_name(value, amap):
    """Devuelve el nombre canónico si hay alias; si no, el valor recortado."""
    if value is None:
        return value
    text = str(value).strip()
    if not text or text == "-":
        return value
    return amap.get(canon_key(text), text)
