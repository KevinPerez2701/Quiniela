#!/usr/bin/env python3
"""
Normaliza los picks de Bota/Balón (Oro/Plata/Bronce) de TODOS los jugadores
dentro de la hoja ADMIN de cada ADMINExcelMundial2026*.xlsx, llevándolos a la
forma canónica de scripts/player_aliases.json ("Mpappe" -> "Kylian Mbappé").

Los puntos de estos premios se calculan con COUNTIF(pick, ganador_real), que
ignora mayúsculas pero NO acentos ni variantes: sin normalizar, "Mbappé" no
matchea "Mbappe". Con los picks canónicos y el ganador real escrito también
canónico (ver update_scores.py --honor), la comparación siempre cierra.

Parchea quirúrgicamente el XML del xlsx (igual que update_scores.py): solo
reescribe las celdas de pick cuyo texto cambia (como inlineStr, conservando el
estilo) y activa fullCalcOnLoad. Nada más del archivo se toca.

Uso:
  python scripts/normalize_honor_picks.py            # aplica sobre el repo
  DRY_RUN=1 python scripts/normalize_honor_picks.py  # solo muestra cambios
  EXCEL_DIR=/otra/ruta ...                           # otra carpeta de xlsx
"""

import glob
import os
import re
import xml.etree.ElementTree as ET

from player_names import load_alias_map, normalize_player_name
from update_scores import cell_attrs, rewrite_zip, set_full_calc_on_load, M_NS, R_NS

import zipfile

REPO_ROOT = os.path.join(os.path.dirname(__file__), "..")
EXCEL_DIR = os.environ.get("EXCEL_DIR", REPO_ROOT)
DRY_RUN = os.environ.get("DRY_RUN") == "1"

# Rótulos (col K de ADMIN) de las 6 filas de premios individuales.
AWARD_LABEL_RE = re.compile(
    r"^(Bota|Bal[oó]n) de (Oro|Plata|Bronce)", re.IGNORECASE)

# Columnas de pick de jugador en ADMIN: 0-based 18, 21, 24, … (stride 3).
FIRST_PLAYER_COL = 19   # 1-based (S)
PLAYER_COL_STRIDE = 3
MAX_PLAYERS = 40


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def find_sheet_path(z, name):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    for sh in wb.find("m:sheets", M_NS):
        if sh.get("name") == name:
            return "xl/" + targets[sh.get(R_NS)].lstrip("/")
    raise KeyError(f"{name} sheet not found in workbook")


def load_shared_strings(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall("m:si", M_NS):
        out.append("".join(t.text or "" for t in si.iter(
            "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
    return out


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def xml_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))


def cell_text(cell_xml, shared):
    """Texto literal de una celda (t=s o inlineStr); None si es fórmula/número."""
    if "<f" in cell_xml:
        return None
    t_attr = re.search(r'\st="([^"]*)"', cell_xml)
    kind = t_attr.group(1) if t_attr else None
    if kind == "s":
        v = re.search(r"<v>(\d+)</v>", cell_xml)
        if v and int(v.group(1)) < len(shared):
            return shared[int(v.group(1))]
        return None
    if kind == "inlineStr":
        raw = "".join(re.findall(r"<t[^>]*>(.*?)</t>", cell_xml, re.S))
        return xml_unescape(raw) or None
    return None


def cell_display_text(cell_xml, shared):
    """Texto visible de una celda, incluyendo el caché de fórmulas t="str"."""
    t_attr = re.search(r'\st="([^"]*)"', cell_xml)
    kind = t_attr.group(1) if t_attr else None
    if kind == "str":
        v = re.search(r"<v[^>]*>(.*?)</v>", cell_xml, re.S)
        return xml_unescape(v.group(1)) if v else None
    return cell_text(cell_xml, shared)


def get_cell(xml_text, ref):
    m = re.search(r'<c r="%s"[^>]*?(?:/>|>.*?</c>)' % ref, xml_text, re.S)
    return m.group(0) if m else None


def normalize_admin_sheet(xml_text, shared, amap):
    """Normaliza los picks de premios; devuelve (xml, [(ref, antes, después)])."""
    # 1) Filas de premios: celdas K (rótulo con fórmula t="str" cacheada) cuyo
    #    texto empieza con Bota/Balón de …
    award_rows = []
    for m in re.finditer(r'<c r="K(\d+)"[^>]*?(?:/>|>.*?</c>)', xml_text, re.S):
        text = cell_display_text(m.group(0), shared)
        if text and AWARD_LABEL_RE.match(re.sub(r"\s+", " ", text).strip()):
            award_rows.append(m.group(1))
    if len(award_rows) != 6:
        raise RuntimeError(
            f"Se esperaban 6 filas de premios en ADMIN y se hallaron "
            f"{len(award_rows)}: {award_rows}")

    # 2) Celdas de pick por jugador en cada fila de premio
    changes = []
    for row in award_rows:
        for k in range(MAX_PLAYERS):
            ref = f"{col_letter(FIRST_PLAYER_COL + k * PLAYER_COL_STRIDE)}{row}"
            cell = get_cell(xml_text, ref)
            if not cell:
                continue
            text = cell_text(cell, shared)
            if not text:
                continue
            fixed = normalize_player_name(text, amap)
            if fixed == text:
                continue
            attrs = cell_attrs(cell)
            attrs = (" " + attrs) if attrs else ""
            new_cell = (f'<c r="{ref}"{attrs} t="inlineStr">'
                        f"<is><t>{xml_escape(fixed)}</t></is></c>")
            xml_text = xml_text.replace(cell, new_cell, 1)
            changes.append((ref, text, fixed))
    return xml_text, changes


def main():
    amap = load_alias_map()
    pattern = os.path.join(EXCEL_DIR, "ADMINExcelMundial2026*.xlsx")
    excel_files = sorted(glob.glob(pattern))
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found matching {pattern}")

    total = 0
    for path in excel_files:
        print(f"Normalizando {os.path.basename(path)} …")
        with zipfile.ZipFile(path) as z:
            sheet_path = find_sheet_path(z, "ADMIN")
            sheet_xml = z.read(sheet_path).decode("utf-8")
            shared = load_shared_strings(z)
            wb_xml = z.read("xl/workbook.xml").decode("utf-8")

        sheet_xml, changes = normalize_admin_sheet(sheet_xml, shared, amap)
        for ref, before, after in changes:
            print(f"  {ref}: {before!r} -> {after!r}")
        if not changes:
            print("  Sin cambios: todos los picks ya están normalizados.")
            continue
        total += len(changes)

        if DRY_RUN:
            print(f"  DRY_RUN — {len(changes)} celda(s) SIN escribir.")
            continue

        wb_xml = set_full_calc_on_load(wb_xml)
        rewrite_zip(path, {
            sheet_path: sheet_xml.encode("utf-8"),
            "xl/workbook.xml": wb_xml.encode("utf-8"),
        })
        print(f"  {len(changes)} celda(s) normalizada(s); fullCalcOnLoad set.")

    print(f"Listo — {total} pick(s) normalizado(s) en {len(excel_files)} archivo(s).")


if __name__ == "__main__":
    main()
