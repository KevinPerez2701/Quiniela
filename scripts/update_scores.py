#!/usr/bin/env python3
"""
Writes the scores from scores.json into the WORLDCUP sheet of every
ADMINExcelMundial2026*.xlsx file by surgically patching the sheet XML
inside the xlsx (a zip). All files get the scores because each one
computes its own players' points from its own WORLDCUP sheet.

Unlike openpyxl's save (which rewrites the whole workbook and corrupts
charts/dropdowns/formula caches), this only edits the goal cells' XML and
leaves every other byte of the file untouched — equivalent to a human
typing the scores in Excel and saving. It also sets fullCalcOnLoad so
Excel/LibreOffice recalculates every formula the next time the file opens.

scores.json format:
  { "<match_num>": { "goal_home": <int|null>, "goal_away": <int|null> }, ... }

It may also carry the actual Cuadro de Honor award winners (written from the
admin panel). They go into the WORLDCUP honor-box input cells; names are
normalized via scripts/player_aliases.json so they match the players' picks
(the sheet compares them with COUNTIF, which is accent-sensitive):
  "honor": { "bota_oro": "Kylian Mbappé", "bota_plata": …, "bota_bronce": …,
             "balon_oro": …, "balon_plata": …, "balon_bronce": … }
Empty/missing honor fields are left untouched (removing a field from
scores.json does NOT undo the Excel cell — to void an already-written award,
save "-" for that field: it matches nobody in the COUNTIF).
Campeón/Subcampeón/3º puesto need no input: the sheet resolves them from the
final/3rd-place scores.
"""

import glob
import json
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

from player_names import load_alias_map, normalize_player_name

REPO_ROOT   = os.path.join(os.path.dirname(__file__), "..")
SCORES_PATH = os.path.join(REPO_ROOT, "scores.json")

# WORLDCUP sheet columns (Excel letters): match number, home goals, away goals.
COL_NUM, COL_GH, COL_GA = "AH", "AC", "AD"
# Penalty-shootout cells (knockout rows only): home penalties, away penalties.
# These are the two "X" cells the sheet enables for decisive phases.
COL_PH, COL_PA = "AB", "AE"

# WORLDCUP honor-box input cells ("Escribe el jugador Bota de Oro", …).
HONOR_CELLS = {
    "bota_oro":     "AA154",
    "bota_plata":   "AA155",
    "bota_bronce":  "AA156",
    "balon_oro":    "AA158",
    "balon_plata":  "AA159",
    "balon_bronce": "AA160",
}

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
M_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def find_worldcup_sheet_path(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    for sh in wb.find("m:sheets", M_NS):
        if sh.get("name") == "WORLDCUP":
            t = targets[sh.get(R_NS)]
            return "xl/" + t.lstrip("/")
    raise KeyError("WORLDCUP sheet not found in workbook")


def cell_attrs(cell_xml):
    """Extract the attributes of a <c> tag, dropping r= (re-added by caller)
    and t= (we write plain numbers)."""
    m = re.match(r'<c\s+([^>]*?)\s*/?>', cell_xml)
    attrs = " " + (m.group(1) if m else "")
    attrs = re.sub(r'\s[rt]="[^"]*"', "", attrs).strip()
    return attrs


def patch_sheet_xml(xml_text, scores):
    """Set goal values for each match row; return (new_xml, updated_count)."""
    # Map Excel row -> match number from the AH column cells.
    row_for_num = {}
    for m in re.finditer(r'<c r="%s(\d+)"[^>]*?(?:/>|>(.*?)</c>)' % COL_NUM, xml_text, re.S):
        body = m.group(2) or ""
        v = re.search(r"<v>([\d.]+)</v>", body)
        if v:
            row_for_num[str(int(float(v.group(1))))] = m.group(1)

    updated = 0
    for num, entry in scores.items():
        row = row_for_num.get(str(num))
        if row is None:
            print(f"  ! match {num} not found in WORLDCUP — skipped")
            continue
        cells = [(COL_GH, entry.get("goal_home")), (COL_GA, entry.get("goal_away"))]
        # Only touch the penalty cells when this entry actually carries penalties,
        # so a knockout result entered by hand in the Excel is never wiped.
        if "pen_home" in entry or "pen_away" in entry:
            cells += [(COL_PH, entry.get("pen_home")), (COL_PA, entry.get("pen_away"))]
        for col, goal in cells:
            ref = f"{col}{row}"
            pat = re.compile(r'<c r="%s"[^>]*?(?:/>|>.*?</c>)' % ref, re.S)
            m = pat.search(xml_text)
            if not m:
                print(f"  ! cell {ref} not found — skipped")
                continue
            attrs = cell_attrs(m.group(0))
            attrs = (" " + attrs) if attrs else ""
            if goal is None:
                new_cell = f'<c r="{ref}"{attrs}/>'
            else:
                new_cell = f'<c r="{ref}"{attrs}><v>{int(goal)}</v></c>'
            xml_text = pat.sub(new_cell, xml_text, count=1)
        updated += 1
    return xml_text, updated


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def patch_honor_cells(xml_text, honor, amap):
    """Escribe los ganadores reales de botas/balones (nombres canónicos) en las
    celdas del honor box, reemplazando la fórmula placeholder por texto literal
    (equivalente a escribirlo a mano en Excel). Campos vacíos no se tocan."""
    from player_names import canon_key
    written = 0
    for key, ref in HONOR_CELLS.items():
        value = honor.get(key)
        if value is None or not str(value).strip():
            continue
        name = normalize_player_name(str(value), amap)
        if name not in ("-",) and canon_key(name) not in amap:
            print(f"  ! '{name}' no está en player_aliases.json — se escribe tal cual "
                  f"(revisa que coincida con los picks o agrégalo como alias)")
        pat = re.compile(r'<c r="%s"[^>]*?(?:/>|>.*?</c>)' % ref, re.S)
        m = pat.search(xml_text)
        if not m:
            print(f"  ! honor cell {ref} ({key}) not found — skipped")
            continue
        attrs = cell_attrs(m.group(0))
        attrs = (" " + attrs) if attrs else ""
        new_cell = (f'<c r="{ref}"{attrs} t="inlineStr">'
                    f"<is><t>{xml_escape(name)}</t></is></c>")
        xml_text = pat.sub(lambda _m: new_cell, xml_text, count=1)
        print(f"  {key} -> {ref}: {name!r}")
        written += 1
    return xml_text, written


def set_full_calc_on_load(workbook_xml):
    if "fullCalcOnLoad" in workbook_xml:
        return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', workbook_xml)
    if "<calcPr" in workbook_xml:
        return re.sub(r"<calcPr\s", '<calcPr fullCalcOnLoad="1" ', workbook_xml, count=1)
    return workbook_xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')


def rewrite_zip(path, replacements, removals=frozenset()):
    """Rewrite the zip at path, swapping only the entries in replacements
    (name -> bytes) and dropping the ones in removals; every other entry is
    copied through byte-for-byte."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in removals:
                continue
            data = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    shutil.move(tmp, path)


def drop_calc_chain(z, replacements, removals):
    """Al escribir un premio se ELIMINA la fórmula placeholder de su celda, y
    xl/calcChain.xml quedaría con entradas huérfanas que hacen que Excel de
    escritorio ofrezca "reparar" el archivo. Se elimina el calcChain completo
    (y su Override en [Content_Types].xml): Excel lo regenera al abrir y
    LibreOffice lo ignora — es lo mismo que hace openpyxl al guardar."""
    if "xl/calcChain.xml" not in z.namelist():
        return
    removals.add("xl/calcChain.xml")
    ct = z.read("[Content_Types].xml").decode("utf-8")
    ct = re.sub(r'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
    replacements["[Content_Types].xml"] = ct.encode("utf-8")


def main():
    with open(SCORES_PATH, encoding="utf-8") as f:
        scores = json.load(f)
    # El bloque "honor" (ganadores reales de botas/balones desde el panel admin)
    # no es un partido: se separa antes de recorrer los marcadores.
    honor = scores.pop("honor", None)
    honor = honor if isinstance(honor, dict) else {}
    honor = {k: v for k, v in honor.items() if k in HONOR_CELLS}
    if not scores and not honor:
        print("scores.json is empty — nothing to update.")
        return
    # Un player_aliases.json faltante/corrupto NUNCA debe frenar la aplicación
    # de marcadores: sin mapa, los nombres se escriben tal cual.
    amap = {}
    if honor:
        try:
            amap = load_alias_map()
        except Exception as e:
            print(f"  ! player_aliases.json ilegible ({e}) — "
                  f"los nombres de honor se escriben sin normalizar")

    pattern = os.path.join(REPO_ROOT, "ADMINExcelMundial2026*.xlsx")
    excel_files = sorted(glob.glob(pattern))
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found matching {pattern}")

    # Every file carries its own WORLDCUP sheet and computes its own players'
    # points from it, so the scores must be written into ALL of them.
    for path in excel_files:
        print(f"Patching {os.path.basename(path)} …")

        replacements = {}
        removals = set()
        with zipfile.ZipFile(path) as z:
            sheet_path = find_worldcup_sheet_path(z)
            sheet_xml = z.read(sheet_path).decode("utf-8")
            wb_xml = z.read("xl/workbook.xml").decode("utf-8")

            sheet_xml, updated = patch_sheet_xml(sheet_xml, scores)
            honor_written = 0
            if honor:
                sheet_xml, honor_written = patch_honor_cells(sheet_xml, honor, amap)
                if honor_written:
                    drop_calc_chain(z, replacements, removals)

        wb_xml = set_full_calc_on_load(wb_xml)
        replacements[sheet_path] = sheet_xml.encode("utf-8")
        replacements["xl/workbook.xml"] = wb_xml.encode("utf-8")

        rewrite_zip(path, replacements, removals)
        print(f"  Updated {updated} match(es), {honor_written} honor cell(s); "
              f"fullCalcOnLoad set.")

    print(f"Done — {len(excel_files)} file(s) patched; formulas will "
          f"recalculate when each file is opened.")


if __name__ == "__main__":
    main()
