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
"""

import glob
import json
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

REPO_ROOT   = os.path.join(os.path.dirname(__file__), "..")
SCORES_PATH = os.path.join(REPO_ROOT, "scores.json")

# WORLDCUP sheet columns (Excel letters): match number, home goals, away goals.
COL_NUM, COL_GH, COL_GA = "AH", "AC", "AD"

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
        for col, goal in ((COL_GH, entry.get("goal_home")), (COL_GA, entry.get("goal_away"))):
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


def set_full_calc_on_load(workbook_xml):
    if "fullCalcOnLoad" in workbook_xml:
        return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', workbook_xml)
    if "<calcPr" in workbook_xml:
        return re.sub(r"<calcPr\s", '<calcPr fullCalcOnLoad="1" ', workbook_xml, count=1)
    return workbook_xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')


def rewrite_zip(path, replacements):
    """Rewrite the zip at path, swapping only the entries in replacements
    (name -> bytes); every other entry is copied through byte-for-byte."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    shutil.move(tmp, path)


def main():
    with open(SCORES_PATH, encoding="utf-8") as f:
        scores = json.load(f)
    if not scores:
        print("scores.json is empty — nothing to update.")
        return

    pattern = os.path.join(REPO_ROOT, "ADMINExcelMundial2026*.xlsx")
    excel_files = sorted(glob.glob(pattern))
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found matching {pattern}")

    # Every file carries its own WORLDCUP sheet and computes its own players'
    # points from it, so the scores must be written into ALL of them.
    for path in excel_files:
        print(f"Patching {os.path.basename(path)} …")

        with zipfile.ZipFile(path) as z:
            sheet_path = find_worldcup_sheet_path(z)
            sheet_xml = z.read(sheet_path).decode("utf-8")
            wb_xml = z.read("xl/workbook.xml").decode("utf-8")

        sheet_xml, updated = patch_sheet_xml(sheet_xml, scores)
        wb_xml = set_full_calc_on_load(wb_xml)

        rewrite_zip(path, {
            sheet_path: sheet_xml.encode("utf-8"),
            "xl/workbook.xml": wb_xml.encode("utf-8"),
        })
        print(f"  Updated {updated} match(es); fullCalcOnLoad set.")

    print(f"Done — {len(excel_files)} file(s) patched; formulas will "
          f"recalculate when each file is opened.")


if __name__ == "__main__":
    main()
