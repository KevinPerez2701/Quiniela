#!/usr/bin/env python3
"""
Reads scores.json and writes the home/away goals back into the WORLDCUP
sheet of every ADMINExcelMundial2026*.xlsx file.

scores.json format:
  { "<match_num>": { "goal_home": <int|null>, "goal_away": <int|null> }, ... }

After this script runs, call LibreOffice headlessly to recalculate formulas,
then call excel_to_json.py to regenerate docs/data.json.
"""

import glob
import json
import os
import warnings

warnings.filterwarnings("ignore")
import openpyxl

REPO_ROOT   = os.path.join(os.path.dirname(__file__), "..")
SCORES_PATH = os.path.join(REPO_ROOT, "scores.json")

# Column indices (0-based) in the WORLDCUP sheet, matching extract_fixture().
COL_NUM    = 33   # match number
COL_GOAL_H = 28   # home goals
COL_GOAL_A = 29   # away goals


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

    # Only the primary (first) file contains the WORLDCUP sheet.
    primary = excel_files[0]
    print(f"Writing scores to {os.path.basename(primary)} …")

    wb = openpyxl.load_workbook(primary)   # NOT data_only — we need to write
    ws = wb["WORLDCUP"]

    updated = 0
    for row in ws.iter_rows():
        num_cell = row[COL_NUM]
        if num_cell.value is None:
            continue
        try:
            key = str(int(num_cell.value))
        except (ValueError, TypeError):
            continue
        if key not in scores:
            continue
        entry = scores[key]
        gh = entry.get("goal_home")
        ga = entry.get("goal_away")
        row[COL_GOAL_H].value = int(gh) if gh is not None else None
        row[COL_GOAL_A].value = int(ga) if ga is not None else None
        updated += 1

    wb.save(primary)
    print(f"Updated {updated} match(es) in {os.path.basename(primary)}")


if __name__ == "__main__":
    main()
