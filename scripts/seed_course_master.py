"""One-off seeder for config/course_master.csv.

Builds the canonical course table from the reference workbook
(kurskod_och_kursnamn.xlsx), de-duplicates, applies known corrections, and adds
codes that are used in the request files but missing from the reference.

The resulting config/course_master.csv is the canonical course source going
forward (edit it by hand to fill in ECTS / fix names). Re-running this script
OVERWRITES it, so only run it to regenerate from scratch.
"""
import csv
import os
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from pipeline.normalize import clean_text, normalize_course_code  # noqa: E402

REF = os.path.join(ROOT, "_info", "om_kurserna_grupperna_utbildningen", "kurskod_och_kursnamn.xlsx")
OUT = os.path.join(ROOT, "config", "course_master.csv")

# code -> (name, ects, notes) for codes used in requests but absent/ambiguous in
# the reference, plus known ECTS. ects is left blank ("") where unknown.
ADDITIONS = {
    "MK-2-118": ("Webbdesign (Online media, åk 1)", "", "used in requests; missing from reference list"),
    "MK-2-177": ("OM - Introduktion till medieproduktion", "5", "new 5 ECTS OM-version of MK-2-113"),
    "MK-2-115": ("Slutproduktion - Online media", "", "new code; becomes 5 ECTS in 2028, not active spring 2027"),
}
# Resolve the FM-3-007 reference clash in favour of the name used in requests.
OVERRIDES = {
    "FM-3-007": ("Gatufotografering", "", "reference reuses FM-3-007 for both Webbdesign and Gatufotografering"),
}


def main():
    courses = {}  # code -> [name, ects, notes]
    wb = openpyxl.load_workbook(REF, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not row or not row[0]:
            continue
        code, _ = normalize_course_code(row[0])
        name = clean_text(row[1]) if len(row) > 1 else ""
        if not code:
            continue
        if code in courses and courses[code][0] != name:
            courses[code][2] = (courses[code][2] + "; " if courses[code][2] else "") + \
                f"reference also lists name '{name}'"
        else:
            courses.setdefault(code, [name, "", ""])
    wb.close()

    for code, (name, ects, note) in OVERRIDES.items():
        courses[code] = [name, ects, note]
    for code, (name, ects, note) in ADDITIONS.items():
        courses[code] = [name, ects, note]

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["code", "name", "ects", "notes"])
        for code in sorted(courses):
            w.writerow([code] + courses[code])
    print(f"Wrote {len(courses)} courses to {OUT}")


if __name__ == "__main__":
    main()
