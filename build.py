"""Step 1 entrypoint.

Run:  py build.py

Reads the 2026-2027 booking-request workbooks under _info/, normalizes them
against the canonical dictionaries, and writes to output/:
  * bookings_2026_2027.csv  - one row per requested session
  * flags.csv               - everything that needed interpretation / failed
  * dict_courses.csv, dict_groups.csv, dict_teachers.csv - the dictionaries used
Also prints a short summary to the console.
"""
from __future__ import annotations

import csv
import os
from dataclasses import asdict, fields

from pipeline.dashboard import generate as generate_dashboard
from pipeline.dictionaries import OUTPUT_DIR, load_all
from pipeline.parse_requests import Booking, parse_all

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = OUTPUT_DIR


def write_csv(path, fieldnames, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def main():
    os.makedirs(OUT, exist_ok=True)
    d = load_all()
    bookings, flags = parse_all(d)

    booking_fields = [f.name for f in fields(Booking)]
    write_csv(os.path.join(OUT, "bookings_2026_2027.csv"), booking_fields,
              [asdict(b) for b in bookings])

    flag_fields = ["file", "sheet", "row", "field", "severity", "issue", "raw_value"]
    write_csv(os.path.join(OUT, "flags.csv"), flag_fields,
              [{k: getattr(fl, k) for k in flag_fields} for fl in flags])

    write_csv(os.path.join(OUT, "dict_courses.csv"), ["code", "name", "ects", "notes"],
              [{"code": c, "name": n, "ects": d.course_ects.get(c, ""),
                "notes": d.course_notes.get(c, "")} for c, n in sorted(d.courses.items())])
    write_csv(os.path.join(OUT, "dict_groups.csv"), ["code", "description"],
              [{"code": c, "description": n} for c, n in sorted(d.groups.items())])
    write_csv(os.path.join(OUT, "dict_teachers.csv"), ["canonical_name"],
              [{"canonical_name": t} for t in sorted(d.teachers)])

    # ---- console summary -------------------------------------------------
    sev = {"error": 0, "warn": 0, "info": 0}
    for fl in flags:
        sev[fl.severity] = sev.get(fl.severity, 0) + 1
    sessions = sum(1 for b in bookings if b.week != "")
    files = sorted({b.source_file for b in bookings})
    print("=" * 64)
    print("STEP 1 - booking-request normalization")
    print("=" * 64)
    print(f"Files parsed         : {len(files)}")
    print(f"Course sheets        : {len(set((b.source_file, b.sheet) for b in bookings))}")
    print(f"Session rows         : {sessions}")
    print(f"Flags                : {len(flags)}  "
          f"(error={sev.get('error',0)}, warn={sev.get('warn',0)}, info={sev.get('info',0)})")
    missing_ects = sorted({b.course_code for b in bookings
                           if b.course_code and not b.course_ects})
    print(f"Courses missing ECTS : {len(missing_ects)}")
    print("\nTop flag types:")
    counts = {}
    for fl in flags:
        counts[fl.issue] = counts.get(fl.issue, 0) + 1
    for issue, c in sorted(counts.items(), key=lambda x: -x[1])[:12]:
        print(f"  {c:3d}  {issue}")
    dash_path, n_events = generate_dashboard(os.path.join(OUT, "bookings_2026_2027.csv"))
    print(f"\nWrote outputs to: {OUT}")
    print(f"Dashboard        : {dash_path} ({n_events} calendar events)")


if __name__ == "__main__":
    main()
