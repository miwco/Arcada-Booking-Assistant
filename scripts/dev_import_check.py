"""Fast, read-only import diagnostic — no .exe rebuild needed.

  py scripts/dev_import_check.py [FOLDER]

Parses a folder of booking Excel files (default: the spring originals) exactly like
the real import, and prints what the importer sees WITHOUT writing anything:

  - rows read / bookings created / skipped (with reasons)
  - cohorts detected (with counts)
  - group codes that failed to match (with a suggested normalised form)
  - teachers / courses that would be generated
  - validation findings grouped by field + severity

Use this to iterate on parsing/validation rules quickly; only rebuild the .exe when
the behaviour is right. Nothing here changes your config or planner.
"""
import collections
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

from pipeline.apply_import import build_bookings, corrections_from_list   # noqa: E402
from pipeline.dictionaries import load_all                               # noqa: E402
from pipeline import discover as disc_mod                                # noqa: E402
from pipeline.validate_import import validate_all                        # noqa: E402

DEFAULT = os.path.join(ROOT, "_info", "original_bokningsönskemål_spring_2027")


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else DEFAULT
    if not os.path.isdir(folder):
        print(f"folder not found: {folder}")
        return
    print(f"== import check: {folder} ==\n")
    d = load_all()
    bk, skipped, n_ctx, applied, report = build_bookings(d, folder, {})

    print(f"files            : {report['files']}")
    print(f"rows read        : {report['rows_read']}")
    print(f"bookings created : {report['bookings_created']}")
    print(f"skipped          : {report['skipped']}")
    print("\ncohorts detected :")
    for c, n in sorted(report["detected"].items(), key=lambda kv: -kv[1]):
        print(f"   {c:<14} {n}")

    if report["unmatched"]:
        print("\ngroup codes that FAILED to match (placed under 'Unassigned'):")
        for u in report["unmatched"][:30]:
            print(f"   {u['sheet'][:24]:<24} '{u['raw']}'  -> suggest: {u['suggest'] or '?'}")

    skip_reasons = collections.Counter(s["reason"] for s in skipped)
    if skip_reasons:
        print("\nskip reasons     :", dict(skip_reasons))

    disc = disc_mod.discover(bk)
    print(f"\nwould generate   : {len(disc['teachers'])} teachers, {len(disc['courses'])} courses, "
          f"{len(disc['groups'])} group codes")
    print("teachers (sample):", ", ".join(disc["teachers"][:12]))

    # validation findings grouped by field + severity
    files = validate_all(folder)
    fcount = collections.Counter()
    for f in files:
        for c in f["courses"]:
            for fd in c["findings"]:
                fcount[(fd["field"], fd["severity"])] += 1
    print("\nvalidation findings (field, severity -> count):")
    for (field, sev), n in sorted(fcount.items(), key=lambda kv: -kv[1]):
        print(f"   {field:<18} {sev:<6} {n}")


if __name__ == "__main__":
    main()
