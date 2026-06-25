"""Generate shareable, name-free stand-ins from the real data.

The only GDPR-sensitive content is **teacher names** (course codes/names and the
workload targets are not personal). This script maps every real name to a generic
label ("Teacher A", "Teacher B", …) and writes:

  py scripts/anonymise.py                  -> config.example/   (anonymised config)
  py scripts/anonymise.py --data OUTDIR     -> also copy the real _info/*.xlsx into
                                               OUTDIR with all names replaced

config.example/ is committed so the public repo runs with no real data: point the
app at it by copying its files into config/ (and using an anonymised data folder).
Aliases/typos that encode real nicknames are dropped and replaced with synthetic
examples that only show the file format.
"""
from __future__ import annotations

import csv
import glob
import json
import os
import re
import shutil
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG = os.path.join(ROOT, "config")
OUT_CONFIG = os.path.join(ROOT, "config.example")
sys.path.insert(0, ROOT)


def _read_csv(path):
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc) as f:
                return list(csv.DictReader(f))
        except (UnicodeDecodeError, FileNotFoundError):
            if not os.path.exists(path):
                return []
            continue
    return []


def _label(i):
    """0->A, 25->Z, 26->AA, …"""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return f"Teacher {s}"


def config_label_map():
    """The known team teachers (config/teacher_aliases.csv) -> 'Teacher A/B/…'."""
    names = [r["canonical_name"].strip() for r in _read_csv(os.path.join(CONFIG, "teacher_aliases.csv"))
             if (r.get("canonical_name") or "").strip()]
    return {n: _label(i) for i, n in enumerate(names)}


NAME_RE = re.compile(r"^[A-Za-zÅÄÖÆØÜåäöæøü][A-Za-zÅÄÖÆØÜåäöæøü .'\-]{1,39}$")


def collect_data_names(src):
    """Every person name that actually appears in the booking data's (cleaned)
    teacher / examiner fields — including guest lecturers NOT in the alias file
    (the gap that matters most for GDPR). Returns a set of name-like strings."""
    names = set()
    try:
        from pipeline.dictionaries import load_all
        from pipeline.parse_requests import parse_all
        d = load_all()
        bookings, _ = parse_all(d, info_dir=src)
        for b in bookings:
            for field in (b.teachers, b.examiner):           # cleaned fields only
                for nm in re.split(r"[;,/&]| och ", field or ""):
                    nm = nm.strip()
                    if NAME_RE.match(nm):                     # name-like, no digits/codes
                        names.add(nm)
    except Exception as e:                       # parsing is best-effort; tokens still cover the rest
        print(f"  (note: could not parse data names from {src}: {e})")
    return names


def build_maps(src=None):
    """Return (label_of, replacements). label_of covers the known team PLUS every
    name found in the data at `src`. replacements is a length-sorted list of
    (real_text, label) — full names, aliases, typos, and individual name tokens —
    so any full name, nickname, or lone first/surname in a cell is scrubbed."""
    label_of = config_label_map()
    nxt = len(label_of)
    if src:
        for nm in sorted(collect_data_names(src)):
            if nm not in label_of:
                label_of[nm] = _label(nxt)
                nxt += 1
    repl = dict(label_of)
    for r in _read_csv(os.path.join(CONFIG, "teacher_aliases.csv")):
        canon = (r.get("canonical_name") or "").strip()
        for a in (r.get("aliases") or "").split(";"):
            a = a.strip()
            if a and canon in label_of:
                repl[a] = label_of[canon]
    for r in _read_csv(os.path.join(CONFIG, "teacher_typos.csv")):
        wrong, correct = (r.get("wrong") or "").strip(), (r.get("correct") or "").strip()
        if wrong and correct in label_of:
            repl[wrong] = label_of[correct]
    # also scrub individual name parts (first names / surnames) so a cell that uses
    # only a first name or surname is still anonymised. Alphabetic tokens only, so a
    # cohort code / year number is never touched.
    for name, label in label_of.items():
        for tok in re.split(r"[\s.]+", name):
            if len(tok) >= 3 and tok.isalpha():
                repl.setdefault(tok, label)
    # longest first so full names are replaced before their shorter parts
    ordered = sorted(repl.items(), key=lambda kv: len(kv[0]), reverse=True)
    return label_of, ordered


def scrub(text, replacements):
    if not isinstance(text, str):
        return text
    for real, label in replacements:
        if real:
            text = re.sub(re.escape(real), label, text, flags=re.IGNORECASE)
    return text


def write_config_example(label_of):
    os.makedirs(OUT_CONFIG, exist_ok=True)
    labels = list(label_of.values()) or [_label(i) for i in range(6)]
    # teacher_aliases.csv — generic names, one example alias to show the format
    with open(os.path.join(OUT_CONFIG, "teacher_aliases.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["canonical_name", "aliases"])
        for i, name in enumerate(labels):
            w.writerow([name, "TA" if i == 0 else ""])
    # teacher_typos.csv — synthetic misspelling examples only
    with open(os.path.join(OUT_CONFIG, "teacher_typos.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["wrong", "correct"])
        if labels:
            w.writerow([labels[0].replace("Teacher", "Techer"), labels[0]])
            if len(labels) > 1:
                w.writerow([labels[1].replace("Teacher", "Teachr"), labels[1]])
    # course data is not personal — copy verbatim
    for fn in ("course_master.csv", "course_code_fixes.csv"):
        src = os.path.join(CONFIG, fn)
        if os.path.exists(src):
            shutil.copyfile(src, os.path.join(OUT_CONFIG, fn))
    # workload_targets.json — keep structure, remap the per-teacher map to labels
    src = os.path.join(CONFIG, "workload_targets.json")
    if os.path.exists(src):
        with open(src, encoding="utf-8-sig") as f:
            wt = json.load(f)
        if isinstance(wt.get("teachers"), dict):
            wt["teachers"] = {label_of.get(k, k): v for k, v in wt["teachers"].items() if k in label_of}
        with open(os.path.join(OUT_CONFIG, "workload_targets.json"), "w", encoding="utf-8") as f:
            json.dump(wt, f, ensure_ascii=False, indent=2)
    # settings example
    with open(os.path.join(OUT_CONFIG, "settings.example.json"), "w", encoding="utf-8") as f:
        json.dump({"_comment": "Copy to config/settings.json; data_dir points at your data folder.",
                   "data_dir": "_info_example"}, f, ensure_ascii=False, indent=2)
    return len(labels)


# Non-personal filenames safe to keep verbatim: cohort codes (media-23-HT-2026),
# elective/group labels (öppnaYH-…, breddstudier-…, KP-25-…).
SAFE_NAME_RE = re.compile(r"^(media-\d{2}-[A-Za-z]{2}-\d{4}|(?:öppna|oppna)yh[-_].*|breddstudier.*|KP-\d{2}.*)$", re.I)


def anonymise_data(src, dst, replacements, generic_names=True):
    """Copy every .xlsx under `src` to `dst` with all teacher names scrubbed from
    the cells. Filenames are made generic (requests_NN.xlsx) unless `generic_names`
    is False — then the scrubbed basename is kept only for known non-personal
    cohort/elective codes; any other (possibly name-bearing) file is skipped."""
    files = [g for g in glob.glob(os.path.join(src, "**", "*.xlsx"), recursive=True)
             if not os.path.basename(g).startswith("~$")]
    n, skipped = 0, []
    for path in files:
        rel = os.path.relpath(path, src)
        if generic_names:
            out_rel = os.path.join(os.path.dirname(rel), f"requests_{n + 1:02d}.xlsx")
        else:
            stem = scrub(os.path.splitext(os.path.basename(path))[0], replacements)
            if not SAFE_NAME_RE.match(stem):
                skipped.append(os.path.basename(path))      # name-bearing -> leave out of the sample
                continue
            out_rel = os.path.join(os.path.dirname(rel), stem + ".xlsx")
        out_path = os.path.join(dst, out_rel)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        wb = openpyxl.load_workbook(path)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        new = scrub(cell.value, replacements)
                        if new != cell.value:
                            cell.value = new
        wb.save(out_path)
        wb.close()
        n += 1
    if skipped:
        print(f"  (skipped {len(skipped)} name-bearing file(s): {', '.join(skipped[:4])}"
              f"{'…' if len(skipped) > 4 else ''})")
    return n, len(files)


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    src = os.environ.get("BA_DATA_DIR") or os.path.join(ROOT, "_info")
    count = write_config_example(config_label_map())     # config.example = the 13 generic team
    print(f"anonymise: wrote config.example/ ({count} generic teachers)")
    # scrub map covers the team PLUS every guest name that appears in the data
    _, replacements = build_maps(src if os.path.isdir(src) else None)
    if "--sample" in argv:
        # committed dummy dataset: only the cohort-coded booking workbooks (safe
        # filenames), cells scrubbed — so a public clone runs out of the box.
        sub = "bokningsönskemålen_2026_2027"
        s = os.path.join(src, sub)
        if not os.path.isdir(s):
            print(f"anonymise: {s} not found — cannot build sample")
        else:
            dst = os.path.join(ROOT, "_info_example", sub)
            if os.path.isdir(dst):
                shutil.rmtree(dst)
            done, total = anonymise_data(s, dst, replacements, generic_names=False)
            print(f"anonymise: built sample _info_example/{sub} ({done}/{total} workbook(s), names kept-but-scrubbed)")
    if "--data" in argv:
        i = argv.index("--data")
        dst = argv[i + 1] if i + 1 < len(argv) else os.path.join(ROOT, "_info_example")
        if not os.path.isdir(src):
            print(f"anonymise: data source {src} not found — skipping data step")
        else:
            done, total = anonymise_data(src, dst, replacements)
            print(f"anonymise: scrubbed {done}/{total} workbook(s) -> {os.path.relpath(dst, ROOT)}")
    if "--data" not in argv and "--sample" not in argv:
        print("anonymise: pass --sample (committed dummy data) or --data [OUTDIR] (full scrub)")


if __name__ == "__main__":
    main()
