"""Canonical dictionaries for courses, groups and teachers.

Per the agreed approach, raw cells are mapped onto these canonical sources:
  * courses  - config/course_master.csv (code, name, ects, notes); seeded from
               _info/.../kurskod_och_kursnamn.xlsx by scripts/seed_course_master.py
  * groups   - the 'Groups' lookup sheet in the spring request files, plus
               every valid Media-YY-X / KP-YY code generated from the known
               programme structure (gruppkoderna.txt)
  * teachers - config/teacher_aliases.csv (canonical name + nickname aliases)
               and config/teacher_typos.csv (misspelling -> canonical)
Code typos are corrected via config/course_code_fixes.csv.
"""
from __future__ import annotations

import csv
import json
import os
import re
import sys
from dataclasses import dataclass, field

import openpyxl

from .normalize import clean_text

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# When packaged with PyInstaller, read-only resources live in the bundle
# (sys._MEIPASS) while everything the app *writes* (output/, export/, _info/uploads)
# must live next to the .exe. Outside a bundle both equal the project root, so the
# normal `py`/source workflow is completely unchanged.
_FROZEN = getattr(sys, "frozen", False)
BUNDLE = getattr(sys, "_MEIPASS", ROOT) if _FROZEN else ROOT      # read-only resources
APP_DIR = os.path.dirname(sys.executable) if _FROZEN else ROOT    # writable base
CONFIG = os.path.join(BUNDLE, "config")
CONFIG_EXAMPLE = os.path.join(BUNDLE, "config.example")
OUTPUT_DIR = os.path.join(APP_DIR, "output")
EXPORT_DIR = os.path.join(APP_DIR, "export")


def cfg_path(name):
    """A config file, searched in: a config/ next to the app (lets a packaged user
    drop in their real config), then the bundled config/, then the committed generic
    config.example/ (so a public clone or fresh .exe still runs)."""
    for base in (os.path.join(APP_DIR, "config"), CONFIG, CONFIG_EXAMPLE):
        p = os.path.join(base, name)
        if os.path.exists(p):
            return p
    return os.path.join(CONFIG, name)


def data_dir():
    """Where the source booking data lives. Defaults to ./_info but can be pointed
    elsewhere (so other programmes can use the app without touching code) via the
    BA_DATA_DIR env var or config/settings.json {"data_dir": "..."}. Relative paths
    resolve against the writable app dir. If no real _info is present, the bundled
    dummy data (_info_example) is used — so a public clone / fresh .exe runs."""
    p = os.environ.get("BA_DATA_DIR")
    if not p:
        for cfg in (os.path.join(APP_DIR, "config", "settings.json"),
                    os.path.join(CONFIG, "settings.json")):
            if os.path.exists(cfg):
                try:
                    with open(cfg, encoding="utf-8-sig") as fh:
                        p = (json.load(fh) or {}).get("data_dir")
                except (ValueError, OSError):
                    p = None
                break
    if not p:
        real = os.path.join(APP_DIR, "_info")
        # only treat _info as real data if it actually holds booking workbooks —
        # an empty _info/ (e.g. just an uploads/ folder) must not hide the dummy data
        if os.path.isdir(os.path.join(real, "bokningsönskemålen_2026_2027")):
            return real
        dummy = os.path.join(BUNDLE, "_info_example")
        return dummy if os.path.isdir(dummy) else real
    return p if os.path.isabs(p) else os.path.join(APP_DIR, p)


def _settings():
    for base in (os.path.join(APP_DIR, "config"), CONFIG):
        path = os.path.join(base, "settings.json")
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8-sig") as fh:
                    return json.load(fh) or {}
            except (ValueError, OSError):
                return {}
    return {}


INFO = data_dir()
COURSE_MASTER = cfg_path("course_master.csv")
COURSE_FIXES = cfg_path("course_code_fixes.csv")
TEACHER_FILE = cfg_path("teacher_aliases.csv")
TEACHER_TYPOS = cfg_path("teacher_typos.csv")
GROUPS_SOURCE = os.path.join(INFO, "bokningsönskemålen_2026_2027", "våren_2027", "media-25-VT-2027.xlsx")

_S = _settings()
SPEC_LETTERS = "".join((_S.get("specializations") or {}).keys()) or "FLMOP"
_Y0, _Y1 = int(_S.get("cohort_year_start", 20)), int(_S.get("cohort_year_end", 26))
COHORT_YEARS = [f"{y:02d}" for y in range(_Y0, _Y1 + 1)]  # configurable cohort range


@dataclass
class Dictionaries:
    courses: dict = field(default_factory=dict)          # code -> name
    course_ects: dict = field(default_factory=dict)      # code -> ects (str, may be "")
    course_notes: dict = field(default_factory=dict)     # code -> notes
    code_fixes: dict = field(default_factory=dict)       # wrong -> (correct, reason)
    groups: dict = field(default_factory=dict)           # code -> description
    teacher_alias: dict = field(default_factory=dict)    # lower(alias) -> canonical
    teachers: list = field(default_factory=list)         # canonical names
    _name_re: object = None                              # regex over known names/aliases

    # ---- course lookups -------------------------------------------------- #
    def lookup_course(self, code: str):
        """Return (canonical_code, name, note).

        Applies code_fixes, then tries an exact match, then a
        zero-pad-insensitive match (so 'TV-2-34' resolves to 'TV-2-034')."""
        note = ""
        if code in self.code_fixes:
            correct, reason = self.code_fixes[code]
            note = f"code corrected {code} -> {correct} ({reason})"
            code = correct
        if code in self.courses:
            return code, self.courses[code], note
        key = _padkey(code)
        if key:
            for c, name in self.courses.items():
                if _padkey(c) == key:
                    pad = f"code '{code}' matched '{c}' ignoring zero-padding"
                    return c, name, (note + "; " + pad if note else pad)
        miss = f"course code '{code}' not in dictionary"
        return None, None, (note + "; " + miss if note else miss)

    def ects(self, code):
        return self.course_ects.get(code, "")

    # ---- teacher lookups ------------------------------------------------- #
    def lookup_teacher(self, name: str):
        """Return (canonical_name, note). Unknown names are kept as-is and,
        per the agreed rule, NOT flagged (note is '')."""
        cleaned = clean_text(name)
        canonical = self.teacher_alias.get(cleaned.lower())
        if canonical:
            note = "" if canonical == cleaned else f"'{cleaned}' -> '{canonical}'"
            return canonical, note
        return cleaned, ""  # unknown -> ignore

    def split_known(self, token: str):
        """If an unrecognized token actually contains several known names run
        together (e.g. 'Alex Smith Robin Lee'), return the canonical names;
        otherwise None."""
        if not self._name_re:
            return None
        found = [self.teacher_alias[m.group(0).lower()] for m in self._name_re.finditer(token)]
        # de-dup while preserving order
        seen, out = set(), []
        for n in found:
            if n not in seen:
                seen.add(n); out.append(n)
        return out if len(out) >= 2 else None

    # ---- group lookups --------------------------------------------------- #
    def is_known_group(self, code: str) -> bool:
        return code in self.groups


def _read_csv(path):
    """Read a config CSV tolerantly: try UTF-8 (with/without BOM), fall back to
    Windows-1252, since hand-edited files may be saved in either encoding."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return list(csv.DictReader(f))


def _padkey(code: str):
    parts = code.split("-")
    if len(parts) < 2 or not parts[-1].isdigit():
        return None
    return ("-".join(parts[:-1]).upper(), int(parts[-1]))


def load_courses():
    courses, ects, notes = {}, {}, {}
    for row in _read_csv(COURSE_MASTER):
        code = clean_text(row["code"]).upper()
        if not code:
            continue
        courses[code] = clean_text(row.get("name"))
        ects[code] = clean_text(row.get("ects"))
        notes[code] = clean_text(row.get("notes"))
    return courses, ects, notes


def load_code_fixes():
    fixes = {}
    if not os.path.exists(COURSE_FIXES):
        return fixes
    for row in _read_csv(COURSE_FIXES):
        wrong = clean_text(row["wrong_code"]).upper()
        correct = clean_text(row["correct_code"]).upper()
        if wrong and correct:
            fixes[wrong] = (correct, clean_text(row.get("reason")))
    return fixes


def load_groups():
    groups = {}
    s = _settings()                                   # honor in-app cohort/spec settings live
    y0, y1 = int(s.get("cohort_year_start", _Y0)), int(s.get("cohort_year_end", _Y1))
    spec_letters = "".join((s.get("specializations") or {}).keys()) or SPEC_LETTERS
    for yy in [f"{y:02d}" for y in range(y0, y1 + 1)]:
        groups[f"Media-{yy}"] = "Film och media (hela årskursen)"
        for sp in spec_letters:
            groups[f"Media-{yy}-{sp}"] = f"Film och media, {sp}"
        groups[f"KP-{yy}"] = "Kulturproducentskap"
    try:
        wb = openpyxl.load_workbook(GROUPS_SOURCE, data_only=True, read_only=True)
        if "Groups" in wb.sheetnames:
            for row in wb["Groups"].iter_rows(values_only=True):
                if row and row[0]:
                    code = clean_text(row[0])
                    desc = clean_text(row[1]) if len(row) > 1 and row[1] else ""
                    groups.setdefault(code, desc or "(cross-programme)")
        wb.close()
    except FileNotFoundError:
        pass
    return groups


def load_teachers():
    alias, canon = {}, []
    for row in _read_csv(TEACHER_FILE):
        name = (row.get("canonical_name") or "").strip()
        if not name:
            continue
        canon.append(name)
        alias[name.lower()] = name
        for a in (row.get("aliases") or "").split(";"):
            a = a.strip()
            if a:
                alias[a.lower()] = name
    if os.path.exists(TEACHER_TYPOS):
        for row in _read_csv(TEACHER_TYPOS):
            wrong = (row.get("wrong") or "").strip()
            correct = (row.get("correct") or "").strip()
            if wrong and correct:
                alias[wrong.lower()] = correct
    return alias, canon


def load_all() -> Dictionaries:
    courses, ects, notes = load_courses()
    alias, canon = load_teachers()
    # Regex of all known names+aliases, longest first, for concatenation splits.
    keys = sorted(alias.keys(), key=len, reverse=True)
    name_re = re.compile("|".join(re.escape(k) for k in keys), re.I) if keys else None
    return Dictionaries(
        courses=courses, course_ects=ects, course_notes=notes,
        code_fixes=load_code_fixes(), groups=load_groups(),
        teacher_alias=alias, teachers=canon, _name_re=name_re,
    )
