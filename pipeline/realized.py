"""Parse realized bookings (the booking-system Staff-Timetable exports).

Format differs from the planned request files: one workbook per teacher, with
real dated sessions (Summary, Type, Date, Start/End, Duration, Building, Room,
Groups, Modules, ...). The Type column is NOT a reliable teaching indicator (real
teaching is often typed "Event"); the Modules column is — a row with a course
module is teaching, a row without is admin/meeting.

We classify each row as:
  * 'course' — has a course module (logged per course; ECTS matched by name where
    possible), or
  * 'admin'  — meetings / events / general info (one big bucket, kept separate
    from the course-workload target).

Co-taught sessions appear once per teacher's file, so we merge by session key and
union the teachers. Output is grouped by academic year, e.g. {"2025-2026": [...]}.
"""
from __future__ import annotations

import datetime as dt
import glob
import os
import re

import openpyxl

from .calendar_model import SPEC, spec_of, specs_of
from .dictionaries import INFO
from .normalize import clean_text

REALIZED_DIRS = {  # folder -> academic year
    "realized_bookings_2025-2026": "2025-2026",
    "actual_bookings_2025-2026": "2025-2026",
}
SPEC_WORD = {"foto och klipp": "F", "ljudarbete": "L", "manus och regi": "M",
             "online media": "O", "producentskap": "P"}


def _norm(s):
    return re.sub(r"[^a-zåäö0-9]+", " ", (s or "").lower()).strip()


def _name_index(d):
    idx = {}
    for code, name in d.courses.items():
        n = _norm(name)
        if n:
            idx.setdefault(n, (code, d.course_ects.get(code, "")))
    return idx


def _match_course(module, idx):
    n = _norm(module)
    if n in idx:
        return idx[n]
    stripped = re.sub(r"^\([^)]*\)\s*", "", re.sub(r"^[^:]*:\s*", "", module))
    n2 = _norm(stripped)
    if n2 in idx:
        return idx[n2]
    best = None
    for k, v in idx.items():
        if len(k) >= 7 and k in n and (best is None or len(k) > len(best[0])):
            best = (k, v)
    return best[1] if best else (None, "")


def _ects_num(v):
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return ""


def parse_groups(s):
    """Long realized group text -> canonical codes (Media-YY-X / KP-YY)."""
    if not s:
        return []
    out = []
    for m in re.finditer(r"film och media,\s*([a-zåäö ]+?),\s*hösten\s*(\d{4})", s, re.I):
        spec = SPEC_WORD.get(m.group(1).strip().lower())
        yy = m.group(2)[-2:]
        code = f"Media-{yy}-{spec}" if spec else f"Media-{yy}"
        if code not in out:
            out.append(code)
    for m in re.finditer(r"kulturproducent(?:skap)?,\s*hösten\s*(\d{4})", s, re.I):
        code = f"KP-{m.group(1)[-2:]}"
        if code not in out:
            out.append(code)
    return out


def _slot(start, end, minutes):
    def hh(t):
        m = re.match(r"(\d{1,2})", str(t) or "")
        return int(m.group(1)) if m else None
    s, e = hh(start), hh(end)
    if s is not None and e is not None:
        if s < 12 and e > 13:
            return "FULL"
        return "AM" if s < 12 else "PM"
    return "FULL" if (minutes or 0) >= 300 else "AM"


def _academic(date):
    if date.month >= 8:
        return f"{date.year}-{date.year+1}", f"autumn {date.year}"
    return f"{date.year-1}-{date.year}", f"spring {date.year}"


def build_realized(d, base=None):
    base = base or INFO
    team_by_key = {n.replace(" ", "").lower(): n for n in d.teachers}
    idx = _name_index(d)
    years = {}
    for folder, _ay in REALIZED_DIRS.items():
        path = os.path.join(base, folder)
        if not os.path.isdir(path):
            continue
        sessions = {}  # key -> event dict (teachers merged)
        for f in sorted(glob.glob(os.path.join(path, "*.xlsx"))):
            m = re.search(r"Staff-Timetable-(.+?)-\d+\.xlsx", os.path.basename(f))
            raw = m.group(1) if m else os.path.basename(f)
            teacher = team_by_key.get(raw.lower(), raw)
            wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
            ws = wb.worksheets[0]
            rows = list(ws.iter_rows(values_only=True))
            H = {h: i for i, h in enumerate(rows[0])}
            for r in rows[1:]:
                if not any(r):
                    continue
                ds = clean_text(r[H["Date"]])
                try:
                    date = dt.datetime.strptime(ds, "%d/%m/%Y").date()
                except (ValueError, TypeError):
                    continue
                module = clean_text(r[H["Modules"]])
                summary = clean_text(r[H["Summary"]])
                room = clean_text(r[H["Room"]])
                start = clean_text(r[H["Start Time"]]); end = clean_text(r[H["End Time"]])
                try:
                    minutes = int(float(r[H["Duration (mins)"]]))
                except (ValueError, TypeError):
                    minutes = 0
                is_course = bool(module) and "allmän information" not in module.lower() \
                    and "allmän information" not in summary.lower()
                ay, sem = _academic(date)
                key = (ay, "course" if is_course else "admin", ds, start, end,
                       (module or summary), room)
                ev = sessions.get(key)
                if ev is None:
                    groups = parse_groups(clean_text(r[H["Groups"]])) if is_course else []
                    gstr = "; ".join(groups)
                    code, ects = _match_course(module, idx) if is_course else (None, "")
                    cohorts = sorted({re.sub(r"-[FLMOP]$", "", g) for g in groups if g.startswith("Media-")}) or \
                        sorted({g for g in groups if g.startswith("KP-")})
                    spec = spec_of(gstr) if is_course else "OTHER"
                    specs, _h = specs_of(gstr) if is_course else (set(), False)
                    sp = SPEC.get(spec, SPEC["OTHER"])
                    ev = {
                        "id": f"R{len(sessions)}", "kind": "course" if is_course else "admin",
                        "course": module if is_course else summary,
                        "course_code": code or ("RZ:" + _norm(module) if is_course else ""),
                        "course_code_real": code or "", "ects": ects,
                        "groups": gstr, "cohort": "/".join(cohorts) if cohorts else ("—" if is_course else ""),
                        "examiner": "", "room": room, "building": clean_text(r[H["Building"]]),
                        "minutes": minutes, "placed_date": date.isoformat(),
                        "week": date.isocalendar()[1], "semester": sem, "academic_year": ay,
                        "slot": _slot(start, end, minutes), "type": clean_text(r[H["Type"]]),
                        "spec": spec, "spec_label": sp[0], "spec_class": sp[1], "specs": sorted(specs),
                        "state": "clean", "external": False, "needs_computer": False,
                        "_teachers": set(),
                    }
                    sessions[key] = ev
                ev["_teachers"].add(teacher)
            wb.close()
        for ev in sessions.values():
            tl = sorted(ev.pop("_teachers"))
            ev["tlist"] = tl
            ev["teachers"] = "; ".join(tl)
            years.setdefault(ev["academic_year"], []).append(ev)
    return years
