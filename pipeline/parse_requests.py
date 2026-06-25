"""Parse the 2026-2027 booking-request workbooks into normalized records.

Each course sheet is a fill-in form with a small header block (course code,
name, student groups, examiner, language) followed by a table of requested
sessions. Column positions in that table are NOT fixed (some sheets insert an
extra 'Om bokningen kräver...' column), so columns are located by header text.
"""
from __future__ import annotations

import glob
import os
from dataclasses import dataclass, field

import openpyxl

from . import normalize as nz
from .dictionaries import Dictionaries

SKIP_SHEETS = {"Groups", "Event_type"}

# header-block label (lowercased substring) -> field name
_HEADER_LABELS = {
    "kurskod": "course_code",
    "namn / name of study unit": "course_name",
    "studentgrupper": "groups",
    "examinator": "examiner",
    "undervisningsspråk": "language",
}

# table-column header (lowercased substring) -> logical column name
_COLUMN_KEYS = [
    ("veckodag", "weekday"),
    ("gånger", "times_per_week"),
    ("minuter", "minutes"),
    ("studenter", "num_students"),
    ("klassrum", "room"),
    ("undervisande", "teachers"),
    ("boknings typ", "booking_type"),
    ("event type", "booking_type"),
    ("utbidlningsprogram", "programme"),
    ("utbildningsprogram", "programme"),
    ("study program", "programme"),
    ("innehåll", "content"),
    ("content of the lecture", "content"),
    ("kommentar", "comments"),
    ("other comments", "comments"),
    ("vecko- nummer", "week"),
    ("week number", "week"),
]


@dataclass
class Flag:
    file: str
    sheet: str
    row: object
    field: str
    severity: str
    issue: str
    raw_value: str = ""


@dataclass
class Booking:
    source_file: str
    cohort: str
    semester: str
    academic_year: str
    sheet: str
    course_code: str = ""
    course_code_raw: str = ""
    course_name: str = ""
    course_name_raw: str = ""
    course_ects: object = ""
    groups: str = ""
    groups_raw: str = ""
    examiner: str = ""
    examiner_raw: str = ""
    language: str = ""
    week: object = ""
    weekdays: str = ""
    weekday_raw: str = ""
    dates_in_cell: str = ""
    times_per_week: object = ""
    minutes: object = ""
    minutes_raw: str = ""
    time_range: str = ""
    num_students: object = ""
    room: str = ""
    teachers: str = ""
    teachers_raw: str = ""
    booking_type: str = ""
    programme: str = ""
    content: str = ""
    comments: str = ""
    comment_time: str = ""
    comment_slot: str = ""
    hard_time: str = ""
    double_ok: str = ""
    desired_room: str = ""
    needs_computer: str = ""
    program: str = ""          # non-empty => other-programme context booking (e.g. KP)
    flags: str = ""


def _meta_from_filename(path):
    """media-25-HT-2026.xlsx -> ('Media-25', 'autumn', '2026-2027')."""
    name = os.path.splitext(os.path.basename(path))[0]
    parts = name.split("-")
    cohort = name
    semester, ay = "", ""
    low = name.lower()
    if "ht-2026" in low:
        semester, ay = "autumn 2026", "2026-2027"
    elif "vt-2027" in low:
        semester, ay = "spring 2027", "2026-2027"
    if parts[0].lower() == "media" and len(parts) >= 2 and parts[1].isdigit():
        cohort = f"Media-{parts[1]}"
    elif low.startswith("oppna"):
        cohort = "ÖppnaYH"
    return cohort, semester, ay


def _read_header_block(rows):
    """Return {field: raw_value} from the label column (col 0)."""
    block = {}
    for r in rows[:8]:
        if not r or not isinstance(r[0], str):
            continue
        label = r[0].lower()
        for key, fld in _HEADER_LABELS.items():
            if key in label and fld not in block:
                value = next((v for v in r[1:8] if v not in (None, "")), None)
                block[fld] = value
    return block


def _find_table(rows):
    """Return (header_row_index, {logical_col: index}) or (None, {})."""
    for i, r in enumerate(rows):
        if r and isinstance(r[0], str) and "vecko" in r[0].lower() and "nummer" in r[0].lower():
            colmap = {"week": 0}
            for j, cell in enumerate(r):
                if not isinstance(cell, str):
                    continue
                low = cell.lower()
                for key, logical in _COLUMN_KEYS:
                    if key in low and logical not in colmap.values():
                        colmap[logical] = j
            return i, colmap
    return None, {}


def parse_file(path, d: Dictionaries):
    bookings, flags = [], []
    cohort, semester, ay = _meta_from_filename(path)
    fname = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        rows = list(ws.iter_rows(values_only=True))
        block = _read_header_block(rows)
        hdr_idx, colmap = _find_table(rows)

        # --- header block: course / groups / examiner ---------------------
        code_raw = nz.clean_text(block.get("course_code"))
        code, cnotes = nz.normalize_course_code(block.get("course_code"))
        cname_raw = nz.clean_text(block.get("course_name"))
        canon_code, canon_name, lnote = d.lookup_course(code) if code else (None, None, "")
        course_code = canon_code or code
        course_name = canon_name or cname_raw

        def add_flag(field, sev, issue, raw="", row=""):
            flags.append(Flag(fname, ws.title, row, field, sev, issue, str(raw)))

        for n in cnotes:
            add_flag("course_code", "info", n, code_raw)  # successfully repaired
        if lnote:
            sev = "error" if "not in dictionary" in lnote else "info"
            add_flag("course_code", sev, lnote, code_raw)
        if canon_name and cname_raw and canon_name.lower() != cname_raw.lower():
            add_flag("course_name", "info",
                     f"sheet name '{cname_raw}' differs from dictionary '{canon_name}'", cname_raw)
        course_ects = d.ects(course_code)
        if canon_code and not course_ects:
            add_flag("ects", "info", "course ECTS unknown — fill in config/course_master.csv", code_raw)

        groups_raw = nz.clean_text(block.get("groups"))
        groups, gnotes = nz.parse_groups(block.get("groups"))
        for n in gnotes:
            add_flag("groups", "warn", n, groups_raw)
        for g in groups:
            if not d.is_known_group(g) and not g.startswith(("Media-", "KP-")):
                add_flag("groups", "warn", f"group '{g}' not in dictionary", groups_raw)

        examiner_raw = nz.clean_text(block.get("examiner"))
        examiner, enote = d.lookup_teacher(examiner_raw) if examiner_raw else ("", "missing examiner")
        if examiner_raw == "":
            add_flag("examiner", "warn", "missing examiner")
        elif enote:
            add_flag("examiner", "warn" if "unknown" in enote else "info", enote, examiner_raw)

        language = nz.clean_text(block.get("language"))

        base = dict(
            source_file=fname, cohort=cohort, semester=semester, academic_year=ay,
            sheet=ws.title, course_code=course_code, course_code_raw=code_raw,
            course_name=course_name, course_name_raw=cname_raw, course_ects=course_ects,
            groups="; ".join(groups), groups_raw=groups_raw,
            examiner=examiner, examiner_raw=examiner_raw, language=language,
        )

        if hdr_idx is None:
            add_flag("table", "error", "no session table header found on sheet")
            bookings.append(Booking(**base))
            continue

        # --- session rows -------------------------------------------------
        def get(row, logical):
            j = colmap.get(logical)
            return row[j] if j is not None and j < len(row) else None

        n_rows = 0
        for ri in range(hdr_idx + 1, len(rows)):
            row = rows[ri]
            if not row:
                continue
            week_cell = get(row, "week")
            teach_cell = get(row, "teachers")
            min_cell = get(row, "minutes")
            room_cell = get(row, "room")
            # skip fully empty rows
            if all(v in (None, "") for v in (week_cell, teach_cell, min_cell, room_cell,
                                             get(row, "weekday"), get(row, "content"))):
                continue
            n_rows += 1
            excel_row = ri + 1
            rflags = []

            week, wn = nz.parse_week(week_cell)
            for n in wn:
                add_flag("week", "warn", n, week_cell, excel_row); rflags.append(n)
            weekdays, dates, ddn = nz.parse_weekdays(get(row, "weekday"))
            for n in ddn:
                add_flag("weekday", "info", n, get(row, "weekday"), excel_row)
            minutes, trange, mn = nz.parse_minutes(min_cell)
            for n in mn:
                sev = "warn" if "could not" in n or "missing" in n else "info"
                add_flag("minutes", sev, n, min_cell, excel_row)

            comments = nz.clean_text(get(row, "comments"))
            pc = nz.parse_comment(comments)
            # Prefer an explicit time range; fall back to a time found in the comment.
            if not trange and pc["time_str"]:
                trange = pc["time_str"]
            room_cell_txt = nz.clean_text(room_cell)
            desired_room = room_cell_txt if room_cell_txt not in ("", "inget rum") else pc["room"]
            if pc["hard"] and not pc["time_str"] and not pc["slot"]:
                add_flag("comment", "info",
                         "comment states a hard time constraint but no time could be parsed — review",
                         comments, excel_row)
            if pc["double_ok"]:
                add_flag("comment", "info", "comment allows double-booking (pre-approved conflict)",
                         comments, excel_row)

            teachers_raw = nz.clean_text(teach_cell)
            tcanon = []
            for t in nz.split_people(teach_cell):
                ct, tnote = d.lookup_teacher(t)
                if tnote:  # nickname or typo correction
                    add_flag("teachers", "info", tnote, teachers_raw, excel_row)
                if ct not in d.teachers:  # unknown token: maybe several names run together
                    parts = d.split_known(t)
                    if parts:
                        add_flag("teachers", "info",
                                 f"split '{t}' -> {', '.join(parts)}", teachers_raw, excel_row)
                        tcanon.extend(parts)
                        continue
                tcanon.append(ct)  # known name, or unknown kept as-is (ignored)

            b = Booking(
                **base,
                week=week if week is not None else "",
                weekdays=",".join(weekdays),
                weekday_raw=nz.clean_text(get(row, "weekday")),
                dates_in_cell="; ".join(dates),
                times_per_week=nz.clean_text(get(row, "times_per_week")),
                minutes=minutes if minutes is not None else "",
                minutes_raw=nz.clean_text(min_cell),
                time_range=trange,
                num_students=nz.clean_text(get(row, "num_students")),
                room=nz.clean_text(room_cell),
                teachers="; ".join(tcanon),
                teachers_raw=teachers_raw,
                booking_type=nz.clean_text(get(row, "booking_type")),
                programme=nz.clean_text(get(row, "programme")),
                content=nz.clean_text(get(row, "content")),
                comments=comments,
                comment_time=pc["time_str"],
                comment_slot=pc["slot"],
                hard_time="yes" if pc["hard"] else "",
                double_ok="yes" if pc["double_ok"] else "",
                desired_room=desired_room,
                needs_computer="yes" if pc["computer"] else "",
            )
            bookings.append(b)

        if n_rows == 0:
            add_flag("table", "info", "course sheet has no session rows")
    wb.close()
    return bookings, flags


def parse_all(d: Dictionaries, info_dir=None):
    from .dictionaries import INFO
    base = info_dir or INFO
    pattern = os.path.join(base, "bokningsönskemålen_2026_2027", "**", "*.xlsx")
    bookings, flags = [], []
    for f in sorted(g for g in glob.glob(pattern, recursive=True)
                    if not os.path.basename(g).startswith("~$")):
        b, fl = parse_file(f, d)
        bookings.extend(b)
        flags.extend(fl)
    return bookings, flags
