"""Turn normalized booking rows into calendar events.

Each request row carries a week number + (optional) weekday(s). Here we:
  * resolve the ISO week number to a real date (autumn weeks -> 2026,
    spring weeks -> 2027),
  * expand a row with several weekdays into one event per weekday,
  * tag each event with the specialization it targets (for colour-coding).
The result is a flat list of events the dashboard renders.
"""
from __future__ import annotations

import csv
import datetime as dt
import re

WEEKDAY_ISO = {"Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5}
WEEKDAY_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri"]

# specialization letter -> (label, css class)
SPEC = {
    "F": ("Foto och klipp", "spec-f"),
    "L": ("Ljudarbete", "spec-l"),
    "M": ("Manus och regi", "spec-m"),
    "O": ("Online media", "spec-o"),
    "P": ("Producing", "spec-p"),
    "FILM": ("Film (flera inriktningar)", "spec-film"),
    "ALLA": ("Hela årskursen", "spec-alla"),
    "OTHER": ("Övrigt / annat program", "spec-other"),
}


def specs_of(groups_str):
    """Return the set of specialization letters a session targets.

    A whole-cohort group (Media-YY with no letter) targets all five; this is
    used for the specialization filter (so a course shows under every relevant
    specialization)."""
    specs, whole_cohort, has_media = set(), False, False
    for g in [x.strip() for x in groups_str.split(";") if x.strip()]:
        m = re.fullmatch(r"Media-\d{2}(?:-([FLMOP]))?", g)
        if m:
            has_media = True
            if m.group(1):
                specs.add(m.group(1))
            else:
                whole_cohort = True
    if whole_cohort:
        specs = set("FLMOP")
    return specs, has_media


def spec_of(groups_str):
    """Single category tag used for colour-coding: one specialization (F/L/M/O/P),
    FILM (several film specializations, no online), ALLA (whole year incl. online),
    or OTHER (another programme)."""
    specs, has_media = specs_of(groups_str)
    if not has_media:
        return "OTHER"
    if len(specs) == 1:
        return next(iter(specs))
    if specs <= {"F", "L", "M", "P"}:   # several film specializations, no online
        return "FILM"
    return "ALLA"                        # whole year (includes online)


def _year_for(semester):
    return 2026 if "2026" in semester else 2027


def iso_to_date(year, week, iso_weekday):
    try:
        return dt.date.fromisocalendar(year, week, iso_weekday)
    except ValueError:
        return None


def _sessions_per_week(weekdays, times_per_week):
    """Return the list of weekdays to create (one event each).

    Honours the rule that 2 sessions/week must land on different days: explicit
    weekdays are kept; if more sessions are requested than weekdays were given,
    the extra ones are added as flexible ('') for the scheduler to place on
    other days. Capped at 5 (one per weekday)."""
    days = [w for w in weekdays.split(",") if w]
    try:
        n = int(times_per_week)
    except (ValueError, TypeError):
        n = 0
    n = n if 1 <= n <= 5 else 0
    if n and n > len(days):
        days = days + [""] * (n - len(days))
    if not days:
        days = [""] * (n or 1)
    return days[:5]


def build_events(bookings_csv, team=None):
    """Return (events, week_dates).

    week_dates maps 'semester|week' -> {weekday: isodate} so the browser can
    resolve a (re)placed weekday to a real date without ISO-week math.
    """
    team = set(team or [])
    events, week_dates = [], {}
    eid = 0
    with open(bookings_csv, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if not r["week"]:
                continue
            try:
                week = int(r["week"])
            except ValueError:
                continue
            year = _year_for(r["semester"])
            wk_key = f"{r['semester']}|{week}"
            if wk_key not in week_dates:
                week_dates[wk_key] = {wd: (iso_to_date(year, week, i).isoformat()
                                           if iso_to_date(year, week, i) else "")
                                      for wd, i in WEEKDAY_ISO.items()}
            spec = spec_of(r["groups"])
            specs, _ = specs_of(r["groups"])
            teachers = [t for t in r["teachers"].split("; ") if t]
            team_involved = any(t in team for t in teachers)
            external = bool(teachers) and not team_involved
            session_group = f"{r['cohort']}|{r['course_code']}|{week}|{r['semester']}"
            for sidx, wd in enumerate(_sessions_per_week(r["weekdays"], r["times_per_week"])):
                eid += 1
                events.append({
                    "id": eid,
                    "session_group": session_group,
                    "cohort": r["cohort"],
                    "semester": r["semester"],
                    "academic_year": r["academic_year"],
                    "week": week,
                    "weekday": wd,
                    "course_code": r["course_code"],
                    "course": r["course_name"] or r["sheet"],
                    "ects": r["course_ects"],
                    "teachers": r["teachers"],
                    "examiner": r["examiner"],
                    "groups": r["groups"],
                    "room": r["desired_room"] or r["room"],
                    "needs_computer": r["needs_computer"] == "yes",
                    "spec": spec,
                    "specs": sorted(specs),
                    "spec_label": SPEC[spec][0],
                    "spec_class": SPEC[spec][1],
                    "type": r["booking_type"],
                    "program": r.get("program", ""),
                    "context": bool(r.get("program", "")),
                    "minutes": int(r["minutes"]) if r["minutes"] else 0,
                    "time_range": r["time_range"],
                    "comment_slot": r["comment_slot"],
                    "hard_time": r["hard_time"] == "yes",
                    "double_ok": r["double_ok"] == "yes",
                    "external": external,
                    "content": r["content"],
                    "comments": r["comments"],
                })
    return events, week_dates
