"""Place booking requests into a week x (day x AM/PM) grid.

Rules agreed with the user:
  * Day has two slots: AM (09:15-12:30) and PM (13:15-16:00). A long session
    that crosses lunch fills the whole day (both slots).
  * A "conflict" is two DIFFERENT courses sharing the same slot on the same day
    for the same teacher (teacher conflict) or the same student group (group
    conflict); both at once is a combined conflict. Two courses on the same day
    in different slots is NOT a conflict.
  * Fixed requests (day + time given) are placed first. Sessions missing a day
    and/or a time are then auto-placed ("AI-placed") into a free slot, marked so.
"""
from __future__ import annotations

import datetime as dt
import re

SLOT_TIMES = {"AM": "09:15–12:30", "PM": "13:15–16:00", "FULL": "09:15–16:00"}
WEEKDAY_ISO = {"Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5}
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]
FULL_DAY_MIN = 300  # sessions this long or more occupy the whole day


SPEC_LETTERS = "FLMOP"


def group_atoms(groups_str):
    """Expand group codes to atomic specialization groups so the year-level code
    overlaps its specialization codes. Media-YY -> {Media-YY-F..P}; Media-YY-X and
    everything else stay as-is. Two bookings share students iff their atoms intersect."""
    out = set()
    for g in (groups_str or "").replace(";", ",").split(","):
        g = g.strip()
        if not g:
            continue
        m = re.fullmatch(r"Media-(\d{2})", g)
        if m:
            out.update(f"Media-{m.group(1)}-{s}" for s in SPEC_LETTERS)
        else:
            out.add(g)
    return out


def _year_for(semester):
    return 2026 if "2026" in semester else 2027


def _date(year, week, weekday):
    try:
        return dt.date.fromisocalendar(year, week, WEEKDAY_ISO[weekday])
    except (ValueError, KeyError):
        return None


def classify_slot(time_range, minutes):
    """Return 'AM' | 'PM' | 'FULL' | None (None = no time info given)."""
    if time_range:
        nums = re.findall(r"(\d{1,2})(?:[.:]\d{2})?", time_range)
        if nums:
            start = int(nums[0])
            end = int(nums[1]) if len(nums) > 1 else start
            if start < 12 and end > 13:
                return "FULL"
            return "AM" if start < 12 else "PM"
    if minutes and minutes >= FULL_DAY_MIN:
        return "FULL"
    return None


def _units(slot):
    return ["AM", "PM"] if slot == "FULL" else [slot]


def schedule(events):
    """Annotate each event in place with placement + conflict info and return it.

    Added keys: slot, placed_weekday, placed_date, placed_date_disp, source
    ('requested'|'ai-slot'|'ai-day'|'ai-both'), ai_placed (bool),
    conflict ('none'|'group'|'teacher'|'both').
    """
    # occupancy: (kind, name, isodate, unit) -> set of course codes
    occ = {}
    group_days = {}  # session_group -> set of weekdays already used (distinct-day rule)

    def res_keys(ev, date, unit):
        for g in group_atoms(ev["groups"]):
            yield ("g", g, date, unit)
        for t in [x.strip() for x in ev["teachers"].split(";") if x.strip()]:
            yield ("t", t, date, unit)

    def clashes(ev, date, slot):
        n = 0
        for unit in _units(slot):
            for k in res_keys(ev, date, unit):
                n += len([c for c in occ.get(k, ()) if c != ev["course_code"]])
        return n

    def occupy(ev, date, slot):
        for unit in _units(slot):
            for k in res_keys(ev, date, unit):
                occ.setdefault(k, set()).add(ev["course_code"])

    for ev in events:
        ev["slot"] = classify_slot(ev.get("time_range", ""), ev.get("minutes", 0)) \
            or (ev.get("comment_slot") or None)
        has_day = bool(ev.get("weekday"))
        has_slot = ev["slot"] is not None
        ev["source"] = ("requested" if has_day and has_slot else
                        "ai-slot" if has_day else
                        "ai-day" if has_slot else "ai-both")
        ev["ai_placed"] = ev["source"] != "requested"
        ev["pre_ok"] = bool(ev.get("double_ok"))

    # Place fixed first (stable, deterministic), then day-fixed, then flexible.
    # Sessions with a fixed weekday (order 0,1) run first so the distinct-day
    # rule can steer the flexible ones (order 2,3) onto other days.
    order = {"requested": 0, "ai-slot": 1, "ai-day": 2, "ai-both": 3}
    for ev in sorted(events, key=lambda e: (order[e["source"]], e["week"], e["course_code"])):
        year = _year_for(ev["semester"])
        sg = ev.get("session_group", "")
        used = group_days.get(sg, set())
        if ev.get("weekday"):
            wdays = [ev["weekday"]]
        else:
            wdays = [w for w in WEEKDAYS if w not in used] or WEEKDAYS
        slots = [ev["slot"]] if ev["slot"] else ["AM", "PM"]
        external = ev.get("external")
        best, best_score = None, None
        for wd in wdays:
            date = _date(year, ev["week"], wd)
            if not date:
                continue
            iso = date.isoformat()
            for slot in slots:
                load = sum(len(occ.get(k, ())) for u in _units(slot) for k in res_keys(ev, iso, u))
                score = (clashes(ev, iso, slot), load, WEEKDAYS.index(wd), 0 if slot != "PM" else 1)
                if best_score is None or score < best_score:
                    best_score, best = score, (wd, iso, date, slot)
        if not best:
            ev.update(placed_weekday="", placed_date="", placed_date_disp="", conflict="none")
            continue
        wd, iso, date, slot = best
        ev["slot"] = slot
        ev["placed_weekday"] = wd
        ev["placed_date"] = iso
        ev["placed_date_disp"] = date.strftime("%d.%m")
        group_days.setdefault(sg, set()).add(wd)
        if not external:  # external bookings are shown only as context, never block ours
            occupy(ev, iso, slot)

    # Conflict pass now that every event is placed (own bookings only).
    for ev in events:
        if not ev.get("placed_date") or ev.get("external"):
            ev["conflict"] = "none"
            continue
        g_conf = any(len(occ.get(("g", g, ev["placed_date"], u), ())) > 1
                     for g in group_atoms(ev["groups"])
                     for u in _units(ev["slot"]))
        t_conf = any(len(occ.get(("t", t, ev["placed_date"], u), ())) > 1
                     for t in [x.strip() for x in ev["teachers"].split(";") if x.strip()]
                     for u in _units(ev["slot"]))
        ev["conflict"] = ("both" if g_conf and t_conf else
                          "group" if g_conf else "teacher" if t_conf else "none")
    return events
