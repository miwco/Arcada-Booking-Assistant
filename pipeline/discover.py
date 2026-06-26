"""Generate the internal lists (teachers, courses, groups, examiners) from the
imported booking Excel — so the import is self-contained and the user doesn't have
to upload separate lists.

`discover()` reads the parsed bookings; `merge_into_config()` adds anything new to
the editable config (teachers / courses / programmes), keeping existing entries and
flagging likely mismatches for review. Any manually-maintained lists stay as they
are and act as reference data for the flags.
"""
from __future__ import annotations

import re

from .dictionaries import COHORT_YEARS
from .normalize import normalize_group_code

_PROG_YEAR = re.compile(r"^([A-Za-zÅÄÖ][A-Za-zÅÄÖ-]*?)-(\d{2})(?:-[A-Za-z]{1,2})?$")


def discover(bookings):
    """-> {teachers:[...], courses:{code:{name,examiner}}, groups:[...]}."""
    teachers, courses, groups = {}, {}, set()
    for b in bookings:
        ex = (b.examiner or "").strip()
        for t in [x.strip() for x in (b.teachers or "").split("; ")] + ([ex] if ex else []):
            if t:
                teachers.setdefault(t.lower(), t)
        code = (b.course_code or "").strip().upper()
        if code:
            c = courses.setdefault(code, {"name": "", "examiner": ""})
            if b.course_name and not c["name"]:
                c["name"] = b.course_name
            if ex and not c["examiner"]:
                c["examiner"] = ex
        for g in (b.groups or "").split("; "):
            norm, _ = normalize_group_code(g, COHORT_YEARS)
            if _PROG_YEAR.match(norm):
                groups.add(norm)
    return {"teachers": sorted(teachers.values()), "courses": courses, "groups": sorted(groups)}


def merge_into_config(disc):
    """Add new teachers/courses/programmes to config; flag likely mismatches.
    Returns {teachers:{added,flagged,flag_list}, courses:{...}, groups:{...}}."""
    from . import config_store
    report = {}

    # --- teachers --------------------------------------------------------
    existing = config_store.get_teachers()
    known = set()
    for t in existing:
        known.add(t["name"].lower())
        known.update(a.lower() for a in t.get("aliases", []))
    known.update(x["wrong"].lower() for x in config_store.get_typos())
    added_t, flag_t = [], []
    for name in disc["teachers"]:
        ln = name.lower()
        if ln in known:
            continue
        if any((ln in k or k in ln) and abs(len(ln) - len(k)) <= 6 for k in known):
            flag_t.append(name)                 # looks like a variant of an existing name
        added_t.append(name)
    if added_t:
        config_store.set_teachers(existing + [{"name": n, "aliases": []} for n in added_t])
    report["teachers"] = {"added": len(added_t), "flagged": len(flag_t), "flag_list": flag_t[:25]}

    # --- courses (with examiner role) ------------------------------------
    ex_courses = {c["code"].upper(): c for c in config_store.get_courses()}
    added_c, flag_c = [], []
    for code, info in disc["courses"].items():
        if code in ex_courses:
            old = ex_courses[code]
            if info["name"] and old.get("name") and info["name"].lower() != old["name"].lower():
                flag_c.append(f"{code}: file says '{info['name']}', list has '{old['name']}'")
            continue
        added_c.append({"code": code, "name": info["name"], "ects": "",
                        "examiner": info["examiner"], "notes": ""})
    if added_c:
        config_store.set_courses(list(ex_courses.values()) + added_c)
    report["courses"] = {"added": len(added_c), "flagged": len(flag_c), "flag_list": flag_c[:25]}

    # --- programmes (new prefixes found in the group codes) --------------
    pdata = config_store.get_programs()
    prog_codes = {p["code"] for p in pdata["programs"]}
    new_progs = []
    for g in disc["groups"]:
        m = _PROG_YEAR.match(g)
        if m and m.group(1) not in prog_codes and m.group(1) != "Media":
            if m.group(1) not in [p["code"] for p in new_progs]:
                new_progs.append({"code": m.group(1), "name": m.group(1), "active": True})
    if new_progs:
        progs = [{"code": p["code"], "name": p["name"], "active": p["active"]} for p in pdata["programs"]]
        config_store.set_programs({"programs": progs + new_progs,
                                   "base_year": pdata["base_year"], "window": pdata["window"],
                                   "tracks": pdata["tracks"], "extra": pdata.get("extra", [])})
    report["groups"] = {"added": len(disc["groups"]), "new_programmes": [p["code"] for p in new_progs],
                        "flagged": 0}
    return report
