"""Saved plannings ("sessions"), keyed by academic year + degree programme.

Each import saves the plan under sessions/<programme>-<year>/ (bookings.csv +
session.json). Re-importing the same year+programme overwrites it. The planner can
switch between saved sessions, so you can look back at how a previous year was
booked. session.json also holds the manually-entered production periods.
"""
from __future__ import annotations

import csv
import datetime
import json
import os
import re
import shutil
from dataclasses import asdict, fields

from .dictionaries import APP_DIR, OUTPUT_DIR
from .parse_requests import Booking

SESSIONS_DIR = os.path.join(APP_DIR, "sessions")
_FIELDS = [f.name for f in fields(Booking)]


def slug(programme, year):
    return re.sub(r"[^a-z0-9._-]", "", f"{programme}-{year}".strip().lower().replace(" ", "-"))


def _dir(programme, year):
    return os.path.join(SESSIONS_DIR, slug(programme, year))


def _active_csv():
    return os.path.join(OUTPUT_DIR, "bookings_2026_2027.csv")


def save(programme, year, bookings):
    """Write/overwrite the session's bookings; keep any existing production periods."""
    d = _dir(programme, year)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "bookings.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=_FIELDS)
        w.writeheader()
        for b in bookings:
            w.writerow(asdict(b))
    m = _read_meta(d)
    now = datetime.datetime.now().isoformat(timespec="seconds")
    m.update(programme=programme, year=year, updated=now)
    m.setdefault("created", now)
    m.setdefault("production_periods", [])
    _write_meta(d, m)
    return m


def _read_meta(d):
    p = os.path.join(d, "session.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as fh:
                return json.load(fh) or {}
        except (ValueError, OSError):
            pass
    return {}


def _write_meta(d, m):
    with open(os.path.join(d, "session.json"), "w", encoding="utf-8") as fh:
        json.dump(m, fh, ensure_ascii=False, indent=2)


def set_active(programme, year):
    from . import config_store
    config_store._save_settings_json({"active_session": slug(programme, year),
                                      "active_programme": programme, "active_year": year})


def activate(programme, year):
    """Make a saved session the active plan: copy its bookings + regenerate dashboard."""
    src = os.path.join(_dir(programme, year), "bookings.csv")
    if not os.path.exists(src):
        return None
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    shutil.copyfile(src, _active_csv())
    set_active(programme, year)
    from .dashboard import generate
    return generate(_active_csv())


def active():
    from .dictionaries import _settings
    s = _settings()
    return {"programme": s.get("active_programme"), "year": s.get("active_year"),
            "slug": s.get("active_session")}


def list_sessions():
    out = []
    if os.path.isdir(SESSIONS_DIR):
        for name in sorted(os.listdir(SESSIONS_DIR)):
            m = _read_meta(os.path.join(SESSIONS_DIR, name))
            if not m:
                continue
            bp = os.path.join(SESSIONS_DIR, name, "bookings.csv")
            n = 0
            if os.path.exists(bp):
                with open(bp, encoding="utf-8-sig") as fh:
                    n = max(0, sum(1 for _ in fh) - 1)
            out.append({"slug": name, "programme": m.get("programme", ""), "year": m.get("year", ""),
                        "bookings": n, "updated": m.get("updated", "")})
    return out


def meta(programme=None, year=None):
    a = active()
    programme = programme or a["programme"]
    year = year or a["year"]
    if not programme or not year:
        return {"production_periods": []}
    m = _read_meta(_dir(programme, year))
    m.setdefault("production_periods", [])
    return m


def set_periods(periods, programme=None, year=None):
    a = active()
    programme = programme or a["programme"]
    year = year or a["year"]
    if not programme or not year:
        return {"ok": False, "error": "no_active_session"}
    d = _dir(programme, year)
    os.makedirs(d, exist_ok=True)
    m = _read_meta(d)
    m["production_periods"] = periods
    m.setdefault("programme", programme)
    m.setdefault("year", year)
    _write_meta(d, m)
    return {"ok": True}
