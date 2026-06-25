"""Read/write the editable config (teachers, courses, settings) for the GUI.

So the user never touches a CSV: the Manage screens call these helpers through the
server. Reads use the normal fallback (config/ -> bundle -> config.example/);
writes always go to a real config/ folder next to the app (APP_DIR), which is
created and seeded on first run by bootstrap().
"""
from __future__ import annotations

import csv
import json
import os
import shutil

from .dictionaries import APP_DIR, CONFIG_EXAMPLE, EXPORT_DIR, INFO, OUTPUT_DIR, cfg_path, data_dir

CONFIG_DIR = os.path.join(APP_DIR, "config")
# files seeded into a fresh install so the app works immediately and is editable
_SEED = ["teacher_aliases.csv", "teacher_typos.csv", "course_master.csv",
         "course_code_fixes.csv", "workload_targets.json"]


def _read_csv(path):
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc) as f:
                return list(csv.DictReader(f))
        except FileNotFoundError:
            return []
        except UnicodeDecodeError:
            continue
    return []


def _writable(name):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    return os.path.join(CONFIG_DIR, name)


def _write_csv(name, fieldnames, rows):
    with open(_writable(name), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fieldnames})


# ---- bootstrap: folders + seed config (first run / install) ----------------
def bootstrap():
    """Create the folders the workflow needs and seed an editable config/ so a
    fresh install runs and can be managed in-app. Safe to call every start."""
    for d in (OUTPUT_DIR, EXPORT_DIR, os.path.join(APP_DIR, "_info", "uploads")):
        os.makedirs(d, exist_ok=True)
    for name in _SEED:
        dest = os.path.join(CONFIG_DIR, name)
        if not os.path.exists(dest):
            src = cfg_path(name)          # resolves to bundle/config or config.example
            if os.path.exists(src):
                os.makedirs(CONFIG_DIR, exist_ok=True)
                shutil.copyfile(src, dest)
    return {"config_dir": CONFIG_DIR, "export_dir": EXPORT_DIR, "output_dir": OUTPUT_DIR}


# ---- teachers --------------------------------------------------------------
def get_teachers():
    rows = _read_csv(cfg_path("teacher_aliases.csv"))
    return [{"name": (r.get("canonical_name") or "").strip(),
             "aliases": [a.strip() for a in (r.get("aliases") or "").split(";") if a.strip()]}
            for r in rows if (r.get("canonical_name") or "").strip()]


def set_teachers(items):
    rows = [{"canonical_name": (it.get("name") or "").strip(),
             "aliases": "; ".join(a.strip() for a in (it.get("aliases") or []) if a.strip())}
            for it in items if (it.get("name") or "").strip()]
    _write_csv("teacher_aliases.csv", ["canonical_name", "aliases"], rows)
    return {"ok": True, "count": len(rows)}


def get_typos():
    return [{"wrong": (r.get("wrong") or "").strip(), "correct": (r.get("correct") or "").strip()}
            for r in _read_csv(cfg_path("teacher_typos.csv"))
            if (r.get("wrong") or "").strip() and (r.get("correct") or "").strip()]


def set_typos(items):
    rows = [{"wrong": (it.get("wrong") or "").strip(), "correct": (it.get("correct") or "").strip()}
            for it in items if (it.get("wrong") or "").strip() and (it.get("correct") or "").strip()]
    _write_csv("teacher_typos.csv", ["wrong", "correct"], rows)
    return {"ok": True, "count": len(rows)}


# ---- courses ---------------------------------------------------------------
def get_courses():
    return [{"code": (r.get("code") or "").strip().upper(), "name": (r.get("name") or "").strip(),
             "ects": (r.get("ects") or "").strip(), "notes": (r.get("notes") or "").strip()}
            for r in _read_csv(cfg_path("course_master.csv")) if (r.get("code") or "").strip()]


def set_courses(items):
    rows = [{"code": (it.get("code") or "").strip().upper(), "name": (it.get("name") or "").strip(),
             "ects": str(it.get("ects") or "").strip(), "notes": (it.get("notes") or "").strip()}
            for it in items if (it.get("code") or "").strip()]
    _write_csv("course_master.csv", ["code", "name", "ects", "notes"], rows)
    return {"ok": True, "count": len(rows)}


# ---- settings (workload targets, data/export folders, cohorts/specs) -------
SPEC_LABELS = {"F": "Foto", "L": "Ljud", "M": "Manus", "O": "Online", "P": "Producing"}


def _settings_json():
    for cfg in (os.path.join(CONFIG_DIR, "settings.json"), os.path.join(os.path.dirname(cfg_path("course_master.csv")), "settings.json")):
        if os.path.exists(cfg):
            try:
                with open(cfg, encoding="utf-8-sig") as f:
                    return json.load(f) or {}
            except (ValueError, OSError):
                return {}
    return {}


def get_settings():
    try:
        with open(cfg_path("workload_targets.json"), encoding="utf-8-sig") as f:
            wt = json.load(f)
    except (ValueError, OSError, FileNotFoundError):
        wt = {}
    s = _settings_json()
    yearly = wt.get("yearly_coursework", {})
    inclass = wt.get("inclass_per_ects", {})
    return {
        "data_dir": s.get("data_dir", os.path.relpath(data_dir(), APP_DIR) if data_dir().startswith(APP_DIR) else data_dir()),
        "export_dir": EXPORT_DIR,
        "output_dir": OUTPUT_DIR,
        "hours_per_ects_coursework": wt.get("hours_per_ects_coursework", 20),
        "yearly_low": yearly.get("target_low", 800),
        "yearly_high": yearly.get("target_high", 1200),
        "inclass_warn_low": inclass.get("warn_low", 8),
        "inclass_warn_high": inclass.get("warn_high", 14),
        "cohort_year_start": int(s.get("cohort_year_start", 20)),
        "cohort_year_end": int(s.get("cohort_year_end", 26)),
        "specializations": s.get("specializations", SPEC_LABELS),
    }


def set_settings(payload):
    # workload_targets.json (merge onto existing)
    try:
        with open(cfg_path("workload_targets.json"), encoding="utf-8-sig") as f:
            wt = json.load(f)
    except (ValueError, OSError, FileNotFoundError):
        wt = {}
    if "hours_per_ects_coursework" in payload:
        wt["hours_per_ects_coursework"] = int(payload["hours_per_ects_coursework"])
    wt.setdefault("yearly_coursework", {})
    if "yearly_low" in payload:
        wt["yearly_coursework"]["target_low"] = int(payload["yearly_low"])
    if "yearly_high" in payload:
        wt["yearly_coursework"]["target_high"] = int(payload["yearly_high"])
    wt.setdefault("inclass_per_ects", {})
    if "inclass_warn_low" in payload:
        wt["inclass_per_ects"]["warn_low"] = int(payload["inclass_warn_low"])
    if "inclass_warn_high" in payload:
        wt["inclass_per_ects"]["warn_high"] = int(payload["inclass_warn_high"])
    with open(_writable("workload_targets.json"), "w", encoding="utf-8") as f:
        json.dump(wt, f, ensure_ascii=False, indent=2)
    # settings.json (data folder, cohorts, specializations)
    s = _settings_json()
    for key in ("data_dir",):
        if payload.get(key):
            s[key] = payload[key]
    for key in ("cohort_year_start", "cohort_year_end"):
        if key in payload:
            s[key] = int(payload[key])
    if isinstance(payload.get("specializations"), dict):
        s["specializations"] = payload["specializations"]
    with open(_writable("settings.json"), "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    return {"ok": True}
