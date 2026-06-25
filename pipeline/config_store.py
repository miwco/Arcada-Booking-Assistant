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

from .dictionaries import (APP_DIR, BUNDLE, DATA_DIR, EXPORT_DIR, IMPORT_DIR, OUTPUT_DIR,
                           TEMPLATES_DIR, active_cohort_years, cfg_path, data_dir,
                           generate_group_codes, media_tracks, programs)

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
    """Create the stable folders the workflow needs (all next to the app — never a
    Temp path) and seed an editable config/ + demo data so a fresh install runs and
    can be set up entirely in-app. Safe to call every start."""
    for d in (DATA_DIR, IMPORT_DIR, EXPORT_DIR, TEMPLATES_DIR, OUTPUT_DIR, CONFIG_DIR,
              os.path.join(IMPORT_DIR, "uploads")):
        os.makedirs(d, exist_ok=True)
    # seed editable config
    for name in _SEED:
        dest = os.path.join(CONFIG_DIR, name)
        if not os.path.exists(dest):
            src = cfg_path(name)          # resolves to bundle/config or config.example
            if os.path.exists(src):
                shutil.copyfile(src, dest)
    # seed a stable demo data folder next to the app (so the source isn't a Temp path)
    demo_sub = "bokningsönskemålen_2026_2027"
    if not os.path.isdir(os.path.join(DATA_DIR, demo_sub)) \
            and not os.path.isdir(os.path.join(APP_DIR, "_info", demo_sub)):
        bundled = os.path.join(BUNDLE, "_info_example", demo_sub)
        if os.path.isdir(bundled):
            shutil.copytree(bundled, os.path.join(DATA_DIR, demo_sub))
    # write the downloadable Excel templates
    try:
        from . import imports
        imports.write_templates(TEMPLATES_DIR)
    except Exception:
        pass
    return {"config_dir": CONFIG_DIR, "data_dir": data_dir(), "import_dir": IMPORT_DIR,
            "export_dir": EXPORT_DIR, "templates_dir": TEMPLATES_DIR}
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


def _save_settings_json(updates):
    s = _settings_json()
    s.update(updates)
    with open(_writable("settings.json"), "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    return s


def get_settings():
    try:
        with open(cfg_path("workload_targets.json"), encoding="utf-8-sig") as f:
            wt = json.load(f)
    except (ValueError, OSError, FileNotFoundError):
        wt = {}
    s = _settings_json()
    yearly = wt.get("yearly_coursework", {})
    inclass = wt.get("inclass_per_ects", {})
    dd = data_dir()
    return {
        "data_dir": s.get("data_dir", os.path.relpath(dd, APP_DIR) if dd.startswith(APP_DIR) else dd),
        "data_dir_full": dd, "export_dir": EXPORT_DIR, "import_dir": IMPORT_DIR,
        "templates_dir": TEMPLATES_DIR, "output_dir": OUTPUT_DIR,
        "hours_per_ects_coursework": wt.get("hours_per_ects_coursework", 20),
        "yearly_low": yearly.get("target_low", 800),
        "yearly_high": yearly.get("target_high", 1200),
        "inclass_warn_low": inclass.get("warn_low", 8),
        "inclass_warn_high": inclass.get("warn_high", 14),
    }


def set_settings(payload):
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
    if payload.get("data_dir"):
        _save_settings_json({"data_dir": payload["data_dir"]})
    return {"ok": True}


# ---- programmes & groups ---------------------------------------------------
def get_programs():
    """{programs:[{code,name,active}], reference:[{code,name}], years:[...],
    base_year, window, tracks:{...}, generated:[codes], extra:[codes]}."""
    from .dictionaries import DEFAULT_PROGRAMS
    s = _settings_json()
    progs = s.get("programs") or dict(DEFAULT_PROGRAMS)
    active = set(s.get("active_programs") or list(progs.keys()))
    years = active_cohort_years()
    base = int(s.get("active_base_year", (int(years[0]) if years else 26)))
    return {
        "programs": [{"code": c, "name": n, "active": c in active} for c, n in progs.items()],
        "reference": [{"code": c, "name": n} for c, n in DEFAULT_PROGRAMS.items()],
        "years": years, "base_year": base, "window": int(s.get("active_window", 4)),
        "tracks": media_tracks(), "extra": s.get("extra_groups") or [],
        "generated": sorted(generate_group_codes().keys()),
    }


def set_programs(payload):
    """payload: {programs:[{code,name,active}], base_year, window, tracks:{}, extra:[]}."""
    upd = {}
    if isinstance(payload.get("programs"), list):
        progs, active = {}, []
        for p in payload["programs"]:
            code = (p.get("code") or "").strip()
            if not code:
                continue
            progs[code] = (p.get("name") or "").strip() or code
            if p.get("active", True):
                active.append(code)
        upd["programs"] = progs
        upd["active_programs"] = active
    if "base_year" in payload:
        upd["active_base_year"] = int(payload["base_year"]) % 100
    if "window" in payload:
        upd["active_window"] = max(1, int(payload["window"]))
    if isinstance(payload.get("tracks"), dict):
        upd["specializations"] = {k.upper()[:2]: v for k, v in payload["tracks"].items() if k.strip()}
    if isinstance(payload.get("extra"), list):
        upd["extra_groups"] = sorted({(g or "").strip() for g in payload["extra"] if (g or "").strip()})
    _save_settings_json(upd)
    return {"ok": True}
