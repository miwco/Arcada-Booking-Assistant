# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the Booking Assistant.

Produces a single BookingAssistant.exe that starts the local server and opens the
dashboard — no Python install needed on the target machine.

SAFETY: only NON-personal files are bundled — the generic config.example/, the
committed course CSVs, and the anonymised _info_example/ dummy data. The real
_info/ and the real config/teacher_*.csv / workload_targets.json are NEVER bundled
(they are not listed below). When the .exe runs, it reads any real data the user
places in an _info/ folder next to the .exe and writes output/ and export/ there.
"""
import os

ROOT = os.path.abspath(os.getcwd())


def tree(rel):
    """All files under rel as (src, dest_dir) datas, preserving structure."""
    out = []
    for dp, _dirs, files in os.walk(os.path.join(ROOT, rel)):
        for f in files:
            full = os.path.join(dp, f)
            out.append((full, os.path.relpath(dp, ROOT)))
    return out


datas = [
    # non-personal config (course data + the settings template)
    ("config/course_master.csv", "config"),
    ("config/course_code_fixes.csv", "config"),
    ("config/settings.example.json", "config"),
]
datas += tree("config.example")     # generic "Teacher A" config
datas += tree("_info_example")      # anonymised dummy booking data

hiddenimports = [
    "serve", "build", "openpyxl",
    "pipeline.dictionaries", "pipeline.dashboard", "pipeline.exporter",
    "pipeline.apply_import", "pipeline.validate_import", "pipeline.ai_assist",
    "pipeline.parse_requests", "pipeline.scheduler", "pipeline.calendar_model",
    "pipeline.normalize", "pipeline.config_store", "pipeline.imports",
]

a = Analysis(
    ["run.py"],
    pathex=[ROOT, os.path.join(ROOT, "scripts")],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="BookingAssistant",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
)
