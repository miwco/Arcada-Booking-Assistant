# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the Booking Assistant.

Produces a single BookingAssistant.exe that starts the local server and opens the
dashboard — no Python install needed on the target machine.

SAFETY: only the generic config.example/ and the settings template are bundled — no
real or example-specific data. A fresh .exe starts empty; the user imports their own
teachers/courses/groups and booking files, and writes output/ and export/ next to
the .exe.
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
    ("config/settings.example.json", "config"),     # generic settings template (no data)
]
datas += tree("config.example")     # generic, non-personal starter config (Example Teacher/Course)
# No real or example-specific booking data is bundled. A fresh app starts empty and
# the user imports their own teachers/courses/groups and booking files.

hiddenimports = [
    "serve", "build", "openpyxl",
    "pipeline.dictionaries", "pipeline.dashboard", "pipeline.exporter",
    "pipeline.apply_import", "pipeline.validate_import", "pipeline.ai_assist",
    "pipeline.parse_requests", "pipeline.scheduler", "pipeline.calendar_model",
    "pipeline.normalize", "pipeline.config_store", "pipeline.imports", "pipeline.discover", "pipeline.sessions", "pipeline.holidays",
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
