"""Build BookingAssistant.exe with PyInstaller.

  py scripts/build_exe.py

Runs PyInstaller against booking_assistant.spec. The result is
dist/BookingAssistant.exe — a single file that needs no Python install. Only the
generic config.example/ is bundled; no real or example-specific data. A fresh app
starts empty and the user imports their own data in the GUI.
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "booking_assistant.spec"],
                   cwd=ROOT, env=env, check=True)
    exe = os.path.join(ROOT, "dist", "BookingAssistant.exe")
    print("\nBuilt:", exe if os.path.exists(exe) else "(not found — check the log above)")


if __name__ == "__main__":
    main()
