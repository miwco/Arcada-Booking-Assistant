"""Build BookingAssistant.exe with PyInstaller.

  py scripts/build_exe.py

Regenerates the anonymised dummy data first (so the bundled sample is current),
then runs PyInstaller against booking_assistant.spec. The result is
dist/BookingAssistant.exe — a single file that needs no Python install. Real data
is never bundled; put an _info/ folder next to the .exe to use real bookings.
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    # refresh config.example/ and the committed dummy data sample
    subprocess.run([sys.executable, os.path.join("scripts", "anonymise.py"), "--sample"],
                   cwd=ROOT, env=env, check=True)
    # build
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "booking_assistant.spec"],
                   cwd=ROOT, env=env, check=True)
    exe = os.path.join(ROOT, "dist", "BookingAssistant.exe")
    print("\nBuilt:", exe if os.path.exists(exe) else "(not found — check the log above)")


if __name__ == "__main__":
    main()
