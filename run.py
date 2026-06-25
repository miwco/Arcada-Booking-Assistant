"""Booking Assistant launcher — double-click run.bat (or `py run.py`).

Starts the local server and opens the dashboard in your browser. Everything else
(upload teacher Excel, validate, plan, resolve, export) happens in the browser.
"""
import os
import sys
import threading
import time
import webbrowser

for _s in (sys.stdout, sys.stderr):          # be UTF-8 safe however we're launched
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
PORT = 8765
URL = f"http://localhost:{PORT}/dashboard.html"


def _open_browser():
    time.sleep(1.3)
    webbrowser.open(URL)


def main():
    from pipeline.dictionaries import OUTPUT_DIR
    if not os.path.exists(os.path.join(OUTPUT_DIR, "dashboard.html")):
        try:
            import build
            build.main()                       # build a dashboard so the page loads
        except Exception as e:                 # fresh/empty setup is fine — upload will create it
            print("(no dashboard yet — upload files in the browser to create one)", e)
    print(f"Opening {URL}")
    threading.Thread(target=_open_browser, daemon=True).start()
    sys.argv = ["serve", str(PORT)]
    import serve
    serve.main()


if __name__ == "__main__":
    main()
