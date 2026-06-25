"""Local app server for the Booking Assistant.

Serves the dashboard (output/) with UTF-8 headers and exposes a small JSON API so
the whole workflow runs in the browser — no terminal:

  GET  /uploads          -> list the uploaded teacher Excel files
  POST /upload           -> {files:[{name, b64}]}  save uploaded .xlsx
  POST /clear-uploads    -> empty the upload folder
  POST /validate         -> run import validation on the uploads -> findings
  POST /apply            -> {approved:[...]}  apply corrections, load into planner
  POST /export           -> {bookings, semester} write the booker Excel files
  GET  /ai/status        -> is the Claude API key present, which models
  POST /ai/interpret     -> {comment, course}  read a messy comment (cheap model)
  POST /ai/suggest       -> {booking, conflicts, candidates}  rank legal fixes (strong model)

Usage:  py scripts/serve.py [port]   ->  http://localhost:8765/dashboard.html
"""
import base64
import functools
import glob
import http.server
import json
import os
import sys

for _s in (sys.stdout, sys.stderr):          # be UTF-8 safe however we're launched
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from pipeline import ai_assist                               # noqa: E402
from pipeline.apply_import import run_upload                 # noqa: E402
from pipeline.dictionaries import APP_DIR, OUTPUT_DIR        # noqa: E402  (frozen-aware paths)
from pipeline.exporter import export_plan                    # noqa: E402
from pipeline.validate_import import validate_all            # noqa: E402

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
UPLOAD_DIR = os.path.join(APP_DIR, "_info", "uploads")       # user uploads live next to the app (writable)


def _uploaded():
    return sorted(os.path.basename(f) for f in glob.glob(os.path.join(UPLOAD_DIR, "*.xlsx"))
                  if not os.path.basename(f).startswith("~$"))


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".html": "text/html; charset=utf-8",
        ".csv": "text/csv; charset=utf-8",
        ".json": "application/json; charset=utf-8",
    }

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        route = self.path.rstrip("/")
        if route == "/uploads":
            return self._json(200, {"files": _uploaded()})
        if route == "/ai/status":
            return self._json(200, ai_assist.status())
        return super().do_GET()

    def do_POST(self):
        route = self.path.rstrip("/")
        try:
            if route == "/upload":
                os.makedirs(UPLOAD_DIR, exist_ok=True)
                saved = []
                for f in self._body().get("files", []):
                    name = os.path.basename(f.get("name", "")).strip()
                    if not name.lower().endswith(".xlsx") or name.startswith("~$"):
                        continue
                    with open(os.path.join(UPLOAD_DIR, name), "wb") as out:
                        out.write(base64.b64decode(f["b64"].split(",")[-1]))
                    saved.append(name)
                return self._json(200, {"ok": True, "saved": saved, "files": _uploaded()})

            if route == "/clear-uploads":
                for f in glob.glob(os.path.join(UPLOAD_DIR, "*.xlsx")):
                    os.remove(f)
                return self._json(200, {"ok": True, "files": []})

            if route == "/validate":
                if not _uploaded():
                    return self._json(200, {"ok": False, "error": "no_files",
                                            "message": "Upload some teacher Excel files first."})
                return self._json(200, {"ok": True, "files": validate_all(UPLOAD_DIR)})

            if route == "/apply":
                if not _uploaded():
                    return self._json(200, {"ok": False, "error": "no_files"})
                return self._json(200, run_upload(UPLOAD_DIR, self._body().get("approved", [])))

            if route == "/export":
                payload = self._body()
                return self._json(200, export_plan(payload.get("bookings", []),
                                                   semester=payload.get("semester") or None,
                                                   decisions=payload.get("decisions")))

            if route == "/ai/interpret":
                b = self._body()
                return self._json(200, ai_assist.interpret_comment(b.get("comment", ""),
                                                                   b.get("course", "")))

            if route == "/ai/suggest":
                return self._json(200, ai_assist.suggest_resolution(self._body()))

            self._json(404, {"ok": False, "error": "not_found"})
        except Exception as e:                                  # surface errors to the UI
            self._json(500, {"ok": False, "error": "exception", "message": str(e)})


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    handler = functools.partial(Handler, directory=OUTPUT_DIR)
    with http.server.ThreadingHTTPServer(("127.0.0.1", PORT), handler) as httpd:
        print(f"Booking Assistant running -> http://localhost:{PORT}/dashboard.html")
        httpd.serve_forever()


if __name__ == "__main__":
    main()
