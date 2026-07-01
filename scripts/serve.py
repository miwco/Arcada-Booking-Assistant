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
from pipeline import ai_assist, config_store, imports, sessions  # noqa: E402
from pipeline.apply_import import run_upload                 # noqa: E402
from pipeline.dictionaries import APP_DIR, EXPORT_DIR, IMPORT_DIR, INFO, OUTPUT_DIR, data_dir  # noqa: E402
from pipeline.exporter import export_plan                    # noqa: E402
from pipeline.validate_import import validate_all            # noqa: E402

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
UPLOAD_DIR = os.path.join(IMPORT_DIR, "uploads")             # uploaded booking files (under import/)


def _status():
    """Summary for the Home screen: where things are + what's loaded."""
    csv_path = os.path.join(OUTPUT_DIR, "bookings_2026_2027.csv")
    n = 0
    if os.path.exists(csv_path):
        with open(csv_path, encoding="utf-8-sig") as fh:
            n = max(0, sum(1 for _ in fh) - 1)
    return {"ok": True, "bookings": n, "has_plan": os.path.exists(csv_path),
            "uploads": len(_uploaded()), "data_dir": data_dir(),
            "export_dir": EXPORT_DIR, "ai": ai_assist.status().get("available", False),
            "teachers": len(config_store.get_teachers()), "courses": len(config_store.get_courses())}


def _rebuild():
    """Regenerate the planner CSV + dashboard from the current data and config."""
    import build
    build.main()
    return _status()


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

    def _binary(self, data, filename):
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        route = self.path.rstrip("/")
        if route == "/uploads":
            return self._json(200, {"files": _uploaded()})
        if route == "/ai/status":
            return self._json(200, ai_assist.status())
        if route == "/status":
            return self._json(200, _status())
        if route == "/config/teachers":
            return self._json(200, {"teachers": config_store.get_teachers(),
                                    "typos": config_store.get_typos()})
        if route == "/config/courses":
            return self._json(200, {"courses": config_store.get_courses()})
        if route == "/config/settings":
            return self._json(200, config_store.get_settings())
        if route == "/config/programs":
            return self._json(200, config_store.get_programs())
        if route == "/config/ai":
            return self._json(200, config_store.get_ai())
        if route == "/config/rules":
            return self._json(200, {"ok": True, "rules": config_store.get_rules(),
                                    "path": config_store.RULES_FILE})
        if route == "/sessions":
            from pipeline.dictionaries import current_academic_year
            act = sessions.active()
            base = int((act.get("year") or current_academic_year())[2:4])
            years = [f"20{y:02d}-20{y + 1:02d}" for y in range(base - 3, base + 2)]
            return self._json(200, {"ok": True, "active": act, "list": sessions.list_sessions(),
                                    "years": years, "production_periods": sessions.meta().get("production_periods", [])})
        if route.startswith("/template/"):
            kind = route.rsplit("/", 1)[-1]
            if kind in ("teachers", "courses", "groups"):
                return self._binary(imports.template_bytes(kind), f"{kind[:-1]}_template.xlsx")
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
                b = self._body()
                return self._json(200, run_upload(UPLOAD_DIR, b.get("approved", []),
                                                  b.get("programme"), b.get("year")))

            if route == "/session/select":
                b = self._body()
                sessions.activate(b.get("programme"), b.get("year"))
                return self._json(200, {"ok": True})

            if route == "/session/periods":
                res = sessions.set_periods(self._body().get("periods", []))
                if res.get("ok"):
                    _rebuild()               # regenerate dashboard so the periods show
                return self._json(200, res)

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

            if route == "/config/teachers":
                b = self._body()
                config_store.set_teachers(b.get("teachers", []))
                config_store.set_typos(b.get("typos", []))
                return self._json(200, {"ok": True})

            if route == "/config/courses":
                return self._json(200, config_store.set_courses(self._body().get("courses", [])))

            if route == "/config/settings":
                return self._json(200, config_store.set_settings(self._body()))

            if route == "/config/programs":
                return self._json(200, config_store.set_programs(self._body()))

            if route == "/config/ai":
                return self._json(200, config_store.set_ai(self._body()))

            if route == "/config/rules":
                return self._json(200, config_store.set_rules(self._body().get("rules", "")))

            if route in ("/import/teachers", "/import/courses", "/import/groups"):
                raw = base64.b64decode((self._body().get("b64") or "").split(",")[-1])
                if route.endswith("teachers"):
                    return self._json(200, {"ok": True, **imports.parse_teachers(raw)})
                if route.endswith("courses"):
                    return self._json(200, {"ok": True, **imports.parse_courses(raw)})
                return self._json(200, {"ok": True, **imports.parse_groups(raw)})

            if route == "/import/realized":
                # keep simple for now: just store the uploaded files for later analysis
                dest = os.path.join(IMPORT_DIR, "realized")
                os.makedirs(dest, exist_ok=True)
                saved = 0
                for f in self._body().get("files", []):
                    name = os.path.basename(f.get("name", "")).strip()
                    if not name or name.startswith("~$"):
                        continue
                    with open(os.path.join(dest, name), "wb") as out:
                        out.write(base64.b64decode((f.get("b64") or "").split(",")[-1]))
                    saved += 1
                return self._json(200, {"ok": True, "saved": saved, "dir": dest})

            if route == "/rebuild":
                return self._json(200, _rebuild())

            self._json(404, {"ok": False, "error": "not_found"})
        except Exception as e:                                  # surface errors to the UI
            self._json(500, {"ok": False, "error": "exception", "message": str(e)})


def main():
    config_store.bootstrap()                    # create folders + seed editable config (first run)
    handler = functools.partial(Handler, directory=OUTPUT_DIR)
    with http.server.ThreadingHTTPServer(("127.0.0.1", PORT), handler) as httpd:
        print(f"Booking Assistant running -> http://localhost:{PORT}/dashboard.html")
        httpd.serve_forever()


if __name__ == "__main__":
    main()
