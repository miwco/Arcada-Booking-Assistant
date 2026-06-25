"""Export the finished plan back into the Bokningsönskemål Excel format.

The original request workbooks are used as TEMPLATES so the format is reproduced
exactly: each workbook is copied, then for every course sheet the header block
(code / name / groups / examiner) is cleaned and the session table is rewritten —
one row per planned session — without changing columns, styling, or the
Groups / Event_type lookup sheets.

The plan comes from the dashboard (resolved placement + your edits). Before
writing anything we validate it; if any conflict is still unresolved (not
approved), we refuse and return the list so the caller can stop and ask.
"""
from __future__ import annotations

import glob
import os
import re

import openpyxl

from .dictionaries import APP_DIR, EXPORT_DIR, INFO, load_all
from .normalize import clean_text, normalize_course_code
from .parse_requests import _COLUMN_KEYS, _meta_from_filename
from .scheduler import SLOT_TIMES

EXPORT_ROOT = EXPORT_DIR
SEM_FOLDER = {"autumn 2026": "autumn_2026", "spring 2027": "spring_2027"}
WD_SV = {"Mon": "Måndag", "Tue": "Tisdag", "Wed": "Onsdag", "Thu": "Torsdag", "Fri": "Fredag"}
WD_ORDER = {"Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5, "": 9}
SLOT_ORDER = {"AM": 0, "FULL": 1, "PM": 2}
GROUP_RE = re.compile(r"^(Media-\d{2}(-[FLMOP])?|KP-\d{2})$")
# Other-programme group codes (EN-26, Em-26, MSE-26, MTH-26, IT-swe-26, ...) and
# elective labels are valid too — only genuinely malformed codes should warn.
OTHER_GROUP_RE = re.compile(r"^[A-Za-zÅÄÖ][A-Za-z-]*-\d{2}$")
ELECTIVE_GROUPS = {"breddstudier", "öppnayh", "öppna yh", "oppnayh"}


def _group_ok(g):
    return bool(GROUP_RE.match(g)) or g.lower() in ELECTIVE_GROUPS or bool(OTHER_GROUP_RE.match(g))


def _label_rows(ws):
    """Map header-block fields to their row number (value lives in column E)."""
    rows = {}
    for r in range(1, 9):
        a = ws.cell(r, 1).value
        if not isinstance(a, str):
            continue
        low = a.lower()
        if "kurskod" in low:
            rows["code"] = r
        elif "namn / name of study unit" in low or "studieavsnitt" in low:
            rows["name"] = r
        elif "studentgrupper" in low:
            rows["groups"] = r
        elif "examinator" in low:
            rows["examiner"] = r
    return rows


def _table(ws):
    """Return (header_row, {logical_col: col_index})."""
    for r in range(1, 15):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and "vecko" in a.lower() and "nummer" in a.lower():
            colmap = {"week": 1}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if not isinstance(v, str):
                    continue
                low = v.lower()
                for key, logical in _COLUMN_KEYS:
                    if key in low and logical not in colmap:
                        colmap[logical] = c
            return r, colmap
    return None, {}


def _comment(rec):
    parts = []
    if rec.get("placed_date") and rec.get("slot"):
        parts.append(f"{rec['placed_date']} · kl. {SLOT_TIMES.get(rec['slot'], '')}")
    if rec.get("approvedFlag") and rec.get("kinds"):
        note = rec.get("conflictNote") or ", ".join(rec["kinds"])
        parts.append(f"⚠ APPROVED DOUBLE BOOKING ({note})")
    if rec.get("comments"):
        parts.append(clean_text(rec["comments"]))
    return " | ".join(parts)


def validate(plan):
    """Return (unresolved, warnings). Unresolved = active (un-approved) conflicts."""
    unresolved, bad_groups = [], set()
    for r in plan:
        if r.get("state") == "conflict":
            unresolved.append(f"{r['course']} ({r['cohort']}) — {','.join(r.get('kinds', []))} "
                              f"on {r.get('placed_date')} {r.get('slot')}")
        for g in [x.strip() for x in (r.get("groups") or "").split(";") if x.strip()]:
            if not _group_ok(g):
                bad_groups.add(g)
    warnings = [f"group code '{g}' not in Media-YY-X / KP-YY format (check it)"
                for g in sorted(bad_groups)]
    return unresolved, warnings


def export_plan(plan, out_dir=None, semester=None, decisions=None):
    """Write the plan into template-based Excel files. `plan` is a list of
    resolved booking records from the dashboard. `semester` (e.g. "autumn 2026"
    or "spring 2027") restricts the export to that batch; None exports both.
    `decisions` (the dashboard's pre-export review list) is saved next to the
    files as a human-readable decision log so every change is on record."""
    out_dir = out_dir or EXPORT_ROOT
    if semester:
        plan = [r for r in plan if r.get("semester") == semester]
    unresolved, warnings = validate(plan)
    if unresolved:
        return {"ok": False, "error": "unresolved_conflicts",
                "unresolved": unresolved,
                "message": f"{len(unresolved)} unresolved conflict(s) — resolve or approve them, "
                           f"then export again."}

    d = load_all()
    # index plan by (cohort, semester, course_code)
    by_key = {}
    for r in plan:
        if r.get("external") or not r.get("placed_date"):
            continue
        by_key.setdefault((r["cohort"], r["semester"], r["course_code"]), []).append(r)

    files_out, total_rows, approved = [], 0, 0
    templates = sorted(g for g in glob.glob(os.path.join(INFO, "bokningsönskemålen_2026_2027", "**", "*.xlsx"),
                                            recursive=True) if not os.path.basename(g).startswith("~$"))
    for tpl in templates:
        cohort, sem, _ = _meta_from_filename(tpl)
        if semester and sem != semester:        # only the chosen batch
            continue
        folder = os.path.join(out_dir, SEM_FOLDER.get(sem, "other"))
        os.makedirs(folder, exist_ok=True)
        wb = openpyxl.load_workbook(tpl)  # keep formatting (template copy)
        wb._external_links = []           # drop external-workbook links (avoids Excel "repaired records")
        sheet_rows = 0
        for ws in wb.worksheets:
            if ws.title in ("Groups", "Event_type"):
                continue
            lr = _label_rows(ws)
            if "code" not in lr:
                continue
            raw_code = ws.cell(lr["code"], 5).value
            norm, _n = normalize_course_code(raw_code)
            canon, name, _note = d.lookup_course(norm) if norm else (None, None, "")
            code = canon or norm
            recs = by_key.get((cohort, sem, code), [])
            recs.sort(key=lambda r: (r["week"], WD_ORDER.get(r.get("weekday", ""), 9),
                                     SLOT_ORDER.get(r.get("slot", ""), 9)))
            # clean header block
            if recs:
                ws.cell(lr["code"], 5).value = code
                if "name" in lr and (name or recs[0].get("course")):
                    ws.cell(lr["name"], 5).value = name or recs[0]["course"]
                if "groups" in lr:
                    ws.cell(lr["groups"], 5).value = recs[0].get("groups", "")
                if "examiner" in lr:
                    ws.cell(lr["examiner"], 5).value = recs[0].get("examiner", "")
            hdr, cm = _table(ws)
            if hdr is None:
                continue
            # unmerge any merged ranges in the data area so every cell is writable
            for mr in list(ws.merged_cells.ranges):
                if mr.min_row > hdr:
                    ws.unmerge_cells(str(mr))
            # clear old data rows
            for r in range(hdr + 1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).value = None
            # write one row per planned session
            row = hdr + 1
            for rec in recs:
                def put(logical, value):
                    if logical in cm:
                        ws.cell(row, cm[logical]).value = value
                put("week", rec["week"])
                put("weekday", WD_SV.get(rec.get("weekday", ""), ""))
                put("times_per_week", 1)
                put("minutes", rec.get("minutes"))
                put("room", rec.get("room"))
                put("teachers", ", ".join([t for t in (rec.get("teachers") or "").split("; ") if t]))
                put("booking_type", rec.get("type") or "Lektion/Lecture")
                put("content", rec.get("content"))
                put("comments", _comment(rec))
                if rec.get("approvedFlag") and rec.get("kinds"):
                    approved += 1
                row += 1
                sheet_rows += 1
        out_path = os.path.join(folder, os.path.basename(tpl))
        wb.save(out_path)
        wb.close()
        files_out.append({"file": os.path.relpath(out_path, APP_DIR), "rows": sheet_rows})
        total_rows += sheet_rows

    log_rel = _write_decision_log(out_dir, semester, decisions, files_out, total_rows, approved)

    return {"ok": True, "files": files_out, "total_rows": total_rows,
            "approved_double_bookings": approved, "warnings": warnings,
            "semester": semester or "both semesters",
            "decision_log": log_rel,
            "export_dir": os.path.relpath(out_dir, APP_DIR)}


def _write_decision_log(out_dir, semester, decisions, files_out, total_rows, approved):
    """Save a JSON + plain-text record of every change made before this export."""
    import datetime
    import json
    decisions = decisions or []
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    tag = SEM_FOLDER.get(semester, "both") if semester else "both"
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.join(out_dir, f"decision_log_{tag}_{stamp}")
    meta = {"exported_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "semester": semester or "both semesters", "sessions_written": total_rows,
            "approved_double_bookings": approved,
            "files": [f["file"] for f in files_out], "decisions": decisions}
    with open(base + ".json", "w", encoding="utf-8") as fh:
        json.dump(meta, fh, ensure_ascii=False, indent=2)
    lines = [f"Booking Assistant — decision log",
             f"Exported: {meta['exported_at']}",
             f"Batch: {meta['semester']}  ·  {total_rows} sessions  ·  {approved} approved double-booking(s)",
             f"Files: {', '.join(meta['files']) or '(none)'}", "",
             f"Changes made before export ({len(decisions)}):"]
    if decisions:
        for dcn in decisions:
            lines.append(f"  - [{dcn.get('type','change')}] {dcn.get('text','')}")
    else:
        lines.append("  (no manual changes — exported as imported)")
    with open(base + ".txt", "w", encoding="utf-8-sig") as fh:
        fh.write("\n".join(lines) + "\n")
    return os.path.relpath(base + ".txt", APP_DIR)
