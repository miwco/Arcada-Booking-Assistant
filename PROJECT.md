# Booking Assistant — Project context (compact)

A planning tool for **Film & Media (medianom) at Arcada** that turns messy course
booking requests into clean, conflict-free Excel files for the bookers, and tracks
teacher workload. This file is the durable summary; `ROADMAP.md` is the finish plan.

## Goals
1. **Main:** produce corrected **Bokningsönskemål Excel** files (same official
   format) that are ready for the bookers and contain **no unresolved conflicts**.
2. **Second:** track **teaching per teacher** — planned now, and **realized**
   (from the real booking system) so planned vs actual can be compared across years.
3. **Real use:** a tool the user runs in daily work — upload teachers' Excel,
   visually resolve conflicts, export corrected Excel. Local is fine (offline ok).

## End-to-end flow (all built and working)
```
original per-examiner Excel  ──▶  validate (catch typos / fields-vs-comments)  ──▶  approve corrections
        │                                                                                    │
        └────────────────────────── apply_import ──────────────────────────────────────────┘
                                          │
                          validated booking rows (output/bookings_2026_2027.csv)
                                          │
                              dashboard.html  (plan: drag, auto-resolve, conflicts,
                                          │     teacher workload, realized follow-up)
                                          │
                              🪄 Resolve all  →  Export Excel (per semester)
                                          │
                          corrected booker Excel  →  export/{autumn_2026,spring_2027}/
```

## Architecture
- **Language/stack:** Python 3.10 + openpyxl; self-contained **HTML/JS dashboard**
  (no build step, no framework) served by a tiny local server. No database.
- **`pipeline/`** (the engine):
  - `normalize.py` — value cleaners (codes, groups, teachers, weekdays, minutes, comments)
  - `dictionaries.py` — canonical courses(+ECTS)/groups/teachers; encoding-tolerant
  - `parse_requests.py` — read the cleaned request workbooks → normalized `Booking` rows
  - `calendar_model.py` — rows → calendar events (week→date, category colour, week expansion)
  - `scheduler.py` — initial placement (AM/PM slots) + conflict detection (group hierarchy, A211)
  - `dashboard.py` — generates the interactive `output/dashboard.html`
  - `exporter.py` — write a plan back into the template Excel format (booker files)
  - `realized.py` — parse the booking-system Staff-Timetable exports (per academic year)
  - `validate_import.py` — validate the **original** files → suggested corrections (interactive report)
  - `apply_import.py` — apply approved corrections → validated planner source
- **`scripts/`**: `serve.py` (serves output/, `POST /export`), `seed_course_master.py`, `fill_ects.py`
- **`config/`** (editable): `teacher_aliases.csv`, `teacher_typos.csv`, `course_master.csv`
  (code/name/ECTS), `course_code_fixes.csv`, `workload_targets.json` (ratios/thresholds)
- **`_info/`** (input data), **`output/`** (CSV + dashboard), **`export/`** (booker files)

## Run
```sh
py build.py                 # planner from the cleaned files
py -m pipeline.apply_import  # planner from the validated originals (alt source)
py scripts/serve.py 8765     # then open http://localhost:8765/dashboard.html
```

## Key rules / decisions baked in
- **Conflict severity:** group (students — hardest) + **A211 film studio** are hard;
  teacher is approvable; rooms are visible conflicts. Colours: group=yellow,
  studio=red, room=orange, teacher=pink. **Category colours are cool hues only**
  (violet/indigo/cyan/blue) so they're never confused with a warning.
- **Group hierarchy:** `Media-YY` = all its `Media-YY-X` specializations (expanded to
  atomic groups before comparing). Two bookings clash iff their student atoms intersect.
- **Slots:** two per day, AM 09:15–12:30 / PM 13:15–16:00; long sessions fill the day.
- **Auto-mover:** moves into fully-clean slots, prefers Tue/Wed/Thu (Thu>Fri), keeps
  studio courses put, **moves rather than removes teachers**.
- **Comments are data:** "Other comments" is parsed for real times/rooms (e.g.
  Gatufotografering `5×45` + "10:00–15:00" → `1×300`). Validator flags fields-vs-comment.
- **Examiner = the in-sheet Examinator field** (not the file creator). Electives
  (öppnaYH / breddstudier) need no specific group.
- **Other-programme context:** a non-FM course (e.g. KP) taught by a team teacher is
  imported as **context** — counted for teacher conflicts, movable, shown distinctly
  (striped + "KP" tag), but **excluded from the FM export**.
- **Workload:** in-class hours (scheduled) ≠ course-work hours (ECTS×20÷team teachers)
  ≠ student workload (~135 h/5 ECTS). Realized = **booked hours only** (target 400–800),
  course vs admin hours kept separate. All thresholds in `config/workload_targets.json`.
- **Export safety:** refuses to export with unresolved conflicts; per-semester export;
  approved double-bookings annotated; external links stripped (no Excel "repair" prompt).
- **Edits** live in the browser (localStorage) with Undo/Reset; the Excel files stay the
  source of truth. Nothing risky is applied silently.

## Known limits (today)
- Single-user, local, manual `py` commands (no one-click launcher yet).
- No built-in AI yet (the heuristics are deterministic; Claude API not wired in).
- ECTS for realized courses is best-effort name-matched (~470/660).
- Real data + real names currently live in the repo tree (not yet GitHub-ready).
