# Booking Assistant — Master Dashboard

A tool to understand, visualize, check and coordinate course bookings for the
**Film and Media (medianom)** programme at Arcada.

## Quick start (one click, no terminal)

1. Install Python 3.10+ once (`py -m pip install openpyxl`).
2. **Double-click `run.bat`** (or `py run.py`). It starts the local server and
   opens the dashboard in your browser.
3. In the browser, the loop is: **📥 Import** (drop the teachers' Excel files →
   Validate & review → Import) → **plan & resolve** (drag, 🪄 Resolve all) →
   **Export Excel** (a review summary, then the booker files).

Everything runs locally and offline. Optional AI assist (reading messy comments,
ranking conflict fixes) turns on when you add a Claude API key — see **AI assist**.

## Status

**Step 1 — data layer (booking requests 2026–2027)**: parses the fill-in
request workbooks into one normalized table, mapping messy raw cells onto
canonical dictionaries for courses (incl. ECTS), groups and teachers, and emits
a flags report.

**Step 2 — interactive planning dashboard** (`output/dashboard.html`): a
self-contained web page. The requests are run through a **scheduler** that
places every session into a week × weekday × slot grid (two slots: AM
09:15–12:30, PM 13:15–16:00; long sessions fill the day). Fixed requests
(day + time) are placed first; sessions missing a day and/or time are
auto-placed into a free slot and marked as AI-placed. Views:
- **Planning calendar** — cohort tabs + a **specialization filter**
  (Foto/Ljud/Manus/Online/Producing), AM/PM slots per day, colour-coded by
  specialization, with distinct markers:
  - **group clash** = yellow ring + `G` (most important — students),
  - **A211 studio clash** = red ring + `🎬`,
  - **room clash** = orange ring + `R`,
  - **teacher clash** = pink ring + `T` (approvable),
  - **AI-placed** = dashed border + `AI` badge; teacher-requested = solid border.

  Colour scheme: **conflict colours (yellow / red / orange / pink) are reserved
  for conflicts only.** Course/category colours are cool hues that can't be
  confused with a warning — film specializations are violet shades (small
  differences), **Film (several inriktningar)** is indigo, **Online media** is
  cyan, **Hela årskursen** is a vivid blue, another programme is slate. The
  conflict ring shows the most important clash; the badge lists all of them.
- **Teacher overview** — **your team only**, exploratory, with four view modes:
  - **Ranked bars** — pick a metric (in-class hours / course-work target / other
    course-work / sessions / courses / examiner-of / conflicts / approved) with an
    optional breakdown (specialization / semester / year).
  - **Scatter** — plot teachers on any two axes (e.g. course-work target vs
    in-class hours), with an equality reference line.
  - **Weekly in-class trend** — in-class hours per week (team, or one teacher via
    the top filter).
  - **Course comparison** — per course: in-class hours vs ECTS, h/ECTS, flagged
    if outside the normal band.

  A **"Possible workload issues"** panel flags (with careful wording) teachers
  who *may be* under/over their yearly course-work target and courses that *may
  have* too few/many in-class hours for their ECTS. Top filters scope everything.
- **Realized bookings** — the same analytics applied to the **actual** sessions
  exported from the booking system (folder `_info/realized_bookings_2025-2026/`),
  with a **year selector**. It's a follow-up/analysis view, separate from
  planning. Realized data has real dates/times but no course codes or examiner;
  the **Modules** column marks teaching (Type is unreliable), so rows with a
  course module are logged **per course** (ECTS matched by name where possible)
  and meetings/events are summed into one **admin-hours** bucket, kept separate
  from the course-workload target (on top of it, within the ~1600 h work year).
  Co-taught sessions are merged across the per-teacher files. For realized the
  focus is **booked hours only** (no per-teacher course-work estimate): course
  hours (overlaps count once — the longer booking) and admin hours are shown
  separately, and the teacher warning uses the `realized_booked` target
  (**400–800 h**, configurable) against booked course hours. Planned-vs-realized
  and multi-year comparisons are the next step (the data is keyed by academic
  year to make that easy).
- **Teacher detail** — click any teacher name (in a lecture popup, the overview,
  or a conflict list) to open their profile: KPIs (hours, expected, sessions,
  courses, examiner-of, conflicts) and a per-course table (role, ECTS, sessions,
  hours) plus every conflict/approved conflict involving them.
- **Conflicts** — the slots where two different courses collide for the same
  teacher or group.

Filters (semester, teacher, free-text search) apply across all views.

Conflict logic: two courses on the same day in **different** slots are not a
conflict — only an overlap in the **same** slot for a shared teacher or group is.
Group codes are **hierarchical**: `Media-YY` (whole year) expands to all its
specialization groups (`Media-YY-F/L/M/O/P`), so a booking on the general year
code conflicts with one on any specialization code (and vice-versa) — same
students, same block. Two different specializations only clash if they overlap or
one side is the full year. This applies to both conflict detection and the
auto-mover (it never places a course where the same year group is already booked
under either coding).

### Planning behaviour

- **Comments are read as booking data.** The "Other comments" column is mined
  for times ("9:15–11:45 absolut tid"), AM/PM words (Förmiddag/Eftermiddag/Hela
  dagen), hard constraints (absolut / necessary), computer-room needs
  (Dataklass → 🖥 badge) and double-booking permission ("får dubbelbokas").
- **Full-day sessions** render as both an AM and a PM block (no separate row).
- **2 sessions/week ⇒ two different days** — extra auto-placed sessions avoid
  the days already used by that course/week.
- **Conflicts are editable.** Click any lecture for a popup that explains the
  clash (which teachers / groups / rooms / courses are involved) and offers:
  **Move course X / Y** (auto-finds a free slot with no teacher/group/room
  clash), **Auto-resolve** (priority order: move one lecture → move the other →
  suggest removing a *non-examiner* teacher → keep), **Approve / keep**, a manual
  move (with a list of free slots), and **Remove teacher** (the examiner is
  marked ⭐ and kept by default). "Får dubbelbokas" requests are pre-approved.
  Edits are saved in the browser (localStorage); conflicts and workload recompute live.
- **🪄 Resolve all** — one click runs a greedy solver that **moves** lectures
  (never removes teachers) into fully-clean slots to clear as many conflicts as
  possible, following the planning rules: students never double-booked, A211
  studio courses kept put, preferred days Tue/Wed/Thu (Thu before Fri), keep
  lectures near their requested week. Anything it can't solve safely is shown in
  a decision prompt with concrete choices (open & **Move course X/Y**, **remove a
  non-examiner teacher**, **approve/keep**, **skip**) plus **"Approve all
  remaining → unblock export"**. A learned-rule checkbox ("auto-approve teacher
  clashes") is remembered in the browser and re-applied next time — the start of
  a planning knowledge base.
- **The planner never blocks you.** Any move/drop is allowed even into a busy
  slot; a clear red **toast warning** names the double-booking it created, and the
  conflict markers update so you can decide. Auto-move and the green/red hints are
  guidance, not gates.
- **Remove lecture** — the popup can move *or* remove a whole lecture (excluded
  from the calendar, conflicts, workload and export; undoable).
- **Undo** — an Undo button and **Ctrl/Cmd+Z** revert the last edit (full history).
- **Reset** — a Reset button reverts everything to the original import after a
  confirmation prompt.
- **ⓘ Help** — a top-right info button explains the page and every marker (AI,
  moved, OK, T/G/R/🎬 conflicts, etc.).
- Full-day lectures are **linked** — AM and PM move together as one.
- The "moved" badge only shows when a lecture differs from its imported slot
  (it clears if you put it back).
- **Auto-move searches across weeks** — it prefers the current week and nearest
  weeks but will move a lecture to another week if that's the only clash-free slot,
  and says so. Free-slot suggestions and the manual Move box include a week selector.
- **Drag-and-drop:** drag a lecture to another slot. Same-week valid slots are
  pre-highlighted **green**; the hovered slot shows live **green/red** validity.
  Dropping anywhere works (within the same cohort); an invalid drop flashes and warns.
- **Room conflicts** can be resolved from the popup: pick a **free room** at that
  slot (suggested) or type one in.

### Conflict severity & how auto-resolve decides

Clashes are ranked, not treated equally:

| Clash | Severity | Auto-mover behaviour |
|---|---|---|
| Student **group** double-booked | **hard** | never auto-placed into it; approval needed to keep |
| **A211 film studio** (`STUDIO`) | **hard** | strongest (pulsing red) marker; never auto-placed into it; approval needed |
| **Teacher** double-booked | medium (approvable) | avoided; tolerated only if no fully-clean slot exists |
| Other **rooms** (A210/A206/…) | **soft** | a gentle warning; never blocks a move |

`slotEval` scores a candidate slot: level 0 = fully clean, +1 soft room clash,
+2 teacher clash, **∞ = hard block** (group / A211 / same course twice a day).
`autoMove` searches every slot in the lecture's semester (sorted by least
disruption: same week → nearest week → same AM/PM → earliest day) and picks the
**lowest-severity** reachable slot, returning immediately on a fully-clean one.
"Move course X" moves that lecture; **Auto-resolve** moves the least-disruptive
lecture in the clash first, then others; if nothing non-hard is reachable it
suggests removing a non-examiner teacher, else keep. Manual moves and drags are
never blocked — they place and warn (A211 and group clashes warn loudest).

**Rooms:** a booking can hold **several rooms at once** — `A211 + A206` (also
`&`, `och`) means both simultaneously (multi-camera: studio + sound), while
`A309, A312` (also `/`, `eller`) is read as alternatives (first one used).
`A211` is the only room treated as a hard conflict for now.
- **Room clashes** are detected for concrete rooms (a single named room — not
  "online" or multi-room alternatives) alongside teacher/group clashes.
- **Filters:** All / Film (every specialization except Online) / Online presets,
  plus independent F/L/M/O/P toggles for any combination (e.g. Producers + Online).
- **Compare view:** two cohorts in one shared table — each week is a single row
  with both cohorts' days side by side, so **weeks stay aligned** (no vertical
  drift). Cross-cohort teacher/room clashes show as conflicts in both.
- **Unresolved teacher overlaps don't double-count hours** — the teacher is
  credited only with the longer of the overlapping lectures.
- **External bookings** (any session with teachers but none on your team) are
  excluded from planning, conflicts and workload; toggle "show external
  (context)" to see them greyed out.

The Excel files stay the source of truth; the finished plan is exported back to
the original Bokningsönskemål format (see **Export** below).

## View the dashboard

```sh
py build.py
py scripts/serve.py 8765      # then open http://localhost:8765/dashboard.html
```

Use `scripts/serve.py` (not plain `http.server`) so the **Export Excel** button
works — it adds a `POST /export` endpoint.

## Export to Excel (the final files for the bookers)

The **⬇ Export Excel** button in the dashboard sends the resolved plan (with your
moves / approvals / room changes) to the local server, which writes finished
workbooks into `export/autumn_2026/` and `export/spring_2027/` — one file per
cohort, in the **exact original Bokningsönskemål format** (the originals are used
as templates, so columns, styling, and the `Groups`/`Event_type` sheets are
preserved). Each planned session becomes one row (week, weekday, duration, room,
teacher, type) with the concrete **date + time block** in the comments.

The export follows the **semester selector** in the header: choose *Autumn 2026*
or *Spring 2027* to export just that batch (you usually work on one at a time), or
*All semesters* for both. Exported files have their external-workbook links
stripped, so Excel opens them without a "Repaired Records" prompt. Non-FM group
codes (Breddstudier, EN-26, Em-26, MSE-26, MTH-26, IT-swe-26, …) are accepted and
no longer warned about.

Before writing, it validates and refuses to export if any conflict in the chosen
batch is still **unresolved** (you must resolve or approve each first). **Approved** double
bookings are exported with a clear `⚠ APPROVED DOUBLE BOOKING (…)` note in the
comments. Course codes, names, group codes and teacher names are canonicalised on
the way out; odd group codes (non `Media-YY-X` / `KP-YY`) are listed as warnings.

`pipeline/exporter.py` does the writing; `scripts/serve.py` exposes it.

## Import validation (original Bokningsönskemål → clean data)

The original request files (`_info/original_bokningsönskemål_spring_2027/`, one
workbook **per examiner**) are messy and contain many blank template tabs.
`pipeline/validate_import.py` validates them before planning:

```sh
py -m pipeline.validate_import     # writes output/import_validation.html + .csv
```

It skips empty/template tabs, identifies the **examiner from the filename** (a
clue to which courses belong), reads the structured fields **and** the
"Other comments" field, and flags issues with a **suggested correction**:

- **Fields contradicting the comment** — the marquee check. E.g. Gatufotografering
  is filled in as `5 × 45 min` but the comment says "Wednesdays 10:00–15:00", so
  it suggests `1 × 300 min`.
- Unknown / messy course codes, course-code↔name mismatches.
- Bad / missing group codes (not `Media-YY-X` / `KP-YY`).
- Examiner field disagreeing with the file owner or the usual examiner.
- Messy text and unusual values (e.g. very short sessions booked many times/week).

`output/import_validation.html` is an **interactive report**: each suggestion has
**Approve / Reject** toggles (saved in the browser), and a **Download approved
corrections** button (`approved_corrections.json`). Nothing is changed
automatically — you approve first. Examiner = each course's **Examinator field**
(authoritative; the filename is only the creator); electives (öppnaYH /
breddstudier) need no specific group.

### Applying the corrections → planner source

```sh
py -m pipeline.apply_import output/approved_corrections.json   # or: --all (apply every suggested fix)
```

`pipeline/apply_import.py` applies the approved corrections to copies of the
original per-examiner files (saved under `export/validated_spring_2027/`), parses
them through the normal cleaning pipeline, regroups each course to its **cohort**
(from its groups; electives → ÖppnaYH), combines with the cleaned **autumn** data, and writes
`output/bookings_2026_2027.csv` + regenerates the dashboard — so **planning now
starts from the validated, corrected originals**. The final **booker Excel** is
then produced from the dashboard's **Export Excel** button.

(Note: `py build.py` rebuilds the planner source from the *cleaned* files;
`py -m pipeline.apply_import` rebuilds it from the *validated originals*. Run
whichever source you want the planner to use.)

**Other-programme context bookings:** a non-Film&Media course (e.g. KP / Culture
Producers) taught by one of *our* teachers is imported as a **context** booking
(its program label is kept). Context bookings appear in their own cohort tab with
a purple **KP** tag and a dashed/striped chip (clearly distinct from FM courses);
they **count for teacher-conflict detection** (so an FM booking won't be scheduled
on top of a teacher's KP class) and **can be moved** to help resolve a clash; but
they are **excluded from the FM booker Excel export**. Non-FM courses with no team
teacher are ignored.

## Workload targets & thresholds — `config/workload_targets.json`

All workload targets and warning thresholds live in `config/workload_targets.json`
(editable, with notes). Key distinction the tool enforces:

- **In-class hours** = scheduled teaching only (what the calendar measures).
- **Course-work hours** = all course-related work (planning, teaching, prep,
  feedback, meetings, evaluation, communication). Default `5 ECTS ≈ 100 h`
  (`hours_per_ects_coursework: 20`).
- **Student workload** (`~135 h / 5 ECTS`) is separate and shown for reference
  only — never mixed into teacher figures.

Configurable: the ECTS→course-work ratio, the in-class h/ECTS band
(`warn_low 8 / normal 10–12 / warn_high 14`, i.e. ~40 / 50–60 / 70 h per 5 ECTS),
the yearly course-work target range (`800–1200 h`, working year `1600 h`), and
per-teacher overrides (`fte`, or explicit `yearly_low`/`yearly_high`). These are
planning guides, not absolute rules — the dashboard only *flags possibilities*.

`output/dashboard.html` is fully self-contained (data embedded), so it can also
be opened directly from disk or shared as a single file.

## AI assist (optional — propose, never apply)

The app works fully without AI. If you add a Claude API key, two assists turn on,
both **suggestion-only** — you approve, the deterministic engine applies:

- **Read the comment** — a messy "Other comments" cell is turned into a structured
  hint (time, slot, room, frequency, double-booking) with a reason (Haiku).
- **Ask AI to rank fixes** — for a conflict, the solver computes the *legal* move
  options and the AI ranks/explains them, recommends one, with confidence (Opus).

Every AI output shows the **change, reason, confidence and the source it read**.
The deterministic core keeps all hard rules (students never double-booked, A211
protected, group hierarchy, the export gate); a bad suggestion can't reach the
files. To enable: copy `.env.example` to `.env`, paste your key, restart. Endpoints:
`GET /ai/status`, `POST /ai/interpret`, `POST /ai/suggest` in `pipeline/ai_assist.py`.

## Trust: review before export + decision log

Export never writes silently. **Confirm before write** shows a review summary of
**every change** for the batch (moves, removed lectures/teachers, room changes,
approved double-bookings, plus how many lectures kept their system-chosen slot).
On confirm, a **decision log** (`export/decision_log_<batch>_<timestamp>.json` and
`.txt`) is saved next to the files, recording what changed and when. Rules you
teach the assistant (e.g. "auto-approve teacher clashes") are listed in a
**🧠 Learned rules** panel (ⓘ Help) where you can remove any of them; the built-in
safety rules are not editable.

## Reusing it elsewhere & GDPR / GitHub

- **Configurable data folder.** The app reads `_info/` by default but can point
  anywhere via `config/settings.json` (`{"data_dir": "..."}`, copy
  `config/settings.example.json`) or the `BA_DATA_DIR` env var — so another
  programme can use it without touching code.
- **Nothing personal is committed.** `.gitignore` excludes `_info/`, `output/`,
  `export/`, the real `config/*.csv` / `workload_targets.json` / `settings.json`,
  and `.env`. Real teacher and booking data stay on your machine.
- **Dummy data for sharing.** `py scripts/anonymise.py` writes `config.example/`
  with generic teachers ("Teacher A", …); add `--data OUTDIR` to also copy the
  `_info` Excel files with every name replaced and generic filenames. Use these to
  run/share the project publicly with no real data.

## Requirements

- Python 3.10+ (`py` launcher on Windows) with `openpyxl` (`py -m pip install openpyxl`).
- Source data lives in `_info/` by default (configurable; see above).

## Run

```sh
py build.py
```

Outputs are written to `output/`:

| File | Contents |
|---|---|
| `bookings_2026_2027.csv` | One row per requested session (normalized). |
| `flags.csv` | Every value that was repaired, ambiguous, or unrecognized. |
| `dict_courses.csv` | Course code → name dictionary actually used. |
| `dict_groups.csv` | Canonical group codes (Media-YY-X, KP-YY, cross-programme). |
| `dict_teachers.csv` | Canonical teacher names. |

`flags.csv` has a `severity` column: `error` (could not resolve, e.g. course
code missing from the reference), `warn` (repaired or unknown, check it), and
`info` (FYI, e.g. minutes parsed from free text).

## How it works

```
pipeline/
  normalize.py      value-level cleaners (codes, groups, teachers, weekdays, minutes)
  dictionaries.py   loads canonical courses (+ECTS) / groups / teachers
  parse_requests.py walks the workbooks, locates each form's header + table, emits records
  calendar_model.py turns booking rows into calendar events (week->date, specialization, weekday expansion)
  scheduler.py      places sessions into week x weekday x AM/PM slots; tags AI-placed + conflicts
  dashboard.py      renders the self-contained output/dashboard.html
build.py            entrypoint: runs the pipeline, writes output/ + dashboard + a summary
scripts/
  seed_course_master.py  (re)generate config/course_master.csv from the reference workbook
  fill_ects.py           derive ECTS from the curriculum box sizes into course_master.csv
config/
  teacher_aliases.csv   editable canonical teacher names + known aliases/typos
```

Design notes (driven by the real data):

- **Columns are located by header text, not position** — some sheets insert an
  extra "Om bokningen kräver…" column that shifts everything right.
- **Course codes** are cleaned of non-breaking hyphens, tabs/newlines and
  version suffixes (`(0)`/`(1)`), then matched zero-pad-insensitively
  (`TV-2-34` → `TV-2-034`).
- **Groups** are canonicalized to `Media-YY-X` / `Media-YY` / `KP-YY`
  (`Media-2024-F` → `Media-24-F`, `media-25` → `Media-25`, `Media-OM 25` →
  `Media-25-O`). Unrecognized / other-programme groups are kept verbatim and flagged.
- **Teachers** are matched against `config/teacher_aliases.csv`. Add new
  canonical names and aliases there as they surface in `flags.csv`.

## Course ECTS (study points)

`config/course_master.csv` has an `ects` column that drives teacher workload
(5 ECTS ≈ 100 h, 10 ECTS ≈ 200 h of lecturer time). Fill it in for each course;
courses with a blank `ects` are reported in `flags.csv` (severity `info`) and in
the build summary so the list of remaining values is easy to track.

## Resolved data corrections (applied automatically)

- Missing/typo course codes are handled: `MK-2-118`, `MK-2-177`, `MK-2-115`
  added to the course master; `ML-2-125 → MK-2-125` and `MK-2-172 → MK-2-131`
  via `config/course_code_fixes.csv`.
- `FM-3-007` is reused for two courses in the reference; resolved to
  "Gatufotografering" (the name used in requests) — see note in the course master.
- Teacher spelling typos are corrected via `config/teacher_typos.csv`; nicknames
  live in `config/teacher_aliases.csv`. Names not on either list (guest / other-
  programme lecturers) are kept verbatim in the output and intentionally not flagged.

## Note

The `_info/.../våren_2027/` spring-2027 request files are not present (the folder
is empty); only the autumn-2026 files are parsed currently.
