# Booking Assistant — Finish plan

Goal: turn the working prototype into something usable in daily work. See
`PROJECT.md` for current state. Recommendations are firm, not a menu.

> **Status (2026-06-25): all four phases below are built.** ✅ 1 one-click launcher
> + in-browser upload (`run.bat`, 📥 Import tab) · ✅ 2 Claude API assist
> (`pipeline/ai_assist.py`, propose-not-apply) · ✅ 3 trust + learning layer
> (pre-export review summary, decision log, editable 🧠 rules) · ✅ 4 reusability +
> GDPR prep (configurable data folder, `scripts/anonymise.py` → `config.example/`,
> LICENSE, README). Optional/not requested: PyInstaller `.exe`; pushing to GitHub.

## Recommended product form: a local web app with a one-click launcher

You already have the two right pieces: a Python engine and a browser dashboard.
Don't change that. Package it so there are **no commands**:

- **Now:** a `run.bat` (and `run.py`) that starts `scripts/serve.py` and opens the
  browser at the dashboard. Double-click → app opens. Offline, no install beyond Python.
- **Later:** a **PyInstaller one-file `.exe`** that bundles Python + the engine, so
  it runs on a machine without Python. Still a local web app under the hood.

Why not the alternatives: an **online/multi-user web app** adds hosting, auth,
accounts and **GDPR exposure** you don't need for personal use. A **native desktop
framework** (Electron/Tauri) is wasted effort — the browser UI already works. Keep
it **local-first**; an internal web version is a later, separate decision.

## What to build next (in order)

1. **One-click launcher + in-browser upload** *(biggest daily-use unlock)*
   - `run.bat`/`run.py`: start server, open browser.
   - An **Upload** screen in the dashboard: drop teachers' original Excel files →
     server runs `validate_import` → shows the interactive validation report →
     you approve → `apply_import` → planner loads. No `py` commands.
   - One clear loop: **Upload → Validate/approve → Plan → Resolve → Export**.

2. **Claude API integration (assist, not autopilot)** — a few local endpoints in
   `serve.py` that call the Anthropic API (key from an untracked `.env`, never the repo):
   - `interpret-comment`: read a messy "Other comments" cell → structured
     suggestion (time, room, frequency) **with the reasoning**. Use a cheap model
     (Haiku) for this high-volume step.
   - `explain-conflict` / `suggest-resolution`: given a conflict + context, propose
     the best fix and **explain why it's good**. Use a stronger model (Opus) here.
   - Every AI output is a **suggestion the user approves** — it feeds the existing
     approve-before-apply flow. The deterministic solver stays in charge of hard rules.

3. **Trust + learning layer**
   - Show on every AI/auto suggestion: the change (from→to), **confidence**, the
     **reason**, and the **source** (which field/comment it read).
   - A **pre-export review summary**: list every correction/auto-move/approval before
     files are written, so risky changes can't slip through.
   - A simple **decision log** (what changed, when, why) saved alongside the export.
   - Promote the current `ba_rules` knowledge base to **explicit, editable rules** the
     user can see/remove; the AI and solver consult them. Learn from decisions, but
     keep rules human-readable and reviewable.

4. **Reusability + GDPR + GitHub prep** (do before any push)
   - A **data-folder setting** so the app points at a configurable folder, not the
     hardcoded `_info/` — lets other programmes use it without touching code.
   - **Dummy data + `config.example/`**: generic teachers/courses/bookings so the
     public repo runs with no real data. An **anonymiser** script to generate dummy
     data from the real structure.
   - `.gitignore` already added (excludes `_info/`, `output/`, `export/`, real
     `config/*.csv`, `workload_targets.json`, `.env`). Add a LICENSE + public README.

## What to avoid for now
- Online hosting, multi-user, accounts, a database — none needed; pure GDPR risk.
- Electron/Tauri or a rewrite of the UI.
- **Full AI autopilot** that applies changes without approval.
- Generalising to other programmes before FM is fully solid.
- Any secret (API key) or real teacher/booking data in the repo.

## How the AI solver stays trustworthy (the core constraint)
1. **Deterministic for hard rules.** Students never double-booked, A211 protected,
   group hierarchy, the export gate — these stay in code, not in the model. The AI
   only assists with fuzzy work (reading comments, ranking move options, explaining).
2. **Propose, never silently apply.** AI suggestions go through the same
   approve-before-change flow as everything else.
3. **Always explain + cite.** Change, reason, confidence, and the source it read.
4. **Validate twice.** The deterministic validator + export gate run regardless of
   what the AI said, so a bad AI suggestion can't reach the booker files.
5. **Review before export.** A summary of every change, so mistakes are visible
   before they're hard to notice in Excel.
6. **Learn transparently.** Decisions become editable rules the user can audit.

## Definition of "finished enough to use daily"
- Double-click to start; upload teacher Excel in the browser.
- Validate → approve → plan → **Resolve all** → export, all without the terminal.
- AI helps read comments and suggest fixes, always with an explanation and your approval.
- Export refuses on unresolved conflicts; a review summary precedes every export.
- Real data stays local; a dummy-data version is ready for GitHub.
