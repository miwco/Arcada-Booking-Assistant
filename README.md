# Arcada Booking Assistant

Turn teachers' booking requests into clean, conflict-free Excel files — all in one
place, no spreadsheets to edit by hand.

## What it does

A local app for planning course bookings:

- **Import** teachers' booking Excel files (and your teachers, courses and groups).
- **Plan** the timetable on a weekly calendar.
- **Resolve conflicts** — with the built-in rules, or with optional AI help.
- **Export** the finished, conflict-free files in the official booking format.

Everything runs on your own computer. Nothing is uploaded anywhere.

## Use the app (.exe — no install)

1. Download **`BookingAssistant.exe`** from the
   [latest release](https://github.com/miwco/Arcada-Booking-Assistant/releases/latest).
2. Put it in a folder and double-click it. It opens in your browser and creates
   `data/`, `import/`, `export/`, `templates/` and `config/` next to itself.
3. Work through the Home screen: **Import → review → plan → resolve → export.**
   Final files appear in the **`export/`** folder.

Set teachers, courses and groups up in **⚙ Manage** — type them in or import from
Excel (download a template there). No CSV editing needed.

### Optional AI conflict solver

In **Manage → Settings** you can add a Claude API key, pick a model, set a monthly
spending cap, and write your own rules for solving conflicts. The key is stored
locally in a `.env` next to the app. Without a key, the rule-based
**“Solve all without AI”** button still works.

## Run from source (developers)

Requires Python 3.10+ with `openpyxl` (`py -m pip install openpyxl`).

```sh
py run.py            # starts the local server + opens the dashboard
```

Build the standalone .exe:

```sh
py -m pip install pyinstaller
py scripts/build_exe.py        # -> dist/BookingAssistant.exe
```

The repo ships with generic example config (`config.example/`) and no real data, so
it runs out of the box; you import your own data in the app.

## License

MIT — see [LICENSE](LICENSE).
