"""Low-level value normalizers.

Every function here is deliberately tolerant: it cleans what it can and returns
a list of human-readable notes describing anything that was ambiguous, repaired
or unrecognized. The pipeline turns those notes into rows of the flags report,
so over-reporting here is intentional.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

# Characters that show up as "hyphens" in the source files but are not ASCII '-'.
_DASH_CHARS = "‐‑‒–—―−"  # ‐ ‑ ‒ – — ― −

SPECIALIZATIONS = {
    "F": "Foto och klipp",
    "L": "Ljudarbete",
    "M": "Manus och regi",
    "O": "Online media",
    "P": "Producing",
}
# Words / abbreviations that map onto a specialization letter.
_SPEC_WORD = {"f": "F", "l": "L", "m": "M", "o": "O", "p": "P", "om": "O", "fk": "F"}

# Swedish / English weekday spellings -> short code.
_WEEKDAYS = {
    "Mon": ["må", "man", "mån", "måndag", "mandag", "mon", "monday"],
    "Tue": ["ti", "tis", "tisdag", "tue", "tues", "tuesday"],
    "Wed": ["on", "ons", "onsdag", "wed", "wednesday"],
    "Thu": ["to", "tor", "tors", "torsdag", "thu", "thur", "thurs", "thursday"],
    "Fri": ["fr", "fre", "fredag", "fri", "friday"],
}
_WEEKDAY_LOOKUP = {spelling: code for code, spellings in _WEEKDAYS.items() for spelling in spellings}
# Full day spellings only (safe for range detection: "måndag till onsdag").
_FULL_DAYS = {"måndag": "Mon", "mandag": "Mon", "tisdag": "Tue", "onsdag": "Wed",
              "torsdag": "Thu", "fredag": "Fri", "monday": "Mon", "tuesday": "Tue",
              "wednesday": "Wed", "thursday": "Thu", "friday": "Fri"}
_WD_ORDER = ["Mon", "Tue", "Wed", "Thu", "Fri"]
_DAY_ALT = "|".join(sorted(_FULL_DAYS, key=len, reverse=True))
_DAYRANGE_RE = re.compile(rf"\b({_DAY_ALT})\b\s*(?:till|t\.?o\.?m\.?\.?|to|[-–—])\s*\b({_DAY_ALT})\b", re.I)

_DATE_RE = re.compile(r"\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})\b")
_TIMERANGE_RE = re.compile(r"(\d{1,2}[.:]\d{2})\s*[-–]\s*(\d{1,2}[.:]\d{2})")


def clean_text(value) -> str:
    """Trim, drop tabs/newlines, collapse internal whitespace."""
    if value is None:
        return ""
    s = str(value).replace("\t", " ").replace("\n", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def _normalize_dashes(s: str) -> str:
    for ch in _DASH_CHARS:
        s = s.replace(ch, "-")
    return s


# --------------------------------------------------------------------------- #
# Course codes
# --------------------------------------------------------------------------- #
def normalize_course_code(raw):
    """Return (canonical_code, notes).

    Canonical form is UPPER, ASCII hyphens, no whitespace, no version suffix.
    Numeric segment is left as-is (we do not invent zero-padding); padding
    differences are detected later against the course dictionary.
    """
    notes = []
    if raw is None or str(raw).strip() == "":
        return "", ["empty course code"]
    original = str(raw)
    s = clean_text(original)
    if any(ch in original for ch in _DASH_CHARS):
        notes.append("non-standard hyphen character in code")
    if original != original.strip() or "\t" in original or "\n" in original:
        notes.append("whitespace/tab/newline in code cell")
    s = _normalize_dashes(s)
    # Strip a trailing version marker like " (0)" or "(1)".
    m = re.search(r"\s*\((\d+)\)\s*$", s)
    if m:
        notes.append(f"dropped version suffix '({m.group(1)})'")
        s = s[: m.start()].strip()
    s = s.replace(" ", "")
    if "--" in s:
        notes.append("double hyphen collapsed in code")
        s = re.sub(r"-+", "-", s)
    s = s.upper()
    return s, notes


# --------------------------------------------------------------------------- #
# Student groups
# --------------------------------------------------------------------------- #
_MEDIA_RE = re.compile(r"media", re.I)
_KP_RE = re.compile(r"\bkp[\s-]*(\d{1,4})", re.I)


def _year2(digits: str):
    """Map a 2- or 4-digit year string to a 2-digit cohort year, or None."""
    if len(digits) >= 4:
        digits = digits[-2:]
    if len(digits) == 1:
        return None
    return digits


def parse_groups(raw):
    """Parse a free-text 'Studentgrupper' cell into canonical group codes.

    Returns (groups, notes). Group codes look like 'Media-25-O', 'Media-25'
    (whole cohort) or 'KP-24'. Non Media/KP programmes are returned verbatim
    (cleaned) and flagged as 'other programme'.
    """
    notes = []
    if raw is None or str(raw).strip() == "":
        return [], ["empty student-group cell"]
    cell = str(raw)
    if cell != cell.strip() or "\t" in cell or "\n" in cell:
        notes.append("whitespace/tab in group cell")
    if "(" in cell:
        notes.append("group cell contains parenthetical note")
    if re.search(r"\balla\b|inriktningar", cell, re.I):
        notes.append("group cell says 'alla/inriktningar' (whole cohort)")

    groups = []

    def add(code):
        if code and code not in groups:
            groups.append(code)

    # Media tokens (handle spec letter before OR after the year).
    for m in _MEDIA_RE.finditer(cell):
        window = cell[m.start(): m.start() + 18].lower()
        year_m = re.search(r"(\d{2,4})", window)
        if not year_m:
            notes.append("Media group without a year")
            continue
        yy = _year2(year_m.group(1))
        if yy is None:
            notes.append(f"unclear Media year in '{window.strip()}'")
            continue
        if len(year_m.group(1)) == 4:
            notes.append("4-digit year normalized to 2-digit")
        # Look for a specialization token anywhere in the window.
        spec = None
        for word in ("om", "fk"):
            if word in window:
                spec = _SPEC_WORD[word]
        if spec is None:
            sm = re.search(r"-([flmop])\b", window)
            if sm:
                spec = _SPEC_WORD[sm.group(1)]
        add(f"Media-{yy}-{spec}" if spec else f"Media-{yy}")

    # KP tokens (cohort level only; KP has no specialization split here).
    for m in _KP_RE.finditer(cell):
        yy = _year2(m.group(1))
        if yy is None:
            notes.append(f"ambiguous KP year 'KP{m.group(1)}'")
            continue
        add(f"KP-{yy}")

    # Anything else comma/semicolon separated that we did not recognize.
    for piece in re.split(r"[;,]", cell):
        p = clean_text(piece)
        if not p or _MEDIA_RE.search(p) or _KP_RE.search(p):
            continue
        p = re.sub(r"\(.*?\)", "", p).strip()
        if p and not re.fullmatch(r"(och|med|i|p\d)", p, re.I):
            add(p)
            notes.append(f"unrecognized / other group '{p}'")

    if not groups:
        notes.append("could not extract any group code")
    return groups, notes


def _canon_prog(p):
    p = re.sub(r"[\s\-_]+$", "", (p or "").strip())
    return "Media" if p.lower() == "media" else p.upper()


def normalize_group_code(token, active_years=None):
    """Normalize ONE group token of any programme to CODE-YY (Media keeps its -X
    track), tolerating messy input. Returns (code, note). Examples:
    'Media-2025'->'Media-25', 'media 25'->'Media-25', 'FT-26'->'FT-26',
    'Media-2'->second study year mapped via the active intake years."""
    t = clean_text(token)
    if not t:
        return "", ""
    m = re.match(r"^\s*([A-Za-zÅÄÖåäö][A-Za-zÅÄÖåäö \-_]*?)[\s\-_]*?(\d{1,4})\s*[-_ ]?\s*([A-Za-zÅÄÖåäö]{1,2})?\s*$", t)
    if not m:
        return t, ""
    prog = _canon_prog(m.group(1))
    digits, spec, note = m.group(2), (m.group(3) or "").upper(), ""
    if len(digits) >= 3:
        yy, note = digits[-2:], "4-digit year normalized to 2-digit"
    elif len(digits) == 2:
        yy = digits
    else:                                   # single digit -> study year -> cohort
        yrs = list(active_years or [])
        idx = int(digits) - 1
        if 0 <= idx < len(yrs):
            yy, note = str(yrs[idx]), f"'{t}' read as study-year {digits} -> cohort {prog}-{yrs[idx]}"
        else:
            return t, f"single-digit year in '{t}' could not be mapped to a cohort"
    code = f"{prog}-{yy}"
    if spec and prog == "Media" and spec[0] in "FLMOP":
        code += f"-{spec[0]}"
    return code, note


# --------------------------------------------------------------------------- #
# Teachers
# --------------------------------------------------------------------------- #
def split_people(raw):
    """Split a teacher cell into individual cleaned names."""
    if raw is None:
        return []
    parts = re.split(r"[;,]|\s&\s|\soch\s", str(raw))
    return [clean_text(p) for p in parts if clean_text(p)]


# Phrases that mean the cell holds an instruction/note, not a teacher name —
# e.g. "Tillsammans med Fiktivt berättande i A211" (co-scheduled with another group).
_NOT_A_NAME = re.compile(
    r"tillsamm|gemensam|samläs|samlas|delas|delad|\bsamma\b|\bmed\b.*\b[a-zåäö]{4,}|"
    r"\bi\s+a\d{2,3}\b|\ba\d{3}\b|\bkl\.?\b|\bgrupp|\bsalen?\b|\bse\b|\bobs\b|together|with the",
    re.I)


def is_probably_name(s):
    """True if a teacher-cell token looks like a person's name (not a note/instruction)."""
    s = clean_text(s)
    if not s:
        return False
    if _NOT_A_NAME.search(s):
        return False
    if len(s.split()) > 4:           # a sentence, not a name
        return False
    if sum(c.isdigit() for c in s):  # names don't contain digits
        return False
    return sum(c.isalpha() for c in s) >= 2


# --------------------------------------------------------------------------- #
# Weekday / minutes / week
# --------------------------------------------------------------------------- #
def parse_weekdays(raw):
    """Return (weekday_codes, embedded_dates, notes)."""
    notes = []
    if raw is None or str(raw).strip() == "":
        return [], [], []
    s = str(raw)
    dates = ["{:02d}.{:02d}.{}".format(int(d), int(mo), y) for d, mo, y in _DATE_RE.findall(s)]
    if dates:
        notes.append("explicit date(s) found in weekday cell")
    # weekday RANGE, e.g. "måndag till onsdag" / "Mon-Wed" -> Mon, Tue, Wed
    rm = _DAYRANGE_RE.search(s)
    if rm:
        a, b = _FULL_DAYS[rm.group(1).lower()], _FULL_DAYS[rm.group(2).lower()]
        ia, ib = _WD_ORDER.index(a), _WD_ORDER.index(b)
        if ia <= ib:
            notes.append(f"weekday range '{rm.group(1)}–{rm.group(2)}' expanded to {', '.join(_WD_ORDER[ia:ib + 1])}")
            return _WD_ORDER[ia:ib + 1], dates, notes
    codes = []
    for tok in re.split(r"[,&/]|\soch\s|\s+", s.lower()):
        tok = re.sub(r"[^a-zåäö]", "", tok)
        if not tok:
            continue
        if tok in _WEEKDAY_LOOKUP:
            code = _WEEKDAY_LOOKUP[tok]
            if code not in codes:
                codes.append(code)
    return codes, dates, notes


def parse_minutes(raw):
    """Return (minutes:int|None, time_range:str, notes)."""
    notes = []
    if raw is None or str(raw).strip() == "":
        return None, "", ["missing minutes"]
    if isinstance(raw, (int, float)):
        return int(raw), "", []
    s = str(raw)
    tr = _TIMERANGE_RE.search(s)
    time_range = f"{tr.group(1)}-{tr.group(2)}" if tr else ""
    m = re.search(r"(\d{2,4})", s)
    if not m:
        return None, time_range, ["could not read minutes from text"]
    notes.append("minutes parsed from free text")
    return int(m.group(1)), time_range, notes


_CTIME_RANGE = re.compile(r"(\d{1,2})(?:[.:](\d{2}))?\s*[-–]\s*(\d{1,2})(?:[.:](\d{2}))?")
_CTIME_FROM = re.compile(r"(?:från|from)\s*kl?\.?\s*(\d{1,2})(?:[.:](\d{2}))?", re.I)
_ROOM_RE = re.compile(r"\b([A-FT]\d{3}|B522|A412|A2|A3)\b")


def parse_comment(text):
    """Mine the 'Other comments' free text for planning hints.

    Returns a dict: time_str, slot ('AM'/'PM'/'FULL'/''), hard (bool, e.g.
    'absolut tid'), double_ok (bool, e.g. 'får dubbelbokas'), room (str),
    computer (bool, e.g. 'Dataklass'). Everything is best-effort and optional.
    """
    out = {"time_str": "", "slot": "", "hard": False, "double_ok": False,
           "room": "", "computer": False}
    if not text:
        return out
    s = str(text)
    low = s.lower()

    sh = eh = None
    m = _CTIME_RANGE.search(s)
    if m:
        sh, eh = int(m.group(1)), int(m.group(3))
        if sh <= 24 and eh <= 24:
            out["time_str"] = (f"{m.group(1)}:{m.group(2) or '00'}-"
                               f"{m.group(3)}:{m.group(4) or '00'}")
    if sh is None:
        mf = _CTIME_FROM.search(s)
        if mf:
            sh = int(mf.group(1))
            out["time_str"] = f"från {mf.group(1)}:{mf.group(2) or '00'}"

    if sh is not None:
        if eh is not None and sh < 12 and eh > 13:
            out["slot"] = "FULL"
        else:
            out["slot"] = "AM" if sh < 12 else "PM"
    if not out["slot"]:
        if "förmiddag" in low:
            out["slot"] = "AM"
        elif "eftermiddag" in low:
            out["slot"] = "PM"
        elif "hela dagen" in low or "heldag" in low:
            out["slot"] = "FULL"

    out["hard"] = bool(re.search(r"absolut|exakt|måste|nödvändig|necessary|\bkrav\b", low))
    out["double_ok"] = "dubbelbok" in low
    out["computer"] = bool(re.search(r"dataklass|datasal|computer", low))
    rm = _ROOM_RE.search(s)
    if rm:
        out["room"] = rm.group(1)
    elif re.search(r"\b(teams|online)\b", low):
        out["room"] = "Online/Teams"
    return out


def parse_week(raw):
    """Return (week:int|None, notes)."""
    if raw is None or str(raw).strip() == "":
        return None, ["missing week number"]
    if isinstance(raw, (int, float)):
        return int(raw), []
    m = re.search(r"(\d{1,2})", str(raw))
    if not m:
        return None, [f"unreadable week value '{clean_text(raw)}'"]
    notes = ["week parsed from text"] if clean_text(raw) != m.group(1) else []
    return int(m.group(1)), notes
