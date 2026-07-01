"""Finnish national public holidays — computed offline for any year.

No internet needed: the fixed dates are known and the movable ones (Good Friday,
Easter Monday, Ascension) come from the Easter date via the standard computus.
Only weekday holidays matter for a Mon–Fri planner, but we return them all.
"""
from __future__ import annotations

import datetime


def _easter(year):
    """Gregorian Easter Sunday (anonymous computus)."""
    a = year % 19
    b, c = year // 100, year % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return datetime.date(year, month, day)


def _friday_between(year, month, d1, d2):
    for d in range(d1, d2 + 1):
        dt = datetime.date(year, month, d)
        if dt.weekday() == 4:            # Friday
            return dt
    return None


def finnish_holidays(year):
    """{ 'YYYY-MM-DD': name } for one calendar year."""
    e = easter = _easter(year)
    out = {}

    def add(d, name):
        if d:
            out[d.isoformat()] = name

    add(datetime.date(year, 1, 1), "Nyårsdagen")
    add(datetime.date(year, 1, 6), "Trettondagen")
    add(e - datetime.timedelta(days=2), "Långfredagen")        # Good Friday
    add(e + datetime.timedelta(days=1), "Annandag påsk")       # Easter Monday
    add(datetime.date(year, 5, 1), "Första maj (Vappu)")
    add(e + datetime.timedelta(days=39), "Kristi himmelsfärdsdag")  # Ascension (Thu)
    add(_friday_between(year, 6, 19, 25), "Midsommarafton")    # Midsummer Eve (Fri)
    add(datetime.date(year, 12, 6), "Självständighetsdagen")
    add(datetime.date(year, 12, 24), "Julafton")
    add(datetime.date(year, 12, 25), "Juldagen")
    add(datetime.date(year, 12, 26), "Annandag jul")
    return out


def holidays_for_academic_year(academic_year):
    """'2026-2027' -> holidays for both calendar years."""
    out = {}
    for part in str(academic_year or "").split("-"):
        try:
            out.update(finnish_holidays(int(part)))
        except ValueError:
            continue
    if not out:                                    # fallback: this year + next
        y = datetime.date.today().year
        out.update(finnish_holidays(y))
        out.update(finnish_holidays(y + 1))
    return out
