"""
Shared date helpers for headings and file names.

Headings carry dates like '16th September 2026'; file names carry '16-09-2026'.
Only the END date of a range is ever rewritten - the fixed/original start date
and any surrounding words (' in U.P', ' in Delhi-NCR', spacing quirks) are
preserved exactly.
"""

import re
from datetime import date, datetime, timedelta, timezone

_DATE_RE = re.compile(r"\d{1,2}(?:st|nd|rd|th)\s+[A-Za-z]+\s+\d{4}")

_MONTHS = {name.lower(): i for i, name in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"], start=1)}
_MONTH_ABBR = {name[:3]: i for name, i in _MONTHS.items()}


def ordinal(n: int) -> str:
    if 11 <= (n % 100) <= 13:
        suf = "th"
    else:
        suf = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"


def heading_date(d: date) -> str:
    """date(2026, 9, 16) -> '16th September 2026'."""
    return f"{ordinal(d.day)} {d.strftime('%B')} {d.year}"


def file_date(d: date) -> str:
    """date(2026, 9, 16) -> '16-09-2026'."""
    return d.strftime("%d-%m-%Y")


def replace_end_date_in_title(title: str, d: date) -> str:
    """Swap only the LAST date-looking token in the title, leaving the fixed
    start date and every other character untouched."""
    matches = list(_DATE_RE.finditer(title))
    if not matches:
        return title
    last = matches[-1]
    return title[:last.start()] + heading_date(d) + title[last.end():]


def first_date_text(text: str):
    """Return the FIRST date-looking substring verbatim (e.g. '01st Jan 2025'),
    preserving whatever zero-padding / month abbreviation the source used.
    Use this when the original phrasing must be kept; use parse_ordinal_date
    when a computable date is needed instead."""
    m = _DATE_RE.search(text)
    return m.group() if m else None


def parse_ordinal_date(text: str):
    """'01st Jan 2025' / '15th October 2021' -> date(...), using the FIRST
    date-looking token in the string. Returns None if none is found or the
    month name is not recognised - the caller decides how to report that."""
    m = _DATE_RE.search(text)
    if not m:
        return None
    dm = re.match(r"(\d{1,2})(?:st|nd|rd|th)\s+([A-Za-z]+)\s+(\d{4})", m.group())
    if not dm:
        return None
    day, mon_txt, year = dm.groups()
    key = mon_txt.strip().lower()
    mon = _MONTHS.get(key) or _MONTH_ABBR.get(key[:3])
    if mon is None:
        return None
    try:
        return date(int(year), mon, int(day))
    except ValueError:
        return None


def today_ist() -> date:
    """Today's date in India (IST, UTC+5:30).

    The hosted app runs on a UTC server, where date.today() is still
    'yesterday' for the first 5.5 hours of every Indian day. Computing IST
    directly also avoids needing a timezone database (zoneinfo needs the
    tzdata package on Windows). India has no daylight saving, so a fixed
    offset is exact.
    """
    return (datetime.now(timezone.utc) + timedelta(hours=5, minutes=30)).date()


def parse_heading_dates(text: str):
    """All dates written like '01st Jan 2025' / '17th September 2026' in a
    heading, as datetime.date objects, in the order they appear."""
    out = []
    for m in _DATE_RE.finditer(text):
        day_s, rest = re.match(r"(\d+)(?:st|nd|rd|th)\s+(.*)", m.group(0)).groups()
        mon, yr = rest.split()
        month = None
        for fmt in ("%B", "%b"):
            try:
                month = datetime.strptime(mon, fmt).month
                break
            except ValueError:
                pass
        if month:
            try:
                out.append(date(int(yr), month, int(day_s)))
            except ValueError:
                pass
    return out
