"""
Independent verifier for Module 3 (state-wise).

Re-derives every output cell straight from the parent workbook, without
importing reconcile_state.py, and checks each generated state file:
  * every agency row's Received / Resolved / NotResolved against the parent
    (MCD = East+North+South summed);
  * every % cell = round-half-up(part/whole) or 0 when whole == 0;
  * Sub Total / Total rows = sum of their block; Grand Total = sum of all
    agency rows;
  * the A1 title's end date and the file name both carry the requested date.

Usage:
    python verify_state.py 2025.xlsx <end YYYY-MM-DD> <out_dir>
"""
import glob
import os
import re
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

PARENT, END, OUTDIR = sys.argv[1], sys.argv[2], sys.argv[3]
end_dt = datetime.strptime(END, "%Y-%m-%d").date()
ERR = {"#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?"}

C_REC, C_RES, C_RESP, C_NR, C_NRP = 3, 4, 5, 6, 7

AGG = {"municipal corporation of delhi":
       ["east delhi municipal corporation",
        "north delhi municipal corporation",
        "south delhi municipal corporation"]}


def norm(t):
    if t is None:
        return ""
    s = re.sub(r"\s+", " ", str(t).replace("\u00a0", " ").replace("\n", " ")).strip().casefold()
    return re.sub(r"[.\-_,'\"]+", "", s)


def num(v):
    if v is None or (isinstance(v, str) and v.strip().upper() in ERR):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def pct(part, whole):
    if whole == 0:
        return 0
    return f"({int(Decimal(str(part / whole * 100)).quantize(Decimal(1), rounding=ROUND_HALF_UP))}%)"


def ordinal(n):
    suf = "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"


# ---- parent facts ----------------------------------------------------------
pws = openpyxl.load_workbook(PARENT, data_only=True).worksheets[0]
parent = {}
for r in range(3, pws.max_row + 1):
    nm = pws.cell(r, 2).value
    if not nm or norm(nm).endswith("total"):
        continue
    parent[norm(nm)] = (int(num(pws.cell(r, C_REC).value) or 0),
                        int(num(pws.cell(r, C_RES).value) or 0),
                        int(num(pws.cell(r, C_NR).value) or 0))


def parent_counts(name_norm):
    if name_norm in AGG:
        t = [0, 0, 0]
        for p in AGG[name_norm]:
            a, b, c = parent[p]
            t[0] += a; t[1] += b; t[2] += c
        return tuple(t)
    return parent.get(name_norm)


expected_end_heading = f"{ordinal(end_dt.day)} {end_dt.strftime('%B')} {end_dt.year}"
expected_end_file = end_dt.strftime("%d-%m-%Y")

fails, checked, files = [], 0, 0

for path in sorted(glob.glob(os.path.join(OUTDIR, "*.xlsx"))):
    files += 1
    fn = os.path.basename(path)
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]

    # filename carries the end date
    checked += 1
    if expected_end_file not in fn:
        fails.append(f"{fn}: filename missing end date {expected_end_file}")
    checked += 1
    if f"01-01-2025 - {expected_end_file}" not in fn:
        fails.append(f"{fn}: filename date range wrong")

    # title carries the end date, once, as the last date token
    a1 = ws.cell(1, 1).value or ""
    checked += 1
    dates = re.findall(r"\d{1,2}(?:st|nd|rd|th)\s+[A-Za-z]+\s+\d{4}", a1)
    if not dates or dates[-1] != expected_end_heading:
        fails.append(f"{fn}: title end date is {dates[-1] if dates else None!r}, "
                     f"expected {expected_end_heading!r}")
    checked += 1
    if "01st Jan 2025" not in a1:
        fails.append(f"{fn}: title lost the fixed start date '01st Jan 2025'")

    # header row
    hrow = next(r for r in range(1, 12)
                if any(norm(ws.cell(r, c).value) == "agency" for c in range(1, 9)))

    all_rows, block = [], []
    for r in range(hrow + 1, ws.max_row + 1):
        nm = ws.cell(r, 2).value
        if not nm:
            continue
        n = norm(nm)
        got = (num(ws.cell(r, C_REC).value), num(ws.cell(r, C_RES).value),
               num(ws.cell(r, C_NR).value))

        if n.endswith("total"):
            rows = all_rows if n == "grand total" else block
            exp = (sum(num(ws.cell(x, C_REC).value) or 0 for x in rows),
                   sum(num(ws.cell(x, C_RES).value) or 0 for x in rows),
                   sum(num(ws.cell(x, C_NR).value) or 0 for x in rows))
            checked += 3
            if got != exp:
                fails.append(f"{fn} row {r} [{nm}]: totals {got} != sum {exp}")
            # percentage cells on total rows
            checked += 2
            if ws.cell(r, C_RESP).value != pct(int(exp[1]), int(exp[0])):
                fails.append(f"{fn} row {r} [{nm}]: res% {ws.cell(r, C_RESP).value!r} "
                             f"!= {pct(int(exp[1]), int(exp[0]))!r}")
            if ws.cell(r, C_NRP).value != pct(int(exp[2]), int(exp[0])):
                fails.append(f"{fn} row {r} [{nm}]: notres% wrong")
            block = []
            continue

        exp = parent_counts(n)
        checked += 3
        if exp is None:
            fails.append(f"{fn} row {r}: agency '{nm}' not in parent")
            continue
        if got != tuple(float(x) for x in exp):
            fails.append(f"{fn} row {r} [{nm}]: counts {got} != parent {exp}")
        checked += 2
        if ws.cell(r, C_RESP).value != pct(exp[1], exp[0]):
            fails.append(f"{fn} row {r} [{nm}]: res% {ws.cell(r, C_RESP).value!r} "
                         f"!= {pct(exp[1], exp[0])!r}")
        if ws.cell(r, C_NRP).value != pct(exp[2], exp[0]):
            fails.append(f"{fn} row {r} [{nm}]: notres% {ws.cell(r, C_NRP).value!r} "
                         f"!= {pct(exp[2], exp[0])!r}")
        all_rows.append(r)
        block.append(r)

print(f"files checked      : {files}")
print(f"assertions checked : {checked}")
if fails:
    print(f"FAILURES: {len(fails)}")
    for f in fails:
        print("   x", f)
    sys.exit(1)
print("ALL CHECKS PASSED")
