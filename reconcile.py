"""
CAQM Agency-wise Reconciliation Engine
======================================

Compares the editable workbook (Agency_wise.xlsx) against two base workbooks
and rewrites the editable one so every figure agrees with the base.

    Complaints Received / Pending / Resolved   <-  agency_analysis.xlsx
    Avg Resolution Period / Avg Pendency Period <- resolution_analysis.xlsx

Design rules
------------
1. NOTHING is changed silently. Every write lands in the change log, and every
   difference too small to act on lands in the "Within Tolerance" sheet.
2. Aggregation (MCD Delhi's 3 zones, NHAI's 3 states) is driven purely by the
   mapping table - counts SUM, periods MEAN. No hardcoded agency logic.
3. The original workbook's formatting, formulas, merges and column widths are
   preserved; only target cells are touched.
4. The engine fails loudly on an unmapped handle rather than dropping data.

Usage
-----
    python reconcile.py --agency-wise Agency_wise.xlsx \
                        --agency-analysis agency_analysis.xlsx \
                        --resolution-analysis resolution_analysis.xlsx \
                        --outdir ./out
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import mapping as CFG
from errors import (Issue, FormatError, DataError, find_header, open_sheet,
                    duplicate_issue, unmatched_issue, col_letter)
from datefmt import heading_date, file_date, replace_end_date_in_title

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
AGENCY_WISE_TEMPLATE = os.path.join(TEMPLATE_DIR, "agency_wise.xlsx")

# Column sets each workbook must have. Used for format validation so a renamed
# or missing column is reported with its file, row and the exact header wanted.
SPEC_AGENCY_ANALYSIS = [
    ("state",    ["State"]),
    ("agency",   ["Agency"]),
    ("pending",  ["Pending"]),
    ("resolved", ["Resolved"]),
    ("total",    ["Grand Total"]),
]
SPEC_RESOLUTION_ANALYSIS = [
    ("state",  ["State"]),
    ("agency", ["Agency"]),
    ("res",    ["Average Resolution Period (Days)", "Average Resolution Period"]),
    ("pen",    ["Average Pendency Period (Days)", "Average Pendency Period"]),
]
SPEC_AGENCY_WISE = [
    ("state",    ["State"]),
    ("agency",   ["Agency"]),
    ("received", ["Complaints Received"]),
    ("pending",  ["Pending Complaints"]),
    ("resolved", ["Resolved Complaints"]),
    ("res",      ["Average Resolution Period (Days)", "Average Resolution Period"]),
    ("pen",      ["Average Pendency Period (Days)", "Average Pendency Period"]),
]


# ===========================================================================
# Helpers
# ===========================================================================

def norm(text) -> str:
    """Normalise a name for matching: casefold, strip punctuation and spacing."""
    if text is None:
        return ""
    s = str(text).replace("\u00a0", " ")
    s = re.sub(r"[\s\n\r\t]+", " ", s).strip().casefold()
    s = re.sub(r"[.\-_,'\"]+", "", s)
    s = re.sub(r"\s+", " ", s)
    return s


def is_error_value(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip().upper() in {e.upper() for e in CFG.ERROR_VALUES}
    return False


def as_number(v):
    """Return a float, or None if the value is not a usable number."""
    if is_error_value(v):
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace(",", "")
        try:
            return float(t)
        except ValueError:
            return None
    return None


def round_half_up(x: float, dp: int = 0):
    q = Decimal(1).scaleb(-dp)
    d = Decimal(str(x)).quantize(q, rounding=ROUND_HALF_UP)
    return int(d) if dp == 0 else float(d)


def find_header_row(ws, must_contain=("agency",), limit=12) -> int:
    for r in range(1, min(limit, ws.max_row) + 1):
        cells = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if all(any(tok == cell or tok in cell for cell in cells) for tok in must_contain):
            return r
    raise ValueError("Could not locate a header row containing: %s" % (must_contain,))


def header_index(ws, row: int) -> dict:
    """Map normalised header text -> column number."""
    out = {}
    for c in range(1, ws.max_column + 1):
        h = norm(ws.cell(row, c).value)
        if h:
            out.setdefault(h, c)
    return out


def pick_col(hdr: dict, *candidates) -> int:
    for cand in candidates:
        n = norm(cand)
        if n in hdr:
            return hdr[n]
    for cand in candidates:
        n = norm(cand)
        for h, c in hdr.items():
            if n in h or h in n:
                return c
    raise KeyError("None of these columns were found: %s (have: %s)"
                   % (list(candidates), list(hdr)))


# ===========================================================================
# Data containers
# ===========================================================================

@dataclass
class BaseRecord:
    """One agency after aggregation across all of its handles."""
    agency: str
    handles: list = field(default_factory=list)
    states: list = field(default_factory=list)
    pending: float = 0.0
    resolved: float = 0.0
    grand_total: float = 0.0
    res_values: list = field(default_factory=list)   # avg resolution, per handle
    pen_values: list = field(default_factory=list)   # avg pendency, per handle
    seen_in_counts: bool = False
    seen_in_periods: bool = False

    @property
    def avg_resolution(self):
        if not self.res_values:
            return None
        return round_half_up(sum(self.res_values) / len(self.res_values),
                             CFG.PERIOD_DECIMALS)

    @property
    def avg_pendency(self):
        if not self.pen_values:
            return None
        return round_half_up(sum(self.pen_values) / len(self.pen_values),
                             CFG.PERIOD_DECIMALS)

    @property
    def is_aggregate(self) -> bool:
        return len(self.handles) > 1


@dataclass
class Change:
    row: int
    state: str
    agency: str
    column: str
    field_name: str
    old: object
    new: object
    rule: str
    sources: str


# ===========================================================================
# Readers
# ===========================================================================

def read_agency_analysis(path):
    """-> (records_by_handle_occurrence, total_rows, integrity_issues)"""
    fname = os.path.basename(path)
    ws, hrow, cmap = find_header(path, SPEC_AGENCY_ANALYSIS, fname)
    c_state, c_agency = cmap["state"], cmap["agency"]
    c_pend, c_res, c_tot = cmap["pending"], cmap["resolved"], cmap["total"]
    seen_handles = {}

    rows, totals, issues = [], {}, []
    current_state = ""
    for r in range(hrow + 1, ws.max_row + 1):
        state = ws.cell(r, c_state).value
        handle = ws.cell(r, c_agency).value
        if state:
            if not handle and norm(state).endswith("total"):
                totals[str(state).strip()] = {
                    "pending": as_number(ws.cell(r, c_pend).value),
                    "resolved": as_number(ws.cell(r, c_res).value),
                    "grand_total": as_number(ws.cell(r, c_tot).value),
                }
                continue
            current_state = str(state).strip()
        if not handle:
            continue

        p = as_number(ws.cell(r, c_pend).value) or 0.0
        v = as_number(ws.cell(r, c_res).value) or 0.0
        g = as_number(ws.cell(r, c_tot).value) or 0.0
        hkey = norm(handle)
        if hkey in seen_handles and hkey != norm("@NHAI_Official"):
            raise DataError(duplicate_issue(fname, ws.title, r,
                                            seen_handles[hkey], "Agency handle",
                                            str(handle).strip()))
        seen_handles.setdefault(hkey, r)
        rows.append({"row": r, "state": current_state, "handle": str(handle).strip(),
                     "pending": p, "resolved": v, "grand_total": g})
        if p + v != g:
            issues.append({"file": os.path.basename(path), "row": r,
                           "state": current_state, "handle": str(handle).strip(),
                           "pending": p, "resolved": v, "grand_total": g,
                           "sum": p + v, "diff": g - (p + v)})
    return rows, totals, issues


def read_resolution_analysis(path):
    """-> (rows, total_rows)"""
    fname = os.path.basename(path)
    ws, hrow, cmap = find_header(path, SPEC_RESOLUTION_ANALYSIS, fname)
    c_state, c_agency = cmap["state"], cmap["agency"]
    c_res, c_pen = cmap["res"], cmap["pen"]

    rows, totals = [], {}
    current_state = ""
    for r in range(hrow + 1, ws.max_row + 1):
        state = ws.cell(r, c_state).value
        handle = ws.cell(r, c_agency).value
        if state:
            if not handle and norm(state).endswith("total"):
                totals[str(state).strip()] = {
                    "res": as_number(ws.cell(r, c_res).value),
                    "pen": as_number(ws.cell(r, c_pen).value),
                }
                continue
            current_state = str(state).strip()
        if not handle:
            continue
        rows.append({"row": r, "state": current_state, "handle": str(handle).strip(),
                     "res": as_number(ws.cell(r, c_res).value),
                     "pen": as_number(ws.cell(r, c_pen).value)})
    return rows, totals


# ===========================================================================
# Aggregation
# ===========================================================================

def build_base_records(count_rows, period_rows):
    """Fold every handle into its target agency. Returns (records, unmapped)."""
    lookup = {norm(h): a for h, a in CFG.HANDLE_TO_AGENCY.items()}
    records, unmapped = {}, []

    def target(handle, src, row):
        a = lookup.get(norm(handle))
        if a is None:
            unmapped.append({"file": src, "row": row, "handle": handle})
        return a

    for r in count_rows:
        a = target(r["handle"], "agency_analysis", r["row"])
        if a is None:
            continue
        rec = records.setdefault(a, BaseRecord(agency=a))
        rec.seen_in_counts = True
        rec.pending += r["pending"]
        rec.resolved += r["resolved"]
        rec.grand_total += r["grand_total"]
        if r["handle"] not in rec.handles:
            rec.handles.append(r["handle"])
        if r["state"] and r["state"] not in rec.states:
            rec.states.append(r["state"])

    for r in period_rows:
        a = target(r["handle"], "resolution_analysis", r["row"])
        if a is None:
            continue
        rec = records.setdefault(a, BaseRecord(agency=a))
        rec.seen_in_periods = True
        if r["res"] is not None:
            rec.res_values.append(r["res"])
        if r["pen"] is not None:
            rec.pen_values.append(r["pen"])
        if r["handle"] not in rec.handles:
            rec.handles.append(r["handle"])
        if r["state"] and r["state"] not in rec.states:
            rec.states.append(r["state"])

    return records, unmapped


# ===========================================================================
# Reconciliation
# ===========================================================================

def reconcile(agency_wise, agency_analysis, resolution_analysis, outdir,
              end_date=None):
    """agency_wise=None uses the fixed template in templates/agency_wise.xlsx.
    end_date (a datetime.date) rewrites the end date in the A1 heading."""
    os.makedirs(outdir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if agency_wise is None:
        agency_wise = AGENCY_WISE_TEMPLATE
        if not os.path.exists(agency_wise):
            raise FormatError(Issue(
                file="templates/agency_wise.xlsx", problem="Template missing",
                detail="The fixed Agency-wise format file is not in templates/.",
                fix="Restore templates/agency_wise.xlsx next to the app."))

    out_xlsx = os.path.join(outdir, "Agency_wise_CORRECTED.xlsx")
    out_log = os.path.join(outdir, f"change_log_{stamp}.xlsx")

    count_rows, count_totals, integrity = read_agency_analysis(agency_analysis)
    period_rows, period_totals = read_resolution_analysis(resolution_analysis)
    records, unmapped = build_base_records(count_rows, period_rows)

    if unmapped:
        raise DataError([Issue(
            file=u["file"] + ".xlsx", row=u["row"], column="Agency",
            problem="Unknown agency handle",
            detail=f"The handle '{u['handle']}' is not in the mapping table, so "
                   f"its numbers would have been dropped.",
            expected="a handle listed in mapping.py (HANDLE_TO_AGENCY)",
            found=u["handle"],
            fix=f"Add a line to mapping.py:  \"{u['handle']}\": \"<Agency name "
                f"exactly as written in the Agency-wise sheet>\",")
            for u in unmapped],
            headline="Some agency handles in the base files are not mapped, so "
                     "nothing was changed (no data is ever dropped silently)")

    # ---- open editable workbook twice: formulas (to write) + values (to read)
    shutil.copyfile(agency_wise, out_xlsx)
    wb = openpyxl.load_workbook(out_xlsx)                     # formulas preserved
    wv = openpyxl.load_workbook(agency_wise, data_only=True)  # cached values
    ws, vs = wb.worksheets[0], wv.worksheets[0]

    # --- heading end date -------------------------------------------------
    if end_date is not None:
        a1 = ws.cell(1, 1).value
        if isinstance(a1, str):
            ws.cell(1, 1).value = replace_end_date_in_title(a1, end_date)

    _, hrow, C = find_header(out_xlsx, SPEC_AGENCY_WISE,
                             os.path.basename(agency_wise))

    merged_anchors = {rng.coord.split(":")[0] for rng in ws.merged_cells.ranges}

    def mirror_received_total(r, label):
        """In base_grand_total mode the per-agency Complaints Received equals the
        base Grand Total, an independent figure. A total row's Received must then
        be the SUM of those cells, not =Pending+Resolved (which disagrees wherever
        the base is inconsistent). We derive the Received-total formula from the
        row's own Pending-total formula by swapping column E->D, guaranteeing it
        aggregates over exactly the same rows."""
        if CFG.RECEIVED_MODE != "base_grand_total":
            return
        e_col = get_column_letter(C["pending"])
        d_col = get_column_letter(C["received"])
        e_cell = ws.cell(r, C["pending"])
        e_formula = e_cell.value if isinstance(e_cell.value, str) and e_cell.value.startswith("=") else None
        if not e_formula or "SUM" not in e_formula.upper():
            return
        want = re.sub(rf"\b{e_col}(\d+)\b", rf"{d_col}\1", e_formula)
        cur_d = ws.cell(r, C["received"]).value
        if cur_d != want:
            set_cell(r, C["received"], want, "Complaints Received",
                     "total = SUM of Grand-Total cells (mirrors Pending-total range)",
                     "agency_analysis", label, label, total_updates)


    def writable(r, c):
        """True unless the cell is a non-anchor member of a merged range."""
        cell = ws.cell(r, c)
        return not isinstance(cell, openpyxl.cell.cell.MergedCell) or cell.coordinate in merged_anchors

    changes, tolerated, matched, unmatched, total_updates = [], [], [], [], []
    used_agencies = set()

    def cur(r, c):
        return vs.cell(r, c).value

    def formula_of(r, c):
        v = ws.cell(r, c).value
        return v if isinstance(v, str) and v.startswith("=") else None

    def set_cell(r, c, new, field_name, rule, rec_sources, state, agency, bucket=None):
        old_f = formula_of(r, c)
        old = old_f if old_f else cur(r, c)
        if not writable(r, c):
            return
        ws.cell(r, c).value = new
        (bucket if bucket is not None else changes).append(
            Change(r, state, agency, get_column_letter(c), field_name, old, new, rule, rec_sources))

    # ---- walk the editable sheet -------------------------------------------
    current_state = ""
    for r in range(hrow + 1, ws.max_row + 1):
        state_v = cur(r, C["state"])
        agency_v = cur(r, C["agency"])
        label = str(state_v).strip() if state_v else ""

        # -- total rows -----------------------------------------------------
        if label and norm(label).endswith("total") and not agency_v:
            mirror_received_total(r, label)
            if CFG.UPDATE_TOTAL_ROWS:
                base_label = CFG.TOTAL_ROW_MAP.get(label)
                t = period_totals.get(base_label) if base_label else None
                if t:
                    _apply_period(set_cell, r, C, cur, t["res"], t["pen"],
                                  label, label, "total row <- base total row",
                                  "resolution_analysis", total_updates)
            continue
        if not agency_v:
            # grand-total row keeps its label in the S.No/merged column
            gl = None
            for c in range(1, C["agency"] + 1):
                v = cur(r, c)
                if v and norm(str(v)).endswith("total"):
                    gl = str(v).strip()
                    break
            if gl:
                mirror_received_total(r, gl)
                if CFG.UPDATE_TOTAL_ROWS:
                    base_label = CFG.TOTAL_ROW_MAP.get(gl)
                    t = period_totals.get(base_label) if base_label else None
                    if t:
                        _apply_period(set_cell, r, C, cur, t["res"], t["pen"],
                                      gl, gl, "total row <- base total row",
                                      "resolution_analysis", total_updates)
            continue

        if state_v and not norm(label).endswith("total"):
            current_state = label

        agency = str(agency_v).strip()
        rec = _find_record(records, agency)

        if rec is None:
            expected = agency in CFG.EXPECTED_MISSING
            unmatched.append({
                "row": r, "state": current_state, "agency": agency,
                "severity": "OK (expected)" if expected else "NOT FOUND",
                "status": ("Known to be absent from the base files - left "
                           "untouched on purpose."
                           if expected else
                           "This agency is in the Agency-wise format but has no "
                           "row in the uploaded base files, so its numbers could "
                           "not be updated."),
                "fix": ("-" if expected else
                        "Check the agency exists in agency_analysis / "
                        "resolution_analysis and that its handle is mapped in "
                        "mapping.py. Its row was left unchanged.")})
            continue

        used_agencies.add(norm(agency))
        src = " + ".join(rec.handles)
        matched.append({"row": r, "state": current_state, "agency": agency,
                        "handles": src, "n_handles": len(rec.handles),
                        "aggregated": "YES (SUM / MEAN)" if rec.is_aggregate else "no",
                        "in_agency_analysis": "yes" if rec.seen_in_counts else "NO",
                        "in_resolution_analysis": "yes" if rec.seen_in_periods else "NO"})

        # -- counts ---------------------------------------------------------
        if rec.seen_in_counts:
            for key, base_val, fname in (
                ("pending", rec.pending, "Pending Complaints"),
                ("resolved", rec.resolved, "Resolved Complaints"),
            ):
                c = C[key]
                old = as_number(cur(r, c))
                new = int(base_val)
                if old is None or int(old) != new:
                    set_cell(r, c, new, fname, f"{'SUM of ' + str(len(rec.handles)) + ' handles' if rec.is_aggregate else 'direct copy'} from agency_analysis",
                             src, current_state, agency)

            c = C["received"]
            if CFG.RECEIVED_MODE == "formula":
                want = "={}{}+{}{}".format(get_column_letter(C["pending"]), r,
                                           get_column_letter(C["resolved"]), r)
                if formula_of(r, c) != want:
                    set_cell(r, c, want, "Complaints Received",
                             "derived: Pending + Resolved (RECEIVED_MODE='formula')",
                             src, current_state, agency)
            else:
                # Complaints Received must be the base Grand Total as a plain
                # number. If the cell still holds a formula (e.g. =E+G) it must
                # be replaced even when its cached value happens to match, or it
                # will silently recompute to Pending+Resolved after recalc.
                old_formula = formula_of(r, c)
                old = as_number(cur(r, c))
                new = int(rec.grand_total)
                if old_formula is not None or old is None or int(old) != new:
                    set_cell(r, c, new, "Complaints Received",
                             f"{'SUM' if rec.is_aggregate else 'copy'} of Grand Total from agency_analysis",
                             src, current_state, agency)

        # -- periods --------------------------------------------------------
        if rec.seen_in_periods:
            rule = (f"MEAN of {len(rec.res_values)} handle values"
                    if rec.is_aggregate else "direct copy")
            _apply_period(set_cell, r, C, cur, rec.avg_resolution, rec.avg_pendency,
                          current_state, agency, rule + " from resolution_analysis",
                          src, None, tolerated)

    # Reverse check: agencies present in the uploaded base files that have no
    # row in the Agency-wise format. Their numbers are simply not carried over,
    # so the user must be told rather than left to notice a missing total.
    for a_name, rec in records.items():
        if norm(a_name) not in used_agencies:
            unmatched.append({
                "row": "-", "state": ", ".join(rec.states) or "-",
                "agency": a_name, "severity": "NOT IN FORMAT",
                "status": ("Present in the uploaded base files (handle(s): "
                           f"{' + '.join(rec.handles)}) but there is no row for "
                           "it in the Agency-wise format, so its numbers were "
                           "not included anywhere."),
                "fix": ("Add a row for this agency to templates/agency_wise.xlsx, "
                        "or correct the name in mapping.py so it points at an "
                        "agency that exists in the sheet.")})

    wb.save(out_xlsx)

    _write_log(out_log, changes, tolerated, total_updates, integrity,
               matched, unmatched, records,
               agency_wise, agency_analysis, resolution_analysis)

    return {"output": out_xlsx, "log": out_log, "changes": changes,
            "tolerated": tolerated, "total_updates": total_updates,
            "integrity": integrity, "matched": matched, "unmatched": unmatched,
            "records": records}


def _find_record(records, agency):
    if agency in records:
        return records[agency]
    n = norm(agency)
    for k, v in records.items():
        if norm(k) == n:
            return v
    return None


def _apply_period(set_cell, r, C, cur, base_res, base_pen, state, agency,
                  rule, src, bucket, tol_bucket=None):
    """Write the two period columns, honouring tolerance and N/A semantics."""
    for key, base_val, fname in (("res", base_res, "Average Resolution Period (Days)"),
                                 ("pen", base_pen, "Average Pendency Period (Days)")):
        c = C[key]
        old_raw = cur(r, c)
        old = as_number(old_raw)

        if base_val is None:
            if not is_error_value(old_raw):
                set_cell(r, c, CFG.NA_TEXT, fname,
                         "base value not computable -> N/A", src, state, agency, bucket)
            continue

        if old is None:
            set_cell(r, c, base_val, fname, rule, src, state, agency, bucket)
            continue

        diff = abs(old - base_val)
        if diff >= CFG.PERIOD_TOLERANCE_DAYS:
            set_cell(r, c, base_val, fname, rule, src, state, agency, bucket)
        elif diff > 0 and tol_bucket is not None:
            tol_bucket.append(Change(r, state, agency, get_column_letter(c), fname,
                                     old, base_val,
                                     f"difference {diff:g} d < tolerance "
                                     f"{CFG.PERIOD_TOLERANCE_DAYS} d - NOT changed", src))


# ===========================================================================
# Change log workbook
# ===========================================================================

def _write_log(path, changes, tolerated, total_updates, integrity,
               matched, unmatched, records, f_edit, f_count, f_period):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    head_font = Font(name="Arial", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F4E79")
    body_font = Font(name="Arial")

    def sheet(title, headers, rows, widths=None):
        ws = wb.create_sheet(title[:31])
        for i, h in enumerate(headers, 1):
            cell = ws.cell(1, i, h)
            cell.font, cell.fill = head_font, head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for ri, row in enumerate(rows, 2):
            for ci, v in enumerate(row, 1):
                cell = ws.cell(ri, ci, v)
                cell.font = body_font
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = (
                widths[i - 1] if widths else max(12, min(46, len(str(h)) + 6)))
        ws.freeze_panes = "A2"
        return ws

    def show(v):
        """Render a value for the log. A leading '=' would otherwise be stored
        as a live formula in the log workbook and read back as blank."""
        if v is None:
            return ""
        if isinstance(v, str) and v.startswith("="):
            return "formula " + v
        return v

    def ch_rows(lst):
        return [[c.row, c.state, c.agency, c.column, c.field_name,
                 show(c.old), show(c.new), c.rule, c.sources] for c in lst]

    # ---- heuristic: which changes deserve a human eye ----------------------
    def suspicious(c):
        if "Period" not in c.field_name:
            return None
        o, n = as_number(c.old), as_number(c.new)
        if o is None or n is None or o == 0 or n == 0:
            return None
        hi, lo = max(o, n), min(o, n)
        if 9.0 < hi / lo < 13.0 and str(int(hi)).startswith(str(int(lo))[:2]):
            return ("POSSIBLE LOST DECIMAL POINT - the two values share their "
                    "leading digits and differ by roughly 10x. One of the two "
                    "workbooks may have dropped a decimal separator. Confirm "
                    "against the raw complaint data before accepting.")
        if "Pendency" in c.field_name and abs(n - o) == 1:
            return ("OFF-BY-ONE - part of a systematic +1 day shift across most "
                    "agencies, consistent with inclusive vs exclusive day "
                    "counting between the two systems. Decide the convention "
                    "once, then re-run.")
        return None

    review = [(c, s) for c in changes if (s := suspicious(c))]

    ch_head = ["Sheet Row", "State", "Agency", "Col", "Field",
               "Old Value", "New Value", "Rule Applied", "Base Source Handle(s)"]
    ch_w = [10, 16, 40, 6, 26, 14, 14, 44, 46]

    sheet("Summary", ["Item", "Value"], [
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Editable workbook", os.path.basename(f_edit)],
        ["Base - counts", os.path.basename(f_count)],
        ["Base - periods", os.path.basename(f_period)],
        ["", ""],
        ["RECEIVED_MODE", CFG.RECEIVED_MODE],
        ["Period tolerance (days)", CFG.PERIOD_TOLERANCE_DAYS],
        ["Period decimals", CFG.PERIOD_DECIMALS],
        ["Totals rows refreshed", str(CFG.UPDATE_TOTAL_ROWS)],
        ["", ""],
        ["Agency rows matched", len(matched)],
        ["Agency rows NOT matched", len(unmatched)],
        ["Cells changed (agency rows)", len(changes)],
        ["Cells changed (total rows)", len(total_updates)],
        ["Differences within tolerance (left alone)", len(tolerated)],
        ["Base integrity warnings", len(integrity)],
        ["Changes flagged for manual review", len(review)],
    ], widths=[42, 40])

    sheet("Review Manually", ["Sheet Row", "State", "Agency", "Col", "Field",
                              "Old Value", "New Value (applied)", "Why this needs review"],
          [[c.row, c.state, c.agency, c.column, c.field_name,
            show(c.old), show(c.new), why] for c, why in review],
          widths=[10, 16, 40, 6, 26, 12, 18, 76])

    sheet("Changes", ch_head, ch_rows(changes), ch_w)
    sheet("Total Row Changes", ch_head, ch_rows(total_updates), ch_w)
    sheet("Within Tolerance", ch_head, ch_rows(tolerated), ch_w)

    sheet("Base Integrity", ["Base File", "Row", "State", "Handle", "Pending",
                             "Resolved", "Grand Total", "Pending+Resolved",
                             "Difference", "Note"],
          [[i["file"], i["row"], i["state"], i["handle"], i["pending"],
            i["resolved"], i["grand_total"], i["sum"], i["diff"],
            "Grand Total disagrees with Pending+Resolved in the BASE file"]
           for i in integrity],
          widths=[22, 8, 16, 26, 11, 11, 13, 17, 12, 56])

    sheet("Aggregation", ["Agency (combined)", "Handles merged", "Count rule",
                          "Pending (sum)", "Resolved (sum)", "Grand Total (sum)",
                          "Period rule", "Avg Resolution", "Avg Pendency",
                          "Source values (resolution)", "Source values (pendency)"],
          [[rec.agency, " + ".join(rec.handles), "SUM",
            rec.pending, rec.resolved, rec.grand_total, "MEAN",
            rec.avg_resolution, rec.avg_pendency,
            ", ".join(str(int(v)) if float(v).is_integer() else str(v) for v in rec.res_values),
            ", ".join(str(int(v)) if float(v).is_integer() else str(v) for v in rec.pen_values)]
           for rec in records.values() if rec.is_aggregate],
          widths=[40, 60, 11, 14, 14, 16, 12, 15, 15, 30, 30])

    sheet("Match Report", ["Sheet Row", "State", "Agency", "Handle(s)",
                           "# Handles", "Aggregated", "In agency_analysis",
                           "In resolution_analysis"],
          [[m["row"], m["state"], m["agency"], m["handles"], m["n_handles"],
            m["aggregated"], m["in_agency_analysis"], m["in_resolution_analysis"]]
           for m in matched],
          widths=[10, 16, 40, 46, 10, 18, 19, 22])

    sheet("Unmatched", ["Sheet Row", "State", "Agency", "Severity", "Status",
                        "How to fix"],
          [[u["row"], u["state"], u["agency"], u.get("severity", ""),
            u["status"], u.get("fix", "")] for u in unmatched],
          widths=[10, 16, 44, 16, 60, 60])

    wb.save(path)


# ===========================================================================
# CLI
# ===========================================================================

def main():
    p = argparse.ArgumentParser(description="Reconcile CAQM Agency-wise workbook against base analyses.")
    p.add_argument("--agency-wise", required=True)
    p.add_argument("--agency-analysis", required=True)
    p.add_argument("--resolution-analysis", required=True)
    p.add_argument("--outdir", default="./out")
    a = p.parse_args()

    res = reconcile(a.agency_wise, a.agency_analysis, a.resolution_analysis, a.outdir)

    print(f"Matched agencies      : {len(res['matched'])}")
    print(f"Unmatched agencies    : {len(res['unmatched'])}")
    print(f"Cells changed         : {len(res['changes'])} (+{len(res['total_updates'])} on total rows)")
    print(f"Within tolerance      : {len(res['tolerated'])}")
    print(f"Base integrity issues : {len(res['integrity'])}")
    print(f"\nOutput    : {res['output']}")
    print(f"Change log: {res['log']}")


if __name__ == "__main__":
    main()
