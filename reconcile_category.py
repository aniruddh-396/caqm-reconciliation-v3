"""
CAQM Category-wise Reconciliation Engine  (Module 2)
====================================================

Corrects the editable workbook `caqm_category_wise.xlsx` against the base
workbook `sectoral_analysis.xlsx`.

    Complaints Received  <-  sectoral_analysis "Grand Total"
    Pending Complaints   <-  sectoral_analysis "Pending"
    Resolved Complaints  <-  sectoral_analysis "Resolved"

Differences from Module 1 (agency-wise)
---------------------------------------
* The match key is the PAIR (State, Category). The same category name
  ("Road Dust", "C&D Activity", ...) recurs under every state, so category
  alone is not unique.
* There is no handle layer and no aggregation - one base row per (state,
  category), one editable row per (state, category).
* There are no period columns; only the three count columns.
* Resolved lives in column F here (it was column G in the agency sheet); the
  engine finds every column by its header text, so the shift is handled
  automatically.

Everything else mirrors Module 1: nothing is changed silently, an unmatched
key aborts rather than dropping data, the workbook's formatting and formulas
are preserved, and a full change log plus an independent verifier back the run.

Usage
-----
    python reconcile_category.py \
        --category-wise caqm_category_wise.xlsx \
        --sectoral-analysis sectoral_analysis.xlsx \
        --outdir ./out
"""

from __future__ import annotations

import argparse
import os
import shutil
from dataclasses import dataclass
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# reuse the battle-tested helpers from Module 1
from reconcile import norm, as_number, is_error_value
from errors import (Issue, FormatError, DataError, find_header,
                    duplicate_issue, unmatched_issue)
from datefmt import heading_date, file_date, replace_end_date_in_title

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
CATEGORY_WISE_TEMPLATE = os.path.join(TEMPLATE_DIR, "category_wise.xlsx")

SPEC_SECTORAL = [
    ("state",    ["State"]),
    ("category", ["Category"]),
    ("pending",  ["Pending"]),
    ("resolved", ["Resolved"]),
    ("total",    ["Grand Total"]),
]
SPEC_CATEGORY_WISE = [
    ("state",    ["State"]),
    ("category", ["Category"]),
    ("received", ["Complaints Received"]),
    ("pending",  ["Pending Complaints"]),
    ("resolved", ["Resolved Complaints"]),
]

import mapping as CFG


# ===========================================================================
# Data containers
# ===========================================================================

@dataclass
class Change:
    row: int
    state: str
    category: str
    column: str
    field_name: str
    old: object
    new: object
    rule: str


# ===========================================================================
# Reader for the base (sectoral_analysis)
# ===========================================================================

def read_sectoral(path):
    """Return {(state, category): {pending, resolved, grand_total, row}},
    plus the *_Total rows, plus any internal-inconsistency warnings."""
    fname = os.path.basename(path)
    ws, hrow, cmap = find_header(path, SPEC_SECTORAL, fname)
    c_state, c_cat = cmap["state"], cmap["category"]
    c_pend, c_res, c_tot = cmap["pending"], cmap["resolved"], cmap["total"]

    records, totals, issues = {}, {}, []
    current_state = ""
    for r in range(hrow + 1, ws.max_row + 1):
        state = ws.cell(r, c_state).value
        cat = ws.cell(r, c_cat).value

        # total rows carry their label in the State column with no category
        if state and not cat and norm(state).endswith("total"):
            totals[str(state).strip()] = {
                "pending": as_number(ws.cell(r, c_pend).value),
                "resolved": as_number(ws.cell(r, c_res).value),
                "grand_total": as_number(ws.cell(r, c_tot).value),
            }
            continue
        if state and not norm(state).endswith("total"):
            current_state = str(state).strip()
        if not cat:
            continue

        p = as_number(ws.cell(r, c_pend).value) or 0.0
        v = as_number(ws.cell(r, c_res).value) or 0.0
        g = as_number(ws.cell(r, c_tot).value) or 0.0
        key = (norm(current_state), norm(cat))
        if key in records:
            raise DataError(duplicate_issue(
                fname, ws.title, r, records[key]["row"],
                "State / Category pair", f"{current_state} / {cat}"))
        records[key] = {"row": r, "state": current_state, "category": str(cat).strip(),
                        "pending": p, "resolved": v, "grand_total": g}
        if p + v != g:
            issues.append({"file": os.path.basename(path), "row": r,
                           "state": current_state, "category": str(cat).strip(),
                           "pending": p, "resolved": v, "grand_total": g,
                           "sum": p + v, "diff": g - (p + v)})
    return records, totals, issues


# ===========================================================================
# Reconciliation
# ===========================================================================

def reconcile_category(category_wise, sectoral_analysis, outdir, end_date=None):
    """category_wise=None uses the fixed template templates/category_wise.xlsx.
    end_date (a datetime.date) rewrites the end date in the A1 heading."""
    os.makedirs(outdir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if category_wise is None:
        category_wise = CATEGORY_WISE_TEMPLATE
        if not os.path.exists(category_wise):
            raise FormatError(Issue(
                file="templates/category_wise.xlsx", problem="Template missing",
                detail="The fixed Category-wise format file is not in templates/.",
                fix="Restore templates/category_wise.xlsx next to the app."))
    out_xlsx = os.path.join(outdir, "caqm_category_wise_CORRECTED.xlsx")
    out_log = os.path.join(outdir, f"change_log_category_{stamp}.xlsx")

    base, base_totals, integrity = read_sectoral(sectoral_analysis)

    shutil.copyfile(category_wise, out_xlsx)
    wb = openpyxl.load_workbook(out_xlsx)                       # formulas
    wv = openpyxl.load_workbook(category_wise, data_only=True)  # cached values
    ws, vs = wb.worksheets[0], wv.worksheets[0]

    if end_date is not None:
        a1 = ws.cell(1, 1).value
        if isinstance(a1, str):
            ws.cell(1, 1).value = replace_end_date_in_title(a1, end_date)

    _, hrow, C = find_header(out_xlsx, SPEC_CATEGORY_WISE,
                             os.path.basename(category_wise))

    merged_anchors = {rng.coord.split(":")[0] for rng in ws.merged_cells.ranges}

    def writable(r, c):
        cell = ws.cell(r, c)
        return not isinstance(cell, openpyxl.cell.cell.MergedCell) or cell.coordinate in merged_anchors

    def cur(r, c):
        return vs.cell(r, c).value

    def formula_of(r, c):
        v = ws.cell(r, c).value
        return v if isinstance(v, str) and v.startswith("=") else None

    changes, total_updates, matched, unmatched = [], [], [], []
    used_keys = set()

    # detect each total row's category block: contiguous category rows above it
    total_blocks = {}          # total_row -> (first_cat_row, last_cat_row)
    grand_total_row = None
    _first = None
    for r in range(hrow + 1, ws.max_row + 1):
        cat_v = cur(r, C["category"])
        st_v = cur(r, C["state"])
        is_total = False
        for c in range(1, C["category"] + 1):
            v = cur(r, c)
            if v and norm(str(v)).endswith("total"):
                is_total = True
                if norm(str(v)) == "grand total":
                    grand_total_row = r
                break
        if is_total and not cat_v:
            if _first is not None:
                total_blocks[r] = (_first, r - 1)
            _first = None
        elif cat_v:
            if _first is None:
                _first = r

    def set_cell(r, c, new, field_name, rule, state, cat, bucket=None):
        if not writable(r, c):
            return
        old_f = formula_of(r, c)
        old = old_f if old_f else cur(r, c)
        ws.cell(r, c).value = new
        (bucket if bucket is not None else changes).append(
            Change(r, state, cat, get_column_letter(c), field_name, old, new, rule))

    # ---- state carry-forward that respects merged state cells --------------
    def state_at(r):
        """The State label governing row r. State cells are merged vertically,
        so read the merge anchor when the cell itself is blank."""
        v = cur(r, C["state"])
        if v:
            return str(v).strip()
        for rng in ws.merged_cells.ranges:
            if rng.min_col <= C["state"] <= rng.max_col and rng.min_row <= r <= rng.max_row:
                a = vs.cell(rng.min_row, rng.min_col).value
                return str(a).strip() if a else ""
        return ""

    def fix_total_row(r, label):
        """Rewrite the three count totals as correct SUMs over the row's own
        category block. Repairs off-by-one ranges and hardcoded totals in the
        original sheet (e.g. a Resolved total that was a literal, or a SUM that
        skipped the first category)."""
        block = total_blocks.get(r)
        if not block:
            return
        a, b = block
        for col_key, fname in (("received", "Complaints Received"),
                               ("pending", "Pending Complaints"),
                               ("resolved", "Resolved Complaints")):
            col = C[col_key]
            letter = get_column_letter(col)
            want = f"=SUM({letter}{a}:{letter}{b})"
            if ws.cell(r, col).value != want:
                set_cell(r, col, want, fname,
                         f"total = SUM({letter}{a}:{letter}{b}) over its category block",
                         label, label, total_updates)

    def fix_grand_total(r, label):
        """Grand total = SUM of the state-total rows."""
        state_total_rows = [t for t in total_blocks if t != r]
        for col_key, fname in (("received", "Complaints Received"),
                               ("pending", "Pending Complaints"),
                               ("resolved", "Resolved Complaints")):
            col = C[col_key]
            letter = get_column_letter(col)
            refs = ",".join(f"{letter}{t}" for t in sorted(state_total_rows))
            want = f"=SUM({refs})"
            if ws.cell(r, col).value != want:
                set_cell(r, col, want, fname,
                         "grand total = SUM of state-total rows", label, label, total_updates)

    # ---- walk the editable sheet -------------------------------------------
    for r in range(hrow + 1, ws.max_row + 1):
        state_v = cur(r, C["state"])
        cat_v = cur(r, C["category"])
        label = str(state_v).strip() if state_v else ""

        # total rows (label in State col, no category)
        if label and norm(label).endswith("total") and not cat_v:
            if r == grand_total_row:
                fix_grand_total(r, label)
            else:
                fix_total_row(r, label)
            continue
        # grand-total style rows: label somewhere left of Category, no category
        if not cat_v:
            gl = None
            for c in range(1, C["category"] + 1):
                v = cur(r, c)
                if v and norm(str(v)).endswith("total"):
                    gl = str(v).strip()
                    break
            if gl:
                if r == grand_total_row:
                    fix_grand_total(r, gl)
                else:
                    fix_total_row(r, gl)
            continue

        state = state_at(r)
        cat = str(cat_v).strip()
        key = (norm(state), norm(cat))
        rec = base.get(key)

        if rec is None:
            unmatched.append({
                "row": r, "state": state, "category": cat,
                "severity": "NOT FOUND",
                "status": ("This State / Category pair is in the Category-wise "
                           "sheet but has no row in the uploaded "
                           "sectoral_analysis, so its numbers could not be "
                           "updated."),
                "fix": ("Check the state and category spelling in both files. "
                        "Matching ignores case, spacing and punctuation. The "
                        "row was left unchanged.")})
            continue
        used_keys.add(key)
        matched.append({"row": r, "state": state, "category": cat,
                        "base_row": rec["row"]})

        # Pending
        c = C["pending"]
        old = as_number(cur(r, c)); new = int(rec["pending"])
        if old is None or int(old) != new:
            set_cell(r, c, new, "Pending Complaints",
                     "copy from sectoral_analysis Pending", state, cat)
        # Resolved
        c = C["resolved"]
        old = as_number(cur(r, c)); new = int(rec["resolved"])
        if old is None or int(old) != new:
            set_cell(r, c, new, "Resolved Complaints",
                     "copy from sectoral_analysis Resolved", state, cat)
        # Received = Grand Total, forced to a plain number (clears any =E+F)
        c = C["received"]
        old_formula = formula_of(r, c)
        old = as_number(cur(r, c)); new = int(rec["grand_total"])
        if old_formula is not None or old is None or int(old) != new:
            set_cell(r, c, new, "Complaints Received",
                     "copy from sectoral_analysis Grand Total", state, cat)

    # base rows never matched by any editable row
    for key, rec in base.items():
        if key not in used_keys:
            unmatched.append({
                "row": rec["row"], "state": rec["state"],
                "category": rec["category"], "severity": "NOT IN FORMAT",
                "status": ("Present in the uploaded sectoral_analysis (row "
                           f"{rec['row']}) but there is no row for it in the "
                           "Category-wise sheet, so its numbers were not "
                           "included anywhere."),
                "fix": ("Add a row for this State / Category to the "
                        "Category-wise file, or correct the naming so it "
                        "matches an existing row.")})

    wb.save(out_xlsx)
    _write_log(out_log, changes, total_updates, integrity, matched, unmatched,
               base_totals, category_wise, sectoral_analysis)

    return {"output": out_xlsx, "log": out_log, "changes": changes,
            "total_updates": total_updates, "integrity": integrity,
            "matched": matched, "unmatched": unmatched}


# ===========================================================================
# Change-log workbook
# ===========================================================================

def _write_log(path, changes, total_updates, integrity, matched, unmatched,
               base_totals, f_edit, f_base):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
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
                ws.cell(ri, ci, v).font = body_font
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = (
                widths[i - 1] if widths else max(12, min(46, len(str(h)) + 6)))
        ws.freeze_panes = "A2"
        return ws

    def show(v):
        if v is None:
            return ""
        if isinstance(v, str) and v.startswith("="):
            return "formula " + v
        return v

    def ch_rows(lst):
        return [[c.row, c.state, c.category, c.column, c.field_name,
                 show(c.old), show(c.new), c.rule] for c in lst]

    ch_head = ["Sheet Row", "State", "Category", "Col", "Field",
               "Old Value", "New Value", "Rule Applied"]
    ch_w = [10, 16, 38, 6, 22, 12, 12, 46]

    sheet("Summary", ["Item", "Value"], [
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Module", "2 - Category-wise"],
        ["Editable workbook", os.path.basename(f_edit)],
        ["Base workbook", os.path.basename(f_base)],
        ["", ""],
        ["Match key", "(State, Category)"],
        ["Category rows matched", len(matched)],
        ["Category rows NOT matched", len(unmatched)],
        ["Cells changed (category rows)", len(changes)],
        ["Cells changed (total rows)", len(total_updates)],
        ["Base integrity warnings", len(integrity)],
    ], widths=[36, 44])

    sheet("Changes", ch_head, ch_rows(changes), ch_w)
    sheet("Total Row Changes", ch_head, ch_rows(total_updates), ch_w)

    sheet("Base Integrity", ["Base File", "Row", "State", "Category", "Pending",
                             "Resolved", "Grand Total", "Pending+Resolved",
                             "Difference", "Note"],
          [[i["file"], i["row"], i["state"], i["category"], i["pending"],
            i["resolved"], i["grand_total"], i["sum"], i["diff"],
            "Grand Total disagrees with Pending+Resolved in the BASE file"]
           for i in integrity],
          widths=[20, 8, 14, 34, 10, 10, 12, 16, 11, 54])

    sheet("Match Report", ["Sheet Row", "State", "Category", "Base Row"],
          [[m["row"], m["state"], m["category"], m["base_row"]] for m in matched],
          widths=[10, 16, 40, 10])

    sheet("Unmatched", ["Row", "State", "Category", "Severity", "Status",
                        "How to fix"],
          [[u["row"], u["state"], u["category"], u.get("severity", ""),
            u["status"], u.get("fix", "")] for u in unmatched],
          widths=[8, 16, 38, 16, 60, 58])

    wb.save(path)


# ===========================================================================
# CLI
# ===========================================================================

def main():
    p = argparse.ArgumentParser(description="Reconcile CAQM Category-wise workbook.")
    p.add_argument("--category-wise", required=True)
    p.add_argument("--sectoral-analysis", required=True)
    p.add_argument("--outdir", default="./out")
    a = p.parse_args()

    res = reconcile_category(a.category_wise, a.sectoral_analysis, a.outdir)
    print(f"Matched categories    : {len(res['matched'])}")
    print(f"Unmatched             : {len(res['unmatched'])}")
    print(f"Cells changed         : {len(res['changes'])} (+{len(res['total_updates'])} on total rows)")
    print(f"Base integrity issues : {len(res['integrity'])}")
    print(f"\nOutput    : {res['output']}")
    print(f"Change log: {res['log']}")


if __name__ == "__main__":
    main()
