"""
Module 4 - Social Media Agency-wise Data (Delhi-NCR), MCD consolidated
========================================================================

Takes ONE flat agency-wise workbook (e.g. "2025-2026.xlsx" or
"2021-2025.xlsx") covering the whole Delhi-NCR list - no State column, one
row per agency, a single Total row - and:

  1. Sums East + North + South Delhi Municipal Corporation into ONE row,
     "Municipal Corporation of Delhi".
  2. Places that combined row at Sr. No. 12 in the output (all other agencies
     keep their existing relative order; whichever agency used to be 12th,
     and everyone after it, shifts down by one - net two fewer rows than the
     source, since three rows become one).
  3. Renumbers S. No. 1..N.
  4. Recomputes the Total row (must equal the source total exactly, since
     summing three rows into one changes nothing about the grand totals -
     this is asserted, not just computed).
  5. Produces BOTH a corrected .xlsx and a matching .pdf, both titled from the
     date range already present in the source file's own heading.

Nothing here touches Modules 1-3.

File naming
-----------
Both output files share one fixed-pattern name:

    Social_Media_Status_<START>_<END>.xlsx
    Social_Media_Status_<START>_<END>.pdf

START and END are written DDMMYYYY. START is the report's fixed start date, read
from the source file's own heading (01012025 for the 2025-2026 report, 15102021
for the 2021-2025 report). END is the end date the user chose.
    e.g.  Social_Media_Status_01012025_26102026.xlsx
          Social_Media_Status_15102021_26102026.pdf

Headings (these are NOT the file name):
  Excel A1 : "Social Media Agency-wise Data (01st Jan 2025 - 26th October 2026) in Delhi-NCR"
  PDF title: "Social Media Agency-wise Data (01st Jan 2025 - 26th October 2026)"

An optional end_date lets the caller override the END date shown, exactly
like the other modules; leave it None to use the date already in the source.
"""

from __future__ import annotations

import os
import re
from copy import copy
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.utils import get_column_letter

from errors import Issue, FormatError, DataError, find_header, open_sheet
from datefmt import replace_end_date_in_title, parse_heading_dates, heading_date

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

MCD_ZONES = [
    "East Delhi Municipal Corporation",
    "North Delhi Municipal Corporation",
    "South Delhi Municipal Corporation",
]
MCD_COMBINED_NAME = "Municipal Corporation of Delhi"
MCD_TARGET_POSITION = 12          # 1-based Sr. No. the combined row must land at

HEADING_SUFFIX = " in Delhi-NCR"

SPEC_NCR = [
    ("sno",     ["S. No.", "S.No.", "Sr. No."]),
    ("agency",  ["Agency"]),
    ("received", ["Total No. of Complaints Received", "Complaints Received"]),
    ("resolved", ["Total No. of Complaints Resolved(% Resolved)",
                  "Total No. of Complaints Resolved", "Complaints Resolved"]),
    ("notres",  ["Total No. of Complaints Not Resolved(% Not Resolved)",
                 "Total No. of Complaints Not Resolved", "Complaints Not Resolved"]),
]


def norm(t) -> str:
    if t is None:
        return ""
    s = re.sub(r"[\s\n\r\t]+", " ", str(t).replace("\u00a0", " ")).strip().casefold()
    s = re.sub(r"[.\-_,'\"]+", "", s)
    return re.sub(r"\s+", " ", s)


STEM_PREFIX = "Social_Media_Status"


def _split_suffix(heading: str):
    """-> (Excel heading ending in ' in Delhi-NCR', PDF heading without it)."""
    heading = heading.strip()
    if heading.endswith(HEADING_SUFFIX):
        return heading, heading[: -len(HEADING_SUFFIX)].strip()
    return heading + HEADING_SUFFIX, heading


def file_stem(heading: str, source_label: str = "") -> str:
    """'Social_Media_Status_<start DDMMYYYY>_<end DDMMYYYY>' from the heading's
    first and last dates."""
    found = parse_heading_dates(heading)
    if len(found) < 2:
        raise DataError(Issue(
            file=source_label, sheet="", row=1, column="A",
            problem="Cannot build the file name",
            detail="The heading needs a start date and an end date, like "
                   "'(01st Jan 2025 - 26th October 2026)', to name the output files.",
            expected="two dates in cell A1", found=heading[:80],
            fix="Add the date range to A1 of the source file, then run again."))
    return f"{STEM_PREFIX}_{found[0]:%d%m%Y}_{found[-1]:%d%m%Y}"


def preview_names(file_like, end_date=None):
    """Read ONLY cell A1 of an uploaded workbook and report what a run would
    produce: (excel_heading, pdf_heading, file_stem). Returns None if the heading
    cannot be read - this is a convenience preview, never a gate."""
    try:
        wb = openpyxl.load_workbook(file_like, read_only=True, data_only=True)
        a1 = next(wb.worksheets[0].iter_rows(min_row=1, max_row=1, values_only=True))[0]
        wb.close()
        if not isinstance(a1, str) or not a1.strip():
            return None
        heading = a1.strip()
        if end_date is not None:
            if not parse_heading_dates(heading):
                return None
            heading = replace_end_date_in_title(heading, end_date)
        excel_heading, pdf_heading = _split_suffix(heading)
        return excel_heading, pdf_heading, file_stem(pdf_heading)
    except Exception:
        return None


def ensure_pdf_support():
    """The PDF needs the 'reportlab' package. Check FIRST so a missing library
    gives a clear instruction instead of a traceback - and so nothing (not even
    the Excel) is written when the PDF cannot be made."""
    try:
        import reportlab  # noqa: F401
    except ImportError:
        raise FormatError(Issue(
            file="(this computer's Python)", problem="PDF library not installed",
            detail="The 'reportlab' package is needed to create the PDF, but it is "
                   "not installed for the Python that is running this app.",
            expected="reportlab installed", found="not installed",
            fix="Stop the app (Ctrl+C in the terminal). In the app folder run:  "
                "python -m pip install -r requirements.txt   then start the app again."),
            headline="The PDF library (reportlab) is not installed, so nothing was generated")


def pct(part: int, whole: int):
    """'(NN%)' rounded half-up. When nothing was received there is no
    percentage to show, so the cell shows 0 (never an Excel #DIV/0! error)."""
    if whole == 0:
        return 0
    p = int(Decimal(str(part / whole * 100)).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    return f"({p}%)"


def pct_formula(part_cell: str, whole_cell: str) -> str:
    """Excel formula for the same rule, so the workbook also shows 0, not #DIV/0!."""
    return (f'=IF({whole_cell}=0,0,"("&ROUND(({part_cell}/{whole_cell}*100),0)&"%)")')


# ---------------------------------------------------------------------------
# Data containers
# ---------------------------------------------------------------------------

@dataclass
class AgencyRow:
    name: str
    received: int
    resolved: int
    notres: int
    src_row: int


# ---------------------------------------------------------------------------
# Reading + validation
# ---------------------------------------------------------------------------

def _as_int(v, ws, row, col, fname):
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(round(v))
    if isinstance(v, str) and v.strip() == "":
        return 0
    raise DataError(Issue(
        file=fname, sheet=ws.title, row=row, column=get_column_letter(col),
        problem="Value is not a number",
        detail="This cell was expected to hold a count but contains text.",
        expected="a number", found=repr(v)[:60],
        fix="Correct this cell in the source file, then run again."))


def read_source(path):
    fname = os.path.basename(path)
    ws, hrow, cmap = find_header(path, SPEC_NCR, fname)
    warnings = []

    # Formulas are read twice: 'ws' holds the cached values, 'wf' the formulas.
    # A formula whose value was never calculated (file produced by a script and
    # never opened in Excel) would otherwise be read as an empty cell and
    # silently counted as 0 - a wrong report with no warning.
    wf = openpyxl.load_workbook(path).worksheets[0]

    a1 = ws.cell(1, 1).value
    if not isinstance(a1, str) or not a1.strip():
        raise FormatError(Issue(
            file=fname, sheet=ws.title, row=1, column="A",
            problem="Missing heading",
            detail="Cell A1 should contain the report title with its date range.",
            expected="a heading like 'Social Media Agency-wise Data "
                     "(<start> - <end>) in Delhi-NCR'",
            fix="Add the heading to A1, or supply an end date to generate one."))

    rows, seen, total_row = [], {}, None
    for r in range(hrow + 1, ws.max_row + 1):
        name = ws.cell(r, cmap["agency"]).value
        if name is None:
            continue
        nname = norm(name)

        if nname.endswith("total"):
            if nname != "total" and nname != "grand total":
                raise FormatError(Issue(
                    file=fname, sheet=ws.title, row=r, column="B",
                    problem="Wrong kind of file",
                    detail=f"Found a '{str(name).strip()}' row. Sub Total rows "
                           "belong to the state-wise reports (Delhi, UP, ...). "
                           "This module needs the single flat Delhi-NCR list "
                           "with ONE Total row at the bottom.",
                    expected="a flat agency list ending in a single 'Total' row",
                    found=f"'{str(name).strip()}' on row {r}",
                    fix="Upload the whole Delhi-NCR workbook (for example the "
                        "2025-2026 or 2021-2025 file), not a state file."))
            if total_row is not None:
                raise FormatError(Issue(
                    file=fname, sheet=ws.title, row=r, column="B",
                    problem="More than one Total row",
                    detail=f"A Total row was already found on row {total_row}.",
                    expected="exactly one Total row", found=f"rows {total_row} and {r}",
                    fix="Remove the extra Total row from the source file."))
            total_row = r
            continue

        for key, label in (("received", "Complaints Received"),
                           ("resolved", "Complaints Resolved")):
            col = cmap[key]
            fv, vv = wf.cell(r, col).value, ws.cell(r, col).value
            if isinstance(fv, str) and fv.startswith("=") and vv is None:
                raise DataError(Issue(
                    file=fname, sheet=ws.title, row=r, column=get_column_letter(col),
                    problem="Formula has no calculated value",
                    detail=f"{label} for '{str(name).strip()}' is a formula whose "
                           "result was never calculated, so it would be read as 0.",
                    expected="a calculated number", found=f"{fv[:40]} (no value)",
                    fix="Open the file in Excel, let it calculate, Save, and upload again."))

        recv = _as_int(ws.cell(r, cmap["received"]).value, ws, r, cmap["received"], fname)
        res = _as_int(ws.cell(r, cmap["resolved"]).value, ws, r, cmap["resolved"], fname)
        nres_cell = ws.cell(r, cmap["notres"]).value
        nres = _as_int(nres_cell, ws, r, cmap["notres"], fname) if nres_cell is not None else recv - res

        if recv < 0 or res < 0 or nres < 0:
            bad = [("Received", recv), ("Resolved", res), ("Not Resolved", nres)]
            bad_field, bad_val = next((f, v) for f, v in bad if v < 0)
            raise DataError(Issue(
                file=fname, sheet=ws.title, row=r, column="C",
                problem="Negative value",
                detail=f"'{name}': {bad_field} came out to {bad_val}, which is "
                       f"impossible. Received={recv}, Resolved={res} - Resolved "
                       f"cannot exceed Received.",
                expected="Resolved <= Received, so Not Resolved is 0 or more",
                found=f"received={recv}, resolved={res}, not resolved={nres}",
                fix="Correct the source data for this agency - Resolved should "
                    "never be greater than Received."))
        if res + nres != recv:
            raise DataError(Issue(
                file=fname, sheet=ws.title, row=r,
                problem="Row does not add up",
                detail=f"'{name}': Resolved ({res}) + Not Resolved ({nres}) "
                       f"does not equal Received ({recv}).",
                expected=f"Resolved + Not Resolved = Received ({recv})",
                found=f"{res} + {nres} = {res + nres}",
                fix="Correct this row in the source file - these three columns "
                    "must always add up, or every downstream total is wrong."))

        key = nname
        if key in seen:
            raise DataError(Issue(
                file=fname, sheet=ws.title, row=r, column="B",
                problem="Duplicate agency",
                detail=f"'{name}' already appeared on row {seen[key]}.",
                expected="each agency once",
                found=f"rows {seen[key]} and {r}",
                fix="Remove or merge the duplicate row in the source file."))
        seen[key] = r

        rows.append(AgencyRow(name=re.sub(r"\s+", " ", str(name).replace("\n", " ")).strip(),
                              received=recv, resolved=res, notres=nres, src_row=r))

    if total_row is None:
        raise FormatError(Issue(
            file=fname, sheet=ws.title, problem="Total row not found",
            detail="No row with 'Total' in the Agency column was found.",
            expected="a row labelled 'Total'",
            fix="Add a Total row at the end of the agency list."))

    # Cross-check the source's own Total row against the rows above it. The
    # output Total is always recomputed from the rows, so a mismatch here means
    # the SOURCE total is stale or wrong - worth telling the user.
    for key, label in (("received", "Complaints Received"), ("resolved", "Complaints Resolved")):
        stated = ws.cell(total_row, cmap[key]).value
        if isinstance(stated, (int, float)):
            actual = sum(getattr(rw, key) for rw in rows)
            if int(round(stated)) != actual:
                warnings.append(
                    f"{fname}: the Total row says {label} = {int(round(stated))}, but the "
                    f"agency rows add up to {actual}. The output uses {actual}. "
                    f"Please check the source file.")

    # confirm every MCD zone is present exactly once
    missing = [z for z in MCD_ZONES if norm(z) not in seen]
    if missing:
        raise DataError([Issue(
            file=fname, sheet=ws.title, problem="MCD zone missing",
            detail=f"'{z}' was not found, so the three zones cannot be summed "
                   f"into one Municipal Corporation of Delhi row.",
            expected=f"a row named '{z}'", found="not present",
            fix="Add the missing zone's row, or correct its spelling - "
                "matching ignores case, spacing and punctuation.")
            for z in missing],
            headline=f"'{fname}' is missing one or more MCD zones, so nothing was generated")

    return ws, hrow, rows, total_row, a1, warnings


# ---------------------------------------------------------------------------
# Transform: merge MCD zones, reposition, renumber
# ---------------------------------------------------------------------------

def build_output_rows(rows):
    zone_keys = {norm(z) for z in MCD_ZONES}
    zones = [r for r in rows if norm(r.name) in zone_keys]
    others = [r for r in rows if norm(r.name) not in zone_keys]

    combined = AgencyRow(
        name=MCD_COMBINED_NAME,
        received=sum(z.received for z in zones),
        resolved=sum(z.resolved for z in zones),
        notres=sum(z.notres for z in zones),
        src_row=min(z.src_row for z in zones),
    )

    insert_at = MCD_TARGET_POSITION - 1          # 0-based index into `others`
    insert_at = max(0, min(insert_at, len(others)))
    out = others[:insert_at] + [combined] + others[insert_at:]

    # sanity: grand totals before and after must be identical - merging rows
    # can never change a sum, so any mismatch here is a bug, not real data.
    assert sum(r.received for r in rows) == sum(r.received for r in out)
    assert sum(r.resolved for r in rows) == sum(r.resolved for r in out)

    return out, combined


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

def _copy_style(src_cell, dst_cell):
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format


def write_excel(src_path, out_path, out_rows, heading_text):
    """Build the corrected workbook by writing fresh formulas for every row
    (all formulas here reference only their OWN row, so this is always
    correct regardless of where a row ends up - no stale references)."""
    src_wb = openpyxl.load_workbook(src_path)
    src_ws = src_wb.worksheets[0]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = src_ws.title

    # column widths
    for col, dim in src_ws.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width

    # title row
    ws.merge_cells("A1:G1")
    ws["A1"] = heading_text
    _copy_style(src_ws["A1"], ws["A1"])
    ws.row_dimensions[1].height = src_ws.row_dimensions[1].height

    # header row (row 2), including its merges
    for c in range(1, 8):
        _copy_style(src_ws.cell(2, c), ws.cell(2, c))
        ws.cell(2, c).value = src_ws.cell(2, c).value
    ws.merge_cells("D2:E2")
    ws.merge_cells("F2:G2")
    ws.row_dimensions[2].height = src_ws.row_dimensions[2].height

    style_row = 3   # template row for per-agency-row styling
    r = 3
    for row in out_rows:
        for c in range(1, 8):
            _copy_style(src_ws.cell(style_row, c), ws.cell(r, c))
        ws.cell(r, 1).value = r - 2                      # S. No.
        ws.cell(r, 2).value = row.name
        ws.cell(r, 3).value = row.received
        ws.cell(r, 4).value = row.resolved
        ws.cell(r, 5).value = pct_formula(f"D{r}", f"C{r}")
        ws.cell(r, 6).value = f"=C{r}-D{r}"
        ws.cell(r, 7).value = pct_formula(f"F{r}", f"C{r}")
        ws.row_dimensions[r].height = src_ws.row_dimensions[style_row].height
        r += 1

    total_r = r
    for c in range(1, 8):
        _copy_style(src_ws.cell(src_ws.max_row, c), ws.cell(total_r, c))
    ws.cell(total_r, 2).value = "Total"
    ws.cell(total_r, 3).value = f"=SUM(C3:C{total_r - 1})"
    ws.cell(total_r, 4).value = f"=SUM(D3:D{total_r - 1})"
    ws.cell(total_r, 5).value = pct_formula(f"D{total_r}", f"C{total_r}")
    ws.cell(total_r, 6).value = f"=C{total_r}-D{total_r}"
    ws.cell(total_r, 7).value = pct_formula(f"F{total_r}", f"C{total_r}")

    wb.save(out_path)
    return total_r


# ---------------------------------------------------------------------------
# PDF output
# ---------------------------------------------------------------------------

def write_pdf(out_path, out_rows, pdf_title, total):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    from xml.sax.saxutils import escape

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("NcrTitle", parent=styles["Title"], fontSize=13,
                                 alignment=TA_CENTER, spaceAfter=10)
    # Agency names are wrapped in a Paragraph so a long name breaks onto a second
    # line instead of running over the cell border. escape() matters: a name like
    # "BSES Delhi (BRPL & BYPL)" contains '&', which Paragraph reads as markup.
    cell_style = ParagraphStyle("NcrCell", parent=styles["Normal"],
                                fontName="Helvetica", fontSize=8, leading=10)

    doc = SimpleDocTemplate(out_path, pagesize=landscape(A4),
                            leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm)

    header = ["S. No.", "Agency", "Complaints\nReceived", "Complaints\nResolved",
              "% Resolved", "Not\nResolved", "% Not\nResolved"]
    data = [header]
    for i, row in enumerate(out_rows, 1):
        data.append([str(i), Paragraph(escape(row.name), cell_style), str(row.received), str(row.resolved),
                    pct(row.resolved, row.received), str(row.notres),
                    pct(row.notres, row.received)])
    data.append(["", "Total", str(total["received"]), str(total["resolved"]),
                pct(total["resolved"], total["received"]), str(total["notres"]),
                pct(total["notres"], total["received"])])

    col_widths = [16 * mm, 82 * mm, 28 * mm, 28 * mm, 22 * mm, 28 * mm, 24 * mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B7C6D9")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#EEF3F8")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DCE6F1")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (1, 0), (1, -1), 6),
        ("RIGHTPADDING", (1, 0), (1, -1), 6),
    ]
    table.setStyle(TableStyle(style))

    story = [Paragraph(pdf_title, title_style), Spacer(1, 4), table]
    doc.build(story)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def generate(source_path, outdir, end_date=None):
    ensure_pdf_support()
    os.makedirs(outdir, exist_ok=True)
    ws, hrow, rows, total_row, a1_text, warnings = read_source(source_path)

    heading = a1_text.strip()
    fname_src = os.path.basename(source_path)
    if end_date is not None:
        found = parse_heading_dates(heading)
        if not found:
            raise DataError(Issue(
                file=fname_src, sheet=ws.title, row=1, column="A",
                problem="Heading has no date range",
                detail="You selected an end date, but cell A1 has no date range "
                       "like '(01st Jan 2025 - 17th September 2026)' to update, so "
                       "the selected date would be silently ignored.",
                expected="a heading such as 'Social Media Agency-wise Data "
                         "(01st Jan 2025 - 17th September 2026) in Delhi-NCR'",
                found=heading[:80],
                fix="Add the date range to A1 in the source file, then run again."))
        if end_date < found[0]:
            raise DataError(Issue(
                file=fname_src, sheet=ws.title, row=1, column="A",
                problem="End date is before the start date",
                detail=f"The selected end date {heading_date(end_date)} is earlier "
                       f"than this file's start date {heading_date(found[0])}.",
                expected=f"an end date on or after {heading_date(found[0])}",
                found=heading_date(end_date),
                fix="Select a later end date."))
        heading = replace_end_date_in_title(heading, end_date)

    heading, base_name = _split_suffix(heading)     # Excel keeps ' in Delhi-NCR'; PDF does not
    stem = file_stem(base_name, os.path.basename(source_path))

    out_rows, combined = build_output_rows(rows)

    total = {"received": sum(r.received for r in out_rows),
             "resolved": sum(r.resolved for r in out_rows),
             "notres": sum(r.notres for r in out_rows)}
    src_total = {"received": sum(r.received for r in rows),
                "resolved": sum(r.resolved for r in rows),
                "notres": sum(r.notres for r in rows)}
    if total != src_total:
        raise DataError(Issue(
            file=os.path.basename(source_path), problem="Total mismatch after merge",
            detail="The grand totals changed after combining the MCD rows, "
                   "which should be mathematically impossible.",
            expected=str(src_total), found=str(total),
            fix="This indicates a bug - please report it rather than using the output."))

    xlsx_name = f"{stem}.xlsx"
    pdf_name = f"{stem}.pdf"
    xlsx_path = os.path.join(outdir, xlsx_name)
    pdf_path = os.path.join(outdir, pdf_name)

    total_row_n = write_excel(source_path, xlsx_path, out_rows, heading)
    write_pdf(pdf_path, out_rows, base_name, total)

    mcd_position = next(i for i, r in enumerate(out_rows, 1) if r.name == MCD_COMBINED_NAME)

    return {
        "source": os.path.basename(source_path),
        "heading": heading,
        "base_name": base_name,          # PDF heading
        "stem": stem,
        "xlsx": xlsx_path, "xlsx_name": xlsx_name,
        "pdf": pdf_path, "pdf_name": pdf_name,
        "row_count": len(out_rows),
        "mcd_combined": combined,
        "mcd_position": mcd_position,
        "total": total,
        "warnings": warnings,
    }


def main():
    import argparse
    p = argparse.ArgumentParser(description="Generate Delhi-NCR agency-wise Excel + PDF with MCD combined.")
    p.add_argument("--source", required=True)
    p.add_argument("--outdir", default="./out_ncr")
    a = p.parse_args()
    res = generate(a.source, a.outdir)
    print(f"MCD combined -> position {res['mcd_position']} "
         f"(received={res['mcd_combined'].received}, resolved={res['mcd_combined'].resolved})")
    print(f"Rows: {res['row_count']}")
    print(f"Excel: {res['xlsx']}")
    print(f"PDF  : {res['pdf']}")


if __name__ == "__main__":
    main()
