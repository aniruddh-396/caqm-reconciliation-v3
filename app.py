"""
CAQM Complaint Data Reconciliation - web app (4 modules)

    pip install streamlit openpyxl reportlab
    streamlit run app.py

Module 1  Agency-wise   correct the Agency-wise report from two base files
Module 2  Category-wise correct the Category-wise report from a base file
Module 3  State-wise    split the parent workbook into 4 state files
Module 4  Delhi-NCR     merge the 3 MCD zones, produce Excel + PDF

Every module requires an end date. All results carry a re-check warning.
"""
import io
import os
import tempfile
import traceback
import zipfile

import streamlit as st

import mapping as CFG
import reconcile as R1
import reconcile_category as R2
import reconcile_state as R3
import reconcile_ncr as R4
from errors import CaqmError
from datefmt import today_ist, heading_date, file_date
from datetime import date

st.set_page_config(page_title="CAQM Complaint Data Reconciliation",
                   page_icon="🌿", layout="wide")

M1 = "Social Media Agency-wise Data of CAQM Tagged Complaints (04th Oct 2024 - 10th September 2026)"
M2 = "Social Media Category-wise Data of CAQM Tagged Complaints (04th Oct 2024 - 10th September 2026)"
M3 = "Social Media State-wise Data (split parent into per-state files)"
M4 = "Social Media Agency-wise Data - Delhi-NCR (MCD consolidated, Excel + PDF)"

SHORT = {
    M1: "1  ·  Agency-wise  -  correct & update",
    M2: "2  ·  Category-wise  -  correct & update",
    M3: "3  ·  State-wise  -  generate 4 state files",
    M4: "4  ·  Delhi-NCR  -  MCD merged, Excel + PDF",
}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------------------
# Look & feel
# ---------------------------------------------------------------------------
CSS = """
<style>
:root { --green:#0B6E4F; --navy:#0B3C5D; --mint:#E9F4EE; --line:#D3E4DB; --ink:#16302B; }

html, body, .stApp, [class*="css"] {
    font-family: "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
}
.block-container { padding-top: 3.4rem; padding-bottom: 2rem; max-width: 1180px; }
footer { visibility: hidden; }

/* ---- hero banner ---- */
.hero {
    background: linear-gradient(120deg, #0B3C5D 0%, #0B6E4F 100%);
    color: #fff; padding: 1.5rem 1.8rem; border-radius: 18px;
    margin: 0 0 1.1rem 0; box-shadow: 0 8px 22px rgba(11,60,93,.22);
}
.hero .pill {
    display: inline-block; background: rgba(255,255,255,.16);
    border: 1px solid rgba(255,255,255,.32); padding: .16rem .75rem;
    border-radius: 999px; font-size: .74rem; letter-spacing: .06em; text-transform: uppercase;
}
.hero .hero-title { font-size: 1.95rem; font-weight: 750; margin: .45rem 0 .25rem 0; letter-spacing: -.01em; line-height: 1.15; }
.hero .hero-sub { color: #D8EBE3; font-size: 1rem; margin: 0; }

/* ---- module chooser: the radio becomes four cards ---- */
section.main [data-testid="stRadio"], [data-testid="stMain"] [data-testid="stRadio"],
section.main [data-testid="stRadio"] > div, [data-testid="stMain"] [data-testid="stRadio"] > div { width: 100%; }
/* Streamlit sizes the radio's outer container to its content; stretch it to the full row */
[data-testid="stMain"] [data-testid="stElementContainer"]:has([data-testid="stRadio"]),
section.main [data-testid="stElementContainer"]:has([data-testid="stRadio"]) { width: 100% !important; }
section.main [data-testid="stRadio"] > div[role="radiogroup"],
[data-testid="stMain"] [data-testid="stRadio"] > div[role="radiogroup"] {
    display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: .6rem;
}
section.main [data-testid="stRadio"] label,
[data-testid="stMain"] [data-testid="stRadio"] label {
    background: #fff; border: 1.5px solid var(--line); border-radius: 12px;
    padding: .7rem 1rem; margin: 0; cursor: pointer; transition: all .15s ease; width: 100%;
}
section.main [data-testid="stRadio"] label:hover,
[data-testid="stMain"] [data-testid="stRadio"] label:hover { border-color: var(--green); background: var(--mint); }
section.main [data-testid="stRadio"] label:has(input:checked),
[data-testid="stMain"] [data-testid="stRadio"] label:has(input:checked) {
    border-color: var(--green); background: var(--mint); box-shadow: inset 0 0 0 1px var(--green);
}
@media (max-width: 820px) {
    section.main [data-testid="stRadio"] > div[role="radiogroup"],
    [data-testid="stMain"] [data-testid="stRadio"] > div[role="radiogroup"] { grid-template-columns: 1fr; }
    .hero .hero-title { font-size: 1.5rem; }
}

/* ---- module title ---- */
.mod-title { font-size: 1.28rem; font-weight: 700; color: var(--navy); margin: .9rem 0 .15rem 0; }
.mod-blurb { color: #4A635B; margin: 0 0 .9rem 0; }

/* ---- step cards ---- */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 14px; border-color: var(--line); background: #fff; }
.step { display: flex; align-items: center; gap: .65rem; margin: .05rem 0 .65rem 0; }
.step .num {
    width: 1.75rem; height: 1.75rem; border-radius: 50%; background: var(--green); color: #fff;
    display: inline-flex; align-items: center; justify-content: center; font-weight: 700; font-size: .9rem;
}
.step .ttl { font-weight: 700; font-size: 1.06rem; color: var(--navy); }
.date-chip {
    display: inline-block; background: var(--mint); border: 1px solid var(--line); color: var(--navy);
    padding: .35rem .85rem; border-radius: 999px; font-size: .9rem; margin-top: .3rem;
}

/* ---- metrics, buttons, tabs ---- */
[data-testid="stMetric"] {
    background: #fff; border: 1px solid var(--line); border-radius: 12px;
    padding: .75rem 1rem; box-shadow: 0 1px 3px rgba(11,60,93,.07);
}
[data-testid="stMetricValue"] { color: var(--navy); font-weight: 700; }
[data-testid="stMetricLabel"] { color: #4A635B; }
.stButton > button, .stDownloadButton > button { border-radius: 10px; font-weight: 600; padding: .5rem 1.2rem; }
button[data-baseweb="tab"] { font-weight: 600; }
[data-testid="stSidebar"] { background: #F3F8F5; border-right: 1px solid var(--line); }

/* ---- re-check warning ---- */
.recheck {
    border-left: 6px solid #D97706; background: #FFF6E3; color: #5B3A00;
    padding: .85rem 1.1rem; border-radius: 10px; margin: .8rem 0 .4rem 0; line-height: 1.5;
}
.recheck b { color: #8A4B00; }
.recheck .hi { display: block; margin-top: .3rem; }
.app-foot { text-align: center; color: #6B7F78; font-size: .82rem; margin-top: .9rem; }
</style>
"""

HERO = """
<div class="hero">
  <span class="pill">CAQM &nbsp;·&nbsp; Social Media Tagged Complaints</span>
  <div class="hero-title">Complaint Data Reconciliation</div>
  <p class="hero-sub">Compare, correct and generate complaint reports &mdash; every change
  is logged and every mismatch is flagged.</p>
</div>
"""

RECHECK_HTML = """
<div class="recheck">
  <b>&#9888;&#65039; Important &mdash; software can make mistakes.</b>
  Always re-check every number, name and date in the generated files against your
  source data before you use, sign or share them.
  <span class="hi">सॉफ़्टवेयर से गलती हो सकती है। उपयोग से पहले हर आँकड़े, नाम और तारीख़ को
  अपने मूल डेटा से दोबारा ज़रूर जाँचें।</span>
</div>
"""

st.markdown(CSS, unsafe_allow_html=True)
st.markdown(HERO, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Small UI helpers
# ---------------------------------------------------------------------------
def step(n, title):
    st.markdown(f'<div class="step"><span class="num">{n}</span>'
                f'<span class="ttl">{title}</span></div>', unsafe_allow_html=True)


def module_header(title, blurb):
    st.markdown(f'<div class="mod-title">{title}</div>'
                f'<div class="mod-blurb">{blurb}</div>', unsafe_allow_html=True)


def date_step(key, start, label="End date"):
    """A REQUIRED date picker. Pre-filled with today (India time) and limited
    to start..today, so a typo like year 2062 is impossible. If the field is
    ever empty, nothing can be generated."""
    today = today_ist()
    d = st.date_input(f"{label}  (required)", value=today, min_value=start,
                      max_value=today, format="DD/MM/YYYY", key=key,
                      help="This date is written into the heading of every output "
                           "file and into its file name. Please confirm it is correct.")
    if d is None:
        st.error("Please select a date. It is required before anything can be generated.")
        return None
    st.markdown(f'<div class="date-chip">📅 Outputs will be dated up to '
                f'<b>{heading_date(d)}</b></div>', unsafe_allow_html=True)
    return d


def file_sig(*uploads):
    return tuple((u.name, u.size) for u in uploads if u is not None)


def stale_notice(meta, cur_date, cur_files):
    """The results on screen belong to the date and files used when they were
    generated. If either has changed since, say so loudly - otherwise a user
    could download an old file believing it reflects the new date."""
    if meta is None:
        return
    if cur_date is not None and meta["date"] != cur_date:
        st.warning(f"These results were generated for **{heading_date(meta['date'])}**, "
                   f"but the date is now set to **{heading_date(cur_date)}**. "
                   f"Click the Generate button again so the files match the new date.")
    if meta["files"] != cur_files:
        st.warning("The uploaded files have changed since these results were generated. "
                   "Click the Generate button again before downloading.")


def recheck_notice():
    st.markdown(RECHECK_HTML, unsafe_allow_html=True)


def save_uploads(**files):
    tmp = tempfile.mkdtemp()
    paths = {}
    for k, up in files.items():
        p = os.path.join(tmp, up.name)
        with open(p, "wb") as fh:
            fh.write(up.getbuffer())
        paths[k] = p
    return tmp, paths


def download_row(output, log, out_name):
    d1, d2 = st.columns(2)
    with open(output, "rb") as fh:
        d1.download_button("⬇️  Download corrected workbook", fh.read(), out_name,
                           XLSX_MIME, type="primary", use_container_width=True)
    with open(log, "rb") as fh:
        d2.download_button("⬇️  Download change log", fh.read(),
                           os.path.basename(log), XLSX_MIME, use_container_width=True)


def show_error(e):
    """Render a CaqmError as a readable headline + a table of issues so the
    user can see exactly which file, row and column is wrong."""
    st.error(e.headline)
    rows = []
    for i in e.issues:
        rows.append({
            "File": i.file,
            "Sheet": i.sheet or "-",
            "Row": i.row if i.row != "" else "-",
            "Column": i.column or "-",
            "Problem": i.problem,
            "Detail": i.detail,
            "Expected": i.expected or "-",
            "Found": i.found or "-",
            "How to fix": i.fix or "-",
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)
    st.caption("Nothing was written. Fix the rows listed above in the source "
               "file and run again.")


def run_guarded(fn, *a, **kw):
    """Call an engine, turning every failure into a clear on-screen report."""
    try:
        return fn(*a, **kw), None
    except CaqmError as e:
        show_error(e)
        return None, "handled"
    except Exception:
        st.error("An unexpected problem stopped the run. Nothing was written.")
        st.code(traceback.format_exc())
        return None, "unexpected"


def show_unmatched_alert(unmatched, label="agency"):
    """Prominent alert naming every entry that could not be matched, split by
    direction so the user knows which file to fix."""
    if not unmatched:
        st.success(f"Every {label} in the format was found in the uploaded files.")
        return

    key = "agency" if label == "agency" else "category"
    not_found = [u for u in unmatched if u.get("severity") == "NOT FOUND"]
    not_in_fmt = [u for u in unmatched if u.get("severity") == "NOT IN FORMAT"]
    expected = [u for u in unmatched if u.get("severity", "").startswith("OK")]

    if not_found:
        names = "\n".join(
            f"- **{u[key]}**" + (f"  (row {u['row']}"
                                 + (f", {u['state']}" if u.get('state') and u['state'] != '-' else "")
                                 + ")" if u.get("row") != "-" else "")
            for u in not_found)
        st.error(
            f"**{len(not_found)} {label}(s) in the format were NOT found in the "
            f"uploaded file(s).** Their rows were left unchanged:\n\n{names}\n\n"
            f"Check the spelling in the uploaded base file, or add the missing "
            f"row there. Matching ignores case, spacing and punctuation.")

    if not_in_fmt:
        names = "\n".join(f"- **{u[key]}**" for u in not_in_fmt)
        st.warning(
            f"**{len(not_in_fmt)} {label}(s) in the uploaded file have no row in "
            f"the format**, so their numbers were not carried into the output:"
            f"\n\n{names}")

    if expected:
        st.info(f"{len(expected)} {label}(s) are known to be absent from the base "
                f"files and were skipped on purpose: "
                + ", ".join(u[key] for u in expected) + ".")


def changes_table(changes, has_source):
    rows = []
    for c in changes:
        label = getattr(c, "agency", None) or getattr(c, "category", None)
        row = {"Row": c.row, "State": c.state,
               ("Agency" if has_source else "Category"): label,
               "Col": c.column, "Field": c.field_name,
               "Old": "" if c.old is None else str(c.old),
               "New": "" if c.new is None else str(c.new),
               "Rule": c.rule}
        if has_source:
            row["Source"] = c.sources
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Sidebar (all modules)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🌿 CAQM Reconciliation")
    st.markdown("**Before you generate**\n\n"
                "1. Upload the correct files\n"
                "2. Confirm the end date\n"
                "3. Generate, then **re-check** the output")
    st.divider()

# ---------------------------------------------------------------------------
# Module chooser
# ---------------------------------------------------------------------------
st.markdown("**Choose a module**")
module = st.radio("Choose a module", [M1, M2, M3, M4], index=0,
                  format_func=lambda m: SHORT[m], label_visibility="collapsed")

START_M1_M2 = date(2024, 10, 4)      # fixed start date in the two templates
START_M3 = date(2025, 1, 1)          # fixed start date in the state files
START_M4 = date(2021, 10, 15)        # earliest start among the NCR reports


# ===========================================================================
# MODULE 1 - AGENCY-WISE
# ===========================================================================
if module == M1:
    prev = st.session_state.get("m1date") or today_ist()
    module_header(f"Social Media Agency-wise Data of CAQM Tagged Complaints "
                  f"(04th Oct 2024 - {heading_date(prev)})",
                  "Corrects the Agency-wise report against agency_analysis (counts) and "
                  "resolution_analysis (periods). MCD Delhi and NHAI are combined first: "
                  "counts summed, periods averaged.")

    with st.sidebar:
        st.markdown("**Module 1 settings**")
        received_mode = st.radio(
            "Complaints Received is derived from",
            ["base_grand_total", "formula"],
            format_func=lambda m: ("Grand Total column in agency_analysis"
                                   if m == "base_grand_total"
                                   else "Pending + Resolved (sheet's own convention)"),
        )
        tolerance = st.number_input("Period tolerance (days)", 0.0, 10.0,
                                    float(CFG.PERIOD_TOLERANCE_DAYS), 0.25)
        update_totals = st.checkbox("Refresh period values on *_Total rows",
                                    value=CFG.UPDATE_TOTAL_ROWS)
        with st.expander("Handle mapping"):
            for h, a in CFG.HANDLE_TO_AGENCY.items():
                st.text(f"{h}  ->  {a}")

    with st.container(border=True):
        step(1, "Upload the two base files")
        c1, c2 = st.columns(2)
        f_count = c1.file_uploader("agency_analysis.xlsx  (counts)", type=["xlsx"], key="m1c")
        f_period = c2.file_uploader("resolution_analysis.xlsx  (periods)", type=["xlsx"], key="m1p")
        with st.expander("Use a different editable file instead of the built-in format"):
            f_edit = st.file_uploader("Custom Agency-wise file (optional)",
                                      type=["xlsx"], key="m1e")
            st.caption("Leave empty to use the built-in format. A custom file must still "
                       "have the standard columns, or you will get a report naming the "
                       "exact row and column.")

    with st.container(border=True):
        step(2, "Choose the end date")
        m1_date = date_step("m1date", START_M1_M2)

    with st.container(border=True):
        step(3, "Compare and correct")
        go = st.button("Compare and correct", type="primary", key="m1run",
                       disabled=not (f_count and f_period and m1_date))
    cur_sig = file_sig(f_count, f_period, f_edit)

    if go:
        CFG.RECEIVED_MODE = received_mode
        CFG.PERIOD_TOLERANCE_DAYS = float(tolerance)
        CFG.UPDATE_TOTAL_ROWS = bool(update_totals)
        ups = {"count": f_count, "period": f_period}
        if f_edit:
            ups["edit"] = f_edit
        _, paths = save_uploads(**ups)
        with st.spinner("Reconciling..."):
            res, err = run_guarded(R1.reconcile, paths.get("edit"),
                                   paths["count"], paths["period"],
                                   tempfile.mkdtemp(), end_date=m1_date)
        st.session_state["m1"] = res
        st.session_state["m1_meta"] = {"date": m1_date, "files": cur_sig}

    res = st.session_state.get("m1")
    if res:
        meta = st.session_state.get("m1_meta")
        st.markdown("#### Results")
        stale_notice(meta, m1_date, cur_sig)
        m = st.columns(5)
        m[0].metric("Matched", len(res["matched"]))
        m[1].metric("Unmatched", len(res["unmatched"]))
        m[2].metric("Cells changed", len(res["changes"]))
        m[3].metric("Total-row edits", len(res["total_updates"]))
        m[4].metric("Base warnings", len(res["integrity"]))

        show_unmatched_alert(res["unmatched"], "agency")

        if res["integrity"]:
            st.warning(f"{len(res['integrity'])} rows in agency_analysis have "
                       "Pending + Resolved != Grand Total - see Base Integrity in the log.")
        d = meta["date"] if meta else today_ist()
        download_row(res["output"], res["log"],
                     f"Agency_wise_CORRECTED ({file_date(d)}).xlsx")
        st.divider()
        t = st.tabs(["Changes", "Total rows", "Aggregated", "Base integrity", "Unmatched"])
        t[0].dataframe(changes_table(res["changes"], True), use_container_width=True, hide_index=True)
        t[1].dataframe(changes_table(res["total_updates"], True), use_container_width=True, hide_index=True)
        t[2].dataframe(
            [{"Agency": r.agency, "Handles": " + ".join(r.handles),
              "Pending": r.pending, "Resolved": r.resolved, "Grand Total": r.grand_total,
              "Avg Resolution": r.avg_resolution, "Avg Pendency": r.avg_pendency}
             for r in res["records"].values() if r.is_aggregate],
            use_container_width=True, hide_index=True)
        t[3].dataframe(res["integrity"], use_container_width=True, hide_index=True)
        t[4].dataframe(res["unmatched"], use_container_width=True, hide_index=True)

# ===========================================================================
# MODULE 2 - CATEGORY-WISE
# ===========================================================================
elif module == M2:
    prev = st.session_state.get("m2date") or today_ist()
    module_header(f"Social Media Category-wise Data of CAQM Tagged Complaints "
                  f"(04th Oct 2024 - {heading_date(prev)})",
                  "Corrects the Category-wise report against sectoral_analysis. Rows are "
                  "matched on the pair (State, Category). Received comes from Grand Total, "
                  "Pending from Pending, Resolved from Resolved.")

    with st.container(border=True):
        step(1, "Upload the base file")
        f_base = st.file_uploader("sectoral_analysis.xlsx  (base)", type=["xlsx"], key="m2b")
        with st.expander("Use a different editable file instead of the built-in format"):
            f_edit = st.file_uploader("Custom Category-wise file (optional)",
                                      type=["xlsx"], key="m2e")
            st.caption("Leave empty to use the built-in format. A custom file must still "
                       "have the standard columns, or you will get a report naming the "
                       "exact row and column.")

    with st.container(border=True):
        step(2, "Choose the end date")
        m2_date = date_step("m2date", START_M1_M2)

    with st.container(border=True):
        step(3, "Compare and correct")
        go = st.button("Compare and correct", type="primary", key="m2run",
                       disabled=not (f_base and m2_date))
    cur_sig = file_sig(f_base, f_edit)

    if go:
        ups = {"base": f_base}
        if f_edit:
            ups["edit"] = f_edit
        _, paths = save_uploads(**ups)
        with st.spinner("Reconciling..."):
            res, err = run_guarded(R2.reconcile_category, paths.get("edit"),
                                   paths["base"], tempfile.mkdtemp(),
                                   end_date=m2_date)
        st.session_state["m2"] = res
        st.session_state["m2_meta"] = {"date": m2_date, "files": cur_sig}

    res = st.session_state.get("m2")
    if res:
        meta = st.session_state.get("m2_meta")
        st.markdown("#### Results")
        stale_notice(meta, m2_date, cur_sig)
        m = st.columns(5)
        m[0].metric("Matched", len(res["matched"]))
        m[1].metric("Unmatched", len(res["unmatched"]))
        m[2].metric("Cells changed", len(res["changes"]))
        m[3].metric("Total-row edits", len(res["total_updates"]))
        m[4].metric("Base warnings", len(res["integrity"]))
        show_unmatched_alert(res["unmatched"], "category")

        if res["integrity"]:
            st.warning(f"{len(res['integrity'])} rows in sectoral_analysis have "
                       "Pending + Resolved != Grand Total - see Base Integrity in the log.")
        d2 = meta["date"] if meta else today_ist()
        download_row(res["output"], res["log"],
                     f"caqm_category_wise_CORRECTED ({file_date(d2)}).xlsx")
        st.divider()
        t = st.tabs(["Changes", "Total rows", "Base integrity", "Unmatched"])
        t[0].dataframe(changes_table(res["changes"], False), use_container_width=True, hide_index=True)
        t[1].dataframe(changes_table(res["total_updates"], False), use_container_width=True, hide_index=True)
        t[2].dataframe(res["integrity"], use_container_width=True, hide_index=True)
        t[3].dataframe(res["unmatched"], use_container_width=True, hide_index=True)

# ===========================================================================
# MODULE 3 - STATE-WISE (split parent -> per-state files)
# ===========================================================================
elif module == M3:
    module_header("Social Media State-wise Data",
                  "Splits the parent workbook (2025.xlsx) into one file per state - Delhi, "
                  "UP, Haryana, Rajasthan - using the fixed layouts. Delhi's three MCD zones "
                  "are summed into one 'Municipal Corporation of Delhi' row. NHAI and "
                  "Northern Railway are not used in any state.")

    with st.container(border=True):
        step(1, "Upload the parent workbook")
        f_parent = st.file_uploader("Parent workbook  (2025.xlsx)", type=["xlsx"], key="m3p")

    with st.container(border=True):
        step(2, "Choose the end date")
        end_dt = date_step("m3date", START_M3, label="End date (the '- to' date)")
        if end_dt:
            st.caption(f"Files will be named e.g. `Delhi_Social Media Agency-wise Complaint "
                       f"Redressal (01-01-2025 - {file_date(end_dt)}).xlsx` and headed "
                       f"`... (01st Jan 2025 - {heading_date(end_dt)}) in Delhi`. "
                       f"The start date is fixed at 01 Jan 2025.")

    with st.container(border=True):
        step(3, "Generate the state files")
        go = st.button("Generate state files", type="primary", key="m3run",
                       disabled=not (f_parent and end_dt))
    cur_sig = file_sig(f_parent)

    if go:
        tmp = tempfile.mkdtemp()
        pp = os.path.join(tmp, f_parent.name)
        with open(pp, "wb") as fh:
            fh.write(f_parent.getbuffer())
        outdir = os.path.join(tmp, "out")
        with st.spinner("Generating..."):
            res, err = run_guarded(R3.generate_all, pp, end_dt, outdir)
        st.session_state["m3"] = res
        st.session_state["m3_meta"] = {"date": end_dt, "files": cur_sig}

    res = st.session_state.get("m3")
    if res:
        meta = st.session_state.get("m3_meta")
        gen_date = meta["date"] if meta else today_ist()     # date the files were made for
        st.markdown("#### Results")
        stale_notice(meta, end_dt, cur_sig)
        st.success(f"Generated {len(res['results'])} state files, dated up to "
                   f"{heading_date(gen_date)}.")

        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            for r in res["results"]:
                z.write(r["output"], arcname=r["filename"])
        zbuf.seek(0)
        st.download_button("⬇️  Download all 4 as ZIP", zbuf.read(),
                           f"CAQM_State-wise_{file_date(gen_date)}.zip",
                           "application/zip", type="primary", use_container_width=True)

        cols = st.columns(len(res["results"]))
        for col, r in zip(cols, res["results"]):
            with open(r["output"], "rb") as fh:
                col.download_button(f"{r['state']}", fh.read(), r["filename"],
                                    XLSX_MIME, use_container_width=True)

        st.divider()
        for r in res["results"]:
            with st.expander(f"{r['state']} - {r['filename']}"):
                st.dataframe(
                    [{"Row": c["row"], "Agency": c["agency"], "Kind": c["kind"],
                      "Received": c["received"], "Resolved": c["resolved"],
                      "Not Resolved": c["notres"]} for c in r["changes"]],
                    use_container_width=True, hide_index=True)

# ===========================================================================
# MODULE 4 - DELHI-NCR AGENCY-WISE, MCD CONSOLIDATED (Excel + PDF)
# ===========================================================================
else:
    prev = st.session_state.get("m4date") or today_ist()
    module_header("Social Media Agency-wise Data - Delhi-NCR",
                  "Upload a flat Delhi-NCR agency-wise workbook (for example the 2025-2026 or "
                  "2021-2025 report). East + North + South Delhi Municipal Corporation are "
                  "summed into one 'Municipal Corporation of Delhi' row placed at Sr. No. 12. "
                  "Produces one corrected Excel and one matching PDF, both named "
                  "Social_Media_Status_&lt;start&gt;_&lt;end&gt; (DDMMYYYY). Zero-complaint rows show 0 in the "
                  "percentage columns.")

    with st.container(border=True):
        step(1, "Upload the source workbook")
        f_src = st.file_uploader("Source workbook  (e.g. 2025-2026.xlsx)", type=["xlsx"], key="m4src")

    with st.container(border=True):
        step(2, "Choose the end date")
        m4_date = date_step("m4date", START_M4)
        if m4_date:
            pv = R4.preview_names(io.BytesIO(f_src.getvalue()), m4_date) if f_src else None
            if pv:
                xh, ph, stem = pv
                st.markdown(f"**Excel heading:** {xh}  \n"
                            f"**PDF heading:** {ph}  \n"
                            f"**File names:** `{stem}.xlsx` and `{stem}.pdf`")
            else:
                st.caption("File names will be `Social_Media_Status_<START>_<END>` (DDMMYYYY). "
                           "The start date is fixed by the source file; the end date is the "
                           "one you chose. Upload the file to see the exact names.")

    pdf_ok = True
    try:
        R4.ensure_pdf_support()
    except CaqmError as e:
        pdf_ok = False
        show_error(e)

    with st.container(border=True):
        step(3, "Generate Excel + PDF")
        go = st.button("Generate Excel + PDF", type="primary", key="m4run",
                       disabled=not (f_src and m4_date and pdf_ok))
    cur_sig = file_sig(f_src)

    if go:
        tmp = tempfile.mkdtemp()
        sp = os.path.join(tmp, f_src.name)
        with open(sp, "wb") as fh:
            fh.write(f_src.getbuffer())
        with st.spinner("Generating..."):
            res, err = run_guarded(R4.generate, sp, os.path.join(tmp, "out"),
                                   end_date=m4_date)
        st.session_state["m4"] = res
        st.session_state["m4_meta"] = {"date": m4_date, "files": cur_sig}

    res = st.session_state.get("m4")
    if res:
        meta = st.session_state.get("m4_meta")
        st.markdown("#### Results")
        stale_notice(meta, m4_date, cur_sig)
        m = st.columns(3)
        m[0].metric("Agencies in output", res["row_count"])
        m[1].metric("MCD combined position", res["mcd_position"])
        m[2].metric("MCD total received", res["mcd_combined"].received)

        for w in res.get("warnings", []):
            st.warning(w)

        st.success(f"Excel heading: *{res['heading']}*  \n"
                   f"PDF heading: *{res['base_name']}*  \n"
                   f"Files: `{res['xlsx_name']}`  and  `{res['pdf_name']}`")

        d1, d2 = st.columns(2)
        with open(res["xlsx"], "rb") as fh:
            d1.download_button("⬇️  Download Excel", fh.read(), res["xlsx_name"],
                               XLSX_MIME, type="primary", use_container_width=True)
        with open(res["pdf"], "rb") as fh:
            d2.download_button("⬇️  Download PDF", fh.read(), res["pdf_name"],
                               "application/pdf", type="primary",
                               use_container_width=True)

        st.divider()
        mcd = res["mcd_combined"]
        st.markdown(f"**Municipal Corporation of Delhi** (East + North + South "
                    f"combined) - Received: {mcd.received}, Resolved: "
                    f"{mcd.resolved}, Not Resolved: {mcd.notres}")
        st.caption(f"Grand Total after merge: Received {res['total']['received']}, "
                   f"Resolved {res['total']['resolved']}, Not Resolved "
                   f"{res['total']['notres']} - identical to the source rows' sum, "
                   f"as merging rows can never change a sum.")

# ---------------------------------------------------------------------------
# Footer - shown on every module, every time
# ---------------------------------------------------------------------------
st.divider()
recheck_notice()
st.markdown('<div class="app-foot">CAQM Complaint Data Reconciliation &nbsp;·&nbsp; '
            'Outputs are only as reliable as the files you upload.</div>',
            unsafe_allow_html=True)
