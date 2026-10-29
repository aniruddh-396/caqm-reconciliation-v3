# CAQM Complaint Data Reconciliation

Two modules on one engine. Each corrects an editable workbook against base
workbook(s) and returns the corrected file plus a full audit trail.

## Module 1 - Social Media Agency-wise Data of CAQM Tagged Complaints (04th Oct 2024 - 10th September 2026)

| Column in Agency_wise | Source | Rule |
|---|---|---|
| Complaints Received | `agency_analysis` -> Grand Total | copy, or SUM across handles |
| Pending Complaints | `agency_analysis` -> Pending | copy, or SUM across handles |
| Resolved Complaints | `agency_analysis` -> Resolved | copy, or SUM across handles |
| Average Resolution Period (Days) | `resolution_analysis` | copy, or MEAN across handles |
| Average Pendency Period (Days) | `resolution_analysis` | copy, or MEAN across handles |

Rows keyed by Twitter handle in the base files, mapped to agency names via
`HANDLE_TO_AGENCY` in `mapping.py`. Multi-handle agencies (MCD Delhi's 3 zones,
NHAI's 3 state entries) combine automatically: counts SUM, periods MEAN.

## Module 2 - Social Media Category-wise Data of CAQM Tagged Complaints (04th Oct 2024 - 10th September 2026)

| Column in caqm_category_wise | Source | Rule |
|---|---|---|
| Complaints Received | `sectoral_analysis` -> Grand Total | copy |
| Pending Complaints | `sectoral_analysis` -> Pending | copy |
| Resolved Complaints | `sectoral_analysis` -> Resolved | copy |

Rows matched on the pair **(State, Category)**, because the same category name
recurs under every state. No handle layer, no aggregation, no period columns.

## Files

```
mapping.py               Module 1 handle map + settings   <- the only file you normally edit
reconcile.py             Module 1 engine + CLI
verify.py                Module 1 independent verifier
reconcile_category.py    Module 2 engine + CLI (reuses reconcile.py helpers)
verify_category.py       Module 2 independent verifier
app.py                   unified Streamlit UI with a module selector
```

Keep all files in the same folder - the Module 2 engine and the app import from
the Module 1 files.

## Run

```bash
python -m pip install -r requirements.txt   # once (streamlit, openpyxl, reportlab)

# web UI (both modules, pick from the radio at the top)
python -m streamlit run app.py

# command line - Module 1
python reconcile.py --agency-wise Agency_wise.xlsx \
    --agency-analysis agency_analysis.xlsx \
    --resolution-analysis resolution_analysis.xlsx --outdir ./out
python verify.py Agency_wise.xlsx agency_analysis.xlsx \
    resolution_analysis.xlsx out/Agency_wise_CORRECTED.xlsx

# command line - Module 2
python reconcile_category.py --category-wise caqm_category_wise.xlsx \
    --sectoral-analysis sectoral_analysis.xlsx --outdir ./out2
python verify_category.py caqm_category_wise.xlsx sectoral_analysis.xlsx \
    out2/caqm_category_wise_CORRECTED.xlsx
```

## Shared design rules

1. **Nothing is changed silently.** Every write is in the change log; small
   differences left alone are logged too.
2. **Match by content, not position.** Header rows are found by their text, so
   an inserted row upstream will not misalign the columns. Module 2 required
   this: "category" appears in the title row, so its header detection keys on
   three columns at once (Category + Pending + Grand Total).
3. **Fail loud, never drop.** An unmapped handle (M1) or an unmatched
   (State, Category) pair (M2) is reported; the run does not quietly skip data.
4. **Preserve the workbook.** Formatting, widths, merges and existing formulas
   survive; only target cells change. Percentage columns are live formulas and
   follow the corrected counts on their own.
5. **Independent verification.** `verify*.py` re-derives every expected value
   straight from the base files, without importing the engine, and asserts the
   output matches - 309 assertions for M1, 159 for M2. Run it after every job.

## Complaints Received = Grand Total

Both modules copy Complaints Received from the base **Grand Total** column as a
plain number, and repair the total rows to sum those cells with correct ranges.

Consequence: Received, Pending and Resolved are three independent copies of
three independent base columns. Wherever a base file has
Pending + Resolved != Grand Total, the output row will too - that is what the
source says. Such rows are listed on the **Base Integrity** sheet.

- Module 1 base (`agency_analysis`): 10 inconsistent rows.
- Module 2 base (`sectoral_analysis`): 0 inconsistent rows (clean).

Total-row repairs found in Module 2's editable file: the Uttar Pradesh Total
Resolved formula skipped its first category (`=SUM(F41:F53)` -> `=SUM(F40:F53)`),
and the Grand Total Resolved was a hardcoded 5204 -> corrected to 5205.

## Module 1: two patterns still worth investigating

Flagged on the **Review Manually** sheet of the Module 1 log.

- **Average Pendency +1** on 34 of 37 agency rows - looks like inclusive vs
  exclusive day counting. Settle the convention and re-run.
- **Eight resolution values differ ~10x** with matching leading digits
  (52/530, 42/424, ...), the signature of a lost decimal separator. Confirm
  against the raw complaint data before accepting.

## Extending

- Module 1, new agency: one line in `HANDLE_TO_AGENCY`. Two handles pointing at
  the same name aggregate automatically.
- Module 2, new state or category: nothing to configure - it is matched by name
  as long as the (State, Category) text agrees between the two files.

## Module 3 - Social Media State-wise (split parent into per-state files)

Splits the parent workbook `2025.xlsx` into four state files. This module
*generates* files from fixed templates rather than correcting an editable file.

| State | Parent Sr | Notes |
|---|---|---|
| Delhi | 1-20 | 3 MCD zones (East+North+South) summed into "Municipal Corporation of Delhi"; grouped into two Sub Totals + Traffic + Grand Total |
| Uttar Pradesh | 21-32 | flat list + Total |
| Haryana | 33-44 | flat list + Total |
| Rajasthan | 45 | single row + Total |

NHAI (sr 46) and Northern Railway (sr 47) are used in NO state.

The layouts live in `templates/` (delhi.xlsx, up.xlsx, haryana.xlsx,
rajasthan.xlsx) - these are the fixed formats you provided and are never
changed. Each run copies a template, matches every agency row to the parent by
name, pours in Received / Resolved / Not Resolved, recomputes both percentage
columns (round-half-up; 0 shown when Received is 0), and recomputes every Sub
Total / Total / Grand Total. An agency in a template that cannot be found in
the parent aborts the run.

**Date + naming.** The user picks the end date; the start stays 01 Jan 2025.
- Heading inside A1: `Social Media Agency-wise Data (01st Jan 2025 - <end>) in <State>`
  where `<end>` becomes e.g. `10th September 2026`. Only the end date changes;
  the `in <State>` suffix and spacing are preserved verbatim.
- File name: `<State>_Social Media Agency-wise Complaint Redressal (01-01-2025 - DD-MM-YYYY).xlsx`.

Run:

```bash
python reconcile_state.py --parent 2025.xlsx --end-date 2026-09-10 --outdir ./out_states
python verify_state.py 2025.xlsx 2026-09-10 ./out_states
```

In the app, Module 3 takes just the parent upload and a date picker, then
offers each state file plus a ZIP of all four.

## Module 4 - Delhi-NCR Agency-wise, MCD consolidated (Excel + PDF)

Takes ONE flat, whole-Delhi-NCR agency-wise workbook (no State column, one row
per agency, a single Total row - e.g. a "2025-2026" or "2021-2025" export) and:

1. Sums East + North + South Delhi Municipal Corporation into one
   "Municipal Corporation of Delhi" row.
2. Places that row at **Sr. No. 12** in the output; every other agency keeps
   its existing relative order (net two fewer rows than the source, since
   three rows become one).
3. Renumbers S. No. 1..N and recomputes the Total row - which must equal the
   source's own total exactly, since merging rows can never change a sum
   (asserted, not just computed).
4. Produces **both** a corrected `.xlsx` and a matching `.pdf`.

**Naming.** The source file's own A1 heading already carries its date range,
e.g. `Social Media Agency-wise Data (01st Jan 2025 - 17th September 2026) in
Delhi-NCR`. The output file name is that heading with the trailing
` in Delhi-NCR` removed:

```
Social Media Agency-wise Data (01st Jan 2025 - 17th September 2026).xlsx
Social Media Agency-wise Data (01st Jan 2025 - 17th September 2026).pdf
```

The Excel heading keeps ` in Delhi-NCR`; the PDF heading is the file name
(without the suffix) - per spec. An optional end-date override works like the
other modules if you need to change just the end date without editing the
source.

Formulas for the recomputed Not-Resolved and both percentage columns are
written fresh for each row's own position (`=C{r}-D{r}`, etc.) rather than
shifted with `insert_rows`/`delete_rows`, which is why row reordering never
leaves a stale formula pointing at the wrong cell - every formula here only
ever references its own row.

Run:

```bash
python reconcile_ncr.py --source "2025-2026.xlsx" --outdir ./out_ncr
python verify_ncr.py "2025-2026.xlsx" "out_ncr/Social Media Agency-wise Data (...).xlsx" "out_ncr/Social Media Agency-wise Data (...).pdf"
```

In the app, Module 4 takes one file upload, an optional end-date override, and
gives two download buttons (Excel + PDF).

## Change log

**Category-wise template updated** - a fifth Rajasthan row, `Road Dust`, was
added to `templates/category_wise.xlsx`. S.No. was renumbered 1-52 so the new
row has a number and there are no gaps. The base file (`sectoral_analysis`)
must contain a matching Rajasthan / Road Dust row; if it does not, the app
shows a red "NOT FOUND" alert naming that row and leaves it unchanged.

**Date pickers now default to today (India time).** All end-date pickers use
`datefmt.today_ist()` instead of a fixed date. IST is computed explicitly
because the hosted server runs on UTC, where "today" would otherwise be
yesterday for the first 5.5 hours of each Indian day.


**Module 4 and app update.**
- Zero-complaint rows now show `0` (not `#DIV/0!`) in the percentage columns of
  both the Excel and the PDF; the Excel uses `=IF(C=0,0,...)` so it stays 0 after
  recalculation.
- Module 4 refuses a state-wise file (it has Sub Total rows), a formula cell that
  was never calculated, and an end date earlier than the file's start date; it warns
  if the source's own Total row disagrees with the rows above it. Long agency names
  wrap in the PDF.
- An end date is now required in every module (pre-filled with today in IST, limited
  to start date .. today). Results show a warning if the date or files change after
  they were generated, and the state-wise ZIP is named from the date it was
  generated for, not the date currently shown.
- New look (banner, module cards, step cards) and a "software can make mistakes -
  always re-check" warning above every download and at the foot of every page.

**Module 4 file names changed.** Both outputs are now named
`Social_Media_Status_<START>_<END>` (DDMMYYYY), e.g.
`Social_Media_Status_01012025_26102026.xlsx` / `.pdf` for the 2025-2026 report and
`Social_Media_Status_15102021_26102026.xlsx` / `.pdf` for the 2021-2025 report.
START is the report's fixed start date (read from the source file's heading);
END is the end date the user chose. Headings are unchanged: the Excel A1 ends with
` in Delhi-NCR`; the PDF heading is the same text without it. The app previews the
exact names and headings as soon as a file and a date are selected.

**Running on your own PC.** Install every dependency in one go, then start the app:
`python -m pip install -r requirements.txt` then `python -m streamlit run app.py`.
Module 4 (PDF) needs `reportlab`; if it is missing the app now says so on screen and
disables the Generate button instead of failing part-way.

**Warning placement.** The "software can make mistakes - always re-check" warning now
appears once, at the foot of every page (not above the download buttons).
#   c a q m - r e c o n c i l i a t i o n - v 3  
 #   c a q m - r e c o n c i l i a t i o n - v 3  
 