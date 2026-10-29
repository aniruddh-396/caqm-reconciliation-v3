"""
CAQM Agency Reconciliation - Mapping & Configuration
=====================================================

This is the ONLY file you should normally need to edit.

The base workbooks (agency_analysis / resolution_analysis) identify agencies by
their X/Twitter handle. The editable workbook (Agency_wise) identifies them by
full agency name. HANDLE_TO_AGENCY is the bridge between the two.

When several handles map to the SAME agency name, the engine automatically
aggregates them:
    - count columns  (Pending / Resolved / Grand Total) -> SUM
    - period columns (Avg Resolution / Avg Pendency)    -> MEAN

That is exactly how MCD Delhi's three zonal handles and NHAI's three state
handles are combined. No special-case code exists for them; they are just
handles that share a target name.
"""

# ---------------------------------------------------------------------------
# 1. HANDLE -> AGENCY NAME
#    Right-hand side MUST match the Agency name in Agency_wise.xlsx column C
#    (matching is case/space/punctuation insensitive, but keep it exact anyway).
# ---------------------------------------------------------------------------

HANDLE_TO_AGENCY = {
    # ---------------- Delhi ----------------
    "@bsesdelhi":            "BSES Delhi (BRPL & BYPL)",
    "@DelhiJalBoard":        "Delhi Jal Board",
    "@DelhiPwd":             "Public Works Department",
    "@DP_Pollution":         "Delhi Police Department",
    "@DPCC_pollution":       "Delhi Pollution Control Committee",
    "@DSIIDC95981579":       "Delhi State Industrial and Infrastructure Development Corporation",
    "@dtptraffic":           "Traffic Police Department",
    "@hqrevenue_":           "Revenue Department",
    "@IFC_Dpt":              "Irrigation and Flood Control Department",
    "@IGL_Care":             "Indraprastha Gas Ltd",
    "@official_dda":         "Delhi Development Authority",
    "@OfficialDMRC":         "Delhi Metro Rail Corporation",
    "@PcdTransport":         "Delhi Transport Department",
    "@tweetndmc":            "New Delhi Municipal Council",

    # --- MCD Delhi: 3 zonal handles -> ONE combined row (SUM / MEAN) ---
    "@MCD_Delhi (EDMC)":      "Municipal Corporation of Delhi",
    "@MCD_Delhi (North DMC)": "Municipal Corporation of Delhi",
    "@MCD_Delhi (SDMC)":      "Municipal Corporation of Delhi",

    # --- Northern Railway is filed under Delhi in the base workbooks ---
    "@drm_dli":              "Northern Railway",

    # ---------------- Haryana ----------------
    "@BallabgarhR":          "HSPCB - RO Ballabhgarh",
    "@BHspcb":               "RO - Bahadurgarh",
    "@FaridabadRegion":      "HSPCB - RO Faridabad",
    "@Haryana_spcb":         "Haryana SPCB",
    "@HspcbN":               "RO - Gurugram North",
    "@HspcbRegion":          "RO - Panipat",
    "@HspcbS":               "RO - Gurugram South",
    "@MCF_Faridabad":        "Municipal Corporation Faridabad",
    "@MunCorpGurugram":      "Municipal Corporation Gurugram",
    "@OfficialGMDA":         "Gurugram Metropolitan Development Authority",
    "@TrafficGGM":           "Traffic Police Gurugram",
    "@ulbharyana":           "Directorate of Urban Local Bodies, Haryana",

    # ---------------- Rajasthan ----------------
    "@RSPCB_Official":       "Rajasthan SPCB",

    # ---------------- Uttar Pradesh ----------------
    "@AMRUTCityGzb":         "Municipal Corporation Ghaziabad",
    "@gdagzb":               "Ghaziabad Development Authority",
    "@Hapur_NPP800737":      "Nagar Palika Parishad Hapur",
    "@MrtNagarNigam":        "Meerut Nagar Nigam",
    "@noida_authority":      "Noida Authority",
    "@noidatraffic":         "Noida Traffic Police",
    "@NppLoni":              "Nagar Palika Parishad Loni",
    "@OfficialGNIDA":        "Greater Noida Authority",
    "@rouppcbgbn":           "UPPCB - RO Noida",
    "@UppcbG":               "UPPCB - RO Ghaziabad",
    "@UppcbN":               "UPPCB - RO Greater Noida",
    "@uppcbofficial":        "UPPCB",

    # --- NHAI: one handle appearing under 3 states -> ONE combined row ---
    #     (the same key can only appear once, so the engine merges the three
    #      occurrences of this handle found across states automatically)
    "@NHAI_Official":        "National Highways Authority of India",
}


# ---------------------------------------------------------------------------
# 2. AGENCIES PRESENT IN Agency_wise BUT DELIBERATELY ABSENT FROM THE BASE
#    These are reported as "expected-missing" instead of as errors.
# ---------------------------------------------------------------------------

EXPECTED_MISSING = {
    "Tata Power Delhi Distribution Ltd.",   # zero complaints received
}


# ---------------------------------------------------------------------------
# 3. ENGINE SETTINGS
# ---------------------------------------------------------------------------

# How should "Complaints Received" (column D) be derived?
#   "formula"          -> write =E+G  (Pending + Resolved).
#   "base_grand_total" -> copy the Grand Total column from agency_analysis
#                         verbatim.  <-- REQUIRED SPEC: Complaints Received
#                         is compared against agency_analysis "Grand Total".
# These two DISAGREE on 10 rows because agency_analysis is internally
# inconsistent. See the "Base Integrity" sheet in the change log.
RECEIVED_MODE = "base_grand_total"

# A period value is only treated as a mismatch when it differs by at least this
# many days. Prevents false positives from display rounding
# (e.g. 187.25 in the sheet vs 187 in the base).
PERIOD_TOLERANCE_DAYS = 0.5

# Decimal places for averaged period values.
PERIOD_DECIMALS = 0

# Text used when a period cannot be computed (no pending / no resolved cases).
NA_TEXT = "N/A"

# Values in the base workbooks that mean "not computable".
ERROR_VALUES = {"#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "N/A", ""}

# Refresh the hardcoded Avg Resolution / Avg Pendency on the *_Total rows from
# the matching total rows in resolution_analysis.
UPDATE_TOTAL_ROWS = True

# Map: label in Agency_wise -> label in the base workbooks.
TOTAL_ROW_MAP = {
    "Delhi Total":          "Delhi Total",
    "Haryana Total":        "Haryana Total",
    "Rajasthan Total":      "Rajasthan Total",
    "Uttar Pradesh Total":  "Uttar Pradesh Total",
    "Grand Total":          "Grand Total",
}
