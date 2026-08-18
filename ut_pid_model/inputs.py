"""
inputs.py — Excel input page for the model.

Provides a single, formatted **Inputs workbook** that exposes every model
variable (mirroring the workbook's "Inputs - First" named ranges) plus a
**Development Inputs** sheet for the year-by-year absorption / value schedules.

  * ``write_inputs_workbook(cfg, dev, path)`` — write a populated, editable
    template (Title | Value | Range Name | Notes), one row per named range.
  * ``load_inputs_workbook(path)`` — read an edited workbook back into a
    ``(ModelConfig, DeveloperProjections)`` pair, so the whole model can be
    driven from Excel.

Round-trips: ``load_inputs_workbook(write_inputs_workbook(cfg, dev))`` reproduces
the same configuration.
"""

from __future__ import annotations

import os
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .config import ModelConfig
from .development import DeveloperProjections


# ── Input specification ──────────────────────────────────────────────────────
# Each entry: (section, label, range_name, attr, kind, note)
# kind: "str" | "float" | "pct" | "int" | "date" | "yesno"
_SPECS: list[tuple] = [
    ("Basic Inputs", "Public Infrastructure District Name", "METRO", "pid_name", "str", ""),
    ("Basic Inputs", "City", "CITY", "city", "str", ""),
    ("Basic Inputs", "County", "COUNTY", "county", "str", ""),
    ("Basic Inputs", "Developer", "DEVELOPER", "developer", "str", "memo addressee"),
    ("Basic Inputs", "Developer Address", "DEVELOPER_ADDRESS", "developer_address", "str", "memo addressee — street"),
    ("Basic Inputs", "Developer City/State/Zip", "DEVELOPER_CITY_STATE_ZIP", "developer_city_state_zip", "str", "memo addressee — city/state/zip"),
    ("Basic Inputs", "Current Certified Value", "CURRENT_CERTIFIED_VALUE", "current_certified_value", "float", "$ certified AV on the rolls"),
    ("Basic Inputs", "Existing Land Value", "EXISTING_VACANT_LAND", "existing_vacant_land", "float", "$ AV — one-time, counted in first collection year only"),
    ("Basic Inputs", "Existing Residential Value", "EXISTING_RESIDENTIAL_VALUE", "existing_residential_value", "float", "$ AV — one-time, counted in first collection year only"),
    ("Basic Inputs", "State Taxable Value", "STATE_ASSESSED", "state_assessed", "float", "$ AV — added to value, held constant"),
    ("Basic Inputs", "Exempt Value", "EXEMPT_VALUE", "exempt_value", "float", "$ AV — subtracted from value, held constant"),
    ("Basic Inputs", "Date of Certification", "CERTIFICATION_DATE", "certification_date", "date", "date county certified the value"),

    ("Financing Toggles", "Refunding (refinance senior new money)", "REFUND_FINANCING", "refund_financing", "yesno", "Yes/No"),
    ("Financing Toggles", "Capitalized Interest", "CAPI", "capi", "yesno", "Yes/No"),
    ("Financing Toggles", "Surplus Toggle", "SURPLUS_ON_OFF", "surplus_on_off", "yesno", "Yes/No"),
    ("Financing Toggles", "Release Surplus Funds?", "SURPLUS_RELEASE_SIZING", "surplus_release_sizing", "yesno", "Yes/No"),
    ("Financing Toggles", "Investment Grade Rating", "IG_RATED", "ig_rated", "yesno", "Yes/No"),

    ("Fees", "Costs of Issuance", "COI", "coi", "float", "$"),
    ("Fees", "Costs of Issuance - Refunding", "COI_REFUNDING", "coi_refunding", "float", "$"),
    ("Fees", "Underwriter's Discount - Senior", "UWD_SENIOR", "uwd_senior", "pct", "% of par"),
    ("Fees", "Underwriter's Discount - Subordinate", "UWD_SUB", "uwd_sub", "pct", "% of par"),
    ("Fees", "Underwriter's Discount - Refunding", "UWD_SENIOR_REFUNDING", "uwd_senior_refunding", "pct", "% of par"),
    ("Fees", "County Collection Cost", "COUNTY_COLLECTION_FEE", "county_collection_fee", "pct", "% of mill revenue; 0% in Utah (§ 59-2-1602)"),
    ("Fees", "Trustee Fee (Senior)", "TRUSTEE_FEE", "trustee_fee", "float", "$/yr"),
    ("Fees", "Trustee Fee (Subordinate)", "TRUSTEE_FEE_SUB", "trustee_fee_sub", "float", "$/yr"),
    ("Fees", "Bond Insurance / Surety Rate", "BOND_INSURANCE_RATE", "bond_insurance_rate", "pct", "% of total DS (refunding)"),

    ("Structuring", "Delivery Date", "DELIVERY", "delivery", "date", ""),
    ("Structuring", "Principal Maturity Month", "PRIN_MATURITY", "prin_maturity", "int", "1-12"),
    ("Structuring", "Interest Maturity Month", "INT_MATURITY", "int_maturity", "int", "1-12"),
    ("Structuring", "Senior Principal Pay Day", "PRIN_MATURITY_DAY_SENIOR", "prin_maturity_day_senior", "int", ""),
    ("Structuring", "Subordinate Principal Pay Day", "PRIN_MATURITY_DAY_SUB", "prin_maturity_day_sub", "int", ""),
    ("Structuring", "Capitalized Interest Period (months)", "CAPI_TERM", "capi_term", "int", ""),
    ("Structuring", "Final Maturity - Senior (yrs)", "FINAL_MAT_YRS", "final_mat_yrs", "int", ""),
    ("Structuring", "Final Maturity - Subordinate (yrs)", "FINAL_MAT_SUB_YRS", "final_mat_sub_yrs", "int", ""),
    ("Structuring", "Final Maturity - Refunding (yrs)", "FINAL_MAT_YRS_REFUNDING", "final_mat_yrs_refunding", "int", ""),
    ("Structuring", "Years to Premium Call", "PREMIUM_CALL_YEARS", "premium_call_years", "int", "first optional redemption"),
    ("Structuring", "Premium Call Price", "PREMIUM_CALL_FIRST_PRICE", "premium_call_price", "float", "e.g. 103; premium steps down 1%/yr to par"),

    ("Interest Rates", "Senior Lien Interest Rate", "SENIOR_INTEREST_RATE", "senior_interest_rate", "pct", ""),
    ("Interest Rates", "Senior Refunding Interest Rate", "SENIOR_REFUNDING_INTEREST_RATE", "senior_refunding_interest_rate", "pct", ""),
    ("Interest Rates", "Subordinate Interest Rate", "SUB_INTEREST_RATE", "sub_interest_rate", "pct", ""),
    ("Interest Rates", "Senior Reoffering Yield (blank=par)", "SENIOR_REOFFERING_YIELD", "senior_reoffering_yield", "pct", "below coupon ⇒ premium"),
    ("Interest Rates", "Refunding Reoffering Yield (blank=par)", "SENIOR_REFUNDING_REOFFERING_YIELD", "senior_refunding_reoffering_yield", "pct", "below coupon ⇒ premium"),

    ("Coverage", "Debt Service Coverage - Senior", "DSC_SENIOR_LIEN_BONDS", "dsc_senior", "float", "x"),
    ("Coverage", "Debt Service Coverage - Subordinate", "DSC_SUB_LIEN_BONDS", "dsc_sub", "float", "x"),
    ("Coverage", "Debt Service Coverage - Refunding", "DSC_REFUNDING_BONDS", "dsc_refunding", "float", "x"),

    ("Funds", "Interest Earnings Rate (reserve/surplus)", "INTEREST_EARN_RATE", "interest_earn_rate", "pct", ""),
    ("Funds", "Surplus Fund Target Factor", "SURPLUS_FUND_TARGET_FACTOR", "surplus_fund_target_factor", "float", "x max senior DS"),

    ("Taxable Value Timing", "AV Lag (years)", "AV_LAG_YEARS", "av_lag_years", "int", ""),
    ("Taxable Value Timing", "Hold Value Flat Between Reassessments", "HOLD_VALUE_FLAT", "hold_value_flat", "yesno", "TRUE/FALSE"),

    ("Stress Testing", "Home Sales Pace (% of forecast)", "ABSORPTION_PACE_FACTOR", "absorption_pace_factor", "pct", "100% = base; 50% halves the monthly pace"),

    ("Tax & Valuation", "First Year (Summary)", "FIRST_YEAR", "first_year", "int", ""),
    ("Tax & Valuation", "Projection Horizon (years from delivery)", "PROJECTION_YEARS", "projection_years", "int", "development, AV, revenue & summary tabs run this many years"),
    ("Tax & Valuation", "Inflation Start Year", "INFLATION_START_YEAR", "inflation_start_year", "int", "home-price inflation begins this year; flat (not deflated) before"),
    ("Tax & Valuation", "Inflation Rate (home prices)", "INFLATION_RATE", "inflation_rate", "pct", ""),
    ("Tax & Valuation", "Inflation Rate (commercial sales)", "INFLATION_RATE_COMM_SALES", "inflation_rate_comm_sales", "pct", ""),
    ("Tax & Valuation", "Reassessment - Residential", "REASSESS_RATE", "reassess_rate", "pct", "applied annually in Utah"),
    ("Tax & Valuation", "Reassessment - Commercial", "REASSESS_COMM_RATE", "reassess_comm_rate", "pct", ""),
    ("Developer Contribution", "Developer Contribution ($)", "DEVELOPER_CONTRIBUTION", "developer_contribution", "float", "cash contributed by the developer (a source of funds)"),
    ("Developer Contribution", "Applied To (Senior / Subordinate / Proportional)", "DEVELOPER_CONTRIBUTION_SERIES", "developer_contribution_series", "str", "which series the contribution funds; Proportional spreads it by par"),
    ("Tax & Valuation", "Primary Residential Taxable Ratio", "RESID_TAXABLE_RATIO", "resid_taxable_ratio", "pct", "55% — the 45% exemption, § 59-2-103"),
    ("Tax & Valuation", "Prior Residential Taxable Ratio", "RESID_TAXABLE_RATIO_PRIOR", "resid_taxable_ratio_prior", "pct", "ratio before the current exemption"),
    ("Tax & Valuation", "Mill Levy Tax Collection %", "TAX_COLLECT_MILL_PRC", "tax_collect_mill_prc", "pct", ""),
    ("Tax & Valuation", "Personal Property Uniform Fee %", "UNIFORM_FEE_PRC", "uniform_fee_prc", "pct", "§ 59-2-405; % of mill revenue"),
    ("Tax & Valuation", "Uniform Fee Taxable Value Threshold", "UNIFORM_FEE_AV_THRESHOLD", "uniform_fee_av_threshold", "float", "$"),
    ("Tax & Valuation", "Reassessment Frequency", "REASSESS_FREQUENCY", "reassess_frequency", "text", "Annual (Utah, § 59-2-303.1) or Biennial (Colorado cadence)"),

    ("Mill Levies", "Mill Levy — Governing Document Cap", "MILL_LEVY_GOVERNING_DOC", "mill_levy_governing_doc", "float", "mills"),
    ("Mill Levies", "Mill Levy — Indenture Cap", "MILL_LEVY_INDENTURE", "mill_levy_indenture", "float", "mills; blank ⇒ no indenture cap"),
    ("Mill Levies", "Mill Levy Target - Debt Service", "MILL_LEVY_DS_TARGET", "mill_levy_ds_target", "float", "mills"),
    ("Mill Levies", "Mill Levy - Commercial", "MILL_LEVY_COMM", "mill_levy_comm", "float", "mills"),
    ("Mill Levies", "Mill Levy Target - Operations", "MILL_LEVY_OPS_TARGET", "mill_levy_ops_target", "float", "mills"),

    ("Lot / Home Ratios", "Builder Lot Inventory Taxable Ratio", "LOT_INVENTORY_TAXABLE_RATIO", "lot_inventory_taxable_ratio", "pct", "55% under Utah Admin. Code R884-24P-52; 100% to tax at full market"),
    ("Lot / Home Valuation", "Platted Lot Value (% of ASP)", "PLATTED_LOT_VALUE", "platted_lot_value", "pct", "lot value = ASP × this (default 10%)"),
    ("Lot / Home Valuation", "Platted Commercial Lot Value", "PLATTED_COMM_LOT_VALUE", "platted_comm_lot_value", "pct", ""),

    ("Centrally Assessed / Commercial", "Centrally Assessed Property (market value)", "CENTRALLY_ASSESSED_VALUE", "centrally_assessed_value", "float", "$ market — § 59-2-201"),
    ("Centrally Assessed / Commercial", "Centrally Assessed Equipment (market value)", "CENTRALLY_ASSESSED_EQUIPMENT", "centrally_assessed_equipment", "float", "$ market"),
    ("Centrally Assessed / Commercial", "Centrally Assessed Taxable Ratio", "CENTRALLY_ASSESSED_RATIO", "centrally_assessed_ratio", "pct", "market value × ratio"),
    ("Centrally Assessed / Commercial", "Include Centrally Assessed in Taxed Value", "CENTRALLY_ASSESSED", "centrally_assessed", "yesno", "Yes/No"),
    ("Centrally Assessed / Commercial", "Commercial Taxable Ratio", "COMMERCIAL_ASSESSMENT_RATIO", "commercial_assessment_ratio", "pct", ""),
    ("Centrally Assessed / Commercial", "Commercial AV Lag (years)", "COMM_ASSESSMENT_LAG_YEARS", "comm_assessment_lag_years", "int", ""),
    ("Centrally Assessed / Commercial", "Include Commercial in Taxed AV", "COMM_NEW_VALUE_ADD", "comm_new_value_add", "yesno", "Yes/No"),

    ("District Costs", "Starting O&M Expense", "OM_EXPENSE", "om_expense", "float", "$ per year of district operations & administration, netted from the revenue available to both liens"),
    ("District Costs", "O&M Expense Growth Rate", "OM_GROWTH_RATE", "om_growth_rate", "pct", "annual inflation on the O&M base"),
    ("District Costs", "O&M Expense Taxable Value Limit", "OM_EXPENSE_AV_LIMIT", "om_expense_av_limit", "float", "$ — above this taxable value the charge stops; 0 ⇒ no limit"),
    ("District Costs", "First Year District Costs Are Charged", "DISTRICT_COST_START_YEAR", "district_cost_start_year", "int", "blank ⇒ two years after closing"),
]

# Range names that live on DeveloperProjections rather than ModelConfig.
_DEV_ATTRS = {"_base_asp": "base_asp"}


# ── Styles ──────────────────────────────────────────────────────────────────
_BLUE = PatternFill("solid", fgColor="1F4E79")
_SECT = PatternFill("solid", fgColor="2E6FA3")
_INPUT = PatternFill("solid", fgColor="FFF2CC")   # editable value cells (yellow)
_HDRF = PatternFill("solid", fgColor="D6E4F0")
_WHITEFONT = Font(bold=True, color="FFFFFF", size=12)
_SECTFONT = Font(bold=True, color="FFFFFF", size=10)
_HDRFONT = Font(bold=True, color="1F4E79", size=9)
_LBL = Font(size=10)
_VAL = Font(bold=True, size=10)
_NOTE = Font(italic=True, size=9, color="595959")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_L = Alignment(horizontal="left", vertical="center")
_R = Alignment(horizontal="right", vertical="center")
_C = Alignment(horizontal="center", vertical="center")

_PCT = '0.000%'           # percentage / rate input cells (3 decimals)
_DOLLAR = '#,##0'
_NUM = '#,##0'
_COVERAGE = '#,##0.000'   # coverage ratios (3 decimals)
_DATEFMT = 'YYYY-MM-DD'


def _fmt_for(kind: str):
    return {"pct": _PCT, "float": _DOLLAR, "date": _DATEFMT}.get(kind)


# ── Write template ───────────────────────────────────────────────────────────
# ── Historical certified-values true-up table (Inputs side table) ─────────────
# Off to the right of the main input list (columns G–K). Up to 4 prior roll years
# of certified lot-inventory, residential, state-assessed, and exempt values, which
# the model uses to true up historical values.  Fixed layout so it can be both
# written and read back regardless of the main list's length.
_HIST_TITLE_ROW = 4          # banner
_HIST_NOTE_ROW = 5           # sub-note
_HIST_HDR_ROW = 6            # column headers
_HIST_FIRST_DATA_ROW = 7     # first of 4 editable year rows
_HIST_N_YEARS = 4
_HIST_YEAR_COL = 7           # G
_HIST_COLS = [               # (column index, header, config key)
    (8,  "Builder Lot Inventory",    "vacant_land"),
    (9,  "Residential",    "residential"),
    (10, "State Assessed", "state_assessed"),
    (11, "Exempt",         "exempt"),
]


def _write_historical_av_table(ws, cfg):
    """Write the historical certified-values true-up entry table (columns G–K)."""
    ws.column_dimensions["F"].width = 3   # spacer between main list and side table
    ws.column_dimensions[get_column_letter(_HIST_YEAR_COL)].width = 13
    for col, _hdr, _k in _HIST_COLS:
        ws.column_dimensions[get_column_letter(col)].width = 16
    last_col = _HIST_COLS[-1][0]

    ws.merge_cells(start_row=_HIST_TITLE_ROW, start_column=_HIST_YEAR_COL,
                   end_row=_HIST_TITLE_ROW, end_column=last_col)
    t = ws.cell(row=_HIST_TITLE_ROW, column=_HIST_YEAR_COL,
                value="HISTORICAL CERTIFIED VALUES (TRUE-UP)")
    t.fill = _BLUE; t.font = _WHITEFONT; t.alignment = _C

    ws.merge_cells(start_row=_HIST_NOTE_ROW, start_column=_HIST_YEAR_COL,
                   end_row=_HIST_NOTE_ROW, end_column=last_col)
    n = ws.cell(row=_HIST_NOTE_ROW, column=_HIST_YEAR_COL,
                value=("Optional — enter up to 4 prior roll years of county-certified "
                       "values. The model uses these to true up historical values."))
    n.font = _NOTE; n.alignment = _L
    ws.row_dimensions[_HIST_NOTE_ROW].height = 26

    # Header row
    hc = ws.cell(row=_HIST_HDR_ROW, column=_HIST_YEAR_COL, value="Roll Year")
    hc.fill = _HDRF; hc.font = _HDRFONT; hc.border = _BORDER; hc.alignment = _C
    for col, hdr, _k in _HIST_COLS:
        c = ws.cell(row=_HIST_HDR_ROW, column=col, value=hdr)
        c.fill = _HDRF; c.font = _HDRFONT; c.border = _BORDER; c.alignment = _C

    # Pre-populate from cfg.historical_av when present, else leave editable blanks.
    hist = getattr(cfg, "historical_av", None) or {}
    years = sorted(hist)[:_HIST_N_YEARS]
    for i in range(_HIST_N_YEARS):
        rr = _HIST_FIRST_DATA_ROW + i
        yr = years[i] if i < len(years) else None
        yc = ws.cell(row=rr, column=_HIST_YEAR_COL, value=yr)
        yc.fill = _INPUT; yc.font = _VAL; yc.border = _BORDER; yc.alignment = _C
        yc.number_format = "0"
        row_data = hist.get(yr, {}) if yr is not None else {}
        for col, _hdr, key in _HIST_COLS:
            vc = ws.cell(row=rr, column=col, value=row_data.get(key))
            vc.fill = _INPUT; vc.font = _VAL; vc.border = _BORDER; vc.alignment = _C
            vc.number_format = _DOLLAR


def _load_historical_av(ws) -> dict | None:
    """Read the Inputs historical-values side table (columns G–K) into a dict."""
    out: dict[int, dict[str, float]] = {}
    for i in range(_HIST_N_YEARS):
        rr = _HIST_FIRST_DATA_ROW + i
        yr = ws.cell(row=rr, column=_HIST_YEAR_COL).value
        if not isinstance(yr, (int, float)):
            continue
        yr = int(yr)
        row = {}
        for col, _hdr, key in _HIST_COLS:
            v = ws.cell(row=rr, column=col).value
            if isinstance(v, (int, float)):
                row[key] = float(v)
        if row:
            out[yr] = row
    return out or None


def write_inputs_workbook(
    cfg: ModelConfig | None = None,
    dev: DeveloperProjections | None = None,
    output_path: str = "output/ut_pid_model_inputs.xlsx",
) -> str:
    """Write a populated, editable Inputs workbook from a config/dev (or defaults)."""
    cfg = cfg or ModelConfig()
    dev = dev or DeveloperProjections()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inputs"
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 34
    ws.column_dimensions["E"].width = 28

    ws.merge_cells("B1:E1")
    t = ws["B1"]
    t.value = f"{cfg.pid_name} — Model Inputs" if cfg.pid_name else "Model Inputs"
    t.fill = _BLUE; t.font = _WHITEFONT; t.alignment = _C
    ws.merge_cells("B2:E2")
    s = ws["B2"]; s.value = "Edit the yellow Value cells, then load with load_inputs_workbook()."
    s.font = _NOTE; s.alignment = _C

    # Column headers
    for col, lbl in [(2, "Title"), (3, "Value"), (4, "Range Name"), (5, "Notes")]:
        c = ws.cell(row=4, column=col, value=lbl)
        c.fill = _HDRF; c.font = _HDRFONT; c.border = _BORDER; c.alignment = _C

    # Dropdowns for boolean cells: Yes/No for string toggles, TRUE/FALSE for
    # bool flags (attached to each value cell as it is written below).
    dv_yesno = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv_tf = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    # Dropdown for the developer-contribution "Applied To" series — restricts input
    # to the four valid choices so the user cannot mistype.
    dv_series = DataValidation(
        type="list", formula1='"Senior,Subordinate,Proportional"',
        allow_blank=True)
    ws.add_data_validation(dv_yesno)
    ws.add_data_validation(dv_tf)
    ws.add_data_validation(dv_series)

    r = 5
    last_section = None
    for section, label, rng, attr, kind, note in _SPECS:
        if section != last_section:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            c = ws.cell(row=r, column=2, value=section)
            c.fill = _SECT; c.font = _SECTFONT; c.alignment = _L
            last_section = section
            r += 1
        value = getattr(dev, _DEV_ATTRS[attr]) if attr in _DEV_ATTRS else getattr(cfg, attr)
        # Default the commercial mill to the DS mill so it starts at the same rate.
        if attr == "mill_levy_comm" and value is None:
            value = cfg.commercial_mill_levy
        is_bool_flag = kind == "yesno" and isinstance(value, bool)
        if is_bool_flag:
            value = "TRUE" if value else "FALSE"
        if isinstance(value, (datetime, date)):
            value = value if isinstance(value, datetime) else datetime(value.year, value.month, value.day)

        ws.cell(row=r, column=2, value=label).font = _LBL
        ws.cell(row=r, column=2).border = _BORDER; ws.cell(row=r, column=2).alignment = _L
        vc = ws.cell(row=r, column=3, value=value)
        vc.fill = _INPUT; vc.font = _VAL; vc.border = _BORDER
        vc.alignment = _C   # center-align all editable (yellow) value cells
        fmt = _fmt_for(kind)
        if kind == "float" and (section == "Coverage"
                                or attr == "surplus_fund_target_factor"):
            fmt = _COVERAGE   # coverage ratios / fund factors shown to 3 decimals
        if fmt and not isinstance(value, str):
            vc.number_format = fmt
        if kind == "yesno":
            (dv_tf if is_bool_flag else dv_yesno).add(vc)
        if attr == "developer_contribution_series":
            dv_series.add(vc)
        rc = ws.cell(row=r, column=4, value=rng); rc.font = _NOTE; rc.border = _BORDER; rc.alignment = _L
        nc = ws.cell(row=r, column=5, value=note); nc.font = _NOTE; nc.border = _BORDER; nc.alignment = _L
        r += 1

    # ── Historical certified values (true-up) — side table, columns G–K ──────
    _write_historical_av_table(ws, cfg)

    ws.freeze_panes = "B5"

    # ── Development Inputs sheet (year-by-year schedules) ────────────────────
    _write_dev_sheet(wb.create_sheet("Development Inputs"), dev)

    # ── Debt Structure sheet (coupon / yield scale + term bonds) ─────────────
    _write_debt_structure_sheet(wb.create_sheet("Debt Structure"), cfg)

    # (The optional-redemption call schedule is a derived OUTPUT — see the
    # "Call Schedule" tab in the model output workbook, not an input.)

    # ── Historical Residential Exemption rates (reference — always the last tab) ─────────────
    _write_reference_sheet(wb.create_sheet("Utah Property Tax Reference"))

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    return output_path


_NPROD = 8   # product columns on the Development Inputs grids


def _write_dev_sheet(ws, dev: DeveloperProjections):
    """
    Development Inputs laid out like the Excel "Residential Development" sheet:
    Product Setup, then a Lot Delivery grid and a Home Closings grid (year rows ×
    product columns + Totals), then an Aggregate fallback and Commercial block.
    """
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 20
    last_col = 3 + _NPROD            # B(year/label) | C..(prod) | Total
    for c in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_col)
    t = ws.cell(row=1, column=2, value="Development Inputs"); t.fill = _BLUE
    t.font = _WHITEFONT; t.alignment = _C
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)
    s = ws.cell(row=2, column=2, value="Per-product grids (up to 8 products). Enter lot "
                "deliveries and home closings by year for each product. Edit yellow cells.")
    s.font = _NOTE; s.alignment = _C

    products = dev.products
    names = [(products[i].name if i < len(products) else None) for i in range(_NPROD)]

    def _marker(r, text):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        c = ws.cell(row=r, column=2, value=text)
        c.fill = _SECT; c.font = _SECTFONT; c.alignment = _L
        return r + 1

    def _cellw(r, c, v, hdr=False, inp=False, num=None, align=None):
        cell = ws.cell(row=r, column=c)
        if v is not None:
            cell.value = v
        cell.border = _BORDER
        if hdr:
            cell.fill = _HDRF; cell.font = _HDRFONT; cell.alignment = _C
        else:
            if inp:
                cell.fill = _INPUT
                cell.alignment = _C   # center-align all editable (yellow) cells
            else:
                cell.alignment = align or _R
        if num:
            cell.number_format = num

    # ── PRODUCT SETUP (one column per product) ───────────────────────────────
    r = _marker(4, "PRODUCT SETUP  (one column per product; up to 8)")
    _cellw(r, 2, "Parameter", hdr=True)
    for i in range(_NPROD):
        _cellw(r, 3 + i, f"Product {i + 1}", hdr=True)
    r += 1
    setup_rows = [
        ("Product Name", [names[i] for i in range(_NPROD)], None, _L),
        ("Total Units", [products[i].existing_units if i < len(products) else None
                         for i in range(_NPROD)], _NUM, _R),
        ("Base ASP ($)", [products[i].asp_base if i < len(products) else None
                          for i in range(_NPROD)], _DOLLAR, _R),
        ("ASP Base Year", [products[i].asp_base_year if i < len(products) else None
                           for i in range(_NPROD)], None, _C),
    ]
    setup_name_row = r
    for label, vals, num, align in setup_rows:
        _cellw(r, 2, label, align=_L)
        for i in range(_NPROD):
            _cellw(r, 3 + i, vals[i], inp=True, num=num, align=align)
        r += 1
    r += 1

    # Year axis for the grids — span the product years and pad out generously so
    # there are plenty of rows to enter a multi-year delivery / closing schedule.
    prod_years = set().union(*[set(p.lot_deliveries) | set(p.home_closings)
                               for p in products]) if products else set()
    start = min(prod_years) if prod_years else dev.asp_base_year
    end = max(prod_years) if prod_years else dev.asp_base_year
    grid_years = list(range(start, max(end + 3, start + 25) + 1))

    def _grid(r, title, getter):
        r = _marker(r, title)
        _cellw(r, 2, "Year", hdr=True)
        for i in range(_NPROD):
            _cellw(r, 3 + i, names[i] or f"Product {i + 1}", hdr=True)
        _cellw(r, last_col, "Total", hdr=True)
        r += 1
        for y in grid_years:
            rowvals = [getter(products[i], y) if i < len(products) else None
                       for i in range(_NPROD)]
            # Year is an editable cell; pre-filled only when the row has data,
            # otherwise left blank for the user to enter.
            _cellw(r, 2, y if any(rowvals) else None, inp=True, align=_C)
            for i in range(_NPROD):
                _cellw(r, 3 + i, rowvals[i] or None, inp=True, num=_NUM)
            ws.cell(row=r, column=last_col).border = _BORDER
            r += 1
        return r + 1

    r = _grid(r, "LOT DELIVERY  (units by year, per product)",
              lambda p, y: p.lot_deliveries.get(y))
    r = _grid(r, "HOME CLOSINGS  (units by year, per product)",
              lambda p, y: p.home_closings.get(y))

    # ── COMMERCIAL MARKET VALUE ──────────────────────────────────────────────
    r = _marker(r, "COMMERCIAL MARKET VALUE  (by year; blank if none)")
    for c, h in [(2, "Year"), (3, "Commercial Market Value")]:
        _cellw(r, c, h, hdr=True)
    r += 1
    for y in sorted(dev.commercial_market_value):
        _cellw(r, 2, y, align=_C)
        _cellw(r, 3, dev.commercial_market_value.get(y), inp=True, num=_DOLLAR)
        r += 1
    for _ in range(4):
        for c in (2, 3):
            _cellw(r, c, None, inp=(c > 2))
        r += 1

    ws.freeze_panes = "C5"




# ── Debt Structure sheet (pricing-day coupon / yield scale + term bonds) ─────
def _write_debt_structure_sheet(ws, cfg: ModelConfig):
    ws.column_dimensions["A"].width = 2
    for col in "BCDE":
        ws.column_dimensions[col].width = 20
    ws.merge_cells("B1:E1")
    t = ws["B1"]; t.value = "Debt Structure (pricing day)"
    t.fill = _BLUE; t.font = _WHITEFONT; t.alignment = _C
    ws.merge_cells("B2:E2")
    ws["B2"].value = ("Leave blank for preliminary analysis (uses the single flat rate / par). "
                      "On pricing day enter per-maturity Coupon & Yield; for a term bond, "
                      "put its final maturity year in the Term column on each sinking row.")
    ws["B2"].font = _NOTE; ws["B2"].alignment = _C

    def _marker(r, text):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(row=r, column=2, value=text)
        c.fill = _SECT; c.font = _SECTFONT; c.alignment = _L
        return r + 1

    def _hdr_row(r):
        for i, h in enumerate(["Maturity Year", "Coupon", "Yield", "Term Final Maturity"]):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.fill = _HDRF; c.font = _HDRFONT; c.border = _BORDER; c.alignment = _C
        return r + 1

    def _block(r, marker, years, coupon_scale, yield_scale, term_bonds):
        r = _marker(r, marker)
        r = _hdr_row(r)
        term_for = {}
        for (f, l, ty) in (term_bonds or []):
            for y in range(f, l + 1):
                term_for[y] = (l, ty)
        for y in years:
            cpn = (coupon_scale or {}).get(y)
            tf = term_for.get(y)
            yld = tf[1] if tf else (yield_scale or {}).get(y)
            vals = [y, cpn, yld, (tf[0] if tf else None)]
            for i in range(4):
                cell = ws.cell(row=r, column=2 + i)
                if vals[i] is not None:
                    cell.value = vals[i]
                cell.fill = _INPUT; cell.border = _BORDER
                cell.alignment = _C   # center-align all editable (yellow) cells
                if i in (1, 2):
                    cell.number_format = _PCT
            r += 1
        return r + 1

    senior_years = list(range(cfg.senior_first_principal_year, cfg.senior_final_year + 1))
    r = _block(4, "SENIOR BONDS", senior_years, cfg.senior_coupon_scale,
               cfg.senior_yield_scale, cfg.senior_term_bonds)
    ref_first = cfg.delivery_refunding.year + 1
    ref_last = cfg.delivery_refunding.year + cfg.final_mat_yrs_refunding
    ref_years = list(range(ref_first, ref_last + 1))
    _block(r, "REFUNDING BONDS", ref_years, cfg.senior_refunding_coupon_scale,
           cfg.senior_refunding_yield_scale, cfg.senior_refunding_term_bonds)
    ws.freeze_panes = "B4"


# ── Historical Utah residential exemption rates — reference ─────────────────
# Residential taxable ratio by tax year.  Used to inform
# RESID_TAXABLE_RATIO and RESID_TAXABLE_RATIO_PRIOR; this is a read-only reference table.
from .config import RESIDENTIAL_EXEMPTION_HISTORY as _RATE_HISTORY
from .config import BUILDER_INVENTORY_HISTORY as _INVENTORY_HISTORY
_RATE_SOURCES = [
    "Utah Const. art. XIII, § 3 — residential exemption, maximum 45% of fair market value",
    "Utah Code § 59-2-103 — primary residential exemption (45%); one acre of land per unit",
    "Utah Admin. Code R884-24P-52 — primary residence; unoccupied property and property "
    "under construction may take the exemption",
    "Utah Code § 59-2-303.1 — annual update of property values by the county assessor",
    "Utah Code § 17D-4-303 — PID levy limited to 0.015 per dollar of taxable value",
    "Utah Code § 59-2-405 — uniform fee on registered personal property, distributed to "
    "taxing entities in the same proportion as real property tax",
    "Utah Code § 59-2-1602 — multicounty assessing and collecting levy (county collection cost)",
    "Utah Code § 59-2-503 — agricultural (Greenbelt) valuation; rollback tax on withdrawal",
]


def _write_rate_table(ws, start_row, title, subtitle, rate_hdr, history,
                      editable=False):
    """
    Render one titled rate table (year / rate / notes) and return the next row.
    When ``editable`` the rate cells are shaded as input (yellow) cells — the
    model reads them back, so the user can retune the rates in Excel.
    """
    ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=4)
    t = ws.cell(row=start_row, column=2, value=title)
    t.fill = _BLUE; t.font = _WHITEFONT; t.alignment = _C
    ws.merge_cells(start_row=start_row + 1, start_column=2, end_row=start_row + 1, end_column=4)
    s = ws.cell(row=start_row + 1, column=2, value=subtitle)
    s.font = _NOTE; s.alignment = _C

    hdr_row = start_row + 3
    for col, lbl in [(2, "Tax / Roll Year(s)"), (3, rate_hdr), (4, "Authority / Notes")]:
        c = ws.cell(row=hdr_row, column=col, value=lbl)
        c.fill = _HDRF; c.font = _HDRFONT; c.border = _BORDER; c.alignment = _C

    r = hdr_row + 1
    for yr, rate, note in history:
        yc = ws.cell(row=r, column=2, value=yr)
        yc.font = _LBL; yc.border = _BORDER; yc.alignment = _C
        rc = ws.cell(row=r, column=3, value=rate)
        rc.font = _VAL; rc.border = _BORDER; rc.alignment = _C
        rc.number_format = '0.000%'
        if editable:
            rc.fill = _INPUT
        nc = ws.cell(row=r, column=4, value=note)
        nc.font = _NOTE; nc.border = _BORDER; nc.alignment = _L
        r += 1
    return r


# Marker text used to locate the editable rate tables on read-back.
_RESIDENTIAL_TABLE_TITLE = "Utah Primary Residential Exemption — Taxable Share of Fair Market Value"
_INVENTORY_TABLE_TITLE = "Utah Builder Lot Inventory Taxable Ratios"


def _load_rate_schedule(wb, table_title: str) -> dict | None:
    """
    Read an editable rate table (located by ``table_title``) off the "Historical
    Residential Exemption Rates" tab into a {roll_year: rate} schedule (carry-forward), keyed by
    the low bound of each row's Tax/Roll Year(s) label.  Returns None if absent.
    """
    from .config import _rate_year_bounds
    if "Utah Property Tax Reference" not in wb.sheetnames:
        return None
    ws = wb["Utah Property Tax Reference"]
    start = None
    for row in ws.iter_rows(min_col=2, max_col=2):
        v = row[0].value
        if isinstance(v, str) and v.strip() == table_title:
            start = row[0].row
            break
    if start is None:
        return None
    sched: dict[int, float] = {}
    r = start + 4                       # title, subtitle, blank, header, then data
    while r <= ws.max_row:
        ylabel = ws.cell(row=r, column=2).value
        rate = ws.cell(row=r, column=3).value
        if ylabel is None or (isinstance(ylabel, str) and not ylabel.strip()):
            break
        if isinstance(ylabel, str) and ylabel.strip().lower() == "sources":
            break
        # Stop if we run into the *next* table's title banner.
        if isinstance(ylabel, str) and ylabel.strip() in (
                _RESIDENTIAL_TABLE_TITLE, _INVENTORY_TABLE_TITLE):
            break
        if isinstance(rate, (int, float)):
            try:
                lo, _hi = _rate_year_bounds(str(ylabel))
                sched[lo] = float(rate)
            except Exception:
                pass
        r += 1
    return sched or None


def _load_vacant_land_schedule(wb) -> dict | None:
    """Editable lot-inventory / nonresidential rate table → {roll_year: rate}."""
    return _load_rate_schedule(wb, _INVENTORY_TABLE_TITLE)


def _load_residential_schedule(wb) -> dict | None:
    """Editable residential (Utah) rate table → {roll_year: rate}."""
    return _load_rate_schedule(wb, _RESIDENTIAL_TABLE_TITLE)


def _write_reference_sheet(ws):
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 70

    r = _write_rate_table(
        ws, 1,
        _RESIDENTIAL_TABLE_TITLE,
        ("EDITABLE — Utah taxes primary residential property on 55% of fair market value "
         "— the 45% exemption of Utah Const. art. XIII, § 3 and Utah Code § 59-2-103, "
         "covering the dwelling and up to one acre of land. The rate has been flat at "
         "55% since 1995. The model READS these yellow cells and applies them to "
         "residential value by roll year (carry-forward past the last row); the "
         "RESID_TAXABLE_RATIO cell on the Inputs tab overrides them when changed."),
        "Residential Taxable Ratio", _RATE_HISTORY, editable=True)

    r += 2
    r = _write_rate_table(
        ws, r,
        _INVENTORY_TABLE_TITLE,
        ("EDITABLE — Utah taxes non-exempt property at 100% of fair market value, but "
         "Utah Admin. Code R884-24P-52 lets the primary residential exemption reach "
         "unoccupied property and property under construction that the assessor determines "
         "will be a primary residence once occupied — so builder lot inventory is carried "
         "at the same 55%. The model READS these yellow rate cells and applies them to "
         "lot-inventory value by roll year (carry-forward past the last row). Enter 100% "
         "to tax inventory at full market value instead."),
        "Lot Inventory Taxable Ratio", _INVENTORY_HISTORY, editable=True)

    # ── Utah property tax calendar ───────────────────────────────────────────
    r += 2
    from .config import UTAH_TAX_CALENDAR
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    tc = ws.cell(row=r, column=2, value="Utah Property Tax Calendar")
    tc.fill = _BLUE; tc.font = _WHITEFONT; tc.alignment = _C
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    sc0 = ws.cell(row=r, column=2, value=
        "Taxes are due in a single payment on 30 November of the roll year, which is why "
        "district principal falls on the following 1 March.")
    sc0.font = _NOTE; sc0.alignment = _C
    r += 2
    for col, lbl in [(2, "Date"), (3, "Event")]:
        c = ws.cell(row=r, column=col, value=lbl)
        c.fill = _HDRF; c.font = _HDRFONT; c.border = _BORDER; c.alignment = _C
    r += 1
    for when, what in UTAH_TAX_CALENDAR:
        dc = ws.cell(row=r, column=2, value=when)
        dc.font = _LBL; dc.border = _BORDER; dc.alignment = _C
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        wc = ws.cell(row=r, column=3, value=what)
        wc.font = _NOTE; wc.alignment = _L
        r += 1

    r += 2
    sc = ws.cell(row=r, column=2, value="Sources")
    sc.font = _HDRFONT; sc.alignment = _L
    r += 1
    for src in _RATE_SOURCES:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=2, value=src)
        c.font = _NOTE; c.alignment = _L
        r += 1
    ws.freeze_panes = "B5"


def _load_debt_structure(wb) -> dict:
    """Parse the Debt Structure sheet into config overrides (None when blank)."""
    if "Debt Structure" not in wb.sheetnames:
        return {}
    ds = wb["Debt Structure"]
    marks = {}
    for row in ds.iter_rows(min_row=1, min_col=2, max_col=2):
        v = row[0].value
        if isinstance(v, str) and v.upper().startswith(("SENIOR BONDS", "REFUNDING BONDS")):
            marks["senior" if v.upper().startswith("SENIOR") else "refunding"] = row[0].row

    def _parse(start):
        if start is None:
            return None, None, None
        coupon, yld = {}, {}
        groups: dict[int, list] = {}
        other = [v for v in marks.values() if v != start]
        r = start + 2
        while r <= ds.max_row:
            if r in marks.values():
                break
            y = ds.cell(row=r, column=2).value
            c = ds.cell(row=r, column=3).value
            yv = ds.cell(row=r, column=4).value
            tf = ds.cell(row=r, column=5).value
            if isinstance(y, (int, float)):
                y = int(y)
                if c:
                    coupon[y] = float(c)
                if tf:
                    groups.setdefault(int(tf), []).append((y, float(yv) if yv else None))
                elif yv:
                    yld[y] = float(yv)
            r += 1
        term_bonds = []
        for tf, members in groups.items():
            years = [yy for yy, _ in members]
            tyield = next((yv for yy, yv in members if yy == tf and yv), None) \
                or next((yv for _, yv in members if yv), None)
            term_bonds.append((min(years), int(tf), tyield))
        return (coupon or None, yld or None, term_bonds or None)

    out = {}
    sc, sy, st = _parse(marks.get("senior"))
    rc, ry, rt = _parse(marks.get("refunding"))
    if sc: out["senior_coupon_scale"] = sc
    if sy: out["senior_yield_scale"] = sy
    if st: out["senior_term_bonds"] = st
    if rc: out["senior_refunding_coupon_scale"] = rc
    if ry: out["senior_refunding_yield_scale"] = ry
    if rt: out["senior_refunding_term_bonds"] = rt
    return out


# ── Read / load ──────────────────────────────────────────────────────────────
def _coerce(value, kind):
    if value is None or value == "":
        return None
    if kind == "str":
        return str(value)
    if kind == "yesno":
        return value  # converted by the caller (Yes/No string vs bool flag)
    if kind == "int":
        return int(round(float(value)))
    if kind in ("float", "pct"):
        return float(value)
    if kind == "date":
        if isinstance(value, datetime):
            return value.date()
        return value
    return value


def load_inputs_workbook(path: str) -> tuple[ModelConfig, DeveloperProjections]:
    """Read an edited Inputs workbook into a (ModelConfig, DeveloperProjections)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Inputs"]

    by_range = {attr: (rng, kind) for _, _, rng, attr, kind, _ in _SPECS}
    rng_to_attr = {rng: (attr, kind) for _, _, rng, attr, kind, _ in _SPECS}

    # Scan the Range Name (col D) / Value (col C) columns.
    overrides_cfg: dict = {}
    overrides_dev: dict = {}
    for row in ws.iter_rows(min_row=5, min_col=3, max_col=4, values_only=True):
        value, rng = row[0], row[1]
        if rng not in rng_to_attr:
            continue
        attr, kind = rng_to_attr[rng]
        coerced = _coerce(value, kind)
        if coerced is None:
            continue
        # Booleans for the strict-biennial flag; Yes/No strings elsewhere.
        if kind == "yesno":
            sval = str(value).strip().lower()
            if attr == "hold_value_flat":
                coerced = sval in ("true", "yes", "y", "1")
            else:
                coerced = "Yes" if sval in ("true", "yes", "y", "1") else "No"
        if attr in _DEV_ATTRS:
            overrides_dev[_DEV_ATTRS[attr]] = coerced
        else:
            overrides_cfg[attr] = coerced

    overrides_cfg.update(_load_debt_structure(wb))   # pricing-day coupon/yield/term scales
    hist_av = _load_historical_av(ws)                 # historical certified-values true-up table
    if hist_av:
        overrides_cfg["historical_av"] = hist_av
    vl_sched = _load_vacant_land_schedule(wb)          # editable lot-inventory rate table
    if vl_sched:
        overrides_cfg["lot_inventory_rate_schedule"] = vl_sched
    res_sched = _load_residential_schedule(wb)         # editable residential (Utah) rate table
    if res_sched:
        overrides_cfg["residential_rate_schedule"] = res_sched
    cfg = ModelConfig(**overrides_cfg)
    dev = _load_dev_sheet(wb, overrides_dev)
    return cfg, dev


def _find_markers(dws):
    """Return {marker_keyword: row} for the Development Inputs section headers."""
    markers = {}
    for row in dws.iter_rows(min_row=1, min_col=2, max_col=2):
        v = row[0].value
        if isinstance(v, str):
            up = v.upper()
            for key in ("PRODUCT SETUP", "LOT DELIVERY", "HOME CLOSINGS",
                        "COMMERCIAL MARKET VALUE"):
                if up.startswith(key):
                    markers[key] = row[0].row
    return markers


def _load_dev_sheet(wb, overrides_dev) -> DeveloperProjections:
    from .development import ProductLine
    dev_kwargs = dict(overrides_dev)
    if "Development Inputs" not in wb.sheetnames:
        return DeveloperProjections(**dev_kwargs)
    dws = wb["Development Inputs"]
    m = _find_markers(dws)
    marker_rows = set(m.values())

    # ── Product setup (parameters as rows, products as columns 3..10) ────────
    names = existing = asps = aspyrs = [None] * _NPROD
    psr = m.get("PRODUCT SETUP")
    if psr:
        for rr in range(psr + 1, psr + 8):
            lbl = dws.cell(row=rr, column=2).value
            if not isinstance(lbl, str):
                continue
            vals = [dws.cell(row=rr, column=3 + i).value for i in range(_NPROD)]
            low = lbl.lower()
            if "product name" in low:
                names = vals
            elif "total unit" in low or "existing" in low:
                existing = vals
            elif "base asp" in low:
                asps = vals
            elif "asp base year" in low:
                aspyrs = vals

    def _read_grid(marker):
        out = [dict() for _ in range(_NPROD)]
        start = m.get(marker)
        if not start:
            return out
        r = start + 2  # skip marker + header
        while r <= dws.max_row:
            if r in marker_rows:
                break
            y = dws.cell(row=r, column=2).value
            if isinstance(y, (int, float)):
                yi = int(y)
                for i in range(_NPROD):
                    v = dws.cell(row=r, column=3 + i).value
                    if v:
                        out[i][yi] = int(round(float(v)))
            r += 1
        return out

    lot_grid = _read_grid("LOT DELIVERY")
    close_grid = _read_grid("HOME CLOSINGS")

    products = []
    for i in range(_NPROD):
        nm = names[i] if names else None
        if not nm:
            continue
        products.append(ProductLine(
            name=str(nm),
            existing_units=int(existing[i]) if (existing and existing[i]) else 0,
            lot_deliveries=lot_grid[i], home_closings=close_grid[i],
            asp_base=float(asps[i]) if (asps and asps[i]) else 515_000.0,
            asp_base_year=int(aspyrs[i]) if (aspyrs and aspyrs[i]) else 2024,
        ))

    # ── Commercial ───────────────────────────────────────────────────────────
    commercial = {}
    cstart = m.get("COMMERCIAL MARKET VALUE")
    if cstart:
        r = cstart + 2
        while r <= dws.max_row:
            if r in marker_rows:
                break
            yr = dws.cell(row=r, column=2).value
            cv = dws.cell(row=r, column=3).value
            if isinstance(yr, (int, float)) and cv:
                commercial[int(yr)] = float(cv)
            r += 1

    if products:
        dev_kwargs["products"] = products
    if commercial:
        dev_kwargs["commercial_market_value"] = commercial
    # No products defined → fall back to the model defaults (plus any overrides).
    return DeveloperProjections(**dev_kwargs)
