"""
workbook.py — writes the Excel output.

Tab list, tab order, row anchors and column positions are the Colorado
metropolitan district template's, unchanged:

     1  Inputs - First                 9  Ops Rev & Exp Projection
     2  Capital Costs                 10  Senior Lien DS - First
     3  Costs of Issuance             11  Sub Lien DS - First (Annual)
     4  Sources and Uses - First      12  Sub Lien DS - First (SA)
     5  Sources and Uses - Refunding  13  Senior Lien DS - Refunding
     6  Summary                       14  CAPI Fund - First
     7  Residential Development       15  Scratch--->>>
     8  Comm Development              16  DBC Output

Every tab keeps the template's four-line title block in rows 1-4, its header
rows, and its anchor rows (development schedules start at row 14, debt service
at row 11, the Summary axis at row 12).  Named ranges are written on
"Inputs - First" at the same cells the template uses, and each one that Utah
re-labels also carries its Colorado alias so the two files stay diff-able.
"""

from __future__ import annotations

from datetime import date
from typing import Any, Iterable, Optional, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from .config import PID_STATUTORY_LEVY_CAP, ModelConfig
from .development import DevelopmentProjections
from .engine import Results

# ── Formatting vocabulary ─────────────────────────────────────────────────────

NAVY = "182957"                      # Tierra navy
LIGHT = "DCE6F1"
INPUT_FILL = "FFF2CC"

FMT_MONEY = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
FMT_MONEY_2 = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"_);_(@_)'
FMT_INT = '#,##0;(#,##0);"-"'
FMT_PCT = '0.00%'
FMT_PCT3 = '0.000%'
FMT_MILLS = '0.000'
FMT_DATE = 'mm/dd/yyyy'
FMT_RATIO = '0.00"x"'
FMT_PRICE = '0.000'

TITLE_FONT = Font(name="Calibri", size=12, bold=True, color=NAVY)
SUB_FONT = Font(name="Calibri", size=10, bold=True, color=NAVY)
HEAD_FONT = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
SECTION_FONT = Font(name="Calibri", size=10, bold=True, underline="single")
BODY_FONT = Font(name="Calibri", size=10)
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
THIN = Side(style="thin", color="BFBFBF")
TOP_BORDER = Border(top=Side(style="thin", color="000000"))
TOTAL_BORDER = Border(top=Side(style="thin", color="000000"),
                      bottom=Side(style="double", color="000000"))


def _put(ws, row: int, col: int, value: Any, fmt: Optional[str] = None,
         font: Optional[Font] = None, fill: Optional[str] = None,
         align: Optional[str] = None, border: Optional[Border] = None,
         wrap: bool = False):
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.font = font or BODY_FONT
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if align or wrap:
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        cell.border = border
    return cell


def _header(ws, row: int, col: int, text: str):
    return _put(ws, row, col, text, font=HEAD_FONT, fill=NAVY,
                align="center", wrap=True)


def _title_block(ws, cfg: ModelConfig, res: Results, col: int = 1) -> None:
    """Rows 1-4 — the same four lines on every tab in the template."""
    _put(ws, 1, col, cfg.district_name, font=TITLE_FONT)
    _put(ws, 2, col, f"City of {cfg.city} / {cfg.county} County", font=SUB_FONT)
    _put(ws, 3, col, cfg.analysis, font=SUB_FONT)
    _put(ws, 4, col,
         f"{cfg.senior_bonds_series} - ${res.senior_par / 1e6:,.2f}MM "
         f"({cfg.dsc_senior_lien_bonds:.2f}x) and {cfg.sub_bonds_series} - "
         f"${res.sub_par / 1e6:,.2f}MM ({cfg.dsc_sub_lien_bonds:.2f}x)",
         font=SUB_FONT)


def _widths(ws, spec: dict[str, float]) -> None:
    for letter, width in spec.items():
        ws.column_dimensions[letter].width = width


# ── Inputs - First ────────────────────────────────────────────────────────────

def _input_rows(cfg: ModelConfig, res: Results) -> list[tuple]:
    """
    (row, label, value, range_name, note, number_format, alias)

    Row numbers are the template's.  `alias` is the Colorado range name kept
    alongside the Utah one.
    """
    c = cfg
    mills_cap = PID_STATUTORY_LEVY_CAP * 1000
    rows: list[tuple] = [
        (5,  "Basic Inputs", None, None, None, None, None),
        (6,  "Public Infrastructure District Name", c.district_name, "PID", None, None, "METRO"),
        (7,  "City", c.city, "CITY", None, None, None),
        (8,  "County", c.county, "COUNTY", None, None, None),
        (9,  "Analysis", c.analysis, "ANALYSIS", None, None, None),
        (10, "Scenario", c.scenario_label, "TITLE4", None, None, None),
        (11, "Developer", c.developer, "DEVELOPER", None, None, None),
        (12, "Senior Bonds", c.senior_bonds_series, "SENIOR_BONDS_SERIES", None, None, None),
        (13, "Subordinate Lien Bonds", c.sub_bonds_series, "SUB_BONDS_SERIES", None, None, None),
        (14, "Refunding Bonds", c.refund_bonds_series, "REFUND_BONDS_SERIES", None, None, None),
        (16, "Second Financing", c.second_financing, "SECOND_FINANCING",
         "Toggle for Second Financing if needed", None, None),
        (17, "Refunding", c.refund_financing, "REFUND_FINANCING", None, None, None),

        (18, "Fees", None, None, None, None, None),
        (19, "Costs of Issuance", c.coi, "COI", None, FMT_MONEY, None),
        (20, "Costs of Issuance - Refunding", c.coi_refunding, "COI_REFUNDING", None, FMT_MONEY, None),
        (21, "Underwriters' Discount - Senior Bonds", c.uwd_senior, "UWD_SENIOR", None, FMT_PCT, None),
        (22, "Underwriters' Discount - Subordinate Lien / Cashflow Notes",
         c.uwd_sub, "UWD_SUB", None, FMT_PCT, None),
        (23, "Underwriters' Discount - Refunding Senior Bonds",
         c.uwd_senior_refunding, "UWD_SENIOR_REFUNDING", None, FMT_PCT, None),
        (24, "County Collection Cost", c.county_treasurer_fee, "COUNTY_COLLECTION_FEE",
         "Utah recovers assessing & collecting through a separate statewide levy "
         "(UCA 59-2-1602), not a haircut on the district's distribution",
         FMT_PCT, "COUNTY_TREASURER_FEE"),
        (25, "Trustee Fee (Senior Bonds)", c.trustee_fee, "TRUSTEE_FEE", None, FMT_MONEY, None),
        (26, "Trustee Fee (Subordinate Bonds)", c.trustee_fee_sub, "TRUSTEE_FEE_SUB", None, FMT_MONEY, None),

        (29, "Structuring Assumptions", None, None, None, None, None),
        (30, "Delivery", c.delivery, "DELIVERY", None, FMT_DATE, None),
        (31, "Delivery - Refunding", c.delivery_refunding, "DELIVERY_REFUNDING", None, FMT_DATE, None),
        (32, "First Interest Senior Bonds", c.first_int, "FIRST_INT", None, FMT_DATE, None),
        (33, "First Interest Subordinate Bonds", c.first_int_sub, "FIRST_INT_SUB", None, FMT_DATE, None),
        (34, "First Interest Refunding Senior Bonds", c.first_int_refunding,
         "FIRST_INT_REFUNDING", None, FMT_DATE, None),
        (35, "First Principal", res.senior_stats.first_maturity, "FIRST_PRIN", None, FMT_DATE, None),
        (36, "Fractional Period Senior Bonds", c.frac, "FRAC", None, '0.0000', None),
        (37, "Fractional Period Refunding Senior Bonds", 0.0, "FRAC_REFUNDING", None, '0.0000', None),
        (38, "Fractional Period Subordinate Bonds", c.frac_sub, "FRAC_SUB", None, '0.0000', None),
        (39, "Fractional Period CAPI", c.frac, "FRAC_CAPI", None, '0.0000', None),
        (40, "Capitalized Interest First Draw", c.capi_first_draw, "CAPI_FIRST_DRAW", None, FMT_DATE, None),
        (41, "Capitalized Interest End Date", c.capi_end_date, "CAPI_END_DATE", None, FMT_DATE, None),
        (42, "Capitalized Interest", c.capi, "CAPI", None, None, None),
        (43, "Capitalized Interest Period (Months)", c.capi_term_months, "CAPI_TERM", None, FMT_INT, None),
        (44, "Surplus Toggle", c.surplus_on_off, "SURPLUS_ON_OFF", None, None, None),
        (45, "Deposit to Surplus Fund / Debt Service Reserve Fund",
         res.surplus_fund_deposit, "SURPLUS_FUND_DEPOSIT",
         "Least of 10% of par, 125% of average annual debt service, and MADS",
         FMT_MONEY, None),
        (46, "Surplus Fund Release Date - Senior Bonds", c.surplus_fund_release_date,
         "SURPLUS_FUND_RELEASE_DATE", None, FMT_DATE, None),
        (47, "Surplus Target Amount - Senior Bonds", res.surplus_fund_target,
         "SURPLUS_FUND_TARGET", None, FMT_MONEY, None),
        (48, "Release Surplus Funds?", c.surplus_release_sizing, "SURPLUS_RELEASE_SIZING", None, None, None),
        (49, "Deposit to Surplus Fund - Refunding Senior Bonds", 0.0,
         "SURPLUS_FUND_DEPOSIT_REFUNDING", None, FMT_MONEY, None),
        (50, "Surplus Fund Release Date - Refunding Bonds",
         c.surplus_fund_release_date_refunding, "SURPLUS_FUND_RELEASE_DATE_REFUNDING", None, FMT_DATE, None),
        (51, "Surplus Target Amount - Refunding Senior Bonds", 0.0,
         "SURPLUS_FUND_TARGET_REFUNDING", None, FMT_MONEY, None),
        (52, "Ending Balance of Accrued Interest for Subordinate Bonds",
         c.end_bal_accrued_sub_bonds, "END_BAL_ACCRUED_SUB_BONDS", None, FMT_DATE, None),
        (53, "Subordinate Lien Bonds Sizing Threshold", c.sub_sizing_threshold,
         "SUB_SIZING_THRESHOLD", None, FMT_MONEY, None),
        (54, "Premium Call Date", c.premium_call_first, "PREMIUM_CALL_FIRST", None, FMT_DATE, None),
        (55, "Premium Call Price", c.premium_call_price, "PREMIUM_CALL_FIRST_PRICE", None, '0', None),
        (56, "Par Call Date", c.par_call_first, "PAR_CALL_FIRST", None, FMT_DATE, None),
        (57, "Final Maturity for Senior Lien Bonds", c.final_mat_yrs, "FINAL_MAT_YRS", None, FMT_INT, None),
        (58, "Final Maturity for Subordinate Lien Bonds", c.final_mat_sub_yrs,
         "FINAL_MAT_SUB_YRS", None, FMT_INT, None),
        (59, "Final Maturity for Refunding Senior Lien Bonds", c.final_mat_yrs_refunding,
         "FINAL_MAT_YRS_REFUNDING", None, FMT_INT, None),
        (60, "Principal Maturity (Month)", c.prin_maturity, "PRIN_MATURITY",
         "Utah property taxes are due 30 November, so principal falls on 1 March",
         FMT_INT, None),
        (61, "Interest Maturity (Month)", c.int_maturity, "INT_MATURITY", None, FMT_INT, None),
        (62, "Senior Lien Bonds Principal Pay Day", c.prin_maturity_day_senior,
         "PRIN_MATURITY_DAY_SENIOR", None, FMT_INT, None),
        (63, "Subordinate Lien Bonds Principal Pay Day", c.prin_maturity_day_sub,
         "PRIN_MATURITY_DAY_SUB", None, FMT_INT, None),
        (64, "Investment Grade Rating", c.ig_rated, "IG_RATED", None, None, None),
        (65, "Senior Lien Bonds Interest Rate", c.senior_interest_rate,
         "SENIOR_INTEREST_RATE", None, FMT_PCT3, None),
        (66, "Senior Lien Bonds Interest Rate - Refunding", c.senior_refunding_interest_rate,
         "SENIOR_REFUNDING_INTEREST_RATE", None, FMT_PCT3, None),
        (67, "Subordinate Lien / Cashflow Notes Interest Rate", c.sub_interest_rate,
         "SUB_INTEREST_RATE", None, FMT_PCT3, None),
        (68, "Debt Service Coverage - Senior Lien Bonds", c.dsc_senior_lien_bonds,
         "DSC_SENIOR_LIEN_BONDS", None, FMT_RATIO, None),
        (69, "Debt Service Coverage - Subordinate Lien Bonds", c.dsc_sub_lien_bonds,
         "DSC_SUB_LIEN_BONDS", None, FMT_RATIO, None),
        (70, "Debt Service Coverage - Refunding", c.dsc_refunding_bonds,
         "DSC_REFUNDING_BONDS", None, FMT_RATIO, None),

        (71, 'First Year for "Summary"', c.first_year, "FIRST_YEAR", None, '0', None),
        (72, "Residential Delivery", c.resid_delivery_year, "RESID_DELIVERY_YEAR", None, FMT_DATE, None),
        (73, "Inflation Rate", c.inflation_rate, "INFLATION_RATE", None, FMT_PCT, None),
        (74, "Inflation Rate Commercial Sales", c.inflation_rate_comm_sales,
         "INFLATION_RATE_COMM_SALES", None, FMT_PCT, None),
        (75, "Inflation Step (in years)", c.inflation_step_years, "INFLATION_STEP_YEARS", None, FMT_INT, None),
        (76, "Inflation Step Year", c.inflation_step_start_year, "INFLATION_STEP_START_YEAR", None, FMT_DATE, None),
        (77, "Reassessment Rate - Residential (Senior)", c.reassess_rate_resid,
         "REASSESS_RATE", 'Also known as "Existing Home Growth"', FMT_PCT, "BIENNIAL_REASSESS_RATE"),
        (78, "Reassessment Rate - Residential (Subordinate)", c.reassess_rate_resid_sub,
         "REASSESS_RATE_SUBORDINATE", 'Also known as "New Home Growth"', FMT_PCT,
         "BIENNIAL_REASSESS_RATE_SUBORDINATE"),
        (79, "Reassessment Rate - Commercial", c.reassess_rate_comm,
         "REASSESS_COMM_RATE", None, FMT_PCT, "BIENNIAL_REASSESS_COMM_RATE"),
        (80, "Prior Residential Taxable Ratio", c.resid_taxable_ratio_prior,
         "RESID_TAXABLE_RATIO_PRIOR",
         "Colorado: prior Gallagher assessment rate. Inert in Utah.",
         FMT_PCT, "TABOR_PRIOR"),
        (81, "Primary Residential Taxable Ratio", c.resid_taxable_ratio,
         "RESID_TAXABLE_RATIO",
         "UCA 59-2-103 — 45% primary residential exemption, so 55% of fair "
         "market value is taxable", FMT_PCT, "TABOR_CURRENT"),
        (82, "Projected Residential Taxable Ratio", c.resid_taxable_ratio,
         "RESID_TAXABLE_RATIO_PROJECTED", None, FMT_PCT, "TABOR_PROJECTED"),
        (83, "Property Tax Collection %", c.tax_collect_mill_prc, "TAX_COLLECT_MILL_PRC", None, FMT_PCT, None),
        (84, "Personal Property Uniform Fee %", c.uniform_fee_prc, "UNIFORM_FEE_PRC",
         "UCA 59-2-405 — distributed to taxing entities in the same proportion "
         "as ad valorem real property tax", FMT_PCT, "TAX_COLLECT_SO_PRC"),
        (85, "Uniform Fee Taxable Value Threshold", c.uniform_fee_av_threshold,
         "UNIFORM_FEE_AV_THRESHOLD", None, FMT_MONEY, "TAX_COLLECT_SO_AV_THRESHOLD"),
        (86, "Interest Earnings Rate", c.interest_earn_rate, "INTEREST_EARN_RATE", None, FMT_PCT, None),
        (87, "Gallagherize the Rate", c.gallagherization, "GALLAGHERIZATION",
         "Colorado mechanism — Utah caps are a fixed rate per dollar of value", None, None),
        (88, "Mill Levy from Governing Document", c.mill_levy_governing_doc,
         "MILL_LEVY_GOVERNING_DOC",
         f"UCA 17D-4-303 caps a PID levy at {mills_cap:,.3f} mills "
         f"(0.015 per dollar of taxable value)", FMT_MILLS, "MILL_LEVY_SERVICE_PLAN"),
        (89, "Mill Levy Target - Debt Service", c.mill_levy_ds_target,
         "MILL_LEVY_DS_TARGET", None, FMT_MILLS, None),
        (90, "Mill Levy Cap - Debt Service", c.mill_levy_ds_cap, "MILL_LEVY_DS_CAP", None, FMT_MILLS, None),
        (91, "Mill Levy Cap - Commercial", c.mill_levy_comm, "MILL_LEVY_COMM", None, FMT_MILLS, None),
        (92, "Mill Levy Target - Operations", c.mill_levy_ops_target, "MILL_LEVY_OPS_TARGET", None, FMT_MILLS, None),
        (93, "Total Mill Levy Cap", c.mill_levy_cap_total, "MILL_LEVY_CAP_TOTAL", None, FMT_MILLS, None),
        (94, "Mill Contribution Rate to Town/City/County", c.contribution_rate,
         "CONTRIBUTION_RATE", None, FMT_MILLS, None),
        (95, "Annual District Administration Costs", c.admin_cost_base, "ADMIN_COST_BASE",
         "Colorado books this as the O&M carveout", FMT_MONEY, "OM_CARVEOUT"),
        (96, "Administration Cost Taxable Value Limit", c.admin_cost_av_limit,
         "ADMIN_COST_AV_LIMIT", None, FMT_MONEY, "OM_CARVEOUT_AV_LIMIT"),
        (97, "Administration Cost Growth", c.admin_cost_growth, "ADMIN_COST_GROWTH", None, FMT_PCT, "OM_GROWTH_RATE"),
        (98, "System Development Fee", c.system_development_fee, "SYSTEM_DEVELOPMENT_FEE",
         "Per residential unit", FMT_MONEY, "CAPITAL_IMPROV_FEE"),
        (99, "State Assessed Valuation", c.state_assessed, "STATE_ASSESSED", None, FMT_MONEY, None),
        (100, "Centrally Assessed Valuation", c.centrally_assessed_value,
         "CENTRALLY_ASSESSED_VALUE",
         "UCA 59-2-201 — utilities, railroads, mines assessed by the State Tax "
         "Commission", FMT_MONEY, "OIL_GAS_ASSETS"),
        (101, "New Residential Value (Senior)", c.resid_new_value_add, "RESID_NEW_VALUE_ADD", None, None, None),
        (102, "New Residential Value (Subordinate)", c.resid_new_value_add_sub,
         "RESID_NEW_VALUE_ADD_SUB", None, None, None),
        (103, "New Commercial Value (Senior)", c.comm_new_value_add, "COMM_NEW_VALUE_ADD", None, None, None),
        (104, "New Commercial Value (Subordinate)", c.comm_new_value_add_sub,
         "COMM_NEW_VALUE_ADD_SUB", None, None, None),
        (105, "Centrally Assessed Property (Senior)", c.centrally_assessed_senior,
         "CENTRALLY_ASSESSED_SENIOR", None, None, "OIL_GAS"),
        (106, "Centrally Assessed Property (Subordinate)", c.centrally_assessed_sub,
         "CENTRALLY_ASSESSED_SUB", None, None, "OIL_GAS_SUBORDINATE"),
        (107, "Reassessment Frequency", c.reassess_frequency, "REASSESS_FREQUENCY",
         "Utah county assessors revalue annually (UCA 59-2-303.1)", None, None),
        (108, "Value Lag (Years)", c.value_lag_years, "VALUE_LAG_YEARS",
         "Utah assesses 1 January and bills the same year — one year. "
         "Colorado's biennial cycle runs two.", FMT_INT, None),

        (111, "Developer Assumptions", None, None, None, None, None),
    ]
    for i in range(12):
        rows.append((113 + i, f"Specific ASP for SFD_{i + 1}",
                     c.asp[i] if i < len(c.asp) else 0.0, f"ASP_SFD_{i + 1}",
                     c.product_labels[i] if i < len(c.product_labels) else None,
                     FMT_MONEY, None))
    rows += [
        (125, "Hypothetical Sizing Scenario", c.hypothetical_scenario, "HYPOTHETICAL_SCENARIO", None, None, None),
        (126, "Lot Delivery Scenario", c.lot_delivery_scenario, "LOT_DELIVERY_SCENARIO", None, FMT_PCT, None),
        (127, "Absorption Scenario", c.absorption_scenario, "ABSORPTION_SCENARIO", None, FMT_PCT, None),
        (128, "First Home Closing", c.home_first_close_date, "HOME_FIRST_CLOSE_DATE", None, FMT_DATE, None),
        (129, "First Townhouse Closing", c.home_first_close_date, "TOWNHOME_FIRST_CLOSE_DATE", None, FMT_DATE, None),
        (130, "First Duplex Closing", c.home_first_close_date, "DUPLEX_FIRST_CLOSE_DATE", None, FMT_DATE, None),
        (131, "Commercial Lag Years (Value for Platted/Developed Lots)",
         c.commercial_lag_years, "COMMERCIAL_LAG_YEARS", None, FMT_INT, None),
        (132, "Home Lot Delivery Date", c.home_lot_delivery_date, "HOME_LOT_DELIVERY_DATE",
         "Also known as Home Starts", FMT_DATE, None),
        (133, "Platted Commercial Lot Value", c.platted_comm_lot_value,
         "PLATTED_COMM_LOT_VALUE", None, FMT_PCT, None),
        (134, "Platted Lot Value", c.platted_lot_value, "PLATTED_LOT_VALUE",
         "Finished-lot market value as a share of home ASP", FMT_PCT, None),
        (135, "Developed Lot Taxable Ratio", c.developed_lot_value, "DEVELOPED_LOT_VALUE",
         "Utah Admin. Code R884-24P-52 lets the residential exemption reach "
         "unoccupied property the assessor finds will become a primary residence",
         FMT_PCT, None),
        (136, "Centrally Assessed Taxable Ratio", c.centrally_assessed_ratio,
         "CENTRALLY_ASSESSED_RATIO", None, FMT_PCT, "OIL_GAS_VALUE"),
    ]
    return rows


def _write_inputs(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Inputs - First")
    _widths(ws, {"A": 3, "B": 52, "C": 26, "D": 34, "E": 62, "F": 12, "G": 12, "H": 12})

    _put(ws, 1, 2, "First Financing Assumptions", font=TITLE_FONT)
    for col, text in ((2, "Title"), (3, "Description"), (4, "Range Name"), (5, "Notes")):
        _header(ws, 3, col, text)

    for row, label, value, name, note, fmt, alias in _input_rows(cfg, res):
        if value is None and name is None:
            _put(ws, row, 2, label, font=SECTION_FONT)
            continue
        _put(ws, row, 2, label)
        cell = _put(ws, row, 3, value, fmt=fmt, fill=INPUT_FILL, align="center")
        _put(ws, row, 4, name, font=Font(name="Consolas", size=9, color="808080"))
        if note:
            _put(ws, row, 5, note, font=Font(name="Calibri", size=8, italic=True,
                                             color="808080"), wrap=True)
        ref = f"'Inputs - First'!$C${row}"
        for n in filter(None, (name, alias)):
            try:
                wb.defined_names.add(DefinedName(n, attr_text=ref))
            except (ValueError, KeyError):
                pass
        _ = cell

    ws.freeze_panes = "B4"


# ── Capital Costs / Costs of Issuance ─────────────────────────────────────────

def _write_capital_costs(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Capital Costs")
    _widths(ws, {"A": 3, "B": 3, "C": 3, "D": 62, "E": 20})
    lines = [
        ("Total Eligible Construction Costs (1)", res.total_reimbursement),
        ("Total Eligible Soft Costs (2)", 0.0),
        ("Approx. Advanced O&M Dollars to District", 0.0),
        ("Future O&M Costs Advance to District at Bond Close", 0.0),
        ("Developer Advance to District for Future Const. Costs (3)", 0.0),
        ("Approx. Developer Interest on Advances (4)", 0.0),
        ("Estimated Bond Formation/Closing Costs (5)",
         cfg.coi + res.senior_uwd + res.sub_uwd),
    ]
    for i, (label, value) in enumerate(lines):
        _put(ws, 4 + i, 4, label)
        _put(ws, 4 + i, 5, value, fmt=FMT_MONEY)
    _put(ws, 11, 4, "Total Bond PAR Amount", font=SUB_FONT, border=TOP_BORDER)
    _put(ws, 11, 5, sum(v for _, v in lines), fmt=FMT_MONEY,
         font=SUB_FONT, border=TOTAL_BORDER)


def _write_costs_of_issuance(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Costs of Issuance")
    _widths(ws, {"A": 3, "B": 34, "C": 12, "D": 12, "E": 16, "F": 12, "G": 12, "H": 18})
    _put(ws, 2, 2, "Detailed Cost of Issuance", font=TITLE_FONT)
    _header(ws, 3, 5, "Total")
    _header(ws, 3, 6, "% of Par")
    items = ["Bond Insurance", "Surety", "Bond Counsel", "Disclosure Counsel",
             "Underwriter's Counsel", "General Counsel", "District Administrator",
             "Municipal Advisor", "Market Study", "Revenue Forecast", "Ratings",
             "Escrow + Trustee", "Verification", "District Photography",
             "Printer", "Title / Recording", "Misc"]
    par = res.total_par or 1.0
    for i, name in enumerate(items):
        r = 4 + i
        _put(ws, r, 2, name)
        _put(ws, r, 5, 0.0, fmt=FMT_MONEY, fill=INPUT_FILL)
        _put(ws, r, 6, 0.0, fmt=FMT_PCT)
    _put(ws, 21, 2, "Total", font=SUB_FONT, border=TOP_BORDER)
    _put(ws, 21, 5, cfg.coi, fmt=FMT_MONEY, font=SUB_FONT, border=TOTAL_BORDER)
    _put(ws, 21, 6, cfg.coi / par, fmt=FMT_PCT, font=SUB_FONT, border=TOTAL_BORDER)
    _put(ws, 23, 2, "Total cost of issuance is carried on 'Inputs - First' (COI); "
                    "this tab is the line-item build-up.",
         font=Font(name="Calibri", size=8, italic=True, color="808080"))


# ── Sources and Uses ──────────────────────────────────────────────────────────

def _write_sources_uses(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Sources and Uses - First")
    _widths(ws, {"A": 46, "B": 20, "C": 3, "D": 22, "E": 3, "F": 20,
                 "G": 3, "H": 14, "I": 14})
    _title_block(ws, cfg, res)
    lots = res.dev.total_units() or 1
    _put(ws, 4, 1,
         f"Scenario: {'Investment Grade Rating' if cfg.ig_rated == 'Yes' else 'Non-Rated'}"
         f" / {cfg.reassess_rate_resid_sub:.2%} Reassessment Rate"
         f" / {lots:,} Lots / Est. Per Lot Reimbursement: "
         f"${res.reimbursement_per_lot:,.0f}", font=SUB_FONT)
    _put(ws, 6, 1, f"---{cfg.senior_bonds_series[:-1]} Financing---", font=SUB_FONT)

    _put(ws, 8, 1, "Sources and Uses", font=SECTION_FONT)
    _header(ws, 10, 2, f"Senior Lien Bonds - {cfg.senior_bonds_series}")
    _header(ws, 10, 4, f"Subordinate Lien Cashflow Bonds - {cfg.sub_bonds_series}")
    _header(ws, 10, 6, "Total")

    sources = [
        (12, "Par Amount of Bonds:", res.senior_par, res.sub_par),
        (13, "Plus: Premium / (Discount)", res.senior_stats.premium, 0.0),
        (14, "Other Sources of Funds:", 0.0, 0.0),
        (15, "Existing Senior Surplus Fund Balance (1)", 0.0, 0.0),
    ]
    for r, label, a, b in sources:
        _put(ws, r, 1, label)
        _put(ws, r, 2, a, fmt=FMT_MONEY)
        _put(ws, r, 4, b, fmt=FMT_MONEY)
        _put(ws, r, 6, a + b, fmt=FMT_MONEY)
    ts_a = sum(s[2] for s in sources)
    ts_b = sum(s[3] for s in sources)
    _put(ws, 17, 1, "TOTAL SOURCES OF FUNDS:", font=SUB_FONT, border=TOP_BORDER)
    for col, val in ((2, ts_a), (4, ts_b), (6, ts_a + ts_b)):
        _put(ws, 17, col, val, fmt=FMT_MONEY, font=SUB_FONT, border=TOTAL_BORDER)

    _put(ws, 20, 1, "Uses of Funds", font=SECTION_FONT)
    _header(ws, 22, 2, f"Senior Lien Bonds - {cfg.senior_bonds_series}")
    _header(ws, 22, 4, f"Subordinate Lien Cashflow Bonds - {cfg.sub_bonds_series}")
    _header(ws, 22, 6, "Total")

    uses = [
        (24, "Estimated Reimbursement Amount", res.senior_reimbursement, res.sub_reimbursement),
        (25, "Refunding Escrow Deposits", 0.0, 0.0),
        (26, "Refinance Prior Obligations (Est.)", 0.0, 0.0),
        (27, "Project Fund Deposit", 0.0, 0.0),
        (28, "Debt Service Reserve / Surplus Fund", res.surplus_fund_deposit, 0.0),
        (29, "Capitalized Interest", res.capitalized_interest_deposit, 0.0),
        (30, "Underwriters' Discount", res.senior_uwd, res.sub_uwd),
        (31, "Costs of Issuance", cfg.coi, 0.0),
    ]
    for r, label, a, b in uses:
        _put(ws, r, 1, label)
        _put(ws, r, 2, a, fmt=FMT_MONEY)
        _put(ws, r, 4, b, fmt=FMT_MONEY)
        _put(ws, r, 6, a + b, fmt=FMT_MONEY)
    tu_a = sum(u[2] for u in uses)
    tu_b = sum(u[3] for u in uses)
    _put(ws, 33, 1, "TOTAL USES OF FUNDS:", font=SUB_FONT, border=TOP_BORDER)
    for col, val in ((2, tu_a), (4, tu_b), (6, tu_a + tu_b)):
        _put(ws, 33, col, val, fmt=FMT_MONEY, font=SUB_FONT, border=TOTAL_BORDER)

    _put(ws, 36, 1, "Key Assumptions:", font=SECTION_FONT)
    capi_months = round((cfg.capi_end_date - cfg.delivery).days / 30.44)
    key = [
        (38, "Delivery Date", cfg.delivery, cfg.delivery, FMT_DATE),
        (39, "First Interest Date", cfg.first_int, cfg.first_int_sub, FMT_DATE),
        (40, "First Maturity Date", res.senior_stats.first_maturity,
         res.sub_stats.first_maturity, FMT_DATE),
        (41, "Final Maturity Date", res.senior_stats.final_maturity,
         res.sub_stats.final_maturity, FMT_DATE),
        (42, "First Par Call Date", cfg.par_call_first, None, FMT_DATE),
        (43, f"Capitalized Interest Period ({capi_months}mos)", cfg.capi_end_date, None, FMT_DATE),
        (44, "Debt Service Coverage", cfg.dsc_senior_lien_bonds, cfg.dsc_sub_lien_bonds, FMT_RATIO),
    ]
    for r, label, a, b, fmt in key:
        _put(ws, r, 1, label)
        if a is not None:
            _put(ws, r, 2, a, fmt=fmt)
        if b is not None:
            _put(ws, r, 4, b, fmt=fmt)
    _put(ws, 45, 1, "Annual Reassessment")
    _put(ws, 45, 6, cfg.reassess_rate_resid, fmt=FMT_PCT)
    _put(ws, 46, 1, "Repayment Ratio (Expected)")
    _put(ws, 46, 6, res.repayment_ratio, fmt=FMT_RATIO)
    _put(ws, 47, 1, "Senior Lien Bonds Surplus Fund Target")
    _put(ws, 47, 2, res.surplus_fund_target, fmt=FMT_MONEY)

    _put(ws, 52, 1, "Bond Statistics:", font=SECTION_FONT)
    stats = [
        (54, "8038 Average Life", res.senior_stats.average_life,
         res.sub_stats.average_life, '0.00'),
        (55, "Arbitrage TIC", res.senior_stats.arbitrage_tic,
         res.sub_stats.arbitrage_tic, FMT_PCT3),
        (56, "All-in TIC", res.senior_stats.all_in_tic, res.sub_stats.all_in_tic, FMT_PCT3),
        (57, "Maximum Annual Debt Service", res.senior_stats.max_annual_debt_service,
         res.sub_stats.max_annual_debt_service, FMT_MONEY),
        (58, "Total Debt Service", res.senior_stats.total_debt_service,
         res.sub_stats.total_debt_service, FMT_MONEY),
    ]
    for r, label, a, b, fmt in stats:
        _put(ws, r, 1, label)
        _put(ws, r, 2, a, fmt=fmt)
        _put(ws, r, 4, b, fmt=fmt)

    _put(ws, 60, 1, "Taxing Authority and Fee Assumptions:", font=SECTION_FONT)
    taxing = [
        (62, "Primary Residential Exemption (UCA 59-2-103)", 1 - cfg.resid_taxable_ratio, FMT_PCT),
        (63, "Residential Taxable Ratio", cfg.resid_taxable_ratio, FMT_PCT),
        (64, "Developed Lot Taxable Ratio", cfg.developed_lot_value, FMT_PCT),
        (66, "Statutory Levy Cap (UCA 17D-4-303)", PID_STATUTORY_LEVY_CAP * 1000, FMT_MILLS),
        (67, "Governing Document Mill Levy Cap", cfg.mill_levy_governing_doc, FMT_MILLS),
        (68, "Maximum Adjusted Mill Cap", cfg.mill_levy_ds_cap, FMT_MILLS),
        (69, "Targeted Mill Levy", cfg.mill_levy_ds_target, FMT_MILLS),
        (71, "Property Tax Collection Rate", cfg.tax_collect_mill_prc, FMT_PCT),
        (72, "Personal Property Uniform Fee (UCA 59-2-405)", cfg.uniform_fee_prc, FMT_PCT),
        (74, "County Collection Cost", cfg.county_treasurer_fee, FMT_PCT),
        (75, "Annual Trustee Fee", cfg.trustee_fee, FMT_MONEY),
    ]
    for r, label, value, fmt in taxing:
        _put(ws, r, 1, label)
        _put(ws, r, 6, value, fmt=fmt)


def _write_sources_uses_refunding(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Sources and Uses - Refunding")
    _widths(ws, {"A": 46, "B": 20, "C": 3, "D": 22, "E": 3, "F": 20})
    _title_block(ws, cfg, res)
    _put(ws, 5, 1, f"---{cfg.refund_bonds_series} Refunding Financing---", font=SUB_FONT)
    if cfg.refund_financing != "Yes":
        _put(ws, 7, 1, 'Refunding is toggled off (REFUND_FINANCING = "No"). '
                       "The tab is retained so the layout matches the template.",
             font=Font(name="Calibri", size=10, italic=True, color="808080"))

    _put(ws, 7, 1, "Sources and Uses", font=SECTION_FONT) if cfg.refund_financing == "Yes" else None
    _header(ws, 9, 2, f"Refunding Senior Lien Bonds - {cfg.refund_bonds_series}")
    _header(ws, 9, 4, f"Subordinate Lien Bonds - {cfg.sub_bonds_series}")
    _header(ws, 9, 6, "Total")
    ref_par = res.refunding_stats.par
    rows = [
        (11, "Par Amount of Bonds:", ref_par, 0.0),
        (12, "Plus: Premium / (Discount)", res.refunding_stats.premium, 0.0),
        (14, "Funds on Hand", 0.0, 0.0),
    ]
    for r, label, a, b in rows:
        _put(ws, r, 1, label)
        _put(ws, r, 2, a, fmt=FMT_MONEY)
        _put(ws, r, 4, b, fmt=FMT_MONEY)
        _put(ws, r, 6, a + b, fmt=FMT_MONEY)
    total_sources = sum(r[2] + r[3] for r in rows)
    _put(ws, 16, 1, "TOTAL SOURCES OF FUNDS:", font=SUB_FONT, border=TOP_BORDER)
    _put(ws, 16, 6, total_sources, fmt=FMT_MONEY, font=SUB_FONT, border=TOTAL_BORDER)

    _put(ws, 19, 1, "Uses of Funds", font=SECTION_FONT)
    escrow = sum(r.principal for r in res.senior
                 if r.payment_date > cfg.par_call_first) if cfg.refund_financing == "Yes" else 0.0
    uwd = cfg.uwd_senior_refunding * ref_par
    coi = cfg.coi_refunding if cfg.refund_financing == "Yes" else 0.0
    uses = [
        (23, "Estimated Reimbursement Amount", total_sources - escrow - uwd - coi),
        (24, "Refunding Escrow Deposits", escrow),
        (25, "Underwriters' Discount", uwd),
        (26, "Costs of Issuance", coi),
    ]
    for r, label, value in uses:
        _put(ws, r, 1, label)
        _put(ws, r, 2, value, fmt=FMT_MONEY)
        _put(ws, r, 6, value, fmt=FMT_MONEY)
    _put(ws, 28, 1, "TOTAL USES OF FUNDS:", font=SUB_FONT, border=TOP_BORDER)
    _put(ws, 28, 6, sum(u[2] for u in uses), fmt=FMT_MONEY, font=SUB_FONT, border=TOTAL_BORDER)

    _put(ws, 31, 1, "Key Assumptions:", font=SECTION_FONT)
    for r, label, value, fmt in [
        (33, "Delivery Date", cfg.delivery_refunding, FMT_DATE),
        (34, "First Interest Date", cfg.first_int_refunding, FMT_DATE),
        (35, "First Maturity Date", res.refunding_stats.first_maturity, FMT_DATE),
        (36, "Final Maturity Date", res.refunding_stats.final_maturity, FMT_DATE),
        (38, "Debt Service Coverage", cfg.dsc_refunding_bonds, FMT_RATIO),
    ]:
        _put(ws, r, 1, label)
        if value is not None:
            _put(ws, r, 2, value, fmt=fmt)
    _put(ws, 44, 1, "Bond Statistics:", font=SECTION_FONT)
    for r, label, value, fmt in [
        (46, "8038 Average Life", res.refunding_stats.average_life, '0.00'),
        (47, "Arbitrage TIC", res.refunding_stats.arbitrage_tic, FMT_PCT3),
        (48, "All-in TIC", res.refunding_stats.all_in_tic, FMT_PCT3),
        (49, "Maximum Annual Debt Service", res.refunding_stats.max_annual_debt_service, FMT_MONEY),
        (50, "Total Debt Service", res.refunding_stats.total_debt_service, FMT_MONEY),
    ]:
        _put(ws, r, 1, label)
        _put(ws, r, 2, value, fmt=fmt)


# ── Summary ───────────────────────────────────────────────────────────────────

SUMMARY_COLUMNS: list[tuple[str, str, str, Optional[str]]] = [
    # (column letter, group, header, attribute)
    ("C", "", "AV Set / Assessment Year", "assessment_date"),
    ("D", "", "Tax Revenue Year", "tax_revenue_date"),
    ("E", "Existing Homes", "Home Market Value Reassessment", "existing_reassessment"),
    ("F", "Existing Homes", "Existing Home Market Value", "existing_market_value"),
    ("G", "Existing Homes", "Taxable Value", "existing_taxable_value"),
    ("H", "New Lots", "Lot Delivery Schedule", "lot_units"),
    ("I", "New Lots", "Market Value of Lots (with lag)", "lot_market_value"),
    ("J", "New Lots", "Taxable Value of Lots", "lot_taxable_value"),
    ("K", "New Homes", "Total Residential Units", "residential_units"),
    ("L", "New Homes", "Home Market Value Reassessment", "new_home_reassessment"),
    ("M", "New Homes", "Home Market Value (with lag)", "new_home_market_value"),
    ("N", "New Homes", "Taxable Value", "new_home_taxable_value"),
    ("P", "New Homes", "Value in Collection Year", "prior_roll_value"),
    ("Q", "New Homes", "Total Residential Taxable Value", "residential_taxable_value"),
    ("R", "Commercial - Platted/Developed Lots", "Cumulative Sq. Ft. Delivered", "comm_sf_delivered"),
    ("S", "Commercial - Platted/Developed Lots", "Commercial Reassessment", "comm_reassessment_platted"),
    ("T", "Commercial - Platted/Developed Lots", "Cumulative Market Value", "comm_market_value_platted"),
    ("U", "Commercial - Platted/Developed Lots", "Taxable Value (lagged)", "comm_taxable_value_platted"),
    ("V", "Commercial - Sold/Leased Up", "Total Commercial Sq. Ft. Sold/Leased", "comm_sf_sold"),
    ("W", "Commercial - Sold/Leased Up", "Commercial Reassessment", "comm_reassessment_sold"),
    ("X", "Commercial - Sold/Leased Up", "Cumulative Market Value", "comm_market_value_sold"),
    ("Y", "Commercial - Sold/Leased Up", "Taxable Value (lagged)", "comm_taxable_value_sold"),
    ("Z", "Commercial - Sold/Leased Up", "Total Commercial Taxable Value", "comm_taxable_value"),
    ("AA", "Centrally Assessed Property", "Market Value", "centrally_assessed_market"),
    ("AD", "Centrally Assessed Property", "Taxable Value", "centrally_assessed_taxable"),
    ("AE", "State Assessed", "Taxable Value", "state_assessed_taxable"),
    ("AG", "Revenue Calculation - Senior Bonds", "Total Taxable Value", "senior_taxable_value"),
    ("AH", "Revenue Calculation - Senior Bonds", "District Debt Service Mill Levy", "senior_mill_levy"),
    ("AI", "Revenue Calculation - Senior Bonds", "Mill Levy Collections", "senior_levy_collections"),
    ("AJ", "Revenue Calculation - Senior Bonds", "Personal Property Uniform Fee", "senior_uniform_fee"),
    ("AK", "Revenue Calculation - Senior Bonds", "Commercial Taxable Value", "senior_comm_taxable_value"),
    ("AL", "Revenue Calculation - Senior Bonds", "Commercial Mill Levy", "senior_comm_mill_levy"),
    ("AM", "Revenue Calculation - Senior Bonds", "Commercial Collections", "senior_comm_collections"),
    ("AN", "Revenue Calculation - Senior Bonds", "Commercial Uniform Fee", "senior_comm_uniform_fee"),
    ("AO", "Revenue Calculation - Senior Bonds", "Annual Taxable Sales Revenue", "taxable_sales"),
    ("AP", "Revenue Calculation - Senior Bonds", "Public Improvement Fee (PIF)", "pif_revenue"),
    ("AQ", "Revenue Calculation - Senior Bonds", "System Development Fee", "system_development_fee"),
    ("AR", "Revenue Calculation - Senior Bonds", "Total Available Revenue for Senior Bonds", "senior_total_revenue"),
    ("AT", "Expenses", "County Collection Cost", "county_fee"),
    ("AU", "Expenses", "Annual Trustee Fee", "trustee_fee"),
    ("AV", "Expenses", "Annual Admin Costs", "admin_costs"),
    ("AX", "Expenses", "Net Revenue Available for Senior Lien Debt Service", "senior_net_revenue"),
    ("AZ", "Revenue Calculation - Subordinate Bonds", "Total Taxable Value", "sub_taxable_value"),
    ("BB", "Revenue Calculation - Subordinate Bonds", "Mill Levy Collections", "sub_levy_collections"),
    ("BC", "Revenue Calculation - Subordinate Bonds", "Personal Property Uniform Fee", "sub_uniform_fee"),
    ("BF", "Revenue Calculation - Subordinate Bonds", "Commercial Collections", "sub_comm_collections"),
    ("BG", "Revenue Calculation - Subordinate Bonds", "Commercial Uniform Fee", "sub_comm_uniform_fee"),
    ("BK", "Revenue Calculation - Subordinate Bonds", "Total Available Revenue for Subordinate Bonds", "sub_total_revenue"),
    ("BM", "Expenses", "Annual Trustee Fee", "sub_trustee_fee"),
    ("BO", "Expenses", "Net Revenue Available for Subordinate Lien Debt Service", "sub_net_revenue"),
    ("BQ", "Senior Lien Debt Service", "Senior Par Net Debt Service", "senior_net_debt_service"),
    ("BR", "Senior Lien Debt Service", "Refunding Par Net Debt Service", "refunding_net_debt_service"),
    ("BS", "Senior Lien Debt Service", "Total Net Debt Service", "total_net_debt_service"),
    ("BT", "Senior Lien Debt Service", "Coverage Test", "coverage"),
    ("BU", "Senior Lien Debt Service", "Senior Surplus", "senior_surplus"),
    ("BW", "Senior Lien Debt Service", "Funds on Hand (Source of Funds)", "funds_on_hand"),
    ("BX", "Senior Lien Debt Service", "Annual Surplus", "annual_surplus"),
    ("BY", "Senior Lien Debt Service", "Senior Surplus Release", "surplus_release"),
    ("BZ", "Senior Lien Debt Service", "Senior Surplus Fund Balance", "surplus_fund_balance"),
    ("CA", "Senior Lien Debt Service", "Cumulative Surplus", "cumulative_surplus"),
    ("CB", "Senior Lien Debt Service", "Senior Debt / Taxable Value Ratio", "debt_to_taxable_value"),
    ("CC", "Senior Lien Debt Service", "Senior Debt / Market Value Ratio", "debt_to_market_value"),
    ("CH", "Subordinate Lien Debt Service", "Surplus Available for Sub. Debt Service", "sub_available"),
    ("CI", "Subordinate Lien Debt Service", "Application of Prior Year Surplus", "prior_year_surplus_applied"),
    ("CJ", "Subordinate Lien Debt Service", "Total Available for Sub. Debt Service", "sub_total_available"),
    ("CK", "Subordinate Lien Debt Service", "Sub. Bond Interest on Balance", "sub_interest_due"),
    ("CL", "Subordinate Lien Debt Service", "Less Payments Toward Sub. Bond Interest", "sub_interest_paid"),
    ("CM", "Subordinate Lien Debt Service", "Accrued Interest + Interest on Balance", "sub_accrued_added"),
    ("CN", "Subordinate Lien Debt Service", "Less Payments Toward Accrued Interest", "sub_accrued_paid"),
    ("CO", "Subordinate Lien Debt Service", "Balance of Accrued Interest", "sub_accrued_balance"),
    ("CQ", "Subordinate Lien Debt Service", "Total Sub. Bonds Principal Issued", "sub_principal_issued"),
    ("CR", "Subordinate Lien Debt Service", "Less Payments Toward Bond Principal", "sub_principal_paid"),
    ("CS", "Subordinate Lien Debt Service", "Balance of Sub. Bond Principal", "sub_principal_balance"),
    ("CU", "Subordinate Lien Debt Service", "Total Sub. Debt Payments", "sub_total_payments"),
    ("CW", "Subordinate Lien Debt Service", "Surplus Cash Flow", "sub_surplus_cashflow"),
    ("CY", "Subordinate Lien Debt Service", "Cumulative Surplus", "sub_cumulative_surplus"),
]

_SUMMARY_FORMATS = {
    "assessment_date": FMT_DATE, "tax_revenue_date": FMT_DATE,
    "lot_units": FMT_INT, "residential_units": FMT_INT,
    "senior_mill_levy": FMT_MILLS, "senior_comm_mill_levy": FMT_MILLS,
    "coverage": FMT_RATIO, "debt_to_taxable_value": FMT_RATIO,
    "debt_to_market_value": FMT_PCT,
    "comm_sf_delivered": FMT_INT, "comm_sf_sold": FMT_INT,
}


def _write_summary(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Summary")
    _title_block(ws, cfg, res, col=5)

    groups: dict[str, list[str]] = {}
    for letter, group, _, _ in SUMMARY_COLUMNS:
        if group:
            groups.setdefault(group, []).append(letter)
    for group, letters in groups.items():
        _put(ws, 9, ws[f"{letters[0]}1"].column, group, font=SUB_FONT)

    for letter, _, header, attr in SUMMARY_COLUMNS:
        col = ws[f"{letter}1"].column
        _header(ws, 10, col, header)
        ws.column_dimensions[letter].width = 15 if attr not in (
            "assessment_date", "tax_revenue_date") else 13

    for i, row in enumerate(res.summary):
        r = 12 + i
        for letter, _, _, attr in SUMMARY_COLUMNS:
            col = ws[f"{letter}1"].column
            value = getattr(row, attr, None)
            fmt = _SUMMARY_FORMATS.get(attr, FMT_MONEY)
            _put(ws, r, col, value, fmt=fmt)
    ws.freeze_panes = "E12"


# ── Development tabs ──────────────────────────────────────────────────────────

def _write_block(ws, anchor_col: int, header_row: int, title: str,
                 product_names: Sequence[str], years: Sequence[date],
                 block: Sequence[Sequence[float]], fmt: str,
                 start_row: int = 14) -> int:
    """One side-by-side block: date column, one column per product, a total."""
    _put(ws, 6, anchor_col, title, font=SUB_FONT)
    _put(ws, 12, anchor_col, "Product")
    for j, name in enumerate(product_names):
        _header(ws, 12, anchor_col + 1 + j, name)
    _header(ws, 12, anchor_col + 1 + len(product_names), "Totals")
    for i, d in enumerate(years):
        r = start_row + i
        _put(ws, r, anchor_col, d, fmt=FMT_DATE)
        for j in range(len(product_names)):
            _put(ws, r, anchor_col + 1 + j, block[i][j], fmt=fmt)
        _put(ws, r, anchor_col + 1 + len(product_names), sum(block[i]), fmt=fmt)
    total_row = start_row + len(years)
    _put(ws, total_row, anchor_col, "Total", font=SUB_FONT, border=TOP_BORDER)
    for j in range(len(product_names) + 1):
        col = anchor_col + 1 + j
        col_total = (sum(block[i][j] for i in range(len(years)))
                     if j < len(product_names)
                     else sum(sum(block[i]) for i in range(len(years))))
        _put(ws, total_row, col, col_total, fmt=fmt, font=SUB_FONT, border=TOTAL_BORDER)
    return anchor_col + len(product_names) + 3


def _write_residential(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Residential Development")
    dev = res.dev
    _title_block(ws, cfg, res, col=4)
    names = [p.name or f"Product {i + 1}" for i, p in enumerate(dev.products)]
    years = dev.assessment_dates(cfg)

    blocks = [
        ("LOT DELIVERY", dev.lot_delivery(cfg), FMT_INT),
        (f"LOT DELIVERY SCENARIO @ {cfg.lot_delivery_scenario:.2%}",
         dev.lot_delivery_scenario(cfg), FMT_INT),
        ("LOT VALUE", dev.lot_value(cfg), FMT_MONEY),
        ("LOT VALUE (LAGGED FOR TAX ROLLS)", dev.lot_value_lagged(cfg), FMT_MONEY),
        ("FINISHED LOTS USED", dev.lots_consumed(cfg), FMT_MONEY),
        ("HOME CLOSINGS", dev.home_closings(cfg), FMT_INT),
        (f"HOME CLOSINGS @ {cfg.absorption_scenario:.2%}",
         dev.closings_scenario(cfg), FMT_INT),
        ("RESIDENTIAL PRICING", dev.pricing(cfg), FMT_MONEY),
        ("TAXABLE VALUE CREATION", dev.av_creation(cfg), FMT_MONEY),
        ("TAXABLE VALUE CREATION (LAGGED FOR TAX ROLLS)",
         dev.av_creation_lagged(cfg), FMT_MONEY),
    ]
    col = 4
    for title, block, fmt in blocks:
        col = _write_block(ws, col, 12, title, names, years, block, fmt)

    # Reference rows carried above every block, as in the template.
    _put(ws, 10, 4, "Total Units", font=SUB_FONT)
    for j, p in enumerate(dev.products):
        _put(ws, 10, 5 + j, p.total_units, fmt=FMT_INT, font=SUB_FONT)
    _put(ws, 10, 5 + len(names), dev.total_units(), fmt=FMT_INT, font=SUB_FONT)
    _put(ws, 11, 4, "Vacant Land Value Per Lot", font=SUB_FONT)
    for j, v in enumerate(dev.vacant_land_value_per_lot(cfg)):
        _put(ws, 11, 5 + j, v, fmt=FMT_MONEY)
    _put(ws, 9, 4, "Product/Builder", font=SUB_FONT)
    for j, p in enumerate(dev.products):
        _put(ws, 9, 5 + j, p.product_type)
    ws.freeze_panes = "E13"


def _write_commercial(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Comm Development")
    dev = res.dev
    _title_block(ws, cfg, res, col=2)
    if not dev.commercial:
        _put(ws, 6, 2, "No commercial component in this scenario.", font=SUB_FONT)
        _put(ws, 8, 2, "The tab is retained so the layout matches the template; "
                       "add CommercialProduct entries to populate it.",
             font=Font(name="Calibri", size=9, italic=True, color="808080"))
        return
    names = [c.name for c in dev.commercial]
    years = dev.assessment_dates(cfg)
    col = 2
    for title, block, fmt in [
        ("SQUARE FEET DELIVERED", dev.comm_sf_delivered(cfg), FMT_INT),
        ("SQUARE FEET SOLD / LEASED", dev.comm_sf_sold(cfg), FMT_INT),
        ("COMMERCIAL VALUE CREATED", dev.comm_value_created(cfg), FMT_MONEY),
    ]:
        col = _write_block(ws, col, 12, title, names, years, block, fmt)


# ── Ops Rev & Exp ─────────────────────────────────────────────────────────────

def _write_ops(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("Ops Rev & Exp Projection")
    _widths(ws, {"A": 6, "B": 3, "C": 14, "D": 20, "E": 14,
                 "F": 18, "G": 18, "H": 20, "I": 14, "J": 18, "K": 14})
    _title_block(ws, cfg, res, col=3)
    heads = ["Year", "Total Taxable Valuation", "Operations Mill Levy",
             f"Total Collections @ {cfg.tax_collect_mill_prc:.2%}",
             f"Personal Property Uniform Fee @ {cfg.uniform_fee_prc:.2%}",
             "Total Available for O&M", "Contribution Mill Levy",
             "Contribution Collections", "Total Mills"]
    for j, text in enumerate(heads):
        _header(ws, 7, 3 + j, text)
    total_mills = (cfg.mill_levy_ds_target + cfg.mill_levy_ops_target
                   + cfg.contribution_rate)
    for i, row in enumerate(res.summary):
        r = 8 + i
        _put(ws, r, 1, i + 1 if i else None, fmt=FMT_INT)
        _put(ws, r, 3, row.assessment_date, fmt=FMT_DATE)
        _put(ws, r, 4, row.senior_taxable_value, fmt=FMT_MONEY)
        _put(ws, r, 5, cfg.mill_levy_ops_target, fmt=FMT_MILLS)
        collections = (row.senior_taxable_value / 1000 * cfg.mill_levy_ops_target
                       * cfg.tax_collect_mill_prc)
        uniform = collections * cfg.uniform_fee_prc
        _put(ws, r, 6, collections, fmt=FMT_MONEY)
        _put(ws, r, 7, uniform, fmt=FMT_MONEY)
        _put(ws, r, 8, collections + uniform, fmt=FMT_MONEY)
        _put(ws, r, 9, cfg.contribution_rate, fmt=FMT_MILLS)
        _put(ws, r, 10, row.senior_taxable_value / 1000 * cfg.contribution_rate
             * cfg.tax_collect_mill_prc, fmt=FMT_MONEY)
        _put(ws, r, 11, total_mills, fmt=FMT_MILLS)
    if cfg.mill_levy_ops_target == 0:
        _put(ws, 5, 3, "This PID levies for debt service only — no separate "
                       "operations levy. Administration is charged against "
                       "pledged revenue on the Summary tab.",
             font=Font(name="Calibri", size=9, italic=True, color="808080"))


# ── Debt service tabs ─────────────────────────────────────────────────────────

_DS_COLUMNS = [
    ("Payment Date", "payment_date", FMT_DATE),
    ("Rate", "rate", FMT_PCT3),
    ("Yield", "yld", FMT_PCT3),
    ("Price", "price", FMT_PRICE),
    ("S/T", None, None),
    ("Premium/OID", "premium_oid", FMT_MONEY),
    ("Principal", "principal", FMT_MONEY),
    ("Interest", "interest", FMT_MONEY_2),
    ("Total", "total", FMT_MONEY_2),
    ("Annual Gross Total", "annual_gross", FMT_MONEY_2),
    ("Capitalized Interest", "capitalized_interest", FMT_MONEY_2),
    ("Surplus Release", "surplus_release", FMT_MONEY),
    ("Interest Earnings", "interest_earnings", FMT_MONEY),
    ("Net Total", "net_total", FMT_MONEY_2),
    ("Annual Net Total", "annual_net", FMT_MONEY_2),
    ("Bond Value", "bond_value", FMT_MONEY),
]
_DS_EXTRA = [
    ("Pledged Revenues", "revenue", FMT_MONEY),
    ("Actual Debt Service Coverage", "actual_coverage", FMT_RATIO),
    ("Coverage Target Factor", "coverage_target", FMT_RATIO),
    ("Bond Years", "bond_years", FMT_MONEY_2),
]


def _write_debt_service(wb: Workbook, title: str, cfg: ModelConfig, res: Results,
                        rows, stats, subtitle: str,
                        columns=None, extra=None) -> None:
    ws = wb.create_sheet(title)
    _title_block(ws, cfg, res)
    _put(ws, 4, 1, subtitle, font=SUB_FONT)
    columns = columns or _DS_COLUMNS
    extra = _DS_EXTRA if extra is None else extra

    all_cols = list(columns) + [("", None, None)] + list(extra)
    for j, (header, _, _) in enumerate(all_cols):
        if header:
            _header(ws, 9, 1 + j, header)
        ws.column_dimensions[get_column_letter(1 + j)].width = 15

    for i, row in enumerate(rows):
        r = 11 + i
        for j, (_, attr, fmt) in enumerate(all_cols):
            if attr is None:
                continue
            _put(ws, r, 1 + j, getattr(row, attr, None), fmt=fmt)

    total_row = 11 + len(rows) + 1
    _put(ws, total_row, 1, "Total", font=SUB_FONT, border=TOP_BORDER)
    for j, (_, attr, fmt) in enumerate(all_cols):
        if attr in (None, "payment_date", "rate", "yld", "price",
                    "actual_coverage", "coverage_target", "bond_value"):
            continue
        _put(ws, total_row, 1 + j, sum(getattr(r, attr, 0.0) or 0.0 for r in rows),
             fmt=fmt, font=SUB_FONT, border=TOTAL_BORDER)

    stat_row = total_row + 3
    _put(ws, stat_row, 1, "Bond Statistics", font=SECTION_FONT)
    for k, (label, value, fmt) in enumerate([
        ("Par Amount", stats.par, FMT_MONEY),
        ("Premium / (Discount)", stats.premium, FMT_MONEY),
        ("Bond Years", stats.bond_years, FMT_MONEY_2),
        ("Average Life (years)", stats.average_life, '0.00'),
        ("Arbitrage TIC", stats.arbitrage_tic, FMT_PCT3),
        ("All-in TIC", stats.all_in_tic, FMT_PCT3),
        ("Maximum Annual Debt Service", stats.max_annual_debt_service, FMT_MONEY),
        ("Total Debt Service", stats.total_debt_service, FMT_MONEY),
        ("First Maturity", stats.first_maturity, FMT_DATE),
        ("Final Maturity", stats.final_maturity, FMT_DATE),
    ]):
        _put(ws, stat_row + 1 + k, 1, label)
        _put(ws, stat_row + 1 + k, 2, value, fmt=fmt)
    ws.freeze_panes = "B11"


_SUB_COLUMNS = [
    ("Payment Date", "payment_date", FMT_DATE),
    ("Rate", "rate", FMT_PCT3),
    ("Yield", "yld", FMT_PCT3),
    ("Price", "price", FMT_PRICE),
    ("Premium/OID", "premium_oid", FMT_MONEY),
    ("", None, None),
    ("Principal", "principal", FMT_MONEY),
    ("Semi Annual Interest", "interest", FMT_MONEY_2),
    ("Annual Interest", "annual_interest", FMT_MONEY_2),
    ("Less Payments Toward Sub. Bond Interest", "interest_paid", FMT_MONEY_2),
    ("Accrued Interest + Interest on Balance", "accrued_added", FMT_MONEY_2),
    ("Less Payments Toward Accrued Interest", "accrued_paid", FMT_MONEY_2),
    ("Balance of Bond Interest + Accrued Interest", "accrued_balance", FMT_MONEY_2),
    ("Aggregate Debt Service on Sub. Bonds", "aggregate_debt_service", FMT_MONEY_2),
    ("Total", "total", FMT_MONEY_2),
    ("Annual Gross Total", "annual_gross", FMT_MONEY_2),
    ("Capitalized Interest", "capitalized_interest", FMT_MONEY_2),
    ("Surplus Release", "surplus_release", FMT_MONEY),
    ("Net Total", "net_total", FMT_MONEY_2),
    ("Annual Net Total", "annual_net", FMT_MONEY_2),
    ("Bond Value", "bond_value", FMT_MONEY),
]


# ── CAPI Fund ─────────────────────────────────────────────────────────────────

def _write_capi(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("CAPI Fund - First")
    _widths(ws, {"A": 14, "B": 16, "C": 16, "D": 16, "E": 3, "F": 16,
                 "G": 10, "H": 12, "I": 14, "J": 16, "K": 16})
    _put(ws, 5, 1, "Capitalized Interest Fund", font=TITLE_FONT)
    _put(ws, 6, 1, "Fund Yield:")
    _put(ws, 6, 2, cfg.interest_earn_rate, fmt=FMT_PCT)
    for col, (a, b) in enumerate([("Draw", "Date"), ("Construction", "Amount"),
                                  ("Capitalized", "Interest"), ("Combined", "Draws"),
                                  ("", ""), ("Beginning", "Balance"), ("Days", ""),
                                  ("Periodic", "Rate"), ("Plus", "Interest"),
                                  ("Less", "Draws"), ("Ending", "Balance")], start=1):
        if a:
            _put(ws, 8, col, a, font=SUB_FONT, align="center")
            _put(ws, 9, col, b, font=SUB_FONT, align="center")
    for i, row in enumerate(res.capi):
        r = 11 + i
        _put(ws, r, 1, row.draw_date, fmt=FMT_DATE)
        _put(ws, r, 2, row.construction_amount, fmt=FMT_MONEY_2)
        _put(ws, r, 3, row.capitalized_interest, fmt=FMT_MONEY_2)
        _put(ws, r, 4, row.combined_draws, fmt=FMT_MONEY_2)
        _put(ws, r, 6, row.beginning_balance, fmt=FMT_MONEY_2)
        _put(ws, r, 7, row.days, fmt=FMT_INT)
        _put(ws, r, 8, row.periodic_rate, fmt=FMT_PCT3)
        _put(ws, r, 9, row.interest, fmt=FMT_MONEY_2)
        _put(ws, r, 10, row.draws, fmt=FMT_MONEY_2)
        _put(ws, r, 11, row.ending_balance, fmt=FMT_MONEY_2)
    end = 11 + len(res.capi) + 1
    _put(ws, end, 1, "Total", font=SUB_FONT, border=TOP_BORDER)
    for col, attr in ((3, "capitalized_interest"), (4, "combined_draws"),
                      (9, "interest"), (10, "draws")):
        _put(ws, end, col, sum(getattr(r, attr) for r in res.capi),
             fmt=FMT_MONEY_2, font=SUB_FONT, border=TOTAL_BORDER)


# ── DBC Output ────────────────────────────────────────────────────────────────

def _write_dbc(wb: Workbook, cfg: ModelConfig, res: Results) -> None:
    ws = wb.create_sheet("DBC Output")
    _widths(ws, {"A": 3, "B": 3, "C": 14, "D": 14, "E": 16, "F": 16, "G": 16, "H": 16})
    _put(ws, 1, 5, "Senior Bonds", font=SUB_FONT, align="center")
    _put(ws, 1, 7, "Sub Bonds", font=SUB_FONT, align="center")
    _put(ws, 2, 5, "Principal", font=SUB_FONT, align="center")
    _put(ws, 2, 6, "DS", font=SUB_FONT, align="center")
    _put(ws, 2, 7, "Principal", font=SUB_FONT, align="center")
    _put(ws, 2, 8, "DS", font=SUB_FONT, align="center")

    senior = {r.payment_date.year: r for r in res.senior
              if r.payment_date.month == cfg.prin_maturity}
    sub = {r.payment_date.year: r for r in res.sub_annual}
    years = sorted(set(senior) | set(sub))
    for i, y in enumerate(years):
        r = 3 + i
        s, b = senior.get(y), sub.get(y)
        _put(ws, r, 3, date(y, cfg.prin_maturity, cfg.prin_maturity_day_sub), fmt=FMT_DATE)
        _put(ws, r, 4, date(y, cfg.prin_maturity, cfg.prin_maturity_day_senior), fmt=FMT_DATE)
        _put(ws, r, 5, s.principal if s else 0.0, fmt=FMT_MONEY)
        _put(ws, r, 6, s.annual_net if s else 0.0, fmt=FMT_MONEY)
        _put(ws, r, 7, b.principal if b else 0.0, fmt=FMT_MONEY)
        _put(ws, r, 8, b.aggregate_debt_service if b else 0.0, fmt=FMT_MONEY)
    end = 3 + len(years) + 1
    _put(ws, end, 3, "Total", font=SUB_FONT, border=TOP_BORDER)
    for col, values in (
        (5, [s.principal for s in senior.values()]),
        (6, [s.annual_net for s in senior.values()]),
        (7, [b.principal for b in sub.values()]),
        (8, [b.aggregate_debt_service for b in sub.values()]),
    ):
        _put(ws, end, col, sum(values), fmt=FMT_MONEY, font=SUB_FONT, border=TOTAL_BORDER)


# ── Entry point ───────────────────────────────────────────────────────────────

def build_workbook(res: Results, path: str) -> str:
    """Write the full workbook and return the path."""
    cfg = res.cfg
    wb = Workbook()
    wb.remove(wb.active)

    _write_inputs(wb, cfg, res)
    _write_capital_costs(wb, cfg, res)
    _write_costs_of_issuance(wb, cfg, res)
    _write_sources_uses(wb, cfg, res)
    _write_sources_uses_refunding(wb, cfg, res)
    _write_summary(wb, cfg, res)
    _write_residential(wb, cfg, res)
    _write_commercial(wb, cfg, res)
    _write_ops(wb, cfg, res)
    _write_debt_service(wb, "Senior Lien DS - First", cfg, res, res.senior,
                        res.senior_stats,
                        f"{cfg.senior_bonds_series} - ${res.senior_par / 1e6:,.2f}MM "
                        f"({cfg.dsc_senior_lien_bonds:.2f}x)")
    _write_debt_service(wb, "Sub Lien DS - First (Annual)", cfg, res, res.sub_annual,
                        res.sub_stats,
                        f"{cfg.sub_bonds_series} - ${res.sub_par / 1e6:,.2f}MM "
                        f"({cfg.dsc_sub_lien_bonds:.2f}x) — annual",
                        columns=_SUB_COLUMNS, extra=[])
    _write_debt_service(wb, "Sub Lien DS - First (SA)", cfg, res, res.sub_sa,
                        res.sub_stats,
                        f"{cfg.sub_bonds_series} - ${res.sub_par / 1e6:,.2f}MM "
                        f"({cfg.dsc_sub_lien_bonds:.2f}x) — semi-annual",
                        columns=_SUB_COLUMNS, extra=[])
    _write_debt_service(wb, "Senior Lien DS - Refunding", cfg, res, res.refunding,
                        res.refunding_stats,
                        f"Refunding {cfg.refund_bonds_series} - "
                        f"${res.refunding_stats.par / 1e6:,.2f}MM")
    _write_capi(wb, cfg, res)

    scratch = wb.create_sheet("Scratch--->>>")
    _put(scratch, 1, 1, "Scratch--->>>", font=TITLE_FONT)
    _put(scratch, 3, 1, "Working area. Everything to the right of this tab is output.",
         font=Font(name="Calibri", size=9, italic=True, color="808080"))

    _write_dbc(wb, cfg, res)

    if res.warnings:
        ws = wb["Inputs - First"]
        _put(ws, 140, 2, "Model notes", font=SECTION_FONT)
        for i, w in enumerate(res.warnings):
            _put(ws, 141 + i, 2, f"• {w}",
                 font=Font(name="Calibri", size=9, italic=True, color="C00000"), wrap=True)

    wb.save(path)
    return path
